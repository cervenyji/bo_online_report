# =============================================================================
# Report vytíženosti pracovišť (Back Office) — vložte do JEDNÉ buňky Jupyter notebooku
#
# Vstupy (v aktuálním adresáři, nebo uveďte plnou cestu níže v KONFIGURACI):
#   bo_data.xlsx               : BRANCH_ID, PRACOVISTE_ID, DATETIME, ZAMESTNANEC, ACTIVITY, DURATION (minuty)
#   qr_codes_bo_online.xlsx    : BRANCH_ID, BRANCH_NAME, NO_WPL (počet pracovišť),
#                                 CAPACITY (týdenní otevírací doba v hodinách),
#                                 VIKENDOVA (bool, otevřeno i o víkendu), POLEDNI_PAUZA (bool)
#   segmenty_pracovist.xlsx    : BRANCH_ID, PRACOVISTE_ID, SEGMENT — volitelné, chybějící
#                                 kombinace se v reportu zobrazí jako "—"
#
# Výstup: samostatný HTML report (Plotly.js vložený přímo v souboru, funguje i offline).
#
# Poznámka: .xlsx soubory se čtou vlastní minimální čtečkou (jen zipfile + xml ze
# standardní knihovny) — NENÍ potřeba mít nainstalovaný openpyxl ani žádný jiný
# balíček pro práci s Excelem. Potřeba je jen pandas, numpy a plotly.
# =============================================================================

from __future__ import annotations

import math
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from IPython.display import IFrame, display

# -----------------------------------------------------------------------------
# 1. KONFIGURACE — upravte cesty a případně prahové hodnoty
# -----------------------------------------------------------------------------

BO_DATA_FILE = "bo_data.xlsx"
WORKSPACES_FILE = "qr_codes_bo_online.xlsx"
SEGMENTS_FILE = WORKSPACES_FILE  # segmenty (BRANCH_ID, PRACOVISTE_ID, SEGMENT) bývají jako list
SEGMENTS_SHEET_NAME = "struktura_pracovist_segmenty"  # uvnitř qr_codes_bo_online.xlsx, ne samostatný soubor
OUTPUT_HTML = "vytizenost_report.html"  # ke jménu se při každém běhu přidá časové razítko (viz níže),
                                         # aby prohlížeč/Jupyter nikdy nezobrazoval starou zkešovanou verzi souboru

THRESHOLD_HIGH = 70.0      # od této hranice (%) je pracoviště "vysoce vytížené"
THRESHOLD_CRITICAL = 90.0  # od této hranice (%) je pracoviště "kriticky vytížené" / na hraně kapacity

# Pobočka provozuje Po-Pá (5 dní/týden), nebo i o víkendu (7 dní/týden) — řídí se
# sloupcem VIKENDOVA z work_spaces.xlsx; denní kapacita = týdenní CAPACITY / počet dní.
DAYS_PER_WEEK_OPEN = 5
DAYS_PER_WEEK_OPEN_WEEKEND = 7

# Polední pauza (pokud POLEDNI_PAUZA=True u pobočky) — vizuálně vyznačena v denním
# rozvrhu pracovišť; přesný čas neznáme ze zdrojových dat, proto je zde jako konstanta.
LUNCH_BREAK_START_HOUR = 12
LUNCH_BREAK_END_HOUR = 13

BLOCK_MINUTES = 10  # velikost bloku pro denní rozvrh pracovišť
MAX_WEEKS_PER_BRANCH = 26  # bezpečnostní strop počtu týdnů v přepínači denních rozvrhů (nejnovější týdny mají přednost)

# --- Pilotní rozšíření (jen pobočka Praha 2 - Jugoslávská) -------------------
# Nepřítomnost zaměstnanců (dovolená, nemoc, home office, ...) se z evidence počítá
# jako "ten den nemohl použít žádné pracoviště" a promítá se do efektivní kapacity.
# Zároveň je tahle pobočka fyzicky rozdělená na dvě budovy s jiným číslováním
# pracovišť - platí to VÝHRADNĚ pro PILOT_BRANCH_ID, nikde jinde v kódu se to
# nezobecňuje.
ABSENCE_FILE = "2026071_14_Nepřítomnosti.xlsx"  # export z HR systému - název souboru se mění dle exportu, upravte
PILOT_BRANCH_ID = 459  # Praha 2 (Jugoslávská)
PILOT_ORG_UNIT_ZKRATKA = "UP_459"  # filtr sloupce 'Zkratka organizační jednotky' v souboru nepřítomností
PILOT_BUILDINGS = {
    "Jugoslávská (budova)": range(1, 5),    # pracoviště 1-4 - zatím nenainstalováno
    "Olbrachtova (budova)": range(5, 21),   # pracoviště 5-20
}

# --- Barevná paleta (validovaná: lightness/chroma/CVD/kontrast) --------------
PAGE_BG = "#f9f9f7"
SURFACE = "#fcfcfb"
TEXT_PRIMARY = "#0b0b0b"
TEXT_SECONDARY = "#52514e"
TEXT_MUTED = "#898781"
GRIDLINE = "#e1e0d9"
BORDER = "rgba(11,11,11,0.10)"

# Status (stav) — rezervováno, vždy s ikonou/popiskem, nikdy jako "5. kategorie"
STATUS_GOOD = "#0ca30c"
STATUS_WARNING = "#fab219"
STATUS_CRITICAL = "#d03b3b"
STATUS_MUTED = "#898781"

BUCKET_COLORS = {"Nízká": STATUS_GOOD, "Vysoká": STATUS_WARNING, "Kritická": STATUS_CRITICAL, "Bez dat": STATUS_MUTED}

# Kategoriální paleta (identita, pevné pořadí, nikdy necyklovat)
CATEGORICAL = ["#2a78d6", "#1baf7a", "#eda100", "#008300", "#4a3aa7", "#e34948", "#e87ba4", "#eb6834"]
BLUE = CATEGORICAL[0]  # výchozí jednobarevná (sekvenční) barva pro grafy bez stavové sémantiky

pd.set_option("display.max_columns", 50)
pd.set_option("display.width", 140)


# -----------------------------------------------------------------------------
# 2. Minimální čtečka .xlsx bez openpyxl (jen standardní knihovna)
# -----------------------------------------------------------------------------

_XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_XLSX_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_XLSX_BUILTIN_DATE_FMT_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 57}


def _xlsx_col_to_index(cell_ref):
    letters = re.match(r"[A-Z]+", cell_ref).group()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _xlsx_is_date_format(fmt_code):
    stripped = re.sub(r'"[^"]*"', "", fmt_code)
    return bool(re.search(r"[ymdhs]", stripped, re.IGNORECASE))


def _xlsx_excel_serial_to_datetime(serial):
    return datetime(1899, 12, 30) + timedelta(days=serial)


def _xlsx_read_styles_date_flags(z):
    if "xl/styles.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/styles.xml"))
    custom_fmts = {}
    numfmts_el = root.find(f"{_XLSX_NS}numFmts")
    if numfmts_el is not None:
        for nf in numfmts_el.findall(f"{_XLSX_NS}numFmt"):
            custom_fmts[int(nf.get("numFmtId"))] = nf.get("formatCode")

    date_flags = []
    cellxfs_el = root.find(f"{_XLSX_NS}cellXfs")
    if cellxfs_el is not None:
        for xf in cellxfs_el.findall(f"{_XLSX_NS}xf"):
            fmt_id = int(xf.get("numFmtId", "0"))
            if fmt_id in _XLSX_BUILTIN_DATE_FMT_IDS:
                date_flags.append(True)
            elif fmt_id in custom_fmts:
                date_flags.append(_xlsx_is_date_format(custom_fmts[fmt_id]))
            else:
                date_flags.append(False)
    return date_flags


class SheetNotFoundError(ValueError):
    """List (podle jména) nebyl v sešitu nalezen — odlišeno od ostatních ValueError,
    aby ho volající mohl chytit zvlášť (např. list se segmenty je nepovinný)."""


def read_xlsx_rows(path, sheet_index=0):
    """Přečte list .xlsx souboru a vrátí seznam řádků (list řádků, každý je list buněk).
    Rozpozná datumové buňky podle formátu ve stylu a rovnou je vrátí jako datetime.
    `sheet_index` může být pořadové číslo (0 = první list), nebo jméno listu (str)."""
    with zipfile.ZipFile(path) as z:
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{_XLSX_NS}si"):
                text = "".join(t.text or "" for t in si.iter(f"{_XLSX_NS}t"))
                shared_strings.append(text)

        date_flags = _xlsx_read_styles_date_flags(z)

        workbook_root = ET.fromstring(z.read("xl/workbook.xml"))
        sheet_els = workbook_root.findall(f"{_XLSX_NS}sheets/{_XLSX_NS}sheet")
        if isinstance(sheet_index, str):
            matches = [i for i, el in enumerate(sheet_els) if el.get("name") == sheet_index]
            if not matches:
                available = [el.get("name") for el in sheet_els]
                raise SheetNotFoundError(f"List '{sheet_index}' nebyl v souboru {path} nalezen. Dostupné listy: {available}")
            sheet_index = matches[0]
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rels = {r.get("Id"): r.get("Target") for r in rels_root}
        target = rels[sheet_els[sheet_index].get(f"{_XLSX_REL_NS}id")]
        sheet_path = target.lstrip("/") if target.startswith("/") else f"xl/{target}"

        sheet_root = ET.fromstring(z.read(sheet_path))
        rows = []
        for row_el in sheet_root.findall(f"{_XLSX_NS}sheetData/{_XLSX_NS}row"):
            row_cells = {}
            max_col = -1
            for c in row_el.findall(f"{_XLSX_NS}c"):
                col_idx = _xlsx_col_to_index(c.get("r"))
                max_col = max(max_col, col_idx)
                cell_type = c.get("t")
                style_idx = int(c.get("s", "0"))
                v_el = c.find(f"{_XLSX_NS}v")

                if v_el is None:
                    is_el = c.find(f"{_XLSX_NS}is")
                    value = "".join(t.text or "" for t in is_el.iter(f"{_XLSX_NS}t")) if is_el is not None else None
                else:
                    raw = v_el.text
                    if cell_type == "s":
                        value = shared_strings[int(raw)]
                    elif cell_type == "b":
                        value = bool(int(raw))
                    elif cell_type == "str":
                        value = raw
                    else:
                        num = float(raw)
                        if style_idx < len(date_flags) and date_flags[style_idx]:
                            value = _xlsx_excel_serial_to_datetime(num)
                        else:
                            value = int(num) if num == int(num) else num
                row_cells[col_idx] = value
            rows.append([row_cells.get(i) for i in range(max_col + 1)])

    width = max((len(r) for r in rows), default=0)
    rows = [r + [None] * (width - len(r)) for r in rows]
    # Excel často eviduje "použitou oblast" o pár řádků větší, než jsou reálná data
    # (formátování, staré řádky apod.) -> takové čistě prázdné řádky zahodíme.
    return [r for r in rows if any(v is not None and v != "" for v in r)]


def read_xlsx(path, sheet_index=0):
    """Přečte .xlsx do pandas DataFrame (první řádek = hlavička). Bez openpyxl."""
    rows = read_xlsx_rows(path, sheet_index)
    header, *data = rows
    return pd.DataFrame(data, columns=header)


def read_xlsx_with_header_row(path, header_row=0, sheet_index=0):
    """Jako read_xlsx, ale hlavička není na prvním řádku listu — `header_row` řádků
    (typicky export-metadata) se přeskočí. Odpovídá pandas.read_excel(skiprows=header_row)."""
    rows = read_xlsx_rows(path, sheet_index)[header_row:]
    header, *data = rows
    return pd.DataFrame(data, columns=header)


# -----------------------------------------------------------------------------
# 3. Načtení a příprava dat
# -----------------------------------------------------------------------------

def _parse_datetime_cell(value):
    """DATETIME buňka může přijít jako text ('02.07.2026 09:25:40') i jako
    skutečné Excel datum (pak už je to hotový datetime z read_xlsx_rows)."""
    if isinstance(value, datetime):
        return value
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    text = str(value).strip()
    try:
        return datetime.strptime(text, "%d.%m.%Y %H:%M:%S")
    except ValueError:
        return pd.to_datetime(text, dayfirst=True, errors="coerce")


def load_activities(path):
    """Načte bo_data.xlsx, sjednotí názvy sloupců (dle pozice, ne dle přesného
    znění hlavičky) a vrátí (očištěná_data, řádky_s_problémem)."""
    raw = read_xlsx(path)
    if raw.shape[1] < 6:
        raise ValueError(
            f"Soubor {path} má jen {raw.shape[1]} sloupců, očekává se 6: "
            "BRANCH_ID, PRACOVISTE_ID, DATETIME, ZAMESTNANEC, ACTIVITY, DURATION"
        )
    df = raw.iloc[:, :6].copy()
    df.columns = ["BRANCH_ID", "WORKSTATION_ID", "DATETIME_RAW", "EMPLOYEE", "ACTIVITY", "DURATION_MIN"]

    df["BRANCH_ID"] = pd.to_numeric(df["BRANCH_ID"], errors="coerce")
    df["WORKSTATION_ID"] = pd.to_numeric(df["WORKSTATION_ID"], errors="coerce")
    df["DURATION_MIN"] = pd.to_numeric(df["DURATION_MIN"], errors="coerce")
    df["DATETIME"] = pd.to_datetime(df["DATETIME_RAW"].map(_parse_datetime_cell), errors="coerce")
    df["DATE"] = df["DATETIME"].dt.normalize()
    df["END_DATETIME"] = df["DATETIME"] + pd.to_timedelta(df["DURATION_MIN"], unit="m")

    bad_mask = (
        df["BRANCH_ID"].isna() | df["WORKSTATION_ID"].isna() | df["DATETIME"].isna()
        | df["DURATION_MIN"].isna() | (df["DURATION_MIN"] <= 0)
    )
    issues = df.loc[bad_mask].copy()
    clean = df.loc[~bad_mask].copy()
    clean["BRANCH_ID"] = clean["BRANCH_ID"].astype(int)
    clean["WORKSTATION_ID"] = clean["WORKSTATION_ID"].astype(int)
    return clean.reset_index(drop=True), issues.reset_index(drop=True)


def _coerce_bool(series):
    """Sloupec VIKENDOVA/POLEDNI_PAUZA může přijít jako skutečný bool, text
    ('ANO'/'NE', 'TRUE'/'FALSE') nebo číslo (0/1) — sjednotíme na bool, chybějící -> False."""
    def _one(v):
        if isinstance(v, bool):
            return v
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        if isinstance(v, (int, float)):
            return bool(v)
        text = str(v).strip().lower()
        return text in {"true", "ano", "yes", "1", "y", "a"}
    return series.map(_one)


def load_workspaces(path):
    """Načte work_spaces.xlsx a sjednotí názvy sloupců dle pozice. CAPACITY je
    týdenní otevírací doba v hodinách — denní kapacita se odvozuje podle toho,
    jestli je pobočka otevřená i o víkendu (VIKENDOVA)."""
    raw = read_xlsx(path)
    if raw.shape[1] < 6:
        raise ValueError(
            f"Soubor {path} má jen {raw.shape[1]} sloupců, očekává se 6: "
            "BRANCH_ID, BRANCH_NAME, NO_WPL, CAPACITY, VIKENDOVA, POLEDNI_PAUZA"
        )
    df = raw.iloc[:, :6].copy()
    df.columns = ["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_WEEK_HOURS", "VIKENDOVA", "POLEDNI_PAUZA"]
    df["BRANCH_ID"] = pd.to_numeric(df["BRANCH_ID"], errors="coerce")
    df["NO_WORKSTATIONS"] = pd.to_numeric(df["NO_WORKSTATIONS"], errors="coerce")
    df["CAPACITY_WEEK_HOURS"] = pd.to_numeric(df["CAPACITY_WEEK_HOURS"], errors="coerce")
    df["VIKENDOVA"] = _coerce_bool(df["VIKENDOVA"])
    df["POLEDNI_PAUZA"] = _coerce_bool(df["POLEDNI_PAUZA"])

    bad_mask = df["BRANCH_ID"].isna() | df["NO_WORKSTATIONS"].isna() | df["CAPACITY_WEEK_HOURS"].isna()
    if bad_mask.any():
        print(f"Pozor: {bad_mask.sum()} řádků ve work_spaces.xlsx bylo přeskočeno (chybí BRANCH_ID/NO_WPL/CAPACITY):")
        display(raw.iloc[:, :6].loc[bad_mask])
    df = df.loc[~bad_mask].copy()

    df["BRANCH_ID"] = df["BRANCH_ID"].astype(int)
    df["NO_WORKSTATIONS"] = df["NO_WORKSTATIONS"].astype(int)
    days_per_week = np.where(df["VIKENDOVA"], DAYS_PER_WEEK_OPEN_WEEKEND, DAYS_PER_WEEK_OPEN)
    df["CAPACITY_DAY_HOURS"] = df["CAPACITY_WEEK_HOURS"] / days_per_week
    df["CAPACITY_MIN"] = df["CAPACITY_DAY_HOURS"] * 60
    return df.reset_index(drop=True)


def load_segments(path, sheet_name=SEGMENTS_SHEET_NAME):
    """Načte převodník BRANCH_ID+PRACOVISTE_ID -> SEGMENT. Bývá jako list `sheet_name`
    uvnitř work_spaces.xlsx (ne samostatný soubor). Pokud soubor/list chybí (např.
    ještě není pro všechny pobočky připravený), vrátí prázdnou tabulku a report jen
    zobrazí segment jako „—" — nic nespadne."""
    try:
        raw = read_xlsx(path, sheet_index=sheet_name)
    except (FileNotFoundError, zipfile.BadZipFile):
        print(f"Pozor: soubor se segmenty '{path}' nebyl nalezen — sloupec SEGMENT bude prázdný.")
        return pd.DataFrame(columns=["BRANCH_ID", "WORKSTATION_ID", "SEGMENT"])
    except SheetNotFoundError:
        print(f"Pozor: list '{sheet_name}' se segmenty v souboru '{path}' nebyl nalezen — sloupec SEGMENT bude prázdný.")
        return pd.DataFrame(columns=["BRANCH_ID", "WORKSTATION_ID", "SEGMENT"])
    if raw.shape[1] < 3:
        raise ValueError(f"Soubor {path} má jen {raw.shape[1]} sloupců, očekává se 3: BRANCH_ID, PRACOVISTE_ID, SEGMENT")
    df = raw.iloc[:, :3].copy()
    df.columns = ["BRANCH_ID", "WORKSTATION_ID", "SEGMENT"]
    df["BRANCH_ID"] = pd.to_numeric(df["BRANCH_ID"], errors="coerce")
    df["WORKSTATION_ID"] = pd.to_numeric(df["WORKSTATION_ID"], errors="coerce")
    df = df.dropna(subset=["BRANCH_ID", "WORKSTATION_ID"]).copy()
    df["BRANCH_ID"] = df["BRANCH_ID"].astype(int)
    df["WORKSTATION_ID"] = df["WORKSTATION_ID"].astype(int)
    return df.reset_index(drop=True)


def load_absences(path, org_unit_zkratka):
    """Načte export nepřítomností z HR systému (Nepřítomnosti.xlsx). V reálném
    exportu bývají před hlavičkou 2 řádky metadat (odpovídá skiprows=2). Filtruje
    jen na `org_unit_zkratka` (pilotní pobočka) a vrátí denní tabulku EMPLOYEE×DATE
    s podílem dne (ABSENCE_FRACTION, 0-1), kdy zaměstnanec nebyl k dispozici -
    v tomto exportu KAŽDÝ typ nepřítomnosti (dovolená, nemoc, home office, ...)
    znamená, že ten den nemohl použít žádné pracoviště."""
    try:
        raw = read_xlsx_with_header_row(path, header_row=2)
    except (FileNotFoundError, zipfile.BadZipFile):
        print(f"Pozor: soubor s nepřítomnostmi '{path}' nebyl nalezen — sekce nepřítomnosti se vynechá.")
        return pd.DataFrame(columns=["EMPLOYEE", "DATE", "ABSENCE_FRACTION"])

    needed = ["Příjmení", "Jméno", "Zkratka organizační jednotky", "Začátek - datum", "Konec - datum", "Počet dní"]
    missing = [c for c in needed if c not in raw.columns]
    if missing:
        raise ValueError(f"V souboru {path} chybí očekávané sloupce: {missing}")

    df = raw.loc[raw["Zkratka organizační jednotky"] == org_unit_zkratka].copy()
    if df.empty:
        return pd.DataFrame(columns=["EMPLOYEE", "DATE", "ABSENCE_FRACTION"])

    df["EMPLOYEE"] = df["Příjmení"].astype(str).str.strip() + " " + df["Jméno"].astype(str).str.strip()
    start = pd.to_datetime(df["Začátek - datum"]).dt.normalize()
    end = pd.to_datetime(df["Konec - datum"]).dt.normalize()
    pocet_dni = pd.to_numeric(df["Počet dní"], errors="coerce").fillna(1.0)

    rows = []
    for employee, s, e, nd in zip(df["EMPLOYEE"], start, end, pocet_dni):
        if pd.isna(s) or pd.isna(e):
            continue
        n_calendar_days = (e - s).days + 1
        if n_calendar_days <= 0:
            continue
        frac_per_day = min(1.0, nd / n_calendar_days)  # rozprostře Počet dní rovnoměrně přes rozsah (i půldenní)
        for i in range(n_calendar_days):
            rows.append({"EMPLOYEE": employee, "DATE": s + pd.Timedelta(days=i), "ABSENCE_FRACTION": frac_per_day})

    if not rows:
        return pd.DataFrame(columns=["EMPLOYEE", "DATE", "ABSENCE_FRACTION"])

    daily = pd.DataFrame(rows).groupby(["EMPLOYEE", "DATE"], as_index=False)["ABSENCE_FRACTION"].sum()
    daily["ABSENCE_FRACTION"] = daily["ABSENCE_FRACTION"].clip(upper=1.0)
    return daily


def compute_absence_daily(absences, employee_base):
    """Denní % nepřítomnosti vůči celkové známé základně zaměstnanců pobočky."""
    total_employees = len(employee_base)
    if total_employees == 0 or absences.empty:
        return pd.DataFrame(columns=["DATE", "N_ABSENT", "ABSENCE_PCT"])
    daily = (
        absences.groupby("DATE", as_index=False)["ABSENCE_FRACTION"].sum()
        .rename(columns={"ABSENCE_FRACTION": "N_ABSENT"})
        .sort_values("DATE")
    )
    daily["ABSENCE_PCT"] = daily["N_ABSENT"] / total_employees * 100
    return daily


def merge_activities_with_workspaces(activities, workspaces):
    """Připojí k aktivitám info o pobočce. Vrátí (spojená_data, aktivity_z_neznámé_pobočky)."""
    merged = activities.merge(workspaces, on="BRANCH_ID", how="left", indicator=True)
    unknown = merged.loc[merged["_merge"] == "left_only"].copy()
    known = merged.loc[merged["_merge"] == "both"].drop(columns="_merge").copy()
    return known.reset_index(drop=True), unknown.reset_index(drop=True)


def full_date_grid_for_branch(dates, include_weekends):
    """Vrátí kompletní řadu dnů pokrývající data JEDNÉ pobočky — pracovní dny Po-Pá,
    nebo všechny dny (podle toho, jestli je pobočka VIKENDOVA)."""
    dates = dates.dropna()
    if dates.empty:
        raise ValueError(
            "Nepodařilo se určit žádné platné datum aktivit po spojení s work_spaces.xlsx. "
            "Nejčastější příčina: všechny řádky v bo_data.xlsx patří pobočkám (BRANCH_ID), "
            "které nejsou ve work_spaces.xlsx — zkontrolujte hlášku 'Pozor: aktivity patří "
            "pobočkám...' o pár buněk výš a porovnejte BRANCH_ID v obou souborech."
        )
    start, end = dates.min(), dates.max()
    return pd.date_range(start, end) if include_weekends else pd.bdate_range(start, end)


def _branch_date_grids(merged, branch_cols):
    """Pro každou pobočku v `merged` sestaví kompletní řadu dnů (viz výše) a vrátí
    je jako jeden DataFrame BRANCH_ID×DATE — základ pro doplnění nulových dnů."""
    branches = merged[branch_cols].drop_duplicates("BRANCH_ID")
    frames = []
    for _, b in branches.iterrows():
        b_dates = merged.loc[merged["BRANCH_ID"] == b["BRANCH_ID"], "DATE"]
        grid = full_date_grid_for_branch(b_dates, bool(b["VIKENDOVA"]))
        frames.append(pd.DataFrame({"BRANCH_ID": b["BRANCH_ID"], "DATE": grid}))
    return branches, pd.concat(frames, ignore_index=True)


# -----------------------------------------------------------------------------
# 4. Výpočet vytíženosti
# -----------------------------------------------------------------------------

def utilization_bucket(pct):
    if pd.isna(pct):
        return "Bez dat"
    if pct >= THRESHOLD_CRITICAL:
        return "Kritická"
    if pct >= THRESHOLD_HIGH:
        return "Vysoká"
    return "Nízká"


def compute_workstation_daily(merged):
    """Denní vytíženost každého jednotlivého pracoviště na pobočce."""
    daily = (
        merged.groupby(["BRANCH_ID", "BRANCH_NAME", "WORKSTATION_ID", "DATE"])
        .agg(DURATION_MIN=("DURATION_MIN", "sum"), N_ACTIVITIES=("DURATION_MIN", "count"))
        .reset_index()
    )
    branches, date_grid = _branch_date_grids(merged, ["BRANCH_ID", "BRANCH_NAME", "CAPACITY_MIN", "VIKENDOVA"])
    workstations = merged[["BRANCH_ID", "WORKSTATION_ID"]].drop_duplicates()

    full_index = workstations.merge(date_grid, on="BRANCH_ID", how="left")
    full_index = full_index.merge(branches[["BRANCH_ID", "BRANCH_NAME", "CAPACITY_MIN"]], on="BRANCH_ID", how="left")

    out = full_index.merge(daily.drop(columns="BRANCH_NAME"), on=["BRANCH_ID", "WORKSTATION_ID", "DATE"], how="left")
    out["DURATION_MIN"] = out["DURATION_MIN"].fillna(0.0)
    out["N_ACTIVITIES"] = out["N_ACTIVITIES"].fillna(0).astype(int)
    out["UTILIZATION_PCT"] = np.where(out["CAPACITY_MIN"] > 0, out["DURATION_MIN"] / out["CAPACITY_MIN"] * 100, np.nan)
    out["BUCKET"] = out["UTILIZATION_PCT"].apply(utilization_bucket)
    return out.sort_values(["BRANCH_ID", "WORKSTATION_ID", "DATE"]).reset_index(drop=True)


def compute_branch_daily(merged):
    """Denní vytíženost celé pobočky (součet přes všechna pracoviště vs. registrovaná kapacita)."""
    branches, date_grid = _branch_date_grids(
        merged, ["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_DAY_HOURS", "CAPACITY_MIN", "VIKENDOVA"]
    )
    branches = branches.copy()
    branches["BRANCH_CAPACITY_MIN"] = branches["NO_WORKSTATIONS"] * branches["CAPACITY_MIN"]

    daily = (
        merged.groupby(["BRANCH_ID", "DATE"])
        .agg(DURATION_MIN=("DURATION_MIN", "sum"), N_ACTIVITIES=("DURATION_MIN", "count"))
        .reset_index()
    )
    out = date_grid.merge(branches, on="BRANCH_ID", how="left").merge(daily, on=["BRANCH_ID", "DATE"], how="left")
    out["DURATION_MIN"] = out["DURATION_MIN"].fillna(0.0)
    out["N_ACTIVITIES"] = out["N_ACTIVITIES"].fillna(0).astype(int)
    out["UTILIZATION_PCT"] = np.where(out["BRANCH_CAPACITY_MIN"] > 0, out["DURATION_MIN"] / out["BRANCH_CAPACITY_MIN"] * 100, np.nan)
    out["BUCKET"] = out["UTILIZATION_PCT"].apply(utilization_bucket)
    return out.sort_values(["BRANCH_ID", "DATE"]).reset_index(drop=True)


def summarize_branches(branch_daily, workspaces):
    """Souhrn na úrovni pobočky za celé sledované období + poznámka o pobočkách bez dat."""
    summary = (
        branch_daily.groupby(["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_DAY_HOURS"])
        .agg(
            PRUMERNA_VYTIZENOST_PCT=("UTILIZATION_PCT", "mean"),
            MAX_VYTIZENOST_PCT=("UTILIZATION_PCT", "max"),
            DNI_KRITICKA=("BUCKET", lambda s: (s == "Kritická").sum()),
            CELKEM_HODIN=("DURATION_MIN", lambda s: s.sum() / 60),
            POCET_DNI=("DATE", "nunique"),
        )
        .reset_index()
    )
    summary["BUCKET"] = summary["PRUMERNA_VYTIZENOST_PCT"].apply(utilization_bucket)

    branches_without_data = workspaces.loc[~workspaces["BRANCH_ID"].isin(summary["BRANCH_ID"])][
        ["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_DAY_HOURS"]
    ].copy()
    if not branches_without_data.empty:
        branches_without_data["PRUMERNA_VYTIZENOST_PCT"] = np.nan
        branches_without_data["MAX_VYTIZENOST_PCT"] = np.nan
        branches_without_data["DNI_KRITICKA"] = 0
        branches_without_data["CELKEM_HODIN"] = 0.0
        branches_without_data["POCET_DNI"] = 0
        branches_without_data["BUCKET"] = "Bez dat"
        summary = pd.concat([summary, branches_without_data], ignore_index=True)

    return summary.sort_values("PRUMERNA_VYTIZENOST_PCT", ascending=False, na_position="last").reset_index(drop=True)


def summarize_workstations(workstation_daily, segments=None):
    """Souhrn za celé sledované období pro každé jednotlivé pracoviště na pobočce —
    aby šlo na jeden pohled vidět, které pracoviště je vytížené a které ne, a na kolik %.
    Pokud je k dispozici převodník segmentů, připojí i sloupec SEGMENT (jinak „—")."""
    summary = (
        workstation_daily.groupby(["BRANCH_ID", "BRANCH_NAME", "WORKSTATION_ID"])
        .agg(
            PRUMERNA_VYTIZENOST_PCT=("UTILIZATION_PCT", "mean"),
            MAX_VYTIZENOST_PCT=("UTILIZATION_PCT", "max"),
            DNI_KRITICKA=("BUCKET", lambda s: (s == "Kritická").sum()),
            CELKEM_HODIN=("DURATION_MIN", lambda s: s.sum() / 60),
            POCET_DNI=("DATE", "nunique"),
        )
        .reset_index()
    )
    summary["BUCKET"] = summary["PRUMERNA_VYTIZENOST_PCT"].apply(utilization_bucket)

    if segments is not None and not segments.empty:
        summary = summary.merge(segments, on=["BRANCH_ID", "WORKSTATION_ID"], how="left")
    else:
        summary["SEGMENT"] = None
    summary["SEGMENT"] = summary["SEGMENT"].fillna("—")

    return summary.sort_values(
        ["BRANCH_ID", "PRUMERNA_VYTIZENOST_PCT"], ascending=[True, False]
    ).reset_index(drop=True)


def compute_capacity_growth_flags(merged, workspaces):
    """Poukazuje na pobočky, kde je v datech využíváno víc pracovišť, než je oficiálně
    registrováno ve work_spaces.xlsx (signál, že kapacitu pobočky je třeba navýšit v evidenci)."""
    used = merged.groupby(["BRANCH_ID", "BRANCH_NAME"])["WORKSTATION_ID"].nunique().reset_index(name="POUZITA_PRACOVISTE")
    out = used.merge(workspaces[["BRANCH_ID", "NO_WORKSTATIONS"]], on="BRANCH_ID", how="left")
    out["ROZDIL"] = out["POUZITA_PRACOVISTE"] - out["NO_WORKSTATIONS"]
    out["PREKROCENO"] = out["ROZDIL"] > 0
    return out.sort_values("ROZDIL", ascending=False).reset_index(drop=True)


def compute_activity_breakdown(merged):
    return (
        merged.groupby("ACTIVITY")
        .agg(CELKEM_HODIN=("DURATION_MIN", lambda s: s.sum() / 60), POCET=("DURATION_MIN", "count"))
        .reset_index()
        .sort_values("CELKEM_HODIN", ascending=False)
        .reset_index(drop=True)
    )


def compute_employee_summary(merged, top_n=None):
    out = (
        merged.groupby(["BRANCH_ID", "BRANCH_NAME", "EMPLOYEE"])
        .agg(
            CELKEM_HODIN=("DURATION_MIN", lambda s: s.sum() / 60),
            POCET_AKTIVIT=("DURATION_MIN", "count"),
            PRUMER_MIN=("DURATION_MIN", "mean"),
            POCET_DNI=("DATE", "nunique"),
        )
        .reset_index()
        .sort_values("CELKEM_HODIN", ascending=False)
        .reset_index(drop=True)
    )
    return out.head(top_n) if top_n else out


# -----------------------------------------------------------------------------
# 5. Grafy (Plotly)
# -----------------------------------------------------------------------------

PLOTLY_FONT = dict(family="system-ui, -apple-system, Segoe UI, sans-serif", color=TEXT_PRIMARY, size=13)


def _style_chart(fig, legend=False):
    """Sjednocený vzhled grafů: světlý povrch, jemné mřížky, tlumený text os,
    žádný zbytečný legend box pro grafy s jednou sérií."""
    fig.update_layout(
        font=PLOTLY_FONT,
        paper_bgcolor=SURFACE,
        plot_bgcolor=SURFACE,
        title_font=dict(size=15, color=TEXT_PRIMARY),
        showlegend=legend,
        margin=dict(l=10, r=10, t=50, b=40),
    )
    fig.update_xaxes(gridcolor=GRIDLINE, linecolor=GRIDLINE, tickfont=dict(color=TEXT_MUTED, size=11), title_font=dict(color=TEXT_SECONDARY))
    fig.update_yaxes(gridcolor=GRIDLINE, linecolor=GRIDLINE, tickfont=dict(color=TEXT_MUTED, size=11), title_font=dict(color=TEXT_SECONDARY))
    return fig


def fig_branch_utilization_bar(branch_summary):
    df = branch_summary.dropna(subset=["PRUMERNA_VYTIZENOST_PCT"]).sort_values("PRUMERNA_VYTIZENOST_PCT")
    colors = df["BUCKET"].map(BUCKET_COLORS).fillna(STATUS_MUTED)
    fig = go.Figure(go.Bar(
        x=df["PRUMERNA_VYTIZENOST_PCT"], y=df["BRANCH_NAME"] + " (" + df["BRANCH_ID"].astype(str) + ")",
        orientation="h", marker_color=colors, marker_line_width=0,
        text=df["PRUMERNA_VYTIZENOST_PCT"].round(1).astype(str) + " %", textposition="outside",
        textfont=dict(color=TEXT_SECONDARY, size=11),
        hovertemplate="%{y}<br>Průměrná vytíženost: %{x:.1f} %<extra></extra>",
    ))
    fig.add_vline(x=THRESHOLD_HIGH, line_dash="dot", line_color=STATUS_WARNING)
    fig.add_vline(x=THRESHOLD_CRITICAL, line_dash="dot", line_color=STATUS_CRITICAL)
    fig.add_vline(x=100, line_color=TEXT_MUTED)
    fig.update_layout(
        title="Průměrná denní vytíženost poboček vůči kapacitě", xaxis_title="Vytíženost (%)", yaxis_title=None,
        height=max(320, 28 * len(df) + 120), bargap=0.35,
    )
    return _style_chart(fig)


def fig_workstation_heatmap(workstation_daily, branch_id, branch_name):
    d = workstation_daily.loc[workstation_daily["BRANCH_ID"] == branch_id].copy()
    d["WORKSTATION_LABEL"] = "Prac. " + d["WORKSTATION_ID"].astype(str)
    pivot = d.pivot_table(index="DATE", columns="WORKSTATION_LABEL", values="UTILIZATION_PCT")
    pivot = pivot.reindex(sorted(pivot.columns, key=lambda c: int(c.split(" ")[1])), axis=1)

    # Přesně 0.0 % (žádná aktivita) se má vykreslit jako světle šedá/prázdná, ne zelená —
    # proto těsně nad nulou barva "skočí" na zelenou (epsilon = zlomek procenta z rozsahu).
    zero_eps = 0.3 / 150
    fig = go.Figure(go.Heatmap(
        z=pivot.values, x=pivot.columns, y=[d.strftime("%d.%m.%Y") for d in pivot.index],
        colorscale=[
            [0.0, GRIDLINE], [zero_eps, GRIDLINE],
            [zero_eps, STATUS_GOOD], [THRESHOLD_HIGH / 150, STATUS_WARNING],
            [THRESHOLD_CRITICAL / 150, STATUS_CRITICAL], [1.0, "#7a1414"],
        ],
        zmin=0, zmax=150, colorbar=dict(title="%", outlinewidth=0, tickfont=dict(color=TEXT_MUTED)),
        hovertemplate="%{x} | %{y}<br>Vytíženost: %{z:.1f} %<extra></extra>",
        xgap=2, ygap=2,
    ))
    fig.update_layout(
        title=f"Denní vytíženost pracovišť — {branch_name} ({branch_id})",
        height=max(260, 24 * len(pivot.index) + 120),
    )
    return _style_chart(fig)




def fig_activity_mix(activity_breakdown):
    fig = px.pie(
        activity_breakdown, names="ACTIVITY", values="CELKEM_HODIN", hole=0.55,
        color_discrete_sequence=CATEGORICAL, title="Skladba aktivit dle odpracovaných hodin",
    )
    # Popisek přímo u výseče (místo legendy) -- identita nezávisí na barvě/legendě
    # a graf tak nepřeteče do sousedního sloupce (vedle medailí zaměstnanců).
    fig.update_traces(
        textinfo="label+percent", textposition="outside",
        textfont=dict(color=TEXT_SECONDARY, size=11),
        marker=dict(line=dict(color=SURFACE, width=2)),
        hovertemplate="%{label}<br>%{value:.1f} h (%{percent})<extra></extra>",
    )
    fig.update_layout(height=380, showlegend=False, margin=dict(l=70, r=70, t=50, b=20))
    return _style_chart(fig, legend=False)


def fig_branch_daily_bar(branch_daily, branch_id, branch_name):
    """Denní vytíženost JEDNÉ pobočky (sloupcový graf po dnech) — obdoba "Room Utilization Rate"."""
    d = branch_daily.loc[branch_daily["BRANCH_ID"] == branch_id].sort_values("DATE")
    colors = d["BUCKET"].map(BUCKET_COLORS).fillna(STATUS_MUTED)
    fig = go.Figure(go.Bar(
        x=[dt.strftime("%d.%m.") for dt in d["DATE"]], y=d["UTILIZATION_PCT"],
        marker_color=colors, marker_line_width=0,
        text=d["UTILIZATION_PCT"].round(1).astype(str) + " %", textposition="outside",
        textfont=dict(color=TEXT_SECONDARY, size=11),
        hovertemplate="%{x}<br>Vytíženost: %{y:.1f} %<extra></extra>",
    ))
    fig.add_hline(y=THRESHOLD_HIGH, line_dash="dot", line_color=STATUS_WARNING)
    fig.add_hline(y=THRESHOLD_CRITICAL, line_dash="dot", line_color=STATUS_CRITICAL)
    fig.add_hline(y=100, line_color=TEXT_MUTED)
    fig.update_layout(
        title=f"Denní vytíženost pobočky vůči kapacitě — {branch_name}",
        xaxis_title=None, yaxis_title="Vytíženost (%)",
        height=360, bargap=0.35,
    )
    return _style_chart(fig)


# --- Pilotní rozšíření: nepřítomnost zaměstnanců a rozdělení na budovy ------

def fig_absence_daily(absence_daily, branch_name, total_employees):
    """Denní % nepřítomnosti (dovolená, nemoc, home office, ...) zaměstnanců pobočky."""
    fig = go.Figure(go.Bar(
        x=absence_daily["DATE"], y=absence_daily["ABSENCE_PCT"],
        marker_color=STATUS_WARNING, marker_line_width=0,
        hovertemplate="%{x|%d.%m.%Y}<br>%{y:.0f} %% nepřítomných<extra></extra>",
    ))
    fig.update_layout(
        title=f"Denní nepřítomnost zaměstnanců — {branch_name} (základna: {total_employees} zam.)",
        xaxis_title=None, yaxis_title="% nepřítomných", height=300, bargap=0.35,
    )
    return _style_chart(fig)


def compute_building_capacity_daily(merged, branch_id, workstation_range, capacity_min_per_workstation, absence_daily):
    """Pro jednu budovu pilotní pobočky (podmnožina WORKSTATION_ID) spočítá denně:
    surovou kapacitu (počet pracovišť budovy × kapacita/pracoviště/den), efektivní
    kapacitu sníženou o poměrný podíl nepřítomných zaměstnanců (dle podílu pracovišť
    této budovy na celkovém počtu pracovišť pobočky) a skutečně odpracované minuty."""
    workstation_ids = list(workstation_range)
    n_workstations = len(workstation_ids)
    raw_capacity_min = n_workstations * capacity_min_per_workstation

    b = merged.loc[(merged["BRANCH_ID"] == branch_id) & (merged["WORKSTATION_ID"].isin(workstation_ids))]
    dates = pd.bdate_range(b["DATE"].min(), b["DATE"].max()) if not b.empty else pd.DatetimeIndex([])
    usage_by_date = b.groupby("DATE")["DURATION_MIN"].sum() if not b.empty else pd.Series(dtype=float)

    total_ws_all_buildings = sum(len(list(r)) for r in PILOT_BUILDINGS.values())
    building_share = n_workstations / total_ws_all_buildings if total_ws_all_buildings else 0.0
    absence_by_date = absence_daily.set_index("DATE")["N_ABSENT"] if not absence_daily.empty else pd.Series(dtype=float)

    out = []
    for d in dates:
        n_absent_building = absence_by_date.get(d, 0.0) * building_share
        effective_capacity_min = max(0.0, raw_capacity_min - n_absent_building * capacity_min_per_workstation)
        out.append({
            "DATE": d,
            "RAW_CAPACITY_MIN": raw_capacity_min,
            "EFFECTIVE_CAPACITY_MIN": effective_capacity_min,
            "ACTUAL_MIN": usage_by_date.get(d, 0.0),
        })
    return pd.DataFrame(out)


def fig_building_capacity(building_daily, building_name, n_workstations):
    """Srovnání surové vs. efektivní (nepřítomností snížené) kapacity budovy se
    skutečným využitím — ukazuje, jestli je nízká vytíženost kvůli nepřítomným
    zaměstnancům, nebo jde o skutečně volnou kapacitu pracovišť."""
    if building_daily.empty:
        return None
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=building_daily["DATE"], y=building_daily["RAW_CAPACITY_MIN"] / 60,
        name="Kapacita (nainstalovaná)", marker_color=GRIDLINE,
    ))
    fig.add_trace(go.Bar(
        x=building_daily["DATE"], y=building_daily["EFFECTIVE_CAPACITY_MIN"] / 60,
        name="Efektivní kapacita (po odečtení nepřítomných)", marker_color=STATUS_WARNING, opacity=0.6,
    ))
    fig.add_trace(go.Scatter(
        x=building_daily["DATE"], y=building_daily["ACTUAL_MIN"] / 60,
        name="Skutečně odpracováno", mode="lines+markers",
        line=dict(color=STATUS_GOOD, width=2), marker=dict(size=6),
    ))
    fig.update_layout(
        title=f"{building_name} — {n_workstations} pracovišť: kapacita vs. skutečnost",
        barmode="overlay", xaxis_title=None, yaxis_title="hodiny / den", height=340,
        legend=dict(orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5),
    )
    return _style_chart(fig, legend=True)


def fig_workstation_time_profile(merged, branch_daily, branch_id, branch_name, block_minutes=BLOCK_MINUTES):
    """Vytíženost pracovišť podle času dne za CELÉ sledované období — stejný rastr
    jako denní rozvrh po blocích, ale místo jednoho dne ukazuje četnost napříč všemi
    dny: barva = v kolika % otevřených dní bylo dané pracoviště v daném čase obsazené
    (od průhledné/prázdné = nikdy, po tmavě modrou = prakticky vždy)."""
    d = merged.loc[merged["BRANCH_ID"] == branch_id].copy()
    ws_ids = sorted(d["WORKSTATION_ID"].unique())
    if not ws_ids or d.empty:
        return None

    total_days = branch_daily.loc[branch_daily["BRANCH_ID"] == branch_id, "DATE"].nunique()
    if total_days == 0:
        return None

    d["TOD_START"] = d["DATETIME"].dt.hour * 60 + d["DATETIME"].dt.minute
    d["TOD_END"] = d["TOD_START"] + d["DURATION_MIN"]

    day_start = int(np.floor(d["TOD_START"].min() / block_minutes) * block_minutes)
    day_end = int(np.ceil(d["TOD_END"].max() / block_minutes) * block_minutes)
    if day_end - day_start < 120:
        day_end = day_start + 120

    n_blocks = (day_end - day_start) // block_minutes
    block_labels = []
    for i in range(n_blocks):
        m = day_start + i * block_minutes
        block_labels.append(f"{(m // 60) % 24:02d}:{m % 60:02d}")

    ws_row = {ws: i for i, ws in enumerate(ws_ids)}
    occ = defaultdict(set)
    for r in d.itertuples(index=False):
        row_i = ws_row[r.WORKSTATION_ID]
        c0 = max(0, int((r.TOD_START - day_start) // block_minutes))
        c1 = min(n_blocks, math.ceil((r.TOD_END - day_start) / block_minutes))
        for c in range(c0, c1):
            occ[(row_i, c)].add(r.DATE)

    z = np.zeros((len(ws_ids), n_blocks))
    for (row_i, c), dates in occ.items():
        z[row_i, c] = len(dates) / total_days * 100

    tick_step = max(1, 60 // block_minutes)
    tick_idx = list(range(0, n_blocks, tick_step))

    fig = go.Figure(go.Heatmap(
        z=z, x=block_labels, y=[f"Prac. {w}" for w in ws_ids],
        zmin=0, zmax=100,
        colorscale=[[0.0, "rgba(42,120,214,0)"], [1.0, "rgba(42,120,214,1)"]],
        colorbar=dict(title="% dní", outlinewidth=0, tickfont=dict(color=TEXT_MUTED)),
        hovertemplate="%{y} · %{x}<br>Obsazeno %{z:.0f} % otevřených dní<extra></extra>",
        xgap=1, ygap=2,
    ))
    fig.update_xaxes(tickmode="array", tickvals=[block_labels[i] for i in tick_idx], title=None)
    fig.update_yaxes(autorange="reversed", title=None)
    fig.update_layout(
        title=f"Vytíženost pracovišť podle času dne — {branch_name} (celé období, {total_days} dní)",
        height=max(260, 34 * len(ws_ids) + 140),
    )
    return _style_chart(fig)


def _discrete_colorscale(colors):
    """Vrátí Plotly colorscale, kde celočíselná hodnota z=i (i=0..len(colors)-1)
    vždy vyjde přesně na barvu colors[i] — pro "kategoriální" heatmapu."""
    n = len(colors)
    scale = []
    for i, color in enumerate(colors):
        scale.append([i / n, color])
        scale.append([(i + 1) / n, color])
    return scale


def fig_workstation_day_blocks(merged, branch_id, branch_name, date, poledni_pauza, block_minutes=BLOCK_MINUTES):
    """Rozvrh pracovišť pro JEDEN den v blocích po `block_minutes` minutách —
    řádky jsou pracoviště (od nejmenšího čísla nahoře), barva bloku = typ aktivity.
    Polední pauza (pokud POLEDNI_PAUZA=True u pobočky) je vyznačená šedým pásem."""
    d_branch = merged.loc[merged["BRANCH_ID"] == branch_id]
    ws_ids = sorted(d_branch["WORKSTATION_ID"].unique())
    d = d_branch.loc[d_branch["DATE"] == date].copy()
    if not ws_ids or d.empty:
        return None

    day_start = d["DATETIME"].min().floor(f"{block_minutes}min")
    day_end = d["END_DATETIME"].max().ceil(f"{block_minutes}min")
    if day_end - day_start < timedelta(hours=2):
        day_end = day_start + timedelta(hours=2)

    n_blocks = int((day_end - day_start) / timedelta(minutes=block_minutes))
    block_starts = [day_start + timedelta(minutes=block_minutes * i) for i in range(n_blocks)]

    activities_present = sorted(d["ACTIVITY"].unique())
    activity_code = {a: i + 1 for i, a in enumerate(activities_present)}
    activity_color = {a: CATEGORICAL[i % len(CATEGORICAL)] for i, a in enumerate(activities_present)}

    ws_row = {ws: i for i, ws in enumerate(ws_ids)}
    z = np.zeros((len(ws_ids), n_blocks))
    hover = np.full((len(ws_ids), n_blocks), "", dtype=object)

    for _, r in d.iterrows():
        row_i = ws_row[r["WORKSTATION_ID"]]
        c0 = max(0, int((r["DATETIME"] - day_start) / timedelta(minutes=block_minutes)))
        c1 = min(n_blocks, math.ceil((r["END_DATETIME"] - day_start) / timedelta(minutes=block_minutes)))
        code = activity_code[r["ACTIVITY"]]
        label = f"{r['ACTIVITY']}<br>{r['EMPLOYEE']}<br>{r['DATETIME']:%H:%M}–{r['END_DATETIME']:%H:%M}"
        for c in range(c0, c1):
            z[row_i, c] = code
            hover[row_i, c] = label

    colors = [GRIDLINE] + [activity_color[a] for a in activities_present]

    fig = go.Figure(go.Heatmap(
        z=z, x=block_starts, y=[f"Prac. {w}" for w in ws_ids],
        zmin=0, zmax=len(colors), colorscale=_discrete_colorscale(colors), showscale=False,
        text=hover, hoverinfo="text", xgap=1, ygap=2,
    ))

    if poledni_pauza:
        lunch_start = day_start.normalize() + pd.Timedelta(hours=LUNCH_BREAK_START_HOUR)
        lunch_end = day_start.normalize() + pd.Timedelta(hours=LUNCH_BREAK_END_HOUR)
        if lunch_start < day_end and lunch_end > day_start:
            fig.add_vrect(
                x0=max(lunch_start, day_start), x1=min(lunch_end, day_end),
                fillcolor=TEXT_MUTED, opacity=0.22, line_width=0,
                annotation_text="polední pauza", annotation_position="top",
                annotation_font=dict(size=9, color=TEXT_MUTED),
            )

    # heatmapa nemá vlastní kategoriální legendu -> "fantomové" body jen pro popisky v legendě
    for a in activities_present:
        fig.add_trace(go.Scatter(
            x=[block_starts[0]], y=[f"Prac. {ws_ids[0]}"], mode="markers",
            marker=dict(size=0.001, color=activity_color[a]), name=a, showlegend=True, hoverinfo="skip",
        ))

    fig.update_yaxes(autorange="reversed", title=None)
    fig.update_xaxes(title=None, tickformat="%H:%M")
    fig.update_layout(
        title=f"{branch_name} — {pd.Timestamp(date):%d.%m.%Y}",
        height=max(220, 34 * len(ws_ids) + 140),
        legend_title_text="Aktivita",
    )
    return _style_chart(fig, legend=True)


# -----------------------------------------------------------------------------
# 6. Sestavení HTML reportu
# -----------------------------------------------------------------------------

_CSS = """
:root { color-scheme: light; }
* { box-sizing: border-box; }
body {
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
    background: #f9f9f7; color: #0b0b0b; margin: 0; padding: 0 0 60px 0; font-size: 14px;
}
.wrap { max-width: 1220px; margin: 0 auto; padding: 0 24px; }
header.page { padding: 22px 0 4px 0; }
header.page .title-row { display: flex; align-items: baseline; justify-content: space-between; flex-wrap: wrap; gap: 10px; }
header.page h1 { margin: 0; font-size: 20px; font-weight: 700; color: #0b0b0b; }
header.page p { margin: 2px 0 0 0; color: #898781; font-size: 12.5px; }
h2 { font-size: 15px; font-weight: 700; color: #0b0b0b; border: none; padding-bottom: 0; margin: 0 0 14px 0; }
h3 { font-size: 13.5px; font-weight: 700; color: #0b0b0b; margin: 18px 0 4px 0; }
.card-row { display: flex; gap: 14px; flex-wrap: wrap; margin: 14px 0 4px 0; }
.card {
    background: #fcfcfb; border-radius: 12px; padding: 16px 18px; flex: 1 1 180px;
    box-shadow: 0 1px 2px rgba(11,11,11,0.05); border: 1px solid rgba(11,11,11,0.08);
    position: relative; overflow: hidden; min-width: 0;
}
.card .label { font-size: 11px; color: #898781; text-transform: uppercase; letter-spacing: .04em; display: flex; align-items: center; gap: 5px; }
.card .value { font-size: 24px; font-weight: 600; margin-top: 5px; color: #0b0b0b; }
.card .delta { font-size: 12px; font-weight: 600; margin-top: 3px; display: inline-flex; align-items: center; gap: 2px; }
.card .delta.up { color: #006300; }
.card .delta.down { color: #d03b3b; }
.card .spark { position: absolute; right: 0; bottom: 0; opacity: 0.9; }
.info-dot {
    display: inline-flex; align-items: center; justify-content: center; width: 13px; height: 13px;
    border-radius: 50%; border: 1px solid #c3c2b7; color: #898781; font-size: 9px; cursor: help;
}
.section { background: #fcfcfb; border-radius: 12px; padding: 20px 22px; margin-top: 16px;
    box-shadow: 0 1px 2px rgba(11,11,11,0.05); border: 1px solid rgba(11,11,11,0.08); }
table.report { border-collapse: collapse; width: 100%; font-size: 12.5px; margin-top: 10px; }
table.report th {
    background: transparent; color: #898781; text-align: left; padding: 8px 10px;
    font-size: 11px; text-transform: uppercase; letter-spacing: .03em; font-weight: 600;
    border-bottom: 1px solid #e1e0d9; position: sticky; top: 0; background-color: #fcfcfb;
}
table.report td { padding: 8px 10px; border-bottom: 1px solid #e1e0d9; color: #0b0b0b; }
table.report tr:hover td { background: #f9f9f7; }
.badge { display: inline-flex; align-items: center; gap: 5px; padding: 2px 9px 2px 6px; border-radius: 12px; font-size: 11.5px; font-weight: 600; color: #0b0b0b; background: #f0efec; }
.badge .dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; }
.table-scroll { max-height: 480px; overflow-y: auto; }
details.branch-block { margin-bottom: 14px; border: 1px solid #e1e0d9; border-radius: 10px; }
details.branch-block summary { padding: 12px 16px; cursor: pointer; font-weight: 600; background: #f9f9f7; border-radius: 10px; list-style: none; }
details.branch-block summary::-webkit-details-marker { display: none; }
details.branch-block[open] summary { border-radius: 10px 10px 0 0; }
details.branch-block > div.branch-body { padding: 18px 20px; }
.note { font-size: 12px; color: #898781; margin-top: 6px; line-height: 1.5; }
footer { text-align: center; color: #898781; font-size: 11.5px; margin-top: 50px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; margin-right: 18px; font-size: 12px; color: #52514e; }
.legend i { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
.branch-summary-line { font-size: 12.5px; color: #898781; font-weight: 400; margin-left: 8px; }
.no-data-panel { padding: 40px 20px; text-align: center; color: #898781; }

/* Plotly grafy se mají přizpůsobit šířce svého (flex) rodiče, ne mít pevnou šířku
   -- jinak se sousední grafy (skladba aktivit / zaměstnanci) překrývaly. */
.js-plotly-plot, .plotly-graph-div { width: 100% !important; }

/* Medailové karty (top 3 zaměstnanci) */
.medal-row { display: flex; gap: 14px; flex-wrap: wrap; margin-top: 10px; }
.medal-card {
    background: #fcfcfb; border: 1px solid rgba(11,11,11,0.08); border-radius: 12px;
    padding: 16px; flex: 1 1 140px; text-align: center;
}
.medal-icon { font-size: 30px; line-height: 1; }
.medal-name { font-weight: 600; font-size: 13px; margin-top: 8px; }
.medal-value { font-size: 20px; font-weight: 700; margin-top: 2px; }
.medal-sub { font-size: 11.5px; color: #898781; margin-top: 2px; }

/* Progress bary (vytíženost podle dne v týdnu) */
.table-scroll-x { overflow-x: auto; }
table.progress-table { min-width: 640px; }
table.progress-table td.ws-cell { font-weight: 600; white-space: nowrap; }
table.progress-table td.closed-cell { color: #c3c2b7; font-size: 11.5px; text-align: center; font-style: italic; }
.progress-cell { display: flex; align-items: center; gap: 8px; min-width: 90px; }
.progress-cell.over .progress-track { border: 1px solid #d03b3b; }
.progress-track { flex: 1 1 auto; height: 8px; border-radius: 4px; background: #e1e0d9; overflow: hidden; }
.progress-fill { height: 100%; border-radius: 4px; }
.progress-label { font-size: 11.5px; color: #52514e; white-space: nowrap; min-width: 34px; text-align: right; }

/* Pilotní rozšíření: budovy (nepřítomnost a kapacita) */
.building-block { margin-top: 18px; padding: 14px 16px; border: 1px solid #e1e0d9; border-radius: 10px; }
.building-block h4 { margin: 0 0 6px 0; font-size: 14px; }

/* Přepínač týdnů (denní rozvrh po blocích) */
.week-nav-wrap { margin-top: 10px; }
.week-nav-controls { display: flex; align-items: center; justify-content: center; gap: 14px; margin-bottom: 12px; }
.week-nav-controls button {
    font-size: 13px; padding: 7px 14px; border-radius: 18px; border: 1px solid #c3c2b7;
    background: #fcfcfb; cursor: pointer; color: #0b0b0b;
}
.week-nav-controls button:disabled { opacity: 0.35; cursor: default; }
.week-nav-controls button:not(:disabled):hover { background: #f0efec; }
.week-label { font-weight: 600; font-size: 13.5px; min-width: 200px; text-align: center; }
.week-page { display: none; }
.week-page.active { display: block; }
.week-day-block { margin-bottom: 14px; border: 1px solid #e1e0d9; border-radius: 8px; padding: 10px; }
.week-day-block.closed { padding: 10px 14px; }
.week-day-heading { font-weight: 600; font-size: 12.5px; color: #52514e; margin-bottom: 6px; }
.week-day-block.closed .week-day-heading { color: #c3c2b7; font-style: italic; margin-bottom: 0; }
"""


def _bucket_badge(bucket):
    color = BUCKET_COLORS.get(bucket, STATUS_MUTED)
    return f'<span class="badge"><i class="dot" style="background:{color}"></i>{bucket}</span>'


def _df_to_html_table(df, bucket_col=None, float_cols=None):
    d = df.copy()
    float_cols = float_cols or {}
    for col, fmt in float_cols.items():
        if col in d.columns:
            d[col] = d[col].map(lambda v: (fmt.format(v) if pd.notna(v) else "—"))
    if bucket_col and bucket_col in d.columns:
        d[bucket_col] = d[bucket_col].map(_bucket_badge)
    return d.to_html(index=False, escape=False, classes="report", border=0)


def _sparkline_svg(values, color, width=110, height=32):
    """Malý inline SVG sparkline (bez závislosti na Plotly) pro KPI dlaždice —
    tenká 2px linie + jemná výplň ~12% pod čarou, poslední bod zvýrazněný."""
    vals = [v for v in values if pd.notna(v)]
    if len(vals) < 2:
        return ""
    lo, hi = min(vals), max(vals)
    span = (hi - lo) or 1.0
    pad = 3
    n = len(vals)
    xs = [pad + i * (width - 2 * pad) / (n - 1) for i in range(n)]
    ys = [height - pad - (v - lo) / span * (height - 2 * pad) for v in vals]
    line_points = " ".join(f"{x:.1f},{y:.1f}" for x, y in zip(xs, ys))
    area_points = f"{xs[0]:.1f},{height} " + line_points + f" {xs[-1]:.1f},{height}"
    return f"""<svg class="spark" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
    <polyline points="{area_points}" fill="{color}" fill-opacity="0.12" stroke="none" />
    <polyline points="{line_points}" fill="none" stroke="{color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" />
    <circle cx="{xs[-1]:.1f}" cy="{ys[-1]:.1f}" r="3" fill="{color}" stroke="{SURFACE}" stroke-width="1.5" />
    </svg>"""


def _trend_delta_pct(values):
    """Změna v % mezi první a druhou polovinou období — pohání šipku/barvu na KPI dlaždici."""
    vals = [v for v in values if pd.notna(v)]
    if len(vals) < 4:
        return None
    mid = len(vals) // 2
    first_half, second_half = vals[:mid], vals[mid:]
    first_avg = sum(first_half) / len(first_half)
    second_avg = sum(second_half) / len(second_half)
    if first_avg == 0:
        return None
    return (second_avg - first_avg) / first_avg * 100


def _stat_tile(label, value, *, delta_pct=None, spark_values=None, spark_color=BLUE, tooltip=None):
    info = f'<span class="info-dot" title="{tooltip}">?</span>' if tooltip else ""
    delta_html = ""
    if delta_pct is not None:
        direction = "up" if delta_pct >= 0 else "down"
        arrow = "▲" if delta_pct >= 0 else "▼"
        delta_html = f'<div class="delta {direction}">{arrow} {abs(delta_pct):.1f} %</div>'
    spark_html = _sparkline_svg(spark_values, spark_color) if spark_values is not None else ""
    return f"""<div class="card">
      <div class="label">{label}{info}</div>
      <div class="value">{value}</div>
      {delta_html}
      {spark_html}
    </div>"""


_MEDALS = ["🥇", "🥈", "🥉"]


def _employee_medals_html(employee_summary, n=3):
    """Top N nejvytíženějších zaměstnanců jako medailové karty (místo grafu),
    aby se sekce nepřekrývala se sousedním koláčovým grafem skladby aktivit."""
    d = employee_summary.sort_values("CELKEM_HODIN", ascending=False).head(n)
    if d.empty:
        return "<p class=\"note\">Žádní zaměstnanci k zobrazení.</p>"
    cards = []
    for medal, (_, row) in zip(_MEDALS, d.iterrows()):
        cards.append(f"""<div class="medal-card">
          <div class="medal-icon">{medal}</div>
          <div class="medal-name">{row['EMPLOYEE']}</div>
          <div class="medal-value">{row['CELKEM_HODIN']:.1f} h</div>
          <div class="medal-sub">{int(row['POCET_AKTIVIT'])} aktivit</div>
        </div>""")
    return f'<div class="medal-row">{"".join(cards)}</div>'


def _progress_bar_html(pct, color):
    fill_pct = min(100, max(0, pct)) if pd.notna(pct) else 0
    label = f"{pct:.0f} %" if pd.notna(pct) else "—"
    over_class = " over" if pd.notna(pct) and pct > 100 else ""
    return f"""<div class="progress-cell{over_class}">
      <div class="progress-track"><div class="progress-fill" style="width:{fill_pct:.0f}%; background:{color}"></div></div>
      <span class="progress-label">{label}</span>
    </div>"""


WEEKDAY_LABELS = ["Po", "Út", "St", "Čt", "Pá", "So", "Ne"]


def build_weekday_progress_table_html(workstation_daily, branch_id):
    """Tabulka pracovišť s progress bary — průměrná vytíženost za jednotlivé dny
    v týdnu (Po–Ne), zprůměrovaná přes všechny výskyty daného dne v období."""
    d = workstation_daily.loc[workstation_daily["BRANCH_ID"] == branch_id].copy()
    if d.empty:
        return "<p class=\"note\">Žádná data.</p>"
    d["WEEKDAY"] = d["DATE"].dt.weekday
    pivot = d.groupby(["WORKSTATION_ID", "WEEKDAY"])["UTILIZATION_PCT"].mean().unstack("WEEKDAY")

    rows_html = []
    for ws in sorted(pivot.index):
        cells = [f'<td class="ws-cell">Prac. {ws}</td>']
        for wd in range(7):
            if wd in pivot.columns and pd.notna(pivot.loc[ws, wd]):
                pct = pivot.loc[ws, wd]
                color = BUCKET_COLORS.get(utilization_bucket(pct), STATUS_MUTED)
                cells.append(f"<td>{_progress_bar_html(pct, color)}</td>")
            else:
                cells.append('<td class="closed-cell">zavřeno</td>')
        rows_html.append(f"<tr>{''.join(cells)}</tr>")

    header = "<th>Pracoviště</th>" + "".join(f"<th>{lbl}</th>" for lbl in WEEKDAY_LABELS)
    return f"""<div class="table-scroll-x">
    <table class="report progress-table">
      <thead><tr>{header}</tr></thead>
      <tbody>{"".join(rows_html)}</tbody>
    </table>
    </div>"""


def build_pilot_absence_section(merged, branch_id, branch_name, capacity_min_per_workstation, absences, fig_html):
    """Pilotní rozšíření JEN pro PILOT_BRANCH_ID: nepřítomnost zaměstnanců a rozdělení
    na dvě budovy (Jugoslávská 1-4 zatím nenainstalováno, Olbrachtova 5-20)."""
    b_merged = merged.loc[merged["BRANCH_ID"] == branch_id]
    employee_base = set(b_merged["EMPLOYEE"].unique()) | (set(absences["EMPLOYEE"].unique()) if not absences.empty else set())
    absence_daily = compute_absence_daily(absences, employee_base)

    absence_fig_html = (
        fig_html(fig_absence_daily(absence_daily, branch_name, len(employee_base)))
        if not absence_daily.empty else '<p class="note">Žádná data o nepřítomnosti pro tuto pobočku/období.</p>'
    )

    building_sections = []
    for name, ws_range in PILOT_BUILDINGS.items():
        b_daily = compute_building_capacity_daily(merged, branch_id, ws_range, capacity_min_per_workstation, absence_daily)
        fig = fig_building_capacity(b_daily, name, len(list(ws_range)))
        if fig is None:
            building_sections.append(f"""
        <div class="building-block">
          <h4>{name} (pracoviště {min(ws_range)}–{max(ws_range)})</h4>
          <p class="note">Zatím žádná data — pracoviště v této budově nejsou v evidenci aktivní.</p>
        </div>""")
        else:
            building_sections.append(f"""
        <div class="building-block">
          <h4>{name} (pracoviště {min(ws_range)}–{max(ws_range)})</h4>
          {fig_html(fig)}
        </div>""")

    return f"""
        <h3 style="margin-top:26px">🧪 Pilotní rozšíření: nepřítomnost a rozdělení na budovy</h3>
        <p class="note">Základna pro výpočet % nepřítomnosti: {len(employee_base)} zaměstnanců (sjednocení jmen
        z aktivit a z evidence nepřítomností). Nepřítomnost zahrnuje dovolenou, nemoc, home office i další typy —
        všechny znamenají, že zaměstnanec ten den nemohl použít žádné pracoviště. Efektivní kapacita = kapacita
        pracovišť budovy snížená o poměrný podíl nepřítomných zaměstnanců.</p>
        {absence_fig_html}
        {"".join(building_sections)}
    """


def build_html_report(
    output_path, *, period_start, period_end, workspaces, merged,
    branch_summary, branch_daily, workstation_daily, growth_flags, segments,
    absences=None,
):
    generated_at = datetime.now().strftime("%d.%m.%Y %H:%M")

    overall_util = branch_daily["UTILIZATION_PCT"].mean()
    total_hours = merged["DURATION_MIN"].sum() / 60
    n_branches_with_data = merged["BRANCH_ID"].nunique()
    n_branches_total = workspaces["BRANCH_ID"].nunique()
    n_workstations_registered = int(workspaces["NO_WORKSTATIONS"].sum())
    n_employees = merged["EMPLOYEE"].nunique()

    daily_all = branch_daily.groupby("DATE").agg(UTILIZATION_PCT=("UTILIZATION_PCT", "mean"), DURATION_MIN=("DURATION_MIN", "sum")).sort_index()
    util_series = daily_all["UTILIZATION_PCT"].tolist()
    hours_series = (daily_all["DURATION_MIN"] / 60).tolist()

    cards = f"""
    <div class="card-row">
      {_stat_tile("Sledované období", f'{period_start:%d.%m.%Y} – {period_end:%d.%m.%Y}')}
      {_stat_tile(
          "Průměrná vytíženost poboček", f"{overall_util:.1f} %",
          delta_pct=_trend_delta_pct(util_series), spark_values=util_series, spark_color=BLUE,
          tooltip="Průměr denní vytíženosti přes všechny pobočky s daty",
      )}
      {_stat_tile(
          "Odpracované hodiny celkem", f"{total_hours:,.0f} h",
          delta_pct=_trend_delta_pct(hours_series), spark_values=hours_series, spark_color=CATEGORICAL[1],
          tooltip="Součet DURATION přes všechny aktivity ve sledovaném období",
      )}
      {_stat_tile("Pobočky s daty / celkem", f"{n_branches_with_data} / {n_branches_total}")}
      {_stat_tile("Registrovaná pracoviště", f"{n_workstations_registered}")}
      {_stat_tile("Aktivních zaměstnanců", f"{n_employees}")}
    </div>
    """

    plotly_cdn_included = False

    def fig_html(fig):
        # Plotly.js se vloží celý přímo do souboru (ne přes CDN), aby report šel
        # otevřít i bez internetu (e-mail, sdílený disk, offline prohlížení).
        # responsive=True + layout autosize -> graf se přizpůsobí šířce svého
        # flex sloupce místo pevné šířky, která způsobovala překryv sousedních grafů.
        nonlocal plotly_cdn_included
        if fig is None:
            return ""
        fig.update_layout(autosize=True)
        html = fig.to_html(
            full_html=False, include_plotlyjs=(True if not plotly_cdn_included else False),
            config={"displaylogo": False, "responsive": True},
            default_width="100%", default_height=None,
        )
        plotly_cdn_included = True
        return html

    # --- Přehled všech poboček (jediné, co zůstává na titulní straně) ----------
    branch_bar_html = fig_html(fig_branch_utilization_bar(branch_summary))

    branch_table_html = _df_to_html_table(
        branch_summary.rename(columns={
            "BRANCH_NAME": "Pobočka", "BRANCH_ID": "ID", "NO_WORKSTATIONS": "Poč. pracovišť",
            "CAPACITY_DAY_HOURS": "Kapacita (h/den)", "PRUMERNA_VYTIZENOST_PCT": "Prům. vytíženost",
            "MAX_VYTIZENOST_PCT": "Max. vytíženost", "DNI_KRITICKA": "Dní kriticky vytíž.",
            "CELKEM_HODIN": "Odprac. hodin", "POCET_DNI": "Dní s daty", "BUCKET": "Stav",
        }),
        bucket_col="Stav",
        float_cols={
            "Kapacita (h/den)": "{:.1f}", "Prům. vytíženost": "{:.1f} %",
            "Max. vytíženost": "{:.1f} %", "Odprac. hodin": "{:.1f}",
        },
    )

    # --- Rozbalovací karta pro každou pobočku -----------------------------------
    workstation_summary = summarize_workstations(workstation_daily, segments)
    growth_by_branch = growth_flags.set_index("BRANCH_ID")
    branches_with_data = set(merged["BRANCH_ID"].unique())

    branch_order = branch_summary["BRANCH_ID"].tolist()  # už seřazeno dle vytíženosti, "Bez dat" naposled
    branch_blocks = []

    for branch_id in branch_order:
        row = branch_summary.loc[branch_summary["BRANCH_ID"] == branch_id].iloc[0]
        branch_name = row["BRANCH_NAME"]
        has_data = branch_id in branches_with_data

        if not has_data:
            branch_blocks.append(f"""
        <details class="branch-block">
          <summary>{branch_name} ({branch_id}) <span class="branch-summary-line">bez dat</span></summary>
          <div class="branch-body no-data-panel">Pro pobočku <strong>{branch_name} ({branch_id})</strong>
          nejsou v bo_data.xlsx zaznamenané žádné aktivity.</div>
        </details>""")
            continue

        b_merged = merged.loc[merged["BRANCH_ID"] == branch_id]
        b_ws_summary = workstation_summary.loc[workstation_summary["BRANCH_ID"] == branch_id]
        b_activity_breakdown = compute_activity_breakdown(b_merged)
        b_employee_summary = compute_employee_summary(b_merged)
        b_growth = growth_by_branch.loc[branch_id] if branch_id in growth_by_branch.index else None
        poledni_pauza = bool(b_merged["POLEDNI_PAUZA"].iloc[0])

        n_activities = len(b_merged)
        b_hours = b_merged["DURATION_MIN"].sum() / 60
        b_used_ws = int(b_merged["WORKSTATION_ID"].nunique())
        b_registered_ws = int(row["NO_WORKSTATIONS"])

        b_daily = branch_daily.loc[branch_daily["BRANCH_ID"] == branch_id].sort_values("DATE")
        b_activities_series = b_daily["N_ACTIVITIES"].tolist()
        b_hours_series = (b_daily["DURATION_MIN"] / 60).tolist()
        b_util_series = b_daily["UTILIZATION_PCT"].tolist()

        kpi_html = f"""
        <div class="card-row">
          {_stat_tile(
              "Celkem aktivit", f"{n_activities:,}",
              delta_pct=_trend_delta_pct(b_activities_series), spark_values=b_activities_series, spark_color=BLUE,
          )}
          {_stat_tile(
              "Celkem hodin", f"{b_hours:.1f} h",
              delta_pct=_trend_delta_pct(b_hours_series), spark_values=b_hours_series, spark_color=CATEGORICAL[1],
          )}
          {_stat_tile(
              "Průměrná vytíženost", f"{row['PRUMERNA_VYTIZENOST_PCT']:.1f} %",
              delta_pct=_trend_delta_pct(b_util_series), spark_values=b_util_series, spark_color=CATEGORICAL[2],
          )}
          {_stat_tile("Max. denní vytíženost", f"{row['MAX_VYTIZENOST_PCT']:.1f} %")}
          {_stat_tile("Využitá / registrovaná pracoviště", f"{b_used_ws} / {b_registered_ws}")}
        </div>
        """

        ws_table_html = _df_to_html_table(
            b_ws_summary.rename(columns={
                "WORKSTATION_ID": "Pracoviště", "SEGMENT": "Segment", "PRUMERNA_VYTIZENOST_PCT": "Prům. vytíženost",
                "MAX_VYTIZENOST_PCT": "Max. vytíženost", "DNI_KRITICKA": "Dní kriticky vytíž.",
                "CELKEM_HODIN": "Odprac. hodin", "POCET_DNI": "Dní s daty", "BUCKET": "Stav",
            })[["Pracoviště", "Segment", "Prům. vytíženost", "Max. vytíženost", "Dní kriticky vytíž.", "Odprac. hodin", "Dní s daty", "Stav"]],
            bucket_col="Stav",
            float_cols={"Prům. vytíženost": "{:.1f} %", "Max. vytíženost": "{:.1f} %", "Odprac. hodin": "{:.1f}"},
        )

        daily_bar_html = fig_html(fig_branch_daily_bar(branch_daily, branch_id, branch_name)) if b_daily["DATE"].nunique() > 1 else ""
        heatmap_html = fig_html(fig_workstation_heatmap(workstation_daily, branch_id, branch_name))
        time_profile_html = fig_html(fig_workstation_time_profile(merged, branch_daily, branch_id, branch_name))
        activity_mix_html = fig_html(fig_activity_mix(b_activity_breakdown)) if not b_activity_breakdown.empty else ""
        employee_medals_html = _employee_medals_html(b_employee_summary)
        weekday_progress_html = build_weekday_progress_table_html(workstation_daily, branch_id)

        growth_note_html = ""
        if b_growth is not None and bool(b_growth["PREKROCENO"]):
            growth_note_html = (
                f'<p class="note" style="color:#d03b3b">⚠ Na pobočce se v datech využívá '
                f'{int(b_growth["POUZITA_PRACOVISTE"])} pracovišť, ale ve work_spaces.xlsx je registrováno jen '
                f'{int(b_growth["NO_WORKSTATIONS"])} — zvažte navýšení kapacity v evidenci.</p>'
            )

        activity_employee_block = f'''
          <div style="margin-top:22px">{activity_mix_html}</div>
          <div style="margin-top:22px">
            <h3 style="margin-top:0">Nejvytíženější zaměstnanci</h3>
            {employee_medals_html}
          </div>'''

        pilot_absence_block = ""
        if branch_id == PILOT_BRANCH_ID:
            pilot_absence_block = build_pilot_absence_section(
                merged, branch_id, branch_name, b_merged["CAPACITY_MIN"].iloc[0],
                absences if absences is not None else pd.DataFrame(columns=["EMPLOYEE", "DATE", "ABSENCE_FRACTION"]),
                fig_html,
            )

        # --- Denní rozvrhy pracovišť po 10minutových blocích, stránkované po týdnech Po-Ne ---
        b_daily_indexed = b_daily.set_index("DATE")
        weeks = {}
        for d in b_merged["DATE"].unique():
            # d může přijít jako numpy.datetime64 (podle verze pandas) — normalizujeme
            # na pd.Timestamp, jinak selže "in" test níže kvůli neshodě hashů mezi typy.
            ts = pd.Timestamp(d)
            week_start = ts - pd.Timedelta(days=ts.weekday())
            weeks.setdefault(week_start, set()).add(ts)
        week_starts_sorted = sorted(weeks.keys())
        shown_week_starts = week_starts_sorted[-MAX_WEEKS_PER_BRANCH:]

        weeks_note = ""
        if len(week_starts_sorted) > MAX_WEEKS_PER_BRANCH:
            weeks_note = (
                f'<p class="note">Zobrazeno {MAX_WEEKS_PER_BRANCH} nejnovějších týdnů '
                f'z celkových {len(week_starts_sorted)}.</p>'
            )

        branch_period_start = b_daily["DATE"].min()
        branch_period_end = b_daily["DATE"].max()

        week_pages = []
        for week_idx, week_start in enumerate(shown_week_starts):
            week_dates_with_data = weeks[week_start]
            week_label = f"Týden {week_start:%d.%m.} – {(week_start + pd.Timedelta(days=6)):%d.%m.%Y}"
            day_sections = []
            for wd_offset in range(7):
                day_date = week_start + pd.Timedelta(days=wd_offset)
                weekday_label = WEEKDAY_LABELS[wd_offset]
                if day_date < branch_period_start or day_date > branch_period_end:
                    # Den je mimo sledované období (ještě nenastal / není v datech) —
                    # neoznačovat jako "zavřeno", prostě ho v týdnu vynecháme.
                    continue
                if day_date in week_dates_with_data:
                    day_row = b_daily_indexed.loc[day_date]
                    day_fig = fig_workstation_day_blocks(merged, branch_id, branch_name, day_date, poledni_pauza)
                    day_sections.append(f"""
          <div class="week-day-block">
            <div class="week-day-heading">{weekday_label} {day_date:%d.%m.%Y} — {day_row["N_ACTIVITIES"]} aktivit,
            {day_row["DURATION_MIN"] / 60:.1f} h, {day_row["UTILIZATION_PCT"]:.0f}% vytíženo</div>
            {fig_html(day_fig)}
          </div>""")
                else:
                    day_sections.append(f"""
          <div class="week-day-block closed">
            <div class="week-day-heading">{weekday_label} {day_date:%d.%m.%Y} — zavřeno / bez aktivit</div>
          </div>""")
            if not day_sections:
                continue
            active_class = " active" if week_idx == len(shown_week_starts) - 1 else ""
            week_pages.append(
                f'<div class="week-page{active_class}" data-label="{week_label}">{"".join(day_sections)}</div>'
            )

        week_nav_id = f"weeknav-{branch_id}"
        default_label = week_pages and shown_week_starts and (
            f"Týden {shown_week_starts[-1]:%d.%m.} – {(shown_week_starts[-1] + pd.Timedelta(days=6)):%d.%m.%Y}"
        ) or "—"
        day_blocks_html = f"""
        <div class="week-nav-wrap" id="{week_nav_id}">
          <div class="week-nav-controls">
            <button class="week-prev" onclick="stepWeek(this,-1)" {"disabled" if len(shown_week_starts) <= 1 else ""}>◀ Předchozí týden</button>
            <span class="week-label">{default_label}</span>
            <button class="week-next" onclick="stepWeek(this,1)" disabled>Další týden ▶</button>
          </div>
          {"".join(week_pages) if week_pages else '<p class="note">Žádná data.</p>'}
        </div>"""

        branch_blocks.append(f"""
        <details class="branch-block">
          <summary>{branch_name} ({branch_id})
            <span class="branch-summary-line">{row['PRUMERNA_VYTIZENOST_PCT']:.0f}% vytíženo · {b_hours:.0f} h · {n_activities} aktivit</span>
          </summary>
          <div class="branch-body">
          {kpi_html}
          {growth_note_html}

          <h3>Vytíženost jednotlivých pracovišť (celé období)</h3>
          {ws_table_html}

          {"" if not daily_bar_html else f'<div style="margin-top:22px">{daily_bar_html}</div>'}

          <h3 style="margin-top:26px">Průměrná vytíženost pracovišť podle dne v týdnu (Po–Ne)</h3>
          <p class="note">Průměrná vytíženost každého pracoviště za každý den v týdnu, zprůměrovaná přes všechny
          výskyty daného dne ve sledovaném období. Dny, kdy je pobočka zavřená, jsou označené „zavřeno".</p>
          {weekday_progress_html}

          <h3 style="margin-top:26px">Vytíženost jednotlivých pracovišť (za jednotlivé dny)</h3>
          <div style="margin-top:6px">{heatmap_html}</div>

          <h3 style="margin-top:26px">Vytíženost pracovišť podle času dne (celé období)</h3>
          <p class="note">Stejný rastr jako denní rozvrh níže (řádek = pracoviště od nejnižšího čísla nahoře,
          sloupec = čas dne po {BLOCK_MINUTES}minutových blocích), ale sečtený přes všechny otevřené dny ve
          sledovaném období. Barva = v kolika % dní bylo dané pracoviště v danou dobu obsazené — od průhledné
          (nikdy) po tmavě modrou (téměř vždy). Ukazuje tak typický denní vzorec vytížení pracoviště.</p>
          {time_profile_html}

          <h3 style="margin-top:26px">Denní rozvrh pracovišť po {BLOCK_MINUTES} minutách</h3>
          <p class="note">Pracoviště seřazená od nejnižšího čísla nahoře, čas po ose x v blocích po
          {BLOCK_MINUTES} minutách, barva bloku = typ aktivity. {"Šedý pás vyznačuje polední pauzu." if poledni_pauza else ""}
          Šipkami přepínáte celý týden Po–Ne najednou — dny bez dat (zavřeno) se zobrazí jako prázdný řádek.</p>
          {weeks_note}
          {day_blocks_html}

          {activity_employee_block}
          {pilot_absence_block}
          </div>
        </details>""")

    branch_blocks_html = "\n".join(branch_blocks)

    html = f"""<!doctype html>
<html lang="cs">
<head>
<meta charset="utf-8">
<title>Report vytíženosti pracovišť</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>{_CSS}</style>
</head>
<body>
<header class="page"><div class="wrap">
  <h1>Report vytíženosti pracovišť poboček</h1>
  <p>Vygenerováno {generated_at} · zdroj: bo_data.xlsx + work_spaces.xlsx</p>
</div></header>
<div class="wrap">

  {cards}

  <div class="section">
    <h2>Přehled všech poboček</h2>
    <div class="legend">
      <span><i style="background:{BUCKET_COLORS['Nízká']}"></i>Nízká (&lt; {THRESHOLD_HIGH:.0f} %)</span>
      <span><i style="background:{BUCKET_COLORS['Vysoká']}"></i>Vysoká ({THRESHOLD_HIGH:.0f}–{THRESHOLD_CRITICAL:.0f} %)</span>
      <span><i style="background:{BUCKET_COLORS['Kritická']}"></i>Kritická (&gt; {THRESHOLD_CRITICAL:.0f} %)</span>
    </div>
    {branch_bar_html}
    <div class="table-scroll">{branch_table_html}</div>
    <p class="note">Vytíženost pobočky = odpracované minuty / (počet pracovišť × kapacita pracoviště v min/den),
    zprůměrováno přes dny, kdy je pobočka otevřená, ve sledovaném období. Denní kapacita pracoviště se počítá
    z týdenní otevírací doby (CAPACITY) dělené počtem otevřených dní v týdnu (5, nebo 7 pro víkendové pobočky).
    Pobočky bez dat v bo_data.xlsx jsou uvedeny se stavem „Bez dat".</p>
  </div>

  <div class="section">
    <h2>Detail pobočky</h2>
    <p class="note">Rozklikněte pobočku — uvidíte vytíženost jejích pracovišť za celé období i po jednotlivých
    dnech, hodinové rozložení aktivit, denní rozvrh pracovišť po {BLOCK_MINUTES} minutách a skladbu aktivit
    se zaměstnanci.</p>
    {branch_blocks_html}
  </div>

</div>
<footer>Report vygenerován automaticky z bo_data.xlsx a work_spaces.xlsx.</footer>
<script>
function stepWeek(btn, delta) {{
  var wrap = btn.closest('.week-nav-wrap');
  var pages = wrap.querySelectorAll('.week-page');
  var cur = 0;
  pages.forEach(function(p, i) {{ if (p.classList.contains('active')) cur = i; }});
  var next = Math.max(0, Math.min(pages.length - 1, cur + delta));
  pages.forEach(function(p, i) {{ p.classList.toggle('active', i === next); }});
  wrap.querySelector('.week-label').textContent = pages[next].dataset.label;
  wrap.querySelector('.week-prev').disabled = (next === 0);
  wrap.querySelector('.week-next').disabled = (next === pages.length - 1);
}}
</script>
</body>
</html>
"""

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    return output_path


# -----------------------------------------------------------------------------
# 7. Spuštění celého výpočtu a generování reportu
# -----------------------------------------------------------------------------

SCRIPT_VERSION = "2026-07-14b (segmenty se čtou z listu struktura_pracovist_segmenty uvnitř qr_codes_bo_online.xlsx)"
print(f"Verze skriptu: {SCRIPT_VERSION}")

activities, data_issues = load_activities(BO_DATA_FILE)
workspaces = load_workspaces(WORKSPACES_FILE)
segments = load_segments(SEGMENTS_FILE)
absences = load_absences(ABSENCE_FILE, PILOT_ORG_UNIT_ZKRATKA)
print(f"Aktivity: {len(activities)} platných řádků, {len(data_issues)} přeskočeno (chybná data).")
print(f"Pobočky (work_spaces.xlsx): {len(workspaces)}")
print(f"Segmenty pracovišť: {len(segments)} řádků")
print(f"Nepřítomnosti (pilotní pobočka {PILOT_ORG_UNIT_ZKRATKA}): {len(absences)} záznamů (zaměstnanec×den)")

merged, unknown_branches = merge_activities_with_workspaces(activities, workspaces)
if not unknown_branches.empty:
    print("Pozor: aktivity patří pobočkám, které nejsou ve work_spaces.xlsx:")
    display(unknown_branches[["BRANCH_ID", "WORKSTATION_ID", "DATETIME", "EMPLOYEE"]])

if merged.empty:
    print("BRANCH_ID v bo_data.xlsx:      ", sorted(activities["BRANCH_ID"].unique()))
    print("BRANCH_ID v work_spaces.xlsx:  ", sorted(workspaces["BRANCH_ID"].unique()))
    raise ValueError(
        "Po spojení s work_spaces.xlsx nezůstala žádná platná aktivita (žádné BRANCH_ID se "
        "neshoduje mezi bo_data.xlsx a work_spaces.xlsx) — viz vypsané seznamy ID výš."
    )

workstation_daily = compute_workstation_daily(merged)
branch_daily = compute_branch_daily(merged)
branch_summary = summarize_branches(branch_daily, workspaces)
growth_flags = compute_capacity_growth_flags(merged, workspaces)

display(branch_summary)
fig_branch_utilization_bar(branch_summary).show()

_output_base = Path(OUTPUT_HTML)
_output_timestamped = _output_base.with_name(f"{_output_base.stem}_{datetime.now():%Y%m%d_%H%M%S}{_output_base.suffix}")

report_path = build_html_report(
    _output_timestamped,
    period_start=branch_daily["DATE"].min(), period_end=branch_daily["DATE"].max(),
    workspaces=workspaces, merged=merged,
    branch_summary=branch_summary, branch_daily=branch_daily, workstation_daily=workstation_daily,
    growth_flags=growth_flags, segments=segments, absences=absences,
)
print(f"\nReport vygenerován (nový soubor, jiný název než minule): {report_path.resolve()}")
print("Otevřete tento konkrétní soubor v prohlížeči — NE starou záložku s předchozí verzí.")

display(IFrame(src=str(report_path), width="100%", height=800))
