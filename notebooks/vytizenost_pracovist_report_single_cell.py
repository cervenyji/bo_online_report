# =============================================================================
# Report vytíženosti pracovišť (Back Office) — vložte do JEDNÉ buňky Jupyter notebooku
#
# Vstupy (v aktuálním adresáři, nebo uveďte plnou cestu níže v KONFIGURACI):
#   bo_data.xlsx      : BRANCH_ID, PRACOVISTE_ID, DATETIME, ZAMESTNANEC, ACTIVITY, DURATION (minuty)
#   work_spaces.xlsx  : BRANCH_ID, BRANCH_NAME, NO_WPL (počet pracovišť), CAPACITY (kapacita 1 pracoviště v hod/den)
#
# Výstup: samostatný HTML report (Plotly.js vložený přímo v souboru, funguje i offline).
#
# Poznámka: .xlsx soubory se čtou vlastní minimální čtečkou (jen zipfile + xml ze
# standardní knihovny) — NENÍ potřeba mít nainstalovaný openpyxl ani žádný jiný
# balíček pro práci s Excelem. Potřeba je jen pandas, numpy a plotly.
# =============================================================================

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
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
WORKSPACES_FILE = "work_spaces.xlsx"
OUTPUT_HTML = "vytizenost_report.html"

BUSINESS_DAYS_ONLY = True  # pobočky mají provoz Po-Pá -> vytíženost se počítá jen pro pracovní dny
THRESHOLD_HIGH = 70.0      # od této hranice (%) je pracoviště "vysoce vytížené"
THRESHOLD_CRITICAL = 90.0  # od této hranice (%) je pracoviště "kriticky vytížené" / na hraně kapacity

BUCKET_COLORS = {"Nízká": "#2E7D32", "Vysoká": "#F9A825", "Kritická": "#C62828"}

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


def read_xlsx_rows(path, sheet_index=0):
    """Přečte list .xlsx souboru a vrátí seznam řádků (list řádků, každý je list buněk).
    Rozpozná datumové buňky podle formátu ve stylu a rovnou je vrátí jako datetime."""
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


def load_workspaces(path):
    """Načte work_spaces.xlsx a sjednotí názvy sloupců dle pozice."""
    raw = read_xlsx(path)
    if raw.shape[1] < 4:
        raise ValueError(
            f"Soubor {path} má jen {raw.shape[1]} sloupců, očekává se 4: "
            "BRANCH_ID, BRANCH_NAME, NO_WPL, CAPACITY"
        )
    df = raw.iloc[:, :4].copy()
    df.columns = ["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_HOURS"]
    df["BRANCH_ID"] = pd.to_numeric(df["BRANCH_ID"], errors="coerce")
    df["NO_WORKSTATIONS"] = pd.to_numeric(df["NO_WORKSTATIONS"], errors="coerce")
    df["CAPACITY_HOURS"] = pd.to_numeric(df["CAPACITY_HOURS"], errors="coerce")

    bad_mask = df["BRANCH_ID"].isna() | df["NO_WORKSTATIONS"].isna() | df["CAPACITY_HOURS"].isna()
    if bad_mask.any():
        print(f"Pozor: {bad_mask.sum()} řádků ve work_spaces.xlsx bylo přeskočeno (chybí BRANCH_ID/NO_WPL/CAPACITY):")
        display(raw.iloc[:, :4].loc[bad_mask])
    df = df.loc[~bad_mask].copy()

    df["BRANCH_ID"] = df["BRANCH_ID"].astype(int)
    df["NO_WORKSTATIONS"] = df["NO_WORKSTATIONS"].astype(int)
    df["CAPACITY_MIN"] = df["CAPACITY_HOURS"] * 60
    return df.reset_index(drop=True)


def merge_activities_with_workspaces(activities, workspaces):
    """Připojí k aktivitám info o pobočce. Vrátí (spojená_data, aktivity_z_neznámé_pobočky)."""
    merged = activities.merge(workspaces, on="BRANCH_ID", how="left", indicator=True)
    unknown = merged.loc[merged["_merge"] == "left_only"].copy()
    known = merged.loc[merged["_merge"] == "both"].drop(columns="_merge").copy()
    return known.reset_index(drop=True), unknown.reset_index(drop=True)


def full_date_grid(dates):
    """Vrátí kompletní řadu dnů pokrývající data (pracovní dny Po-Pá, nebo všechny dny)."""
    dates = dates.dropna()
    if dates.empty:
        raise ValueError(
            "Nepodařilo se určit žádné platné datum aktivit po spojení s work_spaces.xlsx. "
            "Nejčastější příčina: všechny řádky v bo_data.xlsx patří pobočkám (BRANCH_ID), "
            "které nejsou ve work_spaces.xlsx — zkontrolujte hlášku 'Pozor: aktivity patří "
            "pobočkám...' o pár buněk výš a porovnejte BRANCH_ID v obou souborech."
        )
    start, end = dates.min(), dates.max()
    return pd.bdate_range(start, end) if BUSINESS_DAYS_ONLY else pd.date_range(start, end)


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
    grid = full_date_grid(merged["DATE"])
    daily = (
        merged.groupby(["BRANCH_ID", "BRANCH_NAME", "WORKSTATION_ID", "DATE"])
        .agg(DURATION_MIN=("DURATION_MIN", "sum"), N_ACTIVITIES=("DURATION_MIN", "count"))
        .reset_index()
    )
    branches = merged[["BRANCH_ID", "BRANCH_NAME", "CAPACITY_MIN"]].drop_duplicates("BRANCH_ID")
    workstations = merged[["BRANCH_ID", "WORKSTATION_ID"]].drop_duplicates()

    full_index = workstations.merge(pd.Series(grid, name="DATE"), how="cross")
    full_index = full_index.merge(branches, on="BRANCH_ID", how="left")

    out = full_index.merge(daily.drop(columns="BRANCH_NAME"), on=["BRANCH_ID", "WORKSTATION_ID", "DATE"], how="left")
    out["DURATION_MIN"] = out["DURATION_MIN"].fillna(0.0)
    out["N_ACTIVITIES"] = out["N_ACTIVITIES"].fillna(0).astype(int)
    out["UTILIZATION_PCT"] = np.where(out["CAPACITY_MIN"] > 0, out["DURATION_MIN"] / out["CAPACITY_MIN"] * 100, np.nan)
    out["BUCKET"] = out["UTILIZATION_PCT"].apply(utilization_bucket)
    return out.sort_values(["BRANCH_ID", "WORKSTATION_ID", "DATE"]).reset_index(drop=True)


def compute_branch_daily(merged):
    """Denní vytíženost celé pobočky (součet přes všechna pracoviště vs. registrovaná kapacita)."""
    grid = full_date_grid(merged["DATE"])
    branches = merged[["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_HOURS", "CAPACITY_MIN"]].drop_duplicates("BRANCH_ID")
    branches["BRANCH_CAPACITY_MIN"] = branches["NO_WORKSTATIONS"] * branches["CAPACITY_MIN"]

    daily = (
        merged.groupby(["BRANCH_ID", "DATE"])
        .agg(DURATION_MIN=("DURATION_MIN", "sum"), N_ACTIVITIES=("DURATION_MIN", "count"))
        .reset_index()
    )
    full_index = branches[["BRANCH_ID"]].merge(pd.Series(grid, name="DATE"), how="cross")
    out = full_index.merge(branches, on="BRANCH_ID", how="left").merge(daily, on=["BRANCH_ID", "DATE"], how="left")
    out["DURATION_MIN"] = out["DURATION_MIN"].fillna(0.0)
    out["N_ACTIVITIES"] = out["N_ACTIVITIES"].fillna(0).astype(int)
    out["UTILIZATION_PCT"] = np.where(out["BRANCH_CAPACITY_MIN"] > 0, out["DURATION_MIN"] / out["BRANCH_CAPACITY_MIN"] * 100, np.nan)
    out["BUCKET"] = out["UTILIZATION_PCT"].apply(utilization_bucket)
    return out.sort_values(["BRANCH_ID", "DATE"]).reset_index(drop=True)


def summarize_branches(branch_daily, workspaces):
    """Souhrn na úrovni pobočky za celé sledované období + poznámka o pobočkách bez dat."""
    summary = (
        branch_daily.groupby(["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_HOURS"])
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
        ["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_HOURS"]
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


def detect_possible_overlaps(merged):
    """Zjednodušená kontrola kolizí: hledá dvě po sobě jdoucí (dle začátku) rezervace
    stejného pracoviště, které se v čase překrývají (typicky = chybný záznam nebo
    reálný souběh dvou zaměstnanců na jednom místě)."""
    cols = ["BRANCH_ID", "BRANCH_NAME", "WORKSTATION_ID", "EMPLOYEE", "DATETIME", "END_DATETIME", "ACTIVITY"]
    s = merged[cols].sort_values(["BRANCH_ID", "WORKSTATION_ID", "DATETIME"]).reset_index(drop=True)
    grouped = s.groupby(["BRANCH_ID", "WORKSTATION_ID"], group_keys=False)

    prev_end = grouped["END_DATETIME"].shift(1)
    prev_employee = grouped["EMPLOYEE"].shift(1)
    prev_start = grouped["DATETIME"].shift(1)
    overlap_mask = s["DATETIME"] < prev_end

    out = s.loc[overlap_mask].copy()
    out["PREDCHOZI_ZAMESTNANEC"] = prev_employee.loc[overlap_mask]
    out["PREDCHOZI_START"] = prev_start.loc[overlap_mask]
    out["PREDCHOZI_END"] = prev_end.loc[overlap_mask]
    out = out.rename(columns={"EMPLOYEE": "ZAMESTNANEC", "DATETIME": "START", "END_DATETIME": "END"})
    return out[[
        "BRANCH_ID", "BRANCH_NAME", "WORKSTATION_ID",
        "PREDCHOZI_ZAMESTNANEC", "PREDCHOZI_START", "PREDCHOZI_END",
        "ZAMESTNANEC", "START", "END", "ACTIVITY",
    ]].reset_index(drop=True)


# -----------------------------------------------------------------------------
# 5. Grafy (Plotly)
# -----------------------------------------------------------------------------

def fig_branch_utilization_bar(branch_summary):
    df = branch_summary.dropna(subset=["PRUMERNA_VYTIZENOST_PCT"]).sort_values("PRUMERNA_VYTIZENOST_PCT")
    colors = df["BUCKET"].map(BUCKET_COLORS).fillna("#9E9E9E")
    fig = go.Figure(go.Bar(
        x=df["PRUMERNA_VYTIZENOST_PCT"], y=df["BRANCH_NAME"] + " (" + df["BRANCH_ID"].astype(str) + ")",
        orientation="h", marker_color=colors,
        text=df["PRUMERNA_VYTIZENOST_PCT"].round(1).astype(str) + " %", textposition="outside",
        hovertemplate="%{y}<br>Průměrná vytíženost: %{x:.1f} %<extra></extra>",
    ))
    fig.add_vline(x=THRESHOLD_HIGH, line_dash="dot", line_color="#F9A825")
    fig.add_vline(x=THRESHOLD_CRITICAL, line_dash="dot", line_color="#C62828")
    fig.add_vline(x=100, line_color="#616161")
    fig.update_layout(
        title="Průměrná denní vytíženost poboček vůči kapacitě", xaxis_title="Vytíženost (%)", yaxis_title=None,
        height=max(320, 28 * len(df) + 120), margin=dict(l=10, r=10, t=60, b=40), template="plotly_white",
    )
    return fig


def fig_workstation_heatmap(workstation_daily, branch_id, branch_name):
    d = workstation_daily.loc[workstation_daily["BRANCH_ID"] == branch_id].copy()
    d["WORKSTATION_LABEL"] = "Prac. " + d["WORKSTATION_ID"].astype(str)
    pivot = d.pivot_table(index="DATE", columns="WORKSTATION_LABEL", values="UTILIZATION_PCT")
    pivot = pivot.reindex(sorted(pivot.columns, key=lambda c: int(c.split(" ")[1])), axis=1)

    fig = go.Figure(go.Heatmap(
        z=pivot.values, x=pivot.columns, y=[d.strftime("%d.%m.%Y") for d in pivot.index],
        colorscale=[[0.0, "#2E7D32"], [THRESHOLD_HIGH / 150, "#F9A825"], [THRESHOLD_CRITICAL / 150, "#C62828"], [1.0, "#7B0000"]],
        zmin=0, zmax=150, colorbar=dict(title="%"),
        hovertemplate="%{x} | %{y}<br>Vytíženost: %{z:.1f} %<extra></extra>",
    ))
    fig.update_layout(
        title=f"Denní vytíženost pracovišť — {branch_name} ({branch_id})",
        height=max(260, 24 * len(pivot.index) + 120), margin=dict(l=10, r=10, t=60, b=40), template="plotly_white",
    )
    return fig


def fig_utilization_trend(branch_daily, top_n=8):
    top_branches = branch_daily.groupby("BRANCH_NAME")["UTILIZATION_PCT"].mean().sort_values(ascending=False).head(top_n).index
    d = branch_daily.loc[branch_daily["BRANCH_NAME"].isin(top_branches)]
    fig = px.line(
        d, x="DATE", y="UTILIZATION_PCT", color="BRANCH_NAME", markers=True,
        labels={"DATE": "Datum", "UTILIZATION_PCT": "Vytíženost (%)", "BRANCH_NAME": "Pobočka"},
        title=f"Trend vytíženosti v čase (top {top_n} poboček dle průměru)",
    )
    fig.add_hline(y=THRESHOLD_CRITICAL, line_dash="dot", line_color="#C62828")
    fig.update_layout(height=440, template="plotly_white", margin=dict(l=10, r=10, t=60, b=40))
    return fig


def fig_activity_mix(activity_breakdown):
    fig = px.pie(activity_breakdown, names="ACTIVITY", values="CELKEM_HODIN", hole=0.45, title="Skladba aktivit dle odpracovaných hodin")
    fig.update_traces(textinfo="percent+label", hovertemplate="%{label}<br>%{value:.1f} h (%{percent})<extra></extra>")
    fig.update_layout(height=420, template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))
    return fig


def fig_employee_top(employee_summary, n=15):
    d = employee_summary.sort_values("CELKEM_HODIN", ascending=False).head(n).sort_values("CELKEM_HODIN")
    fig = go.Figure(go.Bar(
        x=d["CELKEM_HODIN"], y=d["EMPLOYEE"] + " — " + d["BRANCH_NAME"], orientation="h", marker_color="#1565C0",
        text=d["CELKEM_HODIN"].round(1).astype(str) + " h", textposition="outside",
    ))
    fig.update_layout(
        title=f"Nejvytíženější zaměstnanci (top {n} dle odpracovaných hodin)", xaxis_title="Hodiny celkem", yaxis_title=None,
        height=max(320, 26 * len(d) + 120), template="plotly_white", margin=dict(l=10, r=10, t=60, b=40),
    )
    return fig


# -----------------------------------------------------------------------------
# 6. Sestavení HTML reportu
# -----------------------------------------------------------------------------

_CSS = """
:root { color-scheme: light; }
* { box-sizing: border-box; }
body {
    font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
    background: #F4F6F8; color: #1A1A1A; margin: 0; padding: 0 0 60px 0;
}
.wrap { max-width: 1180px; margin: 0 auto; padding: 0 24px; }
header.page { background: #0D2A4A; color: white; padding: 36px 0; margin-bottom: 28px; }
header.page h1 { margin: 0 0 6px 0; font-size: 26px; }
header.page p { margin: 0; opacity: 0.85; font-size: 14px; }
h2 { font-size: 19px; border-bottom: 2px solid #E0E4E8; padding-bottom: 8px; margin-top: 44px; }
h3 { font-size: 15px; color: #333; }
.card-row { display: flex; gap: 16px; flex-wrap: wrap; margin: 16px 0 8px 0; }
.card {
    background: white; border-radius: 10px; padding: 18px 22px; flex: 1 1 200px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-left: 4px solid #0D2A4A;
}
.card .label { font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: .04em; }
.card .value { font-size: 26px; font-weight: 700; margin-top: 4px; }
.section { background: white; border-radius: 10px; padding: 20px 24px; margin-top: 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
table.report { border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 10px; }
table.report th { background: #0D2A4A; color: white; text-align: left; padding: 8px 10px; position: sticky; top: 0; }
table.report td { padding: 7px 10px; border-bottom: 1px solid #EEE; }
table.report tr:nth-child(even) { background: #FAFBFC; }
.badge { display: inline-block; padding: 2px 9px; border-radius: 12px; color: white; font-size: 12px; font-weight: 600; }
.table-scroll { max-height: 480px; overflow-y: auto; }
details.branch-block { margin-bottom: 14px; border: 1px solid #E0E4E8; border-radius: 8px; }
details.branch-block summary { padding: 10px 14px; cursor: pointer; font-weight: 600; background: #F0F3F6; border-radius: 8px; }
details.branch-block[open] summary { border-radius: 8px 8px 0 0; }
.note { font-size: 12.5px; color: #666; margin-top: 6px; }
footer { text-align: center; color: #888; font-size: 12px; margin-top: 50px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; margin-right: 18px; font-size: 12.5px; }
.legend i { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }
"""


def _bucket_badge(bucket):
    color = BUCKET_COLORS.get(bucket, "#9E9E9E")
    return f'<span class="badge" style="background:{color}">{bucket}</span>'


def _df_to_html_table(df, bucket_col=None, float_cols=None):
    d = df.copy()
    float_cols = float_cols or {}
    for col, fmt in float_cols.items():
        if col in d.columns:
            d[col] = d[col].map(lambda v: (fmt.format(v) if pd.notna(v) else "—"))
    if bucket_col and bucket_col in d.columns:
        d[bucket_col] = d[bucket_col].map(_bucket_badge)
    return d.to_html(index=False, escape=False, classes="report", border=0)


def build_html_report(
    output_path, *, period_start, period_end, activities, workspaces, merged,
    branch_summary, branch_daily, workstation_daily, growth_flags,
    activity_breakdown, employee_summary, overlaps, data_issues, unknown_branches,
):
    generated_at = datetime.now().strftime("%d.%m.%Y %H:%M")

    overall_util = branch_daily["UTILIZATION_PCT"].mean()
    total_hours = merged["DURATION_MIN"].sum() / 60
    n_branches_with_data = merged["BRANCH_ID"].nunique()
    n_branches_total = workspaces["BRANCH_ID"].nunique()
    n_workstations_registered = int(workspaces["NO_WORKSTATIONS"].sum())
    n_employees = merged["EMPLOYEE"].nunique()

    cards = f"""
    <div class="card-row">
      <div class="card"><div class="label">Sledované období</div>
        <div class="value" style="font-size:18px">{period_start:%d.%m.%Y} – {period_end:%d.%m.%Y}</div></div>
      <div class="card"><div class="label">Průměrná vytíženost poboček</div>
        <div class="value">{overall_util:.1f} %</div></div>
      <div class="card"><div class="label">Odpracované hodiny celkem</div>
        <div class="value">{total_hours:,.0f} h</div></div>
      <div class="card"><div class="label">Pobočky s daty / celkem</div>
        <div class="value">{n_branches_with_data} / {n_branches_total}</div></div>
      <div class="card"><div class="label">Registrovaná pracoviště</div>
        <div class="value">{n_workstations_registered}</div></div>
      <div class="card"><div class="label">Aktivních zaměstnanců</div>
        <div class="value">{n_employees}</div></div>
    </div>
    """

    plotly_cdn_included = False

    def fig_html(fig):
        # Plotly.js se vloží celý přímo do souboru (ne přes CDN), aby report šel
        # otevřít i bez internetu (e-mail, sdílený disk, offline prohlížení).
        nonlocal plotly_cdn_included
        html = fig.to_html(full_html=False, include_plotlyjs=(True if not plotly_cdn_included else False), config={"displaylogo": False})
        plotly_cdn_included = True
        return html

    branch_bar_html = fig_html(fig_branch_utilization_bar(branch_summary))
    trend_html = fig_html(fig_utilization_trend(branch_daily)) if branch_daily["DATE"].nunique() > 1 else ""
    activity_html = fig_html(fig_activity_mix(activity_breakdown)) if not activity_breakdown.empty else ""
    employee_html = fig_html(fig_employee_top(employee_summary)) if not employee_summary.empty else ""

    heatmap_blocks = []
    for branch_id in sorted(workstation_daily["BRANCH_ID"].unique()):
        name = workstation_daily.loc[workstation_daily["BRANCH_ID"] == branch_id, "BRANCH_NAME"].iloc[0]
        fig = fig_workstation_heatmap(workstation_daily, branch_id, name)
        heatmap_blocks.append(
            f'<details class="branch-block"><summary>{name} ({branch_id})</summary>'
            f'<div style="padding:14px">{fig_html(fig)}</div></details>'
        )
    heatmaps_html = "\n".join(heatmap_blocks) if heatmap_blocks else "<p>Žádná data k dispozici.</p>"

    branch_table_html = _df_to_html_table(
        branch_summary.rename(columns={
            "BRANCH_NAME": "Pobočka", "BRANCH_ID": "ID", "NO_WORKSTATIONS": "Poč. pracovišť",
            "CAPACITY_HOURS": "Kapacita (h/den)", "PRUMERNA_VYTIZENOST_PCT": "Prům. vytíženost",
            "MAX_VYTIZENOST_PCT": "Max. vytíženost", "DNI_KRITICKA": "Dní kriticky vytíž.",
            "CELKEM_HODIN": "Odprac. hodin", "POCET_DNI": "Dní s daty", "BUCKET": "Stav",
        }),
        bucket_col="Stav",
        float_cols={"Prům. vytíženost": "{:.1f} %", "Max. vytíženost": "{:.1f} %", "Odprac. hodin": "{:.1f}"},
    )

    exceeded = growth_flags.loc[growth_flags["PREKROCENO"]].drop(columns="PREKROCENO")
    if not exceeded.empty:
        growth_html = _df_to_html_table(exceeded.rename(columns={
            "BRANCH_NAME": "Pobočka", "BRANCH_ID": "ID",
            "POUZITA_PRACOVISTE": "Využitá pracoviště (dle dat)",
            "NO_WORKSTATIONS": "Registrováno ve work_spaces.xlsx", "ROZDIL": "Rozdíl",
        }))
    else:
        growth_html = "<p>Ve sledovaném období nebyla zjištěna žádná pobočka, kde by se využívalo víc pracovišť, než je registrováno.</p>"

    overlaps_html = (
        _df_to_html_table(overlaps.head(200).rename(columns={
            "BRANCH_ID": "ID pobočky", "BRANCH_NAME": "Pobočka", "WORKSTATION_ID": "Pracoviště",
            "PREDCHOZI_ZAMESTNANEC": "Předchozí zaměstnanec", "PREDCHOZI_START": "Předchozí od",
            "PREDCHOZI_END": "Předchozí do", "ZAMESTNANEC": "Zaměstnanec",
            "START": "Od", "END": "Do", "ACTIVITY": "Aktivita",
        }))
        if not overlaps.empty
        else "<p>Nebyly nalezeny žádné časové kolize mezi sousedícími rezervacemi stejného pracoviště.</p>"
    )

    data_quality_bits = []
    if not data_issues.empty:
        data_quality_bits.append(
            f"<li>{len(data_issues)} řádků v bo_data.xlsx bylo přeskočeno kvůli chybějícím/neplatným hodnotám "
            "(BRANCH_ID, PRACOVISTE_ID, DATETIME nebo DURATION).</li>"
        )
    if not unknown_branches.empty:
        ids = ", ".join(str(x) for x in sorted(unknown_branches["BRANCH_ID"].dropna().unique()))
        data_quality_bits.append(
            f"<li>{len(unknown_branches)} aktivit patří pobočkám, které nejsou ve work_spaces.xlsx (ID: {ids}) "
            "— tyto aktivity nejsou v reportu zahrnuty.</li>"
        )
    data_quality_html = f"<ul>{''.join(data_quality_bits)}</ul>" if data_quality_bits else "<p>Bez zjištěných problémů v datech.</p>"

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
    <h2>Vytíženost poboček</h2>
    <div class="legend">
      <span><i style="background:{BUCKET_COLORS['Nízká']}"></i>Nízká (&lt; {THRESHOLD_HIGH:.0f} %)</span>
      <span><i style="background:{BUCKET_COLORS['Vysoká']}"></i>Vysoká ({THRESHOLD_HIGH:.0f}–{THRESHOLD_CRITICAL:.0f} %)</span>
      <span><i style="background:{BUCKET_COLORS['Kritická']}"></i>Kritická (&gt; {THRESHOLD_CRITICAL:.0f} %)</span>
    </div>
    {branch_bar_html}
    <div class="table-scroll">{branch_table_html}</div>
    <p class="note">Vytíženost pobočky = odpracované minuty / (počet pracovišť × kapacita pracoviště v min/den),
    zprůměrováno přes pracovní dny ve sledovaném období. Pobočky bez dat v bo_data.xlsx jsou uvedeny se stavem „Bez dat".</p>
  </div>

  {"" if not trend_html else f'''
  <div class="section">
    <h2>Trend vytíženosti v čase</h2>
    {trend_html}
  </div>'''}

  <div class="section">
    <h2>Vytíženost jednotlivých pracovišť</h2>
    <p class="note">Rozklikněte pobočku pro zobrazení mapy vytíženosti jednotlivých pracovišť podle dne.</p>
    {heatmaps_html}
  </div>

  <div class="section">
    <h2>Kapacita: je pracovišť dost?</h2>
    <p class="note">Porovnání skutečně využívaných pracovišť (dle unikátních PRACOVISTE_ID v datech) s počtem
    registrovaným ve work_spaces.xlsx. Časem může na pobočce přibývat pracovišť — tato tabulka na to upozorní.</p>
    {growth_html}
  </div>

  {"" if not activity_html and not employee_html else f'''
  <div class="section">
    <h2>Skladba aktivit a zaměstnanci</h2>
    <div style="display:flex; gap:24px; flex-wrap:wrap">
      <div style="flex:1 1 420px">{activity_html}</div>
      <div style="flex:1 1 420px">{employee_html}</div>
    </div>
  </div>'''}

  <div class="section">
    <h2>Kontrola kolizí rezervací</h2>
    <p class="note">Zjednodušená kontrola: hledá sousedící (dle času začátku) rezervace stejného pracoviště,
    které se časově překrývají — může jít o chybu v datech nebo o reálný souběh dvou lidí na jednom místě.</p>
    <div class="table-scroll">{overlaps_html}</div>
  </div>

  <div class="section">
    <h2>Kvalita dat</h2>
    {data_quality_html}
  </div>

</div>
<footer>Report vygenerován automaticky z bo_data.xlsx a work_spaces.xlsx.</footer>
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

SCRIPT_VERSION = "2026-07-03c (diagnostika NaT / neshoda BRANCH_ID)"
print(f"Verze skriptu: {SCRIPT_VERSION}")

activities, data_issues = load_activities(BO_DATA_FILE)
workspaces = load_workspaces(WORKSPACES_FILE)
print(f"Aktivity: {len(activities)} platných řádků, {len(data_issues)} přeskočeno (chybná data).")
print(f"Pobočky (work_spaces.xlsx): {len(workspaces)}")

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
activity_breakdown = compute_activity_breakdown(merged)
employee_summary = compute_employee_summary(merged)
overlaps = detect_possible_overlaps(merged)

display(branch_summary)
fig_branch_utilization_bar(branch_summary).show()

for _branch_id in sorted(workstation_daily["BRANCH_ID"].unique()):
    _name = workstation_daily.loc[workstation_daily["BRANCH_ID"] == _branch_id, "BRANCH_NAME"].iloc[0]
    fig_workstation_heatmap(workstation_daily, _branch_id, _name).show()

if branch_daily["DATE"].nunique() > 1:
    fig_utilization_trend(branch_daily).show()

fig_activity_mix(activity_breakdown).show()
fig_employee_top(employee_summary).show()

report_path = build_html_report(
    OUTPUT_HTML,
    period_start=branch_daily["DATE"].min(), period_end=branch_daily["DATE"].max(),
    activities=activities, workspaces=workspaces, merged=merged,
    branch_summary=branch_summary, branch_daily=branch_daily, workstation_daily=workstation_daily,
    growth_flags=growth_flags, activity_breakdown=activity_breakdown, employee_summary=employee_summary,
    overlaps=overlaps, data_issues=data_issues, unknown_branches=unknown_branches,
)
print(f"\nReport vygenerován: {report_path.resolve()}")

display(IFrame(src=str(report_path), width="100%", height=800))
