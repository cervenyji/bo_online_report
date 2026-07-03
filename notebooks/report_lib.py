"""
Knihovna pro výpočet a vizualizaci vytíženosti pracovišť (BO) vůči denní kapacitě.

Vstupy:
- bo_data.xlsx      : BRANCH_ID, PRACOVISTE_ID, DATETIME, ZAMESTNANEC, ACTIVITY, DURATION (minuty)
- work_spaces.xlsx  : BRANCH_ID, BRANCH_NAME, NO_WPL (počet pracovišť), CAPACITY (kapacita 1 pracoviště v hod/den)

Používá se z notebooku vytizenost_pracovist_report.ipynb.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

# ---------------------------------------------------------------------------
# Konfigurace
# ---------------------------------------------------------------------------

BUSINESS_DAYS_ONLY = True  # pobočky mají provoz Po-Pá -> vytíženost se počítá jen pro pracovní dny

THRESHOLD_HIGH = 70.0      # od této hranice (%) je pracoviště "vysoce vytížené"
THRESHOLD_CRITICAL = 90.0  # od této hranice (%) je pracoviště "kriticky vytížené" / na hraně kapacity

BUCKET_COLORS = {
    "Nízká": "#2E7D32",
    "Vysoká": "#F9A825",
    "Kritická": "#C62828",
}

ACTIVITY_ORDER_HINT = None  # lze doplnit pevné pořadí aktivit, jinak se použije pořadí dle celkového trvání


# ---------------------------------------------------------------------------
# Načtení a příprava dat
# ---------------------------------------------------------------------------

def load_activities(path: str | Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Načte bo_data.xlsx, sjednotí názvy sloupců (dle pozice, ne dle přesného
    znění hlavičky) a vrátí (očištěná_data, řádky_s_problémem)."""
    raw = pd.read_excel(path)
    if raw.shape[1] < 6:
        raise ValueError(
            f"Soubor {path} má jen {raw.shape[1]} sloupců, očekává se 6: "
            "BRANCH_ID, PRACOVISTE_ID, DATETIME, ZAMESTNANEC, ACTIVITY, DURATION"
        )
    df = raw.iloc[:, :6].copy()
    df.columns = [
        "BRANCH_ID", "WORKSTATION_ID", "DATETIME_RAW",
        "EMPLOYEE", "ACTIVITY", "DURATION_MIN",
    ]

    df["BRANCH_ID"] = pd.to_numeric(df["BRANCH_ID"], errors="coerce")
    df["WORKSTATION_ID"] = pd.to_numeric(df["WORKSTATION_ID"], errors="coerce")
    df["DURATION_MIN"] = pd.to_numeric(df["DURATION_MIN"], errors="coerce")

    parsed = pd.to_datetime(df["DATETIME_RAW"], format="%d.%m.%Y %H:%M:%S", errors="coerce")
    still_missing = parsed.isna() & df["DATETIME_RAW"].notna()
    if still_missing.any():
        parsed.loc[still_missing] = pd.to_datetime(
            df.loc[still_missing, "DATETIME_RAW"], dayfirst=True, errors="coerce"
        )
    df["DATETIME"] = parsed
    df["DATE"] = df["DATETIME"].dt.normalize()
    df["END_DATETIME"] = df["DATETIME"] + pd.to_timedelta(df["DURATION_MIN"], unit="m")

    bad_mask = (
        df["BRANCH_ID"].isna()
        | df["WORKSTATION_ID"].isna()
        | df["DATETIME"].isna()
        | df["DURATION_MIN"].isna()
        | (df["DURATION_MIN"] <= 0)
    )
    issues = df.loc[bad_mask].copy()
    clean = df.loc[~bad_mask].copy()
    clean["BRANCH_ID"] = clean["BRANCH_ID"].astype(int)
    clean["WORKSTATION_ID"] = clean["WORKSTATION_ID"].astype(int)

    return clean.reset_index(drop=True), issues.reset_index(drop=True)


def load_workspaces(path: str | Path) -> pd.DataFrame:
    """Načte work_spaces.xlsx a sjednotí názvy sloupců dle pozice."""
    raw = pd.read_excel(path)
    if raw.shape[1] < 4:
        raise ValueError(
            f"Soubor {path} má jen {raw.shape[1]} sloupců, očekává se 4: "
            "BRANCH_ID, BRANCH_NAME, NO_WPL, CAPACITY"
        )
    df = raw.iloc[:, :4].copy()
    df.columns = ["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_HOURS"]
    df["BRANCH_ID"] = pd.to_numeric(df["BRANCH_ID"], errors="coerce").astype(int)
    df["NO_WORKSTATIONS"] = pd.to_numeric(df["NO_WORKSTATIONS"], errors="coerce").astype(int)
    df["CAPACITY_HOURS"] = pd.to_numeric(df["CAPACITY_HOURS"], errors="coerce")
    df["CAPACITY_MIN"] = df["CAPACITY_HOURS"] * 60
    return df.reset_index(drop=True)


def merge_activities_with_workspaces(
    activities: pd.DataFrame, workspaces: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Připojí k aktivitám info o pobočce. Vrátí (spojená_data, aktivity_z_neznámé_pobočky)."""
    merged = activities.merge(workspaces, on="BRANCH_ID", how="left", indicator=True)
    unknown = merged.loc[merged["_merge"] == "left_only"].copy()
    known = merged.loc[merged["_merge"] == "both"].drop(columns="_merge").copy()
    return known.reset_index(drop=True), unknown.reset_index(drop=True)


def full_date_grid(dates: pd.Series) -> pd.DatetimeIndex:
    """Vrátí kompletní řadu dnů pokrývající data (pracovní dny Po-Pá, nebo všechny dny)."""
    dates = dates.dropna()
    if dates.empty:
        raise ValueError(
            "Nepodařilo se určit žádné platné datum aktivit po spojení s work_spaces.xlsx. "
            "Nejčastější příčina: všechny řádky v bo_data.xlsx patří pobočkám (BRANCH_ID), "
            "které nejsou ve work_spaces.xlsx — porovnejte BRANCH_ID v obou souborech "
            "(viz proměnná `unknown_branches`)."
        )
    start, end = dates.min(), dates.max()
    if BUSINESS_DAYS_ONLY:
        return pd.bdate_range(start, end)
    return pd.date_range(start, end)


# ---------------------------------------------------------------------------
# Výpočet vytíženosti
# ---------------------------------------------------------------------------

def compute_workstation_daily(merged: pd.DataFrame) -> pd.DataFrame:
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

    out = full_index.merge(
        daily.drop(columns="BRANCH_NAME"), on=["BRANCH_ID", "WORKSTATION_ID", "DATE"], how="left"
    )
    out["DURATION_MIN"] = out["DURATION_MIN"].fillna(0.0)
    out["N_ACTIVITIES"] = out["N_ACTIVITIES"].fillna(0).astype(int)
    out["UTILIZATION_PCT"] = np.where(
        out["CAPACITY_MIN"] > 0, out["DURATION_MIN"] / out["CAPACITY_MIN"] * 100, np.nan
    )
    out["BUCKET"] = out["UTILIZATION_PCT"].apply(utilization_bucket)
    return out.sort_values(["BRANCH_ID", "WORKSTATION_ID", "DATE"]).reset_index(drop=True)


def compute_branch_daily(merged: pd.DataFrame) -> pd.DataFrame:
    """Denní vytíženost celé pobočky (součet přes všechna pracoviště vs. registrovaná kapacita)."""
    grid = full_date_grid(merged["DATE"])
    branches = merged[
        ["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_HOURS", "CAPACITY_MIN"]
    ].drop_duplicates("BRANCH_ID")
    branches["BRANCH_CAPACITY_MIN"] = branches["NO_WORKSTATIONS"] * branches["CAPACITY_MIN"]

    daily = (
        merged.groupby(["BRANCH_ID", "DATE"])
        .agg(DURATION_MIN=("DURATION_MIN", "sum"), N_ACTIVITIES=("DURATION_MIN", "count"))
        .reset_index()
    )

    full_index = branches[["BRANCH_ID"]].merge(pd.Series(grid, name="DATE"), how="cross")
    out = full_index.merge(branches, on="BRANCH_ID", how="left").merge(
        daily, on=["BRANCH_ID", "DATE"], how="left"
    )
    out["DURATION_MIN"] = out["DURATION_MIN"].fillna(0.0)
    out["N_ACTIVITIES"] = out["N_ACTIVITIES"].fillna(0).astype(int)
    out["UTILIZATION_PCT"] = np.where(
        out["BRANCH_CAPACITY_MIN"] > 0,
        out["DURATION_MIN"] / out["BRANCH_CAPACITY_MIN"] * 100,
        np.nan,
    )
    out["BUCKET"] = out["UTILIZATION_PCT"].apply(utilization_bucket)
    return out.sort_values(["BRANCH_ID", "DATE"]).reset_index(drop=True)


def utilization_bucket(pct: float) -> str:
    if pd.isna(pct):
        return "Bez dat"
    if pct >= THRESHOLD_CRITICAL:
        return "Kritická"
    if pct >= THRESHOLD_HIGH:
        return "Vysoká"
    return "Nízká"


def summarize_branches(branch_daily: pd.DataFrame, workspaces: pd.DataFrame) -> pd.DataFrame:
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

    branches_without_data = workspaces.loc[
        ~workspaces["BRANCH_ID"].isin(summary["BRANCH_ID"])
    ][["BRANCH_ID", "BRANCH_NAME", "NO_WORKSTATIONS", "CAPACITY_HOURS"]].copy()
    if not branches_without_data.empty:
        branches_without_data["PRUMERNA_VYTIZENOST_PCT"] = np.nan
        branches_without_data["MAX_VYTIZENOST_PCT"] = np.nan
        branches_without_data["DNI_KRITICKA"] = 0
        branches_without_data["CELKEM_HODIN"] = 0.0
        branches_without_data["POCET_DNI"] = 0
        branches_without_data["BUCKET"] = "Bez dat"
        summary = pd.concat([summary, branches_without_data], ignore_index=True)

    return summary.sort_values("PRUMERNA_VYTIZENOST_PCT", ascending=False, na_position="last").reset_index(
        drop=True
    )


def compute_capacity_growth_flags(merged: pd.DataFrame, workspaces: pd.DataFrame) -> pd.DataFrame:
    """Poukazuje na pobočky, kde je v datech využíváno víc pracovišť, než je oficiálně
    registrováno ve work_spaces.xlsx (signál, že kapacitu pobočky je třeba navýšit v evidenci)."""
    used = (
        merged.groupby(["BRANCH_ID", "BRANCH_NAME"])["WORKSTATION_ID"]
        .nunique()
        .reset_index(name="POUZITA_PRACOVISTE")
    )
    out = used.merge(
        workspaces[["BRANCH_ID", "NO_WORKSTATIONS"]], on="BRANCH_ID", how="left"
    )
    out["ROZDIL"] = out["POUZITA_PRACOVISTE"] - out["NO_WORKSTATIONS"]
    out["PREKROCENO"] = out["ROZDIL"] > 0
    return out.sort_values("ROZDIL", ascending=False).reset_index(drop=True)


def compute_activity_breakdown(merged: pd.DataFrame) -> pd.DataFrame:
    return (
        merged.groupby("ACTIVITY")
        .agg(CELKEM_HODIN=("DURATION_MIN", lambda s: s.sum() / 60), POCET=("DURATION_MIN", "count"))
        .reset_index()
        .sort_values("CELKEM_HODIN", ascending=False)
        .reset_index(drop=True)
    )


def compute_employee_summary(merged: pd.DataFrame, top_n: int | None = None) -> pd.DataFrame:
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


def detect_possible_overlaps(merged: pd.DataFrame) -> pd.DataFrame:
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
    out = out.rename(
        columns={
            "EMPLOYEE": "ZAMESTNANEC",
            "DATETIME": "START",
            "END_DATETIME": "END",
        }
    )
    return out[
        [
            "BRANCH_ID", "BRANCH_NAME", "WORKSTATION_ID",
            "PREDCHOZI_ZAMESTNANEC", "PREDCHOZI_START", "PREDCHOZI_END",
            "ZAMESTNANEC", "START", "END", "ACTIVITY",
        ]
    ].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Grafy (Plotly)
# ---------------------------------------------------------------------------

def fig_branch_utilization_bar(branch_summary: pd.DataFrame) -> go.Figure:
    df = branch_summary.dropna(subset=["PRUMERNA_VYTIZENOST_PCT"]).sort_values(
        "PRUMERNA_VYTIZENOST_PCT"
    )
    colors = df["BUCKET"].map(BUCKET_COLORS).fillna("#9E9E9E")
    fig = go.Figure(
        go.Bar(
            x=df["PRUMERNA_VYTIZENOST_PCT"],
            y=df["BRANCH_NAME"] + " (" + df["BRANCH_ID"].astype(str) + ")",
            orientation="h",
            marker_color=colors,
            text=df["PRUMERNA_VYTIZENOST_PCT"].round(1).astype(str) + " %",
            textposition="outside",
            hovertemplate="%{y}<br>Průměrná vytíženost: %{x:.1f} %<extra></extra>",
        )
    )
    fig.add_vline(x=THRESHOLD_HIGH, line_dash="dot", line_color="#F9A825")
    fig.add_vline(x=THRESHOLD_CRITICAL, line_dash="dot", line_color="#C62828")
    fig.add_vline(x=100, line_color="#616161")
    fig.update_layout(
        title="Průměrná denní vytíženost poboček vůči kapacitě",
        xaxis_title="Vytíženost (%)",
        yaxis_title=None,
        height=max(320, 28 * len(df) + 120),
        margin=dict(l=10, r=10, t=60, b=40),
        template="plotly_white",
    )
    return fig


def fig_workstation_heatmap(workstation_daily: pd.DataFrame, branch_id: int, branch_name: str) -> go.Figure:
    d = workstation_daily.loc[workstation_daily["BRANCH_ID"] == branch_id].copy()
    d["WORKSTATION_LABEL"] = "Prac. " + d["WORKSTATION_ID"].astype(str)
    pivot = d.pivot_table(index="DATE", columns="WORKSTATION_LABEL", values="UTILIZATION_PCT")
    pivot = pivot.reindex(sorted(pivot.columns, key=lambda c: int(c.split(" ")[1])), axis=1)

    fig = go.Figure(
        go.Heatmap(
            z=pivot.values,
            x=pivot.columns,
            y=[d.strftime("%d.%m.%Y") for d in pivot.index],
            colorscale=[
                [0.0, "#2E7D32"], [THRESHOLD_HIGH / 150, "#F9A825"],
                [THRESHOLD_CRITICAL / 150, "#C62828"], [1.0, "#7B0000"],
            ],
            zmin=0, zmax=150,
            colorbar=dict(title="%"),
            hovertemplate="%{x} | %{y}<br>Vytíženost: %{z:.1f} %<extra></extra>",
        )
    )
    fig.update_layout(
        title=f"Denní vytíženost pracovišť — {branch_name} ({branch_id})",
        height=max(260, 24 * len(pivot.index) + 120),
        margin=dict(l=10, r=10, t=60, b=40),
        template="plotly_white",
    )
    return fig


def fig_utilization_trend(branch_daily: pd.DataFrame, top_n: int = 8) -> go.Figure:
    top_branches = (
        branch_daily.groupby("BRANCH_NAME")["UTILIZATION_PCT"].mean().sort_values(ascending=False).head(top_n).index
    )
    d = branch_daily.loc[branch_daily["BRANCH_NAME"].isin(top_branches)]
    fig = px.line(
        d, x="DATE", y="UTILIZATION_PCT", color="BRANCH_NAME", markers=True,
        labels={"DATE": "Datum", "UTILIZATION_PCT": "Vytíženost (%)", "BRANCH_NAME": "Pobočka"},
        title=f"Trend vytíženosti v čase (top {top_n} poboček dle průměru)",
    )
    fig.add_hline(y=THRESHOLD_CRITICAL, line_dash="dot", line_color="#C62828")
    fig.update_layout(height=440, template="plotly_white", margin=dict(l=10, r=10, t=60, b=40))
    return fig


def fig_activity_mix(activity_breakdown: pd.DataFrame) -> go.Figure:
    fig = px.pie(
        activity_breakdown, names="ACTIVITY", values="CELKEM_HODIN", hole=0.45,
        title="Skladba aktivit dle odpracovaných hodin",
    )
    fig.update_traces(textinfo="percent+label", hovertemplate="%{label}<br>%{value:.1f} h (%{percent})<extra></extra>")
    fig.update_layout(height=420, template="plotly_white", margin=dict(l=10, r=10, t=60, b=10))
    return fig


def fig_employee_top(employee_summary: pd.DataFrame, n: int = 15) -> go.Figure:
    d = employee_summary.sort_values("CELKEM_HODIN", ascending=False).head(n).sort_values("CELKEM_HODIN")
    fig = go.Figure(
        go.Bar(
            x=d["CELKEM_HODIN"], y=d["EMPLOYEE"] + " — " + d["BRANCH_NAME"], orientation="h",
            marker_color="#1565C0",
            text=d["CELKEM_HODIN"].round(1).astype(str) + " h",
            textposition="outside",
        )
    )
    fig.update_layout(
        title=f"Nejvytíženější zaměstnanci (top {n} dle odpracovaných hodin)",
        xaxis_title="Hodiny celkem", yaxis_title=None,
        height=max(320, 26 * len(d) + 120), template="plotly_white",
        margin=dict(l=10, r=10, t=60, b=40),
    )
    return fig


# ---------------------------------------------------------------------------
# Sestavení HTML reportu
# ---------------------------------------------------------------------------

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


def _bucket_badge(bucket: str) -> str:
    color = BUCKET_COLORS.get(bucket, "#9E9E9E")
    return f'<span class="badge" style="background:{color}">{bucket}</span>'


def _df_to_html_table(df: pd.DataFrame, bucket_col: str | None = None, float_cols: dict | None = None) -> str:
    d = df.copy()
    float_cols = float_cols or {}
    for col, fmt in float_cols.items():
        if col in d.columns:
            d[col] = d[col].map(lambda v: (fmt.format(v) if pd.notna(v) else "—"))
    if bucket_col and bucket_col in d.columns:
        d[bucket_col] = d[bucket_col].map(_bucket_badge)
    return d.to_html(index=False, escape=False, classes="report", border=0)


def build_html_report(
    output_path: str | Path,
    *,
    period_start,
    period_end,
    activities: pd.DataFrame,
    workspaces: pd.DataFrame,
    merged: pd.DataFrame,
    branch_summary: pd.DataFrame,
    branch_daily: pd.DataFrame,
    workstation_daily: pd.DataFrame,
    growth_flags: pd.DataFrame,
    activity_breakdown: pd.DataFrame,
    employee_summary: pd.DataFrame,
    overlaps: pd.DataFrame,
    data_issues: pd.DataFrame,
    unknown_branches: pd.DataFrame,
) -> Path:
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

    def fig_html(fig: go.Figure) -> str:
        # Plotly.js se vloží celý přímo do souboru (ne přes CDN), aby report šel
        # otevřít i bez internetu (e-mail, sdílený disk, offline prohlížení).
        nonlocal plotly_cdn_included
        html = fig.to_html(
            full_html=False,
            include_plotlyjs=(True if not plotly_cdn_included else False),
            config={"displaylogo": False},
        )
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

    growth_html = ""
    exceeded = growth_flags.loc[growth_flags["PREKROCENO"]].drop(columns="PREKROCENO")
    if not exceeded.empty:
        growth_html = _df_to_html_table(
            exceeded.rename(columns={
                "BRANCH_NAME": "Pobočka", "BRANCH_ID": "ID",
                "POUZITA_PRACOVISTE": "Využitá pracoviště (dle dat)",
                "NO_WORKSTATIONS": "Registrováno ve work_spaces.xlsx", "ROZDIL": "Rozdíl",
            })
        )
    else:
        growth_html = "<p>Ve sledovaném období nebyla zjištěna žádná pobočka, kde by se využívalo víc pracovišť, než je registrováno.</p>"

    overlaps_html = (
        _df_to_html_table(
            overlaps.head(200).rename(columns={
                "BRANCH_ID": "ID pobočky", "BRANCH_NAME": "Pobočka", "WORKSTATION_ID": "Pracoviště",
                "PREDCHOZI_ZAMESTNANEC": "Předchozí zaměstnanec", "PREDCHOZI_START": "Předchozí od",
                "PREDCHOZI_END": "Předchozí do", "ZAMESTNANEC": "Zaměstnanec",
                "START": "Od", "END": "Do", "ACTIVITY": "Aktivita",
            })
        )
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
    data_quality_html = (
        f"<ul>{''.join(data_quality_bits)}</ul>" if data_quality_bits else "<p>Bez zjištěných problémů v datech.</p>"
    )

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
