#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import os
import re
import tempfile
import traceback
from pathlib import Path
from threading import RLock
from typing import Iterable
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
from dash import Dash, Input, Output, State, dcc, html
from dash.exceptions import PreventUpdate


STATE_TO_REGION = {
    "AK": "West", "AZ": "West", "CA": "West", "CO": "West", "HI": "West", "ID": "West", "MT": "West", "NV": "West", "NM": "West", "OR": "West", "UT": "West", "WA": "West", "WY": "West",
    "AR": "Central", "IA": "Central", "IL": "Central", "IN": "Central", "KS": "Central", "LA": "Central", "MI": "Central", "MN": "Central", "MO": "Central", "ND": "Central", "NE": "Central", "OH": "Central", "OK": "Central", "SD": "Central", "TX": "Central", "WI": "Central",
    "CT": "Northeast", "DC": "Northeast", "DE": "Northeast", "MA": "Northeast", "MD": "Northeast", "ME": "Northeast", "NH": "Northeast", "NJ": "Northeast", "NY": "Northeast", "PA": "Northeast", "RI": "Northeast", "VA": "Northeast", "VT": "Northeast", "WV": "Northeast",
    "AL": "Southeast", "FL": "Southeast", "GA": "Southeast", "KY": "Southeast", "MS": "Southeast", "NC": "Southeast", "SC": "Southeast", "TN": "Southeast",
}

DATA_DIR = Path("data")
DEFAULT_UPLOAD_FILENAME = "Cases_Final_Dashboard_CURRENT.xlsx"
RAILWAY_DATA_DIR = Path("/data")
DATA_DIR = Path(os.environ.get("DATA_DIR", str(RAILWAY_DATA_DIR if RAILWAY_DATA_DIR.exists() else DATA_DIR)))
CURRENT_FILE = Path(os.environ.get("CURRENT_FILE", str(DATA_DIR / DEFAULT_UPLOAD_FILENAME)))
MAX_UPLOAD_SIZE_BYTES = 25 * 1024 * 1024
APP_TIMEZONE = ZoneInfo(os.environ.get("APP_TIMEZONE", "America/New_York"))
SHORT_DESCRIPTION_MAX_LEN = 160

DATA_LOCK = RLock()
DATA_CACHE = {
    "dispatch_df": None,
    "source_path": None,
}


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u200c", " ").replace("\xa0", " ")).strip()


def infer_header_row(df: pd.DataFrame, required_tokens: Iterable[str], max_scan: int = 8) -> int:
    for idx in range(min(max_scan, len(df))):
        row = [clean_text(v).lower() for v in df.iloc[idx].tolist()]
        if all(any(tok.lower() == cell for cell in row) for tok in required_tokens):
            return idx
    return 0


def read_sheet_with_inferred_header(
    path: Path,
    sheet_name: str,
    required_tokens: Iterable[str],
    max_scan: int = 8,
    keep_columns: Iterable[str] | None = None,
) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=max_scan)
    header_row = infer_header_row(raw, required_tokens, max_scan=max_scan)
    usecols = None
    if keep_columns:
        keep = {clean_text(c).lower() for c in keep_columns}
        usecols = lambda c: clean_text(c).lower() in keep
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row, usecols=usecols)
    df.columns = [clean_text(c) for c in df.columns]
    return df


def current_timestamp() -> pd.Timestamp:
    return pd.Timestamp.now(tz=APP_TIMEZONE)


def current_day_start() -> pd.Timestamp:
    return current_timestamp().normalize()


def file_mtime_timestamp(path: Path) -> pd.Timestamp:
    return pd.Timestamp.fromtimestamp(path.stat().st_mtime, tz=APP_TIMEZONE)


def friendly_dt(dt) -> str:
    if pd.isna(dt) or dt is None:
        return "-"
    ts = pd.to_datetime(dt)
    return ts.strftime("%b %d, %Y at %I:%M %p").replace(" 0", " ")


def friendly_sched(dt) -> str:
    if pd.isna(dt) or dt is None or clean_text(dt) == "":
        return ""
    ts = pd.to_datetime(dt, errors="coerce")
    if pd.isna(ts):
        return clean_text(dt)
    if ts.hour == 0 and ts.minute == 0 and ts.second == 0:
        return ts.strftime("%b %d, %Y").replace(" 0", " ")
    return ts.strftime("%b %d, %I:%M %p").replace(" 0", " ")


def short_preview(text: str, limit: int = 120) -> str:
    txt = clean_text(text)
    if len(txt) <= limit:
        return txt
    return txt[: max(0, limit - 1)].rstrip() + "..."


def empty_dispatch_df() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "WO", "CS", "Company", "Location", "WO State", "Assignment Group", "State", "City",
            "Short Description", "Scheduled Start", "Region", "Scheduled Date", "Week Start",
            "Escalation Status", "Escalated Flag", "Warning Flag", "RCE Remote Support Flag",
            "Visit Count",
        ]
    )


def normalize_case_flags(path: Path) -> pd.DataFrame:
    try:
        case_df = read_sheet_with_inferred_header(
            path,
            "Per_Case_Dashboard",
            ["Case number", "Visit List"],
            keep_columns=[
                "Case number", "State / Province", "Escalation Status", "Visit Count",
                "Summary of case and visits", "Work notes", "Next Steps",
            ],
        )
    except Exception:
        return pd.DataFrame(columns=["CS", "Escalation Status", "Escalated Flag", "Warning Flag", "RCE Remote Support Flag", "Visit Count", "Region"])

    rename_map = {
        "Case number": "CS",
        "State / Province": "State",
    }
    case_df = case_df.rename(columns={k: v for k, v in rename_map.items() if k in case_df.columns})
    for col in ["CS", "State", "Escalation Status"]:
        if col not in case_df.columns:
            case_df[col] = ""
    for col in case_df.columns:
        if case_df[col].dtype == object:
            case_df[col] = case_df[col].map(clean_text)

    status = case_df["Escalation Status"].astype(str).str.lower()
    note_cols = [c for c in ["Summary of case and visits", "Work notes", "Next Steps"] if c in case_df.columns]
    notes = case_df[note_cols].astype(str).agg(" ".join, axis=1).str.lower() if note_cols else pd.Series("", index=case_df.index)
    case_df["Escalated Flag"] = status.str.contains("escalated", na=False)
    case_df["Warning Flag"] = status.str.contains("warning", na=False)
    case_df["RCE Remote Support Flag"] = notes.str.contains(r"rce|remote support|remote engineering", regex=True, na=False)
    case_df["Visit Count"] = pd.to_numeric(case_df.get("Visit Count", 0), errors="coerce").fillna(0).astype(int)
    case_df["Region"] = case_df["State"].map(STATE_TO_REGION).fillna("Non-US / Unknown")
    return case_df[["CS", "Escalation Status", "Escalated Flag", "Warning Flag", "RCE Remote Support Flag", "Visit Count", "Region"]].drop_duplicates("CS")


def normalize_dispatch_df(path: Path) -> pd.DataFrame:
    with pd.ExcelFile(path) as xl:
        first_sheet = xl.sheet_names[0]
    raw = pd.read_excel(path, sheet_name=first_sheet, header=None, nrows=12)
    header_row = infer_header_row(raw, ["Work order", "Company", "Location"], max_scan=12)
    dispatch_keep_columns = {
        "Work order", "Work Order", "Number2",
        "Case Number", "Case number", "Number", "CS",
        "Work Order State", "WO State",
        "Company", "Location",
        "Work order Assignment Group", "Work Order Assignment Group", "Assignment Group",
        "Scheduled start", "Scheduled Start",
        "Short description", "Short Description",
        "State / Province", "State",
        "City", "City / Town", "Town",
        "Summary of case and visits", "Summary", "Description", "Work notes",
    }
    dispatch_keep = {clean_text(c).lower() for c in dispatch_keep_columns}
    df = pd.read_excel(
        path,
        sheet_name=first_sheet,
        header=header_row,
        usecols=lambda c: clean_text(c).lower() in dispatch_keep,
    )
    df.columns = [clean_text(c) for c in df.columns]
    rename_candidates = {
        "Work order": "WO", "Work Order": "WO", "Number2": "WO",
        "Case Number": "CS", "Case number": "CS", "Number": "CS", "CS": "CS",
        "Work Order State": "WO State", "WO State": "WO State",
        "Company": "Company", "Location": "Location",
        "Work order Assignment Group": "Assignment Group", "Work Order Assignment Group": "Assignment Group", "Assignment Group": "Assignment Group",
        "Scheduled start": "Scheduled Start", "Scheduled Start": "Scheduled Start",
        "Short description": "Short Description", "Short Description": "Short Description",
        "State / Province": "State", "State": "State",
        "City": "City", "City / Town": "City", "Town": "City",
    }
    df = df.rename(columns={k: v for k, v in rename_candidates.items() if k in df.columns})
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(clean_text)
    for col in ["WO", "CS", "Company", "Location", "WO State", "Assignment Group", "State", "City", "Short Description"]:
        if col not in df.columns:
            df[col] = ""
    df["Scheduled Start"] = pd.to_datetime(df["Scheduled Start"], errors="coerce") if "Scheduled Start" in df.columns else pd.NaT
    if df["CS"].eq("").all():
        summary_cols = [c for c in df.columns if str(c).lower() in {"summary of case and visits", "summary", "description", "short description", "work notes"}]
        if summary_cols:
            df["CS"] = df[summary_cols[0]].astype(str).str.extract(r"(CS\d+)", expand=False).fillna("")
    df["Region"] = df["State"].map(STATE_TO_REGION).fillna("Non-US / Unknown")
    df["Scheduled Date"] = pd.to_datetime(df["Scheduled Start"], errors="coerce").dt.normalize()
    weekday = pd.to_datetime(df["Scheduled Date"], errors="coerce").dt.weekday
    df["Week Start"] = df["Scheduled Date"] - pd.to_timedelta(weekday.fillna(0), unit="D")
    return df


def trim_dispatch_columns(dispatch_df: pd.DataFrame) -> pd.DataFrame:
    required_cols = [
        "WO", "CS", "Company", "Location", "WO State", "Assignment Group", "State", "City",
        "Short Description", "Scheduled Start", "Region", "Scheduled Date", "Week Start",
        "Escalation Status", "Escalated Flag", "Warning Flag", "RCE Remote Support Flag",
        "Visit Count",
    ]
    out = dispatch_df.copy()
    for col in required_cols:
        if col not in out.columns:
            if col in {"Scheduled Start", "Scheduled Date", "Week Start"}:
                out[col] = pd.NaT
            elif col in {"Escalated Flag", "Warning Flag", "RCE Remote Support Flag"}:
                out[col] = False
            elif col == "Visit Count":
                out[col] = 0
            else:
                out[col] = ""
    out["Short Description"] = out["Short Description"].map(lambda x: short_preview(x, SHORT_DESCRIPTION_MAX_LEN))
    return out[required_cols]


def build_dispatch_dataset(case_file: Path) -> pd.DataFrame:
    dispatch_df = normalize_dispatch_df(case_file)
    case_flags = normalize_case_flags(case_file)
    if not case_flags.empty:
        dispatch_df = dispatch_df.merge(case_flags, on="CS", how="left", suffixes=("", "_case"))
        if "Region_case" in dispatch_df.columns:
            dispatch_df["Region"] = dispatch_df["Region"].where(dispatch_df["Region"].ne("Non-US / Unknown"), dispatch_df["Region_case"])
            dispatch_df = dispatch_df.drop(columns=["Region_case"])
    for col in ["Escalated Flag", "Warning Flag", "RCE Remote Support Flag"]:
        dispatch_df[col] = pd.Series(dispatch_df.get(col, False)).astype("boolean").fillna(False).astype(bool)
    dispatch_df["Visit Count"] = pd.to_numeric(dispatch_df.get("Visit Count", 0), errors="coerce").fillna(0).astype(np.int32)
    dispatch_df = trim_dispatch_columns(dispatch_df)
    for col in ["Region", "State", "WO State", "Assignment Group", "Company"]:
        if col in dispatch_df.columns and dispatch_df[col].dtype == object:
            dispatch_df[col] = dispatch_df[col].astype("category")
    return dispatch_df


def resolve_source_path(case_file: Path | None) -> Path | None:
    if CURRENT_FILE.exists():
        return CURRENT_FILE
    if case_file and Path(case_file).exists():
        return Path(case_file)
    return None


def load_dispatch_data(case_file: Path | None) -> tuple[pd.DataFrame, Path | None]:
    source_path = resolve_source_path(case_file)
    if not source_path:
        return empty_dispatch_df(), None
    return build_dispatch_dataset(source_path), source_path


def update_data_cache(dispatch_df: pd.DataFrame, source_path: Path | None) -> None:
    with DATA_LOCK:
        DATA_CACHE["dispatch_df"] = dispatch_df
        DATA_CACHE["source_path"] = source_path


def get_cached_dispatch_df() -> pd.DataFrame:
    with DATA_LOCK:
        df = DATA_CACHE.get("dispatch_df")
    if isinstance(df, pd.DataFrame):
        return df
    return empty_dispatch_df()


def dispatch_filter_options(dispatch_df: pd.DataFrame) -> tuple[list[str], list[str], list[str], list[str]]:
    if dispatch_df.empty:
        return [], [], [], []
    week_options = sorted(
        [str(pd.to_datetime(x).date()) for x in pd.to_datetime(dispatch_df["Week Start"], errors="coerce").dropna().unique().tolist()]
    )
    wo_state_options = sorted([x for x in dispatch_df.get("WO State", pd.Series(dtype=object)).dropna().unique().tolist() if clean_text(x)])
    company_options = sorted([x for x in dispatch_df.get("Company", pd.Series(dtype=object)).dropna().unique().tolist() if clean_text(x)])
    region_options = sorted([x for x in dispatch_df.get("Region", pd.Series(dtype=object)).dropna().unique().tolist() if clean_text(x)])
    return week_options, wo_state_options, company_options, region_options


def current_week_start_from_data(dispatch_df: pd.DataFrame):
    today = current_day_start().tz_localize(None)
    current = today - pd.Timedelta(days=today.weekday())
    vals = pd.to_datetime(dispatch_df["Week Start"], errors="coerce").dropna()
    if vals.empty:
        return None
    if current in set(vals):
        return str(current.date())
    future = vals[vals >= current]
    if not future.empty:
        return str(future.min().date())
    return str(vals.max().date())


def validate_upload(filename: str | None, bin_data: bytes) -> None:
    safe_name = clean_text(filename or "")
    if not safe_name.lower().endswith(".xlsx"):
        raise ValueError("Only .xlsx uploads are allowed.")
    if len(bin_data) > MAX_UPLOAD_SIZE_BYTES:
        raise ValueError(f"File is too large. Max size is {MAX_UPLOAD_SIZE_BYTES // (1024 * 1024)} MB.")


def save_uploaded_workbook(bin_data: bytes) -> Path:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix="upload_", suffix=".xlsx", dir=str(DATA_DIR))
    try:
        with os.fdopen(fd, "wb") as tmp_file:
            tmp_file.write(bin_data)
        temp_path = Path(tmp_name)
        build_dispatch_dataset(temp_path)
        os.replace(temp_path, CURRENT_FILE)
        return CURRENT_FILE
    except Exception:
        try:
            Path(tmp_name).unlink(missing_ok=True)
        except Exception:
            pass
        raise


def build_dispatch_rows(
    dispatch_df: pd.DataFrame,
    week_start: str,
    wo_states: list[str],
    flagged_only: list[str],
    company_filter: list[str],
    region_filter: list[str],
    flag_filter: list[str],
    show_weekend: list[str],
):
    df = dispatch_df
    if week_start:
        df = df[pd.to_datetime(df["Week Start"], errors="coerce").dt.strftime("%Y-%m-%d") == week_start]
    if wo_states:
        df = df[df["WO State"].isin(wo_states)]
    if company_filter:
        df = df[df["Company"].isin(company_filter)]
    if region_filter:
        df = df[df["Region"].isin(region_filter)]
    if df.empty:
        return []
    df = df.assign(DayName=pd.to_datetime(df["Scheduled Date"], errors="coerce").dt.day_name().str[:3])
    weekday_order = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    if show_weekend and "show" in show_weekend:
        weekday_order += ["Sat", "Sun"]
    df = df[df["DayName"].isin(weekday_order)]
    rows = []
    for company, grp in df.groupby("Company", dropna=False, observed=True):
        if grp.empty:
            continue
        days_present = sorted(set(grp["DayName"]))
        day_nums = sorted({weekday_order.index(d) for d in days_present})
        consecutive = any((b - a) == 1 for a, b in zip(day_nums, day_nums[1:]))
        same_case_multi_day = grp.groupby("CS")["DayName"].nunique().max() > 1 if ("CS" in grp.columns and not grp["CS"].isna().all()) else False
        multi_group = grp["Assignment Group"].fillna("").replace("", np.nan).nunique(dropna=True) > 1
        escalated = bool(grp["Escalated Flag"].fillna(False).any())
        warning = bool(grp["Warning Flag"].fillna(False).any())
        rce = bool(grp["RCE Remote Support Flag"].fillna(False).any())
        same_site_consecutive = False
        for _, site_grp in grp.groupby("Location", dropna=False, observed=True):
            site_days = sorted({weekday_order.index(d) for d in site_grp["DayName"].dropna().tolist()})
            if any((b - a) == 1 for a, b in zip(site_days, site_days[1:])):
                same_site_consecutive = True
                break
        flags = []
        if consecutive:
            flags.append("Consecutive")
        if same_site_consecutive:
            flags.append("Same Site")
        if same_case_multi_day:
            flags.append("Same Case")
        if multi_group:
            flags.append("Multi-Group")
        if escalated:
            flags.append("Escalated")
        if warning:
            flags.append("Warning")
        if rce:
            flags.append("RCE")
        if flagged_only and "show" in flagged_only and not flags:
            continue
        if flag_filter and not any(f in flags for f in flag_filter):
            continue
        cells = {}
        for day in weekday_order:
            day_rows = grp[grp["DayName"] == day].sort_values(["Scheduled Start", "WO"])
            chips = []
            for _, r in day_rows.iterrows():
                sched_str = friendly_sched(r.get("Scheduled Start"))
                summary_txt = short_preview(r.get("Short Description", ""), 120)
                tags = []
                if clean_text(r.get("WO State", "")).lower() == "assigned":
                    tags.append("Assigned")
                if bool(r.get("Escalated Flag", False)):
                    tags.append("Escalated")
                elif bool(r.get("Warning Flag", False)):
                    tags.append("Warning")
                if bool(r.get("RCE Remote Support Flag", False)):
                    tags.append("RCE")
                chip_class = "wo-chip"
                if bool(r.get("Escalated Flag", False)):
                    chip_class += " escalated"
                elif bool(r.get("Warning Flag", False)):
                    chip_class += " warning"
                elif bool(r.get("RCE Remote Support Flag", False)):
                    chip_class += " rce"
                city_state = ", ".join([x for x in [clean_text(r.get("City", "")), clean_text(r.get("State", ""))] if x])
                chips.append(html.Div([
                    html.Div(f'{clean_text(r.get("WO", ""))} | {clean_text(r.get("CS", ""))}', className="chip-line chip-strong"),
                    html.Div(clean_text(r.get("Location", "")), className="chip-line"),
                    html.Div(city_state, className="chip-line chip-small") if city_state else None,
                    html.Div(clean_text(r.get("Assignment Group", "")), className="chip-line chip-small"),
                    html.Div(sched_str, className="chip-line chip-small") if sched_str else None,
                    html.Div(summary_txt, className="chip-line chip-small chip-summary") if summary_txt else None,
                    html.Div(" | ".join(tags), className="chip-tags") if tags else None,
                ], className=chip_class))
            cells[day] = chips
        row_class = "dispatch-row"
        if flags:
            row_class += " flagged"
        if escalated:
            row_class += " escalated-row"
        elif warning:
            row_class += " warning-row"
        elif rce:
            row_class += " rce-row"
        rows.append(html.Tr(
            [
                html.Td(company, className="sticky-col sticky-col-1"),
                html.Td(clean_text(grp["Region"].dropna().iloc[0] if not grp["Region"].dropna().empty else ""), className="sticky-col sticky-col-2"),
                html.Td(int(grp["WO"].nunique()), className="sticky-col sticky-col-3"),
                html.Td(int(grp["CS"].nunique()), className="sticky-col sticky-col-4"),
                html.Td(int(grp["Location"].nunique()), className="sticky-col sticky-col-5"),
                html.Td(int(grp["Assignment Group"].replace("", np.nan).nunique(dropna=True)), className="sticky-col sticky-col-6"),
                html.Td(" | ".join(flags), className="sticky-col sticky-col-7"),
            ] + [html.Td(cells.get(day, [])) for day in weekday_order],
            className=row_class,
        ))
    return rows


def build_dispatch_layout(dispatch_df: pd.DataFrame, source_path: Path | None) -> html.Div:
    source_file_mtime = friendly_dt(file_mtime_timestamp(source_path)) if source_path else "No workbook loaded"
    current_week = current_week_start_from_data(dispatch_df)
    week_options, wo_state_options, company_options, region_options = dispatch_filter_options(dispatch_df)
    dispatch_flag_options = ["Consecutive", "Same Site", "Same Case", "Multi-Group", "Escalated", "Warning", "RCE"]
    upload_prompt = "Upload Cases_Final_Dashboard_CURRENT.xlsx to load the board." if source_path is None else ""

    return html.Div([
        dcc.Store(id="refresh-store", data={"last_refresh": friendly_dt(current_timestamp())}),
        dcc.Store(id="upload-signal", data={"ts": None}),
        html.Div([
            html.H1("Weekly Client Dispatch Board"),
            html.Div(id="timestamp-line", children=f"Source file updated: {source_file_mtime} | Board refreshed: {friendly_dt(current_timestamp())}"),
            html.Div(id="source-status", children=upload_prompt),
        ], className="header"),
        html.Div([
            dcc.Dropdown(id="week-select", options=[{"label": x, "value": x} for x in week_options], value=current_week, clearable=False, style={"width": "180px"}),
            dcc.Dropdown(id="wo-state-select", options=[{"label": x, "value": x} for x in wo_state_options], value=["Assigned"] if "Assigned" in wo_state_options else [], multi=True, placeholder="Work Order State", style={"minWidth": "260px"}),
            dcc.Dropdown(id="dispatch-company-filter", options=[{"label": x, "value": x} for x in company_options], value=[], multi=True, placeholder="Company", style={"minWidth": "220px"}),
            dcc.Dropdown(id="dispatch-region-filter", options=[{"label": x, "value": x} for x in region_options], value=[], multi=True, placeholder="Region", style={"minWidth": "180px"}),
            dcc.Dropdown(id="dispatch-flag-filter", options=[{"label": x, "value": x} for x in dispatch_flag_options], value=[], multi=True, placeholder="Flags", style={"minWidth": "220px"}),
            dcc.Checklist(id="flagged-only", options=[{"label": "Show only flagged rows", "value": "show"}], value=[]),
            dcc.Checklist(id="show-weekend", options=[{"label": "Show Saturday / Sunday", "value": "show"}], value=[]),
            html.Button("Refresh Data", id="refresh-button", n_clicks=0),
            dcc.Upload(id="upload-data", children=html.Button("Upload Dataset"), multiple=False, style={"display": "inline-block"}),
            html.Div(id="upload-status", className="upload-status"),
        ], className="toolbar"),
        html.Div("Clients appear only if they have scheduled work in the selected week.", className="filters-label"),
        html.Div(id="dispatch-board-container"),
    ], className="wrapper")


def create_app(case_file: Path | None) -> Dash:
    dispatch_df, source_path = load_dispatch_data(case_file)
    update_data_cache(dispatch_df, source_path)
    app = Dash(__name__)
    app.title = "Weekly Client Dispatch Board"
    app.layout = build_dispatch_layout(dispatch_df, source_path)

    @app.callback(
        Output("refresh-store", "data"),
        Output("timestamp-line", "children"),
        Output("source-status", "children"),
        Input("refresh-button", "n_clicks"),
        Input("upload-signal", "data"),
        prevent_initial_call=True,
    )
    def refresh_data(n_clicks, upload_signal):
        dispatch, src = load_dispatch_data(case_file)
        update_data_cache(dispatch, src)
        now_txt = friendly_dt(current_timestamp())
        prompt = ""
        try:
            source_txt = friendly_dt(file_mtime_timestamp(src))
        except Exception:
            source_txt = "No workbook loaded"
            prompt = "Upload Cases_Final_Dashboard_CURRENT.xlsx to load the board."
        return {"last_refresh": now_txt}, f"Source file updated: {source_txt} | Board refreshed: {now_txt}", prompt

    @app.callback(
        Output("upload-signal", "data"),
        Output("upload-status", "children"),
        Input("upload-data", "contents"),
        State("upload-data", "filename"),
        prevent_initial_call=True,
    )
    def handle_upload(contents, filename):
        if not contents:
            raise PreventUpdate
        try:
            _, b64 = contents.split(",", 1) if "," in contents else ("", contents)
            bin_data = base64.b64decode(b64)
            validate_upload(filename, bin_data)
            saved_path = save_uploaded_workbook(bin_data)
            ts = friendly_dt(current_timestamp())
            return {"ts": ts}, f"Saved {DEFAULT_UPLOAD_FILENAME} to {saved_path} at {ts}"
        except Exception as exc:
            return {"ts": None}, f"Upload failed: {exc}"

    @app.callback(
        Output("dispatch-board-container", "children"),
        Input("refresh-store", "data"),
        Input("week-select", "value"),
        Input("wo-state-select", "value"),
        Input("flagged-only", "value"),
        Input("dispatch-company-filter", "value"),
        Input("dispatch-region-filter", "value"),
        Input("dispatch-flag-filter", "value"),
        Input("show-weekend", "value"),
    )
    def render_dispatch_board(refresh_state, week_start, wo_states, flagged_only, company_filter, region_filter, flag_filter, show_weekend):
        df = get_cached_dispatch_df()
        rows = build_dispatch_rows(df, week_start, wo_states or [], flagged_only or [], company_filter or [], region_filter or [], flag_filter or [], show_weekend or [])
        day_headers = ["Mon", "Tue", "Wed", "Thu", "Fri"]
        if show_weekend and "show" in show_weekend:
            day_headers += ["Sat", "Sun"]
        header = html.Thead(html.Tr(
            [
                html.Th("Company", className="sticky-col sticky-col-1"),
                html.Th("Region", className="sticky-col sticky-col-2"),
                html.Th("Weekly WOs", className="sticky-col sticky-col-3"),
                html.Th("Unique Cases", className="sticky-col sticky-col-4"),
                html.Th("Unique Sites", className="sticky-col sticky-col-5"),
                html.Th("Unique Groups", className="sticky-col sticky-col-6"),
                html.Th("Flags", className="sticky-col sticky-col-7"),
            ] + [html.Th(d) for d in day_headers]
        ))
        if not rows:
            return html.Div("No scheduled rows for the selected filters.", className="panel")
        return html.Div([
            html.Div(html.Div(className="dispatch-top-scroll-inner"), className="dispatch-top-scroll", id="dispatch-top-scroll"),
            html.Div(html.Table([header, html.Tbody(rows)], className="dispatch-table"), className="dispatch-wrap", id="dispatch-wrap"),
        ], className="panel")

    @app.callback(
        Output("week-select", "options"),
        Output("week-select", "value"),
        Output("wo-state-select", "options"),
        Output("wo-state-select", "value"),
        Output("dispatch-company-filter", "options"),
        Output("dispatch-company-filter", "value"),
        Output("dispatch-region-filter", "options"),
        Output("dispatch-region-filter", "value"),
        Input("refresh-store", "data"),
        State("week-select", "value"),
        State("wo-state-select", "value"),
        State("dispatch-company-filter", "value"),
        State("dispatch-region-filter", "value"),
    )
    def sync_dispatch_filters(refresh_state, week_value, wo_values, company_value, region_value):
        df = get_cached_dispatch_df()
        week_vals, wo_vals, company_vals, region_vals = dispatch_filter_options(df)
        next_week = week_value if week_value in week_vals else current_week_start_from_data(df)
        next_wo = [x for x in (wo_values or []) if x in wo_vals]
        if not next_wo and "Assigned" in wo_vals:
            next_wo = ["Assigned"]
        next_company = [x for x in (company_value or []) if x in company_vals]
        next_region = [x for x in (region_value or []) if x in region_vals]
        return (
            [{"label": x, "value": x} for x in week_vals], next_week,
            [{"label": x, "value": x} for x in wo_vals], next_wo,
            [{"label": x, "value": x} for x in company_vals], next_company,
            [{"label": x, "value": x} for x in region_vals], next_region,
        )

    app.index_string = """
    <!DOCTYPE html><html><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}
    <style>
    body { font-family: Arial, Helvetica, sans-serif; margin: 0; background: #f4f7fb; color: #1f2937; }
    .header { background: #0f3b5f; color: white; padding: 20px 24px; }
    .header h1 { margin: 0 0 8px 0; font-size: 28px; letter-spacing: 0; }
    .wrapper { padding: 0 0 20px 0; }
    .toolbar { display:flex; gap:12px; align-items:center; margin: 16px 20px 12px 20px; flex-wrap: wrap; }
    .filters-label { margin: 0 20px 12px 20px; color:#5b6675; }
    .upload-status { color:#1f2937; max-width: 520px; }
    button { background:#0f3b5f; color:#fff; border:2px solid #0a2942; border-radius:9px; padding:8px 12px; cursor:pointer; font-weight:600; }
    .panel { background:#fff; border-radius:14px; padding:12px; box-shadow:0 1px 7px rgba(0,0,0,.08); margin: 0 20px; }
    .dispatch-wrap { margin: 0 0 8px 0; overflow:auto; cursor: grab; }
    .dispatch-wrap.dragging { cursor: grabbing; user-select:none; }
    .dispatch-top-scroll { margin: 0 0 8px 0; overflow-x:auto; overflow-y:hidden; height:16px; position:relative; z-index:1; }
    .dispatch-top-scroll-inner { height:1px; }
    .dispatch-table { border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; font-size: 12px; table-layout: fixed; }
    .dispatch-table th, .dispatch-table td { border-bottom:3px solid #dbe4ef; border-right:3px solid #edf2f7; padding:6px; vertical-align: top; background:#fff; min-width:135px; }
    .dispatch-table thead th { position: sticky; top: 0; z-index: 4; background:#eaf1f8; border-top:3px solid #dbe4ef; }
    .dispatch-table .sticky-col { position: sticky; background:#f8fbff; z-index: 5; min-width:92px; max-width:92px; }
    .dispatch-table .sticky-col-1 { left: 0; min-width:160px; max-width:160px; z-index: 6; }
    .dispatch-table .sticky-col-2 { left: 160px; }
    .dispatch-table .sticky-col-3 { left: 252px; }
    .dispatch-table .sticky-col-4 { left: 344px; }
    .dispatch-table .sticky-col-5 { left: 436px; }
    .dispatch-table .sticky-col-6 { left: 528px; }
    .dispatch-table .sticky-col-7 { left: 620px; min-width:110px; max-width:110px; }
    .wo-chip { border:1px solid #dbe7f3; border-left:4px solid #4c78a8; border-radius:8px; padding:4px 6px; margin-bottom:5px; background:#fafdff; width:100%; box-sizing:border-box; display:block; }
    .wo-chip.escalated { border-left-color:#d62728; background:#fff5f5; }
    .wo-chip.warning { border-left-color:#ff7f0e; background:#fff9ef; }
    .wo-chip.rce { border-left-color:#1f77b4; background:#f4f9ff; }
    .chip-line { line-height:1.1; margin-bottom:1px; }
    .chip-strong { font-weight:700; }
    .chip-small { font-size:10px; color:#5b6675; }
    .chip-tags { font-size:11px; color:#334e68; margin-top:4px; }
    .chip-summary { color:#44556b; white-space: normal; overflow-wrap: anywhere; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; }
    .dispatch-row.flagged > td.sticky-col-7 { background:#fff9df; }
    .dispatch-row.escalated-row > td.sticky-col-1 { border-left:4px solid #d62728; }
    .dispatch-row.warning-row > td.sticky-col-1 { border-left:4px solid #ff7f0e; }
    .dispatch-row.rce-row > td.sticky-col-1 { border-left:4px solid #1f77b4; }
    </style></head><body>{%app_entry%}
    <script>
    (function() {
      function setupDispatchPan() {
        const wrap = document.getElementById('dispatch-wrap');
        const top = document.getElementById('dispatch-top-scroll');
        if (!wrap || !top) return;
        const table = wrap.querySelector('.dispatch-table');
        const inner = top.querySelector('.dispatch-top-scroll-inner');
        if (!table || !inner) return;

        inner.style.width = table.scrollWidth + 'px';

        if (!wrap.dataset.syncBound) {
          wrap.addEventListener('scroll', () => {
            if (Math.abs(top.scrollLeft - wrap.scrollLeft) > 1) top.scrollLeft = wrap.scrollLeft;
          });
          top.addEventListener('scroll', () => {
            if (Math.abs(wrap.scrollLeft - top.scrollLeft) > 1) wrap.scrollLeft = top.scrollLeft;
          });
          wrap.dataset.syncBound = '1';
        }

        function bindDrag(el, target) {
          if (el.dataset.dragBound) return;
          let isDown = false, startX = 0, startScroll = 0;
          el.addEventListener('mousedown', (e) => {
            if (e.target.closest('.wo-chip')) return;
            isDown = true;
            startX = e.pageX;
            startScroll = target.scrollLeft;
            wrap.classList.add('dragging');
            document.body.style.userSelect = 'none';
          });
          window.addEventListener('mouseup', () => {
            isDown = false;
            wrap.classList.remove('dragging');
            document.body.style.userSelect = '';
          });
          el.addEventListener('mousemove', (e) => {
            if (!isDown) return;
            e.preventDefault();
            const dx = e.pageX - startX;
            target.scrollLeft = startScroll - dx;
          });
          el.dataset.dragBound = '1';
        }

        bindDrag(wrap, wrap);
        bindDrag(top, top);
      }

      function initDispatchObserver() {
        setupDispatchPan();
        const root = document.getElementById('react-entry-point') || document.body;
        const obs = new MutationObserver(() => {
          window.requestAnimationFrame(setupDispatchPan);
        });
        obs.observe(root, { childList: true, subtree: true });
        window.addEventListener('resize', setupDispatchPan);
      }

      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initDispatchObserver);
      } else {
        initDispatchObserver();
      }
    })();
    </script>
<footer>{%config%}{%scripts%}{%renderer%}</footer></body></html>
    """
    return app


_env_case = os.environ.get("CASE_FILE")
try:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    app = create_app(Path(_env_case).resolve() if _env_case else None)
    server = app.server
except Exception as exc:
    traceback.print_exc()
    raise RuntimeError("Failed to initialize Weekly Client Dispatch Board") from exc


def main():
    ap = argparse.ArgumentParser(description="Standalone Weekly Client Dispatch Board.")
    ap.add_argument("--case-file", required=False)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8052)
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()
    case_arg = args.case_file or os.environ.get("CASE_FILE")
    case_path = Path(case_arg).resolve() if case_arg else None
    local_app = create_app(case_path)
    local_app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
