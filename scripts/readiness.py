from __future__ import annotations

from datetime import time
from pathlib import Path

import pandas as pd


DEFAULT_PLAN_UNITS = 14


def _time_text(value: object) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    return str(value)


def add_event_datetimes(df: pd.DataFrame) -> pd.DataFrame:
    data = df.copy()
    date_text = pd.to_datetime(data["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    time_text = data["Awal"].apply(_time_text)
    data["BD_Start"] = pd.to_datetime(date_text + " " + time_text, errors="coerce")
    data["BD_End"] = data["BD_Start"] + pd.to_timedelta(
        pd.to_numeric(data["Duration_Real"], errors="coerce"), unit="h"
    )
    return data


def load_or_create_readiness_plan(
    path: str | Path,
    dates: pd.Series,
    default_plan_units: int = DEFAULT_PLAN_UNITS,
) -> pd.DataFrame:
    plan_path = Path(path)
    if plan_path.exists():
        try:
            xls = pd.ExcelFile(plan_path)
            if "plan" in xls.sheet_names:
                plan = pd.read_excel(plan_path, sheet_name="plan")
            else:
                plan = pd.read_excel(plan_path, sheet_name=xls.sheet_names[0])
        except Exception:
            plan = pd.DataFrame(columns=["Date", "Plan_Ready_Units"])

        if "Date" not in plan.columns:
            plan = pd.DataFrame(columns=["Date", "Plan_Ready_Units"])
        if "Plan_Ready_Units" not in plan.columns:
            if "Plan" in plan.columns:
                plan = plan.rename(columns={"Plan": "Plan_Ready_Units"})
            else:
                plan["Plan_Ready_Units"] = default_plan_units

        plan["Date"] = pd.to_datetime(plan["Date"], errors="coerce").dt.date
        plan["Plan_Ready_Units"] = pd.to_numeric(plan["Plan_Ready_Units"], errors="coerce").fillna(default_plan_units)
        plan = plan.dropna(subset=["Date"]).copy()
        if not plan.empty:
            return plan

    clean_dates = pd.to_datetime(dates, errors="coerce").dropna()
    if clean_dates.empty:
        plan = pd.DataFrame(columns=["Date", "Plan_Ready_Units"])
    else:
        all_dates = pd.date_range(clean_dates.min().date(), clean_dates.max().date(), freq="D")
        plan = pd.DataFrame(
            {
                "Date": all_dates.date,
                "Plan_Ready_Units": default_plan_units,
            }
        )

    plan_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(plan_path, engine="openpyxl") as writer:
        plan.to_excel(writer, sheet_name="plan", index=False)
    return plan


def build_hourly_readiness(
    df: pd.DataFrame,
    plan: pd.DataFrame,
    ready_threshold_minutes: int = 30,
    units: list[str] | None = None,
) -> pd.DataFrame:
    events = add_event_datetimes(df).dropna(subset=["BD_Start", "BD_End", "CN Unit"]).copy()
    if units is None:
        units = sorted(events["CN Unit"].dropna().astype(str).unique())
    else:
        units = sorted({str(unit) for unit in units if pd.notna(unit)})
    dates = pd.to_datetime(plan["Date"], errors="coerce").dropna()

    if not units or dates.empty:
        return pd.DataFrame(
            columns=[
                "Slot_Start",
                "Date",
                "Hour",
                "CN Unit",
                "Downtime_Minutes",
                "Ready_Minutes",
                "Readiness_Status",
                "Is_Ready",
                "Plan_Ready_Units",
            ]
        )

    start = dates.min().floor("D")
    end = dates.max().floor("D") + pd.Timedelta(hours=23)
    slots = pd.date_range(start, end, freq="h")
    grid = pd.MultiIndex.from_product([slots, units], names=["Slot_Start", "CN Unit"]).to_frame(index=False)
    grid["Downtime_Minutes"] = 0.0

    downtime_lookup = grid.set_index(["Slot_Start", "CN Unit"])["Downtime_Minutes"].to_dict()
    for _, row in events.iterrows():
        unit = str(row["CN Unit"])
        event_start = row["BD_Start"]
        event_end = row["BD_End"]
        if pd.isna(event_start) or pd.isna(event_end) or event_end <= event_start:
            continue

        event_slots = pd.date_range(event_start.floor("h"), (event_end - pd.Timedelta(seconds=1)).floor("h"), freq="h")
        for slot in event_slots:
            slot_end = slot + pd.Timedelta(hours=1)
            overlap_start = max(event_start, slot)
            overlap_end = min(event_end, slot_end)
            overlap_minutes = max(0.0, (overlap_end - overlap_start).total_seconds() / 60)
            key = (slot, unit)
            if key in downtime_lookup:
                downtime_lookup[key] += overlap_minutes

    hourly = pd.DataFrame(
        [(slot, unit, minutes) for (slot, unit), minutes in downtime_lookup.items()],
        columns=["Slot_Start", "CN Unit", "Downtime_Minutes"],
    )
    hourly["Downtime_Minutes"] = hourly["Downtime_Minutes"].clip(upper=60)
    hourly["Ready_Minutes"] = 60 - hourly["Downtime_Minutes"]
    hourly["Is_Ready"] = hourly["Ready_Minutes"] >= ready_threshold_minutes
    hourly["Readiness_Status"] = hourly["Is_Ready"].map({True: "READY", False: "BREAKDOWN"})
    hourly["Date"] = hourly["Slot_Start"].dt.date
    hourly["Hour"] = hourly["Slot_Start"].dt.hour

    plan_cols = plan[["Date", "Plan_Ready_Units"]].copy()
    hourly = hourly.merge(plan_cols, on="Date", how="left")
    hourly["Plan_Ready_Units"] = hourly["Plan_Ready_Units"].fillna(DEFAULT_PLAN_UNITS)
    return hourly


def summarize_hourly_readiness(hourly: pd.DataFrame) -> pd.DataFrame:
    summary = (
        hourly.groupby(["Slot_Start", "Date", "Hour"], dropna=False)
        .agg(
            Actual_Ready_Units=("Is_Ready", "sum"),
            Breakdown_Units=("Is_Ready", lambda s: int((~s).sum())),
            Total_Downtime_Minutes=("Downtime_Minutes", "sum"),
            Plan_Ready_Units=("Plan_Ready_Units", "max"),
        )
        .reset_index()
    )
    summary["Readiness_Gap"] = summary["Actual_Ready_Units"] - summary["Plan_Ready_Units"]
    summary["Ready_Achievement_%"] = (
        summary["Actual_Ready_Units"] / summary["Plan_Ready_Units"].replace(0, pd.NA) * 100
    ).fillna(0)
    return summary


def summarize_daily_readiness(hourly_summary: pd.DataFrame) -> pd.DataFrame:
    daily = (
        hourly_summary.groupby("Date", dropna=False)
        .agg(
            Actual_Ready_Units=("Actual_Ready_Units", "mean"),
            Min_Ready_Units=("Actual_Ready_Units", "min"),
            Breakdown_Unit_Hours=("Breakdown_Units", "sum"),
            Total_Downtime_Minutes=("Total_Downtime_Minutes", "sum"),
            Plan_Ready_Units=("Plan_Ready_Units", "max"),
        )
        .reset_index()
    )
    daily["Readiness_Gap"] = daily["Actual_Ready_Units"] - daily["Plan_Ready_Units"]
    daily["Actual_Ready_Units_Rounded"] = daily["Actual_Ready_Units"].round().astype("Int64")
    daily["Ready_Achievement_%"] = (
        daily["Actual_Ready_Units"] / daily["Plan_Ready_Units"].replace(0, pd.NA) * 100
    ).fillna(0)
    return daily
