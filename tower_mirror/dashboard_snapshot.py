from __future__ import annotations

import json
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from .database import connect, initialize_database


PROCESS_NAMES = {"10": "사출", "20": "분리", "45": "하드레이션", "55": "접착", "80": "누수규격"}
GRADE_ORDER = {"위험": 0, "경고": 1, "주의": 2, "양호": 3}
CLASSIFICATION_CODE_NAMES = {
    "01010101": "1-Day_Sph",
    "01010201": "1-Day_Color_Sph",
    "01010202": "1-Day_Color_Toric",
    "01010207": "1-Day_Color_Sph_Fix",
    "01020101": "FRP_Sph",
    "01020102": "FRP_Toric",
    "01020103": "FRP_M/F",
    "01020201": "FRP_Color_Sph",
    "02010101": "Si_1-Day_Sph",
    "02010102": "Si_1-Day_Toric",
    "02010103": "Si_1-Day_M/F",
    "02010201": "Si_1-Day_Color_Sph",
    "02010202": "Si_1-Day_Color_Toric",
    "02020101": "Si_FRP_Sph",
    "02020102": "Si_FRP_Toric",
    "02020103": "Si_FRP_M/F",
    "02020201": "Si_FRP_Color_Sph",
}


def analysis_lookback_days(db_path: str | Path) -> int:
    initialize_database(db_path)
    with connect(db_path) as con:
        row = con.execute(
            "SELECT setting_value FROM app_setting WHERE setting_key='analysis_lookback_days'"
        ).fetchone()
    try:
        return max(1, min(31, int(row[0]) if row else 10))
    except (TypeError, ValueError):
        return 10


def _rate(numerator: float, denominator: float) -> float:
    return numerator / denominator * 100.0 if denominator > 0 else 0.0


def _grade(score: float, profile: dict[str, Any]) -> str:
    if score >= float(profile["grade_good"]):
        return "양호"
    if score >= float(profile["grade_caution"]):
        return "주의"
    if score >= float(profile["grade_warning"]):
        return "경고"
    return "위험"


def _factory_grade(grades: list[str]) -> str:
    if not grades:
        return "위험"
    counts = Counter(grades)
    most = max(counts.values())
    selected = min((grade for grade, count in counts.items() if count == most), key=lambda grade: GRADE_ORDER[grade])
    if "위험" in grades and GRADE_ORDER[selected] > GRADE_ORDER["경고"]:
        return "경고"
    return selected


def _score(spec: float, exact: float, over: float, nonstandard: float, profile: dict[str, Any]) -> float:
    raw = (
        spec * float(profile["weight_spec"])
        + exact * float(profile["weight_exact"])
        + (100 - min(max(over, 0), 100)) * float(profile["weight_over"])
        + (100 - min(max(nonstandard, 0), 100)) * float(profile["weight_nonstandard"])
    )
    if spec >= float(profile["cap_spec_high"]):
        cap = float(profile["cap_score_high"])
    elif spec >= float(profile["cap_spec_mid"]):
        cap = float(profile["cap_score_mid"])
    elif spec >= float(profile["cap_spec_low"]):
        cap = float(profile["cap_score_low"])
    else:
        cap = float(profile["cap_score_floor"])
    return max(0.0, min(100.0, raw, cap))


def _classification_name_map(rows: list[Any]) -> dict[str, str]:
    """Map ERP classification codes to APS summary labels found on matched SKUs."""
    candidates: dict[str, Counter[str]] = {}
    for row in rows:
        production = str(row["production_classification"] or "").strip()
        aps = str(row["aps_classification"] or "").strip()
        if production and aps:
            candidates.setdefault(production, Counter())[aps] += 1
    inferred = {
        production: counts.most_common(1)[0][0]
        for production, counts in candidates.items()
        if counts
    }
    return {**CLASSIFICATION_CODE_NAMES, **inferred}


def _classification_label(row: Any, name_map: dict[str, str]) -> str:
    production = str(row["production_classification"] or "").strip()
    aps = str(row["aps_classification"] or "").strip()
    return name_map.get(production) or aps or production or "미분류"


def _classification_summaries(
    rows: list[dict[str, Any]], profile: dict[str, Any]
) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for row in rows:
        grouped.setdefault(
            (str(row["factory"]), str(row["process_code"]), str(row["classification"])),
            [],
        ).append(row)

    summaries: list[dict[str, Any]] = []
    for (factory, process_code, classification), items in sorted(grouped.items()):
        actual = sum(float(item["actual_qty"]) for item in items)
        need = sum(float(item["need_qty"]) for item in items)
        effective = sum(float(item["effective_qty"]) for item in items)
        excess = sum(float(item["excess_qty"]) for item in items)
        nonstandard = sum(float(item["nonstandard_qty"]) for item in items)
        produced = sum(int(item["production_sku_count"]) for item in items)
        responded = sum(int(item["responded_sku_count"]) for item in items)
        spec = _rate(responded, produced) if produced > 0 else None
        exact = _rate(effective, actual)
        over_rate = _rate(excess, actual)
        nonstandard_rate = _rate(nonstandard, actual)
        score = _score(spec, exact, over_rate, nonstandard_rate, profile) if spec is not None else None
        summaries.append({
            "factory": factory,
            "process_code": process_code,
            "process": PROCESS_NAMES.get(process_code, process_code),
            "classification": classification,
            "actual_qty": actual,
            "need_qty": need,
            "effective_qty": effective,
            "excess_qty": excess,
            "nonstandard_qty": nonstandard,
            "production_sku_count": produced,
            "responded_sku_count": responded,
            "spec_rate": spec,
            "exact_rate": exact,
            "excess_rate": over_rate,
            "nonstandard_rate": nonstandard_rate,
            "score": score,
            "grade": _grade(score, profile) if score is not None else "집계없음",
        })
    return summaries


def _classification_daily_details(
    rows: list[dict[str, Any]],
    process_rows: list[dict[str, Any]],
    profile: dict[str, Any],
) -> list[dict[str, Any]]:
    """Return dated classification rows with a usable score for every active row.

    Some ERP classification rows have quantities but no classification-level SKU
    counts. In that case the enclosing date/factory/process specification rate is
    the most specific available fallback; the other rates remain classification
    specific.
    """
    process_lookup = {
        (
            str(row["analysis_date"]),
            str(row["factory"]),
            str(row["process_code"]),
        ): row
        for row in process_rows
    }
    details: list[dict[str, Any]] = []
    for row in rows:
        day = str(row["analysis_date"])
        factory = str(row["factory"])
        process_code = str(row["process_code"])
        actual = float(row["actual_qty"])
        need = float(row["need_qty"])
        effective = float(row["effective_qty"])
        excess = float(row["excess_qty"])
        nonstandard = float(row["nonstandard_qty"])
        produced = int(row["production_sku_count"])
        responded = int(row["responded_sku_count"])
        process_row = process_lookup.get((day, factory, process_code))
        if produced > 0:
            spec = _rate(responded, produced)
            spec_source = "신규분류"
        elif process_row is not None:
            spec = float(process_row["spec_rate"])
            spec_source = "공정 보완"
        else:
            spec = 0.0
            spec_source = "미집계"
        exact = _rate(effective, actual)
        over_rate = _rate(excess, actual)
        nonstandard_rate = _rate(nonstandard, actual)
        has_activity = actual > 0 or need > 0
        score = _score(spec, exact, over_rate, nonstandard_rate, profile) if has_activity else None
        details.append({
            "analysis_date": day,
            "factory": factory,
            "process_code": process_code,
            "process": PROCESS_NAMES.get(process_code, process_code),
            "classification": str(row["classification"]),
            "actual_qty": actual,
            "need_qty": need,
            "effective_qty": effective,
            "excess_qty": excess,
            "nonstandard_qty": nonstandard,
            "production_sku_count": produced,
            "responded_sku_count": responded,
            "spec_rate": spec,
            "spec_source": spec_source,
            "exact_rate": exact,
            "excess_rate": over_rate,
            "nonstandard_rate": nonstandard_rate,
            "score": score,
            "grade": _grade(score, profile) if score is not None else "집계없음",
        })
    return details


def rebuild_dashboard_snapshot(
    db_path: str | Path,
    *,
    date_from: str | None = None,
    date_to: str | None = None,
    analysis_run_id: int | None = None,
) -> dict[str, Any]:
    initialize_database(db_path)
    lookback = analysis_lookback_days(db_path)
    yesterday = date.today() - timedelta(days=1)
    requested_end = date.fromisoformat(date_to) if date_to else yesterday
    end = min(requested_end, yesterday)
    start = date.fromisoformat(date_from) if date_from else end - timedelta(days=lookback - 1)
    date_from = start.isoformat(); date_to = end.isoformat()
    generated_at = datetime.now().isoformat(timespec="seconds")

    with connect(db_path) as con:
        profile_row = con.execute("SELECT * FROM score_profile WHERE is_active=1 ORDER BY profile_id DESC LIMIT 1").fetchone()
        if not profile_row:
            raise RuntimeError("활성 공정 밸런스 기준이 없습니다.")
        profile = dict(profile_row)
        profile_id = int(profile["profile_id"])
        source_rows = con.execute(
            "SELECT r.*,d.aps_snapshot_id AS day_snapshot_id FROM analysis_result r "
            "JOIN analysis_day d ON d.analysis_date=r.analysis_date "
            "WHERE r.analysis_date BETWEEN ? AND ?",
            (date_from, date_to),
        ).fetchall()
        con.execute("BEGIN IMMEDIATE")
        source_days = sorted({str(row["analysis_date"]) for row in source_rows})
        if source_days:
            placeholders = ",".join("?" for _ in source_days)
            con.execute(
                f"DELETE FROM analysis_factory_daily WHERE analysis_date IN ({placeholders})",
                source_days,
            )
            con.execute(
                f"DELETE FROM analysis_process_daily WHERE analysis_date IN ({placeholders})",
                source_days,
            )
            con.execute(
                f"DELETE FROM analysis_classification_daily WHERE analysis_date IN ({placeholders})",
                source_days,
            )

        factory_groups: dict[tuple[str, str], list[Any]] = {}
        process_groups: dict[tuple[str, str, str], list[Any]] = {}
        class_groups: dict[tuple[str, str, str, str], list[Any]] = {}
        classification_map = _classification_name_map(source_rows)
        for row in source_rows:
            day = str(row["analysis_date"]); factory = str(row["factory"]); process = str(row["process_code"])
            process_groups.setdefault((day, factory, process), []).append(row)
            classification = _classification_label(row, classification_map)
            class_groups.setdefault((day, factory, process, classification), []).append(row)
            if process == "80":
                factory_groups.setdefault((day, factory), []).append(row)

        for (day, factory), rows in factory_groups.items():
            actual = sum(float(row["actual_qty"] or 0) for row in rows)
            effective = sum(float(row["effective_qty"] or 0) for row in rows)
            excess = sum(float(row["excess_qty"] or 0) for row in rows)
            nonstandard = sum(float(row["unnecessary_qty"] or 0) for row in rows)
            produced = sum(1 for row in rows if float(row["actual_qty"] or 0) > 0)
            responded = sum(1 for row in rows if float(row["actual_qty"] or 0) > 0 and float(row["need_qty"] or 0) > 0)
            con.execute(
                "INSERT INTO analysis_factory_daily VALUES(?,?,?,?,?,?,?,?,?)",
                (day, factory, int(rows[0]["day_snapshot_id"]), produced, responded, actual, effective, excess, nonstandard),
            )

        for (day, factory, process), rows in process_groups.items():
            actual = sum(float(row["actual_qty"] or 0) for row in rows)
            need = sum(float(row["need_qty"] or 0) for row in rows)
            effective = sum(float(row["effective_qty"] or 0) for row in rows)
            over = max(actual - effective, 0)
            nonstandard = sum(float(row["unnecessary_qty"] or 0) for row in rows)
            produced = sum(1 for row in rows if float(row["actual_qty"] or 0) > 0)
            responded = sum(1 for row in rows if float(row["actual_qty"] or 0) > 0 and float(row["need_qty"] or 0) > 0)
            spec = _rate(responded, produced); exact = _rate(effective, actual)
            over_rate = _rate(over, actual); nonstandard_rate = _rate(nonstandard, actual)
            score = _score(spec, exact, over_rate, nonstandard_rate, profile)
            con.execute(
                "INSERT INTO analysis_process_daily VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (day, factory, process, int(rows[0]["day_snapshot_id"]), produced, responded, actual, need,
                 effective, over, nonstandard, spec, exact, over_rate, nonstandard_rate, score,
                 _grade(score, profile), profile_id),
            )

        for (day, factory, process, classification), rows in class_groups.items():
            actual = sum(float(row["actual_qty"] or 0) for row in rows)
            need = sum(float(row["need_qty"] or 0) for row in rows)
            effective = sum(float(row["effective_qty"] or 0) for row in rows)
            produced = sum(1 for row in rows if float(row["actual_qty"] or 0) > 0)
            responded = sum(
                1 for row in rows
                if float(row["actual_qty"] or 0) > 0 and float(row["need_qty"] or 0) > 0
            )
            con.execute(
                "INSERT INTO analysis_classification_daily("
                "analysis_date,factory,process_code,classification,actual_qty,need_qty,effective_qty,"
                "shortage_qty,excess_qty,nonstandard_qty,production_sku_count,responded_sku_count"
                ") VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                (day, factory, process, classification, actual, need, effective, max(need - actual, 0),
                 sum(float(row["excess_qty"] or 0) for row in rows),
                 sum(float(row["unnecessary_qty"] or 0) for row in rows), produced, responded),
            )

        final_rows = [dict(row) for row in con.execute(
            "SELECT * FROM analysis_factory_daily WHERE analysis_date BETWEEN ? AND ? ORDER BY analysis_date,factory",
            (date_from, date_to),
        ).fetchall()]
        process_rows = [dict(row) for row in con.execute(
            "SELECT * FROM analysis_process_daily WHERE analysis_date BETWEEN ? AND ? ORDER BY analysis_date,factory,process_code",
            (date_from, date_to),
        ).fetchall()]
        class_rows = [dict(row) for row in con.execute(
            "SELECT * FROM analysis_classification_daily "
            "WHERE analysis_date BETWEEN ? AND ? ORDER BY analysis_date,factory,process_code,classification",
            (date_from, date_to),
        ).fetchall()]
        if not final_rows and not process_rows:
            raise RuntimeError(f"{date_from} ~ {date_to} 분석 결과가 없습니다.")
        stale_days = int(con.execute(
            "SELECT COUNT(*) FROM analysis_day d JOIN aps_snapshot s ON s.business_date=d.analysis_date AND s.is_baseline=1 "
            "WHERE d.analysis_date BETWEEN ? AND ? AND d.aps_snapshot_id<>s.snapshot_id",
            (date_from, date_to),
        ).fetchone()[0])

        def production_summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
            actual = sum(float(row["actual_qty"]) for row in rows)
            effective = sum(float(row["effective_qty"]) for row in rows)
            excess = sum(float(row["excess_qty"]) for row in rows)
            nonstandard = sum(float(row["nonstandard_qty"]) for row in rows)
            produced = sum(int(row["production_sku_count"]) for row in rows)
            responded = sum(int(row["responded_sku_count"]) for row in rows)
            return {"actual_qty": actual, "effective_qty": effective, "excess_qty": excess,
                    "nonstandard_qty": nonstandard, "spec_rate": _rate(responded, produced),
                    "exact_rate": _rate(effective, actual), "excess_rate": _rate(excess, actual),
                    "nonstandard_rate": _rate(nonstandard, actual)}

        final_process_rows = [
            row for row in process_rows if str(row["process_code"]) == "80"
        ]

        def final_need(*, day: str | None = None, factory: str | None = None) -> float:
            return sum(
                float(row["need_qty"])
                for row in final_process_rows
                if (day is None or str(row["analysis_date"]) == day)
                and (factory is None or str(row["factory"]) == factory)
            )

        production = production_summary(final_rows)
        production["need_qty"] = final_need()
        production["factories"] = []
        for factory in sorted({str(row["factory"]) for row in final_rows}):
            item = production_summary([row for row in final_rows if row["factory"] == factory])
            item["factory"] = factory
            item["need_qty"] = final_need(factory=factory)
            production["factories"].append(item)
        production["daily"] = []
        for day in sorted({str(row["analysis_date"]) for row in final_rows}, reverse=True):
            item = production_summary([row for row in final_rows if row["analysis_date"] == day])
            item["analysis_date"] = day
            item["need_qty"] = final_need(day=day)
            production["daily"].append(item)
        production["factory_daily"] = []
        for day in sorted({str(row["analysis_date"]) for row in final_rows}, reverse=True):
            for factory in sorted({str(row["factory"]) for row in final_rows if str(row["analysis_date"]) == day}):
                item = production_summary([
                    row for row in final_rows
                    if str(row["analysis_date"]) == day and str(row["factory"]) == factory
                ])
                item["analysis_date"] = day
                item["factory"] = factory
                item["need_qty"] = final_need(day=day, factory=factory)
                production["factory_daily"].append(item)
        classification_summaries = _classification_summaries(class_rows, profile)
        classification_daily_details = _classification_daily_details(
            class_rows, process_rows, profile
        )
        production["classifications"] = [
            row for row in classification_summaries if row["process_code"] == "80"
        ]

        def weighted(rows: list[dict[str, Any]]) -> float:
            weight = sum(float(row["actual_qty"]) for row in rows)
            return sum(float(row["process_score"]) * float(row["actual_qty"]) for row in rows) / weight if weight else 0.0

        overall = weighted(process_rows)
        balance: dict[str, Any] = {"overall_score": overall, "overall_grade": _grade(overall, profile), "processes": [], "factories": []}
        for process in PROCESS_NAMES:
            rows = [row for row in process_rows if row["process_code"] == process]
            score = weighted(rows)
            balance["processes"].append({"process_code": process, "process": PROCESS_NAMES[process], "score": score, "grade": _grade(score, profile)})
        for factory in sorted({str(row["factory"]) for row in process_rows}):
            factory_rows = [row for row in process_rows if row["factory"] == factory]
            process_items = []
            for process in PROCESS_NAMES:
                rows = [row for row in factory_rows if row["process_code"] == process]
                score = weighted(rows); process_items.append({"process_code": process, "process": PROCESS_NAMES[process], "score": score, "grade": _grade(score, profile)})
            balance["factories"].append({"factory": factory, "score": weighted(factory_rows), "grade": _factory_grade([item["grade"] for item in process_items]), "processes": process_items})
        balance["daily"] = process_rows
        balance["classifications"] = classification_daily_details

        con.execute("UPDATE dashboard_snapshot SET is_current=0 WHERE is_current=1")
        cursor = con.execute(
            "INSERT INTO dashboard_snapshot(analysis_run_id,date_from,date_to,lookback_days,score_profile_id,generated_at,"
            "stale_days,production_payload,balance_payload,is_current) VALUES(?,?,?,?,?,?,?,?,?,1)",
            (analysis_run_id, date_from, date_to, lookback, profile_id, generated_at, stale_days,
             json.dumps(production, ensure_ascii=False, separators=(",", ":")),
             json.dumps(balance, ensure_ascii=False, separators=(",", ":"))),
        )
        snapshot_id = int(cursor.lastrowid)
        con.execute("DELETE FROM dashboard_snapshot WHERE dashboard_snapshot_id NOT IN (SELECT dashboard_snapshot_id FROM dashboard_snapshot ORDER BY dashboard_snapshot_id DESC LIMIT 10)")
    return {"dashboard_snapshot_id": snapshot_id, "date_from": date_from, "date_to": date_to,
            "generated_at": generated_at, "stale_days": stale_days, "source_rows": len(source_rows)}


def current_dashboard_snapshot(db_path: str | Path) -> dict[str, Any] | None:
    initialize_database(db_path)
    with connect(db_path) as con:
        row = con.execute("SELECT * FROM dashboard_snapshot WHERE is_current=1").fetchone()
    if not row:
        return None
    result = dict(row)
    result["production"] = json.loads(str(result.pop("production_payload")))
    result["balance"] = json.loads(str(result.pop("balance_payload")))
    return result


def available_analysis_range(db_path: str | Path) -> tuple[str | None, str | None]:
    initialize_database(db_path)
    yesterday = (date.today() - timedelta(days=1)).isoformat()
    with connect(db_path) as con:
        row = con.execute(
            "SELECT MIN(analysis_date),MAX(analysis_date) "
            "FROM analysis_factory_daily WHERE analysis_date<=?",
            (yesterday,),
        ).fetchone()
    return (str(row[0]) if row and row[0] else None, str(row[1]) if row and row[1] else None)


def available_analysis_periods(db_path: str | Path) -> list[tuple[int, int]]:
    """Return only year/month periods that actually contain analysis data.

    The newest period is first.  UI filters use this instead of manufacturing
    every calendar month between the minimum and maximum dates, so a newly
    collected month appears automatically while empty months never do.
    """
    initialize_database(db_path)
    yesterday = (date.today() - timedelta(days=1)).isoformat()
    with connect(db_path) as con:
        rows = con.execute(
            "SELECT DISTINCT substr(analysis_date,1,7) AS period "
            "FROM analysis_factory_daily "
            "WHERE analysis_date<=? AND length(analysis_date)>=7 "
            "ORDER BY period DESC",
            (yesterday,),
        ).fetchall()
    periods: list[tuple[int, int]] = []
    for row in rows:
        try:
            year_text, month_text = str(row[0]).split("-", 1)
            year, month = int(year_text), int(month_text)
        except (TypeError, ValueError):
            continue
        if 1 <= month <= 12:
            periods.append((year, month))
    return periods


def monthly_response_payload(
    db_path: str | Path,
    *,
    selected_month: str | None = None,
    month_count: int = 6,
) -> dict[str, Any] | None:
    """Return weighted monthly response rates and same-day prior-month context.

    Monthly specification response is calculated from summed SKU counts, never
    from an unweighted average of daily percentages.
    """
    initialize_database(db_path)
    yesterday = date.today() - timedelta(days=1)
    with connect(db_path) as con:
        latest_row = con.execute(
            "SELECT MAX(analysis_date) FROM analysis_factory_daily WHERE analysis_date<=?",
            (yesterday.isoformat(),),
        ).fetchone()
        earliest_row = con.execute(
            "SELECT MIN(analysis_date) FROM analysis_factory_daily WHERE analysis_date<=?",
            (yesterday.isoformat(),),
        ).fetchone()
        latest_text = str(latest_row[0]) if latest_row and latest_row[0] else None
        earliest_text = str(earliest_row[0]) if earliest_row and earliest_row[0] else None
        if not latest_text or not earliest_text:
            return None

        latest_day = date.fromisoformat(latest_text)
        selected_key = selected_month or latest_day.strftime("%Y-%m")
        try:
            selected_start = date.fromisoformat(f"{selected_key}-01")
        except ValueError:
            selected_start = latest_day.replace(day=1)
            selected_key = selected_start.strftime("%Y-%m")

        def shift_month(value: date, offset: int) -> date:
            absolute = value.year * 12 + value.month - 1 + offset
            return date(absolute // 12, absolute % 12 + 1, 1)

        selected_next = shift_month(selected_start, 1)
        selected_calendar_end = selected_next - timedelta(days=1)
        selected_end = min(selected_calendar_end, latest_day, yesterday)
        if selected_end < selected_start:
            selected_start = latest_day.replace(day=1)
            selected_key = selected_start.strftime("%Y-%m")
            selected_next = shift_month(selected_start, 1)
            selected_calendar_end = selected_next - timedelta(days=1)
            selected_end = min(selected_calendar_end, latest_day, yesterday)

        trend_start = shift_month(selected_start, -(max(2, month_count) - 1))
        previous_start = shift_month(selected_start, -1)
        previous_calendar_end = selected_start - timedelta(days=1)
        previous_end = min(
            previous_calendar_end,
            previous_start + timedelta(days=selected_end.day - 1),
        )
        query_start = min(trend_start, previous_start)
        rows = [
            dict(row)
            for row in con.execute(
                "SELECT analysis_date,factory,production_sku_count,responded_sku_count,"
                "actual_qty,effective_qty,excess_qty,nonstandard_qty "
                "FROM analysis_factory_daily WHERE analysis_date BETWEEN ? AND ? "
                "ORDER BY analysis_date,factory",
                (query_start.isoformat(), selected_end.isoformat()),
            ).fetchall()
        ]

    def summarize(items: list[dict[str, Any]]) -> dict[str, Any]:
        produced = sum(int(row["production_sku_count"] or 0) for row in items)
        responded = sum(int(row["responded_sku_count"] or 0) for row in items)
        actual = sum(float(row["actual_qty"] or 0) for row in items)
        effective = sum(float(row["effective_qty"] or 0) for row in items)
        excess = sum(float(row["excess_qty"] or 0) for row in items)
        nonstandard = sum(float(row["nonstandard_qty"] or 0) for row in items)
        return {
            "production_sku_count": produced,
            "responded_sku_count": responded,
            "actual_qty": actual,
            "effective_qty": effective,
            "excess_qty": excess,
            "nonstandard_qty": nonstandard,
            "spec_rate": _rate(responded, produced),
            "exact_rate": _rate(effective, actual),
            "excess_rate": _rate(excess, actual),
            "nonstandard_rate": _rate(nonstandard, actual),
        }

    factories = sorted({str(row["factory"]) for row in rows})
    months: list[dict[str, Any]] = []
    cursor = trend_start
    while cursor <= selected_start:
        key = cursor.strftime("%Y-%m")
        month_rows = [row for row in rows if str(row["analysis_date"]).startswith(key)]
        if month_rows:
            factory_values = {
                factory: summarize([row for row in month_rows if str(row["factory"]) == factory])
                for factory in factories
            }
            months.append({
                "key": key,
                "label": f"{cursor.month}월",
                "is_selected": key == selected_key,
                "is_mtd": key == selected_key and selected_end < selected_calendar_end,
                "overall": summarize(month_rows),
                "factories": factory_values,
            })
        cursor = shift_month(cursor, 1)

    selected_rows = [
        row for row in rows
        if selected_start.isoformat() <= str(row["analysis_date"]) <= selected_end.isoformat()
    ]
    previous_rows = [
        row for row in rows
        if previous_start.isoformat() <= str(row["analysis_date"]) <= previous_end.isoformat()
    ]
    daily: list[dict[str, Any]] = []
    for day_text in sorted({str(row["analysis_date"]) for row in selected_rows}):
        day_rows = [row for row in selected_rows if str(row["analysis_date"]) == day_text]
        daily.append({
            "analysis_date": day_text,
            "overall": summarize(day_rows),
            "factories": {
                factory: summarize([
                    row for row in day_rows if str(row["factory"]) == factory
                ])
                for factory in factories
            },
        })
    return {
        "earliest_date": earliest_text,
        "latest_date": latest_text,
        "selected_month": selected_key,
        "selected_date_from": selected_start.isoformat(),
        "selected_date_to": selected_end.isoformat(),
        "previous_date_from": previous_start.isoformat(),
        "previous_date_to": previous_end.isoformat(),
        "is_mtd": selected_end < selected_calendar_end,
        "months": months,
        "daily": daily,
        "factories": factories,
        "selected": {
            "overall": summarize(selected_rows),
            "factories": {
                factory: summarize([row for row in selected_rows if str(row["factory"]) == factory])
                for factory in factories
            },
        },
        "previous": {
            "overall": summarize(previous_rows),
            "factories": {
                factory: summarize([row for row in previous_rows if str(row["factory"]) == factory])
                for factory in factories
            },
        },
    }


def response_payload_for_period(
    db_path: str | Path,
    date_from: str,
    date_to: str,
) -> dict[str, Any] | None:
    """Return weighted response metrics for an arbitrary dashboard period."""
    initialize_database(db_path)
    yesterday = date.today() - timedelta(days=1)
    with connect(db_path) as con:
        range_row = con.execute(
            "SELECT MIN(analysis_date),MAX(analysis_date) "
            "FROM analysis_factory_daily WHERE analysis_date<=?",
            (yesterday.isoformat(),),
        ).fetchone()
        if not range_row or not range_row[0] or not range_row[1]:
            return None
        earliest = date.fromisoformat(str(range_row[0]))
        latest = min(date.fromisoformat(str(range_row[1])), yesterday)

        requested_start = date.fromisoformat(date_from)
        requested_end = date.fromisoformat(date_to)
        selected_start = max(min(requested_start, requested_end), earliest)
        selected_end = min(max(requested_start, requested_end), latest)
        if selected_start > selected_end:
            return None

        def shift_same_day(value: date, offset: int) -> date:
            absolute = value.year * 12 + value.month - 1 + offset
            shifted_start = date(absolute // 12, absolute % 12 + 1, 1)
            next_absolute = absolute + 1
            shifted_next = date(
                next_absolute // 12,
                next_absolute % 12 + 1,
                1,
            )
            shifted_last_day = (shifted_next - timedelta(days=1)).day
            return shifted_start.replace(day=min(value.day, shifted_last_day))

        if (
            selected_start.year == selected_end.year
            and selected_start.month == selected_end.month
        ):
            previous_start = shift_same_day(selected_start, -1)
            previous_end = shift_same_day(selected_end, -1)
            comparison_label = "전월 동기간"
        else:
            period_days = (selected_end - selected_start).days + 1
            previous_end = selected_start - timedelta(days=1)
            previous_start = previous_end - timedelta(days=period_days - 1)
            comparison_label = "직전 동기간"

        rows = [
            dict(row)
            for row in con.execute(
                "SELECT analysis_date,factory,production_sku_count,responded_sku_count,"
                "actual_qty,effective_qty,excess_qty,nonstandard_qty "
                "FROM analysis_factory_daily WHERE analysis_date BETWEEN ? AND ? "
                "ORDER BY analysis_date,factory",
                (
                    max(previous_start, earliest).isoformat(),
                    selected_end.isoformat(),
                ),
            ).fetchall()
        ]

    def summarize(items: list[dict[str, Any]]) -> dict[str, Any]:
        produced = sum(int(row["production_sku_count"] or 0) for row in items)
        responded = sum(int(row["responded_sku_count"] or 0) for row in items)
        actual = sum(float(row["actual_qty"] or 0) for row in items)
        effective = sum(float(row["effective_qty"] or 0) for row in items)
        excess = sum(float(row["excess_qty"] or 0) for row in items)
        nonstandard = sum(float(row["nonstandard_qty"] or 0) for row in items)
        return {
            "production_sku_count": produced,
            "responded_sku_count": responded,
            "actual_qty": actual,
            "effective_qty": effective,
            "excess_qty": excess,
            "nonstandard_qty": nonstandard,
            "spec_rate": _rate(responded, produced),
            "exact_rate": _rate(effective, actual),
            "excess_rate": _rate(excess, actual),
            "nonstandard_rate": _rate(nonstandard, actual),
        }

    factories = sorted({str(row["factory"]) for row in rows})
    selected_rows = [
        row for row in rows
        if selected_start.isoformat() <= str(row["analysis_date"]) <= selected_end.isoformat()
    ]
    previous_rows = [
        row for row in rows
        if previous_start.isoformat() <= str(row["analysis_date"]) <= previous_end.isoformat()
    ]
    daily: list[dict[str, Any]] = []
    for day_text in sorted({str(row["analysis_date"]) for row in selected_rows}):
        day_rows = [row for row in selected_rows if str(row["analysis_date"]) == day_text]
        daily.append({
            "analysis_date": day_text,
            "overall": summarize(day_rows),
            "factories": {
                factory: summarize([
                    row for row in day_rows if str(row["factory"]) == factory
                ])
                for factory in factories
            },
        })
    return {
        "earliest_date": earliest.isoformat(),
        "latest_date": latest.isoformat(),
        "selected_date_from": selected_start.isoformat(),
        "selected_date_to": selected_end.isoformat(),
        "previous_date_from": previous_start.isoformat(),
        "previous_date_to": previous_end.isoformat(),
        "comparison_label": comparison_label,
        "daily": daily,
        "factories": factories,
        "selected": {
            "overall": summarize(selected_rows),
            "factories": {
                factory: summarize([
                    row for row in selected_rows if str(row["factory"]) == factory
                ])
                for factory in factories
            },
        },
        "previous": {
            "overall": summarize(previous_rows),
            "factories": {
                factory: summarize([
                    row for row in previous_rows if str(row["factory"]) == factory
                ])
                for factory in factories
            },
        },
    }


def dashboard_payload_for_range(db_path: str | Path, date_from: str, date_to: str) -> dict[str, Any] | None:
    """Build a lightweight view from daily aggregates, never from item-level matching rows."""
    initialize_database(db_path)
    yesterday = (date.today() - timedelta(days=1)).isoformat()
    date_to = min(date_to, yesterday)
    if date_from > date_to:
        return None
    with connect(db_path) as con:
        profile_row = con.execute("SELECT * FROM score_profile WHERE is_active=1 ORDER BY profile_id DESC LIMIT 1").fetchone()
        if not profile_row:
            return None
        profile = dict(profile_row)
        final_rows = [dict(row) for row in con.execute(
            "SELECT * FROM analysis_factory_daily WHERE analysis_date BETWEEN ? AND ? ORDER BY analysis_date,factory",
            (date_from, date_to),
        ).fetchall()]
        process_rows = [dict(row) for row in con.execute(
            "SELECT * FROM analysis_process_daily WHERE analysis_date BETWEEN ? AND ? ORDER BY analysis_date,factory,process_code",
            (date_from, date_to),
        ).fetchall()]
        class_rows = [dict(row) for row in con.execute(
            "SELECT * FROM analysis_classification_daily "
            "WHERE analysis_date BETWEEN ? AND ? ORDER BY analysis_date,factory,process_code,classification",
            (date_from, date_to),
        ).fetchall()]
        if not final_rows and not process_rows:
            return None
        stale_days = int(con.execute(
            "SELECT COUNT(*) FROM analysis_day d JOIN aps_snapshot s ON s.business_date=d.analysis_date AND s.is_baseline=1 "
            "WHERE d.analysis_date BETWEEN ? AND ? AND d.aps_snapshot_id<>s.snapshot_id",
            (date_from, date_to),
        ).fetchone()[0])
        current = con.execute("SELECT generated_at FROM dashboard_snapshot WHERE is_current=1").fetchone()

    for row in process_rows:
        score = _score(float(row["spec_rate"]), float(row["exact_rate"]), float(row["over_rate"]), float(row["nonstandard_rate"]), profile)
        row["process_score"] = score; row["grade"] = _grade(score, profile)

    def production_summary(rows: list[dict[str, Any]]) -> dict[str, Any]:
        actual = sum(float(row["actual_qty"]) for row in rows)
        effective = sum(float(row["effective_qty"]) for row in rows)
        excess = sum(float(row["excess_qty"]) for row in rows)
        nonstandard = sum(float(row["nonstandard_qty"]) for row in rows)
        produced = sum(int(row["production_sku_count"]) for row in rows)
        responded = sum(int(row["responded_sku_count"]) for row in rows)
        return {"actual_qty": actual, "effective_qty": effective, "excess_qty": excess,
                "nonstandard_qty": nonstandard, "spec_rate": _rate(responded, produced),
                "exact_rate": _rate(effective, actual), "excess_rate": _rate(excess, actual),
                "nonstandard_rate": _rate(nonstandard, actual)}

    final_process_rows = [
        row for row in process_rows if str(row["process_code"]) == "80"
    ]

    def final_need(*, day: str | None = None, factory: str | None = None) -> float:
        return sum(
            float(row["need_qty"])
            for row in final_process_rows
            if (day is None or str(row["analysis_date"]) == day)
            and (factory is None or str(row["factory"]) == factory)
        )

    production = production_summary(final_rows)
    production["need_qty"] = final_need()
    production["factories"] = []
    for factory in sorted({str(row["factory"]) for row in final_rows}):
        item = production_summary([row for row in final_rows if row["factory"] == factory])
        item["factory"] = factory
        item["need_qty"] = final_need(factory=factory)
        production["factories"].append(item)
    production["daily"] = []
    for day in sorted({str(row["analysis_date"]) for row in final_rows}, reverse=True):
        item = production_summary([row for row in final_rows if row["analysis_date"] == day])
        item["analysis_date"] = day
        item["need_qty"] = final_need(day=day)
        production["daily"].append(item)
    production["factory_daily"] = []
    for day in sorted({str(row["analysis_date"]) for row in final_rows}, reverse=True):
        for factory in sorted({str(row["factory"]) for row in final_rows if str(row["analysis_date"]) == day}):
            item = production_summary([
                row for row in final_rows
                if str(row["analysis_date"]) == day and str(row["factory"]) == factory
            ])
            item["analysis_date"] = day
            item["factory"] = factory
            item["need_qty"] = final_need(day=day, factory=factory)
            production["factory_daily"].append(item)
    classification_summaries = _classification_summaries(class_rows, profile)
    classification_daily_details = _classification_daily_details(
        class_rows, process_rows, profile
    )
    production["classifications"] = [
        row for row in classification_summaries if row["process_code"] == "80"
    ]

    def weighted(rows: list[dict[str, Any]]) -> float:
        weight = sum(float(row["actual_qty"]) for row in rows)
        return sum(float(row["process_score"]) * float(row["actual_qty"]) for row in rows) / weight if weight else 0.0

    overall = weighted(process_rows)
    balance: dict[str, Any] = {"overall_score": overall, "overall_grade": _grade(overall, profile), "processes": [], "factories": [], "daily": process_rows}
    for process in PROCESS_NAMES:
        rows = [row for row in process_rows if row["process_code"] == process]
        score = weighted(rows)
        balance["processes"].append({"process_code": process, "process": PROCESS_NAMES[process], "score": score, "grade": _grade(score, profile)})
    for factory in sorted({str(row["factory"]) for row in process_rows}):
        factory_rows = [row for row in process_rows if row["factory"] == factory]
        process_items = []
        for process in PROCESS_NAMES:
            rows = [row for row in factory_rows if row["process_code"] == process]
            score = weighted(rows); process_items.append({"process_code": process, "process": PROCESS_NAMES[process], "score": score, "grade": _grade(score, profile)})
        balance["factories"].append({"factory": factory, "score": weighted(factory_rows),
                                     "grade": _factory_grade([item["grade"] for item in process_items]), "processes": process_items})
    balance["classifications"] = classification_daily_details
    data_dates = [str(row["analysis_date"]) for row in final_rows] or [str(row["analysis_date"]) for row in process_rows]
    return {"date_from": date_from, "date_to": date_to, "data_date_from": min(data_dates), "data_date_to": max(data_dates),
            "generated_at": str(current[0]) if current else None, "stale_days": stale_days,
            "production": production, "balance": balance}
