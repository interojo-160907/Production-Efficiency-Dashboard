import os
from pathlib import Path
import calendar
from datetime import datetime, date
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import io
import json
from xlsxwriter.utility import xl_rowcol_to_cell
import textwrap


# ====== Dashboard color system (global) ======
# 목표: 생산운영현황/공정밸런스 탭 간 색상 통일(공장/의미색)
FACTORY_COLOR_MAP = {
    # 공장 라벨이 여러 형태로 등장할 수 있어 키를 넓게 잡음
    "A관": "#6366F1",  # indigo
    "A관(1공장)": "#6366F1",
    "C관": "#8B5CF6",  # violet
    "C관(2공장)": "#8B5CF6",
    "S관": "#EC4899",  # pink
    "S관(3공장)": "#EC4899",
}

BALANCE_COLORS = {
    "정확": "#7C3AED",  # purple
    "초과": "#F43F5E",  # rose
    # 비정형은 살짝 톤다운(너무 튀는 주황 방지)
    "비정형": "#FB923C",  # soft orange
    "초과+비정형": "#F43F5E",
}

# ====== Score grading (공정/공장 공통) ======
# 공정 등급 기준(점수 0~100)
# - 양호: 70 이상
# - 주의: 65 이상
# - 경고: 60 이상
# - 위험: 60 미만
GRADE_RANK = {"양호": 3, "주의": 2, "경고": 1, "위험": 0}
INV_GRADE_RANK = {v: k for k, v in GRADE_RANK.items()}


def grade_of(score: float) -> str:
    try:
        s = float(score)
    except Exception:
        return "위험"
    if np.isnan(s):
        return "위험"
    if s >= 70:
        return "양호"
    if s >= 65:
        return "주의"
    if s >= 60:
        return "경고"
    return "위험"


def grade_text_color(grade: str) -> str:
    # 직관: 양호/주의=차분(블루/블랙), 경고/위험=레드
    if grade == "주의":
        return "#1D4ED8"
    if grade in {"경고", "위험"}:
        return "#B91C1C"
    return "#111827"


def majority_grade(grades: list[str]) -> str:
    # 공정 5개 등급의 다수결(3개 이상) / 2-2-1이면 낮은 등급 선택
    counts: dict[str, int] = {}
    for g in grades:
        counts[g] = counts.get(g, 0) + 1
    majority = next((g for g, n in counts.items() if n >= 3), None)
    if majority is not None:
        picked = majority
    else:
        max_n = max(counts.values()) if counts else 0
        top_grades = [g for g, n in counts.items() if n == max_n]
        picked = INV_GRADE_RANK[min(GRADE_RANK.get(g, 0) for g in top_grades)] if top_grades else "위험"
    # 안전장치: '위험' 공정이 1개라도 있으면 종합은 최대 '경고'
    if "위험" in grades and picked in {"양호", "주의"}:
        picked = "경고"
    return picked


def _factory_color_discrete_map(factories: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for f in factories:
        key = str(f)
        if "A관" in key:
            out[key] = FACTORY_COLOR_MAP["A관"]
        elif "C관" in key:
            out[key] = FACTORY_COLOR_MAP["C관"]
        elif "S관" in key:
            out[key] = FACTORY_COLOR_MAP["S관"]
    return out


def _safe_sheet_name(name: str) -> str:
    name = str(name).strip().replace("/", "_").replace("\\", "_").replace(":", "_")
    if not name:
        name = "Sheet"
    return name[:31]


def _df_to_sheet(
    writer: pd.ExcelWriter,
    *,
    sheet_name: str,
    df: pd.DataFrame,
    startrow: int,
    startcol: int,
) -> tuple[int, int, int, int]:
    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow, startcol=startcol)
    first_row = startrow
    last_row = startrow + len(df)  # include header row
    first_col = startcol
    last_col = startcol + max(len(df.columns) - 1, 0)
    return first_row, last_row, first_col, last_col


def _apply_table_formats(workbook, worksheet, *, df: pd.DataFrame, startrow: int, startcol: int) -> None:
    # IMPORTANT: Do not use set_column formats here because multiple tables share the same worksheet
    # and column-level formats would overwrite each other. Apply formats to the table cell ranges only.
    fmt_int = workbook.add_format({"num_format": "#,##0"})
    fmt_pct = workbook.add_format({"num_format": "0.0\"%\""})
    fmt_date = workbook.add_format({"num_format": "yyyy-mm-dd"})

    nrows = len(df)
    if nrows <= 0:
        return

    data_first_row = startrow + 1
    data_last_row = startrow + nrows

    for idx, col in enumerate(df.columns):
        c = startcol + idx
        name = str(col)

        # Width is safe to set at column level (style is not).
        if name in {"날짜", "기간"}:
            worksheet.set_column(c, c, 12)
        elif name in {"공정점수", "종합점수", "평균점수"} or ("점수" in name):
            worksheet.set_column(c, c, 12)
        elif "(pcs)" in name or name in {"총실적", "유효생산량", "과생산량", "불필요생산량", "총부족수량", "실적수량", "필요수량", "부족수량"}:
            worksheet.set_column(c, c, 16)
        elif "(%)" in name:
            worksheet.set_column(c, c, 14)
        else:
            worksheet.set_column(c, c, 16 if len(name) <= 10 else 20)

        if name in {"날짜", "기간"}:
            fmt = fmt_date
        elif name in {"공정점수", "종합점수", "평균점수"} or ("점수" in name):
            fmt = workbook.add_format({"num_format": "0.0"})
        elif "(pcs)" in name or name in {"총실적", "유효생산량", "과생산량", "불필요생산량", "총부족수량", "실적수량", "필요수량", "부족수량"}:
            fmt = fmt_int
        elif "(%)" in name or name in {"선택지표"}:
            fmt = fmt_pct
        else:
            fmt = None

        if fmt is None:
            continue

        rng = f"{xl_rowcol_to_cell(data_first_row, c)}:{xl_rowcol_to_cell(data_last_row, c)}"
        # Apply number format visually without overriding column formats of other tables.
        worksheet.conditional_format(rng, {"type": "formula", "criteria": "TRUE", "format": fmt})


@st.cache_data(show_spinner=False)
def build_two_sheet_excel(summary_df: pd.DataFrame, detail_df: pd.DataFrame, *, sheet1: str, sheet2: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        s1 = _safe_sheet_name(sheet1)
        ws1 = workbook.add_worksheet(s1)
        writer.sheets[s1] = ws1
        r1, _, c1, _ = _df_to_sheet(writer, sheet_name=s1, df=summary_df, startrow=0, startcol=0)
        _apply_table_formats(workbook, ws1, df=summary_df, startrow=r1, startcol=c1)

        s2 = _safe_sheet_name(sheet2)
        ws2 = workbook.add_worksheet(s2)
        writer.sheets[s2] = ws2
        r2, _, c2, _ = _df_to_sheet(writer, sheet_name=s2, df=detail_df, startrow=0, startcol=0)
        _apply_table_formats(workbook, ws2, df=detail_df, startrow=r2, startcol=c2)

    output.seek(0)
    return output.getvalue()


@st.cache_data(show_spinner=False)
def build_balance_tables_for_export(proc: pd.DataFrame, det: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """공정밸런스 탭 하단 '공장별 요약/상세 테이블' 다운로드용 데이터(집계 완료)를 생성."""
    target_order = ["사출", "분리", "하드레이션", "접착", "누수규격"]
    factory_order = ["A관", "C관", "S관"]

    summary = (
        proc.groupby(["날짜_date", "공장그룹", "공장", "공정"], dropna=False)[
            [c for c in ["실적수량", "유효생산량", "과생산량", "불필요생산량", "생산SKU수", "필요SKU수", "규격대응SKU수"] if c in proc.columns]
        ]
        .sum()
        .reset_index()
    )
    summary["규격대응률(%)"] = np.where(
        pd.to_numeric(summary.get("생산SKU수", 0), errors="coerce").fillna(0) > 0,
        pd.to_numeric(summary.get("규격대응SKU수", 0), errors="coerce").fillna(0)
        / pd.to_numeric(summary.get("생산SKU수", 0), errors="coerce").fillna(0)
        * 100,
        0.0,
    )
    summary["정확대응비중(%)"] = np.where(
        pd.to_numeric(summary.get("실적수량", 0), errors="coerce").fillna(0) > 0,
        pd.to_numeric(summary.get("유효생산량", 0), errors="coerce").fillna(0)
        / pd.to_numeric(summary.get("실적수량", 0), errors="coerce").fillna(0)
        * 100,
        0.0,
    )
    summary["초과생산비중(%)"] = np.where(
        pd.to_numeric(summary.get("실적수량", 0), errors="coerce").fillna(0) > 0,
        pd.to_numeric(summary.get("과생산량", 0), errors="coerce").fillna(0)
        / pd.to_numeric(summary.get("실적수량", 0), errors="coerce").fillna(0)
        * 100,
        0.0,
    )
    summary["비정형생산비중(%)"] = np.where(
        pd.to_numeric(summary.get("실적수량", 0), errors="coerce").fillna(0) > 0,
        pd.to_numeric(summary.get("불필요생산량", 0), errors="coerce").fillna(0)
        / pd.to_numeric(summary.get("실적수량", 0), errors="coerce").fillna(0)
        * 100,
        0.0,
    )
    for c in ["규격대응률(%)", "정확대응비중(%)", "초과생산비중(%)", "비정형생산비중(%)"]:
        summary[c] = pd.to_numeric(summary[c], errors="coerce").fillna(0)
    summary["규격대응률(%)"] = summary["규격대응률(%)"].clip(0, 100)
    summary["정확대응비중(%)"] = summary["정확대응비중(%)"].clip(0, 100)
    summary["초과생산비중(%)"] = summary["초과생산비중(%)"].clip(0, 300)
    summary["비정형생산비중(%)"] = summary["비정형생산비중(%)"].clip(0, 300)

    summary["공정점수_raw"] = (
        summary["규격대응률(%)"] * 0.45
        + summary["정확대응비중(%)"] * 0.25
        + (100 - summary["초과생산비중(%)"].clip(0, 100)) * 0.10
        + (100 - summary["비정형생산비중(%)"].clip(0, 100)) * 0.20
    ).clip(0, 100)
    cap = np.select(
        [
            summary["규격대응률(%)"] >= 85,
            summary["규격대응률(%)"] >= 70,
            summary["규격대응률(%)"] >= 55,
        ],
        [100.0, 75.0, 65.0],
        default=55.0,
    )
    summary["공정점수"] = np.minimum(summary["공정점수_raw"], cap).clip(0, 100)
    summary["상태"] = np.select(
        [summary["공정점수"] >= 70, summary["공정점수"] >= 65, summary["공정점수"] >= 60],
        ["양호", "주의", "경고"],
        default="위험",
    )
    if "공장그룹" in summary.columns:
        summary["공장그룹"] = pd.Categorical(summary["공장그룹"], categories=factory_order + ["기타"], ordered=True)
    if "공정" in summary.columns:
        summary["공정"] = pd.Categorical(summary["공정"], categories=target_order, ordered=True)
    summary = summary.sort_values(["날짜_date", "공장그룹", "공장", "공정"], ascending=[True, True, True, True])

    summary_show_cols = [
        c
        for c in [
            "날짜_date",
            "공장",
            "공정",
            "실적수량",
            "규격대응률(%)",
            "정확대응비중(%)",
            "초과생산비중(%)",
            "비정형생산비중(%)",
            "공정점수",
            "상태",
        ]
        if c in summary.columns
    ]
    summary_view = summary[summary_show_cols].copy() if summary_show_cols else summary.copy()

    det_show = det.copy()
    group_cols = [c for c in ["날짜_date", "공장그룹", "공장", "공정", "신규분류요약"] if c in det_show.columns]
    value_cols = [c for c in ["실적수량", "필요수량", "부족수량", "유효생산량", "과생산량", "불필요생산량"] if c in det_show.columns]
    if group_cols and value_cols:
        det_show = det_show.groupby(group_cols, dropna=False)[value_cols].sum().reset_index()
    show_cols = [c for c in ["날짜_date", "공장", "공정", "신규분류요약"] if c in det_show.columns] + value_cols
    det_show = det_show[show_cols].copy() if show_cols else det_show
    if "공정" in det_show.columns:
        det_show["공정"] = pd.Categorical(det_show["공정"], categories=target_order, ordered=True)
    if "공장그룹" in det_show.columns:
        det_show["공장그룹"] = pd.Categorical(det_show["공장그룹"], categories=factory_order + ["기타"], ordered=True)
    sort_cols = [c for c in ["날짜_date", "공장그룹", "공장", "공정", "신규분류요약"] if c in det_show.columns]
    det_show = det_show.sort_values(sort_cols, ascending=[True] * len(sort_cols)) if sort_cols else det_show

    return summary_view.reset_index(drop=True), det_show.reset_index(drop=True)


def _write_chart_source_df(
    writer: pd.ExcelWriter,
    data_sheet_name: str,
    *,
    df: pd.DataFrame,
    startrow: int,
    startcol: int,
) -> tuple[int, int, int, int]:
    # Ensure datetime column stays datetime for Excel axis formatting.
    df2 = df.copy()
    if len(df2.columns) > 0 and str(df2.columns[0]) in {"기간", "날짜"}:
        df2[df2.columns[0]] = pd.to_datetime(df2[df2.columns[0]], errors="coerce")
    return _df_to_sheet(writer, sheet_name=data_sheet_name, df=df2, startrow=startrow, startcol=startcol)


def _excel_col_width_to_pixels(width: float) -> int:
    # XlsxWriter uses an Excel-like character width unit. Approximate conversion.
    # See: Excel column width ≈ number of '0' characters. Typical pixel conversion:
    # pixels = trunc(width * 7 + 5) for widths >= 1.
    if width is None:
        width = 8.43
    if width <= 0:
        return 0
    return int(width * 7 + 5)


def _excel_row_height_to_pixels(points: float) -> int:
    # Excel row height is in points. 1 point = 1/72 inch. At 96 dpi => 96/72 = 1.333 px/pt.
    if points is None:
        points = 15.0
    return int(points * 96.0 / 72.0)


def _chart_box_pixels(
    *,
    col_widths: dict[int, float],
    row_height_points: float,
    first_col: int,
    last_col: int,
    first_row: int,
    last_row: int,
    pad_px: int = 8,
) -> tuple[int, int]:
    # rows/cols are 0-based inclusive.
    width_px = 0
    for c in range(first_col, last_col + 1):
        width_px += _excel_col_width_to_pixels(col_widths.get(c))
    height_px = _excel_row_height_to_pixels(row_height_points) * (last_row - first_row + 1)
    width_px = max(200, width_px - pad_px)
    height_px = max(160, height_px - pad_px)
    return width_px, height_px


def _build_excel_report_bytes(
    *,
    metric_order: list[str],
    metric_sheet_map: dict[str, str],
    metric_desc: dict[str, str],
    export_payload: dict[str, dict[str, object]],
    start_date_str: str,
    end_date_str: str,
    tz_name: str,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        fmt_title = workbook.add_format({"bold": True, "font_size": 14})
        fmt_section = workbook.add_format({"bold": True, "font_size": 12})
        fmt_note = workbook.add_format({"font_size": 10, "font_color": "#6b7280"})
        fmt_header_bg = workbook.add_format({"bold": True, "bg_color": "#f3f4f6", "border": 1})

        # Hidden data sheet for chart source ranges (Excel charts must reference cells on a worksheet).
        # Keep it hidden so the report sheets remain clean.
        # Create DATA sheet last (so the workbook opens on the first report sheet).
        # We still need stable cell ranges for charts, so we pre-allocate row blocks
        # and write the actual DATA sheet at the end.
        data_sheet_name = "DATA"
        data_next_row = 0
        pending_data_blocks: list[tuple[pd.DataFrame, int, int]] = []

        chart_font = "Malgun Gothic"
        title_font = {"name": chart_font, "size": 20, "bold": True, "color": "#111827"}
        axis_title_font = {"name": chart_font, "size": 11, "bold": False, "color": "#374151"}
        axis_num_font = {"name": chart_font, "size": 10, "color": "#374151"}
        legend_font = {"name": chart_font, "size": 10, "color": "#374151"}
        # XlsxWriter doesn't reliably support alpha on gridlines across Excel versions,
        # so approximate "50% transparency" with a lighter color.
        gridline_color = "#eef2f7"

        # Chart boxes are computed per-sheet to fit target cell ranges.

        def _ymax_0_100(series_max: float | None) -> int:
            if series_max is None:
                return 100
            try:
                v = float(series_max)
            except Exception:
                return 100
            if not np.isfinite(v):
                return 100
            v = max(0.0, min(100.0, v))
            # round up to nearest 10, but at least 20 for readability
            ymax = int(np.ceil(v / 10.0) * 10.0)
            ymax = max(20, min(100, ymax))
            return ymax

        # Parse dates for chart bucketing (Excel axis)
        try:
            _start_date_obj = pd.to_datetime(start_date_str).date()
            _end_date_obj = pd.to_datetime(end_date_str).date()
        except Exception:
            _start_date_obj = None
            _end_date_obj = None

        for metric in metric_order:
            sheet_name = _safe_sheet_name(metric_sheet_map.get(metric, metric))
            payload = export_payload.get(metric, {})

            factory_table = payload.get("factory_table")
            daily_table = payload.get("daily_table")
            factory_daily_table = payload.get("factory_daily_table")

            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet

            # Keep a deterministic row height so "cell-range sized" charts are consistent.
            row_height_points = 15.0
            worksheet.set_default_row(row_height_points)

            # Layout (0-based): match requested template (chart on top, table starts at fixed rows)
            col0 = 0
            sec1_top = 4          # row 5 (1-based): "선택지표 (공장 비교)"
            sec1_chart_row = 5    # row 6: chart
            sec1_table_row = 23   # row 24: table header row

            sec2_top_min = 31     # row 32: "일별요약"
            sec2_chart_row_min = 33  # row 34: line chart (match 규격대응률 sheet)
            sec2_table_row_min = 50  # row 51: table header row (chart above)

            # Chart sizing: keep height stable (avoid covering tables) and only widen as needed.
            bar_chart_scale = {"x_scale": 1.45, "y_scale": 1.0}
            line_chart_scale = {"x_scale": 1.65, "y_scale": 1.0}
            chart_gap_after_table = 6

            # Column widths (0-based col index) used for chart sizing.
            # Keep these consistent with table readability.
            col_widths: dict[int, float] = {
                0: 12,  # A
                1: 16,  # B
                2: 16,  # C
                3: 14,  # D
                4: 14,  # E
                5: 16,  # F
                6: 13,  # G
                7: 13,  # H
                8: 13,  # I
                9: 13,  # J
                10: 13,  # K
                11: 13,  # L
                12: 2,  # M (unused on report sheets)
            }
            for c, w in col_widths.items():
                worksheet.set_column(c, c, w)

            now_txt = datetime.now(ZoneInfo(tz_name)).strftime("%Y-%m-%d %H:%M")
            title = f"{metric} 리포트 ({start_date_str} ~ {end_date_str})  생성: {now_txt}"
            worksheet.write(0, 0, title, fmt_title)

            desc = metric_desc.get(metric)
            if desc:
                worksheet.write(2, 0, f"설명: {desc}")

            # ---- KPI summary (no merges; stable grid) ----
            kpi_total_prod = payload.get("kpi_total_prod")
            kpi_spec_rate = payload.get("kpi_spec_rate")
            kpi_valid = payload.get("kpi_valid")
            kpi_over = payload.get("kpi_over")
            kpi_waste = payload.get("kpi_waste")

            fmt_kpi_box = workbook.add_format({"bg_color": "#f3f6fb", "border": 1, "border_color": "#e5e7eb"})
            fmt_kpi_label = workbook.add_format({"font_name": chart_font, "font_size": 11, "bold": True, "color": "#111827"})
            fmt_kpi_value = workbook.add_format({"font_name": chart_font, "font_size": 22, "bold": True, "color": "#111827"})
            fmt_kpi_sub = workbook.add_format({"font_name": chart_font, "font_size": 10, "color": "#6b7280"})

            def _fill_box(r1: int, c1: int, r2: int, c2: int) -> None:
                for rr in range(r1, r2 + 1):
                    for cc in range(c1, c2 + 1):
                        worksheet.write_blank(rr, cc, None, fmt_kpi_box)

            def _write_kpi(slot_col: int, label: str, value: str, sub: str | None, color: str | None) -> None:
                vfmt = fmt_kpi_value if color is None else workbook.add_format({"font_name": chart_font, "font_size": 22, "bold": True, "color": color})
                worksheet.write(1, slot_col, label, fmt_kpi_label)
                worksheet.write(2, slot_col, value, vfmt)
                if sub:
                    worksheet.write(3, slot_col, sub, fmt_kpi_sub)

            # Slots: 2 columns each starting at A,C,E,G,I (no merge)
            # A:B, C:D, E:F, G:H, I:J
            slots = [(0, 1), (2, 3), (4, 5), (6, 7), (8, 9)]
            for c1, c2 in slots:
                _fill_box(1, c1, 3, c2)

            if kpi_total_prod is not None:
                _write_kpi(0, "총 생산량 (pcs)", f"{int(kpi_total_prod):,}", None, None)
            if kpi_spec_rate is not None:
                _write_kpi(2, "규격 대응률 (%)", f"{float(kpi_spec_rate):.1f}%", "일자별(필요SKU∩생산SKU) / 생산SKU", "#1d4ed8")
            if kpi_valid is not None:
                _write_kpi(4, "정확 대응 비중", f"{float(kpi_valid[0]):.1f}%", f"수량: {int(kpi_valid[1]):,} pcs", "#047857")
            if kpi_over is not None:
                _write_kpi(6, "초과 생산 비중", f"{float(kpi_over[0]):.1f}%", f"수량: {int(kpi_over[1]):,} pcs", "#b91c1c")
            if kpi_waste is not None:
                _write_kpi(8, "비정형 생산 비중", f"{float(kpi_waste[0]):.1f}%", f"수량: {int(kpi_waste[1]):,} pcs", "#b45309")

            # ---- Section 1: Factory bar chart + table ----
            worksheet.write(sec1_top, 0, "선택지표 (공장 비교)", fmt_section)
            chart_row = sec1_chart_row
            table_row = sec1_table_row
            if isinstance(factory_table, pd.DataFrame) and len(factory_table) > 0:
                # Write table at fixed row for reporting
                _df_to_sheet(writer, sheet_name=sheet_name, df=factory_table, startrow=table_row, startcol=col0)
                _apply_table_formats(workbook, worksheet, df=factory_table, startrow=table_row, startcol=col0)

                # Build bar chart from table range
                data_first_row = table_row + 1
                data_last_row = table_row + len(factory_table)

                # Determine columns by name
                cols = list(factory_table.columns)
                cat_col = cols.index("공장")
                val_col = cols.index("선택지표") if "선택지표" in cols else (len(cols) - 1)
                y_max = _ymax_0_100(pd.to_numeric(factory_table["선택지표"], errors="coerce").max() if "선택지표" in cols else None)

                categories = f"='{sheet_name}'!{xl_rowcol_to_cell(data_first_row, cat_col)}:{xl_rowcol_to_cell(data_last_row, cat_col)}"
                values = f"='{sheet_name}'!{xl_rowcol_to_cell(data_first_row, val_col)}:{xl_rowcol_to_cell(data_last_row, val_col)}"

                chart = workbook.add_chart({"type": "column"})
                chart.add_series(
                    {
                        "name": metric,
                        "categories": categories,
                        "values": values,
                        "data_labels": {
                            "value": True,
                            "num_format": "0.0\"%\"",
                            "font": {"name": chart_font, "size": 16, "bold": True, "color": "#111827"},
                        },
                        "gap": 70,
                        "overlap": 0,
                        "points": [
                            {"fill": {"color": FACTORY_COLOR_MAP["A관"]}},  # A관
                            {"fill": {"color": FACTORY_COLOR_MAP["C관"]}},  # C관
                            {"fill": {"color": FACTORY_COLOR_MAP["S관"]}},  # S관
                        ],
                        "border": {"none": True},
                    }
                )
                chart.set_title({"name": f"공장별 {metric} (%)", "name_font": title_font})
                chart.set_x_axis(
                    {
                        "name": "",
                        "name_font": axis_title_font,
                        "num_font": axis_num_font,
                        "major_gridlines": {"visible": False},
                        "line": {"none": True},
                        "tick_mark": "none",
                    }
                )
                chart.set_y_axis(
                    {
                        "name": "",
                        "min": 0,
                        "max": 100,
                        "name_font": axis_title_font,
                        "num_font": axis_num_font,
                        "major_gridlines": {"visible": True, "line": {"color": gridline_color}},
                        "line": {"none": True},
                        "tick_mark": "none",
                    }
                )
                chart.set_legend({"none": True})
                chart.set_style(10)
                chart.set_plotarea({"border": {"none": True}, "fill": {"color": "#ffffff"}})
                chart.set_chartarea({"border": {"none": True}, "fill": {"color": "#ffffff"}})
                # Fit chart into A6:G23 (1-based). 0-based rows 5..22, cols 0..6.
                wpx, hpx = _chart_box_pixels(
                    col_widths=col_widths,
                    row_height_points=row_height_points,
                    first_col=0,
                    last_col=6,
                    first_row=5,
                    last_row=22,
                )
                chart.set_size({"width": wpx, "height": hpx})
                worksheet.insert_chart(chart_row, 0, chart)
            else:
                worksheet.write(table_row, 0, "데이터 없음")

            # ---- Section 2: Daily line chart + table ----
            sec2_top = max(sec2_top_min, table_row + (len(factory_table) + chart_gap_after_table if isinstance(factory_table, pd.DataFrame) else 12))
            worksheet.write(sec2_top, 0, "일별요약", fmt_section)

            chart2_row = max(sec2_chart_row_min, sec2_top + 1)
            table2_row = max(sec2_table_row_min, chart2_row + 18)

            line_ts_df = payload.get("line_ts_df")
            if isinstance(line_ts_df, pd.DataFrame) and len(line_ts_df) > 0:
                # Pivot to wide for chart source
                tmp = line_ts_df.copy()
                tmp["기간"] = pd.to_datetime(tmp["기간"], errors="coerce")
                tmp = tmp.dropna(subset=["기간"])
                wide = tmp.pivot_table(index="기간", columns="공장", values="값", aggfunc="mean").reset_index()

                # Extend x-axis like dashboard:
                # - 당월/전월: extend to month-end (values after last data remain blank)
                # - 기간조회: keep selected end_date
                filter_option = payload.get("filter_option")
                axis_start = _start_date_obj
                axis_end = _end_date_obj
                axis_bucket = "D"
                if axis_start is not None and axis_end is not None and filter_option in {"당월", "전월"}:
                    try:
                        axis_end = _month_end(axis_start)
                    except Exception:
                        axis_end = _end_date_obj

                if axis_start is not None and axis_end is not None:
                    # Match dashboard behavior:
                    # - 당월/전월: always use daily axis (D) even if we extend to month-end.
                    # - 기간조회: choose D/W/M by span.
                    if filter_option in {"당월", "전월"}:
                        axis_bucket = "D"
                    else:
                        span_days = (axis_end - axis_start).days + 1
                        if span_days <= 30:
                            axis_bucket = "D"
                        elif span_days <= 210:
                            axis_bucket = "W"
                        else:
                            axis_bucket = "M"
                    axis_idx = _build_axis(axis_start, axis_end, axis_bucket)
                    full_df = pd.DataFrame({"기간": axis_idx})
                    wide["기간"] = pd.to_datetime(wide["기간"], errors="coerce")
                    wide = full_df.merge(wide, on="기간", how="left")
                # Put chart source into DATA sheet (written at end; keep ranges stable now)
                src_row = data_next_row
                src_col = 0
                pending_data_blocks.append((wide, src_row, src_col))
                date_col = src_col
                date_first = src_row + 1
                date_last = src_row + len(wide)
                data_next_row = date_last + 3

                chart2 = workbook.add_chart({"type": "line"})
                y2_max = _ymax_0_100(pd.to_numeric(tmp["값"], errors="coerce").max())
                chart2.set_title({"name": f"공장별 {metric} 추이", "name_font": title_font})
                # Keep x-axis settings consistent with the axis bucket used for the chart source.
                bucket = axis_bucket

                x_axis_opts = {
                    "name": "기간",
                    "name_font": axis_title_font,
                    "num_font": axis_num_font,
                    "num_format": "yyyy-mm-dd",
                    "label_position": "low",
                    "major_gridlines": {"visible": False},
                    "line": {"none": True},
                    "tick_mark": "none",
                    "date_axis": True,
                }
                if bucket == "W":
                    x_axis_opts.update({"major_unit": 7, "major_unit_type": "days"})
                elif bucket == "M":
                    x_axis_opts.update({"num_format": "yyyy-mm", "major_unit": 1, "major_unit_type": "months"})
                chart2.set_x_axis(x_axis_opts)
                chart2.set_y_axis(
                    {
                        "name": "",
                        "min": 0,
                        "max": 100,
                        "name_font": axis_title_font,
                        "num_font": axis_num_font,
                        "major_gridlines": {"visible": True, "line": {"color": gridline_color}},
                        "line": {"none": True},
                        "tick_mark": "none",
                    }
                )
                chart2.set_legend({"position": "top", "font": legend_font})
                chart2.set_style(10)
                chart2.set_plotarea({"border": {"none": True}, "fill": {"color": "#ffffff"}})
                chart2.set_chartarea({"border": {"none": True}, "fill": {"color": "#ffffff"}})

                series_colors = [FACTORY_COLOR_MAP["A관"], FACTORY_COLOR_MAP["C관"], FACTORY_COLOR_MAP["S관"]]

                for j, col_name in enumerate(wide.columns[1:], start=1):
                    val_c = src_col + j
                    chart2.add_series(
                        {
                            "name": str(col_name),
                            "categories": f"='{data_sheet_name}'!{xl_rowcol_to_cell(date_first, date_col)}:{xl_rowcol_to_cell(date_last, date_col)}",
                            "values": f"='{data_sheet_name}'!{xl_rowcol_to_cell(date_first, val_c)}:{xl_rowcol_to_cell(date_last, val_c)}",
                            "line": {"width": 2.75, "color": series_colors[(j - 1) % len(series_colors)]},
                            "marker": {"type": "none"},
                            "smooth": True,
                        }
                    )

                # Fit chart into A34:T51 (1-based). 0-based rows 33..50, cols 0..19.
                # Use default width for columns beyond our set; approximate at 8.43.
                wpx, hpx = _chart_box_pixels(
                    col_widths=col_widths,
                    row_height_points=row_height_points,
                    first_col=0,
                    last_col=19,
                    first_row=33,
                    last_row=50,
                )
                chart2.set_size({"width": wpx, "height": hpx})
                worksheet.insert_chart(chart2_row, 0, chart2)

            if isinstance(daily_table, pd.DataFrame) and len(daily_table) > 0:
                _df_to_sheet(writer, sheet_name=sheet_name, df=daily_table, startrow=table2_row, startcol=col0)
                _apply_table_formats(workbook, worksheet, df=daily_table, startrow=table2_row, startcol=col0)
            else:
                worksheet.write(table2_row, 0, "데이터 없음")

            # ---- Section 3: Factory daily detail table ----
            sec3_top = table2_row + (len(daily_table) + 6 if isinstance(daily_table, pd.DataFrame) else 14)
            worksheet.write(sec3_top, 0, "관별(공장별) 일별상세", fmt_section)
            table3_row = sec3_top + 1
            if isinstance(factory_daily_table, pd.DataFrame) and len(factory_daily_table) > 0:
                _df_to_sheet(writer, sheet_name=sheet_name, df=factory_daily_table, startrow=table3_row, startcol=col0)
                _apply_table_formats(workbook, worksheet, df=factory_daily_table, startrow=table3_row, startcol=col0)
            else:
                worksheet.write(table3_row, 0, "데이터 없음")

            # Freeze title + KPI rows
            worksheet.freeze_panes(4, 0)
            # Do not override per-column number formats set by _apply_table_formats.
            # A4 landscape print-friendly
            try:
                worksheet.set_landscape()
                worksheet.set_paper(9)  # A4
                worksheet.fit_to_pages(1, 0)
            except Exception:
                pass

        # Write DATA sheet last so it appears at the end of tabs.
        if pending_data_blocks:
            for df_block, r0, c0 in pending_data_blocks:
                _write_chart_source_df(writer, data_sheet_name, df=df_block, startrow=r0, startcol=c0)

    output.seek(0)
    return output.getvalue()


def _build_factory_bar_fig(*, factory_data: pd.DataFrame, metric_option: str) -> tuple[pd.DataFrame, go.Figure | None]:
    metric_map = {
        "규격 대응률": ("규격대응률(%)", "유효생산량"),
        "정확 대응 비중": ("유효비율(%)", "유효생산량"),
        "초과 생산 비중": ("과생산비율(%)", "과생산량"),
        "비정형 생산 비중": ("불필요비율(%)", "불필요생산량"),
    }
    metric_col, pcs_col = metric_map[metric_option]

    df = factory_data.copy()
    if metric_col not in df.columns:
        df[metric_col] = np.nan
    df["선택지표"] = pd.to_numeric(df[metric_col], errors="coerce").replace([np.inf, -np.inf], 0).fillna(0).clip(0, 100)

    hover_data = {
        "총실적": ":,",
        "유효생산량": ":,",
        "과생산량": ":,",
        "불필요생산량": ":,",
        "생산SKU수": ":,",
        "필요대응SKU수": ":,",
        "규격대응률(%)": ":.1f",
        "유효비율(%)": ":.1f",
        "과생산비율(%)": ":.1f",
        "불필요비율(%)": ":.1f",
        "선택지표": ":.1f",
    }
    hover_data = {k: v for k, v in hover_data.items() if k in df.columns}

    table_cols = ["공장", "총실적", pcs_col, metric_col, "선택지표"]
    table_cols = [c for c in table_cols if c in df.columns]
    # Export/UI 공통 컬러(공장별)
    factories = [f for f in df.get("공장", pd.Series([], dtype="object")).dropna().astype(str).unique().tolist()]
    color_map = _factory_color_discrete_map(factories)

    try:
        fig = px.bar(
            df,
            x="공장",
            y="선택지표",
            color="공장",
            title=f"공장별 {metric_option} (%)",
            text="선택지표",
            color_discrete_map=color_map if color_map else None,
        )
        fig.update_traces(
            texttemplate="%{text:.1f}%",
            textposition="outside",
            textfont=dict(size=24, family="Arial", color="#222222"),
            marker=dict(cornerradius="15"),
            cliponaxis=False,
        )
        fig.update_layout(
            height=520,
            showlegend=False,
            margin=dict(l=0, r=0, t=60, b=0),
            yaxis=dict(range=[0, 105], title=dict(text=f"{metric_option} (%)", font=dict(size=16, family="Arial", color="#222222"))),
            xaxis=dict(
                title=dict(text="공장", font=dict(size=16, family="Arial", color="#222222")),
                tickfont=dict(size=18, family="Arial", color="#222222"),
            ),
            title=dict(font=dict(size=22, family="Arial", color="#111111")),
        )
    except Exception:
        fig = None

    return df[table_cols].copy(), fig


def _build_factory_line_fig(
    *,
    metric_option: str,
    factory_summary_filtered: pd.DataFrame,
    sku_daily_factory: pd.DataFrame | None,
    sku_daily_all: pd.DataFrame | None,
    start_date,
    end_date,
    today,
) -> go.Figure | None:
    if factory_summary_filtered is None or len(factory_summary_filtered) == 0:
        return None

    display_start_date = start_date
    display_end_date = end_date

    span_days = (display_end_date - display_start_date).days + 1
    if span_days <= 30:
        bucket = "D"
    elif span_days <= 210:
        bucket = "W"
    else:
        bucket = "M"

    axis = _build_axis(display_start_date, display_end_date, bucket)
    tickvals, ticktext = _build_tick_labels(axis, bucket)

    factories = [f for f in factory_summary_filtered["공장"].dropna().astype(str).unique().tolist()]
    if not factories:
        return None

    ts_rows: list[dict] = []
    if metric_option != "규격 대응률":
        base_ts = factory_summary_filtered[
            ["생산일자_date", "공장", "총실적", "유효생산량", "과생산량", "불필요생산량"]
        ].copy()
        base_ts["date"] = pd.to_datetime(base_ts["생산일자_date"], errors="coerce")
        base_ts = base_ts.dropna(subset=["date"])
        base_ts["period"] = _period_start(base_ts["date"], bucket)
        agg = base_ts.groupby(["period", "공장"], dropna=False).agg(
            total=("총실적", "sum"),
            valid=("유효생산량", "sum"),
            over=("과생산량", "sum"),
            waste=("불필요생산량", "sum"),
        ).reset_index()

        num_col = {
            "정확 대응 비중": "valid",
            "초과 생산 비중": "over",
            "비정형 생산 비중": "waste",
        }[metric_option]
        agg["value"] = np.where(agg["total"] > 0, agg[num_col] / agg["total"] * 100, np.nan)
        agg["value"] = pd.to_numeric(agg["value"], errors="coerce").clip(0, 100)

        for _, r in agg.iterrows():
            ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
    else:
        spec_done = False
        required_cols_ts = {"날짜_date", "공장", "생산SKU수", "필요대응SKU수"}
        if sku_daily_factory is not None and len(sku_daily_factory) > 0 and required_cols_ts.issubset(set(sku_daily_factory.columns)):
            day_counts_ts = sku_daily_factory[
                (sku_daily_factory["날짜_date"] >= start_date) &
                (sku_daily_factory["날짜_date"] <= end_date) &
                (sku_daily_factory["날짜_date"] != today)
            ].copy()
            if len(day_counts_ts) > 0:
                day_counts_ts["date"] = pd.to_datetime(day_counts_ts["날짜_date"], errors="coerce")
                day_counts_ts = day_counts_ts.dropna(subset=["date"])
                day_counts_ts["period"] = _period_start(day_counts_ts["date"], bucket)
                agg_ts = day_counts_ts.groupby(["period", "공장"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                agg_ts["value"] = np.where(
                    agg_ts["생산SKU수"] > 0,
                    agg_ts["필요대응SKU수"] / agg_ts["생산SKU수"] * 100,
                    np.nan,
                )
                agg_ts["value"] = pd.to_numeric(agg_ts["value"], errors="coerce").clip(0, 100)
                for _, r in agg_ts.iterrows():
                    ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
                spec_done = True

        if (not spec_done) and (sku_daily_all is not None) and (len(sku_daily_all) > 0) and {"날짜_date", "생산SKU수", "필요대응SKU수"}.issubset(set(sku_daily_all.columns)):
            daily_spec = sku_daily_all[
                (sku_daily_all["날짜_date"] >= start_date) &
                (sku_daily_all["날짜_date"] <= end_date) &
                (sku_daily_all["날짜_date"] != today)
            ].copy()
            if len(daily_spec) > 0:
                daily_spec["date"] = pd.to_datetime(daily_spec["날짜_date"], errors="coerce")
                daily_spec = daily_spec.dropna(subset=["date"])
                daily_spec["period"] = _period_start(daily_spec["date"], bucket)
                agg_spec = daily_spec.groupby(["period"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                agg_spec["value"] = np.where(
                    agg_spec["생산SKU수"] > 0,
                    agg_spec["필요대응SKU수"] / agg_spec["생산SKU수"] * 100,
                    np.nan,
                )
                agg_spec["value"] = pd.to_numeric(agg_spec["value"], errors="coerce").clip(0, 100)
                for _, r in agg_spec.iterrows():
                    for f in factories:
                        ts_rows.append({"기간": r["period"], "공장": f, "값": r["value"]})

    ts_df = pd.DataFrame(ts_rows)
    if len(ts_df) == 0:
        return None

    ts_df["기간"] = pd.to_datetime(ts_df["기간"], errors="coerce")
    full_grid = pd.MultiIndex.from_product([axis, factories], names=["기간", "공장"]).to_frame(index=False)
    ts_df = full_grid.merge(ts_df, on=["기간", "공장"], how="left")
    label_map = {pd.Timestamp(v): t for v, t in zip(tickvals, ticktext, strict=False)}
    ts_df["x_label"] = ts_df["기간"].map(label_map)

    color_map = _factory_color_discrete_map(factories)
    line_fig = px.line(
        ts_df,
        x="기간",
        y="값",
        color="공장",
        title=f"공장별 {metric_option} 추이",
        markers=False,
        custom_data=["x_label"],
        color_discrete_map=color_map if color_map else None,
    )
    line_fig.update_traces(line=dict(width=3.5), hovertemplate="공장=%{legendgroup}<br>기간=%{customdata[0]}<br>값=%{y:.1f}%<extra></extra>")
    line_fig.update_layout(
        height=320,
        margin=dict(l=0, r=0, t=60, b=0),
        yaxis=dict(range=[0, 105], title=f"{metric_option} (%)", tickformat=".1f"),
        xaxis=dict(tickmode="array", tickvals=tickvals, ticktext=ticktext, tickangle=-45, tickfont=dict(size=10)),
        legend_title_text="공장",
    )
    return line_fig


def _build_factory_line_ts_df(
    *,
    metric_option: str,
    factory_summary_filtered: pd.DataFrame,
    sku_daily_factory: pd.DataFrame | None,
    sku_daily_all: pd.DataFrame | None,
    start_date,
    end_date,
    today,
) -> pd.DataFrame:
    if factory_summary_filtered is None or len(factory_summary_filtered) == 0:
        return pd.DataFrame()

    span_days = (end_date - start_date).days + 1
    if span_days <= 30:
        bucket = "D"
    elif span_days <= 210:
        bucket = "W"
    else:
        bucket = "M"

    axis = _build_axis(start_date, end_date, bucket)
    factories = [f for f in factory_summary_filtered["공장"].dropna().astype(str).unique().tolist()]
    if not factories:
        return pd.DataFrame()

    ts_rows: list[dict] = []
    if metric_option != "규격 대응률":
        base_ts = factory_summary_filtered[
            ["생산일자_date", "공장", "총실적", "유효생산량", "과생산량", "불필요생산량"]
        ].copy()
        base_ts["date"] = pd.to_datetime(base_ts["생산일자_date"], errors="coerce")
        base_ts = base_ts.dropna(subset=["date"])
        base_ts["period"] = _period_start(base_ts["date"], bucket)
        agg = base_ts.groupby(["period", "공장"], dropna=False).agg(
            total=("총실적", "sum"),
            valid=("유효생산량", "sum"),
            over=("과생산량", "sum"),
            waste=("불필요생산량", "sum"),
        ).reset_index()

        num_col = {
            "정확 대응 비중": "valid",
            "초과 생산 비중": "over",
            "비정형 생산 비중": "waste",
        }[metric_option]
        agg["value"] = np.where(agg["total"] > 0, agg[num_col] / agg["total"] * 100, np.nan)
        agg["value"] = pd.to_numeric(agg["value"], errors="coerce").clip(0, 100)
        for _, r in agg.iterrows():
            ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
    else:
        spec_done = False
        required_cols_ts = {"날짜_date", "공장", "생산SKU수", "필요대응SKU수"}
        if sku_daily_factory is not None and len(sku_daily_factory) > 0 and required_cols_ts.issubset(set(sku_daily_factory.columns)):
            day_counts_ts = sku_daily_factory[
                (sku_daily_factory["날짜_date"] >= start_date) &
                (sku_daily_factory["날짜_date"] <= end_date) &
                (sku_daily_factory["날짜_date"] != today)
            ].copy()
            if len(day_counts_ts) > 0:
                day_counts_ts["date"] = pd.to_datetime(day_counts_ts["날짜_date"], errors="coerce")
                day_counts_ts = day_counts_ts.dropna(subset=["date"])
                day_counts_ts["period"] = _period_start(day_counts_ts["date"], bucket)
                agg_ts = day_counts_ts.groupby(["period", "공장"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                agg_ts["value"] = np.where(
                    agg_ts["생산SKU수"] > 0,
                    agg_ts["필요대응SKU수"] / agg_ts["생산SKU수"] * 100,
                    np.nan,
                )
                agg_ts["value"] = pd.to_numeric(agg_ts["value"], errors="coerce").clip(0, 100)
                for _, r in agg_ts.iterrows():
                    ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
                spec_done = True

        if (not spec_done) and (sku_daily_all is not None) and (len(sku_daily_all) > 0) and {"날짜_date", "생산SKU수", "필요대응SKU수"}.issubset(set(sku_daily_all.columns)):
            daily_spec = sku_daily_all[
                (sku_daily_all["날짜_date"] >= start_date) &
                (sku_daily_all["날짜_date"] <= end_date) &
                (sku_daily_all["날짜_date"] != today)
            ].copy()
            if len(daily_spec) > 0:
                daily_spec["date"] = pd.to_datetime(daily_spec["날짜_date"], errors="coerce")
                daily_spec = daily_spec.dropna(subset=["date"])
                daily_spec["period"] = _period_start(daily_spec["date"], bucket)
                agg_spec = daily_spec.groupby(["period"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                agg_spec["value"] = np.where(
                    agg_spec["생산SKU수"] > 0,
                    agg_spec["필요대응SKU수"] / agg_spec["생산SKU수"] * 100,
                    np.nan,
                )
                agg_spec["value"] = pd.to_numeric(agg_spec["value"], errors="coerce").clip(0, 100)
                for _, r in agg_spec.iterrows():
                    for f in factories:
                        ts_rows.append({"기간": r["period"], "공장": f, "값": r["value"]})

    ts_df = pd.DataFrame(ts_rows)
    if len(ts_df) == 0:
        return pd.DataFrame()

    ts_df["기간"] = pd.to_datetime(ts_df["기간"], errors="coerce")
    full_grid = pd.MultiIndex.from_product([axis, factories], names=["기간", "공장"]).to_frame(index=False)
    ts_df = full_grid.merge(ts_df, on=["기간", "공장"], how="left")
    ts_df["값"] = pd.to_numeric(ts_df["값"], errors="coerce")
    return ts_df


def _month_end(d: datetime.date) -> datetime.date:
    last_day = calendar.monthrange(d.year, d.month)[1]
    return datetime(d.year, d.month, last_day).date()


def _period_start(ts: pd.Series, bucket: str) -> pd.Series:
    if bucket == "D":
        return ts.dt.normalize()
    if bucket == "W":
        return ts.dt.normalize() - pd.to_timedelta(ts.dt.weekday, unit="D")
    return ts.dt.to_period("M").dt.to_timestamp()


def _build_axis(start_d: datetime.date, end_d: datetime.date, bucket: str) -> pd.DatetimeIndex:
    start_ts = pd.Timestamp(start_d)
    end_ts = pd.Timestamp(end_d)
    if bucket == "D":
        return pd.date_range(start_ts, end_ts, freq="D")
    if bucket == "W":
        start_monday = start_ts.normalize() - pd.to_timedelta(start_ts.weekday(), unit="D")
        end_monday = end_ts.normalize() - pd.to_timedelta(end_ts.weekday(), unit="D")
        return pd.date_range(start_monday, end_monday, freq="W-MON")
    start_ms = start_ts.to_period("M").to_timestamp()
    end_ms = end_ts.to_period("M").to_timestamp()
    return pd.date_range(start_ms, end_ms, freq="MS")


def _build_tick_labels(axis: pd.DatetimeIndex, bucket: str) -> tuple[list[pd.Timestamp], list[str]]:
    tickvals = [pd.Timestamp(x) for x in axis.to_list()]
    if bucket == "D":
        ticktext = [x.strftime("%m-%d") for x in tickvals]
        return tickvals, ticktext
    if bucket == "W":
        years = {x.year for x in tickvals}
        iso_weeks = pd.Series(tickvals).dt.isocalendar().week.astype(int).tolist()
        if len(years) > 1:
            ticktext = [f"{x.year % 100:02d}W{w}" for x, w in zip(tickvals, iso_weeks, strict=False)]
        else:
            ticktext = [f"W{w}" for w in iso_weeks]
        return tickvals, ticktext
    # bucket == "M"
    years = {x.year for x in tickvals}
    if len(years) > 1:
        ticktext = [f"{x.year}-{x.month}월" for x in tickvals]
    else:
        ticktext = [f"{x.month}월" for x in tickvals]
    return tickvals, ticktext

# 페이지 설정
DASHBOARD_TITLE = "생산 운영 현황 대시보드"
KPI_LABEL_MAP = {
    "총실적": "총 생산량",
    "총부족수량": "필요 수량",
    "유효생산량": "정확 대응 생산량",
    "과생산량": "초과 생산량",
    "불필요생산량": "비정형 생산량",
}
RATE_LABEL_MAP = {
    "유효비율(%)": "정확 대응 비중(%)",
    "과생산비율(%)": "초과 생산 비중(%)",
    "불필요비율(%)": "비정형 생산 비중(%)",
}

st.set_page_config(page_title=DASHBOARD_TITLE, layout="wide", initial_sidebar_state="collapsed")

# CSS 스타일링
st.markdown("""
<style>
    [data-testid="metric.container"] {
        background-color: #f0f4f8;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .kpi-card {
        background-color: #f0f4f8;
        border-radius: 12px;
        padding: 16px 16px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        height: 100%;
    }
    .kpi-head {
        display: flex;
        justify-content: space-between;
        align-items: baseline;
        gap: 10px;
        margin-bottom: 6px;
    }
    .kpi-title {
        font-size: 14px;
        font-weight: 700;
        color: #374151;
        line-height: 1.2;
    }
    .kpi-right {
        font-size: 13px;
        font-weight: 800;
        white-space: nowrap;
    }
    .kpi-value {
        font-size: clamp(22px, 2.2vw, 34px);
        font-weight: 900;
        color: #111827;
        letter-spacing: 0.3px;
        line-height: 1.0;
        margin: 0;
    }
    .kpi-sub {
        margin-top: 8px;
        font-size: 12px;
        color: #6b7280;
        line-height: 1.2;
    }
    .kpi-split {
        background-color: #f0f4f8;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        display: flex;
        overflow: hidden;
        height: 100%;
    }
    .kpi-cell {
        padding: 16px 16px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        min-width: 0;
    }
    .kpi-cell.left {
        flex: 1.6;
    }
    .kpi-cell.right {
        flex: 1.0;
    }
    .kpi-divider {
        width: 1px;
        background: rgba(17, 24, 39, 0.12);
        margin: 14px 0;
        flex: 0 0 1px;
    }
    .kpi-cell-title {
        font-size: 14px;
        font-weight: 700;
        color: #374151;
        line-height: 1.2;
        margin-bottom: 6px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .kpi-cell-value {
        font-size: clamp(22px, 2.2vw, 34px);
        font-weight: 900;
        color: #111827;
        letter-spacing: 0.3px;
        line-height: 1.0;
        margin: 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .kpi-cell-sub {
        margin-top: 8px;
        font-size: 12px;
        color: #6b7280;
        line-height: 1.2;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    @media (max-width: 900px) {
        .kpi-split {
            flex-direction: column;
        }
        .kpi-divider {
            width: 100%;
            height: 1px;
            margin: 0 14px;
        }
        .kpi-cell.left, .kpi-cell.right {
            flex: unset;
        }
    }
    h1 {
        text-align: center;
        color: #1f3a93;
        margin-bottom: 90px;
    }
    h2 {
        color: #2c5aa0;
        border-bottom: 3px solid #2c5aa0;
        padding-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)


def render_kpi_card(title: str, value: str, right_label: str | None = None, right_value: float | None = None, right_color: str = "#111827", sub: str | None = None) -> None:
    right_html = ""
    if right_label is not None and right_value is not None:
        right_html = f"<span class='kpi-right' style='color:{right_color};'>{right_label} {right_value:.1f}%</span>"
    sub_html = f"<div class='kpi-sub'>{sub}</div>" if sub else ""
    st.markdown(
        f"""
<div class="kpi-card">
  <div class="kpi-head">
    <div class="kpi-title">{title}</div>
    {right_html}
  </div>
  <div class="kpi-value">{value}</div>
  {sub_html}
</div>
""",
        unsafe_allow_html=True,
    )


def render_kpi_split_card(
    left_title: str,
    left_value: str,
    right_title: str,
    right_value: str,
    right_sub: str | None = None,
) -> None:
    right_sub_html = f"<div class='kpi-cell-sub'>{right_sub}</div>" if right_sub else ""
    st.markdown(
        f"""
<div class="kpi-split">
  <div class="kpi-cell left">
    <div class="kpi-cell-title">{left_title}</div>
    <div class="kpi-cell-value">{left_value}</div>
  </div>
  <div class="kpi-divider"></div>
  <div class="kpi-cell right">
    <div class="kpi-cell-title">{right_title}</div>
    <div class="kpi-cell-value">{right_value}</div>
    {right_sub_html}
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_result_excels(
    result_paths: tuple[str, ...],
    mtime_nss: tuple[int, ...],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """결과 엑셀 파일(여러 개 가능) 로드 + 전처리 (Streamlit 캐시 적용)

    - 월별로 파일이 분리되어 저장되는 경우(예: 유효생산량_결과_2026-05.xlsx)에도
      전월/기간조회가 동작하도록 여러 결과 파일을 합쳐서 사용합니다.
    - mtime_nss는 파일 변경 시 캐시 무효화를 위해 사용됩니다.
    """
    _ = mtime_nss  # cache key only

    required_sheets = ["매칭결과", "일별요약", "공장_신규분류별"]
    matching_frames: list[pd.DataFrame] = []
    daily_frames: list[pd.DataFrame] = []
    factory_frames: list[pd.DataFrame] = []

    for path_str, mtime_ns in zip(result_paths, mtime_nss, strict=False):
        path = Path(path_str)
        sheets = pd.read_excel(path, sheet_name=required_sheets)
        required = set(required_sheets)
        missing = required - set(sheets.keys())
        if missing:
            raise ValueError(f"결과 엑셀에 필요한 시트가 없습니다({path.name}): {', '.join(sorted(missing))}")

        mr = sheets["매칭결과"].copy()
        ds = sheets["일별요약"].copy()
        fs = sheets["공장_신규분류별"].copy()

        mr["_source_mtime_ns"] = mtime_ns
        ds["_source_mtime_ns"] = mtime_ns
        fs["_source_mtime_ns"] = mtime_ns

        matching_frames.append(mr)
        daily_frames.append(ds)
        factory_frames.append(fs)

    matching_result = pd.concat(matching_frames, ignore_index=True) if matching_frames else pd.DataFrame()
    daily_summary = pd.concat(daily_frames, ignore_index=True) if daily_frames else pd.DataFrame()
    factory_summary = pd.concat(factory_frames, ignore_index=True) if factory_frames else pd.DataFrame()

    if len(matching_result) > 0:
        matching_result["날짜"] = pd.to_datetime(matching_result["날짜"], errors="coerce")
        matching_result["생산일자"] = pd.to_datetime(matching_result["생산일자"], errors="coerce")
        matching_result = matching_result.sort_values("_source_mtime_ns", kind="stable")
        dedup_cols = [c for c in ["날짜", "생산일자", "공장", "신규분류요약", "제품코드"] if c in matching_result.columns]
        if dedup_cols:
            matching_result = matching_result.drop_duplicates(subset=dedup_cols, keep="last")
        matching_result["날짜_date"] = matching_result["날짜"].dt.date
        matching_result["생산일자_date"] = matching_result["생산일자"].dt.date
        matching_result = matching_result.drop(columns=["_source_mtime_ns"], errors="ignore")

    if len(daily_summary) > 0:
        daily_summary["날짜"] = pd.to_datetime(daily_summary["날짜"], errors="coerce")
        daily_summary = daily_summary[daily_summary["날짜"].notna()].copy()
        daily_summary = daily_summary.sort_values("_source_mtime_ns", kind="stable")
        if "날짜" in daily_summary.columns:
            daily_summary = daily_summary.drop_duplicates(subset=["날짜"], keep="last")
        daily_summary["날짜_date"] = daily_summary["날짜"].dt.date
        daily_summary = daily_summary.drop(columns=["_source_mtime_ns"], errors="ignore")

    if len(factory_summary) > 0:
        factory_summary["생산일자"] = pd.to_datetime(factory_summary["생산일자"], errors="coerce")
        factory_summary = factory_summary[factory_summary["생산일자"].notna()].copy()
        factory_summary = factory_summary.sort_values("_source_mtime_ns", kind="stable")
        dedup_cols = [c for c in ["생산일자", "공장", "신규분류요약"] if c in factory_summary.columns]
        if dedup_cols:
            factory_summary = factory_summary.drop_duplicates(subset=dedup_cols, keep="last")
        factory_summary["생산일자_date"] = factory_summary["생산일자"].dt.date
        factory_summary = factory_summary.drop(columns=["_source_mtime_ns"], errors="ignore")

    # SKU 기반 일자 규격 대응률(전사/공장별) 프리컴퓨트 (대용량 groupby는 1회만 수행)
    sku_daily_all = pd.DataFrame()
    sku_daily_factory = pd.DataFrame()
    required_cols = {"날짜_date", "제품코드", "양품수량", "부족수량", "유효생산량"}
    if len(matching_result) > 0 and required_cols.issubset(set(matching_result.columns)):
        base = matching_result[matching_result["제품코드"].notna()].copy()
        for col in ["양품수량", "부족수량", "유효생산량"]:
            base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0)
        base["_need_qty"] = (base["유효생산량"] + base["부족수량"]).fillna(0)

        by_day_sku = base.groupby(["날짜_date", "제품코드"], dropna=False).agg(
            prod_qty=("양품수량", "sum"),
            need_qty=("_need_qty", "sum"),
        ).reset_index()
        by_day_sku["produced_flag"] = by_day_sku["prod_qty"] > 0
        by_day_sku["need_flag"] = by_day_sku["need_qty"] > 0

        produced_skus = (
            by_day_sku[by_day_sku["produced_flag"]]
            .groupby("날짜_date", dropna=False)["제품코드"]
            .nunique()
            .rename("생산SKU수")
        )
        needed_skus = (
            by_day_sku[by_day_sku["produced_flag"] & by_day_sku["need_flag"]]
            .groupby("날짜_date", dropna=False)["제품코드"]
            .nunique()
            .rename("필요대응SKU수")
        )
        sku_daily_all = pd.concat([produced_skus, needed_skus], axis=1).fillna(0).reset_index()
        sku_daily_all["규격대응률(%)"] = np.where(
            sku_daily_all["생산SKU수"] > 0,
            sku_daily_all["필요대응SKU수"] / sku_daily_all["생산SKU수"] * 100,
            0,
        )
        sku_daily_all["규격대응률(%)"] = sku_daily_all["규격대응률(%)"].clip(0, 100)
        sku_daily_all = sku_daily_all.sort_values("날짜_date").reset_index(drop=True)

        if "공장" in matching_result.columns:
            base_f = matching_result[(matching_result["공장"].notna()) & (matching_result["제품코드"].notna())].copy()
            for col in ["양품수량", "부족수량", "유효생산량"]:
                base_f[col] = pd.to_numeric(base_f[col], errors="coerce").fillna(0)
            base_f["_need_qty"] = (base_f["유효생산량"] + base_f["부족수량"]).fillna(0)

            by_day_factory_sku = base_f.groupby(["날짜_date", "공장", "제품코드"], dropna=False).agg(
                prod_qty=("양품수량", "sum"),
                need_qty=("_need_qty", "sum"),
            ).reset_index()
            by_day_factory_sku["produced_flag"] = by_day_factory_sku["prod_qty"] > 0
            by_day_factory_sku["need_flag"] = by_day_factory_sku["need_qty"] > 0

            produced_f = (
                by_day_factory_sku[by_day_factory_sku["produced_flag"]]
                .groupby(["날짜_date", "공장"], dropna=False)["제품코드"]
                .nunique()
                .rename("생산SKU수")
            )
            needed_f = (
                by_day_factory_sku[by_day_factory_sku["produced_flag"] & by_day_factory_sku["need_flag"]]
                .groupby(["날짜_date", "공장"], dropna=False)["제품코드"]
                .nunique()
                .rename("필요대응SKU수")
            )
            sku_daily_factory = pd.concat([produced_f, needed_f], axis=1).fillna(0).reset_index()
            sku_daily_factory["규격대응률(%)"] = np.where(
                sku_daily_factory["생산SKU수"] > 0,
                sku_daily_factory["필요대응SKU수"] / sku_daily_factory["생산SKU수"] * 100,
                0,
            )
            sku_daily_factory["규격대응률(%)"] = sku_daily_factory["규격대응률(%)"].clip(0, 100)
            sku_daily_factory = sku_daily_factory.sort_values(["날짜_date", "공장"]).reset_index(drop=True)

    return matching_result, daily_summary, factory_summary, sku_daily_all, sku_daily_factory


def _months_between(start_d: date, end_d: date) -> tuple[str, ...]:
    if start_d is None or end_d is None:
        return tuple()
    if end_d < start_d:
        start_d, end_d = end_d, start_d
    periods = pd.period_range(start=start_d, end=end_d, freq="M")
    return tuple(str(p) for p in periods)


def _store_dir_from_user_input(*, base_dir: Path) -> Path:
    env = os.environ.get("APS_YIELD_STORE_PATH", "").strip()
    if env:
        return Path(env)
    return base_dir / "outputs" / "store"


def _store_has_table(store_dir: Path, table: str) -> bool:
    tdir = store_dir / table
    return tdir.exists() and tdir.is_dir() and any(tdir.glob("*.parquet"))


@st.cache_data(show_spinner=False)
def list_store_months(store_dir_str: str, table: str) -> tuple[str, ...]:
    store_dir = Path(store_dir_str)
    tdir = store_dir / table
    if not tdir.exists():
        return tuple()
    months = sorted([p.stem for p in tdir.glob("*.parquet") if p.is_file()])
    return tuple(months)


@st.cache_data(show_spinner=False)
def load_store_table(
    store_dir_str: str,
    table: str,
    months: tuple[str, ...],
    columns: tuple[str, ...] | None = None,
) -> pd.DataFrame:
    store_dir = Path(store_dir_str)
    tdir = store_dir / table
    if not tdir.exists():
        return pd.DataFrame()

    frames: list[pd.DataFrame] = []
    for ym in months:
        p = tdir / f"{ym}.parquet"
        if not p.exists():
            continue
        try:
            frames.append(pd.read_parquet(p, columns=list(columns) if columns else None))
        except Exception:
            continue

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _ensure_date_column(df: pd.DataFrame, *, src_col: str, out_col: str) -> pd.DataFrame:
    if df is None or len(df) == 0:
        return df if df is not None else pd.DataFrame()
    if out_col in df.columns:
        return df
    if src_col not in df.columns:
        return df
    df = df.copy()
    df[src_col] = pd.to_datetime(df[src_col], errors="coerce")
    df[out_col] = df[src_col].dt.date
    return df


def _normalize_spec_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or len(df) == 0:
        return df if df is not None else pd.DataFrame()
    df = df.copy()
    # 전처리 산출(부족대응SKU수)을 대시보드 기대(필요대응SKU수)로 호환
    if "필요대응SKU수" not in df.columns and "부족대응SKU수" in df.columns:
        df["필요대응SKU수"] = df["부족대응SKU수"]
    return df


@st.cache_data(show_spinner=False)
def load_process_balance_excels(
    result_paths: tuple[str, ...],
    mtime_nss: tuple[int, ...],
) -> tuple[pd.DataFrame, bool]:
    """유효생산량_결과2 엑셀 로드 (공정별_일별실적) + 전처리

    - 시트가 없는 파일은 스킵합니다(오류 없이 안내용).
    - 반환: (dataframe, has_any_sheet)
    """
    _ = mtime_nss  # cache key only

    frames: list[pd.DataFrame] = []
    has_any_sheet = False

    for path_str, mtime_ns in zip(result_paths, mtime_nss, strict=False):
        path = Path(path_str)
        try:
            df = pd.read_excel(path, sheet_name="공정별_일별실적")
        except Exception:
            continue

        if df is None or len(df) == 0:
            has_any_sheet = True
            continue

        has_any_sheet = True
        df = df.copy()
        df["_source_mtime_ns"] = mtime_ns
        frames.append(df)

    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if len(out) == 0:
        return out, has_any_sheet

    if "날짜" in out.columns:
        out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce")
        out = out[out["날짜"].notna()].copy()
        out = out.sort_values("_source_mtime_ns", kind="stable")
        dedup_cols = [c for c in ["날짜", "공장", "공정"] if c in out.columns]
        if dedup_cols:
            out = out.drop_duplicates(subset=dedup_cols, keep="last")
        out["날짜_date"] = out["날짜"].dt.date
        out = out.drop(columns=["_source_mtime_ns"], errors="ignore")

    for col in ["실적수량", "부족수량", "과생산수량"]:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    return out.reset_index(drop=True), has_any_sheet


@st.cache_data(show_spinner=False)
def load_process_balance_detail_excels(
    result_paths: tuple[str, ...],
    mtime_nss: tuple[int, ...],
) -> tuple[pd.DataFrame, bool]:
    """유효생산량_결과2 엑셀 로드 (매칭결과) + 전처리

    - 시트가 없는 파일은 스킵합니다(오류 없이 안내용).
    - 반환: (dataframe, has_any_sheet)
    """
    _ = mtime_nss  # cache key only

    frames: list[pd.DataFrame] = []
    has_any_sheet = False

    for path_str, mtime_ns in zip(result_paths, mtime_nss, strict=False):
        path = Path(path_str)
        try:
            df = pd.read_excel(path, sheet_name="매칭결과")
        except Exception:
            continue

        if df is None or len(df) == 0:
            has_any_sheet = True
            continue

        has_any_sheet = True
        df = df.copy()
        df["_source_mtime_ns"] = mtime_ns
        frames.append(df)

    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if len(out) == 0:
        return out, has_any_sheet

    # 최신 파일 우선(중복 제거용)
    out = out.sort_values("_source_mtime_ns", kind="stable")

    if "날짜" in out.columns:
        out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce")
        out = out[out["날짜"].notna()].copy()
        out["날짜_date"] = out["날짜"].dt.date

    dedup_cols = [c for c in ["날짜_date", "공장", "공정", "신규분류요약", "제품코드"] if c in out.columns]
    if dedup_cols:
        out = out.drop_duplicates(subset=dedup_cols, keep="last")

    for col in ["실적수량", "필요수량", "부족수량", "유효생산량", "과생산량", "불필요생산량"]:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    out = out.drop(columns=["_source_mtime_ns"], errors="ignore")
    return out.reset_index(drop=True), has_any_sheet


@st.cache_data(show_spinner=False)
def load_process_balance_prepared(
    result_paths: tuple[str, ...],
    mtime_nss: tuple[int, ...],
) -> tuple[pd.DataFrame, pd.DataFrame, bool]:
    """공정 밸런스 계산용으로 '매칭결과'를 미리 집계해 둔 데이터 반환(빠른 기간 필터/리런 목적).

    반환:
    - proc_base: 일자/공장/공정 단위 집계(수량 + SKU 카운트)
    - det_base:  일자/공장/공정/신규분류요약 단위 집계(상세 테이블용)
    - has_any_sheet
    """
    _ = mtime_nss  # cache key only

    det, has_any_sheet = load_process_balance_detail_excels(result_paths, mtime_nss)
    if det is None or len(det) == 0:
        return pd.DataFrame(), pd.DataFrame(), has_any_sheet

    det = det.copy()

    # Backward-compat for older 결과2 파일
    if "공정" in det.columns:
        det["공정"] = det["공정"].replace({"최종공정": "누수규격"})

    target_order = ["사출", "분리", "하드레이션", "접착", "누수규격"]
    if "공정" in det.columns:
        det = det[det["공정"].isin(target_order)].copy()

    # 공장 그룹(A/C/S관) 매핑
    if "공장" in det.columns:
        det["공장그룹"] = np.select(
            [
                det["공장"].astype(str).str.contains("A관", na=False),
                det["공장"].astype(str).str.contains("C관", na=False),
                det["공장"].astype(str).str.contains("S관", na=False),
            ],
            ["A관", "C관", "S관"],
            default="기타",
        )
    else:
        det["공장그룹"] = "기타"

    # 제품명: 제품코드 앞자리 5글자(예: Q1230)
    if "제품코드" in det.columns:
        det["제품명"] = det["제품코드"].astype(str).str.slice(0, 5)

    # 정확/초과/비정형 재계산(집계 전에 수행)
    if {"실적수량", "필요수량"}.issubset(set(det.columns)):
        _prod = pd.to_numeric(det["실적수량"], errors="coerce").fillna(0).clip(lower=0)
        _need = pd.to_numeric(det["필요수량"], errors="coerce").fillna(0).clip(lower=0)
        det["유효생산량"] = np.minimum(_prod, _need)
        det["불필요생산량"] = np.where(_need <= 0, _prod, 0.0)
        det["과생산량"] = np.where(_need > 0, np.maximum(_prod - _need, 0.0), 0.0)

    group_keys = [c for c in ["날짜_date", "공장", "공장그룹", "공정"] if c in det.columns]
    if not group_keys:
        return pd.DataFrame(), pd.DataFrame(), has_any_sheet

    # 수량 집계(일자/공장/공정)
    agg_cols = [c for c in ["실적수량", "유효생산량", "과생산량", "불필요생산량", "부족수량", "필요수량"] if c in det.columns]
    qty = det.groupby(group_keys, dropna=False)[agg_cols].sum().reset_index() if agg_cols else det[group_keys].drop_duplicates()
    for c in agg_cols:
        qty[c] = pd.to_numeric(qty[c], errors="coerce").fillna(0)

    # SKU 집계(제품명 기준) - 한번만(비용 큰 연산)
    if "제품명" in det.columns:
        produced = (
            det[pd.to_numeric(det.get("실적수량", 0), errors="coerce").fillna(0) > 0]
            .groupby(group_keys, dropna=False)["제품명"]
            .nunique()
            .rename("생산SKU수")
        )
        needed = (
            det[pd.to_numeric(det.get("필요수량", 0), errors="coerce").fillna(0) > 0]
            .groupby(group_keys, dropna=False)["제품명"]
            .nunique()
            .rename("필요SKU수")
        )
        inter = (
            det[
                (pd.to_numeric(det.get("실적수량", 0), errors="coerce").fillna(0) > 0)
                & (pd.to_numeric(det.get("필요수량", 0), errors="coerce").fillna(0) > 0)
            ]
            .groupby(group_keys, dropna=False)["제품명"]
            .nunique()
            .rename("규격대응SKU수")
        )
        sku = pd.concat([produced, needed, inter], axis=1).fillna(0).reset_index()
    else:
        sku = pd.DataFrame(columns=group_keys + ["생산SKU수", "필요SKU수", "규격대응SKU수"])

    proc_base = qty.merge(sku, on=group_keys, how="left")
    for c in ["생산SKU수", "필요SKU수", "규격대응SKU수"]:
        if c in proc_base.columns:
            proc_base[c] = pd.to_numeric(proc_base[c], errors="coerce").fillna(0)

    # 상세 테이블용 집계(일자/공장/공정/신규분류요약)
    det_group_cols = [c for c in ["날짜_date", "공장", "공장그룹", "공정", "신규분류요약"] if c in det.columns]
    det_value_cols = [c for c in ["실적수량", "필요수량", "부족수량", "유효생산량", "과생산량", "불필요생산량"] if c in det.columns]
    if det_group_cols and det_value_cols:
        det_base = det.groupby(det_group_cols, dropna=False)[det_value_cols].sum().reset_index()
    else:
        det_base = pd.DataFrame()

    return proc_base.reset_index(drop=True), det_base.reset_index(drop=True), has_any_sheet


def _truncate_err_message(msg: str, *, max_chars: int = 600) -> str:
    msg = str(msg or "")
    msg = " ".join(msg.split())
    if len(msg) <= max_chars:
        return msg
    return textwrap.shorten(msg, width=max_chars, placeholder=" …(truncated)")


def _safe_dataframe(
    df: pd.DataFrame,
    *,
    fmt: dict[str, str] | None = None,
    max_style_rows: int = 2000,
    height: int | None = None,
    hide_index: bool = True,
) -> None:
    kwargs: dict[str, Any] = {"use_container_width": True, "hide_index": hide_index}
    if height is not None:
        kwargs["height"] = int(height)
    if df is None:
        st.dataframe(pd.DataFrame(), **kwargs)
        return
    if fmt and len(df) <= max_style_rows:
        st.dataframe(df.style.format(fmt), **kwargs)
        return
    if fmt and len(df) > max_style_rows:
        st.caption(f"표시 행이 많아({len(df):,}행) 스타일링을 생략하고 표시합니다.")
    st.dataframe(df, **kwargs)


# 결과 파일 선택(월별 분리 저장 지원)
try:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))
    base_dir = Path(BASE_PATH)

    st.sidebar.markdown("### 데이터 소스")
    _default_store_dir = _store_dir_from_user_input(base_dir=base_dir)
    store_dir_str = st.sidebar.text_input("Parquet store 경로", value=str(_default_store_dir))
    store_dir = Path(store_dir_str) if store_dir_str else _default_store_dir
    store_available = _store_has_table(store_dir, "result1_daily")
    use_store = st.sidebar.toggle("Parquet(store) 사용(추천)", value=store_available, disabled=not store_available)
    if st.sidebar.button("캐시 비우기"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.sidebar.success("캐시를 비웠습니다. 새로고침하세요.")

    # 결과 파일이 repo 루트뿐 아니라 `outputs/` 아래에 저장되는 경우도 있어 함께 검색합니다.
    search_dirs = [base_dir, base_dir / "outputs", base_dir / "outputs" / "archive"]

    if use_store:
        # result1은 Parquet(store)로만 로드(엑셀 로드는 생략)
        result_candidates = []

        # 공정 밸런스용 결과2 파일(전공정)은 기존 엑셀도 지원(호환 목적)
        _cands2: list[Path] = []
        for d in search_dirs:
            if not d.exists():
                continue
            _cands2.extend([p for p in d.glob("유효생산량_결과2*.xlsx") if not p.name.startswith("~$")])

        _seen2: set[str] = set()
        result2_candidates = []
        for p in _cands2:
            rp = str(p.resolve())
            if rp in _seen2:
                continue
            _seen2.add(rp)
            result2_candidates.append(p)

        result2_candidates = sorted(
            result2_candidates,
            key=lambda p: p.stat().st_mtime_ns if p.exists() else 0,
            reverse=True,
        )
    else:
        _cands: list[Path] = []
        for d in search_dirs:
            if not d.exists():
                continue
            _cands.extend(
                [
                    p
                    for p in d.glob("유효생산량_결과*.xlsx")
                    if (not p.name.startswith("~$")) and (not p.name.startswith("유효생산량_결과2"))
                ]
            )

        _seen: set[str] = set()
        result_candidates = []
        for p in _cands:
            rp = str(p.resolve())
            if rp in _seen:
                continue
            _seen.add(rp)
            result_candidates.append(p)

        result_candidates = sorted(
            result_candidates,
            key=lambda p: p.stat().st_mtime_ns if p.exists() else 0,
            reverse=True,
        )
        if not result_candidates:
            st.error(
                "⚠️ 결과 파일을 찾을 수 없습니다. 검색 경로: "
                + ", ".join(str(d) for d in search_dirs)
            )
            st.info("전처리 완료된 결과 파일(`유효생산량_결과*.xlsx`)을 repo 루트 또는 `outputs/`에 넣어주세요.")
            st.stop()

        # 공정 밸런스용 결과2 파일(전공정) 후보 검색
        _cands2: list[Path] = []
        for d in search_dirs:
            if not d.exists():
                continue
            _cands2.extend([p for p in d.glob("유효생산량_결과2*.xlsx") if not p.name.startswith("~$")])

        _seen2: set[str] = set()
        result2_candidates = []
        for p in _cands2:
            rp = str(p.resolve())
            if rp in _seen2:
                continue
            _seen2.add(rp)
            result2_candidates.append(p)

        result2_candidates = sorted(
            result2_candidates,
            key=lambda p: p.stat().st_mtime_ns if p.exists() else 0,
            reverse=True,
        )
except Exception as e:
    st.error("❌ 초기화(파일 검색) 중 오류가 발생했습니다.")
    st.code(_truncate_err_message(str(e)), language="text")
    st.stop()

try:
    if use_store:
        _store_dir_str = str(store_dir)
        _months_all = list_store_months(_store_dir_str, "result1_daily")
        daily_summary = load_store_table(_store_dir_str, "result1_daily", _months_all)
        daily_summary = _ensure_date_column(daily_summary, src_col="날짜", out_col="날짜_date")

        matching_result = pd.DataFrame()
        factory_summary = pd.DataFrame()
        sku_daily_all = pd.DataFrame()
        sku_daily_factory = pd.DataFrame()

        # 공정 밸런스(전공정): 기간 선택 후 필요한 월만 로드(초기에는 빈 DF)
        process_daily = pd.DataFrame()
        process_has_sheet = False
        process_detail = pd.DataFrame()
        process_detail_has_sheet = False
        process_proc_base = pd.DataFrame()
        process_det_base = pd.DataFrame()
    else:
        # 최신 파일이 월별로 분리되어 저장될 수 있어, 후보 파일들을 합쳐서 사용
        result_paths = tuple(str(p) for p in result_candidates)
        mtime_nss = tuple(int(p.stat().st_mtime_ns) for p in result_candidates)
        matching_result, daily_summary, factory_summary, sku_daily_all, sku_daily_factory = load_result_excels(result_paths, mtime_nss)

        # 공정 밸런스: 결과2 로드(없으면 빈 DF)
        process_daily = pd.DataFrame()
        process_has_sheet = False
        process_detail = pd.DataFrame()
        process_detail_has_sheet = False
        process_proc_base = pd.DataFrame()
        process_det_base = pd.DataFrame()
        if result2_candidates:
            result2_paths = tuple(str(p) for p in result2_candidates)
            mtime2_nss = tuple(int(p.stat().st_mtime_ns) for p in result2_candidates)
            process_daily, process_has_sheet = load_process_balance_excels(result2_paths, mtime2_nss)
            process_detail, process_detail_has_sheet = load_process_balance_detail_excels(result2_paths, mtime2_nss)
            process_proc_base, process_det_base, _ = load_process_balance_prepared(result2_paths, mtime2_nss)

    # 금일 데이터 제외 (아직 생산 중이므로) - KST 기준
    now_kst = datetime.now(ZoneInfo("Asia/Seoul"))
    today = now_kst.date()
    st.caption(f"기준 시각(KST): {now_kst.strftime('%Y-%m-%d %H:%M:%S')}")

    # 필수 데이터 검증/정규화
    if daily_summary is None or len(daily_summary) == 0 or "날짜_date" not in daily_summary.columns:
        st.error("⚠️ `일별요약` 시트에 날짜 데이터가 없어서 대시보드를 표시할 수 없습니다. (컬럼: `날짜`)")
        st.info("`유효생산량_결과*.xlsx`를 최신 버전으로 다시 생성한 뒤, repo 루트 또는 `outputs/`에 넣어주세요.")
        st.stop()

    for col in ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]:
        if col in daily_summary.columns:
            daily_summary[col] = pd.to_numeric(daily_summary[col], errors="coerce").fillna(0)

    factory_has_dates = factory_summary is not None and len(factory_summary) > 0 and "생산일자_date" in factory_summary.columns
    if factory_has_dates:
        for col in ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]:
            if col in factory_summary.columns:
                factory_summary[col] = pd.to_numeric(factory_summary[col], errors="coerce").fillna(0)
    else:
        factory_summary = pd.DataFrame()

    # 제목
    st.markdown(f"<h1 style='text-align:center; color:#1f3a93; margin:0;'>🏭 {DASHBOARD_TITLE}</h1>", unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # 메인 탭(세션 유지): 기간조회 선택 후에도 탭 이동 시 첫 탭으로 돌아가는 현상 방지
    main_tab = st.radio(
        "메인 탭",
        ["생산 운영 현황", "공정 밸런스"],
        horizontal=True,
        label_visibility="collapsed",
        key="main_tab",
    )

    # 기간 필터 (기본: 당월)
    filter_option = st.radio("조회 기간", ["당월", "전월", "기간조회"], horizontal=True, label_visibility="collapsed")

    # 날짜 범위 계산
    current_month_start = datetime(today.year, today.month, 1).date()
    data_max_date = daily_summary["날짜_date"].max()
    current_month_end = (today - pd.Timedelta(days=1))  # 어제까지 (date)
    if pd.notna(data_max_date):
        current_month_end = min(current_month_end, data_max_date)

    # 전월 계산
    first_day_current = current_month_start
    last_day_prev = first_day_current - pd.Timedelta(days=1)
    prev_month_start = datetime(last_day_prev.year, last_day_prev.month, 1).date()

    # 전체 기간(데이터 기준) 계산 (기간조회 범위 제한용)
    full_min_date = daily_summary[daily_summary["날짜_date"] != today]["날짜_date"].min()
    full_max_date = daily_summary[daily_summary["날짜_date"] != today]["날짜_date"].max()

    # 공정 밸런스는 결과2(매칭결과) 기준으로 가능한 최소일을 사용 (없으면 전체 기간과 동일)
    process_full_min_date = None
    process_full_max_date = None
    if process_detail is not None and len(process_detail) > 0 and "날짜_date" in process_detail.columns:
        _proc_dates = process_detail[process_detail["날짜_date"] != today]["날짜_date"]
        if len(_proc_dates) > 0:
            process_full_min_date = _proc_dates.min()
            process_full_max_date = _proc_dates.max()

    # 날짜 범위 결정
    if filter_option == "당월":
        start_date = current_month_start
        end_date = current_month_end
    elif filter_option == "전월":
        start_date = prev_month_start
        end_date = last_day_prev
    else:  # 기간조회
        if main_tab == "공정 밸런스" and process_full_min_date is not None and process_full_max_date is not None:
            min_date = process_full_min_date
            max_date = process_full_max_date
        else:
            min_date = full_min_date
            max_date = full_max_date
        if pd.isna(min_date) or pd.isna(max_date):
            st.warning("선택 가능한 날짜 범위를 계산할 수 없습니다. (데이터 없음)")
            min_date = today
            max_date = today

        col_filter1, col_space, col_filter2 = st.columns([1.5, 0.2, 1.5])

        def _clamp_date(d: date, lo: date, hi: date) -> date:
            return max(lo, min(hi, d))

        # 기간조회 기본값:
        # - 과거 맨 첫날로 가지 않고, "현재 월 1일 ~ 데이터 최신일"로 시작 (예: 6/3이면 6/1~6/3)
        _default_start = min_date if isinstance(min_date, date) else today
        _default_end = max_date if isinstance(max_date, date) else today
        try:
            _month_start = datetime(today.year, today.month, 1).date()
            if isinstance(min_date, date) and isinstance(max_date, date):
                if _month_start <= max_date:
                    _default_start = max(min_date, _month_start)
                _default_end = min(max_date, today) if today <= max_date else max_date
        except Exception:
            pass
        _ss_start = st.session_state.get("range_start", _default_start)
        _ss_end = st.session_state.get("range_end", _default_end)
        _ss_start = _clamp_date(_ss_start, min_date, max_date)
        _ss_end = _clamp_date(_ss_end, min_date, max_date)

        with col_filter1:
            start_date = st.date_input("시작 날짜", value=_ss_start, min_value=min_date, max_value=max_date, key="range_start")

        with col_filter2:
            end_date = st.date_input("종료 날짜", value=_ss_end, min_value=min_date, max_value=max_date, key="range_end")

    if start_date > end_date:
        st.warning("시작 날짜가 종료 날짜보다 커서 자동으로 교체했습니다.")
        start_date, end_date = end_date, start_date

    # 공정 밸런스는 결과2에 존재하는 기간 내로 강제(공정 밸런스 최초일 이전 데이터는 신뢰 불가)
    if main_tab == "공정 밸런스" and process_full_min_date is not None and process_full_max_date is not None:
        if start_date < process_full_min_date:
            st.info(f"공정 밸런스 데이터는 {process_full_min_date}부터 있어 시작일을 자동 조정했습니다.")
            start_date = process_full_min_date
            if start_date > end_date:
                end_date = start_date
        if end_date > process_full_max_date:
            st.info(f"공정 밸런스 데이터는 {process_full_max_date}까지 있어 종료일을 자동 조정했습니다.")
            end_date = process_full_max_date
            if start_date > end_date:
                start_date = end_date

    st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)

    # Parquet(store) 모드: 기간 선택 후 해당 월 데이터만 로드(메모리/속도 안정화)
    if use_store:
        _store_dir_str = str(store_dir)
        _months_in_range = _months_between(start_date, end_date)

        # 대용량(매칭결과)은 기본 미로드. 필요한 화면/기능에서만 사용하도록 옵션 제공.
        detail_available = _store_has_table(store_dir, "result1_matching")
        load_detail = st.sidebar.toggle(
            "상세(매칭결과) 로드",
            value=False,
            disabled=not detail_available,
        )
        if not detail_available:
            st.sidebar.caption("상세는 저장되지 않았습니다(WRITE_DETAIL_STORE=0).")

        matching_result = pd.DataFrame()
        if load_detail:
            matching_result = load_store_table(
                _store_dir_str,
                "result1_matching",
                _months_in_range,
                # 컬럼을 제한하면 메모리가 크게 줄어듭니다.
                columns=("날짜", "생산일자", "공장", "신규분류요약", "제품코드", "양품수량", "부족수량"),
            )

        factory_summary = load_store_table(_store_dir_str, "result1_factory", _months_in_range)
        sku_daily_all = _normalize_spec_cols(load_store_table(_store_dir_str, "result1_spec_daily", _months_in_range))
        sku_daily_factory = _normalize_spec_cols(load_store_table(_store_dir_str, "result1_spec_factory_daily", _months_in_range))
        sku_daily_factory_class = _normalize_spec_cols(load_store_table(_store_dir_str, "result1_spec_factory_class_daily", _months_in_range))

        matching_result = _ensure_date_column(matching_result, src_col="날짜", out_col="날짜_date")
        matching_result = _ensure_date_column(matching_result, src_col="생산일자", out_col="생산일자_date")
        factory_summary = _ensure_date_column(factory_summary, src_col="생산일자", out_col="생산일자_date")
        sku_daily_all = _ensure_date_column(sku_daily_all, src_col="날짜", out_col="날짜_date")
        sku_daily_factory = _ensure_date_column(sku_daily_factory, src_col="날짜", out_col="날짜_date")
        sku_daily_factory_class = _ensure_date_column(sku_daily_factory_class, src_col="날짜", out_col="날짜_date")

        # store 모드에서는 이 시점에 factory_summary가 확정되므로, 이후 필터링에 사용될 플래그를 갱신합니다.
        factory_has_dates = factory_summary is not None and len(factory_summary) > 0 and "생산일자_date" in factory_summary.columns
        if factory_has_dates:
            for col in ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]:
                if col in factory_summary.columns:
                    factory_summary[col] = pd.to_numeric(factory_summary[col], errors="coerce").fillna(0)

        # 공정 밸런스(전공정): store 우선 사용(엑셀 로드로 인한 메모리 폭증 방지)
        if main_tab == "공정 밸런스":
            process_daily = load_store_table(_store_dir_str, "result2_process_daily", _months_in_range)
            process_daily = _ensure_date_column(process_daily, src_col="날짜", out_col="날짜_date")
            process_has_sheet = len(process_daily) > 0

            process_proc_base = load_store_table(_store_dir_str, "result2_proc_base", _months_in_range)
            process_det_base = load_store_table(_store_dir_str, "result2_det_base", _months_in_range)
            process_proc_base = _ensure_date_column(process_proc_base, src_col="날짜", out_col="날짜_date")
            process_det_base = _ensure_date_column(process_det_base, src_col="날짜", out_col="날짜_date")
            process_detail = pd.DataFrame()
            process_detail_has_sheet = len(process_proc_base) > 0

    # 필터 적용 (기간 범위)
    daily_summary_filtered = daily_summary[
        (daily_summary["날짜_date"] >= start_date) &
        (daily_summary["날짜_date"] <= end_date) &
        (daily_summary["날짜_date"] != today)
    ]

    if factory_has_dates and len(factory_summary) > 0:
        factory_summary_filtered = factory_summary[
            (factory_summary["생산일자_date"] >= start_date) &
            (factory_summary["생산일자_date"] <= end_date) &
            (factory_summary["생산일자_date"] != today)
        ]
    else:
        factory_summary_filtered = pd.DataFrame()

    # 메트릭 계산
    total_prod = int(daily_summary_filtered["총실적"].sum()) if len(daily_summary_filtered) > 0 else 0
    valid_prod = int(daily_summary_filtered["유효생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    over_prod = int(daily_summary_filtered["과생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    waste_prod = int(daily_summary_filtered["불필요생산량"].sum()) if len(daily_summary_filtered) > 0 else 0
    prod_days = int(daily_summary_filtered["날짜_date"].nunique()) if len(daily_summary_filtered) > 0 else 0

    valid_rate = (valid_prod / total_prod * 100) if total_prod > 0 else 0
    over_rate = (over_prod / total_prod * 100) if total_prod > 0 else 0
    waste_rate = (waste_prod / total_prod * 100) if total_prod > 0 else 0

    # 규격 대응률(SKU 기준): "그날 생산한 SKU" 중 "그날 필요(수요)가 있던 SKU" 비율
    # - 사용자 정의: (일자별 필요 SKU ∩ 일자별 생산 SKU) / 일자별 생산 SKU
    # - 공장별 규격 대응률도 동일 기준으로 계산하려면 `매칭결과` 시트에 `공장` 컬럼이 필요합니다.
    shortage_prod_daily = None
    shortage_prod_rate = None
    if sku_daily_all is not None and len(sku_daily_all) > 0 and {"날짜_date", "생산SKU수", "필요대응SKU수", "규격대응률(%)"}.issubset(set(sku_daily_all.columns)):
        shortage_prod_daily = sku_daily_all[
            (sku_daily_all["날짜_date"] >= start_date) &
            (sku_daily_all["날짜_date"] <= end_date) &
            (sku_daily_all["날짜_date"] != today)
        ].copy()
        if len(shortage_prod_daily) > 0:
            produced_skus_total = float(pd.to_numeric(shortage_prod_daily["생산SKU수"], errors="coerce").fillna(0).sum())
            need_responded_skus_total = float(pd.to_numeric(shortage_prod_daily["필요대응SKU수"], errors="coerce").fillna(0).sum())
            shortage_prod_rate = (need_responded_skus_total / produced_skus_total * 100) if produced_skus_total > 0 else None

    if main_tab == "생산 운영 현황":
        colA, col3, col4, col5 = st.columns([2.6, 1.1, 1.1, 1.1])
        with colA:
            spec_value = f"{shortage_prod_rate:.1f}%" if shortage_prod_rate is not None else "-"
            spec_sub = "일자별 (필요SKU∩생산SKU) / 생산SKU"
            if shortage_prod_rate is None:
                spec_sub = "계산 불가: 매칭결과에 제품코드/수량 필요"
            render_kpi_split_card(
                f"{KPI_LABEL_MAP['총실적']} (pcs)",
                f"{total_prod:,}",
                "규격 대응률 (%)",
                f"<span style='color:#1d4ed8'>{spec_value}</span>",
                right_sub=spec_sub,
            )

        with col3:
            render_kpi_card(
                "정확 대응 비중",
                f"<span style='color:#047857'>{valid_rate:.1f}%</span>",
                sub=f"수량: {valid_prod:,} pcs",
            )
        with col4:
            render_kpi_card(
                "초과 생산 비중",
                f"<span style='color:#b91c1c'>{over_rate:.1f}%</span>",
                sub=f"수량: {over_prod:,} pcs",
            )
        with col5:
            render_kpi_card(
                "비정형 생산 비중",
                f"<span style='color:#b45309'>{waste_rate:.1f}%</span>",
                sub=f"수량: {waste_prod:,} pcs",
            )

        st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

        with st.expander("지표 정의/상세 보기", expanded=False):
            st.markdown(
                "- `규격 대응률` : 일자별 `(필요 SKU ∩ 생산 SKU) ÷ 생산 SKU` 의 비율\n"
                "- `정확 대응 생산량` : SKU별 `min(생산, 필요)`의 합\n"
                "- `정확 대응 비중` : `정확 대응 생산량` ÷ `총 생산량`\n"
                "- `초과 생산량` : SKU별 `max(생산-필요, 0)`의 합\n"
                "- `초과 생산 비중` : `초과 생산량` ÷ `총 생산량`\n"
                "- `비정형 생산량` : 필요 SKU 외 생산(필요=0인데 생산>0)\n"
                "- `비정형 생산 비중` : `비정형 생산량` ÷ `총 생산량`"
            )
            st.caption("참고: 공장별 `규격 대응률(SKU 기준)`은 `매칭결과` 시트에 `공장`/`제품코드`가 있어야 계산 가능합니다.")

        st.markdown("<div style='margin-top:50px'></div>", unsafe_allow_html=True)

        st.markdown("### 📈 공장별 운영 현황")

        if len(factory_summary_filtered) == 0:
            st.info("선택한 기간에 공장별 데이터가 없습니다.")
        else:
            # 공장별 기간 집계 (정확/초과/비정형 분해)
            factory_data = factory_summary_filtered.groupby("공장", dropna=False).agg(
                {
                    "총실적": "sum",
                    "유효생산량": "sum",
                    "과생산량": "sum",
                    "불필요생산량": "sum",
                }
            ).reset_index()

            factory_data["유효비율(%)"] = (factory_data["유효생산량"] / factory_data["총실적"] * 100).fillna(0)
            factory_data["과생산비율(%)"] = (factory_data["과생산량"] / factory_data["총실적"] * 100).fillna(0)
            factory_data["불필요비율(%)"] = (factory_data["불필요생산량"] / factory_data["총실적"] * 100).fillna(0)

            # 공장별 KPI (정확도 기반)
            factory_data["유효 대응률(수량)(%)"] = factory_data["유효비율(%)"]

            # 공장별 규격 대응률(SKU 기준): 일자별 생산 SKU 중 필요 SKU 비중의 기간 합산
            # - 정의: (Σ 일자별 필요대응SKU수) / (Σ 일자별 생산SKU수)
            sku_coverage_available = False
            sku_coverage_unavailable_reason: str | None = None
            required_cols = {"날짜_date", "공장", "생산SKU수", "필요대응SKU수"}
            if sku_daily_factory is None or len(sku_daily_factory) == 0 or not required_cols.issubset(set(sku_daily_factory.columns)):
                sku_coverage_unavailable_reason = "필수 컬럼 누락(`공장/제품코드/수량/날짜`)"
            else:
                day_counts = sku_daily_factory[
                    (sku_daily_factory["날짜_date"] >= start_date) &
                    (sku_daily_factory["날짜_date"] <= end_date) &
                    (sku_daily_factory["날짜_date"] != today)
                ].copy()
                if len(day_counts) == 0:
                    sku_coverage_unavailable_reason = "선택 기간에 매칭결과 데이터 없음"
                else:
                    sku_counts = day_counts.groupby("공장", dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                    sku_counts["규격대응률(%)"] = np.where(
                        sku_counts["생산SKU수"] > 0,
                        sku_counts["필요대응SKU수"] / sku_counts["생산SKU수"] * 100,
                        0,
                    )
                    sku_counts["규격대응률(%)"] = sku_counts["규격대응률(%)"].clip(0, 100)
                    factory_data = factory_data.merge(
                        sku_counts[["공장", "생산SKU수", "필요대응SKU수", "규격대응률(%)"]],
                        on="공장",
                        how="left",
                    )
                    if "생산SKU수" in factory_data.columns:
                        factory_data["생산SKU수"] = pd.to_numeric(factory_data["생산SKU수"], errors="coerce").fillna(0)
                    else:
                        factory_data["생산SKU수"] = 0
                    if "필요대응SKU수" in factory_data.columns:
                        factory_data["필요대응SKU수"] = pd.to_numeric(factory_data["필요대응SKU수"], errors="coerce").fillna(0)
                    else:
                        factory_data["필요대응SKU수"] = 0
                    if "규격대응률(%)" in factory_data.columns:
                        factory_data["규격대응률(%)"] = (
                            pd.to_numeric(factory_data["규격대응률(%)"], errors="coerce")
                            .replace([np.inf, -np.inf], 0)
                            .fillna(0)
                        )
                    else:
                        factory_data["규격대응률(%)"] = 0.0
                    sku_coverage_available = True

            # NOTE: 규격 대응률은 메인 지표이므로 항상 노출합니다.
            # 계산이 불가한 경우(예: 매칭결과에 공장 없음)에는 "전사 규격 대응률"을 동일 적용해 표시합니다.
            if "규격대응률(%)" not in factory_data.columns:
                factory_data["규격대응률(%)"] = np.nan
            if not sku_coverage_available:
                factory_data["규격대응률(%)"] = float(shortage_prod_rate) if shortage_prod_rate is not None else np.nan
            metric_choices = ["규격 대응률", "정확 대응 비중", "초과 생산 비중", "비정형 생산 비중"]
            radio_key = "factory_metric_option"
            if radio_key not in st.session_state or st.session_state[radio_key] not in metric_choices:
                st.session_state[radio_key] = metric_choices[0]
            metric_option = st.radio("공장 비교 지표", metric_choices, horizontal=True, key=radio_key)
            metric_desc = {
                "규격 대응률": "생산한 SKU(제품코드) 중 필요가 있었던 SKU 비중",
                "정확 대응 비중": "총 생산량 중 정확 대응 생산량이 차지하는 비중",
                "초과 생산 비중": "총 생산량 중 초과 생산량이 차지하는 비중",
                "비정형 생산 비중": "총 생산량 중 비정형 생산량이 차지하는 비중",
            }
            st.caption(f"설명: {metric_desc[metric_option]}")

            metric_map = {
                "규격 대응률": ("규격대응률(%)", "유효생산량"),
                "정확 대응 비중": ("유효비율(%)", "유효생산량"),
                "초과 생산 비중": ("과생산비율(%)", "과생산량"),
                "비정형 생산 비중": ("불필요비율(%)", "불필요생산량"),
            }
            metric_col, pcs_col = metric_map[metric_option]
            factory_data["선택지표"] = factory_data[metric_col].replace([np.inf, -np.inf], 0).fillna(0)
            if metric_option == "규격 대응률" and not sku_coverage_available:
                reason = sku_coverage_unavailable_reason or "원인 미상"
                st.warning(f"공장별 `규격 대응률(SKU 기준)` 계산 불가: {reason}. (전사 규격 대응률을 동일 적용해 표시)")

            hover_data = {
                "총실적": ":,",
                "유효생산량": ":,",
                "과생산량": ":,",
                "불필요생산량": ":,",
                "생산SKU수": ":,",
                "필요대응SKU수": ":,",
                "규격대응률(%)": ":.1f",
                "유효비율(%)": ":.1f",
                "과생산비율(%)": ":.1f",
                "불필요비율(%)": ":.1f",
                "선택지표": ":.1f",
            }
            hover_data = {k: v for k, v in hover_data.items() if k in factory_data.columns}

            _factories_ui = [f for f in factory_data["공장"].dropna().astype(str).unique().tolist()]
            _factory_colors_ui = _factory_color_discrete_map(_factories_ui)

            fig = px.bar(
                factory_data,
                x="공장",
                y="선택지표",
                color="공장",
                title=f"공장별 {metric_option} (%)",
                text="선택지표",
                hover_data=hover_data,
                color_discrete_map=_factory_colors_ui if _factory_colors_ui else None,
            )
            fig.update_traces(
                texttemplate="%{text:.1f}%",
                textposition="outside",
                textfont=dict(size=24, family="Arial", color="#222222"),
                marker=dict(cornerradius="15"),
                cliponaxis=False,
            )
            fig.update_layout(
                height=520,
                showlegend=False,
                margin=dict(l=0, r=0, t=60, b=0),
                yaxis=dict(range=[0, 105], title=dict(text=f"{metric_option} (%)", font=dict(size=16, family="Arial", color="#222222"))),
                xaxis=dict(
                    title=dict(text="공장", font=dict(size=16, family="Arial", color="#222222")),
                    tickfont=dict(size=18, family="Arial", color="#222222")
                ),
                title=dict(font=dict(size=22, family="Arial", color="#111111"))
            )
            st.plotly_chart(fig, use_container_width=True)

            # 공장별 날짜 추이 (라인 차트)
            display_start_date = start_date
            display_end_date = end_date
            if filter_option == "당월":
                display_end_date = _month_end(display_start_date)

            if filter_option in {"당월", "전월"}:
                bucket = "D"
            else:
                span_days = (display_end_date - display_start_date).days + 1
                if span_days <= 30:
                    bucket = "D"
                elif span_days <= 210:
                    bucket = "W"
                else:
                    bucket = "M"
            axis = _build_axis(display_start_date, display_end_date, bucket)
            tickvals, ticktext = _build_tick_labels(axis, bucket)

            factories = [f for f in factory_data["공장"].dropna().astype(str).unique().tolist()]
            ts_rows: list[dict] = []

            if metric_option != "규격 대응률":
                base_ts = factory_summary_filtered[
                    ["생산일자_date", "공장", "총실적", "유효생산량", "과생산량", "불필요생산량"]
                ].copy()
                base_ts["date"] = pd.to_datetime(base_ts["생산일자_date"], errors="coerce")
                base_ts = base_ts.dropna(subset=["date"])
                base_ts["period"] = _period_start(base_ts["date"], bucket)
                agg = base_ts.groupby(["period", "공장"], dropna=False).agg(
                    total=("총실적", "sum"),
                    valid=("유효생산량", "sum"),
                    over=("과생산량", "sum"),
                    waste=("불필요생산량", "sum"),
                ).reset_index()

                num_col = {
                    "정확 대응 비중": "valid",
                    "초과 생산 비중": "over",
                    "비정형 생산 비중": "waste",
                }[metric_option]
                agg["value"] = np.where(agg["total"] > 0, agg[num_col] / agg["total"] * 100, np.nan)
                agg["value"] = pd.to_numeric(agg["value"], errors="coerce").clip(0, 100)

                for _, r in agg.iterrows():
                    ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
            else:
                spec_done = False
                required_cols_ts = {"날짜_date", "공장", "생산SKU수", "필요대응SKU수"}
                if sku_daily_factory is not None and len(sku_daily_factory) > 0 and required_cols_ts.issubset(set(sku_daily_factory.columns)):
                    day_counts_ts = sku_daily_factory[
                        (sku_daily_factory["날짜_date"] >= start_date) &
                        (sku_daily_factory["날짜_date"] <= end_date) &
                        (sku_daily_factory["날짜_date"] != today)
                    ].copy()
                    if len(day_counts_ts) > 0:
                        day_counts_ts["date"] = pd.to_datetime(day_counts_ts["날짜_date"], errors="coerce")
                        day_counts_ts = day_counts_ts.dropna(subset=["date"])
                        day_counts_ts["period"] = _period_start(day_counts_ts["date"], bucket)
                        agg_ts = day_counts_ts.groupby(["period", "공장"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                        agg_ts["value"] = np.where(
                            agg_ts["생산SKU수"] > 0,
                            agg_ts["필요대응SKU수"] / agg_ts["생산SKU수"] * 100,
                            np.nan,
                        )
                        agg_ts["value"] = pd.to_numeric(agg_ts["value"], errors="coerce").clip(0, 100)
                        for _, r in agg_ts.iterrows():
                            ts_rows.append({"기간": r["period"], "공장": r["공장"], "값": r["value"]})
                        spec_done = True

                if (not spec_done) and (sku_daily_all is not None) and (len(sku_daily_all) > 0) and {"날짜_date", "생산SKU수", "필요대응SKU수"}.issubset(set(sku_daily_all.columns)):
                    daily_spec = sku_daily_all[
                        (sku_daily_all["날짜_date"] >= start_date) &
                        (sku_daily_all["날짜_date"] <= end_date) &
                        (sku_daily_all["날짜_date"] != today)
                    ].copy()
                    if len(daily_spec) > 0:
                        daily_spec["date"] = pd.to_datetime(daily_spec["날짜_date"], errors="coerce")
                        daily_spec = daily_spec.dropna(subset=["date"])
                        daily_spec["period"] = _period_start(daily_spec["date"], bucket)
                        agg_spec = daily_spec.groupby(["period"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                        agg_spec["value"] = np.where(
                            agg_spec["생산SKU수"] > 0,
                            agg_spec["필요대응SKU수"] / agg_spec["생산SKU수"] * 100,
                            np.nan,
                        )
                        agg_spec["value"] = pd.to_numeric(agg_spec["value"], errors="coerce").clip(0, 100)
                        for _, r in agg_spec.iterrows():
                            for f in factories:
                                ts_rows.append({"기간": r["period"], "공장": f, "값": r["value"]})

            ts_df = pd.DataFrame(ts_rows)
            if len(ts_df) > 0:
                ts_df["기간"] = pd.to_datetime(ts_df["기간"], errors="coerce")
                full_grid = pd.MultiIndex.from_product([axis, factories], names=["기간", "공장"]).to_frame(index=False)
                ts_df = full_grid.merge(ts_df, on=["기간", "공장"], how="left")
                label_map = {pd.Timestamp(v): t for v, t in zip(tickvals, ticktext, strict=False)}
                ts_df["x_label"] = ts_df["기간"].map(label_map)

                line_fig = px.line(
                    ts_df,
                    x="기간",
                    y="값",
                    color="공장",
                    title=f"공장별 {metric_option} 추이",
                    markers=False,
                    custom_data=["x_label"],
                    color_discrete_map=_factory_colors_ui if _factory_colors_ui else None,
                )
                line_fig.update_traces(
                    line=dict(width=3.5),
                    hovertemplate="공장=%{legendgroup}<br>기간=%{customdata[0]}<br>값=%{y:.1f}%<extra></extra>",
                )
                line_fig.update_layout(
                    height=360,
                    margin=dict(l=0, r=0, t=60, b=0),
                    yaxis=dict(range=[0, 105], title=f"{metric_option} (%)", tickformat=".1f"),
                    xaxis=dict(
                        tickmode="array",
                        tickvals=tickvals,
                        ticktext=ticktext,
                        tickangle=-45,
                        tickfont=dict(size=10),
                    ),
                    legend_title_text="공장",
                )
                st.plotly_chart(line_fig, use_container_width=True)

            st.markdown(f"**선택 지표: {metric_option} (%)**")
            if not sku_coverage_available:
                st.caption("Tip: 공장별 `규격 대응률(SKU 기준)`은 `매칭결과` 시트에 `공장` 컬럼이 있어야 계산 가능합니다.")

            if metric_option == "규격 대응률":
                if not sku_coverage_available:
                    st.info("공장별 SKU 집계가 불가합니다: `매칭결과` 시트에 `공장` 컬럼이 필요합니다.")
                else:
                    # 1) 전처리 산출(공장별_일별_신규분류SKU)이 있으면 그걸 우선 사용(가벼움)
                    if "sku_daily_factory_class" in globals() and sku_daily_factory_class is not None and len(sku_daily_factory_class) > 0:
                        src = sku_daily_factory_class[
                            (sku_daily_factory_class["날짜_date"] >= start_date) &
                            (sku_daily_factory_class["날짜_date"] <= end_date) &
                            (sku_daily_factory_class["날짜_date"] != today)
                        ].copy()
                        if len(src) == 0:
                            st.info("선택한 기간에 신규분류 기준 SKU 집계 데이터가 없습니다.")
                        else:
                            for col in ["생산SKU수", "필요대응SKU수"]:
                                if col in src.columns:
                                    src[col] = pd.to_numeric(src[col], errors="coerce").fillna(0)
                            sku_counts = src.groupby(["공장", "신규분류요약"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                            sku_counts["규격대응률(%)"] = np.where(
                                sku_counts["생산SKU수"] > 0,
                                sku_counts["필요대응SKU수"] / sku_counts["생산SKU수"] * 100,
                                0,
                            )
                            sku_counts["규격대응률(%)"] = sku_counts["규격대응률(%)"].clip(0, 100)
                            factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
                            sku_counts["_factory_sort"] = sku_counts["공장"].map(factory_order)
                            sku_counts = sku_counts.sort_values(["_factory_sort", "신규분류요약"]).reset_index(drop=True).drop("_factory_sort", axis=1)
                    # 2) 없으면 기존 방식(매칭결과로 계산)
                    elif matching_result is None or len(matching_result) == 0 or not {"공장", "신규분류요약", "제품코드", "양품수량", "부족수량", "유효생산량", "날짜_date"}.issubset(set(matching_result.columns)):
                        st.info("신규분류 기준 SKU 상세 집계를 위해 전처리 산출물(공장별_일별_신규분류SKU) 또는 `매칭결과`가 필요합니다.")
                    else:
                        mf2 = matching_result[
                            (matching_result["날짜_date"] >= start_date) &
                            (matching_result["날짜_date"] <= end_date) &
                            (matching_result["날짜_date"] != today) &
                            (matching_result["공장"].notna()) &
                            (matching_result["제품코드"].notna())
                        ].copy()

                        if len(mf2) == 0:
                            st.info("선택한 기간에 신규분류 기준 SKU 집계 데이터가 없습니다.")
                        else:
                            for col in ["양품수량", "부족수량", "유효생산량"]:
                                mf2[col] = pd.to_numeric(mf2[col], errors="coerce").fillna(0)
                            mf2["_need_qty"] = (mf2["유효생산량"] + mf2["부족수량"]).fillna(0)

                            by_day_factory_class_sku = mf2.groupby(["날짜_date", "공장", "신규분류요약", "제품코드"], dropna=False).agg(
                                prod_qty=("양품수량", "sum"),
                                need_qty=("_need_qty", "sum"),
                            ).reset_index()
                            by_day_factory_class_sku["produced_flag"] = by_day_factory_class_sku["prod_qty"] > 0
                            by_day_factory_class_sku["need_flag"] = by_day_factory_class_sku["need_qty"] > 0

                            produced = (
                                by_day_factory_class_sku[by_day_factory_class_sku["produced_flag"]]
                                .groupby(["날짜_date", "공장", "신규분류요약"], dropna=False)["제품코드"]
                                .nunique()
                                .rename("생산SKU수")
                            )
                            needed = (
                                by_day_factory_class_sku[by_day_factory_class_sku["produced_flag"] & by_day_factory_class_sku["need_flag"]]
                                .groupby(["날짜_date", "공장", "신규분류요약"], dropna=False)["제품코드"]
                                .nunique()
                                .rename("필요대응SKU수")
                            )
                            day_counts = pd.concat([produced, needed], axis=1).fillna(0).reset_index()
                            sku_counts = day_counts.groupby(["공장", "신규분류요약"], dropna=False)[["생산SKU수", "필요대응SKU수"]].sum().reset_index()
                            sku_counts["규격대응률(%)"] = np.where(
                                sku_counts["생산SKU수"] > 0,
                                sku_counts["필요대응SKU수"] / sku_counts["생산SKU수"] * 100,
                                0,
                            )
                            sku_counts["규격대응률(%)"] = sku_counts["규격대응률(%)"].clip(0, 100)

                            factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
                            sku_counts["_factory_sort"] = sku_counts["공장"].map(factory_order)
                            sku_counts = sku_counts.sort_values(["_factory_sort", "신규분류요약"]).reset_index(drop=True).drop("_factory_sort", axis=1)

                            sku_counts_fmt = sku_counts.copy()
                            sku_counts_fmt["생산SKU수"] = sku_counts_fmt["생산SKU수"].map("{:,.0f}".format)
                            sku_counts_fmt["필요대응SKU수"] = sku_counts_fmt["필요대응SKU수"].map("{:,.0f}".format)
                            sku_counts_fmt["규격대응률(%)"] = sku_counts_fmt["규격대응률(%)"].map("{:.1f}%".format)
                            sku_counts_fmt["신규분류요약"] = sku_counts_fmt["신규분류요약"].fillna("미분류")

                            html_parts = []
                            header_lines = [
                                "<style>",
                                ".custom-table { width: 100%; border-collapse: collapse; font-size: 14px; }",
                                ".custom-table th, .custom-table td { padding: 10px 12px; border: 1px solid #e2e8f0; }",
                                ".custom-table th { background: #f8fafc; color: #111827; text-align: left; }",
                                ".custom-table td { vertical-align: middle; }",
                                ".custom-table td.number { text-align: right; }",
                                ".custom-table tbody tr:nth-child(even) { background: #f8fafc22; }",
                                "</style>",
                                "<table class=\"custom-table\">",
                                "<thead>",
                                "<tr>",
                                "<th>공장</th>",
                                "<th>신규분류요약</th>",
                                "<th>총 생산 SKU</th>",
                                "<th>규격 대응 SKU</th>",
                                "<th>규격 대응률(%)</th>",
                                "</tr>",
                                "</thead>",
                                "<tbody>",
                            ]
                            html_parts.append("\n".join(header_lines) + "\n")

                            grouped = sku_counts_fmt.groupby("공장", sort=False)
                            for factory_name, group in grouped:
                                rowspan = len(group)
                                for idx, row in group.iterrows():
                                    html_parts.append("<tr>")
                                    if idx == group.index[0]:
                                        html_parts.append(f"<td rowspan='{rowspan}' style='vertical-align: middle; font-weight: 600;'>{factory_name}</td>")
                                    html_parts.append(f"<td>{row['신규분류요약']}</td>")
                                    html_parts.append(f"<td class='number'>{row['생산SKU수']}</td>")
                                    html_parts.append(f"<td class='number'>{row['필요대응SKU수']}</td>")
                                    html_parts.append(f"<td class='number'>{row['규격대응률(%)']}</td>")
                                    html_parts.append("</tr>")

                            html_parts.append("</tbody></table>")
                            st.markdown("".join(html_parts), unsafe_allow_html=True)
            else:
                # 공장_신규분류별 통합 현황
                combined_metric_option = metric_option if metric_option in {"정확 대응 비중", "초과 생산 비중", "비정형 생산 비중"} else "정확 대응 비중"
                combined_summary = factory_summary_filtered.groupby(["공장", "신규분류요약"], dropna=False).agg({
                    "총실적": "sum",
                    "유효생산량": "sum",
                    "과생산량": "sum",
                    "불필요생산량": "sum"
                }).reset_index()

                # 비율 계산
                combined_summary["유효비율(%)"] = (combined_summary["유효생산량"] / combined_summary["총실적"] * 100).fillna(0)
                combined_summary["과생산비율(%)"] = (combined_summary["과생산량"] / combined_summary["총실적"] * 100).fillna(0)
                combined_summary["불필요비율(%)"] = (combined_summary["불필요생산량"] / combined_summary["총실적"] * 100).fillna(0)

                combined_summary["유효 대응률(수량)(%)"] = combined_summary["유효비율(%)"]

                # 선택지표 추가 (공장 비교 지표와 동일 3종)
                metric_map = {
                    "정확 대응 비중": ("유효비율(%)", "유효생산량"),
                    "초과 생산 비중": ("과생산비율(%)", "과생산량"),
                    "비정형 생산 비중": ("불필요비율(%)", "불필요생산량"),
                }
                metric_col, pcs_col = metric_map[combined_metric_option]
                combined_summary["선택지표"] = combined_summary[metric_col].fillna(0)

                # 테이블 표시
                base_cols = ["공장", "신규분류요약", "총실적"]
                display_combined = combined_summary[base_cols + [pcs_col, "선택지표"]].copy()
                total_hdr = f"{KPI_LABEL_MAP['총실적']} (pcs)"
                pcs_hdr = f"{KPI_LABEL_MAP[pcs_col]} (pcs)"
                rate_hdr = f"{combined_metric_option} (%)"
                display_combined.columns = ["공장", "신규분류요약", total_hdr, pcs_hdr, rate_hdr]

                # 공장 순서 지정 (A관 > C관 > S관)
                factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
                display_combined["_factory_sort"] = display_combined["공장"].map(factory_order)
                display_combined = display_combined.sort_values(["_factory_sort", "신규분류요약"]).reset_index(drop=True)
                display_combined = display_combined.drop("_factory_sort", axis=1)

                display_combined[total_hdr] = display_combined[total_hdr].map("{:,.0f}".format)
                display_combined[pcs_hdr] = display_combined[pcs_hdr].map("{:,.0f}".format)
                display_combined[rate_hdr] = display_combined[rate_hdr].map("{:.1f}%".format)

                html_parts = []
                # NOTE: Markdown에서는 4칸 이상 들여쓰기된 HTML이 코드블록으로 취급될 수 있어,
                # 모든 라인을 "맨 앞 공백 없이" 생성합니다.
                header_lines = [
                    "<style>",
                    ".custom-table { width: 100%; border-collapse: collapse; font-size: 14px; }",
                    ".custom-table th, .custom-table td { padding: 10px 12px; border: 1px solid #e2e8f0; }",
                    ".custom-table th { background: #f8fafc; color: #111827; text-align: left; }",
                    ".custom-table td { vertical-align: middle; }",
                    ".custom-table td.number { text-align: right; }",
                    ".custom-table tbody tr:nth-child(even) { background: #f8fafc22; }",
                    "</style>",
                    "<table class=\"custom-table\">",
                    "<thead>",
                    "<tr>",
                    "<th>공장</th>",
                    "<th>신규분류요약</th>",
                    f"<th>{total_hdr}</th>",
                ]
                header_lines.extend(
                    [
                        f"<th>{pcs_hdr}</th>",
                        f"<th>{rate_hdr}</th>",
                        "</tr>",
                        "</thead>",
                        "<tbody>",
                    ]
                )
                html_parts.append("\n".join(header_lines) + "\n")

                grouped = display_combined.groupby("공장", sort=False)
                for factory_name, group in grouped:
                    rowspan = len(group)
                    for idx, row in group.iterrows():
                        html_parts.append("<tr>")
                        if idx == group.index[0]:
                            html_parts.append(f"<td rowspan='{rowspan}' style='vertical-align: middle; font-weight: 600;'>{factory_name}</td>")
                        html_parts.append(f"<td>{row['신규분류요약']}</td>")
                        html_parts.append(f"<td class='number'>{row[total_hdr]}</td>")
                        html_parts.append(f"<td class='number'>{row[pcs_hdr]}</td>")
                        html_parts.append(f"<td class='number'>{row[rate_hdr]}</td>")
                        html_parts.append("</tr>")
                html_parts.append("</tbody></table>")
                st.markdown("".join(html_parts), unsafe_allow_html=True)

        # ============== 일별 요약 ==============
        st.markdown("### 📊 일별 요약")

        daily_display = daily_summary_filtered[
            [
                "날짜",
                "날짜_date",
                "총실적",
                "총부족수량",
                "유효생산량",
                "과생산량",
                "불필요생산량",
                "유효비율(%)",
                "과생산비율(%)",
                "불필요비율(%)",
            ]
        ].copy()

        # 일자별 규격 대응률(가능한 경우) 병합
        if shortage_prod_daily is not None and len(shortage_prod_daily) > 0:
            spec_rate = shortage_prod_daily[["날짜_date", "규격대응률(%)"]].copy()
            daily_display = daily_display.merge(spec_rate, on="날짜_date", how="left")
            daily_display["규격대응률(%)"] = daily_display["규격대응률(%)"].fillna(0)

        # 날짜는 일자까지만 표시 (시간 제거)
        daily_display["날짜"] = daily_display["날짜"].dt.strftime("%Y-%m-%d")

        # pcs 컬럼은 콤마 표시 및 컬럼명에 (pcs) 추가
        pcs_cols = ["총실적", "총부족수량", "유효생산량", "과생산량", "불필요생산량"]
        daily_display.rename(
            columns={c: f"{KPI_LABEL_MAP.get(c, c)} (pcs)" for c in pcs_cols},
            inplace=True,
        )
        if "규격대응률(%)" in daily_display.columns:
            daily_display.rename(columns={"규격대응률(%)": "규격 대응률(%)"}, inplace=True)
        daily_display.rename(columns=RATE_LABEL_MAP, inplace=True)

        # 컬럼 순서 정리 (비중(%) 우선, 수량(pcs)은 뒤쪽)
        daily_cols = [
            "날짜",
            f"{KPI_LABEL_MAP['총실적']} (pcs)",
            f"{KPI_LABEL_MAP['총부족수량']} (pcs)",
        ]
        if "규격 대응률(%)" in daily_display.columns:
            daily_cols.append("규격 대응률(%)")
        daily_cols.extend(
            [
                RATE_LABEL_MAP["유효비율(%)"],
                RATE_LABEL_MAP["과생산비율(%)"],
                RATE_LABEL_MAP["불필요비율(%)"],
                f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
            ]
        )
        daily_display = daily_display[daily_cols].copy()

        _safe_dataframe(
            daily_display,
            fmt={
                **{f"{KPI_LABEL_MAP.get(c, c)} (pcs)": "{:,.0f}" for c in pcs_cols},
                **({"규격 대응률(%)": "{:.1f}%"} if "규격 대응률(%)" in daily_display.columns else {}),
                RATE_LABEL_MAP["유효비율(%)"]: "{:.1f}%",
                RATE_LABEL_MAP["과생산비율(%)"]: "{:.1f}%",
                RATE_LABEL_MAP["불필요비율(%)"]: "{:.1f}%",
            },
        )

        with st.expander("🔎 관별(공장별) 일별 상세 펼치기", expanded=False):
            if len(factory_summary_filtered) == 0:
                st.info("선택한 기간에 공장별 데이터가 없습니다.")
            else:
                factory_daily = factory_summary_filtered.groupby(["생산일자_date", "공장"], dropna=False).agg({
                    "총실적": "sum",
                    "총부족수량": "sum",
                    "유효생산량": "sum",
                    "과생산량": "sum",
                    "불필요생산량": "sum",
                }).reset_index()

                factory_daily[RATE_LABEL_MAP["유효비율(%)"]] = (factory_daily["유효생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
                factory_daily[RATE_LABEL_MAP["과생산비율(%)"]] = (factory_daily["과생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)
                factory_daily[RATE_LABEL_MAP["불필요비율(%)"]] = (factory_daily["불필요생산량"] / factory_daily["총실적"] * 100).replace([np.inf, -np.inf], 0).fillna(0)

                # 관별(공장별) 일자 규격 대응률(SKU 기준)
                factory_daily_spec = None
                factory_daily_spec_by_day = None
                if sku_daily_factory is not None and len(sku_daily_factory) > 0 and {"날짜_date", "공장", "규격대응률(%)"}.issubset(set(sku_daily_factory.columns)):
                    factory_daily_spec = sku_daily_factory[
                        (sku_daily_factory["날짜_date"] >= start_date) &
                        (sku_daily_factory["날짜_date"] <= end_date) &
                        (sku_daily_factory["날짜_date"] != today)
                    ][["날짜_date", "공장", "규격대응률(%)"]].copy()
                    factory_daily_spec.rename(columns={"규격대응률(%)": "규격 대응률(%)"}, inplace=True)
                elif shortage_prod_daily is not None and len(shortage_prod_daily) > 0:
                    factory_daily_spec_by_day = shortage_prod_daily[["날짜_date", "규격대응률(%)"]].copy()
                    factory_daily_spec_by_day.rename(columns={"날짜_date": "날짜", "규격대응률(%)": "규격 대응률(%)"}, inplace=True)

                factory_need_label = f"{KPI_LABEL_MAP['총부족수량']} (pcs)"
                factory_daily_display = factory_daily.rename(columns={
                    "생산일자_date": "날짜",
                    "총실적": f"{KPI_LABEL_MAP['총실적']} (pcs)",
                    "총부족수량": factory_need_label,
                    "유효생산량": f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                    "과생산량": f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                    "불필요생산량": f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
                }).copy()

                if factory_daily_spec is not None and len(factory_daily_spec) > 0:
                    factory_daily_display = factory_daily_display.merge(
                        factory_daily_spec.rename(columns={"날짜_date": "날짜"}),
                        on=["날짜", "공장"],
                        how="left",
                    )
                elif factory_daily_spec_by_day is not None and len(factory_daily_spec_by_day) > 0:
                    st.warning(
                        "관별(공장별) `규격 대응률(SKU 기준)` 계산 불가: 선택 기간 데이터에 `공장` 값이 비어있거나 SKU 집계에 필요한 컬럼이 누락되었습니다. "
                        "(전사 일자 규격 대응률을 동일 적용해 표시)"
                    )
                    factory_daily_display = factory_daily_display.merge(
                        factory_daily_spec_by_day,
                        on=["날짜"],
                        how="left",
                    )

                if "규격 대응률(%)" in factory_daily_display.columns:
                    factory_daily_display["규격 대응률(%)"] = (
                        pd.to_numeric(factory_daily_display["규격 대응률(%)"], errors="coerce")
                        .replace([np.inf, -np.inf], 0)
                    )
                else:
                    factory_daily_display["규격 대응률(%)"] = np.nan

                factory_daily_display["날짜"] = pd.to_datetime(factory_daily_display["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")

                # 공장 순서 지정 (A관 > C관 > S관)
                factory_order = {"A관(1공장)": 1, "C관(2공장)": 2, "S관(3공장)": 3}
                factory_daily_display["_factory_sort"] = factory_daily_display["공장"].map(factory_order)
                factory_daily_display = factory_daily_display.sort_values(["날짜", "_factory_sort", "공장"]).drop(columns=["_factory_sort"]).reset_index(drop=True)

                # 컬럼 순서 정리 (비중(%) 우선, 수량(pcs)은 뒤쪽)
                factory_daily_cols = [
                    "날짜",
                    "공장",
                    f"{KPI_LABEL_MAP['총실적']} (pcs)",
                    factory_need_label,
                    "규격 대응률(%)",
                ]
                factory_daily_cols.extend(
                    [
                        RATE_LABEL_MAP["유효비율(%)"],
                        RATE_LABEL_MAP["과생산비율(%)"],
                        RATE_LABEL_MAP["불필요비율(%)"],
                        f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                        f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                        f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
                    ]
                )
                factory_daily_display = factory_daily_display[factory_daily_cols].copy()

                _safe_dataframe(
                    factory_daily_display,
                    fmt={
                        f"{KPI_LABEL_MAP['총실적']} (pcs)": "{:,.0f}",
                        factory_need_label: "{:,.0f}",
                        f"{KPI_LABEL_MAP['유효생산량']} (pcs)": "{:,.0f}",
                        f"{KPI_LABEL_MAP['과생산량']} (pcs)": "{:,.0f}",
                        f"{KPI_LABEL_MAP['불필요생산량']} (pcs)": "{:,.0f}",
                        **({"규격 대응률(%)": "{:.1f}%"} if "규격 대응률(%)" in factory_daily_display.columns else {}),
                        RATE_LABEL_MAP["유효비율(%)"]: "{:.1f}%",
                        RATE_LABEL_MAP["과생산비율(%)"]: "{:.1f}%",
                        RATE_LABEL_MAP["불필요비율(%)"]: "{:.1f}%",
                    },
                )

        # ============== 자료 다운로드 ==============
        st.markdown("### 📥 자료 다운로드")
        if len(daily_summary_filtered) == 0:
            st.info("선택한 기간에 다운로드할 데이터가 없습니다.")
        else:
            metric_order = ["규격 대응률", "정확 대응 비중", "초과 생산 비중", "비정형 생산 비중"]
            metric_sheet_map = {
                "규격 대응률": "규격대응률",
                "정확 대응 비중": "정확대응비중",
                "초과 생산 비중": "초과생산비중",
                "비정형 생산 비중": "비정형생산비중",
            }
            metric_desc = {
                "규격 대응률": "생산한 SKU(제품코드) 중 필요가 있었던 SKU 비중",
                "정확 대응 비중": "총 생산량 중 정확 대응 생산량이 차지하는 비중",
                "초과 생산 비중": "총 생산량 중 초과 생산량이 차지하는 비중",
                "비정형 생산 비중": "총 생산량 중 비정형 생산량이 차지하는 비중",
            }

            rate_col_map = {
                "규격 대응률": "규격 대응률(%)",
                "정확 대응 비중": RATE_LABEL_MAP["유효비율(%)"],
                "초과 생산 비중": RATE_LABEL_MAP["과생산비율(%)"],
                "비정형 생산 비중": RATE_LABEL_MAP["불필요비율(%)"],
            }
            pcs_col_map = {
                "규격 대응률": None,
                "정확 대응 비중": f"{KPI_LABEL_MAP['유효생산량']} (pcs)",
                "초과 생산 비중": f"{KPI_LABEL_MAP['과생산량']} (pcs)",
                "비정형 생산 비중": f"{KPI_LABEL_MAP['불필요생산량']} (pcs)",
            }

            def _daily_table_for(metric: str) -> pd.DataFrame:
                cols = [
                    "날짜",
                    f"{KPI_LABEL_MAP['총실적']} (pcs)",
                    f"{KPI_LABEL_MAP['총부족수량']} (pcs)",
                    rate_col_map.get(metric),
                    pcs_col_map.get(metric),
                ]
                cols = [c for c in cols if c and c in daily_display.columns]
                return daily_display[cols].copy() if cols else pd.DataFrame()

            def _factory_daily_table_for(metric: str) -> pd.DataFrame:
                cols = [
                    "날짜",
                    "공장",
                    f"{KPI_LABEL_MAP['총실적']} (pcs)",
                    factory_need_label,
                    rate_col_map.get(metric),
                    pcs_col_map.get(metric),
                ]
                cols = [c for c in cols if c and c in factory_daily_display.columns]
                return factory_daily_display[cols].copy() if cols else pd.DataFrame()

            can_export = "factory_data" in locals() and isinstance(factory_data, pd.DataFrame) and len(factory_data) > 0
            if not can_export:
                st.info("공장별 데이터가 없어 다운로드 파일을 생성할 수 없습니다.")
            else:
                excel_key = "export_excel_bytes"
                excel_sig_key = "export_excel_signature"
                signature = (
                    str(start_date),
                    str(end_date),
                    int(daily_summary_filtered["날짜_date"].nunique()) if "날짜_date" in daily_summary_filtered.columns else 0,
                    int(factory_summary_filtered["공장"].nunique()) if "공장" in factory_summary_filtered.columns else 0,
                    int(factory_summary_filtered["생산일자_date"].nunique()) if "생산일자_date" in factory_summary_filtered.columns else 0,
                )

                needs_rebuild = (
                    excel_key not in st.session_state
                    or excel_sig_key not in st.session_state
                    or st.session_state.get(excel_sig_key) != signature
                )

                if needs_rebuild:
                    with st.spinner("다운로드 파일 준비 중... (차트 PNG 포함)"):
                        try:
                            export_payload: dict[str, dict[str, object]] = {}
                            for metric in metric_order:
                                factory_table, bar_fig = _build_factory_bar_fig(factory_data=factory_data, metric_option=metric)
                                line_fig = _build_factory_line_fig(
                                    metric_option=metric,
                                    factory_summary_filtered=factory_summary_filtered,
                                    sku_daily_factory=sku_daily_factory,
                                    sku_daily_all=sku_daily_all,
                                    start_date=start_date,
                                    end_date=end_date,
                                    today=today,
                                )
                                line_ts_df = _build_factory_line_ts_df(
                                    metric_option=metric,
                                    factory_summary_filtered=factory_summary_filtered,
                                    sku_daily_factory=sku_daily_factory,
                                    sku_daily_all=sku_daily_all,
                                    start_date=start_date,
                                    end_date=end_date,
                                    today=today,
                                )
                                export_payload[metric] = {
                                    "factory_table": factory_table,
                                    "daily_table": _daily_table_for(metric),
                                    "factory_daily_table": _factory_daily_table_for(metric) if "factory_daily_display" in locals() else pd.DataFrame(),
                                    "bar_fig": bar_fig,
                                    "line_fig": line_fig,
                                    "line_ts_df": line_ts_df,
                                    "kpi_total_prod": total_prod,
                                    "kpi_spec_rate": shortage_prod_rate,
                                    "kpi_valid": (valid_rate, valid_prod),
                                    "kpi_over": (over_rate, over_prod),
                                    "kpi_waste": (waste_rate, waste_prod),
                                    "filter_option": filter_option,
                                }

                            start_date_str = start_date.strftime("%Y-%m-%d")
                            end_date_str = end_date.strftime("%Y-%m-%d")
                            st.session_state[excel_key] = _build_excel_report_bytes(
                                metric_order=metric_order,
                                metric_sheet_map=metric_sheet_map,
                                metric_desc=metric_desc,
                                export_payload=export_payload,
                                start_date_str=start_date_str,
                                end_date_str=end_date_str,
                                tz_name="Asia/Seoul",
                            )
                            st.session_state[excel_sig_key] = signature
                        except ModuleNotFoundError as e:
                            st.session_state[excel_key] = None
                            st.error(f"다운로드 파일 생성에 필요한 패키지가 없습니다: {e}. `pip install -r requirements.txt` 후 재실행/재시도해주세요.")
                        except Exception as e:
                            st.session_state[excel_key] = None
                            st.error(f"다운로드 파일 생성 실패: {e}")

                if st.session_state.get(excel_key):
                    start_date_str = start_date.strftime("%Y%m%d")
                    end_date_str = end_date.strftime("%Y%m%d")
                    filename = f"공장비교_리포트_{start_date_str}_{end_date_str}.xlsx"
                    st.download_button(
                        "공장비교 리포트 다운로드",
                        data=st.session_state[excel_key],
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                    )

                # RAWDATA download (original result excel)
                st.markdown("")
                raw_candidates = result_candidates if "result_candidates" in globals() else []
                if raw_candidates:
                    raw_path = raw_candidates[0]  # newest (already sorted desc by mtime)
                    try:
                        raw_bytes = Path(raw_path).read_bytes()
                        st.download_button(
                            "ROWDATA 다운로드",
                            data=raw_bytes,
                            file_name=Path(raw_path).name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    except Exception as e:
                        st.error(f"ROWDATA 다운로드 준비 실패: {e}")

    if main_tab == "공정 밸런스":
        st.markdown("### ⚖️ 공정 밸런스")
        if (process_proc_base is None) or (len(process_proc_base) == 0):
            st.info("공정 밸런스 집계 데이터가 없습니다. (result2_proc_base) 전처리 결과를 확인해주세요.")
        else:
            # 미리 집계된 proc_base/det_base를 기간 필터만 적용(리런 속도↑)
            proc_base = process_proc_base
            det_base = process_det_base
            proc = proc_base[(proc_base["날짜_date"] >= start_date) & (proc_base["날짜_date"] <= end_date) & (proc_base["날짜_date"] != today)].copy()
            det = det_base[(det_base["날짜_date"] >= start_date) & (det_base["날짜_date"] <= end_date) & (det_base["날짜_date"] != today)].copy() if len(det_base) else pd.DataFrame()

            if len(proc) == 0:
                st.info("선택한 기간에 공정별 데이터가 없습니다.")
            else:
                target_order = ["사출", "분리", "하드레이션", "접착", "누수규격"]
                factory_order = ["A관", "C관", "S관"]

                proc["공장그룹"] = pd.Categorical(proc["공장그룹"], categories=factory_order + ["기타"], ordered=True)

                # ---- 공정 점수 산출(부족수량 기반 필요수량 사용 X) ----
                # 1) 규격 대응 SKU 점수화: (필요 SKU ∩ 생산 SKU) / 생산 SKU
                # 2) 정확대응 비중: 유효생산량 / 실적수량
                # 3) 초과생산 비중: 과생산량 / 실적수량
                # 4) 비정형 생산 비중: 불필요생산량 / 실적수량
                # 점수 합성(가중): SKU 25%, 정확 45%, 초과감점 20%, 비정형감점 10%
                # proc는 이미 일자/공장/공정 단위로 수량 및 SKU 카운트까지 집계됨

                proc["규격대응률(%)"] = np.where(
                    proc.get("생산SKU수", 0) > 0,
                    proc.get("규격대응SKU수", 0) / proc.get("생산SKU수", 0) * 100,
                    0.0,
                )
                proc["정확대응비중(%)"] = np.where(
                    proc.get("실적수량", 0) > 0,
                    proc.get("유효생산량", 0) / proc.get("실적수량", 0) * 100,
                    0.0,
                )
                proc["초과생산비중(%)"] = np.where(
                    proc.get("실적수량", 0) > 0,
                    proc.get("과생산량", 0) / proc.get("실적수량", 0) * 100,
                    0.0,
                )
                proc["비정형생산비중(%)"] = np.where(
                    proc.get("실적수량", 0) > 0,
                    proc.get("불필요생산량", 0) / proc.get("실적수량", 0) * 100,
                    0.0,
                )

                proc["규격대응률(%)"] = pd.to_numeric(proc["규격대응률(%)"], errors="coerce").fillna(0).clip(0, 100)
                proc["정확대응비중(%)"] = pd.to_numeric(proc["정확대응비중(%)"], errors="coerce").fillna(0).clip(0, 100)
                proc["초과생산비중(%)"] = pd.to_numeric(proc["초과생산비중(%)"], errors="coerce").fillna(0).clip(0, 300)
                proc["비정형생산비중(%)"] = pd.to_numeric(proc["비정형생산비중(%)"], errors="coerce").fillna(0).clip(0, 300)

                # 공정점수(0~100)
                # 목표: "필요 대응(규격대응률)"을 가장 중요하게, "비정형" 감점을 더 강하게 반영
                #  - 규격대응률(%)      : 필요가 있었던 SKU를 생산했는가(했냐/안했냐 성격)
                #  - 정확대응비중(%)    : 수량 관점의 대응 정도
                #  - 초과/비정형비중(%) : 불필요 생산에 대한 감점 (비정형 가중↑)
                proc["공정점수_raw"] = (
                    proc["규격대응률(%)"] * 0.45
                    + proc["정확대응비중(%)"] * 0.25
                    + (100 - proc["초과생산비중(%)"].clip(0, 100)) * 0.10
                    + (100 - proc["비정형생산비중(%)"].clip(0, 100)) * 0.20
                ).clip(0, 100)
                # 규격대응률이 낮으면(필요 대응 자체가 안 됨) 점수 상한을 낮춰 "했냐/안했냐"를 강조
                cap = np.select(
                    [
                        proc["규격대응률(%)"] >= 85,
                        proc["규격대응률(%)"] >= 70,
                        proc["규격대응률(%)"] >= 55,
                    ],
                    [100.0, 75.0, 65.0],
                    default=55.0,
                )
                proc["공정점수"] = np.minimum(proc["공정점수_raw"], cap).clip(0, 100)

                # 등급(공정별 기준)
                proc["상태"] = np.select(
                    [
                        proc["공정점수"] >= 70,
                        proc["공정점수"] >= 65,
                        proc["공정점수"] >= 60,
                    ],
                    ["양호", "주의", "경고"],
                    default="위험",
                )

                # 집계(가중: 실적수량)
                w = pd.to_numeric(proc.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0)
                overall = float((proc["공정점수"] * w).sum() / w.sum()) if float(w.sum()) > 0 else float(proc["공정점수"].mean())
                by_proc = (
                    proc.groupby("공정", dropna=False)
                    .apply(lambda g: (g["공정점수"] * pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0)).sum() / max(float(pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0).sum()), 1.0))
                    .rename("평균점수")
                    .reset_index()
                )
                by_proc["평균점수"] = pd.to_numeric(by_proc["평균점수"], errors="coerce").fillna(0)
                by_proc["공정"] = pd.Categorical(by_proc["공정"], categories=target_order, ordered=True)
                by_proc = by_proc.sort_values("공정")
                worst_proc = by_proc.sort_values("평균점수").iloc[0]["공정"] if len(by_proc) else "-"
                risk_count = int((by_proc["평균점수"] < 70).sum()) if len(by_proc) else 0

                proc_score_map = {str(r["공정"]): float(r["평균점수"]) for _, r in by_proc.iterrows()} if len(by_proc) else {}

                proc_grades = [grade_of(float(proc_score_map.get(p, 0.0))) for p in target_order]
                overall_status = majority_grade(proc_grades)
                overall_status_html = f"<span style='color:{grade_text_color(overall_status)}'>{overall_status}</span>"

                k_cols = st.columns([1.25, 1, 1, 1, 1, 1])
                with k_cols[0]:
                    render_kpi_card(
                        "공정 밸런스 종합점수",
                        f"<span style='color:#1d4ed8'>{overall:.1f}점</span>",
                        sub=f"등급: {overall_status_html}",
                    )
                with k_cols[1]:
                    v = float(proc_score_map.get("사출", 0.0))
                    render_kpi_card("사출 점수", f"{v:.1f}점", sub=f"등급: {grade_of(v)}")
                with k_cols[2]:
                    v = float(proc_score_map.get("분리", 0.0))
                    render_kpi_card("분리 점수", f"{v:.1f}점", sub=f"등급: {grade_of(v)}")
                with k_cols[3]:
                    v = float(proc_score_map.get("하드레이션", 0.0))
                    render_kpi_card("하드레이션 점수", f"{v:.1f}점", sub=f"등급: {grade_of(v)}")
                with k_cols[4]:
                    v = float(proc_score_map.get("접착", 0.0))
                    render_kpi_card("접착 점수", f"{v:.1f}점", sub=f"등급: {grade_of(v)}")
                with k_cols[5]:
                    v = float(proc_score_map.get("누수규격", 0.0))
                    render_kpi_card("누수규격 점수", f"{v:.1f}점", sub=f"등급: {grade_of(v)}")

                st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

                with st.expander("지표 정의/상세 보기", expanded=False):
                    st.markdown(
                        "- `규격대응률(%)` : 일자/공장/공정별 `(필요 SKU ∩ 생산 SKU) ÷ 생산 SKU` 의 비율\n"
                        "- `정확대응비중(%)` : `유효생산량 ÷ 실적수량`\n"
                        "- `초과생산비중(%)` : `과생산량 ÷ 실적수량`\n"
                        "- `비정형생산비중(%)` : `불필요생산량 ÷ 실적수량`\n"
                        "- `공정점수(0~100)` : `0.45×규격대응률 + 0.25×정확대응비중 + 0.10×(100-초과생산비중) + 0.20×(100-비정형생산비중)` (규격대응률이 낮으면 상한 적용)\n"
                        "- `등급(공정)` : 70↑ 양호 / 65↑ 주의 / 60↑ 경고 / 60↓ 위험\n"
                        "- `등급(공장 종합)` : 공정 5개 등급의 다수결(동률이면 낮은 등급, 위험 1개면 최대 경고)\n"
                        "- `종합점수(공장/전체)` : 각 `공정점수`의 `실적수량` 가중평균"
                    )
                    st.caption("참고: 공정 밸런스는 `유효생산량_결과2.xlsx`의 `매칭결과` 시트를 기반으로 계산합니다.")

                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

                proc_acs = proc[proc["공장그룹"].isin(factory_order)].copy()

                # 공장별(좌) + 공정별(우) 한 행 배치
                c_left, c_right = st.columns([1, 2], gap="large")

                # 공정 팔레트(레퍼런스 톤): 인디고/시안/민트/앰버/핑크레드
                proc_color_map = {
                    "사출": "#6366F1",
                    "분리": "#22D3EE",
                    "하드레이션": "#34D399",
                    "접착": "#FBBF24",
                    "누수규격": "#F43F5E",
                }
                factory_display = {"A관": "A관(1공장)", "C관": "C관(2공장)", "S관": "S관(3공장)"}
                chart_height = 520

                with c_left:
                    st.markdown("#### 공장별 종합 점수")
                    if len(proc_acs) == 0:
                        st.info("선택한 기간에 A/C/S관 데이터가 없습니다.")
                    else:
                        by_factory = (
                            proc_acs.groupby("공장그룹", dropna=False)
                            .apply(
                                lambda g: (g["공정점수"] * pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0)).sum()
                                / max(float(pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0).sum()), 1.0)
                            )
                            .rename("종합점수")
                            .reset_index()
                        )
                        by_factory["공장그룹"] = pd.Categorical(by_factory["공장그룹"], categories=factory_order, ordered=True)
                        by_factory = by_factory.sort_values("공장그룹")
                        by_factory["공장"] = by_factory["공장그룹"].astype(str).map(factory_display)

                        # 공장별 종합 등급(공정 5개 등급의 다수결)
                        by_fac_proc_scores = (
                            proc_acs.groupby(["공장그룹", "공정"], dropna=False)
                            .apply(
                                lambda g: (g["공정점수"] * pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0)).sum()
                                / max(float(pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0).sum()), 1.0)
                            )
                            .rename("평균점수")
                            .reset_index()
                        )
                        by_fac_proc_scores["평균점수"] = pd.to_numeric(by_fac_proc_scores["평균점수"], errors="coerce").fillna(0)
                        fac_grade_map: dict[str, str] = {}
                        for fg in factory_order:
                            subp = by_fac_proc_scores[by_fac_proc_scores["공장그룹"].astype(str) == str(fg)]
                            grades = [grade_of(float(subp[subp["공정"].astype(str) == p]["평균점수"].mean())) for p in target_order]
                            fac_grade_map[str(fg)] = majority_grade(grades)
                        by_factory["등급"] = by_factory["공장그룹"].astype(str).map(fac_grade_map).fillna("위험")
                        by_factory["표시"] = by_factory["종합점수"].apply(lambda v: f"{float(v):.1f}")

                        fig_factory = px.bar(
                            by_factory,
                            x="공장",
                            y="종합점수",
                            range_y=[0, 100],
                            category_orders={"공장": [factory_display[k] for k in factory_order]},
                            color="공장",
                            text="표시",
                            color_discrete_map={
                                factory_display["A관"]: "#6366F1",
                                factory_display["C관"]: "#8B5CF6",
                                factory_display["S관"]: "#EC4899",
                            },
                        )
                        fig_factory.update_traces(
                            texttemplate="%{text}",
                            textposition="outside",
                            textfont=dict(size=32, family="Arial", color="#111827"),
                            marker=dict(cornerradius=18),
                            hovertemplate="공장=%{x}<br>종합점수=%{y:.1f}<extra></extra>",
                            cliponaxis=False,
                        )

                        fig_factory.update_layout(
                            height=chart_height,
                            margin=dict(l=95, r=10, t=10, b=10),
                            showlegend=False,
                            xaxis_title="공장",
                            yaxis_title="종합 점수",
                            xaxis=dict(
                                tickfont=dict(size=18, family="Arial", color="#111827"),
                                title_font=dict(size=18, family="Arial", color="#111827"),
                            ),
                            yaxis=dict(
                                range=[0, 110],
                                tickfont=dict(size=14, family="Arial", color="#111827"),
                                title_font=dict(size=18, family="Arial", color="#111827"),
                                automargin=True,
                            ),
                            uniformtext_minsize=10,
                            uniformtext_mode="hide",
                        )
                        st.plotly_chart(fig_factory, use_container_width=True)

                with c_right:
                    st.markdown("#### 공장별 공정 평균 점수")
                    if len(proc_acs) == 0:
                        st.info("선택한 기간에 A/C/S관 데이터가 없습니다.")
                    else:
                        chart_col, legend_col = st.columns([6.3, 0.9], gap="large")
                        by_fac_proc = (
                            proc_acs.groupby(["공장그룹", "공정"], dropna=False)
                            .apply(
                                lambda g: (g["공정점수"] * pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0)).sum()
                                / max(float(pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0).sum()), 1.0)
                            )
                            .rename("평균점수")
                            .reset_index()
                        )
                        by_fac_proc["평균점수"] = pd.to_numeric(by_fac_proc["평균점수"], errors="coerce").fillna(0)
                        by_fac_proc["공정"] = pd.Categorical(by_fac_proc["공정"], categories=target_order, ordered=True)
                        by_fac_proc = by_fac_proc.sort_values(["공장그룹", "공정"])
                        by_fac_proc["등급"] = by_fac_proc["평균점수"].apply(lambda v: grade_of(float(v)))
                        by_fac_proc["표시"] = by_fac_proc["평균점수"].apply(lambda v: f"{float(v):.1f}")

                        fig_fac = px.bar(
                            by_fac_proc,
                            x="공정",
                            y="평균점수",
                            color="공정",
                            facet_row="공장그룹",
                            facet_row_spacing=0.10,
                            category_orders={"공장그룹": factory_order, "공정": target_order},
                            range_y=[0, 100],
                            text="표시",
                            color_discrete_map=proc_color_map,
                            custom_data=["등급"],
                        )
                        fig_fac.update_traces(
                            marker=dict(cornerradius=14),
                            texttemplate="%{text}",
                            textposition="outside",
                            textfont=dict(size=22, family="Arial", color="#111827"),
                            cliponaxis=False,
                            hovertemplate="공정=%{x}<br>점수=%{y:.1f}<br>등급=%{customdata[0]}<extra></extra>",
                        )
                        fig_fac.update_yaxes(range=[0, 110])
                        fig_fac.update_layout(
                            height=chart_height,
                            margin=dict(l=230, r=10, t=30, b=10),
                            showlegend=False,
                            xaxis_title="공정",
                            yaxis_title=None,
                            legend_title_text="공정",
                            uniformtext_minsize=18,
                            uniformtext_mode="hide",
                        )
                        # facet 라벨을 왼쪽에 보이도록 위치/텍스트 보정
                        for ann in fig_fac.layout.annotations:
                            if isinstance(ann.text, str) and "공장그룹=" in ann.text:
                                raw = ann.text.replace("공장그룹=", "")
                                ann.text = factory_display.get(raw, raw)
                                ann.x = -0.18
                                ann.xanchor = "left"
                                ann.yanchor = "middle"
                                ann.textangle = 0
                                ann.font = dict(size=16, family="Arial", color="#111827")
                        # y축 제목은 제거하고(기울어진 텍스트), tick/격자만 유지
                        fig_fac.update_yaxes(title_text="", tickfont=dict(size=14, family="Arial", color="#111827"))
                        fig_fac.update_xaxes(tickfont=dict(size=18, family="Arial", color="#111827"), title_font=dict(size=18, family="Arial", color="#111827"))
                        # 구분선(패널 사이)
                        fig_fac.add_shape(type="line", xref="paper", yref="paper", x0=0, x1=1, y0=2/3, y1=2/3, line=dict(color="#E5E7EB", width=2))
                        fig_fac.add_shape(type="line", xref="paper", yref="paper", x0=0, x1=1, y0=1/3, y1=1/3, line=dict(color="#E5E7EB", width=2))
                        with chart_col:
                            st.plotly_chart(fig_fac, use_container_width=True)
                        with legend_col:
                            legend_items = "".join(
                                [
                                    f"<div style='display:flex; align-items:center; gap:8px;'><span style='width:10px; height:10px; border-radius:3px; background:{proc_color_map.get(p, '#64748B')}; display:inline-block;'></span><b>{p}</b></div>"
                                    for p in target_order
                                ]
                            )
                            st.markdown(
                                "<div style='padding:10px 10px; border:1px solid #E5E7EB; border-radius:12px; background:#F9FAFB;'>"
                                "<div style='font-weight:800; font-size:13px; color:#111827; margin-bottom:8px;'>범례(공정)</div>"
                                f"<div style='display:flex; flex-direction:column; gap:8px; font-size:13px;'>{legend_items}</div>"
                                "</div>",
                                unsafe_allow_html=True,
                            )

                # 공장별 종합점수 추이(선그래프): 기본 접힘(토글)로 제공
                with st.expander("공장별 종합점수 추이", expanded=False):
                    req_cols_ts = {"날짜_date", "공장그룹", "공정점수", "실적수량"}
                    if len(proc_acs) == 0 or not req_cols_ts.issubset(set(proc_acs.columns)):
                        st.info("추이 그래프를 그리기 위한 데이터가 부족합니다.")
                    else:
                        display_start_date = start_date
                        display_end_date = end_date
                        if filter_option == "당월":
                            # 당월은 1일~말일까지 축을 만들어 추이를 한 번에 보기 좋게 표시
                            display_end_date = _month_end(display_start_date)

                        if filter_option in {"당월", "전월"}:
                            bucket = "D"
                        else:
                            span_days = (display_end_date - display_start_date).days + 1
                            if span_days <= 30:
                                bucket = "D"
                            elif span_days <= 210:
                                bucket = "W"
                            else:
                                bucket = "M"

                        axis = _build_axis(display_start_date, display_end_date, bucket)
                        tickvals, ticktext = _build_tick_labels(axis, bucket)

                        ts_base = proc_acs[["날짜_date", "공장그룹", "공정점수", "실적수량"]].copy()
                        ts_base["date"] = pd.to_datetime(ts_base["날짜_date"], errors="coerce")
                        ts_base = ts_base.dropna(subset=["date"])
                        ts_base["period"] = _period_start(ts_base["date"], bucket)
                        ts_base["w"] = pd.to_numeric(ts_base["실적수량"], errors="coerce").fillna(0).clip(lower=0)
                        ts_base["s"] = pd.to_numeric(ts_base["공정점수"], errors="coerce").fillna(0).clip(0, 100)

                        agg_ts = (
                            ts_base.groupby(["period", "공장그룹"], dropna=False)
                            .apply(lambda g: float((g["s"] * g["w"]).sum() / max(float(g["w"].sum()), 1.0)))
                            .rename("종합점수")
                            .reset_index()
                        )
                        agg_ts["period"] = pd.to_datetime(agg_ts["period"], errors="coerce")

                        factories = [f for f in factory_order if f in proc_acs["공장그룹"].astype(str).unique().tolist()]
                        if not factories:
                            factories = factory_order

                        full_grid = pd.MultiIndex.from_product([axis, factories], names=["period", "공장그룹"]).to_frame(index=False)
                        ts_df = full_grid.merge(agg_ts, on=["period", "공장그룹"], how="left")
                        ts_df["공장"] = ts_df["공장그룹"].astype(str).map(factory_display).fillna(ts_df["공장그룹"].astype(str))

                        label_map = {pd.Timestamp(v): t for v, t in zip(tickvals, ticktext, strict=False)}
                        ts_df["x_label"] = ts_df["period"].map(label_map)

                        line_fig = px.line(
                            ts_df,
                            x="period",
                            y="종합점수",
                            color="공장",
                            markers=False,
                            custom_data=["x_label"],
                            color_discrete_map={
                                factory_display["A관"]: FACTORY_COLOR_MAP["A관"],
                                factory_display["C관"]: FACTORY_COLOR_MAP["C관"],
                                factory_display["S관"]: FACTORY_COLOR_MAP["S관"],
                            },
                        )
                        line_fig.update_traces(
                            line=dict(width=3.5),
                            hovertemplate="공장=%{legendgroup}<br>기간=%{customdata[0]}<br>종합점수=%{y:.1f}점<extra></extra>",
                        )
                        line_fig.update_layout(
                            height=330,
                            margin=dict(l=0, r=0, t=10, b=0),
                            yaxis=dict(range=[0, 105], title="종합점수", tickformat=".1f"),
                            xaxis=dict(
                                tickmode="array",
                                tickvals=tickvals,
                                ticktext=ticktext,
                                tickangle=-45,
                                tickfont=dict(size=10),
                                title=None,
                            ),
                            legend_title_text="공장",
                            showlegend=True,
                        )
                        st.plotly_chart(line_fig, use_container_width=True)

                st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
                st.markdown("#### 관별 감점요인 · 공정별 생산 비율")
                if len(proc_acs) == 0:
                    st.info("선택한 기간에 A/C/S관 데이터가 없습니다.")
                else:
                    from plotly.subplots import make_subplots

                    qty_cols = [c for c in ["실적수량", "유효생산량", "과생산량", "불필요생산량"] if c in proc_acs.columns]
                    comp = proc_acs.groupby("공장그룹", dropna=False)[qty_cols].sum().reset_index() if qty_cols else pd.DataFrame()
                    for c in ["실적수량", "유효생산량", "과생산량", "불필요생산량"]:
                        if c in comp.columns:
                            comp[c] = pd.to_numeric(comp[c], errors="coerce").fillna(0)
                        else:
                            comp[c] = 0.0

                    comp["초과(%)"] = np.where(comp["실적수량"] > 0, comp["과생산량"] / comp["실적수량"] * 100, 0.0)
                    comp["비정형(%)"] = np.where(comp["실적수량"] > 0, comp["불필요생산량"] / comp["실적수량"] * 100, 0.0)
                    for c in ["초과(%)", "비정형(%)"]:
                        comp[c] = pd.to_numeric(comp[c], errors="coerce").fillna(0).clip(0, 100)

                    # 관별 종합점수(공정점수의 실적 가중평균) -> 감점(=100-종합점수) 계산용
                    by_factory_score = (
                        proc_acs.groupby("공장그룹", dropna=False)
                        .apply(
                            lambda g: (g["공정점수"] * pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0)).sum()
                            / max(float(pd.to_numeric(g.get("실적수량", 0), errors="coerce").fillna(0).clip(lower=0).sum()), 1.0)
                        )
                        .rename("종합점수")
                        .reset_index()
                    )
                    by_factory_score["종합점수"] = pd.to_numeric(by_factory_score["종합점수"], errors="coerce").fillna(0).clip(0, 100)
                    score_map = {str(r["공장그룹"]): float(r["종합점수"]) for _, r in by_factory_score.iterrows()}

                    comp_map = {str(r["공장그룹"]): r for _, r in comp.iterrows()}
                    donut_cols = st.columns([1, 1, 1, 0.55], gap="large")
                    donut_colors = {"초과": BALANCE_COLORS["초과"], "비정형": BALANCE_COLORS["비정형"]}

                    with donut_cols[3]:
                        donut_legend_html = f"""
                        <div style='padding:12px 12px; border:1px solid #E5E7EB; border-radius:12px; background:#F9FAFB;'>
                          <div style='font-weight:800; color:#111827; margin-bottom:10px;'>범례(도넛)</div>
                          <div style='color:#6B7280; font-size:12px; margin:-4px 0 10px 0;'>감점요인(초과+비정형) 내 구성비</div>
                          <div style='display:flex; flex-direction:column; gap:10px;'>
                            <div style='display:flex; align-items:center; gap:10px;'><span style='width:12px; height:12px; border-radius:3px; background:{BALANCE_COLORS['초과']}; display:inline-block;'></span><b>초과</b></div>
                            <div style='display:flex; align-items:center; gap:10px;'><span style='width:12px; height:12px; border-radius:3px; background:{BALANCE_COLORS['비정형']}; display:inline-block;'></span><b>비정형</b></div>
                          </div>
                        </div>
                        """
                        st.markdown(donut_legend_html, unsafe_allow_html=True)

                    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
                    for idx, g in enumerate(factory_order):
                        row = comp_map.get(str(g))
                        if row is None:
                            with donut_cols[idx]:
                                st.caption(factory_display.get(g, str(g)))
                                st.info("데이터 없음")
                            continue

                        over = float(row.get("초과(%)", 0.0))
                        waste = float(row.get("비정형(%)", 0.0))
                        # 감점(전체)은 100 - 종합점수로 표시 (점수의 부족분)
                        score = float(score_map.get(str(g), 0.0))
                        penalty_total = max(0.0, min(100.0, 100.0 - score))

                        # 도넛은 "초과 vs 비정형" 비중만 표현 (둘 합이 0이면 0/0 방지)
                        denom = over + waste
                        over_share = (over / denom * 100.0) if denom > 0 else 0.0
                        waste_share = (waste / denom * 100.0) if denom > 0 else 0.0
                        over_share = float(np.clip(over_share, 0, 100))
                        waste_share = float(np.clip(waste_share, 0, 100))

                        fig_donut = go.Figure(
                            data=[
                                go.Pie(
                                    labels=["초과", "비정형"],
                                    values=[over_share, waste_share],
                                    hole=0.56,
                                    sort=False,
                                    direction="clockwise",
                                    marker=dict(colors=[donut_colors["초과"], donut_colors["비정형"]]),
                                    textinfo="none",
                                    texttemplate="%{percent:.1%}",
                                    textposition="outside",
                                    outsidetextfont=dict(size=24, family="Arial Black", color="#0B1220"),
                                    hovertemplate="%{label}<br>비중=%{percent}<br>실적대비=%{customdata:.1f}%<extra></extra>",
                                    customdata=[over, waste],
                                )
                            ]
                        )
                        fig_donut.update_layout(
                            height=320,
                            margin=dict(l=10, r=10, t=50, b=10),
                            showlegend=False,
                            title=dict(text=factory_display.get(g, str(g)), x=0.5, xanchor="center", font=dict(size=20, family="Arial", color="#111827")),
                            uniformtext_minsize=18,
                            uniformtext_mode="show",
                            annotations=[
                                dict(
                                    text=f"<span style='font-size:22px; font-weight:900; opacity:0.95;'>감점</span><br><span style='font-size:30px; font-weight:900;'>{penalty_total:.1f}</span>",
                                    x=0.5,
                                    y=0.5,
                                    font=dict(size=30, family="Arial Black", color="#B91C1C"),
                                    showarrow=False,
                                )
                            ],
                        )

                        with donut_cols[idx]:
                            st.plotly_chart(fig_donut, use_container_width=True)

                    # 공정별은 "정확대응 vs (초과+비정형)" 100% 누적 가로 막대로 표현(비중)
                    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
                    # 공정별 비교는 '정확대응비중' vs '(초과생산비중+비정형생산비중)'을 100%로 정규화해서 표현
                    proc_comp = (
                        proc_acs.groupby(["공장그룹", "공정"], dropna=False)[
                            [c for c in ["실적수량", "유효생산량", "과생산량", "불필요생산량"] if c in proc_acs.columns]
                        ]
                        .sum()
                        .reset_index()
                        if len(proc_acs) > 0 else pd.DataFrame(columns=["공장그룹", "공정"])
                    )
                    for c in ["실적수량", "유효생산량", "과생산량", "불필요생산량"]:
                        if c in proc_comp.columns:
                            proc_comp[c] = pd.to_numeric(proc_comp[c], errors="coerce").fillna(0)
                        else:
                            proc_comp[c] = 0.0

                    proc_comp["정확대응비중(%)"] = np.where(
                        proc_comp["실적수량"] > 0,
                        proc_comp["유효생산량"] / proc_comp["실적수량"] * 100,
                        0.0,
                    )
                    proc_comp["초과+비정형비중(%)"] = np.where(
                        proc_comp["실적수량"] > 0,
                        (proc_comp["과생산량"] + proc_comp["불필요생산량"]) / proc_comp["실적수량"] * 100,
                        0.0,
                    )
                    proc_comp["정확대응비중(%)"] = pd.to_numeric(proc_comp["정확대응비중(%)"], errors="coerce").fillna(0).clip(0, 100)
                    proc_comp["초과+비정형비중(%)"] = pd.to_numeric(proc_comp["초과+비정형비중(%)"], errors="coerce").fillna(0).clip(0, 300)

                    _den = (proc_comp["정확대응비중(%)"] + proc_comp["초과+비정형비중(%)"]).replace(0, np.nan)
                    proc_comp["정확(%)"] = (proc_comp["정확대응비중(%)"] / _den * 100.0).fillna(0.0).clip(0, 100)
                    proc_comp["초과+비정형(%)"] = (proc_comp["초과+비정형비중(%)"] / _den * 100.0).fillna(0.0).clip(0, 100)

                    proc_comp["공정"] = pd.Categorical(proc_comp["공정"].astype(str), categories=target_order, ordered=True)
                    proc_comp = proc_comp[proc_comp["공정"].notna()].copy()
                    proc_comp = proc_comp.sort_values("공정")

                    # 공정비교(가로막대) 행:
                    # - Plotly stacked bar는 좌측 라운딩이 제한적이라, HTML 막대로 표현(양끝 라운드)
                    # - hover는 CSS tooltip로 즉시 표시
                    bar_row_cols = st.columns([1, 1, 1, 0.55], gap="large")
                    with bar_row_cols[3]:
                        bar_legend_html = f"""
                        <div style='padding:12px 12px; border:1px solid #E5E7EB; border-radius:12px; background:#F9FAFB;'>
                          <div style='font-weight:800; color:#111827; margin-bottom:10px;'>범례(공정비교)</div>
                          <div style='color:#6B7280; font-size:12px; margin:-4px 0 10px 0;'>공정별 (정확 vs 감점요인) 구성비</div>
                          <div style='display:flex; flex-direction:column; gap:10px;'>
                            <div style='display:flex; align-items:center; gap:10px;'><span style='width:12px; height:12px; border-radius:3px; background:{BALANCE_COLORS['정확']}; display:inline-block;'></span><b>정확</b></div>
                            <div style='display:flex; align-items:center; gap:10px;'><span style='width:12px; height:12px; border-radius:3px; background:{BALANCE_COLORS['초과+비정형']}; display:inline-block;'></span><b>초과+비정형</b></div>
                          </div>
                        </div>
                        """
                        st.markdown(bar_legend_html, unsafe_allow_html=True)

                    st.markdown(
                        """
                        <style>
                          .pb-wrap{position:relative;}
                          .pb-tip{
                            position:absolute; left:84px; top:-2px; transform:translateY(-100%);
                            background:#FFFFFF; color:#111827; border:1px solid #E5E7EB;
                            border-radius:10px; padding:8px 10px; font-size:12px;
                            box-shadow:0 8px 24px rgba(15,23,42,0.10);
                            opacity:0; pointer-events:none; transition:opacity 0.08s ease-in;
                            z-index:9999; white-space:nowrap;
                          }
                          .pb-wrap:hover .pb-tip{opacity:1;}
                        </style>
                        """,
                        unsafe_allow_html=True,
                    )
                    for idx, g in enumerate(factory_order):
                        with bar_row_cols[idx]:
                            sub = proc_comp[proc_comp["공장그룹"].astype(str) == str(g)].copy()
                            if len(sub) == 0:
                                st.info("공정별 데이터 없음")
                                continue

                            base = sub[["공정", "정확(%)", "초과+비정형(%)"]].copy()
                            base["공정"] = base["공정"].astype(str)

                            rows_html: list[str] = []
                            for _, r in base.iterrows():
                                proc_name = str(r["공정"])
                                p_ok = float(pd.to_numeric(r.get("정확(%)", 0), errors="coerce"))
                                p_bad = float(pd.to_numeric(r.get("초과+비정형(%)", 0), errors="coerce"))
                                p_ok = max(0.0, min(100.0, p_ok))
                                p_bad = max(0.0, min(100.0, p_bad))
                                tip = f"{proc_name} · 정확 {p_ok:.1f}% / 초과+비정형 {p_bad:.1f}%"
                                rows_html.append(
                                    "<div class='pb-wrap' style='display:flex; align-items:center; gap:10px; margin:10px 0;'>"
                                    f"<div style='width:72px; color:#111827; font-size:13px;'>{proc_name}</div>"
                                    f"<div class='pb-tip'>{tip}</div>"
                                    "<div style='flex:1; height:28px; border-radius:14px; background:#E5E7EB; overflow:hidden;'>"
                                    "<div style='display:flex; height:100%; width:100%;'>"
                                    f"<div style='width:{p_ok:.2f}%; background:{BALANCE_COLORS['정확']}; display:flex; align-items:center; justify-content:center; font-size:14px; color:white; font-weight:800; border-top-left-radius:14px; border-bottom-left-radius:14px;'>{p_ok:.1f}%</div>"
                                    f"<div style='width:{p_bad:.2f}%; background:{BALANCE_COLORS['초과+비정형']}; display:flex; align-items:center; justify-content:center; font-size:14px; color:white; font-weight:800; border-top-right-radius:14px; border-bottom-right-radius:14px;'>{p_bad:.1f}%</div>"
                                    "</div>"
                                    "</div>"
                                    "</div>"
                                )
                            st.markdown("".join(rows_html), unsafe_allow_html=True)

                st.markdown("<div style='height:22px'></div>", unsafe_allow_html=True)
                ctl_l, ctl_r = st.columns([2.2, 1.0])
                with ctl_l:
                    show_tables = st.toggle("공장별 요약/상세 테이블 보기", value=False)
                with ctl_r:
                    # 원클릭 다운로드를 위해(생산운영현황 탭과 동일 패턴):
                    # - 필요 시 미리(리런 중) 생성해 session_state에 보관
                    # - download_button은 준비된 bytes를 바로 내려줌
                    export_key = "balance_export_excel_bytes"
                    export_sig_key = "balance_export_excel_signature"
                    signature = (
                        str(start_date),
                        str(end_date),
                        int(len(proc)),
                        int(len(det)) if isinstance(det, pd.DataFrame) else 0,
                    )
                    needs_rebuild = (
                        export_key not in st.session_state
                        or export_sig_key not in st.session_state
                        or st.session_state.get(export_sig_key) != signature
                    )
                    if needs_rebuild:
                        with st.spinner("다운로드 파일 준비 중..."):
                            summary_view, det_show = build_balance_tables_for_export(proc, det)
                            st.session_state[export_key] = build_two_sheet_excel(summary_view, det_show, sheet1="공장별요약", sheet2="상세테이블")
                            st.session_state[export_sig_key] = signature
                    data = st.session_state.get(export_key, b"")
                    st.download_button(
                        "요약+상세 다운로드 (xlsx)",
                        data=data,
                        file_name=f"공정밸런스_요약+상세_{start_date}_{end_date}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary",
                        disabled=not bool(data),
                    )

                if show_tables:
                    summary_view, det_show = build_balance_tables_for_export(proc, det)

                    with st.expander("공장별 요약", expanded=False):
                        _summary_fmt: dict[str, str] = {}
                        for c in ["실적수량", "유효생산량", "과생산량", "불필요생산량", "부족수량", "필요수량"]:
                            if c in summary_view.columns:
                                _summary_fmt[c] = "{:,.0f}"
                        for c in summary_view.columns:
                            if isinstance(c, str) and c.endswith("(%)"):
                                _summary_fmt[c] = "{:.1f}%"
                        if "공정점수" in summary_view.columns:
                            _summary_fmt["공정점수"] = "{:.1f}"
                        st.dataframe(summary_view.style.format(_summary_fmt), use_container_width=True, height=420)

                    with st.expander("상세 테이블", expanded=False):
                        value_cols = [c for c in ["실적수량", "필요수량", "부족수량", "유효생산량", "과생산량", "불필요생산량"] if c in det_show.columns]
                        _det_fmt: dict[str, str] = {c: "{:,.0f}" for c in value_cols}
                        if len(det_show) <= 3000:
                            st.dataframe(det_show.style.format(_det_fmt), use_container_width=True, height=520)
                        else:
                            st.caption(f"표시 행이 많아({len(det_show):,}행) 빠른 렌더링 모드로 표시합니다.")
                            st.dataframe(det_show, use_container_width=True, height=520)
except Exception as e:
    st.error("❌ 오류가 발생했습니다.")
    st.code(_truncate_err_message(str(e)), language="text")
    st.info("결과 파일을 다시 생성해주세요.")
