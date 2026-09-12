import io
from database import get_all_periods, get_metrics_for_period, get_covenants_for_period, METRIC_NAMES, COVENANT_DEFINITIONS


def build_excel_export() -> bytes:
    import xlsxwriter

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    # ── Formats ────────────────────────────────────────────────────────────────
    header_fmt = wb.add_format({
        "bold": True, "bg_color": "#1B3A6B", "font_color": "white",
        "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
    })
    label_fmt = wb.add_format({
        "bold": True, "bg_color": "#E8EDF5", "border": 1,
        "align": "left", "valign": "vcenter",
    })
    money_fmt = wb.add_format({"num_format": '$#,##0', "border": 1, "align": "right"})
    money_k_fmt = wb.add_format({"num_format": '$#,##0,', "border": 1, "align": "right"})
    pct_fmt = wb.add_format({"num_format": "0.0%", "border": 1, "align": "right"})
    num_fmt = wb.add_format({"num_format": "#,##0.0", "border": 1, "align": "right"})
    text_fmt = wb.add_format({"border": 1, "align": "left"})
    pass_fmt = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#276221", "bold": True, "border": 1, "align": "center"})
    fail_fmt = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True, "border": 1, "align": "center"})
    na_fmt = wb.add_format({"bg_color": "#EDEDED", "border": 1, "align": "center"})
    section_fmt = wb.add_format({
        "bold": True, "bg_color": "#2E5FA3", "font_color": "white",
        "border": 1, "align": "left", "valign": "vcenter",
    })

    periods = get_all_periods()

    # ── Summary Sheet ──────────────────────────────────────────────────────────
    ws = wb.add_worksheet("Financial Summary")
    ws.set_zoom(90)
    ws.freeze_panes(2, 1)

    # Title row
    ws.merge_range(0, 0, 0, len(periods), "Mitra EV – Financial & Operational Summary", header_fmt)
    ws.set_row(0, 24)

    # Column headers
    ws.write(1, 0, "Metric", header_fmt)
    ws.set_column(0, 0, 36)
    for j, p in enumerate(periods):
        ws.write(1, j + 1, p, header_fmt)
        ws.set_column(j + 1, j + 1, 16)

    REVENUE_METRICS = [
        "EV Vehicle Leases Revenue",
        "DCFC Charging Fees Revenue",
        "LCFS Credits Revenue",
        "Other Revenue",
        "Total Revenue",
        "EBITDA",
    ]
    BALANCE_METRICS = [
        "Cash Position",
        "Total Debt",
    ]
    OPS_METRICS = [
        "Vehicles in Service",
        "Vehicles Under MLA",
        "Truck MRR",
        "DCFC in Service",
        "DCFC Under SHA",
        "DCFC Avg. Utilization Rate",
    ]

    period_data = {p: get_metrics_for_period(p) for p in periods}

    def write_section(row, title, metrics):
        ws.merge_range(row, 0, row, len(periods), title, section_fmt)
        row += 1
        for metric in metrics:
            ws.write(row, 0, metric, label_fmt)
            for j, p in enumerate(periods):
                info = period_data[p].get(metric, {})
                val = info.get("value") if info else None
                unit = (info.get("unit") or "") if info else ""
                if val is None:
                    ws.write(row, j + 1, "—", na_fmt)
                elif unit == "$":
                    ws.write_number(row, j + 1, val, money_fmt)
                elif unit == "%":
                    ws.write_number(row, j + 1, val / 100 if val > 1 else val, pct_fmt)
                else:
                    ws.write_number(row, j + 1, val, num_fmt)
            row += 1
        return row

    row = 2
    row = write_section(row, "Revenue & Profitability ($)", REVENUE_METRICS)
    row = write_section(row, "Balance Sheet ($)", BALANCE_METRICS)
    row = write_section(row, "Operational KPIs", OPS_METRICS)

    # ── Covenants Sheet ────────────────────────────────────────────────────────
    cov_ws = wb.add_worksheet("Covenant Tracking")
    cov_ws.set_zoom(90)
    cov_ws.freeze_panes(2, 1)
    cov_ws.merge_range(0, 0, 0, len(periods) * 2, "Mitra EV – Covenant Compliance Tracker", header_fmt)
    cov_ws.set_row(0, 24)
    cov_ws.write(1, 0, "Covenant", header_fmt)
    cov_ws.set_column(0, 0, 30)
    cov_ws.write(1, 1, "Threshold", header_fmt)
    cov_ws.set_column(1, 1, 16)

    col = 2
    for p in periods:
        cov_ws.merge_range(1, col, 1, col + 1, p, header_fmt)
        cov_ws.set_column(col, col, 14)
        cov_ws.set_column(col + 1, col + 1, 10)
        col += 2

    period_covs = {p: get_covenants_for_period(p) for p in periods}
    cov_names = list(COVENANT_DEFINITIONS.keys())

    for i, cname in enumerate(cov_names):
        r = i + 2
        cov_ws.write(r, 0, cname, label_fmt)
        defn = COVENANT_DEFINITIONS[cname]
        cov_ws.write(r, 1, defn.get("threshold_text", "—"), text_fmt)
        col = 2
        for p in periods:
            info = period_covs[p].get(cname, {})
            actual = info.get("actual") if info else None
            status = info.get("status") if info else None
            if actual is None:
                cov_ws.write(r, col, "—", na_fmt)
            else:
                cov_ws.write_number(r, col, actual, num_fmt)
            if status == "pass":
                cov_ws.write(r, col + 1, "PASS", pass_fmt)
            elif status == "fail":
                cov_ws.write(r, col + 1, "FAIL", fail_fmt)
            else:
                cov_ws.write(r, col + 1, "N/A", na_fmt)
            col += 2

    wb.close()
    return output.getvalue()
