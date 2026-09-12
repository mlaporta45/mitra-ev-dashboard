import os
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# Pull API key from Streamlit secrets if available (Streamlit Cloud deployment)
if "ANTHROPIC_API_KEY" in st.secrets and not os.environ.get("ANTHROPIC_API_KEY"):
    os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]

from database import (
    get_all_periods,
    get_metrics_for_period,
    get_covenants_for_period,
    get_metric_timeseries,
    get_all_metrics_wide,
    get_documents,
    insert_document,
    upsert_metrics,
    upsert_covenants,
    METRIC_NAMES,
    COVENANT_DEFINITIONS,
)
from extractor import extract_text, file_hash
from analyzer import extract_and_analyze
from excel_export import build_excel_export

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Mitra EV Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styles ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stSidebar"] { background-color: #1B3A6B; }
  [data-testid="stSidebar"] * { color: white !important; }
  .kpi-card { background:#F0F4FB; border-radius:10px; padding:16px 20px; margin-bottom:8px; border-left:4px solid #2E5FA3; }
  .kpi-label { font-size:0.78rem; color:#555; text-transform:uppercase; letter-spacing:0.05em; }
  .kpi-value { font-size:1.55rem; font-weight:700; color:#1B3A6B; }
  .kpi-delta { font-size:0.82rem; }
  .badge-pass { background:#C6EFCE; color:#276221; border-radius:6px; padding:3px 10px; font-weight:700; }
  .badge-fail { background:#FFC7CE; color:#9C0006; border-radius:6px; padding:3px 10px; font-weight:700; }
  .badge-na   { background:#EDEDED; color:#555;    border-radius:6px; padding:3px 10px; font-weight:600; }
  h1, h2, h3 { color:#1B3A6B; }
</style>
""", unsafe_allow_html=True)

LOGO_COLOR = "#2E5FA3"

# ── Sidebar nav ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚡ Mitra EV")
    st.markdown("---")
    page = st.radio(
        "Navigation",
        ["📊 Overview", "📤 Upload", "📈 Trends", "📋 Data"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.caption("Mitra EV Financial Dashboard")


# ── Helpers ────────────────────────────────────────────────────────────────────
def fmt_money(v):
    if v is None:
        return "—"
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"${v/1_000:.1f}K"
    return f"${v:,.0f}"


def fmt_val(v, unit):
    if v is None:
        return "—"
    if unit == "$":
        return fmt_money(v)
    if unit == "%":
        pct = v * 100 if v < 1 else v
        return f"{pct:.1f}%"
    return f"{v:,.1f}"


def delta_pct(curr, prev):
    if curr is None or prev is None or prev == 0:
        return None
    return (curr - prev) / abs(prev) * 100


def covenant_badge(status):
    if status == "pass":
        return '<span class="badge-pass">PASS</span>'
    if status == "fail":
        return '<span class="badge-fail">FAIL</span>'
    return '<span class="badge-na">N/A</span>'


def kpi_card(label, value_str, delta=None, delta_label="vs prior period"):
    delta_html = ""
    if delta is not None:
        arrow = "▲" if delta >= 0 else "▼"
        color = "green" if delta >= 0 else "red"
        delta_html = f'<div class="kpi-delta" style="color:{color}">{arrow} {abs(delta):.1f}% {delta_label}</div>'
    return f"""
<div class="kpi-card">
  <div class="kpi-label">{label}</div>
  <div class="kpi-value">{value_str}</div>
  {delta_html}
</div>"""


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Overview
# ══════════════════════════════════════════════════════════════════════════════
if page == "📊 Overview":
    st.title("Mitra EV — Financial Overview")

    periods = get_all_periods()
    if not periods:
        st.info("No data yet. Go to **Upload** to add your first report.")
        st.stop()

    # Period selector
    col_sel, col_exp = st.columns([3, 1])
    with col_sel:
        selected = st.selectbox("Reporting Period", periods[::-1], index=0)
    with col_exp:
        st.markdown("<br>", unsafe_allow_html=True)
        excel_bytes = build_excel_export()
        st.download_button(
            "⬇ Export to Excel",
            data=excel_bytes,
            file_name="mitra_ev_dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    metrics = get_metrics_for_period(selected)
    covenants = get_covenants_for_period(selected)

    # Prior period for delta
    idx = periods.index(selected)
    prior_metrics = get_metrics_for_period(periods[idx - 1]) if idx > 0 else {}

    def mv(name):
        info = metrics.get(name, {})
        return info.get("value") if info else None

    def pmv(name):
        info = prior_metrics.get(name, {})
        return info.get("value") if info else None

    # ── Covenants ──────────────────────────────────────────────────────────────
    st.markdown("### Covenant Compliance")
    cov_cols = st.columns(len(COVENANT_DEFINITIONS))
    for col, (cname, defn) in zip(cov_cols, COVENANT_DEFINITIONS.items()):
        info = covenants.get(cname, {})
        actual = info.get("actual") if info else None
        status = info.get("status") if info else None
        threshold_text = defn.get("threshold_text", "")
        with col:
            st.markdown(
                f"""<div class="kpi-card">
  <div class="kpi-label">{cname}</div>
  <div class="kpi-value">{fmt_val(actual, '$' if 'Liquidity' in cname or 'Worth' in cname else 'x')}</div>
  <div>Threshold: {threshold_text} &nbsp; {covenant_badge(status)}</div>
</div>""",
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # ── Revenue KPIs ───────────────────────────────────────────────────────────
    st.markdown("### Revenue & Profitability")
    rev_metrics = [
        ("Total Revenue", "$"),
        ("EBITDA", "$"),
        ("EV Vehicle Leases Revenue", "$"),
        ("DCFC Charging Fees Revenue", "$"),
        ("LCFS Credits Revenue", "$"),
        ("Other Revenue", "$"),
    ]
    for i in range(0, len(rev_metrics), 3):
        cols = st.columns(3)
        for col, (name, unit) in zip(cols, rev_metrics[i:i+3]):
            curr = mv(name)
            prev = pmv(name)
            with col:
                st.markdown(
                    kpi_card(name, fmt_val(curr, unit), delta_pct(curr, prev)),
                    unsafe_allow_html=True,
                )

    st.markdown("---")

    # ── Ops KPIs ───────────────────────────────────────────────────────────────
    st.markdown("### Operational KPIs")
    ops_metrics = [
        ("Vehicles in Service", "units"),
        ("Vehicles Under MLA", "units"),
        ("Truck MRR", "$"),
        ("DCFC in Service", "units"),
        ("DCFC Under SHA", "units"),
        ("DCFC Avg. Utilization Rate", "%"),
    ]
    for i in range(0, len(ops_metrics), 3):
        cols = st.columns(3)
        for col, (name, unit) in zip(cols, ops_metrics[i:i+3]):
            curr = mv(name)
            prev = pmv(name)
            with col:
                st.markdown(
                    kpi_card(name, fmt_val(curr, unit), delta_pct(curr, prev)),
                    unsafe_allow_html=True,
                )

    st.markdown("---")

    # ── Revenue waterfall chart ────────────────────────────────────────────────
    if any(mv(m) is not None for m in ["EV Vehicle Leases Revenue", "DCFC Charging Fees Revenue", "LCFS Credits Revenue", "Other Revenue"]):
        st.markdown("### Revenue Breakdown")
        cats = ["EV Leases", "DCFC Fees", "LCFS Credits", "Other"]
        vals = [
            mv("EV Vehicle Leases Revenue") or 0,
            mv("DCFC Charging Fees Revenue") or 0,
            mv("LCFS Credits Revenue") or 0,
            mv("Other Revenue") or 0,
        ]
        fig = px.bar(
            x=cats, y=vals,
            labels={"x": "Revenue Stream", "y": "Amount ($)"},
            color=cats,
            color_discrete_sequence=px.colors.qualitative.Safe,
        )
        fig.update_layout(showlegend=False, height=320)
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Upload
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📤 Upload":
    st.title("Upload Mitra EV Reports")
    st.markdown(
        "Upload financial reports, board decks, or statements. "
        "The AI will automatically detect the reporting period and extract all metrics."
    )

    _env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    _has_env_key = bool(_env_key)
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        value=_env_key,
        help="Pre-filled from environment/secrets. Leave blank if already set there." if _has_env_key else "Paste your Anthropic API key here",
        placeholder="sk-ant-..." if not _has_env_key else "",
    )

    uploaded = st.file_uploader(
        "Drop files here",
        type=["pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
    )

    if st.button("⚡ Process Files", type="primary", disabled=not uploaded):
        for uf in uploaded:
            data = uf.read()
            fhash = file_hash(data)

            with st.expander(f"📄 {uf.name}", expanded=True):
                log = st.empty()

                log.info("Extracting text...")
                try:
                    text = extract_text(uf.name, data)
                except Exception as e:
                    log.error(f"Extraction failed: {e}")
                    continue

                log.info("Analyzing with AI (this may take ~20 seconds)...")
                try:
                    result = extract_and_analyze(text, api_key or None)
                except Exception as e:
                    log.error(f"AI analysis failed: {e}")
                    continue

                period = result.get("period")
                if not period:
                    log.warning("Could not detect reporting period — skipping.")
                    continue

                doc_id = insert_document(uf.name, period, fhash)
                upsert_metrics(doc_id, period, result.get("metrics", {}))
                upsert_covenants(doc_id, period, result.get("covenants", {}))

                confidence = result.get("confidence", "unknown")
                label = result.get("period_label", period)
                commentary = result.get("commentary", "")

                log.success(
                    f"✅ Saved — Period: **{label}** | Confidence: **{confidence}**"
                )
                if commentary:
                    st.markdown(f"> {commentary}")

                # Preview extracted metrics
                mets = result.get("metrics", {})
                rows = []
                for name, info in mets.items():
                    if isinstance(info, dict) and info.get("value") is not None:
                        rows.append({"Metric": name, "Value": info.get("value_text") or str(info.get("value")), "Unit": info.get("unit", "")})
                if rows:
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### Previously Uploaded Documents")
    docs = get_documents()
    if docs:
        st.dataframe(pd.DataFrame(docs)[["filename", "period", "uploaded_at"]], use_container_width=True, hide_index=True)
    else:
        st.caption("No documents uploaded yet.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Trends
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📈 Trends":
    st.title("Trends — Month Over Month")

    periods = get_all_periods()
    if len(periods) < 2:
        st.info("Upload at least **2 periods** of data to see trends.")
        st.stop()

    TREND_GROUPS = {
        "Revenue": ["EV Vehicle Leases Revenue", "DCFC Charging Fees Revenue", "LCFS Credits Revenue", "Other Revenue", "Total Revenue"],
        "Profitability": ["Total Revenue", "EBITDA"],
        "Balance Sheet": ["Cash Position", "Total Debt"],
        "Fleet": ["Vehicles in Service", "Vehicles Under MLA"],
        "Charging": ["DCFC in Service", "DCFC Under SHA", "DCFC Avg. Utilization Rate"],
        "Revenue per Unit": ["Truck MRR"],
    }

    tab_names = list(TREND_GROUPS.keys())
    tabs = st.tabs(tab_names)

    for tab, (group_name, metric_list) in zip(tabs, TREND_GROUPS.items()):
        with tab:
            fig = go.Figure()
            has_data = False
            for metric in metric_list:
                series = get_metric_timeseries(metric)
                if not series:
                    continue
                xs = [r["period"] for r in series]
                ys = [r["value"] for r in series]
                unit = series[0].get("unit", "")
                if unit == "%" :
                    ys = [v * 100 if v is not None and v < 1 else v for v in ys]
                fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines+markers", name=metric))
                has_data = True

            if not has_data:
                st.caption("No data for this group yet.")
            else:
                fig.update_layout(
                    height=420,
                    xaxis_title="Period",
                    yaxis_title="Value",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    margin=dict(l=40, r=20, t=30, b=40),
                )
                st.plotly_chart(fig, use_container_width=True)

    # ── Covenant trend ─────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Covenant Compliance Over Time")

    cov_rows = []
    for p in periods:
        covs = get_covenants_for_period(p)
        for cname, info in covs.items():
            cov_rows.append({
                "Period": p,
                "Covenant": cname,
                "Actual": info.get("actual"),
                "Status": info.get("status", "N/A"),
            })

    if cov_rows:
        cov_df = pd.DataFrame(cov_rows)
        for cname in cov_df["Covenant"].unique():
            sub = cov_df[cov_df["Covenant"] == cname].dropna(subset=["Actual"])
            if sub.empty:
                continue
            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=sub["Period"], y=sub["Actual"],
                mode="lines+markers",
                name=cname,
                marker=dict(
                    color=["green" if s == "pass" else "red" if s == "fail" else "gray"
                           for s in sub["Status"]],
                    size=10,
                ),
            ))
            defn = COVENANT_DEFINITIONS.get(cname, {})
            thresh = defn.get("threshold")
            if thresh is not None:
                fig2.add_hline(y=thresh, line_dash="dash", line_color="orange",
                               annotation_text=f"Threshold: {defn.get('threshold_text','')}")
            fig2.update_layout(title=cname, height=300, margin=dict(l=40, r=20, t=40, b=40))
            st.plotly_chart(fig2, use_container_width=True)
    else:
        st.caption("No covenant data yet.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Data
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Data":
    st.title("Raw Data")

    col1, col2 = st.columns([1, 1])
    with col1:
        periods = get_all_periods()
        if not periods:
            st.info("No data yet.")
            st.stop()
        filter_period = st.selectbox("Filter by period (blank = all)", ["All"] + periods[::-1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        excel_bytes = build_excel_export()
        st.download_button(
            "⬇ Export All to Excel",
            data=excel_bytes,
            file_name="mitra_ev_dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    df = get_all_metrics_wide()
    if filter_period != "All":
        df = df[df["period"] == filter_period]

    # Pivot for readability
    if not df.empty:
        pivot = df.pivot_table(index="metric_name", columns="period", values="value", aggfunc="first")
        pivot = pivot.reindex(columns=sorted(pivot.columns))
        st.dataframe(pivot, use_container_width=True)
    else:
        st.caption("No metrics in database.")

    st.markdown("---")
    st.markdown("### Covenant Detail")
    if periods:
        for p in (periods if filter_period == "All" else [filter_period]):
            covs = get_covenants_for_period(p)
            if covs:
                rows = []
                for cname, info in covs.items():
                    rows.append({
                        "Period": p,
                        "Covenant": cname,
                        "Actual": info.get("actual"),
                        "Threshold": info.get("threshold_text"),
                        "Status": info.get("status", "N/A"),
                    })
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
