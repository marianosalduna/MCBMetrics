"""
AUMI Global — Cash Book Dashboard
Run: streamlit run app.py
"""

import io
import warnings
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AUMI Global | Cash Book Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    div[data-testid="metric-container"] {
        background: #f0f4f8;
        border: 1px solid #d0d8e4;
        border-radius: 10px;
        padding: 12px 16px;
    }
    h2 { border-bottom: 2px solid #0066cc; padding-bottom: 6px; }
    .stDownloadButton > button { font-size: 13px; }
</style>
""",
    unsafe_allow_html=True,
)

# ─── Helpers ─────────────────────────────────────────────────────────────────
EXCEL_EPOCH = datetime(1899, 12, 30)

# Plausible modern Excel date serial range (~1998–2060)
_MIN_SERIAL = 36_000
_MAX_SERIAL = 58_000


def excel_serial_to_date(val):
    """Convert an Excel date serial to Python datetime; return None on failure."""
    if isinstance(val, (datetime, pd.Timestamp)):
        return pd.Timestamp(val)
    try:
        n = float(val)
        if _MIN_SERIAL <= n <= _MAX_SERIAL:
            return pd.Timestamp(EXCEL_EPOCH + timedelta(days=int(n)))
    except (TypeError, ValueError):
        pass
    return None


def fmt_usd(val):
    try:
        return f"${float(val):,.2f}"
    except Exception:
        return str(val)


def fmt_pct(val, decimals=2):
    try:
        return f"{float(val):.{decimals}f}%"
    except Exception:
        return str(val)


# ─── Data loading (cached) ────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_and_process(file_bytes: bytes):
    """
    Returns
    -------
    cb : pd.DataFrame   — processed Master Cash Book rows
    metrics : pd.DataFrame — processed monthly Cash Metrics rows
    """

    # ── Master Cash Book ──────────────────────────────────────────────────────
    # Column positions (0-indexed, Excel columns A=0 … AG=32):
    #   E  = 4  → CASH RECEIVED DATE
    #   F  = 5  → DEPOSIT AMOUNT
    #   J  = 9  → Member Name
    #   O  = 14 → Assigned to (worker initials)
    #   P  = 15 → UNMATCHED CASH (UMC)
    #   Y–AG = 24:33 → matching amounts (9 columns)

    raw_cb = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="MASTER CASH BOOK",
        header=0,
        engine="openpyxl",
    )

    cb = pd.DataFrame(
        {
            "Date_Raw": raw_cb.iloc[:, 4],
            "Deposit": pd.to_numeric(raw_cb.iloc[:, 5], errors="coerce"),
            "Member": raw_cb.iloc[:, 9].astype(str).str.strip(),
            "Worker": raw_cb.iloc[:, 14].astype(str).str.strip(),
            "UMC": pd.to_numeric(raw_cb.iloc[:, 15], errors="coerce").fillna(0),
        }
    )

    # Sum Y:AG (indices 24–32 inclusive → slice [24:33])
    match_block = raw_cb.iloc[:, 24:33].apply(pd.to_numeric, errors="coerce").fillna(0)
    cb["Matched"] = match_block.sum(axis=1)

    # Parse dates (may come as datetime objects or as Excel serials / floats)
    cb["Date"] = cb["Date_Raw"].apply(
        lambda x: pd.Timestamp(x) if isinstance(x, (datetime, pd.Timestamp))
        else excel_serial_to_date(x)
    )
    cb["Date"] = pd.to_datetime(cb["Date"], errors="coerce")

    # Keep only valid rows
    bad_workers = {"nan", "", "none", "None", "NaN"}
    cb = cb[
        cb["Deposit"].notna()
        & (cb["Deposit"] > 0)
        & cb["Date"].notna()
        & ~cb["Worker"].isin(bad_workers)
        & cb["Worker"].notna()
    ].copy()

    # Replace "nan" member names with "Unknown"
    cb["Member"] = cb["Member"].replace("nan", "Unknown")

    # Per-row percentages (capped at 100 % for matched)
    dep = cb["Deposit"].replace(0, np.nan)
    cb["Pct_Matched"] = (cb["Matched"] / dep * 100).clip(upper=200).round(2)
    cb["Pct_UMC"] = (cb["UMC"] / dep * 100).round(2)

    # ── Cash Metrics ──────────────────────────────────────────────────────────
    # Sheet layout (Excel rows, columns O:S):
    #   Row 1  → "Cash Movement – Do Not change layout"
    #   Row 2  → "US $"
    #   Row 3  → column headers (Month | Cash Received | Total Reconciled | AutoRek | Unallocated)
    #   Row 4  → annual total labelled "2023"
    #   Row 5+ → monthly rows with Excel date serial in col O

    raw_cm = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="Cash Metrics",
        header=None,
        engine="openpyxl",
    )

    # Slice: rows 3-40 (0-indexed), columns O-S (indices 14-18)
    section = raw_cm.iloc[3:41, 14:19].copy()
    section.columns = ["Month_Raw", "Cash_Received", "Total_Reconciled", "AutoRek", "UMC_Monthly"]
    section = section.reset_index(drop=True)

    section["Month"] = section["Month_Raw"].apply(excel_serial_to_date)

    for col in ["Cash_Received", "Total_Reconciled", "AutoRek", "UMC_Monthly"]:
        section[col] = pd.to_numeric(section[col], errors="coerce").fillna(0)

    metrics = section[
        section["Month"].notna() & (section["Cash_Received"] > 0)
    ].copy()
    metrics["Month"] = pd.to_datetime(metrics["Month"])

    # Percentages
    rcv = metrics["Cash_Received"].replace(0, np.nan)
    metrics["Pct_Reconciled"] = (metrics["Total_Reconciled"] / rcv * 100).round(2)
    metrics["Pct_AutoRek"] = (metrics["AutoRek"] / rcv * 100).round(2)
    metrics["Pct_UMC"] = (metrics["UMC_Monthly"] / rcv * 100).round(4)
    metrics["Month_Label"] = metrics["Month"].dt.strftime("%b %Y")

    return cb, metrics


# ─── Header ───────────────────────────────────────────────────────────────────
st.title("💰 AUMI Global | Cash Book Dashboard")
st.markdown("*Deposit matching & cash allocation report — January 2025 to present*")
st.divider()

uploaded = st.file_uploader(
    "📁 Drop the **MASTER CASH BOOK** Excel file here to load the report",
    type=["xlsx"],
)

if uploaded is None:
    st.info(
        "**Upload your file above to generate the report.**\n\n"
        "The dashboard will show:\n"
        "- **Worker & Member Report** — every deposit grouped by assigned worker, "
        "with amounts matched, UMC, and match-rate percentages\n"
        "- **Monthly Cash Metrics** — month-by-month cash received, reconciled, "
        "and outstanding UMC from the Cash Metrics sheet (Jan 2025 →)"
    )
    st.stop()

# ─── Load ─────────────────────────────────────────────────────────────────────
with st.spinner("Loading and processing data…"):
    try:
        file_bytes = uploaded.getvalue()
        cb_all, metrics_all = load_and_process(file_bytes)
    except Exception as exc:
        st.error(f"❌ Could not read the file: {exc}")
        st.stop()

START = pd.Timestamp("2025-01-01")
TODAY = pd.Timestamp(datetime.today().date())

cb = cb_all[(cb_all["Date"] >= START) & (cb_all["Date"] <= TODAY)].copy()
metrics = metrics_all[
    (metrics_all["Month"] >= START) & (metrics_all["Month"] <= TODAY)
].sort_values("Month")

if cb.empty:
    st.warning("No records found from January 2025 onwards in the Master Cash Book.")
    st.stop()

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["👥 Worker & Member Report", "📈 Monthly Cash Metrics"])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Worker & Member Report
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:

    # Top-level KPIs
    total_dep = cb["Deposit"].sum()
    total_matched = cb["Matched"].sum()
    total_umc = cb["UMC"].sum()
    overall_pct = total_matched / total_dep * 100 if total_dep else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total Deposits", fmt_usd(total_dep))
    k2.metric("Total Matched", fmt_usd(total_matched))
    k3.metric("Overall Match Rate", fmt_pct(overall_pct, 1))
    k4.metric("Total UMC", fmt_usd(total_umc))
    k5.metric("Total Records", f"{len(cb):,}")

    st.divider()

    # Worker filter
    all_workers = sorted(cb["Worker"].unique())
    selected_workers = st.multiselect(
        "Filter workers (all selected by default):",
        all_workers,
        default=all_workers,
    )

    if not selected_workers:
        st.warning("Please select at least one worker.")
        st.stop()

    cb_view = cb[cb["Worker"].isin(selected_workers)].copy()

    # ── Worker summary ──────────────────────────────────────────────────────
    wsum = (
        cb_view.groupby("Worker")
        .agg(
            Total_Deposit=("Deposit", "sum"),
            Total_Matched=("Matched", "sum"),
            Total_UMC=("UMC", "sum"),
            Num_Deposits=("Deposit", "count"),
            Num_Members=("Member", "nunique"),
        )
        .reset_index()
    )
    wsum["Pct_Matched"] = (wsum["Total_Matched"] / wsum["Total_Deposit"] * 100).round(2)
    wsum["Pct_UMC"] = (wsum["Total_UMC"] / wsum["Total_Deposit"] * 100).round(2)
    wsum = wsum.sort_values("Total_Deposit", ascending=False)

    st.subheader("Worker Summary")

    # Chart
    fig_w = go.Figure()
    fig_w.add_trace(
        go.Bar(
            x=wsum["Worker"],
            y=wsum["Total_Matched"],
            name="Matched",
            marker_color="#28a745",
            text=wsum["Pct_Matched"].apply(lambda x: f"{x:.1f}%"),
            textposition="inside",
            insidetextanchor="middle",
        )
    )
    fig_w.add_trace(
        go.Bar(
            x=wsum["Worker"],
            y=wsum["Total_UMC"],
            name="UMC (Unmatched)",
            marker_color="#dc3545",
            text=wsum["Pct_UMC"].apply(lambda x: f"{x:.2f}%"),
            textposition="inside",
            insidetextanchor="middle",
        )
    )
    fig_w.update_layout(
        barmode="stack",
        title="Total Deposits by Worker — Matched vs Unmatched (UMC)",
        xaxis_title="Worker",
        yaxis=dict(tickformat="$,.0f", title="Amount (USD)"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=380,
        margin=dict(t=60, b=20),
    )
    st.plotly_chart(fig_w, use_container_width=True)

    # Summary table
    disp_wsum = pd.DataFrame(
        {
            "Worker": wsum["Worker"],
            "# Members": wsum["Num_Members"].astype(int),
            "# Deposits": wsum["Num_Deposits"].astype(int),
            "Total Deposit": wsum["Total_Deposit"].apply(fmt_usd),
            "Matched": wsum["Total_Matched"].apply(fmt_usd),
            "% Matched": wsum["Pct_Matched"].apply(lambda x: fmt_pct(x, 2)),
            "Unmatched (UMC)": wsum["Total_UMC"].apply(fmt_usd),
            "% UMC": wsum["Pct_UMC"].apply(lambda x: fmt_pct(x, 4)),
        }
    )
    st.dataframe(disp_wsum, use_container_width=True, hide_index=True)

    csv_wsum = wsum.to_csv(index=False).encode()
    st.download_button(
        "⬇️ Download Worker Summary CSV",
        csv_wsum,
        "worker_summary.csv",
        "text/csv",
        key="dl_wsum",
    )

    st.divider()
    st.subheader("Detail by Worker")
    st.caption(
        "Each section shows a member-level breakdown followed by every individual deposit. "
        "Expand a worker row to drill down."
    )

    # ── Per-worker expanders ────────────────────────────────────────────────
    for worker in sorted(selected_workers):
        wdf = cb_view[cb_view["Worker"] == worker].copy()
        if wdf.empty:
            continue

        w_row = wsum[wsum["Worker"] == worker]
        if w_row.empty:
            continue
        w_row = w_row.iloc[0]

        pct = w_row["Pct_Matched"]
        icon = "✅" if pct >= 95 else ("⚠️" if pct >= 80 else "🔴")
        umc_str = fmt_usd(w_row["Total_UMC"])

        label = (
            f"**{worker}** — "
            f"{int(w_row['Num_Members'])} members · "
            f"{int(w_row['Num_Deposits'])} deposits · "
            f"{fmt_usd(w_row['Total_Deposit'])} total · "
            f"{icon} {pct:.1f}% matched · "
            f"UMC: {umc_str}"
        )

        with st.expander(label):

            # Member summary
            msum = (
                wdf.groupby("Member")
                .agg(
                    Total_Deposit=("Deposit", "sum"),
                    Total_Matched=("Matched", "sum"),
                    Total_UMC=("UMC", "sum"),
                    Deposits=("Deposit", "count"),
                )
                .reset_index()
            )
            dep_s = msum["Total_Deposit"].replace(0, np.nan)
            msum["Pct_Matched"] = (msum["Total_Matched"] / dep_s * 100).round(2)
            msum["Pct_UMC"] = (msum["Total_UMC"] / dep_s * 100).round(4)
            msum = msum.sort_values("Total_Deposit", ascending=False)

            st.markdown("**Member Summary**")
            disp_msum = pd.DataFrame(
                {
                    "Member": msum["Member"],
                    "# Deposits": msum["Deposits"].astype(int),
                    "Total Deposit": msum["Total_Deposit"].apply(fmt_usd),
                    "Matched": msum["Total_Matched"].apply(fmt_usd),
                    "% Matched": msum["Pct_Matched"].apply(lambda x: fmt_pct(x, 2)),
                    "Unmatched (UMC)": msum["Total_UMC"].apply(fmt_usd),
                    "% UMC": msum["Pct_UMC"].apply(lambda x: fmt_pct(x, 4)),
                }
            )
            st.dataframe(disp_msum, use_container_width=True, hide_index=True)

            # All deposits
            st.markdown("**All Deposits**")
            deps = wdf[
                ["Date", "Member", "Deposit", "Matched", "UMC", "Pct_Matched", "Pct_UMC"]
            ].copy()
            deps = deps.sort_values(["Member", "Date"])
            disp_deps = pd.DataFrame(
                {
                    "Date": deps["Date"].dt.strftime("%Y-%m-%d"),
                    "Member": deps["Member"],
                    "Deposit Amount": deps["Deposit"].apply(fmt_usd),
                    "Matched": deps["Matched"].apply(fmt_usd),
                    "% Matched": deps["Pct_Matched"].apply(lambda x: fmt_pct(x, 2)),
                    "Unmatched (UMC)": deps["UMC"].apply(fmt_usd),
                    "% UMC": deps["Pct_UMC"].apply(lambda x: fmt_pct(x, 4)),
                }
            )
            st.dataframe(disp_deps, use_container_width=True, hide_index=True)

            csv_w = deps.to_csv(index=False).encode()
            st.download_button(
                f"⬇️ Download {worker} CSV",
                csv_w,
                f"worker_{worker}_deposits.csv",
                "text/csv",
                key=f"dl_w_{worker}",
            )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Monthly Cash Metrics
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:

    if metrics.empty:
        st.warning("No monthly Cash Metrics data found for January 2025 onwards.")
    else:
        total_recv = metrics["Cash_Received"].sum()
        total_rec = metrics["Total_Reconciled"].sum()
        total_ar = metrics["AutoRek"].sum()
        total_umc_m = metrics["UMC_Monthly"].sum()
        pct_rec = total_rec / total_recv * 100 if total_recv else 0
        pct_umc_m = total_umc_m / total_recv * 100 if total_recv else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Cash Received (2025+)", fmt_usd(total_recv))
        k2.metric("Total Reconciled", fmt_usd(total_rec), fmt_pct(pct_rec, 1) + " of received")
        k3.metric("AutoRek Matched", fmt_usd(total_ar))
        k4.metric("Total Outstanding UMC", fmt_usd(total_umc_m), fmt_pct(pct_umc_m, 4) + " of received")

        st.divider()

        # ── Cash flow chart ──────────────────────────────────────────────────
        st.subheader("Monthly Cash Flow")

        chart_type = st.radio("View as:", ["Grouped Bars", "Lines"], horizontal=True)

        if chart_type == "Grouped Bars":
            fig_m = go.Figure()
            for name, col, color in [
                ("Cash Received", "Cash_Received", "#0066cc"),
                ("Reconciled", "Total_Reconciled", "#28a745"),
                ("Outstanding UMC", "UMC_Monthly", "#dc3545"),
            ]:
                fig_m.add_trace(
                    go.Bar(
                        x=metrics["Month_Label"],
                        y=metrics[col],
                        name=name,
                        marker_color=color,
                    )
                )
            fig_m.update_layout(barmode="group")
        else:
            fig_m = go.Figure()
            for name, col, color, dash in [
                ("Cash Received", "Cash_Received", "#0066cc", "solid"),
                ("Reconciled", "Total_Reconciled", "#28a745", "solid"),
                ("Outstanding UMC", "UMC_Monthly", "#dc3545", "dot"),
            ]:
                fig_m.add_trace(
                    go.Scatter(
                        x=metrics["Month_Label"],
                        y=metrics[col],
                        name=name,
                        mode="lines+markers",
                        line=dict(color=color, width=2, dash=dash),
                    )
                )

        fig_m.update_layout(
            title="Monthly: Cash Received vs Reconciled vs UMC",
            xaxis_title="Month",
            yaxis=dict(tickformat="$,.0f", title="Amount (USD)"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=420,
            margin=dict(t=60, b=20),
        )
        st.plotly_chart(fig_m, use_container_width=True)

        # ── Reconciliation rate trend ────────────────────────────────────────
        fig_pct = go.Figure()
        fig_pct.add_trace(
            go.Scatter(
                x=metrics["Month_Label"],
                y=metrics["Pct_Reconciled"],
                name="% Reconciled",
                mode="lines+markers",
                line=dict(color="#28a745", width=2),
                fill="tozeroy",
                fillcolor="rgba(40,167,69,0.08)",
            )
        )
        fig_pct.add_hline(
            y=100,
            line_dash="dash",
            line_color="gray",
            annotation_text="100%",
            annotation_position="top right",
        )
        fig_pct.update_layout(
            title="Monthly Reconciliation Rate",
            xaxis_title="Month",
            yaxis=dict(tickformat=".1f", ticksuffix="%", title="% of Cash Received"),
            height=320,
            margin=dict(t=60, b=20),
        )
        st.plotly_chart(fig_pct, use_container_width=True)

        # ── Monthly data table ───────────────────────────────────────────────
        st.subheader("Monthly Data Table")
        disp_m = pd.DataFrame(
            {
                "Month": metrics["Month_Label"],
                "Cash Received": metrics["Cash_Received"].apply(fmt_usd),
                "Total Reconciled": metrics["Total_Reconciled"].apply(fmt_usd),
                "% Reconciled": metrics["Pct_Reconciled"].apply(lambda x: fmt_pct(x, 2)),
                "AutoRek Matched": metrics["AutoRek"].apply(fmt_usd),
                "% AutoRek": metrics["Pct_AutoRek"].apply(lambda x: fmt_pct(x, 2)),
                "Outstanding UMC": metrics["UMC_Monthly"].apply(fmt_usd),
                "% UMC": metrics["Pct_UMC"].apply(lambda x: fmt_pct(x, 4)),
            }
        )
        st.dataframe(disp_m, use_container_width=True, hide_index=True)

        csv_m = metrics.drop(columns=["Month_Raw"], errors="ignore").to_csv(index=False).encode()
        st.download_button(
            "⬇️ Download Monthly Metrics CSV",
            csv_m,
            "monthly_cash_metrics.csv",
            "text/csv",
            key="dl_metrics",
        )

# ─── Footer ───────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "Data filtered: January 1, 2025 → today · "
    "Matched = sum of columns Y–AG (MASTER CASH BOOK) · "
    "UMC = column P · "
    "Monthly metrics from Cash Metrics tab, columns O–S"
)
