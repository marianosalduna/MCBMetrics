"""
AUMI Global — Cash Book Dashboard  v2
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
    .block-container { padding-top: 1.5rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ─── Constants ────────────────────────────────────────────────────────────────
EXCEL_EPOCH = datetime(1899, 12, 30)
_MIN_SERIAL, _MAX_SERIAL = 36_000, 58_000
START_DATE = pd.Timestamp("2025-01-01")


# ─── Helpers ─────────────────────────────────────────────────────────────────
def excel_serial_to_date(val):
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


def fmt_m(val):
    """Format large dollar amounts as $X.Xm"""
    try:
        v = float(val)
        if abs(v) >= 1_000_000:
            return f"${v/1_000_000:.2f}M"
        return f"${v:,.0f}"
    except Exception:
        return str(val)


def fmt_pct(val, decimals=2):
    try:
        return f"{float(val):.{decimals}f}%"
    except Exception:
        return str(val)


def worker_label(initials, names_dict):
    name = names_dict.get(initials.strip(), "")
    return f"{initials} — {name}" if name else initials


def pct_delta_str(curr, prev):
    """Return formatted delta string like '+12.3%' or '-5.1%'."""
    if not prev or prev == 0:
        return None
    d = (curr - prev) / abs(prev) * 100
    sign = "+" if d >= 0 else ""
    return f"{sign}{d:.1f}%"


# ─── Data loading ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_and_process(file_bytes: bytes):
    # ── 1. Master Cash Book ───────────────────────────────────────────────────
    # Columns (0-indexed): A=0 Alias, B=1 CMG, C=2 Insurer, E=4 Date,
    # F=5 Deposit, J=9 Member, O=14 Worker, P=15 UMC, Y-AG=24:33 Matched
    raw_cb = pd.read_excel(
        io.BytesIO(file_bytes), sheet_name="MASTER CASH BOOK",
        header=0, engine="openpyxl",
    )

    cb = pd.DataFrame({
        "Alias":    raw_cb.iloc[:, 0].astype(str).str.strip(),
        "CMG":      raw_cb.iloc[:, 1].astype(str).str.strip(),
        "Insurer":  raw_cb.iloc[:, 2].astype(str).str.strip(),
        "Date_Raw": raw_cb.iloc[:, 4],
        "Deposit":  pd.to_numeric(raw_cb.iloc[:, 5], errors="coerce"),
        "Member":   raw_cb.iloc[:, 9].astype(str).str.strip(),
        "Worker":   raw_cb.iloc[:, 14].astype(str).str.strip(),
        "UMC":      pd.to_numeric(raw_cb.iloc[:, 15], errors="coerce").fillna(0),
    })

    match_block = raw_cb.iloc[:, 24:33].apply(pd.to_numeric, errors="coerce").fillna(0)
    cb["Matched"] = match_block.sum(axis=1)

    cb["Date"] = cb["Date_Raw"].apply(
        lambda x: pd.Timestamp(x) if isinstance(x, (datetime, pd.Timestamp))
        else excel_serial_to_date(x)
    )
    cb["Date"] = pd.to_datetime(cb["Date"], errors="coerce")

    # Clean up
    bad = {"nan", "", "none", "None", "NaN"}
    cb = cb[
        cb["Deposit"].notna() & (cb["Deposit"] > 0) &
        cb["Date"].notna() &
        ~cb["Worker"].isin(bad) & cb["Worker"].notna()
    ].copy()

    cb["Member"]  = cb["Member"].replace("nan", "Unknown")
    cb["Alias"]   = cb["Alias"].replace("nan", "")
    cb["CMG"]     = cb["CMG"].replace("nan", "")
    cb["Insurer"] = cb["Insurer"].replace("nan", "")

    dep = cb["Deposit"].replace(0, np.nan)
    cb["Pct_Matched"] = (cb["Matched"] / dep * 100).clip(upper=200).round(2)
    cb["Pct_UMC"]     = (cb["UMC"].abs() / dep * 100).round(4)   # abs: negative UMC still counts
    cb["Month_Period"] = cb["Date"].dt.to_period("M")

    # ── 2. Cash Metrics monthly (O2:S40) ─────────────────────────────────────
    # Excel rows 2-40 → pandas indices 1-39; data rows from index 3 onward
    # (index 0=title, 1=US$, 2=headers, 3=2023 total, 4+=monthly)
    raw_cm = pd.read_excel(
        io.BytesIO(file_bytes), sheet_name="Cash Metrics",
        header=None, engine="openpyxl",
    )

    section = raw_cm.iloc[3:41, 14:19].copy()
    section.columns = ["Month_Raw", "Cash_Received", "Total_Reconciled", "AutoRek", "UMC_Monthly"]
    section = section.reset_index(drop=True)
    section["Month"] = section["Month_Raw"].apply(excel_serial_to_date)

    for col in ["Cash_Received", "Total_Reconciled", "AutoRek", "UMC_Monthly"]:
        section[col] = pd.to_numeric(section[col], errors="coerce").fillna(0)

    metrics = section[section["Month"].notna() & (section["Cash_Received"] > 0)].copy()
    metrics["Month"] = pd.to_datetime(metrics["Month"])

    rcv = metrics["Cash_Received"].replace(0, np.nan)
    metrics["Pct_Reconciled"] = (metrics["Total_Reconciled"] / rcv * 100).round(2)
    metrics["Pct_AutoRek"]    = (metrics["AutoRek"] / rcv * 100).round(2)
    metrics["Pct_UMC"]        = (metrics["UMC_Monthly"].abs() / rcv * 100).round(4)
    metrics["Month_Label"]    = metrics["Month"].dt.strftime("%b %Y")

    # ── 3. UMC Backlog (O49:P87) ──────────────────────────────────────────────
    # Excel rows 49-87 → pandas indices 48-86
    # Row 49 = "Unallocated", Row 50 = col headers, Row 51+ = data
    umc_bl_raw = raw_cm.iloc[50:87, 14:16].copy()
    umc_bl_raw.columns = ["Month_Raw", "UMC_Backlog"]
    umc_bl_raw = umc_bl_raw.reset_index(drop=True)
    umc_bl_raw["Month"] = umc_bl_raw["Month_Raw"].apply(excel_serial_to_date)
    umc_bl_raw["UMC_Backlog"] = pd.to_numeric(umc_bl_raw["UMC_Backlog"], errors="coerce").fillna(0)

    umc_bl = umc_bl_raw[umc_bl_raw["Month"].notna() & (umc_bl_raw["UMC_Backlog"] > 0)].copy()
    umc_bl["Month"] = pd.to_datetime(umc_bl["Month"])

    # ── 4. Worker names (Team Assignments K:L, rows 1-37) ────────────────────
    # K=index 10, L=index 11 (0-based)
    try:
        raw_team = pd.read_excel(
            io.BytesIO(file_bytes), sheet_name="Team Assignments",
            header=None, engine="openpyxl",
            usecols=[10, 11], nrows=37,
        )
        worker_names = {}
        for _, row in raw_team.iterrows():
            name     = str(row.iloc[0]).strip()
            initials = str(row.iloc[1]).strip()
            if name not in ("", "nan", "Names") and initials not in ("", "nan", "Initials"):
                worker_names[initials] = name
    except Exception:
        worker_names = {}

    return cb, metrics, umc_bl, worker_names


# ─── Header & uploader ───────────────────────────────────────────────────────
st.title("💰 AUMI Global | Cash Book Dashboard")
st.markdown("*Deposit matching & cash allocation — January 2025 → present*")
st.divider()

uploaded = st.file_uploader(
    "📁 Drop the **MASTER CASH BOOK** Excel file here",
    type=["xlsx"],
)

if uploaded is None:
    st.info(
        "**Upload the Excel file above to generate the report.**\n\n"
        "- **Worker Report** — monthly performance, deposits by member with match %, color-coded\n"
        "- **Member Search** — look up any member by name, alias or CMG\n"
        "- **Monthly Metrics** — month-by-month cash received, reconciled & UMC trends"
    )
    st.stop()

with st.spinner("Loading and processing…"):
    try:
        file_bytes = uploaded.getvalue()
        cb_all, metrics_all, umc_bl_all, WORKER_NAMES = load_and_process(file_bytes)
    except Exception as exc:
        st.error(f"❌ Could not read the file: {exc}")
        st.stop()

TODAY = pd.Timestamp(datetime.today().date())

cb      = cb_all[(cb_all["Date"] >= START_DATE) & (cb_all["Date"] <= TODAY)].copy()
metrics = metrics_all[(metrics_all["Month"] >= START_DATE) & (metrics_all["Month"] <= TODAY)].sort_values("Month").reset_index(drop=True)
umc_bl  = umc_bl_all[(umc_bl_all["Month"] >= START_DATE) & (umc_bl_all["Month"] <= TODAY)].sort_values("Month")

if cb.empty:
    st.warning("No records found from January 2025 onwards.")
    st.stop()

# Precompute month options (most recent first)
all_month_periods = sorted(cb["Month_Period"].unique(), reverse=True)
month_label_map   = {m: pd.Timestamp(m.start_time).strftime("%b %Y") for m in all_month_periods}
month_options     = [month_label_map[m] for m in all_month_periods]
month_label_to_period = {v: k for k, v in month_label_map.items()}

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["👥 Worker Report", "🔍 Member Search", "📈 Monthly Metrics"])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Worker Report
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:

    # ── Current Month KPIs (from Cash Metrics) ────────────────────────────────
    if len(metrics) >= 2:
        curr = metrics.iloc[-1]
        prev = metrics.iloc[-2]

        st.subheader(f"Current Month: {curr['Month_Label']}")

        k1, k2, k3, k4 = st.columns(4)

        d_recv = pct_delta_str(curr["Cash_Received"], prev["Cash_Received"])
        k1.metric("Cash Received", fmt_m(curr["Cash_Received"]),
                  delta=f"{d_recv} vs {prev['Month_Label']}" if d_recv else None)

        d_rec = pct_delta_str(curr["Total_Reconciled"], prev["Total_Reconciled"])
        k2.metric("Reconciled", fmt_m(curr["Total_Reconciled"]),
                  delta=f"{d_rec} vs {prev['Month_Label']}" if d_rec else None)

        d_pct = curr["Pct_Reconciled"] - prev["Pct_Reconciled"]
        k3.metric("Reconciliation Rate", fmt_pct(curr["Pct_Reconciled"], 1),
                  delta=f"{d_pct:+.1f}pp vs {prev['Month_Label']}")

        d_umc_pct = curr["Pct_UMC"] - prev["Pct_UMC"]
        k4.metric("UMC Rate", fmt_pct(curr["Pct_UMC"], 4),
                  delta=f"{d_umc_pct:+.4f}pp vs {prev['Month_Label']}",
                  delta_color="inverse")

    st.divider()

    # ── Filters ───────────────────────────────────────────────────────────────
    all_workers = sorted(cb["Worker"].unique())
    worker_options = [worker_label(w, WORKER_NAMES) for w in all_workers]

    col_f1, col_f2 = st.columns([2, 3])
    with col_f1:
        selected_worker_labels = st.multiselect(
            "Workers", worker_options, default=worker_options,
            help="Filter by worker"
        )
    selected_workers = [w.split(" — ")[0] for w in selected_worker_labels]

    if not selected_workers:
        st.warning("Select at least one worker.")
        st.stop()

    cb_view = cb[cb["Worker"].isin(selected_workers)].copy()

    # ── Worker Summary (all-time Jan 2025→) ───────────────────────────────────
    st.subheader("Worker Summary — Jan 2025 to Present")

    wsum = (
        cb_view.groupby("Worker")
        .agg(Total_Deposit=("Deposit", "sum"), Total_Matched=("Matched", "sum"),
             Total_UMC=("UMC", lambda x: x.abs().sum()),
             Num_Deposits=("Deposit", "count"), Num_Members=("Member", "nunique"))
        .reset_index()
    )
    dep_s = wsum["Total_Deposit"].replace(0, np.nan)
    wsum["Pct_Matched"] = (wsum["Total_Matched"] / dep_s * 100).round(2)
    wsum["Pct_UMC"]     = (wsum["Total_UMC"] / dep_s * 100).round(4)
    wsum["Worker_Label"] = wsum["Worker"].apply(lambda w: worker_label(w, WORKER_NAMES))
    wsum = wsum.sort_values("Total_Deposit", ascending=False)

    fig_w = go.Figure()
    fig_w.add_trace(go.Bar(
        x=wsum["Worker_Label"], y=wsum["Total_Matched"], name="Matched",
        marker_color="#28a745",
        text=wsum["Pct_Matched"].apply(lambda x: f"{x:.1f}%"), textposition="inside",
        insidetextanchor="middle",
    ))
    fig_w.add_trace(go.Bar(
        x=wsum["Worker_Label"], y=wsum["Total_UMC"], name="UMC",
        marker_color="#dc3545",
        text=wsum["Pct_UMC"].apply(lambda x: f"{x:.2f}%"), textposition="inside",
        insidetextanchor="middle",
    ))
    fig_w.update_layout(
        barmode="stack", title="Total Deposits: Matched vs UMC by Worker",
        xaxis_title="Worker", yaxis=dict(tickformat="$,.0f", title="Amount (USD)"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=360, margin=dict(t=60, b=20),
    )
    st.plotly_chart(fig_w, use_container_width=True)

    # Summary table with color coding
    wsum_display = pd.DataFrame({
        "Worker":        wsum["Worker_Label"],
        "# Members":     wsum["Num_Members"].astype(int),
        "# Deposits":    wsum["Num_Deposits"].astype(int),
        "Total Deposit": wsum["Total_Deposit"],
        "Matched":       wsum["Total_Matched"],
        "% Matched":     wsum["Pct_Matched"],
        "UMC":           wsum["Total_UMC"],
        "% UMC":         wsum["Pct_UMC"],
    })
    styled_wsum = (
        wsum_display.style
        .format({
            "Total Deposit": "${:,.2f}",
            "Matched": "${:,.2f}",
            "UMC": "${:,.2f}",
            "% Matched": "{:.2f}%",
            "% UMC": "{:.4f}%",
        })
        .hide(axis="index")
    )
    st.dataframe(styled_wsum, use_container_width=True)

    csv_wsum = wsum.to_csv(index=False).encode()
    st.download_button("⬇️ Worker Summary CSV", csv_wsum, "worker_summary.csv", "text/csv", key="dl_wsum")

    st.divider()

    # ── Worker Detail — Monthly View ──────────────────────────────────────────
    st.subheader("Worker Detail — Monthly View")

    col_mf, _ = st.columns([3, 2])
    with col_mf:
        selected_month_labels = st.multiselect(
            "Select month(s) to inspect",
            month_options,
            default=[month_options[0]],   # default: most recent month
            help="Filter deposits shown in each worker's expander",
        )

    if not selected_month_labels:
        st.info("Select at least one month above.")
    else:
        selected_periods = [month_label_to_period[lbl] for lbl in selected_month_labels]
        cb_month = cb_view[cb_view["Month_Period"].isin(selected_periods)].copy()

        if cb_month.empty:
            st.warning("No deposits found for the selected month(s) and workers.")
        else:
            # Per-worker expanders
            for worker in sorted(selected_workers):
                wdf = cb_month[cb_month["Worker"] == worker].copy()
                if wdf.empty:
                    continue

                # Worker totals for selected period
                w_dep  = wdf["Deposit"].sum()
                w_mat  = wdf["Matched"].sum()
                w_umc  = wdf["UMC"].abs().sum()
                w_pct  = w_mat / w_dep * 100 if w_dep else 0
                w_icon = "✅" if w_pct >= 95 else ("⚠️" if w_pct >= 80 else "🔴")
                w_name = worker_label(worker, WORKER_NAMES)

                label = (
                    f"**{w_name}** — {wdf['Member'].nunique()} members · "
                    f"{len(wdf)} deposits · {fmt_usd(w_dep)} · "
                    f"{w_icon} {w_pct:.1f}% matched · UMC: {fmt_usd(w_umc)}"
                )

                with st.expander(label):
                    # Monthly sub-breakdown (if multiple months selected)
                    if len(selected_periods) > 1:
                        monthly_grp = (
                            wdf.groupby("Month_Period")
                            .agg(Deposit=("Deposit", "sum"), Matched=("Matched", "sum"),
                                 UMC=("UMC", lambda x: x.abs().sum()),
                                 Members=("Member", "nunique"), Deposits=("Deposit", "count"))
                            .reset_index()
                        )
                        monthly_grp["Pct_Matched"] = (monthly_grp["Matched"] / monthly_grp["Deposit"].replace(0, np.nan) * 100).round(2)
                        monthly_grp["Pct_UMC"]     = (monthly_grp["UMC"] / monthly_grp["Deposit"].replace(0, np.nan) * 100).round(4)
                        monthly_grp["Month"] = monthly_grp["Month_Period"].apply(lambda p: pd.Timestamp(p.start_time).strftime("%b %Y"))
                        monthly_grp = monthly_grp.sort_values("Month_Period")

                        st.markdown("**Month-by-Month**")
                        mo_disp = pd.DataFrame({
                            "Month":       monthly_grp["Month"],
                            "# Members":   monthly_grp["Members"].astype(int),
                            "# Deposits":  monthly_grp["Deposits"].astype(int),
                            "Deposit":     monthly_grp["Deposit"],
                            "Matched":     monthly_grp["Matched"],
                            "% Matched":   monthly_grp["Pct_Matched"],
                            "UMC":         monthly_grp["UMC"],
                            "% UMC":       monthly_grp["Pct_UMC"],
                        })
                        styled_mo = (
                            mo_disp.style
                            .format({"Deposit": "${:,.2f}", "Matched": "${:,.2f}", "UMC": "${:,.2f}",
                                     "% Matched": "{:.2f}%", "% UMC": "{:.4f}%"})
                            .hide(axis="index")
                        )
                        st.dataframe(styled_mo, use_container_width=True)
                        st.markdown("---")

                    # Member summary
                    msum = (
                        wdf.groupby("Member")
                        .agg(Total_Deposit=("Deposit", "sum"), Total_Matched=("Matched", "sum"),
                             Total_UMC=("UMC", lambda x: x.abs().sum()), Deposits=("Deposit", "count"))
                        .reset_index()
                    )
                    dep_ms = msum["Total_Deposit"].replace(0, np.nan)
                    msum["Pct_Matched"] = (msum["Total_Matched"] / dep_ms * 100).round(2)
                    msum["Pct_UMC"]     = (msum["Total_UMC"] / dep_ms * 100).round(4)
                    msum = msum.sort_values("Total_Deposit", ascending=False)

                    st.markdown("**Member Summary**")
                    ms_disp = pd.DataFrame({
                        "Member":      msum["Member"],
                        "# Deposits":  msum["Deposits"].astype(int),
                        "Deposit":     msum["Total_Deposit"],
                        "Matched":     msum["Total_Matched"],
                        "% Matched":   msum["Pct_Matched"],
                        "UMC":         msum["Total_UMC"],
                        "% UMC":       msum["Pct_UMC"],
                    })
                    styled_ms = (
                        ms_disp.style
                        .format({"Deposit": "${:,.2f}", "Matched": "${:,.2f}", "UMC": "${:,.2f}",
                                 "% Matched": "{:.2f}%", "% UMC": "{:.4f}%"})
                        .hide(axis="index")
                    )
                    st.dataframe(styled_ms, use_container_width=True)

                    # All deposits
                    st.markdown("**All Deposits**")
                    deps = wdf[["Date", "Member", "Deposit", "Matched", "Pct_Matched", "UMC", "Pct_UMC"]].copy()
                    deps = deps.sort_values(["Member", "Date"])
                    deps_disp = pd.DataFrame({
                        "Date":        deps["Date"].dt.strftime("%Y-%m-%d"),
                        "Member":      deps["Member"],
                        "Deposit":     deps["Deposit"],
                        "Matched":     deps["Matched"],
                        "% Matched":   deps["Pct_Matched"],
                        "UMC":         deps["UMC"].abs(),
                        "% UMC":       deps["Pct_UMC"],
                    })
                    styled_deps = (
                        deps_disp.style
                        .format({"Deposit": "${:,.2f}", "Matched": "${:,.2f}", "UMC": "${:,.2f}",
                                 "% Matched": "{:.2f}%", "% UMC": "{:.4f}%"})
                        .hide(axis="index")
                    )
                    st.dataframe(styled_deps, use_container_width=True)

                    csv_w = deps.to_csv(index=False).encode()
                    st.download_button(
                        f"⬇️ Download {worker} CSV",
                        csv_w, f"worker_{worker}.csv", "text/csv",
                        key=f"dl_{worker}_{'-'.join(selected_month_labels)}",
                    )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Member Search
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Member Search")
    st.caption("Search by alias, member name, or CMG code. Results grouped by insurer.")

    search_q = st.text_input("Search:", placeholder="Type alias, member name, or CMG…").strip().lower()

    if search_q:
        mask = (
            cb["Alias"].str.lower().str.contains(search_q, na=False) |
            cb["Member"].str.lower().str.contains(search_q, na=False) |
            cb["CMG"].str.lower().str.contains(search_q, na=False)
        )
        results = cb[mask].copy()

        if results.empty:
            st.warning(f"No members found matching '{search_q}'.")
        else:
            # Unique members matched
            member_groups = results.groupby(["Member", "Alias", "CMG"])["Worker"].agg(
                lambda x: ", ".join(sorted(x.unique()))
            ).reset_index()

            st.markdown(f"**{len(member_groups)} member(s) found:**")

            for _, mg in member_groups.iterrows():
                member_name = mg["Member"]
                alias       = mg["Alias"]
                cmg         = mg["CMG"]
                workers_str = mg["Worker"]

                worker_full = ", ".join(
                    worker_label(w.strip(), WORKER_NAMES) for w in workers_str.split(",")
                )

                with st.expander(
                    f"**{member_name}** | Alias: {alias} | CMG: {cmg} | Worker: {worker_full}"
                ):
                    mem_df = results[results["Member"] == member_name].copy()

                    # Group by Insurer (Col C)
                    insurers = sorted(mem_df["Insurer"].dropna().unique())

                    for insurer in insurers:
                        ins_df = mem_df[mem_df["Insurer"] == insurer].sort_values("Date")

                        ins_dep = ins_df["Deposit"].sum()
                        ins_mat = ins_df["Matched"].sum()
                        ins_umc = ins_df["UMC"].abs().sum()
                        ins_pct = ins_mat / ins_dep * 100 if ins_dep else 0

                        if len(insurers) > 1:
                            st.markdown(f"**Insurer: {insurer}**")

                        summary_cols = st.columns(4)
                        summary_cols[0].metric("Total Deposit", fmt_usd(ins_dep))
                        summary_cols[1].metric("Matched", fmt_usd(ins_mat))
                        summary_cols[2].metric("Match Rate", fmt_pct(ins_pct, 1))
                        summary_cols[3].metric("UMC", fmt_usd(ins_umc))

                        ins_disp = pd.DataFrame({
                            "Date":      ins_df["Date"].dt.strftime("%Y-%m-%d"),
                            "Deposit":   ins_df["Deposit"],
                            "Matched":   ins_df["Matched"],
                            "% Matched": ins_df["Pct_Matched"],
                            "UMC":       ins_df["UMC"].abs(),
                            "% UMC":     ins_df["Pct_UMC"],
                            "Worker":    ins_df["Worker"].apply(lambda w: worker_label(w, WORKER_NAMES)),
                        })
                        styled_ins = (
                            ins_disp.style
                            .format({"Deposit": "${:,.2f}", "Matched": "${:,.2f}", "UMC": "${:,.2f}",
                                     "% Matched": "{:.2f}%", "% UMC": "{:.4f}%"})
                            .hide(axis="index")
                        )
                        st.dataframe(styled_ins, use_container_width=True)

                        if len(insurers) > 1:
                            st.markdown("---")
    else:
        st.info("Type a search term above to find a member.")


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Monthly Cash Metrics
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:

    if metrics.empty:
        st.warning("No monthly Cash Metrics data found for January 2025 onwards.")
        st.stop()

    # Merge with UMC backlog
    combined = metrics.merge(
        umc_bl[["Month", "UMC_Backlog"]], on="Month", how="left"
    )
    combined["UMC_Backlog"] = combined["UMC_Backlog"].fillna(0)
    rcv2 = combined["Cash_Received"].replace(0, np.nan)
    combined["Pct_Backlog"] = (combined["UMC_Backlog"] / rcv2 * 100).round(4)

    # KPIs (aggregate over full filtered period)
    total_recv   = combined["Cash_Received"].sum()
    total_rec    = combined["Total_Reconciled"].sum()
    total_umc_m  = combined["UMC_Monthly"].abs().sum()
    latest_bl    = combined["UMC_Backlog"].iloc[-1] if len(combined) else 0
    pct_rec_ov   = total_rec / total_recv * 100 if total_recv else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Received (2025+)", fmt_m(total_recv))
    k2.metric("Total Reconciled",       fmt_m(total_rec), f"{pct_rec_ov:.1f}% of received")
    k3.metric("Total Monthly UMC",      fmt_m(total_umc_m))
    k4.metric("Latest UMC Backlog",     fmt_m(latest_bl),
              help="Total outstanding unallocated cash as of last closed month")

    st.divider()

    # ── Charts ────────────────────────────────────────────────────────────────
    st.subheader("Monthly Cash Flow")

    chart_type = st.radio("View as:", ["Grouped Bars", "Lines"], horizontal=True)

    series = [
        ("Cash Received",    "Cash_Received",    "#0066cc"),
        ("Reconciled",       "Total_Reconciled", "#28a745"),
        ("Monthly UMC",      "UMC_Monthly",      "#dc3545"),
        ("UMC Backlog",      "UMC_Backlog",      "#fd7e14"),
    ]

    fig_m = go.Figure()
    for name, col, color in series:
        y = combined[col]
        if chart_type == "Grouped Bars":
            fig_m.add_trace(go.Bar(x=combined["Month_Label"], y=y, name=name, marker_color=color))
        else:
            fig_m.add_trace(go.Scatter(
                x=combined["Month_Label"], y=y, name=name, mode="lines+markers",
                line=dict(color=color, width=2),
            ))

    fig_m.update_layout(
        barmode="group" if chart_type == "Grouped Bars" else None,
        title="Monthly Cash Flow",
        xaxis_title="Month",
        yaxis=dict(tickformat="$,.0f", title="Amount (USD)"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        height=420, margin=dict(t=60, b=20),
    )
    st.plotly_chart(fig_m, use_container_width=True)

    # Reconciliation rate
    fig_r = go.Figure()
    fig_r.add_trace(go.Scatter(
        x=combined["Month_Label"], y=combined["Pct_Reconciled"],
        name="% Reconciled", mode="lines+markers",
        line=dict(color="#28a745", width=2),
        fill="tozeroy", fillcolor="rgba(40,167,69,0.08)",
    ))
    fig_r.add_hline(y=100, line_dash="dash", line_color="gray",
                    annotation_text="100%", annotation_position="top right")
    fig_r.update_layout(
        title="Monthly Reconciliation Rate",
        xaxis_title="Month",
        yaxis=dict(tickformat=".1f", ticksuffix="%", title="% of Cash Received"),
        height=300, margin=dict(t=60, b=20),
    )
    st.plotly_chart(fig_r, use_container_width=True)

    # ── Monthly Data Table ────────────────────────────────────────────────────
    st.subheader("Monthly Data Table")

    tbl = pd.DataFrame({
        "Month":           combined["Month_Label"],
        "Cash Received":   combined["Cash_Received"],
        "Reconciled":      combined["Total_Reconciled"],
        "% Reconciled":    combined["Pct_Reconciled"],
        "AutoRek":         combined["AutoRek"],
        "Monthly UMC":     combined["UMC_Monthly"].abs(),
        "% UMC (monthly)": combined["Pct_UMC"],
        "Total UMC Backlog": combined["UMC_Backlog"],
        "% Backlog / Recv":  combined["Pct_Backlog"],
    })
    styled_tbl = (
        tbl.style
        .format({
            "Cash Received":   "${:,.2f}",
            "Reconciled":      "${:,.2f}",
            "% Reconciled":    "{:.2f}%",
            "AutoRek":         "${:,.2f}",
            "Monthly UMC":     "${:,.2f}",
            "% UMC (monthly)": "{:.4f}%",
            "Total UMC Backlog": "${:,.2f}",
            "% Backlog / Recv":  "{:.4f}%",
        })
        .hide(axis="index")
    )
    st.dataframe(styled_tbl, use_container_width=True)

    csv_m = combined.drop(columns=["Month_Raw"], errors="ignore").to_csv(index=False).encode()
    st.download_button("⬇️ Monthly Metrics CSV", csv_m, "monthly_metrics.csv", "text/csv", key="dl_metrics")

# ─── Footer ───────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "Jan 2025 → present · Matched = sum Y–AG · UMC = col P (abs value) · "
    "Monthly metrics: Cash Metrics O2:S40 · UMC Backlog: O49:P87"
)
