"""
pages/8_Business_Performance.py — Hierarchical Business Performance MIS Report

Displays data in a management report format mirroring the AOP Excel structure.
Each location (Delhi / Hyderabad / Goa) is shown with:
  - Individual outlet rows
  - Subtotal rows (bold, indented label)
  - TOTAL EHPL grand total row

Columns: Performance | Rev (Current) | Rev (Compare) | Rev YOY% |
         PAX (Current) | PAX (Compare) | PAX YOY% |
         Traffic (Cur) | Traffic (Cmp) | Traffic Δ% |
         PEN % (Cur) | PEN % (Cmp) | PEN Δ% |
         SPP (Cur) | SPP (Cmp) | SPP Δ% |
         AOP Target | AOP Var %
"""

from __future__ import annotations

import pandas as pd
import streamlit as st
from modules import comparison_widget, database, date_picker, revenue_analysis as ra, table_style
from modules.formatting import format_money, format_pax, format_pct, format_spp
from modules.outlet_groups import (
    DELHI_GROUPS, DELHI_SUBTOTALS, HYD_GROUPS, HYD_SUBTOTALS, GOA_GROUPS, GOA_SUBTOTALS,
    get_display_name,
)
from modules.session import bootstrap_session, default_active_date, set_active_date
from modules.app_logger import safe_run, log_exception, show_friendly_error

st.set_page_config(page_title="Business Performance", page_icon="📊", layout="wide")
bootstrap_session()

st.title("📊 Business Performance Report")
st.caption(
    "Grouped summary mirroring the AOP management report structure. "
    "Penetration % = PAX ÷ Terminal Traffic. "
    "SPP = Revenue ÷ Terminal Traffic (not Revenue ÷ PAX)."
)

available_dates = database.get_available_dates()
if not available_dates:
    st.info("No data available. Upload a report on the main page first.")
    st.stop()

anchor_date = date_picker.render_date_dropdown(
    available_dates, key_prefix="bp_anchor", label="Report Date",
    default_date=default_active_date(),
)
set_active_date(anchor_date)

ranges = comparison_widget.render_comparison_selector(anchor_date, key_prefix="bp_cmp")
current_df = database.load_for_date_range(ranges["current_start"], ranges["current_end"])
compare_df = database.load_for_date_range(ranges["compare_start"], ranges["compare_end"])

if current_df.empty:
    st.warning("No revenue data for the selected period.")
    st.stop()

current_label = ranges["current_label"]
compare_label = ranges["compare_label"]
aop_df = database.load_aop_targets_for_range(ranges["current_start"], ranges["current_end"])


# ---------------------------------------------------------------------------
# Core data builder
# ---------------------------------------------------------------------------

def _build_group_summary(
    df: pd.DataFrame,
    groups: dict[str, list[str]],
    subtotals: list[tuple[str, list[str]]],
    aop_df: pd.DataFrame | None = None,
    compare_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Build a flat DataFrame of outlet rows + subtotal rows in group order.
    Each row has: label, indent_level, is_subtotal, is_total, + metric cols.
    """
    rows = []

    cur_agg = df.groupby(["outlet", "location"], as_index=False).agg(
        revenue=("revenue", "sum"), pax=("pax", "sum")
    )
    cmp_agg = (
        compare_df.groupby(["outlet", "location"], as_index=False).agg(
            revenue=("revenue", "sum"), pax=("pax", "sum")
        )
        if compare_df is not None and not compare_df.empty
        else pd.DataFrame(columns=["outlet", "location", "revenue", "pax"])
    )

    # Per-outlet terminal traffic
    outlet_traffic_cur = database.join_revenue_with_traffic_by_outlet(df) \
        if df is not None and not df.empty else None
    outlet_traffic_cmp = database.join_revenue_with_traffic_by_outlet(compare_df) \
        if compare_df is not None and not compare_df.empty else None

    traffic_map: dict = {}
    compare_traffic_map: dict = {}
    if outlet_traffic_cur is not None and not outlet_traffic_cur.empty:
        for _, r in outlet_traffic_cur[outlet_traffic_cur["traffic"].notna()].iterrows():
            raw_key = (r["outlet"], r["location"])
            disp_key = (get_display_name(r["outlet"], r.get("location", "")), r["location"])
            traffic_map[raw_key]  = r["traffic"]
            traffic_map[disp_key] = r["traffic"]
    if outlet_traffic_cmp is not None and not outlet_traffic_cmp.empty:
        for _, r in outlet_traffic_cmp[outlet_traffic_cmp["traffic"].notna()].iterrows():
            raw_key = (r["outlet"], r["location"])
            disp_key = (get_display_name(r["outlet"], r.get("location", "")), r["location"])
            compare_traffic_map[raw_key]  = r["traffic"]
            compare_traffic_map[disp_key] = r["traffic"]

    # AOP lookup
    aop_map: dict = {}
    if aop_df is not None and not aop_df.empty and "aop" in aop_df.columns:
        aop_agg = aop_df.groupby("outlet", as_index=False)["aop"].sum()
        aop_map = aop_agg.set_index("outlet")["aop"].to_dict()

    def _outlet_row(outlet_name: str, group_label: str, indent: int = 1) -> dict | None:
        cur = cur_agg[cur_agg["outlet"] == outlet_name]
        if cur.empty:
            return None
        cur_rev = cur["revenue"].sum()
        cmp_check = cmp_agg[cmp_agg["outlet"] == outlet_name]
        cmp_rev_check = cmp_check["revenue"].sum() if not cmp_check.empty else 0
        if cur_rev == 0 and cmp_rev_check == 0:
            return None
        cur_pax = cur["pax"].sum()
        location = cur["location"].iloc[0]
        traffic = traffic_map.get((outlet_name, location))
        cmp = cmp_agg[cmp_agg["outlet"] == outlet_name]
        cmp_rev = cmp["revenue"].sum() if not cmp.empty else None
        cmp_pax = cmp["pax"].sum() if not cmp.empty else None
        cmp_traffic = compare_traffic_map.get((outlet_name, location))
        pen_cur = ra.safe_div(cur_pax, traffic) * 100 if traffic else None
        pen_cmp = ra.safe_div(cmp_pax, cmp_traffic) * 100 if cmp_traffic else None
        spp_cur = ra.safe_div(cur_rev, traffic) if traffic else None
        spp_cmp = ra.safe_div(cmp_rev, cmp_traffic) if cmp_traffic else None
        aop = aop_map.get(outlet_name)
        return {
            "label":        get_display_name(outlet_name, location),
            "indent":       indent,
            "is_subtotal":  False,
            "is_total":     False,
            "cur_rev": cur_rev,  "cmp_rev": cmp_rev,
            "rev_yoy": ra.pct_change(cur_rev, cmp_rev),
            "cur_pax": cur_pax,  "cmp_pax": cmp_pax,
            "pax_yoy": ra.pct_change(cur_pax, cmp_pax),
            "cur_traffic": traffic, "cmp_traffic": cmp_traffic,
            "traffic_chg": ra.pct_change(traffic, cmp_traffic),
            "pen_cur": pen_cur, "pen_cmp": pen_cmp,
            "pen_chg": ra.pct_change(pen_cur, pen_cmp),
            "spp_cur": spp_cur, "spp_cmp": spp_cmp,
            "spp_chg": ra.pct_change(spp_cur, spp_cmp),
            "aop": aop,
            "aop_var": ra.pct_change(cur_rev, aop) if aop else None,
        }

    group_totals: dict = {}
    subtotal_keys = {s[0] for s in subtotals}
    matched_outlets: set = set()

    for group_name, outlets in groups.items():
        if group_name in subtotal_keys:
            continue
        g_rows = []
        seen_display: set = set()
        _tab_loc = df["location"].iloc[0] if not df.empty else ""
        for outlet in outlets:
            dn = get_display_name(outlet, _tab_loc)
            if dn in seen_display:
                matched_outlets.add(outlet)
                continue
            r = _outlet_row(outlet, group_name, indent=1)
            if r:
                seen_display.add(dn)
                rows.append(r)
                g_rows.append(r)
                matched_outlets.add(outlet)
        if g_rows:
            group_totals[group_name] = {
                "cur_rev":     sum(r["cur_rev"] or 0 for r in g_rows),
                "cmp_rev":     sum(r["cmp_rev"] or 0 for r in g_rows if r["cmp_rev"]),
                "cur_pax":     sum(r["cur_pax"] or 0 for r in g_rows),
                "cmp_pax":     sum(r["cmp_pax"] or 0 for r in g_rows if r["cmp_pax"]),
                "aop":         sum(r["aop"] or 0 for r in g_rows if r["aop"]),
                # Traffic: all outlets in a group share the same terminal pool value
                # (e.g. all T3 Dom outlets get T3 Dom Dep traffic = 11,22,458).
                # Use the first non-None value — do NOT sum (that would multiply
                # the same pool value by the number of outlets).
                "cur_traffic": next((r["cur_traffic"] for r in g_rows if r["cur_traffic"]), None),
                "cmp_traffic": next((r["cmp_traffic"] for r in g_rows if r["cmp_traffic"]), None),
            }

    # Unmatched → Other
    all_data_outlets = set(cur_agg["outlet"].unique()) if not cur_agg.empty else set()
    unmatched = all_data_outlets - matched_outlets
    if unmatched:
        g_rows_other = []
        seen_other: set = set()
        _tab_loc2 = df["location"].iloc[0] if not df.empty else ""
        for outlet in sorted(unmatched):
            dn = get_display_name(outlet, _tab_loc2)
            if dn in seen_other:
                continue
            r = _outlet_row(outlet, "Other", indent=1)
            if r:
                seen_other.add(dn)
                rows.append(r)
                g_rows_other.append(r)

    # Subtotal rows
    for sub_label, source_groups in subtotals:
        sub_cur_rev = sum(group_totals.get(g, {}).get("cur_rev", 0) for g in source_groups)
        sub_cmp_rev = sum(group_totals.get(g, {}).get("cmp_rev", 0) for g in source_groups)
        sub_cur_pax = sum(group_totals.get(g, {}).get("cur_pax", 0) for g in source_groups)
        sub_cmp_pax = sum(group_totals.get(g, {}).get("cmp_pax", 0) for g in source_groups)
        sub_aop     = sum(group_totals.get(g, {}).get("aop", 0) for g in source_groups)
        # Traffic for subtotals — hybrid MAX/SUM logic matching Excel:
        #
        # Rule: collect UNIQUE traffic values across source groups, then:
        #   - If max value ≥ whole-airport (Atithya level) → use MAX
        #     (whole-airport subsumes all smaller pools: Porter ⊂ Atithya)
        #   - Otherwise → SUM unique values
        #     (non-overlapping pools genuinely add: T1 Dep + T3 Dom Dep)
        #
        # This correctly produces:
        #   Total T3 Dom      = 11,22,458     (5 groups, same pool → SUM dedups to 1)
        #   T1+T2+T3 Dom      = 22,95,375     (T1=11,72,917 + T3=11,22,458 → SUM)
        #   Atithya subtotal  = 61,48,937     (Porter ⊂ Atithya → MAX)
        #   TOTAL EHPL        = 61,48,937     (all pools ⊂ airport → MAX)
        def _sub_traffic(source_groups, traffic_key):
            vals = set(
                group_totals.get(g, {}).get(traffic_key)
                for g in source_groups
                if group_totals.get(g, {}).get(traffic_key)
            )
            if not vals:
                return None
            # Whole-airport traffic = Atithya (M&G) value — largest possible
            whole = group_totals.get("Atithya (M&G)", {}).get(traffic_key) or 0
            mx = max(vals)
            return mx if (whole and mx >= whole) else sum(vals)

        sub_cur_tr = _sub_traffic(source_groups, "cur_traffic")
        sub_cmp_tr = _sub_traffic(source_groups, "cmp_traffic")
        sub_pen_cur = ra.safe_div(sub_cur_pax, sub_cur_tr) * 100 if sub_cur_tr else None
        sub_pen_cmp = ra.safe_div(sub_cmp_pax, sub_cmp_tr) * 100 if sub_cmp_tr else None
        sub_spp_cur = ra.safe_div(sub_cur_rev, sub_cur_tr) if sub_cur_tr else None
        sub_spp_cmp = ra.safe_div(sub_cmp_rev, sub_cmp_tr) if sub_cmp_tr else None
        rows.append({
            "label": sub_label, "indent": 0, "is_subtotal": True, "is_total": False,
            "cur_rev": sub_cur_rev, "cmp_rev": sub_cmp_rev if sub_cmp_rev else None,
            "rev_yoy": ra.pct_change(sub_cur_rev, sub_cmp_rev),
            "cur_pax": sub_cur_pax, "cmp_pax": sub_cmp_pax if sub_cmp_pax else None,
            "pax_yoy": ra.pct_change(sub_cur_pax, sub_cmp_pax),
            "cur_traffic": sub_cur_tr, "cmp_traffic": sub_cmp_tr,
            "traffic_chg": ra.pct_change(sub_cur_tr, sub_cmp_tr),
            "pen_cur": sub_pen_cur, "pen_cmp": sub_pen_cmp,
            "pen_chg": ra.pct_change(sub_pen_cur, sub_pen_cmp),
            "spp_cur": sub_spp_cur, "spp_cmp": sub_spp_cmp,
            "spp_chg": ra.pct_change(sub_spp_cur, sub_spp_cmp),
            "aop": sub_aop if sub_aop else None,
            "aop_var": ra.pct_change(sub_cur_rev, sub_aop) if sub_aop else None,
        })

    # Sort rows by DELHI_GROUPS/HYD_GROUPS/GOA_GROUPS key order
    # Build position map: outlet display name or subtotal label → position
    order: dict[str, int] = {}
    pos = 0
    _loc = df["location"].iloc[0] if not df.empty else ""
    for group_name, outlets in groups.items():
        if group_name in subtotal_keys:
            order[group_name] = pos
            pos += 1
        else:
            seen: set = set()
            for outlet in outlets:
                dn = get_display_name(outlet, _loc)
                if dn not in seen:
                    order[dn] = pos
                    seen.add(dn)
                    pos += 1

    def _sort_key(r):
        key = r["label"]
        return order.get(key, pos + 1000)

    rows.sort(key=_sort_key)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Renderer — hierarchical MIS table
# ---------------------------------------------------------------------------

def _render_mis_table(
    summary_df: pd.DataFrame,
    cur_label: str,
    cmp_label: str,
    location: str,
) -> None:
    """Render the MIS table: one 'Performance' column + metric columns.

    Row types:
      - Outlet row   : normal weight, indented with thin left border via text prefix
      - Subtotal row : bold text
      - Total row    : bold text, uppercase
    """
    if summary_df.empty:
        st.info(f"No {location} data for this period.")
        return

    out = summary_df.copy()

    # ── Performance label column ──────────────────────────────────────────
    # Subtotals get no prefix; outlet rows get a thin space indent
    def _label(row):
        if row["is_total"]:
            return row["label"].upper()
        if row["is_subtotal"]:
            return row["label"]
        return "  " + row["label"]   # 2-space indent for outlet rows

    out["Performance"] = out.apply(_label, axis=1)

    # ── Metric columns ────────────────────────────────────────────────────
    out[f"Rev ({cur_label})"] = out["cur_rev"].apply(format_money)
    out[f"Rev ({cmp_label})"] = out["cmp_rev"].apply(lambda v: format_money(v) if v else "—")
    out["Rev YOY%"]           = out["rev_yoy"].apply(lambda v: format_pct(v) if v is not None and v == v else "—")
    out[f"PAX ({cur_label})"] = out["cur_pax"].apply(format_pax)
    out[f"PAX ({cmp_label})"] = out["cmp_pax"].apply(lambda v: format_pax(v) if v else "—")
    out["PAX YOY%"]           = out["pax_yoy"].apply(lambda v: format_pct(v) if v is not None and v == v else "—")

    display_cols = [
        "Performance",
        f"Rev ({cur_label})", f"Rev ({cmp_label})", "Rev YOY%",
        f"PAX ({cur_label})", f"PAX ({cmp_label})", "PAX YOY%",
    ]

    if out["cur_traffic"].notna().any():
        out[f"Traffic ({cur_label})"] = out["cur_traffic"].apply(lambda v: format_pax(v) if v else "—")
        out[f"Traffic ({cmp_label})"] = out["cmp_traffic"].apply(lambda v: format_pax(v) if v else "—")
        out["Traffic Δ%"]             = out["traffic_chg"].apply(lambda v: format_pct(v) if v is not None and v == v else "—")
        out[f"PEN % ({cur_label})"]   = out["pen_cur"].apply(lambda v: f"{v:.2f}%" if v is not None and v == v and v != 0 else "—")
        out[f"PEN % ({cmp_label})"]   = out["pen_cmp"].apply(lambda v: f"{v:.2f}%" if v is not None and v == v and v != 0 else "—")
        out["PEN Δ%"]                 = out["pen_chg"].apply(lambda v: format_pct(v) if v is not None and v == v else "—")
        out[f"SPP ({cur_label})"]     = out["spp_cur"].apply(lambda v: format_spp(v) if v is not None else "—")
        out[f"SPP ({cmp_label})"]     = out["spp_cmp"].apply(lambda v: format_spp(v) if v is not None else "—")
        out["SPP Δ%"]                 = out["spp_chg"].apply(lambda v: format_pct(v) if v is not None and v == v else "—")
        display_cols += [
            f"Traffic ({cur_label})", f"Traffic ({cmp_label})", "Traffic Δ%",
            f"PEN % ({cur_label})", f"PEN % ({cmp_label})", "PEN Δ%",
            f"SPP ({cur_label})", f"SPP ({cmp_label})", "SPP Δ%",
        ]

    if out["aop"].notna().any():
        out["AOP Target"] = out["aop"].apply(lambda v: format_money(v) if v else "—")
        out["AOP Var %"]  = out["aop_var"].apply(format_pct)
        display_cols += ["AOP Target", "AOP Var %"]

    pct_cols = [c for c in ["Rev YOY%", "PAX YOY%", "Traffic Δ%", "PEN Δ%", "SPP Δ%", "AOP Var %"]
                if c in out.columns]

    display_df = out[display_cols].reset_index(drop=True)
    is_subtotal = out["is_subtotal"].reset_index(drop=True)
    is_total    = out["is_total"].reset_index(drop=True)

    # ── Styling: bold for subtotals/totals, colour for pct cols ──────────
    BOLD     = "font-weight: bold"
    PCT_POS  = "color: #1E7F1E; font-weight: bold"
    PCT_NEG  = "color: #C00000; font-weight: bold"
    PCT_POS_B = "color: #1E7F1E; font-weight: bold"
    PCT_NEG_B = "color: #C00000; font-weight: bold"

    def _style(df):
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        for i in df.index:
            if is_total.iloc[i]:
                styles.iloc[i] = "font-weight: bold; text-transform: uppercase"
                styles.loc[i, "Performance"] = "font-weight: bold; text-transform: uppercase"
            elif is_subtotal.iloc[i]:
                styles.iloc[i] = BOLD
            for col in pct_cols:
                if col not in df.columns:
                    continue
                val = df[col].iloc[i]
                if isinstance(val, str) and val not in ("—", "-", ""):
                    try:
                        num = float(val.replace("%","").replace("+","").replace(",",""))
                        is_bold = is_subtotal.iloc[i] or is_total.iloc[i]
                        if num > 0:
                            styles.loc[i, col] = PCT_POS_B if is_bold else PCT_POS
                        elif num < 0:
                            styles.loc[i, col] = PCT_NEG_B if is_bold else PCT_NEG
                    except ValueError:
                        pass
        return styles

    styled = display_df.style.apply(_style, axis=None)

    col_cfg = {
        "Performance": st.column_config.TextColumn(
            "Performance", width="medium", pinned=True
        ),
    }

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        column_config=col_cfg,
        column_order=display_cols,
    )
    if out["cur_traffic"].notna().any():
        st.caption(
            "PEN % = PAX ÷ Terminal Traffic × 100.  "
            "SPP = Revenue ÷ Terminal Traffic.  "
            "Traffic is outlet's assigned terminal pool."
        )


# ---------------------------------------------------------------------------
# TOTAL EHPL row builder
# ---------------------------------------------------------------------------

def _total_ehpl_row(summary_df: pd.DataFrame, cur_label: str, cmp_label: str) -> None:
    """Render the TOTAL EHPL grand total as a highlighted metric row."""
    if summary_df.empty:
        return
    cur_rev = summary_df["cur_rev"].sum()
    cmp_rev = summary_df["cmp_rev"].dropna().sum()
    cur_pax = summary_df["cur_pax"].sum()
    cmp_pax = summary_df["cmp_pax"].dropna().sum()
    aop_tot = summary_df["aop"].dropna().sum()
    rev_yoy = ra.pct_change(cur_rev, cmp_rev)
    pax_yoy = ra.pct_change(cur_pax, cmp_pax)
    aop_var = ra.pct_change(cur_rev, aop_tot) if aop_tot else None

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("TOTAL EHPL — Rev (Current)",  format_money(cur_rev))
    c2.metric(f"Rev ({cmp_label})",           format_money(cmp_rev))
    c3.metric("Rev YOY%",                     format_pct(rev_yoy) if rev_yoy is not None else "—",
              delta=format_pct(rev_yoy) if rev_yoy is not None else None,
              delta_color="normal")
    c4.metric("PAX (Current)",                format_pax(cur_pax))
    if aop_var is not None:
        c5.metric("AOP Variance",             format_pct(aop_var),
                  delta=format_pct(aop_var), delta_color="normal")
    else:
        c5.metric("PAX YOY%",                 format_pct(pax_yoy) if pax_yoy is not None else "—")


# ---------------------------------------------------------------------------
# Location tabs
# ---------------------------------------------------------------------------

tab_delhi, tab_hyd, tab_goa = st.tabs(["🏙️ Delhi", "🏙️ Hyderabad", "🏖️ Goa"])

with tab_delhi:
    with safe_run("Delhi Business Performance", error_type="comparison_error"):
        df_cur = current_df[current_df["location"] == "Delhi"]
        df_cmp = compare_df[compare_df["location"] == "Delhi"] if not compare_df.empty else None
        df_aop = aop_df[aop_df["location"] == "Delhi"] \
            if aop_df is not None and not aop_df.empty else None

        if df_cur.empty:
            st.info("No Delhi data for this period.")
        else:
            st.subheader("🏙️ Delhi")
            summary = _build_group_summary(
                df_cur, DELHI_GROUPS, DELHI_SUBTOTALS,
                aop_df=df_aop, compare_df=df_cmp,
            )
            _render_mis_table(summary, current_label, compare_label, "Delhi")
            st.divider()
            _total_ehpl_row(summary, current_label, compare_label)

with tab_hyd:
    with safe_run("Hyderabad Business Performance", error_type="comparison_error"):
        df_cur = current_df[current_df["location"] == "Hyderabad"]
        df_cmp = compare_df[compare_df["location"] == "Hyderabad"] if not compare_df.empty else None
        df_aop = aop_df[aop_df["location"] == "Hyderabad"] \
            if aop_df is not None and not aop_df.empty else None

        if df_cur.empty:
            st.info("No Hyderabad data for this period.")
        else:
            st.subheader("🏙️ Hyderabad")
            summary = _build_group_summary(
                df_cur, HYD_GROUPS, HYD_SUBTOTALS,
                aop_df=df_aop, compare_df=df_cmp,
            )
            _render_mis_table(summary, current_label, compare_label, "Hyderabad")
            st.divider()
            _total_ehpl_row(summary, current_label, compare_label)

with tab_goa:
    with safe_run("Goa Business Performance", error_type="comparison_error"):
        df_cur = current_df[current_df["location"] == "Goa"]
        df_cmp = compare_df[compare_df["location"] == "Goa"] if not compare_df.empty else None
        df_aop = aop_df[aop_df["location"] == "Goa"] \
            if aop_df is not None and not aop_df.empty else None

        if df_cur.empty:
            st.info("No Goa data for this period.")
        else:
            st.subheader("🏖️ Goa")
            summary = _build_group_summary(
                df_cur, GOA_GROUPS, GOA_SUBTOTALS,
                aop_df=df_aop, compare_df=df_cmp,
            )
            _render_mis_table(summary, current_label, compare_label, "Goa")
            st.divider()
            _total_ehpl_row(summary, current_label, compare_label)
