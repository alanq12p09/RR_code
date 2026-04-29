import re
import pandas as pd
import numpy as np
from presto_conn import query_df


# =========================
# 可配置参数
# =========================
INCLUDE_CURRENT_MONTH_IN_RECOMMENDATION = False
MIN_RECENT_MONTHS_WITH_RR = 3          # 最近4个月中至少几个月有RR（1~4）
AP_SUBGEO_MIN_RECOMMENDED_RR = 10      # AP按SubGeo分配后，低于此值的行不输出


def get_rr_data():
    sql = """
SELECT
    substr(rsd, 1, 7) AS RSD_Month,
    geo AS Geo,
    CASE 
        WHEN sub_geo = 'GTAP' THEN region_new
        WHEN sub_geo = 'HTK' THEN 'CAPK'
        WHEN sub_geo = 'ASEAN' THEN 'CAPK'
        WHEN sub_geo = 'LJP' THEN 'JAPAN'
        WHEN sub_geo IN ('IN','INDO') THEN 'INDIA'
        ELSE sub_geo
    END AS SubGeo,
    mtm AS PN,
    SUM(line_qty) AS Qty
FROM prd_sc_common_model.v_pcsd_order_detail_cml_brand_option
WHERE
    substr(rsd, 1, 7) >= DATE_FORMAT(DATE_ADD('month', -12, DATE_TRUNC('quarter', NOW())), '%Y-%m')                     
    AND substr(rsd, 1, 7) <= DATE_FORMAT(NOW(), '%Y-%m')    
    AND substr(rsd, 1, 4) <> '9999'
    AND shipping_point = 'SC03'
    AND brand = 'Option'
    AND order_type IN ('Customer Order', 'Replenishment', 'Free of Charge', 'Internal Order')
    AND site_name = 'LSSC'
    AND status <> 'Order Cancelled'
    AND product_group IN ('TBG', 'LBG')
    AND (so_route <> 'CXLSSC' OR so_route IS NULL)
    AND (end_customer_name <> 'Lenovo Information Products(Shenzhen)Co.' OR end_customer_name IS NULL)
GROUP BY
    substr(rsd, 1, 7),
    geo,
    CASE 
        WHEN sub_geo = 'GTAP' THEN region_new
        WHEN sub_geo = 'HTK' THEN 'CAPK'
        WHEN sub_geo = 'ASEAN' THEN 'CAPK'
        WHEN sub_geo = 'LJP' THEN 'JAPAN'
        WHEN sub_geo IN ('IN','INDO') THEN 'INDIA'
        ELSE sub_geo
    END,
    mtm
"""
    return query_df(sql)


def load_pn_info(pn_info_path):
    df = pd.read_excel(pn_info_path)
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {}
    for c in df.columns:
        cu = c.upper()
        if cu in ["PN", "MTM", "PART NUMBER", "MATERIAL"]:
            rename_map[c] = "PN"
        elif cu in ["COMMENTS", "COMMENT", "REMARK", "REMARKS"]:
            rename_map[c] = "Comments"
        elif cu in ["CATEGORY1", "CATEGORY 1", "BU"]:
            rename_map[c] = "Category1"

    df = df.rename(columns=rename_map)

    required = ["PN", "Comments", "Category1"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"PN_INFO 缺少必要列: {missing}")

    df["PN"] = df["PN"].astype(str).str.strip()
    df["Comments"] = df["Comments"].fillna("").astype(str)
    df["Category1"] = df["Category1"].fillna("").astype(str).str.upper().str.strip()
    return df


def prepare_rr(rr_raw):
    rr = rr_raw.copy()
    rr.columns = [str(c).strip() for c in rr.columns]

    required = ["RSD_Month", "Geo", "SubGeo", "PN", "Qty"]
    missing = [c for c in required if c not in rr.columns]
    if missing:
        raise ValueError(f"RR 数据缺少必要列: {missing}")

    rr["RSD_Month"] = pd.to_datetime(rr["RSD_Month"].astype(str).str[:7] + "-01", errors="coerce")
    rr["Geo"] = rr["Geo"].astype(str).str.strip()
    rr["SubGeo"] = rr["SubGeo"].astype(str).str.strip()
    rr["PN"] = rr["PN"].astype(str).str.strip()
    rr["Qty"] = pd.to_numeric(rr["Qty"], errors="coerce").fillna(0)

    rr = rr.dropna(subset=["RSD_Month"])
    return rr


def filter_pn_info(pn_info):
    df = pn_info.copy()

    pattern = re.compile(r"(EOL|LTB|LTS)", flags=re.IGNORECASE)
    df["Exclude_By_Comments"] = df["Comments"].apply(lambda x: bool(pattern.search(str(x))))
    df["Exclude_By_DOCK"] = df["Category1"].eq("DOCK")
    df["Valid_PN"] = (~df["Exclude_By_Comments"]) & (~df["Exclude_By_DOCK"])

    valid_pn = df.loc[df["Valid_PN"], ["PN"]].drop_duplicates()
    return df, valid_pn


def get_time_windows(rr):
    """
    基准月份改为系统当前月份，不再依赖数据最大月份。
    同时校验数据是否严重滞后，如果滞后超过1个月则打印警告。
    """
    current_month = pd.Timestamp.now().to_period("M").to_timestamp()
    current_quarter = current_month.to_period("Q")

    # 校验数据时效性
    data_max_month = rr["RSD_Month"].max()
    if data_max_month < (current_month.to_period("M") - 1).to_timestamp():
        print(
            f"⚠️  警告: 数据最新月份为 {data_max_month:%Y-%m}，"
            f"落后于系统当前月份 {current_month:%Y-%m} 超过1个月，请检查数据源是否正常。"
        )

    recent_4m = [(current_month.to_period("M") - i).to_timestamp() for i in range(4)]
    recent_4m = sorted(recent_4m)

    cutoff_month_6_ago = (current_month.to_period("M") - 6).to_timestamp()
    prev_4_quarters = [current_quarter - i for i in range(1, 5)]

    return {
        "current_month": current_month,
        "current_quarter": current_quarter,
        "recent_4m": recent_4m,
        "cutoff_month_6_ago": cutoff_month_6_ago,
        "prev_4_quarters": prev_4_quarters
    }


def build_geo_pn_summary(rr, valid_pn, time_windows, min_recent_months=3):
    rr2 = rr.merge(valid_pn, on="PN", how="inner").copy()
    rr2["Quarter"] = rr2["RSD_Month"].dt.to_period("Q")

    recent_4m = set(time_windows["recent_4m"])
    cutoff_month_6_ago = time_windows["cutoff_month_6_ago"]
    prev_4_quarters = time_windows["prev_4_quarters"]

    geo_pn_month = rr2.groupby(["PN", "Geo", "RSD_Month"], as_index=False)["Qty"].sum()
    geo_pn_qtr = rr2.groupby(["PN", "Geo", "Quarter"], as_index=False)["Qty"].sum()

    pn_total = (
        rr2.groupby("PN", as_index=False)["Qty"].sum()
        .rename(columns={"Qty": "PN_Total_Qty_AllGeo"})
    )
    geo_total = (
        rr2.groupby(["PN", "Geo"], as_index=False)["Qty"].sum()
        .rename(columns={"Qty": "PN_Geo_Total_Qty"})
    )

    summary = geo_total.merge(pn_total, on="PN", how="left")
    summary["Geo_Share_of_PN"] = np.where(
        summary["PN_Total_Qty_AllGeo"] > 0,
        summary["PN_Geo_Total_Qty"] / summary["PN_Total_Qty_AllGeo"],
        0
    )

    # 条件1：过去4个季度（不含当前季度）平均RR < 500
    for q in prev_4_quarters:
        temp = geo_pn_qtr.loc[geo_pn_qtr["Quarter"] == q, ["PN", "Geo", "Qty"]].copy()
        temp.rename(columns={"Qty": f"Qty_{str(q)}"}, inplace=True)
        summary = summary.merge(temp, on=["PN", "Geo"], how="left")

    q_cols = [c for c in summary.columns if c.startswith("Qty_20") and "Q" in c]
    for c in q_cols:
        summary[c] = summary[c].fillna(0)

    summary["Avg_Last4Q_Excl_CQ"] = summary[q_cols].mean(axis=1) if q_cols else 0
    summary["Flag_Avg_Last4Q_LT_500"] = summary["Avg_Last4Q_Excl_CQ"] < 500

    # 条件2：最近4个月中至少 min_recent_months 个月有RR（可配置）
    recent_rr_detail = geo_pn_month.loc[
        geo_pn_month["RSD_Month"].isin(recent_4m)
    ].copy()

    recent_rr_check = (
        recent_rr_detail.groupby(["PN", "Geo"], as_index=False)
        .agg(
            Qty_Recent4M=("Qty", "sum"),
            Recent4M_Month_Count=("RSD_Month", "nunique")
        )
    )

    summary = summary.merge(recent_rr_check, on=["PN", "Geo"], how="left")
    summary["Qty_Recent4M"] = summary["Qty_Recent4M"].fillna(0)
    summary["Recent4M_Month_Count"] = summary["Recent4M_Month_Count"].fillna(0).astype(int)
    summary["Flag_Recent4M_HasRR"] = summary["Recent4M_Month_Count"] >= min_recent_months

    # 条件3：6个月前及更早，在当前数据集任意月份有RR
    older_rr = (
        geo_pn_month.loc[geo_pn_month["RSD_Month"] <= cutoff_month_6_ago]
        .groupby(["PN", "Geo"], as_index=False)["Qty"].sum()
        .rename(columns={"Qty": "Qty_6M_Ago_And_Earlier"})
    )
    summary = summary.merge(older_rr, on=["PN", "Geo"], how="left")
    summary["Qty_6M_Ago_And_Earlier"] = summary["Qty_6M_Ago_And_Earlier"].fillna(0)
    summary["Flag_6M_Ago_And_Earlier_HasRR"] = summary["Qty_6M_Ago_And_Earlier"] > 0

    # 条件4：PN总量>500，但某GEO占比<3%，排除该GEO
    summary["Exclude_Geo_LowShare_When_PNHigh"] = (
        (summary["PN_Total_Qty_AllGeo"] > 500) &
        (summary["Geo_Share_of_PN"] < 0.03)
    )

    summary["Qualified"] = (
        summary["Flag_Avg_Last4Q_LT_500"] &
        summary["Flag_Recent4M_HasRR"] &
        summary["Flag_6M_Ago_And_Earlier_HasRR"] &
        (~summary["Exclude_Geo_LowShare_When_PNHigh"])
    )

    return summary, rr2


def trimmed_mean(series):
    vals = pd.to_numeric(series, errors="coerce").dropna().tolist()
    vals = sorted(vals)

    if len(vals) == 0:
        return np.nan
    if len(vals) == 1:
        return vals[0]
    if len(vals) == 2:
        return np.mean(vals)
    return np.mean(vals[1:-1])


def get_recommendation_base_data(rr_filtered, summary, current_month, include_current_month=False):
    qualified_keys = summary.loc[summary["Qualified"], ["PN", "Geo"]].drop_duplicates()
    base = rr_filtered.merge(qualified_keys, on=["PN", "Geo"], how="inner").copy()

    if not include_current_month:
        base = base.loc[base["RSD_Month"] < current_month].copy()

    return base


def build_recommendation(rr_filtered, summary, time_windows,
                         include_current_month=False,
                         ap_subgeo_min_rr=10):
    """
    非AP：按 PN + GEO 直接算 recommendation
    AP：先算 PN + AP 的 recommendation，再按 SubGeo 历史占比分配
         分配后低于 ap_subgeo_min_rr 的行会被过滤掉
    """
    current_month = time_windows["current_month"]
    base = get_recommendation_base_data(
        rr_filtered=rr_filtered,
        summary=summary,
        current_month=current_month,
        include_current_month=include_current_month
    )

    # =========================
    # 非 AP：按 PN + GEO 直接算
    # =========================
    non_ap = base.loc[base["Geo"] != "AP"].copy()
    non_ap_month = (
        non_ap.groupby(["PN", "Geo", "RSD_Month"], as_index=False)["Qty"].sum()
    )

    non_ap_rec = (
        non_ap_month.groupby(["PN", "Geo"])["Qty"]
        .apply(trimmed_mean)
        .reset_index(name="Recommended_RR")
    )
    non_ap_rec["SubGeo"] = ""
    non_ap_rec["Recommend_Level"] = "GEO"
    non_ap_rec["Recommend_To"] = non_ap_rec["Geo"]

    non_ap_stats = (
        non_ap_month.groupby(["PN", "Geo"])["Qty"]
        .agg(Month_Count="count", Hist_Min="min", Hist_Max="max", Hist_Avg="mean")
        .reset_index()
    )
    non_ap_stats["SubGeo"] = ""

    non_ap_result = non_ap_rec.merge(
        non_ap_stats[["PN", "Geo", "SubGeo", "Month_Count", "Hist_Min", "Hist_Max", "Hist_Avg"]],
        on=["PN", "Geo", "SubGeo"],
        how="left"
    )

    # =========================
    # AP：先算 PN + AP recommendation
    # 再按 SubGeo 历史占比分配
    # =========================
    ap = base.loc[base["Geo"] == "AP"].copy()

    # AP总体月度
    ap_month_total = (
        ap.groupby(["PN", "Geo", "RSD_Month"], as_index=False)["Qty"].sum()
    )

    ap_geo_rec = (
        ap_month_total.groupby(["PN", "Geo"])["Qty"]
        .apply(trimmed_mean)
        .reset_index(name="AP_Geo_Recommended_RR")
    )

    ap_geo_stats = (
        ap_month_total.groupby(["PN", "Geo"])["Qty"]
        .agg(AP_Month_Count="count", AP_Hist_Min="min", AP_Hist_Max="max", AP_Hist_Avg="mean")
        .reset_index()
    )

    # SubGeo 历史占比：基于 recommendation 使用的同一份历史窗口
    ap_subgeo_hist = (
        ap.groupby(["PN", "Geo", "SubGeo"], as_index=False)["Qty"]
        .sum()
        .rename(columns={"Qty": "SubGeo_Hist_Qty"})
    )

    ap_pn_hist = (
        ap.groupby(["PN", "Geo"], as_index=False)["Qty"]
        .sum()
        .rename(columns={"Qty": "AP_Hist_Qty"})
    )

    ap_alloc = ap_subgeo_hist.merge(ap_pn_hist, on=["PN", "Geo"], how="left")
    ap_alloc["SubGeo_Share"] = np.where(
        ap_alloc["AP_Hist_Qty"] > 0,
        ap_alloc["SubGeo_Hist_Qty"] / ap_alloc["AP_Hist_Qty"],
        0
    )

    ap_result = (
        ap_alloc
        .merge(ap_geo_rec, on=["PN", "Geo"], how="left")
        .merge(ap_geo_stats, on=["PN", "Geo"], how="left")
    )

    ap_result["Recommended_RR"] = ap_result["AP_Geo_Recommended_RR"] * ap_result["SubGeo_Share"]
    ap_result["Recommend_Level"] = "SubGeo_Allocated_From_AP"
    ap_result["Recommend_To"] = ap_result["SubGeo"]
    ap_result["Month_Count"] = ap_result["AP_Month_Count"]
    ap_result["Hist_Min"] = ap_result["AP_Hist_Min"]
    ap_result["Hist_Max"] = ap_result["AP_Hist_Max"]
    ap_result["Hist_Avg"] = ap_result["AP_Hist_Avg"]

    # 过滤掉分配后推荐值过低的SubGeo行
    ap_before_filter = len(ap_result)
    ap_result = ap_result.loc[ap_result["Recommended_RR"] >= ap_subgeo_min_rr].copy()
    ap_filtered_count = ap_before_filter - len(ap_result)
    if ap_filtered_count > 0:
        print(
            f"ℹ️  AP SubGeo 分配后过滤: {ap_filtered_count} 行推荐值 < {ap_subgeo_min_rr}，已移除"
        )

    ap_result = ap_result[[
        "PN", "Geo", "SubGeo", "Recommend_Level", "Recommend_To",
        "Recommended_RR", "Month_Count", "Hist_Min", "Hist_Max", "Hist_Avg",
        "SubGeo_Hist_Qty", "AP_Hist_Qty", "SubGeo_Share", "AP_Geo_Recommended_RR"
    ]]

    non_ap_result["SubGeo_Hist_Qty"] = np.nan
    non_ap_result["AP_Hist_Qty"] = np.nan
    non_ap_result["SubGeo_Share"] = np.nan
    non_ap_result["AP_Geo_Recommended_RR"] = np.nan

    rec = pd.concat([
        non_ap_result[[
            "PN", "Geo", "SubGeo", "Recommend_Level", "Recommend_To",
            "Recommended_RR", "Month_Count", "Hist_Min", "Hist_Max", "Hist_Avg",
            "SubGeo_Hist_Qty", "AP_Hist_Qty", "SubGeo_Share", "AP_Geo_Recommended_RR"
        ]],
        ap_result
    ], ignore_index=True)

    rec["Recommended_RR"] = rec["Recommended_RR"].round(2)
    rec["Hist_Avg"] = rec["Hist_Avg"].round(2)
    rec["Hist_Min"] = rec["Hist_Min"].round(2)
    rec["Hist_Max"] = rec["Hist_Max"].round(2)
    rec["SubGeo_Share"] = rec["SubGeo_Share"].round(4)
    rec["AP_Geo_Recommended_RR"] = rec["AP_Geo_Recommended_RR"].round(2)

    return rec.sort_values(["PN", "Geo", "SubGeo"]).reset_index(drop=True)


def build_history_reference(rr_filtered, summary):
    qualified_keys = summary.loc[summary["Qualified"], ["PN", "Geo"]].drop_duplicates()
    hist = rr_filtered.merge(qualified_keys, on=["PN", "Geo"], how="inner")
    hist = (
        hist.groupby(["PN", "Geo", "SubGeo", "RSD_Month"], as_index=False)["Qty"]
        .sum()
        .sort_values(["PN", "Geo", "SubGeo", "RSD_Month"])
    )
    hist["RSD_Month"] = hist["RSD_Month"].dt.strftime("%Y-%m")
    return hist


def export_results(output_path, summary, history, recommendation, config_df):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        config_df.to_excel(writer, sheet_name="Config", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        summary.loc[summary["Qualified"]].to_excel(writer, sheet_name="Qualified", index=False)
        history.to_excel(writer, sheet_name="History_RR", index=False)
        recommendation.to_excel(writer, sheet_name="Recommendation", index=False)


def main():
    pn_info_path = r"C:\Users\xuecz1\Lenovo\Option Share Files - DeptDocument\Power BI Group\PN Information.xlsx"
    output_path = r"D:\Project\RR_Forecast\Data\RR_Low_RR_Recommendation.xlsx"

    rr_raw = get_rr_data()
    rr = prepare_rr(rr_raw)

    pn_info = load_pn_info(pn_info_path)
    _, valid_pn = filter_pn_info(pn_info)

    time_windows = get_time_windows(rr)

    summary, rr_filtered = build_geo_pn_summary(
        rr, valid_pn, time_windows,
        min_recent_months=MIN_RECENT_MONTHS_WITH_RR
    )
    history = build_history_reference(rr_filtered, summary)

    recommendation = build_recommendation(
        rr_filtered=rr_filtered,
        summary=summary,
        time_windows=time_windows,
        include_current_month=INCLUDE_CURRENT_MONTH_IN_RECOMMENDATION,
        ap_subgeo_min_rr=AP_SUBGEO_MIN_RECOMMENDED_RR
    )

    config_df = pd.DataFrame({
        "Parameter": [
            "INCLUDE_CURRENT_MONTH_IN_RECOMMENDATION",
            "MIN_RECENT_MONTHS_WITH_RR",
            "AP_SUBGEO_MIN_RECOMMENDED_RR",
            "Current_Month (System Time)",
            "Data_Max_Month",
            "Recent_4M_Rule",
            "AP_Recommendation_Method"
        ],
        "Value": [
            INCLUDE_CURRENT_MONTH_IN_RECOMMENDATION,
            MIN_RECENT_MONTHS_WITH_RR,
            AP_SUBGEO_MIN_RECOMMENDED_RR,
            time_windows["current_month"].strftime("%Y-%m"),
            rr["RSD_Month"].max().strftime("%Y-%m"),
            f"最近4个月中至少 {MIN_RECENT_MONTHS_WITH_RR} 个月有RR",
            "先算AP的Geo recommendation，再按SubGeo历史占比分配，"
            f"分配后低于 {AP_SUBGEO_MIN_RECOMMENDED_RR} 的SubGeo行被过滤"
        ]
    })

    export_results(output_path, summary, history, recommendation, config_df)

    print("=" * 100)
    print(f"结果已保存到: {output_path}")
    print(f"Current Month (System Time): {time_windows['current_month'].strftime('%Y-%m')}")
    print(f"Data Max Month: {rr['RSD_Month'].max().strftime('%Y-%m')}")
    print(f"Include Current Month In Recommendation: {INCLUDE_CURRENT_MONTH_IN_RECOMMENDATION}")
    print(f"Min Recent Months With RR: {MIN_RECENT_MONTHS_WITH_RR}")
    print(f"AP SubGeo Min Recommended RR: {AP_SUBGEO_MIN_RECOMMENDED_RR}")
    print(f"Qualified GEO+PN 数量: {summary['Qualified'].sum():,}")
    print(f"Recommendation 行数: {len(recommendation):,}")
    print("=" * 100)

    if recommendation.empty:
        print("没有可输出的建议值。")
    else:
        print(recommendation.head(30).to_string(index=False))


if __name__ == "__main__":
    main()
