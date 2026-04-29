from presto_conn import query_df

def main():
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
    --mtm_desc AS Description,
    SUM(line_qty) AS Qty
FROM prd_sc_common_model.v_pcsd_order_detail_cml_brand_option
WHERE
    substr(rsd, 1, 7) >= DATE_FORMAT(DATE_ADD('month', -15, DATE_TRUNC('quarter', NOW())), '%Y-%m')
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
    df = query_df(sql)
    print(df)
    # Save to Excel on Desktop
    import os
    # desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    output_path = os.path.join("D:\\Project\\RR_Forecast\\Data", "RR_C2.xlsx")
    df.to_excel(output_path, index=False)
    print(f"结果已保存到: {output_path}")

if __name__ == "__main__":
    main()

    # Save to Excel on Desktop