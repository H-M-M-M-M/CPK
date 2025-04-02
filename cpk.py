import streamlit as st
import pandas as pd
import numpy as np

st.title("探头测试数据 CPK 计算 (按季度统计)")

uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("选择工作表", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    st.write("Excel 列名:", list(df.columns))
    
    type_col = st.selectbox("选择型号列", df.columns)
    sn_col = st.selectbox("选择探头 SN 列", df.columns)
    date_col = st.selectbox("选择日期列", df.columns)
    time_col = st.selectbox("选择时间列", df.columns)
    test_value_col = st.selectbox("选择测试值列", df.columns)
    upper_limit_col = st.selectbox("选择上限列", df.columns)
    lower_limit_col = st.selectbox("选择下限列", df.columns)
    
    df[date_col] = df[date_col].astype(str).str.strip()
    df[time_col] = df[time_col].astype(str).str.strip()
    
    df["TestDateTime"] = pd.to_datetime(df[date_col] + " " + df[time_col], errors='coerce')
    df["Quarter"] = df["TestDateTime"].dt.to_period("Q").astype(str).str.replace("-", "Q")

    df[test_value_col] = pd.to_numeric(df[test_value_col], errors='coerce')
    df[upper_limit_col] = pd.to_numeric(df[upper_limit_col], errors='coerce')
    df[lower_limit_col] = pd.to_numeric(df[lower_limit_col], errors='coerce')

    # 统计每个型号在每个季度的唯一 SN 数量
    df_type_stats = df.groupby([type_col, "Quarter"]).agg(
        total_tests=(sn_col, "count"),
        unique_sn_count=(sn_col, "nunique")  
    ).reset_index()

    # 取每个 SN 在每个季度的第一次测试数据
    df_first = df.sort_values("TestDateTime").groupby([sn_col, "Quarter"]).first().reset_index()

    df_cpk = df_first[[type_col, "Quarter", test_value_col, upper_limit_col, lower_limit_col]]
    df_stats = df_cpk.groupby([type_col, "Quarter"]).agg(
        mean=(test_value_col, "mean"),
        std=(test_value_col, "std"),
        upper_limit=(upper_limit_col, "mean"),
        lower_limit=(lower_limit_col, "mean")
    ).reset_index()

    df_stats["std"].replace(0, np.nan, inplace=True)

    df_stats["cpk_upper"] = ((df_stats["upper_limit"] - df_stats["mean"]) / (3 * df_stats["std"])).round(3)
    df_stats["cpk_lower"] = ((df_stats["mean"] - df_stats["lower_limit"]) / (3 * df_stats["std"])).round(3)
    df_stats["cpk"] = df_stats[["cpk_upper", "cpk_lower"]].min(axis=1).round(3)

    df_stats = df_stats.merge(df_type_stats, on=[type_col, "Quarter"], how="left")

    # 计算 Code 和 Rate
    df_stats["Code"] = df_stats["cpk"].apply(lambda x: "100%" if x < 0.95 else "A")
    df_stats["Rate"] = df_stats["cpk"].apply(lambda x: "1:1" if x < 0.95 else "1:8")

    # 透视表（Pivot）
    df_pivot = df_stats.pivot(index=type_col, columns="Quarter", values=["cpk", "std", "Code", "Rate", "unique_sn_count"])

    df_pivot.columns = [f"{q} {m}" if m != "unique_sn_count" else f"{q} Total SN" for m, q in df_pivot.columns]
    df_pivot = df_pivot.reset_index()

    df_pivot = df_pivot.round(3)

    st.subheader("CPK 计算结果（按季度）")
    st.dataframe(df_pivot)

    available_quarters = sorted(df_stats["Quarter"].unique())
    selected_quarters = st.multiselect("选择要显示的季度", available_quarters, default=available_quarters)

    selected_cols = [col for col in df_pivot.columns if any(q in col for q in selected_quarters)] + [type_col]
    df_filtered = df_pivot[selected_cols]

    if any("cpk" in col for col in df_filtered.columns):
        st.subheader("CPK 变化趋势（按季度）")
        cpk_cols = [col for col in df_filtered.columns if "cpk" in col]
        st.line_chart(df_filtered.set_index(type_col)[cpk_cols])
    else:
        st.warning("未找到 CPK 计算结果，请检查数据。")
