import streamlit as st
import pandas as pd

# Streamlit UI
st.title("RFB/Thermal测试数据 CPK 计算")

# 文件上传
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    # 读取 Excel
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("选择工作表", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()  # 去除列名空格

    # 去除所有列的空格
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # 显示列名，供用户选择
    # st.write("Excel 列名:", list(df.columns))
    
    # 用户自定义列名
    type_col = st.selectbox("选择型号列", df.columns)
    sn_col = st.selectbox("选择探头 SN 列", df.columns)
    date_col = st.selectbox("选择日期列", df.columns)
    time_col = st.selectbox("选择时间列", df.columns)
    test_value_col = st.selectbox("选择测试值列", df.columns)
    upper_limit_col = st.selectbox("选择上限列", df.columns)
    lower_limit_col = st.selectbox("选择下限列", df.columns)
    
    # 确保日期和时间列为字符串
    df[date_col] = df[date_col].astype(str).str.strip()
    df[time_col] = df[time_col].astype(str).str.strip()
    
    # 生成测试时间列
    df["TestDateTime"] = pd.to_datetime(df[date_col] + " " + df[time_col], errors='coerce')
    
    # 计算总测试次数和唯一探头数量
    total_tests = len(df)
    unique_sn_count = df[sn_col].nunique()
    
    # 统计每种型号的总测试次数和唯一探头数量
    df_type_stats = df.groupby(type_col).agg(
        total_tests=(sn_col, "count"),
        unique_sn_count=(sn_col, "nunique")
    ).reset_index()
    
    # 取每个探头的第一次测试数据
    df_first = df.sort_values("TestDateTime").groupby(sn_col).first().reset_index()
    
    # 计算 CPK
    df_cpk = df_first[[type_col, test_value_col, upper_limit_col, lower_limit_col]]
    df_stats = df_cpk.groupby(type_col).agg(
        mean=(test_value_col, "mean"),
        std=(test_value_col, "std"),
        upper_limit=(upper_limit_col, "mean"),
        lower_limit=(lower_limit_col, "mean")
    ).reset_index()
    
    df_stats["cpk_upper"] = (df_stats["upper_limit"] - df_stats["mean"]) / (3 * df_stats["std"])
    df_stats["cpk_lower"] = (df_stats["mean"] - df_stats["lower_limit"]) / (3 * df_stats["std"])
    df_stats["cpk"] = df_stats[["cpk_upper", "cpk_lower"]].min(axis=1)
    
    # 合并统计数据
    df_stats = df_stats.merge(df_type_stats, on=type_col, how="left")
    
    # 显示统计信息
    st.subheader("测试数据统计")
    st.write(f"Total Tests: {total_tests}")
    st.write(f"Unique SN Count: {unique_sn_count}")
    
    # 显示按型号统计的测试次数和唯一探头数
    #st.subheader("按型号统计")
    #st.dataframe(df_type_stats)
    
    # 显示 CPK 计算结果
    st.subheader("CPK 计算结果")
    st.dataframe(df_stats)
    
    # 可视化
    st.bar_chart(df_stats.set_index(type_col)["cpk"])
