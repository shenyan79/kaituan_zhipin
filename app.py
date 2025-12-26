import streamlit as st
import pandas as pd
import os
import io

# 设置页面配置（必须是 Streamlit 命令的第一行）
st.set_page_config(page_title="汇总表转换工具", layout="wide")

def is_valid_number(x):
    return pd.notna(x) and isinstance(x, (int, float))

def transform_excel_streamlit(uploaded_file, mode="detail"):
    # 提取文件名
    name_part = os.path.splitext(uploaded_file.name)[0]
    output_name = f"改_{name_part}_{'重量表' if mode == 'weight' else '详情表'}.xlsx"

    # 读取 Excel
    # 注意：确保 header=None，因为后续逻辑是按索引 iloc 读取的
    try:
        df = pd.read_excel(uploaded_file, header=None, engine="openpyxl")
    except Exception as e:
        st.error(f"读取 Excel 失败: {e}")
        return None, None, None

    # ---------- 1. 分类（第2行，索引1） ----------
    col_to_category = {}
    for col in range(2, df.shape[1]):
        v = df.iloc[1, col]
        col_to_category[col] = str(v).strip() if pd.notna(v) and str(v).strip() else ""

    # ---------- 2. 制品名称（第3行，索引2） ----------
    product_names = {}
    for col in range(2, df.shape[1]):
        v = df.iloc[2, col]
        if pd.isna(v) or str(v).strip() == "":
            break
        product_names[col] = str(v).strip()

    # ---------- 3. 重量（第1行，索引0） ----------
    product_weights = {
        col: float(df.iloc[0, col]) if is_valid_number(df.iloc[0, col]) else None
        for col in product_names
    }

    # ---------- 4. 单价（第4行，索引3） ----------
    product_prices = {
        col: float(df.iloc[3, col]) if is_valid_number(df.iloc[3, col]) else 0.0
        for col in product_names
    }

    results = []

    # ---------- 5. 人员数据（名字在第2列即B列，从第6行即索引5起） ----------
    # 这里通过 len(df) 动态获取行数，确保 df 已定义
    for i in range(5, len(df)):
        name_cell = df.iloc[i, 1]  # B列 = 索引1

        # 跳过空行
        if pd.isna(name_cell) or str(name_cell).strip() == "":
            continue

        name = str(name_cell).strip()
        detail_list = []

        total_count = 0
        total_weight = 0.0
        total_money = 0.0

        for col, item in product_names.items():
            cnt = df.iloc[i, col]

            if not is_valid_number(cnt) or cnt <= 0:
                continue

            cnt = int(cnt)
            total_count += cnt

            cat = col_to_category.get(col, "")
            weight = product_weights.get(col)
            price = product_prices.get(col, 0.0)

            if weight is not None:
                total_weight += cnt * weight

            total_money += cnt * price

            prefix = f"（{cat}）" if cat else ""
            detail_list.append(f"{prefix}{item}×{cnt}")

        if not detail_list:
            continue

        row = {
            "名字": name,
            "（分类）制品×数量": " / ".join(detail_list),
            "总点数": total_count,
            "总金额": round(total_money, 3)
        }

        if mode == "weight":
            row["总重量"] = round(total_weight, 2)

        results.append(row)

    if not results:
        return pd.DataFrame(), None, output_name

    result_df = pd.DataFrame(results)

    # 输出 Excel 到内存
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False)
    buffer.seek(0)
    
    return result_df, buffer, output_name


# ================= Streamlit UI =================

st.title("📊 汇总表 → 详情表 / 重量表")
st.markdown("请确保 Excel 格式：第1行重量，第2行分类，第3行品名，第4行单价，第6行起为人员数据。")

uploaded_file = st.file_uploader(
    "上传汇总表 Excel（.xlsx）",
    type=["xlsx"]
)

mode = st.radio(
    "选择生成模式",
    options=["detail", "weight"],
    format_func=lambda x: "详情表 (不含重量)" if x == "detail" else "重量表 (包含总重量)"
)

if uploaded_file:
    # 增加预览功能
    with st.expander("查看原始文件预览"):
        preview_df = pd.read_excel(uploaded_file, header=None).head(10)
        st.dataframe(preview_df)

    if st.button("🚀 开始转换"):
        with st.spinner("正在处理，请稍候..."):
            df_result, excel_buffer, filename = transform_excel_streamlit(uploaded_file, mode)

        if df_result is not None and not df_result.empty:
            st.success("✅ 转换成功！")
            st.dataframe(df_result, use_container_width=True)

            st.download_button(
                label="⬇ 下载转换后的 Excel",
                data=excel_buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif df_result is not None and df_result.empty:
            st.warning("⚠️ 转换完成，但未发现有效的人员数据，请检查 Excel 格式。")
else:
    st.info("📌 请先上传 Excel 文件以开始转换。")
