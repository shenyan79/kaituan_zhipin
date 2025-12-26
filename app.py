import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="汇总表 → 详情表 / 重量表", layout="wide")

st.title("📊 汇总表 → 详情表 / 重量表")

# -------------------------------
# 核心处理函数
# -------------------------------
def transform_excel_streamlit(uploaded_file, mode="detail"):
    df = pd.read_excel(uploaded_file, header=None)

    # ===== 基础结构约定 =====
    # 第 1 行：分类
    # 第 2 行：制品分类
    # 第 3 行：种类
    # 第 4 行：单价（关键）
    # 第 5 行开始：人员数据

    name_col = 0
    product_start_col = 2
    price_row = 3        # 单价行（0-based）
    data_start_row = 5   # 人员数据起始行（0-based）

    prices = df.iloc[price_row, product_start_col:].fillna(0)

    result_rows = []

    for i in range(data_start_row, len(df)):
        name = df.iloc[i, name_col]

        if pd.isna(name):
            continue

        quantities = df.iloc[i, product_start_col:].fillna(0)

        total_qty = quantities.sum()
        total_amount = (quantities * prices).sum()

        for col_idx, qty in quantities.items():
            if qty == 0:
                continue

            product_name = df.iloc[2, col_idx]
            price = prices[col_idx]
            amount = qty * price

            if mode == "detail":
                result_rows.append({
                    "名字": name,
                    "制品": product_name,
                    "数量": int(qty),
                    "单价": round(float(price), 3),
                    "金额": round(float(amount), 3)
                })

        if mode == "weight":
            result_rows.append({
                "名字": name,
                "总点数": int(total_qty),
                "总金额": round(float(total_amount), 3)
            })

    df_result = pd.DataFrame(result_rows)

    # ==========================
    # 导出 Excel
    # ==========================
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False)

    buffer.seek(0)

    filename = f"{'详情表' if mode=='detail' else '重量表'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return df_result, buffer, filename


# -------------------------------
# Streamlit UI
# -------------------------------
uploaded_file = st.file_uploader(
    "上传汇总表 Excel（.xlsx）",
    type=["xlsx"]
)

mode = st.radio(
    "选择生成模式",
    ["详情表", "重量表"]
)

if uploaded_file and st.button("🚀 生成 Excel"):
    with st.spinner("处理中..."):
        df_result, excel_buffer, filename = transform_excel_streamlit(
            uploaded_file,
            mode="detail" if mode == "详情表" else "weight"
        )

    st.success("✅ 生成完成")

    st.dataframe(df_result, use_container_width=True)

    st.download_button(
        label="⬇️ 下载 Excel",
        data=excel_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
