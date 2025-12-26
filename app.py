import streamlit as st
import pandas as pd
import os
import io


def is_valid_number(x):
    return pd.notna(x) and isinstance(x, (int, float))


def transform_excel_streamlit(uploaded_file, mode="detail"):
    name_part = os.path.splitext(uploaded_file.name)[0]
    output_name = f"改_{name_part}_{'重量表' if mode == 'weight' else '详情表'}.xlsx"

    # 读取 Excel
    df = pd.read_excel(uploaded_file, header=None, engine="openpyxl")

    # ---------- 1. 分类（第2行） ----------
    col_to_category = {}
    for col in range(2, df.shape[1]):
        v = df.iloc[1, col]
        col_to_category[col] = str(v).strip() if pd.notna(v) and str(v).strip() else ""

    # ---------- 2. 制品名称（第3行） ----------
    product_names = {}
    for col in range(2, df.shape[1]):
        v = df.iloc[2, col]
        if pd.isna(v) or str(v).strip() == "":
            break
        product_names[col] = str(v).strip()

    # ---------- 3. 重量（第1行） ----------
    product_weights = {
        col: float(df.iloc[0, col]) if is_valid_number(df.iloc[0, col]) else None
        for col in product_names
    }

    # ---------- 4. 单价（第4行） ----------
    product_prices = {
        col: float(df.iloc[3, col]) if is_valid_number(df.iloc[3, col]) else 0.0
        for col in product_names
    }

    results = []

    # ---------- 5. 人员数据（名字在第2列，第6行起） ----------
    for i in range(5, len(df)):
        name_cell = df.iloc[i, 1]  # B列 = 名字

        if not isinstance(name_cell, str) or not name_cell.strip():
            continue

        name = name_cell.strip()
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

            cat = col_to_category[col]
            weight = product_weights[col]
            price = product_prices[col]

            if weight is not None:
                total_weight += cnt * weight

            total_money += cnt * price

            prefix = f"（{cat}）" if cat else ""
            detail_list.append(f"{prefix}{item}✖{cnt}")

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

    result_df = pd.DataFrame(results)

    # 输出 Excel 到内存
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False)

    buffer.seek(0)
    return result_df, buffer, output_name


# ================= Streamlit UI =================

st.set_page_config(page_title="汇总表转换工具", layout="wide")

st.title("📊 汇总表 → 详情表 / 重量表")

uploaded_file = st.file_uploader(
    "上传汇总表 Excel（.xlsx）",
    type=["xlsx"]
)

mode = st.radio(
    "选择生成模式",
    options=["detail", "weight"],
    format_func=lambda x: "详情表" if x == "detail" else "重量表"
)

if uploaded_file:
    if st.button("🚀 生成 Excel"):
        with st.spinner("正在处理，请稍候..."):
            df_result, excel_buffer, filename = transform_excel_streamlit(
                uploaded_file,
                mode
            )

        st.success("✅ 生成完成")

        st.dataframe(df_result, use_container_width=True)

        st.download_button(
            label="⬇ 下载 Excel",
            data=excel_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("📌 请先上传 Excel 文件")
