import streamlit as st
import pandas as pd
import os
import io

# 设置页面配置
st.set_page_config(page_title="汇总表转换工具", layout="wide")

def is_valid_number(x):
    """判断是否为有效数字"""
    return pd.notna(x) and isinstance(x, (int, float, complex))

def transform_excel_streamlit(uploaded_file, mode="detail"):
    # 1. 准备文件名
    name_part = os.path.splitext(uploaded_file.name)[0]
    output_name = f"改_{name_part}_{'重量表' if mode == 'weight' else '详情表'}.xlsx"

    # 2. 读取 Excel (header=None 确保我们可以通过索引精准访问行)
    try:
        df = pd.read_excel(uploaded_file, header=None, engine="openpyxl")
    except Exception as e:
        st.error(f"读取失败: {e}")
        return None, None, None

    # ---------- 核心索引校准 ----------
    # 第1行 (index 0): 重量
    # 第2行 (index 1): 分类
    # 第3行 (index 2): 制品名称
    # 第4行 (index 3): 单价 (金额)
    # 第6行起 (index 5): 人员数据
    # 第2列 (index 1): 名字 (B列)
    # 第3列起 (index 2): 制品数据 (C列往后)

    # 获取有效制品的列范围
    product_cols = []
    for col in range(2, df.shape[1]):
        v = df.iloc[2, col] # 检查第3行品名
        if pd.isna(v) or str(v).strip() == "":
            break
        product_cols.append(col)

    # 提前提取属性，避免在循环中重复计算
    product_names = {c: str(df.iloc[2, c]).strip() for c in product_cols}
    product_categories = {c: (str(df.iloc[1, c]).strip() if pd.notna(df.iloc[1, c]) else "") for c in product_cols}
    product_weights = {c: (float(df.iloc[0, c]) if is_valid_number(df.iloc[0, c]) else 0.0) for c in product_cols}
    # 对应你说的：制品对应金额在第四行 (index 3)
    product_prices = {c: (float(df.iloc[3, c]) if is_valid_number(df.iloc[3, c]) else 0.0) for c in product_cols}

    results = []

    # 从第6行 (index 5) 开始遍历人员
    for i in range(5, len(df)):
        name_cell = df.iloc[i, 1]  # B列 = 名字
        
        # 名字为空则跳过
        if pd.isna(name_cell) or str(name_cell).strip() == "":
            continue

        name = str(name_cell).strip()
        detail_list = []
        total_count = 0
        total_weight = 0.0
        total_money = 0.0

        for col in product_cols:
            cnt = df.iloc[i, col]

            if not is_valid_number(cnt) or cnt <= 0:
                continue

            cnt = float(cnt) # 支持半件或整数
            total_count += cnt
            
            # 计算逻辑
            total_weight += cnt * product_weights[col]
            total_money += cnt * product_prices[col]

            cat = product_categories[col]
            item = product_names[col]
            prefix = f"（{cat}）" if cat else ""
            
            # 格式化数量：如果是整数则显示整数，否则显示小数
            cnt_str = int(cnt) if cnt == int(cnt) else cnt
            detail_list.append(f"{prefix}{item}✖{cnt_str}")

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

    # 输出到 Excel 内存缓冲
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False)
    buffer.seek(0)

    return result_df, buffer, output_name

# ================= Streamlit UI =================

st.title("📊 汇总表 → 转换工具")

col1, col2 = st.columns([1, 1])
with col1:
    uploaded_file = st.file_uploader("1. 上传汇总表 Excel", type=["xlsx"])
with col2:
    mode = st.radio("2. 选择模式", ["detail", "weight"], 
                    format_func=lambda x: "详情表 (含金额)" if x=="detail" else "重量表 (含重量+金额)")

if uploaded_file:
    if st.button("🚀 点击开始转换"):
        with st.spinner("处理中..."):
            res_df, excel_out, fn = transform_excel_streamlit(uploaded_file, mode)
            
            if res_df is not None:
                if res_df.empty:
                    st.warning("转换完成，但没找到有效数据。请检查：B列是否有名字，第6行以下是否有数字。")
                else:
                    st.success(f"处理成功！共处理 {len(res_df)} 行数据。")
                    st.dataframe(res_df, use_container_width=True)
                    st.download_button("⬇ 下载结果", excel_out, file_name=fn)
