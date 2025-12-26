# ---------- 5. 人员数据（第6行起，名字在第2列） ----------
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
