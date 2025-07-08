import streamlit as st 
import pandas as pd
import math
import os

st.title("使用量と必要本数シミュレーター")

# === ✅ Excelファイルの選択肢（data/フォルダ構成） ===
file_options = {
    "K40": os.path.join("data", "32Rk40.xlsx"),
    "1085G": os.path.join("data", "1085G使用量.xlsx"),
    "E51G-JP": os.path.join("data", "E51G-JP使用量.xlsx")
}

# ファイル選択
selected_file_key = st.sidebar.selectbox("材質選択", list(file_options.keys()))
file_path = file_options[selected_file_key]

try:
    df = pd.read_excel(file_path, engine="openpyxl")

    # 使用量列の単位除去と変換（"127.5g" や " 127.5 g" に対応）
    df["使用量"] = (
        df["使用量"]
        .astype(str)
        .replace(r"[gG\s]", "", regex=True)
        .replace("", "0")
        .astype(float)
    )

    st.sidebar.header("⚙️ 設定")

    selected_processes = st.sidebar.multiselect(
        "工程を選択", options=df["工程"].unique(), default=list(df["工程"].unique())
    )
    operating_days = st.sidebar.number_input("稼働日数（生産）", min_value=1, value=20)
    production_units = st.sidebar.number_input("1日あたり生産台数", min_value=1, value=1100)
    drum_capacity = st.sidebar.number_input("ドラム缶容量 (kg)", min_value=1.0, value=250.0, step=10.0)
    split_days = st.sidebar.number_input("振り分け日数（搬入）", min_value=1, value=15)
    loss_per_drum = st.sidebar.number_input("1本交換時のロス量 (kg)", min_value=0.0, max_value=drum_capacity - 1, value=20.0)

    # 実質使用可能容量（ロスを除いた容量）
    usable_capacity = drum_capacity - loss_per_drum

    # フィルタリング
    filtered_df = df[df["工程"].isin(selected_processes)]

    # 工程ごとの1台あたり使用量（g）
    per_unit = filtered_df.groupby("工程")["使用量"].sum().reset_index()
    per_unit.columns = ["工程", "1台あたり使用量（g）"]

    # 総使用量（kg） = 使用量 × 台数 × 稼働日数 / 1000
    per_unit["総使用量（kg）"] = (
        per_unit["1台あたり使用量（g）"] * production_units * operating_days / 1000
    )

    # 必要ドラム缶数（ロス考慮後の使用可能容量で割って切り上げ）
    per_unit["必要ドラム缶数"] = per_unit["総使用量（kg）"].apply(
        lambda x: math.ceil(x / usable_capacity)
    )

    # 全体集計
    total_required_kg = per_unit["総使用量（kg）"].sum()
    total_drum_count = total_required_kg / usable_capacity
    daily_drum_count = total_drum_count / split_days
    total_loss_kg = per_unit["必要ドラム缶数"].sum() * loss_per_drum

    # 表示用データ
    per_unit_display = per_unit.copy()
    per_unit_display["総使用量（kg）"] = per_unit_display["総使用量（kg）"].map(lambda x: f"{x:.1f} kg")
    per_unit_display["必要ドラム缶数"] = per_unit_display["必要ドラム缶数"].astype(str) + " 本"

    # 表示
    st.subheader(f"📋 工程ごとの必要本数（kg）と必要ドラム缶数 [{selected_file_key}]")
    st.dataframe(per_unit_display)

    st.subheader("📌 総使用量の合計と日別振り分け（ドラム缶本数）")
    st.markdown(f"✅ 全工程の必要本数 合計: **{total_drum_count:.1f} 本**")
    st.markdown(f"📅 {split_days}日で振り分けた場合：**1日あたり {daily_drum_count:.1f} 本**")
    st.markdown(f"♻️ ドラム交換による総ロス見込み: **{total_loss_kg:.1f} kg**")

   

    # ===== 📆 発注スケジュール入力エリア =====
    st.subheader("📆 発注スケジュール（週×曜日）入力")

    # 週・曜日の設定
    days = ["月", "火", "水", "木", "金"]
    weeks = [f"{i+1}週目" for i in range((split_days // 5) + 1)]

    schedule = {}
    total_input = 0

    for week in weeks:
        st.markdown(f"**{week}**")
        cols = st.columns(len(days))
        for i, day in enumerate(days):
            key = f"{week}_{day}"
            val = cols[i].number_input(f"{day}", key=key, min_value=0, step=1, value=0)
            schedule[key] = val
            total_input += val

    st.markdown("---")
    st.markdown(f"🧮 入力した合計本数: **{total_input} 本**")
    st.markdown(f"🔢 自動計算した必要本数: **{math.ceil(total_drum_count)} 本**")

    if total_input != math.ceil(total_drum_count):
        st.warning("⚠️ 入力された本数が必要本数と一致していません。")
    else:
        st.success("✅ 入力されたスケジュールと必要本数が一致しています。")

except FileNotFoundError:
    st.error("❌ Excelファイルが見つかりません。dataフォルダに対象ファイルが存在するか確認してください。")

except Exception as e:
    st.error(f"⚠️ エラーが発生しました: {e}")



import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# 📄 ファイルパス
template_path = r"C:\Users\J0134011\OneDrive - Honda\デスクトップ\シーラー管理\calendar_template.xlsx"

# 📌 設定
工程_材質リスト = [
    ("D3/D4", "K40"),
    ("D3/D4", "E51G-JP"),
    ("D7", "ペンギンセメント1085G")
]

曜日リスト = ["月", "火", "水", "木", "金"] * 4  # 週×5日 → 20列分

# 📋 入力
st.title("🗓 ドラム缶発注スケジュール入力")

selected_item = st.selectbox("工程・材質の組み合わせを選択", [f"{k[0]}・{k[1]}" for k in 工程_材質リスト])
選択_工程, 選択_材質 = selected_item.split("・")

入力値 = []
st.markdown("#### 各日のドラム缶本数を入力")

cols = st.columns(5)
for i in range(4):  # 4週分
    st.markdown(f"**{i+1}週目**")
    for j in range(5):  # 月〜金
        day_idx = i * 5 + j
        with cols[j]:
            val = st.number_input(f"{曜日リスト[day_idx]}", min_value=0, step=1, key=f"{i}_{j}")
            入力値.append(val)

if st.button("✅ 確定してExcelに保存"):

    try:
        # Excelテンプレートを読み込む
        wb = load_workbook(template_path)
        ws = wb.active

        # 工程・材質の行を探す
        row_num = None
        for row in range(2, ws.max_row + 1):  # 2行目以降（ヘッダー除く）
            process_cell = str(ws.cell(row=row, column=1).value).strip()
            material_cell = str(ws.cell(row=row, column=2).value).strip()

            if process_cell == 選択_工程 and material_cell == 選択_材質:
                row_num = row
                break

        if row_num is None:
            st.error("❌ 該当する工程・材質の行がテンプレートに見つかりません。")
        else:
            # 入力値を3列目以降に順に書き込み（カレンダーの各日）
            for col_index, val in enumerate(入力値, start=3):
                ws.cell(row=row_num, column=col_index, value=val)

            # 保存
            wb.save(template_path)
            st.success("✅ 入力内容をExcelに保存しました！")

    except Exception as e:
        st.error(f"❌ 保存中にエラーが発生しました: {e}")

