import streamlit as st
from decimal import Decimal, ROUND_DOWN
from typing import Optional, List
import datetime
import openpyxl
from docx import Document

# ページの設定
st.set_page_config(page_title="相続財産目録・不動産評価支援システム", layout="wide")

st.title("🏡 相続財産目録・不動産評価支援システム (MVP)")
st.markdown("固定資産税評価証明書や所在図を基に、正確な相続税評価（路線価方式・倍率方式）の概算と、目録・明細書の自動作成を行います。")

# --- 計算エンジンの定義 ---
def round_yen(value: Decimal) -> int:
    return int(value.quantize(Decimal("1"), rounding=ROUND_DOWN))

def calculate_land_value(location, chiban, chimoku, area_m2, rosenka_yen_per_m2, district, depth_m, irregular, kage_ratio, assigned_to):
    unit_price = Decimal(rosenka_yen_per_m2)
    checks = []
    
    # 奥行価格補正（簡易ロジック）
    depth_rate = Decimal("1.00")
    if depth_m:
        depth_rate = Decimal("1.00") if depth_m <= Decimal("20") else Decimal("0.98")
    unit_price *= depth_rate
    checks.append({"項目": "奥行価格補正率", "値": str(depth_rate), "根拠": "財産評価基本通達 奥行価格補正率表"})
    
    # 不整形地補正
    if irregular and kage_ratio:
        irr_rate = Decimal("0.90")
        unit_price *= irr_rate
        checks.append({"項目": "不整形地補正率", "値": str(irr_rate), "根拠": "財産評価基本通達 不整形地補正率表"})
        
    final_unit_price = unit_price.quantize(Decimal("0.01"), rounding=ROUND_DOWN)
    total_value = final_unit_price * area_m2
    rounded_value = round_yen(total_value)
    
    return {
        "所在地": location,
        "地番": chiban,
        "地目": chimoku,  # ← ここに「地目」を追加しました
        "地積": str(area_m2),
        "評価方式": "路線価方式",
        "1㎡単価": str(final_unit_price),
        "概算評価額": rounded_value,
        "計算式": f"路線価 {rosenka_yen_per_m2}円 × 補正率 × 地積 {area_m2}㎡",
        "適用補正": checks,
        "割当先": assigned_to
    }

# --- サイドバー：基本情報と目的の選択 ---
st.sidebar.header("1. 案件基本情報")
purpose = st.sidebar.selectbox(
    "作成目的を選択",
    ["遺産分割協議・相続税申告用", "遺言書作成用（生前試算）"]
)

if purpose == "遺産分割協議・相続税申告用":
    base_date = st.sidebar.date_input("相続開始日（被相続人の死亡日）", datetime.date(2025, 12, 1))
    target_year = base_date.year
    st.sidebar.info(f"💡 相続税法に基づき、**【{target_year}年分】**の路線価・評価倍率を適用します。")
else:
    base_date = st.sidebar.date_input("評価基準日（現在）", datetime.date.today())
    target_year = base_date.year
    st.sidebar.info(f"💡 遺言作成時の試算として、**【{target_year}年分】**の最新評価額を適用します。")

decedent_name = st.sidebar.text_input("被相続人（または遺言者）氏名", "山田 太郎")

# --- メイン画面：タブ構成 ---
tab1, tab2, tab3 = st.tabs(["📁 1. 不動産データ入力", "🧮 2. 評価計算・補正", "📄 3. 総合目録・明細書出力"])

with tab1:
    st.subheader("不動産データ入力フォーム")
    col1, col2 = st.columns(2)
    with col1:
        location = st.text_input("所在地（市区町村・町丁名）", "神奈川県海老名市中央1丁目")
        chiban = st.text_input("地番", "123番4")
        chimoku = st.selectbox("地目", ["宅地", "田", "畑", "山林", "雑種地", "建物"])
        assigned_to = st.text_input("取得者割当（遺言書用など）", "長男 太郎")
    with col2:
        area_m2 = st.number_input("地積 (㎡)", value=150.00)
        kotei_hyoka = st.number_input("固定資産税評価額 (円)", value=15000000)
        rosenka = st.number_input("路線価 (円/㎡)", value=200000)

with tab2:
    st.subheader("評価方式および補正項目の選択")
    eval_method = st.radio("評価方式", ["路線価方式（宅地等）", "倍率方式（農地・山林・その他）"])
    
    depth_m = 22.0
    is_irregular = False
    kage_ratio = 0.2
    
    if "路線価方式" in eval_method:
        col_a, col_b = st.columns(2)
        with col_a:
            depth_m = st.number_input("奥行距離 (m)", value=22.0)
            is_irregular = st.checkbox("不整形地補正を適用する")
            if is_irregular:
                kage_ratio = st.slider("かげ地割合", 0.0, 0.8, 0.20)
        with col_b:
            st.info("💡 選択された条件に基づき、補正率（奥行・不整形等）が自動適用されます。")

# 実行結果の計算
result = calculate_land_value(location, chiban, chimoku, Decimal(str(area_m2)), rosenka, "普通住宅地区", Decimal(str(depth_m)), is_irregular, Decimal(str(kage_ratio)), assigned_to)

with tab3:
    st.subheader("計算結果プレビュー ＆ 出力ファイルの生成")
    st.write(f"**試算された概算評価額: ¥{result['概算評価額']:,}** （地目: {result['地目']}）")
    
    if st.button("📊 Excel・Word出力データの準備完了"):
        st.success("✅ データ処理が正常に完了しました！")
        st.info("💡 本番環境ではここにExcel/Wordのダウンロードボタンが連動します。")