# -*- coding: utf-8 -*-
"""
4期1Qと4期2Qの卒業生の卒業時平均投稿数を比較するスクリプト。

【4期1Q】
- 8月卒業: 3名
- 9月卒業: 0名
- 10月卒業: 6名

【4期2Q】
- 11月卒業: 8名
- 12月卒業: 10名
- 1月卒業: 5名

【データ出所】コミットプラン (4).xlsx の「新 月次投稿数」シート
【卒業時投稿数】6ヶ月目の時点での合計投稿数（または0-6ヶ月目の合計）
"""
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

# 入力パス
BASE = Path(__file__).parent.parent
INPUT_PATH = BASE / "コミットプラン (4).xlsx"
if not INPUT_PATH.exists():
    INPUT_PATH = Path.home() / "Downloads" / "コミットプラン (4).xlsx"
OUTPUT_PATH = Path(__file__).parent / "4期1Q_2Q_卒業時平均投稿数比較結果.xlsx"
REPORT_PATH = Path(__file__).parent.parent / "分析結果" / "4期1Q_2Q_卒業時平均投稿数比較.md"

# セッション実施状況管理
SESS_HEADER_ROW = 9
SESS_DATA_START = 10
SESS_COL_NO = 0
SESS_COL_NAME = 7
SESS_COL_MG = 19  # T列: 担当MG名
SESS_COL_FIRST_NORMAL = 22  # W列: 初回の通常セッション日
SESS_COL_6TH_DATE = None  # 6回目実施日の列を探す必要がある

# 新 月次投稿数
MONTH_HEADER_ROW = 10
MONTH_DATA_START = 11
MONTH_COL_NO = 0
MONTH_COL_NAME = 4
MONTH_COL_0M = 15   # P列: 0ヶ月目
MONTH_COL_6M = 21   # V列: 6ヶ月目

# 卒業生リスト
GRADUATES_4Q1Q = {
    "2025-08": [  # 8月卒業
        "てらもとまさゆき",
        "しまむらまりな",
        "かわさきちあき",
    ],
    "2025-09": [],  # 9月卒業: 0名
    "2025-10": [  # 10月卒業
        "きべまき",
        "きくちことの",
        "さえきもえか",
        "かとうともか",
        "おおしろにいな",
        "ひろたはるな",
    ],
}

GRADUATES_4Q2Q = {
    "2025-11": [  # 11月卒業
        "らぶひとみ",
        "かげやまこゆき",
        "かわらまき",
        "のぶとうまさこ",
        "こうごたかひろ",
        "やまぐちちづる",
        "ひらやまみか",
        "しむらまなぶ",
    ],
    "2025-12": [  # 12月卒業
        "またよししょうや",
        "ながいけいこ",
        "じくまるみほ",
        "はらだたかよ",
        "おおつばきかなこ",
        "のざきせいか",
        "ろばーつあゆみ",
        "いまむら やすえ",
        "かきのきまゆみ",
        "すのうちなお",
    ],
    "2026-01": [  # 1月卒業
        "なかおしょうや",
        "みやざとせいぎ",
        "たむらやすのり",
        "よしだえみ",
        "まえだりお",
    ],
}


def normalize_name(name):
    """名前を正規化（全角スペース→半角スペース、小文字化など）"""
    if pd.isna(name):
        return ""
    s = str(name).strip()
    # 全角スペースを半角スペースに統一
    s = s.replace("　", " ")
    # 小文字化
    s = s.lower()
    return s


def to_num(val):
    """数値に変換"""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ("ー", "－", "-", ""):
        return None
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def to_display_val(val):
    """表示用の値に変換（数値または「ー」）"""
    if pd.isna(val):
        return "ー"
    s = str(val).strip()
    if s in ("ー", "－", "-", ""):
        return "ー"
    try:
        num = int(float(s))
        return num
    except (ValueError, TypeError):
        return "ー"


def find_6th_session_col(df_sess):
    """6回目実施日の列を探す"""
    # ヘッダー行を確認
    header_row = df_sess.iloc[SESS_HEADER_ROW]
    for i, col in enumerate(header_row):
        if pd.notna(col):
            col_str = str(col).strip()
            if "6回目" in col_str and ("実施日" in col_str or "日" in col_str):
                return i
    return None


def calculate_total_posts(df_month, row_idx, start_col, end_col):
    """指定範囲の合計投稿数を計算"""
    total = 0
    for col_idx in range(start_col, end_col + 1):
        val = to_num(df_month.iloc[row_idx, col_idx])
        if val is not None:
            total += val
    return total if total > 0 else None


def main():
    df_sess = pd.read_excel(INPUT_PATH, sheet_name="セッション実施状況管理", header=None)
    df_month = pd.read_excel(INPUT_PATH, sheet_name="新 月次投稿数", header=None)
    
    # 6回目実施日の列を探す
    sess_6th_col = find_6th_session_col(df_sess)
    if sess_6th_col is None:
        print("警告: 6回目実施日の列が見つかりません。名前マッチングのみで判定します。")
    
    # no. -> 名前、6回目実施日、初回セッション日 のマッピング
    sess_map = {}
    for i in range(SESS_DATA_START, len(df_sess)):
        no_ = df_sess.iloc[i, SESS_COL_NO]
        if pd.isna(no_):
            continue
        try:
            no_int = int(float(no_))
        except (ValueError, TypeError):
            continue
        
        name = df_sess.iloc[i, SESS_COL_NAME]
        mg_name = df_sess.iloc[i, SESS_COL_MG]
        first_sess = df_sess.iloc[i, SESS_COL_FIRST_NORMAL]
        sess_6th = None
        if sess_6th_col is not None:
            sess_6th = df_sess.iloc[i, sess_6th_col]
        
        sess_map[no_int] = {
            "name": normalize_name(name),
            "mg": mg_name if pd.notna(mg_name) else "",
            "first_sess": pd.Timestamp(first_sess) if pd.notna(first_sess) and isinstance(first_sess, (pd.Timestamp, datetime)) else None,
            "6th_sess": pd.Timestamp(sess_6th) if pd.notna(sess_6th) and isinstance(sess_6th, (pd.Timestamp, datetime)) else None,
        }
    
    # 全卒業生データを収集（卒業の定義: 6回目実施日がある、または6ヶ月目のデータがある）
    all_graduates_results = []
    
    # 新 月次投稿数シートから全卒業生データを取得
    for i in range(MONTH_DATA_START, len(df_month)):
        no_ = df_month.iloc[i, MONTH_COL_NO]
        if pd.isna(no_):
            continue
        try:
            no_int = int(float(no_))
        except (ValueError, TypeError):
            continue
        
        name_raw = df_month.iloc[i, MONTH_COL_NAME]
        if pd.isna(name_raw) or str(name_raw).strip() == "":
            continue
        
        sess_info = sess_map.get(no_int)
        if sess_info is None:
            continue
        
        # 卒業判定: 6回目実施日がある、または6ヶ月目のデータがある（0でもデータがあれば卒業とみなす）
        is_graduated = False
        if sess_info["6th_sess"] is not None:
            is_graduated = True
        else:
            # 6ヶ月目の列にデータがあるか確認（値が0でもデータがあれば卒業とみなす）
            month_6_raw = df_month.iloc[i, MONTH_COL_6M]
            if pd.notna(month_6_raw):
                # 値が「ー」や空文字でなければデータがあるとみなす
                month_6_str = str(month_6_raw).strip()
                if month_6_str not in ("ー", "－", "-", ""):
                    is_graduated = True
        
        if not is_graduated:
            continue
        
        # 各月の値を個別に取得（表示用と計算用）
        monthly_posts_display = {}
        monthly_posts_calc = {}
        for month_idx in range(7):  # 0-6ヶ月目
            col_idx = MONTH_COL_0M + month_idx
            val_raw = df_month.iloc[i, col_idx]
            val_display = to_display_val(val_raw)
            val_calc = to_num(val_raw)
            monthly_posts_display[f"{month_idx}m"] = val_display
            monthly_posts_calc[f"{month_idx}ヶ月目"] = val_calc if val_calc is not None else 0
        
        # 卒業時投稿数 = 0-6ヶ月目の合計（数値のみで計算）
        total_calc = sum(monthly_posts_calc.values())
        
        # 卒業月を6回目実施日から取得、なければ6ヶ月目のデータがある月を推定
        grad_month = ""
        if sess_info["6th_sess"] is not None:
            grad_month = sess_info["6th_sess"].strftime("%Y-%m")
        else:
            # 6ヶ月目のデータがある場合、初回セッションから6ヶ月後を推定
            if sess_info["first_sess"] is not None:
                grad_date = sess_info["first_sess"] + pd.DateOffset(months=6)
                grad_month = grad_date.strftime("%Y-%m")
        
        all_graduates_results.append({
            "no.": no_int,
            "生徒名": name_raw if not pd.isna(name_raw) else sess_info["name"],
            "担当MG": sess_info.get("mg", ""),
            "卒業月": grad_month,
            "卒業時投稿数": total_calc,
            "0m": monthly_posts_display.get("0m", "ー"),
            "1m": monthly_posts_display.get("1m", "ー"),
            "2m": monthly_posts_display.get("2m", "ー"),
            "3m": monthly_posts_display.get("3m", "ー"),
            "4m": monthly_posts_display.get("4m", "ー"),
            "5m": monthly_posts_display.get("5m", "ー"),
            "6m": monthly_posts_display.get("6m", "ー"),
            "初回セッション日": sess_info["first_sess"].strftime("%Y-%m-%d") if sess_info["first_sess"] else "",
            "6回目実施日": sess_info["6th_sess"].strftime("%Y-%m-%d") if sess_info["6th_sess"] else "",
        })
    
    # 4期1Qと4期2Qの特定の卒業生データを収集
    results = []
    
    # 全卒業生リストを統合
    all_graduates = {}
    for period, names in GRADUATES_4Q1Q.items():
        all_graduates[period] = names
    for period, names in GRADUATES_4Q2Q.items():
        all_graduates[period] = names
    
    # 新 月次投稿数シートからデータを取得
    for i in range(MONTH_DATA_START, len(df_month)):
        no_ = df_month.iloc[i, MONTH_COL_NO]
        if pd.isna(no_):
            continue
        try:
            no_int = int(float(no_))
        except (ValueError, TypeError):
            continue
        
        name_raw = df_month.iloc[i, MONTH_COL_NAME]
        if pd.isna(name_raw) or str(name_raw).strip() == "":
            continue
        
        name = normalize_name(name_raw)
        sess_info = sess_map.get(no_int)
        
        if sess_info is None:
            continue
        
        # 名前マッチングで卒業生を特定（より厳密に）
        matched_period = None
        matched_grad_name = None
        for period, grad_names in all_graduates.items():
            for grad_name in grad_names:
                grad_name_norm = normalize_name(grad_name)
                # 完全一致または、grad_nameがnameに含まれる（括弧内の情報を除く）
                name_clean = name.split("（")[0].split("(")[0].strip()
                grad_name_clean = grad_name_norm.split("（")[0].split("(")[0].strip()
                
                if grad_name_clean == name_clean or grad_name_clean in name_clean:
                    matched_period = period
                    matched_grad_name = grad_name
                    break
            if matched_period:
                break
        
        if not matched_period:
            continue
        
        # 6ヶ月目の投稿数を取得（卒業時投稿数）
        # 方法1: 6ヶ月目の列の値（その月の投稿数）
        month_6_val = to_num(df_month.iloc[i, MONTH_COL_6M])
        
        # 方法2: 0-6ヶ月目の合計投稿数
        total_0_to_6 = calculate_total_posts(df_month, i, MONTH_COL_0M, MONTH_COL_6M)
        
        # 各月の値を個別に取得（表示用と計算用）
        monthly_posts_display = {}
        monthly_posts_calc = {}
        for month_idx in range(7):  # 0-6ヶ月目
            col_idx = MONTH_COL_0M + month_idx
            val_raw = df_month.iloc[i, col_idx]
            val_display = to_display_val(val_raw)
            val_calc = to_num(val_raw)
            monthly_posts_display[f"{month_idx}m"] = val_display
            monthly_posts_calc[f"{month_idx}ヶ月目"] = val_calc if val_calc is not None else 0
        
        # 卒業時投稿数 = 0-6ヶ月目の合計（数値のみで計算）
        total_calc = sum(monthly_posts_calc.values())
        
        results.append({
            "no.": no_int,
            "生徒名": name_raw if not pd.isna(name_raw) else sess_info["name"],
            "担当MG": sess_info.get("mg", ""),
            "マッチした卒業生名": matched_grad_name,
            "卒業月": matched_period,
            "卒業時投稿数": total_calc,
            "0m": monthly_posts_display.get("0m", "ー"),
            "1m": monthly_posts_display.get("1m", "ー"),
            "2m": monthly_posts_display.get("2m", "ー"),
            "3m": monthly_posts_display.get("3m", "ー"),
            "4m": monthly_posts_display.get("4m", "ー"),
            "5m": monthly_posts_display.get("5m", "ー"),
            "6m": monthly_posts_display.get("6m", "ー"),
            "初回セッション日": sess_info["first_sess"].strftime("%Y-%m-%d") if sess_info["first_sess"] else "",
            "6回目実施日": sess_info["6th_sess"].strftime("%Y-%m-%d") if sess_info["6th_sess"] else "",
        })
    
    if not results:
        print("エラー: 卒業生が見つかりませんでした。名前のマッチングを確認してください。")
        return
    
    result_df = pd.DataFrame(results)
    
    # 4期1Qと4期2Qに分類
    result_df["期"] = result_df["卒業月"].apply(
        lambda x: "4期1Q" if x in ["2025-08", "2025-09", "2025-10"] else "4期2Q"
    )
    
    # 4期1Qと4期2Qを分離
    q1_df = result_df[result_df["期"] == "4期1Q"].copy()
    q2_df = result_df[result_df["期"] == "4期2Q"].copy()
    
    # 卒業月でソート
    q1_df = q1_df.sort_values(["卒業月", "生徒名"])
    q2_df = q2_df.sort_values(["卒業月", "生徒名"])
    
    # 表示用の列を選択（必要な列のみ）
    display_cols = ["生徒名", "担当MG", "卒業時投稿数", "0m", "1m", "2m", "3m", "4m", "5m", "6m"]
    q1_display = q1_df[display_cols].copy()
    q2_display = q2_df[display_cols].copy()
    
    # 集計用の計算
    q1_total_count = len(q1_df)
    q1_avg = round(q1_df["卒業時投稿数"].mean(), 2) if q1_total_count > 0 else 0
    
    q2_total_count = len(q2_df)
    q2_avg = round(q2_df["卒業時投稿数"].mean(), 2) if q2_total_count > 0 else 0
    
    # 全卒業生82名のデータフレーム
    all_graduates_df = pd.DataFrame(all_graduates_results)
    
    # MG別分析（MGが空でないもののみ）
    # 全期間（全卒業生82名）
    all_graduates_df_mg = all_graduates_df[all_graduates_df["担当MG"].notna() & (all_graduates_df["担当MG"] != "")].copy()
    mg_all = all_graduates_df_mg.groupby("担当MG").agg({
        "卒業時投稿数": ["count", "mean", "sum", "std", "min", "max"]
    }).round(2)
    mg_all.columns = ["人数", "平均投稿数", "合計投稿数", "標準偏差", "最小値", "最大値"]
    mg_all = mg_all.sort_values("平均投稿数", ascending=False)
    
    # 1Q
    q1_df_mg = q1_df[q1_df["担当MG"].notna() & (q1_df["担当MG"] != "")].copy()
    mg_q1 = q1_df_mg.groupby("担当MG").agg({
        "卒業時投稿数": ["count", "mean", "sum", "std", "min", "max"]
    }).round(2)
    mg_q1.columns = ["人数", "平均投稿数", "合計投稿数", "標準偏差", "最小値", "最大値"]
    mg_q1 = mg_q1.sort_values("平均投稿数", ascending=False)
    
    # 2Q
    q2_df_mg = q2_df[q2_df["担当MG"].notna() & (q2_df["担当MG"] != "")].copy()
    mg_q2 = q2_df_mg.groupby("担当MG").agg({
        "卒業時投稿数": ["count", "mean", "sum", "std", "min", "max"]
    }).round(2)
    mg_q2.columns = ["人数", "平均投稿数", "合計投稿数", "標準偏差", "最小値", "最大値"]
    mg_q2 = mg_q2.sort_values("平均投稿数", ascending=False)
    
    # 相関分析: 全期間 vs 1Q, 全期間 vs 2Q
    # MG別平均投稿数をマージ
    mg_comparison = pd.DataFrame({
        "全期間_平均投稿数": mg_all["平均投稿数"],
        "全期間_人数": mg_all["人数"],
    })
    
    # 1Qと2Qのデータをマージ（該当MGのみ）
    mg_comparison = mg_comparison.merge(
        mg_q1[["平均投稿数", "人数"]].rename(columns={"平均投稿数": "1Q_平均投稿数", "人数": "1Q_人数"}),
        left_index=True, right_index=True, how="left"
    )
    mg_comparison = mg_comparison.merge(
        mg_q2[["平均投稿数", "人数"]].rename(columns={"平均投稿数": "2Q_平均投稿数", "人数": "2Q_人数"}),
        left_index=True, right_index=True, how="left"
    )
    
    # 相関係数の計算（データがあるMGのみ）
    # 全期間 vs 1Q
    mg_all_vs_q1 = mg_comparison[mg_comparison["1Q_平均投稿数"].notna()]
    if len(mg_all_vs_q1) > 1:
        corr_all_q1 = mg_all_vs_q1["全期間_平均投稿数"].corr(mg_all_vs_q1["1Q_平均投稿数"])
    else:
        corr_all_q1 = None
    
    # 全期間 vs 2Q
    mg_all_vs_q2 = mg_comparison[mg_comparison["2Q_平均投稿数"].notna()]
    if len(mg_all_vs_q2) > 1:
        corr_all_q2 = mg_all_vs_q2["全期間_平均投稿数"].corr(mg_all_vs_q2["2Q_平均投稿数"])
    else:
        corr_all_q2 = None
    
    # 1Q vs 2Q
    mg_q1_vs_q2 = mg_comparison[mg_comparison["1Q_平均投稿数"].notna() & mg_comparison["2Q_平均投稿数"].notna()]
    if len(mg_q1_vs_q2) > 1:
        corr_q1_q2 = mg_q1_vs_q2["1Q_平均投稿数"].corr(mg_q1_vs_q2["2Q_平均投稿数"])
    else:
        corr_q1_q2 = None
    
    # Excel出力
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as w:
        # 4期1Qの詳細データ
        q1_display.to_excel(w, sheet_name="4期1Q", index=False)
        
        # 4期2Qの詳細データ
        q2_display.to_excel(w, sheet_name="4期2Q", index=False)
        
        # 全データ（詳細情報付き）
        result_df.to_excel(w, sheet_name="全データ", index=False)
        
        # 集計サマリ
        summary = pd.DataFrame([
            {"項目": "4期1Q 卒業生総数", "値": q1_total_count},
            {"項目": "4期1Q 平均卒業時投稿数", "値": q1_avg},
            {"項目": "4期2Q 卒業生総数", "値": q2_total_count},
            {"項目": "4期2Q 平均卒業時投稿数", "値": q2_avg},
            {"項目": "差（2Q - 1Q）", "値": round(q2_avg - q1_avg, 2)},
        ])
        summary.to_excel(w, sheet_name="集計", index=False)
        
        # MG別分析
        mg_all.to_excel(w, sheet_name="MG別_全期間（全卒業生82名）", index=True)
        mg_q1.to_excel(w, sheet_name="MG別_1Q", index=True)
        mg_q2.to_excel(w, sheet_name="MG別_2Q", index=True)
        
        # 相関分析結果
        mg_comparison.to_excel(w, sheet_name="MG別_相関分析", index=True)
        
        # 相関係数サマリ
        corr_summary = pd.DataFrame([
            {"比較項目": "全期間 vs 1Q", "相関係数": round(corr_all_q1, 3) if corr_all_q1 is not None else "データ不足", "対象MG数": len(mg_all_vs_q1)},
            {"比較項目": "全期間 vs 2Q", "相関係数": round(corr_all_q2, 3) if corr_all_q2 is not None else "データ不足", "対象MG数": len(mg_all_vs_q2)},
            {"比較項目": "1Q vs 2Q", "相関係数": round(corr_q1_q2, 3) if corr_q1_q2 is not None else "データ不足", "対象MG数": len(mg_q1_vs_q2)},
        ])
        corr_summary.to_excel(w, sheet_name="相関係数サマリ", index=False)
        
        # 全卒業生82名の詳細データ（参考用）
        all_graduates_display = all_graduates_df[["生徒名", "担当MG", "卒業時投稿数", "0m", "1m", "2m", "3m", "4m", "5m", "6m"]].copy()
        all_graduates_display.to_excel(w, sheet_name="全卒業生82名", index=False)
    
    # Markdownレポート（詳細データ付き）
    report_lines = [
        "# 4期1Q vs 4期2Q 卒業時平均投稿数比較",
        "",
        "## サマリ",
        "",
        f"- **4期1Q 平均卒業時投稿数**: {q1_avg}投稿（{q1_total_count}名）",
        f"- **4期2Q 平均卒業時投稿数**: {q2_avg}投稿（{q2_total_count}名）",
        f"- **差（2Q - 1Q）**: {round(q2_avg - q1_avg, 2):+.2f}投稿",
        "",
        "---",
        "",
        "## 4期1Q 詳細データ",
        "",
        "| 生徒名 | 担当MG | 卒業時投稿数 | 0m | 1m | 2m | 3m | 4m | 5m | 6m |",
        "|--------|--------|-------------|----|----|----|----|----|----|----|",
    ]
    
    for _, row in q1_display.iterrows():
        name = str(row["生徒名"])
        mg = str(row["担当MG"]) if pd.notna(row["担当MG"]) else ""
        total = row["卒業時投稿数"]
        m0 = row["0m"]
        m1 = row["1m"]
        m2 = row["2m"]
        m3 = row["3m"]
        m4 = row["4m"]
        m5 = row["5m"]
        m6 = row["6m"]
        report_lines.append(f"| {name} | {mg} | {total} | {m0} | {m1} | {m2} | {m3} | {m4} | {m5} | {m6} |")
    
    report_lines.extend([
        "",
        "---",
        "",
        "## 4期2Q 詳細データ",
        "",
        "| 生徒名 | 担当MG | 卒業時投稿数 | 0m | 1m | 2m | 3m | 4m | 5m | 6m |",
        "|--------|--------|-------------|----|----|----|----|----|----|----|",
    ])
    
    for _, row in q2_display.iterrows():
        name = str(row["生徒名"])
        mg = str(row["担当MG"]) if pd.notna(row["担当MG"]) else ""
        total = row["卒業時投稿数"]
        m0 = row["0m"]
        m1 = row["1m"]
        m2 = row["2m"]
        m3 = row["3m"]
        m4 = row["4m"]
        m5 = row["5m"]
        m6 = row["6m"]
        report_lines.append(f"| {name} | {mg} | {total} | {m0} | {m1} | {m2} | {m3} | {m4} | {m5} | {m6} |")
    
    report_lines.extend([
        "",
        "---",
        "",
        "## MG別分析",
        "",
        f"### 全期間 MG別 卒業時平均投稿数（全卒業生{len(all_graduates_df)}名）",
        "",
        "| 担当MG | 人数 | 平均投稿数 | 合計投稿数 | 標準偏差 | 最小値 | 最大値 |",
        "|--------|------|-----------|-----------|---------|--------|--------|",
    ])
    
    for mg, row in mg_all.iterrows():
        std_val = f"{row['標準偏差']:.2f}" if pd.notna(row['標準偏差']) else "-"
        report_lines.append(
            f"| {mg} | {int(row['人数'])} | **{row['平均投稿数']}** | {int(row['合計投稿数'])} | "
            f"{std_val} | {int(row['最小値'])} | {int(row['最大値'])} |"
        )
    
    report_lines.extend([
        "",
        "### 4期1Q MG別 卒業時平均投稿数",
        "",
        "| 担当MG | 人数 | 平均投稿数 | 合計投稿数 | 標準偏差 | 最小値 | 最大値 |",
        "|--------|------|-----------|-----------|---------|--------|--------|",
    ])
    
    for mg, row in mg_q1.iterrows():
        std_val = f"{row['標準偏差']:.2f}" if pd.notna(row['標準偏差']) else "-"
        report_lines.append(
            f"| {mg} | {int(row['人数'])} | **{row['平均投稿数']}** | {int(row['合計投稿数'])} | "
            f"{std_val} | {int(row['最小値'])} | {int(row['最大値'])} |"
        )
    
    report_lines.extend([
        "",
        "### 4期2Q MG別 卒業時平均投稿数",
        "",
        "| 担当MG | 人数 | 平均投稿数 | 合計投稿数 | 標準偏差 | 最小値 | 最大値 |",
        "|--------|------|-----------|-----------|---------|--------|--------|",
    ])
    
    for mg, row in mg_q2.iterrows():
        std_val = f"{row['標準偏差']:.2f}" if pd.notna(row['標準偏差']) else "-"
        report_lines.append(
            f"| {mg} | {int(row['人数'])} | **{row['平均投稿数']}** | {int(row['合計投稿数'])} | "
            f"{std_val} | {int(row['最小値'])} | {int(row['最大値'])} |"
        )
    
    report_lines.extend([
        "",
        "---",
        "",
        "## 相関分析: 全期間 vs 4期1Q・2Q",
        "",
        "### 相関係数",
        "",
        "| 比較項目 | 相関係数 | 対象MG数 |",
        "|---------|---------|---------|",
    ])
    
    if corr_all_q1 is not None:
        report_lines.append(f"| 全期間 vs 1Q | **{corr_all_q1:.3f}** | {len(mg_all_vs_q1)} |")
    else:
        report_lines.append(f"| 全期間 vs 1Q | データ不足 | {len(mg_all_vs_q1)} |")
    
    if corr_all_q2 is not None:
        report_lines.append(f"| 全期間 vs 2Q | **{corr_all_q2:.3f}** | {len(mg_all_vs_q2)} |")
    else:
        report_lines.append(f"| 全期間 vs 2Q | データ不足 | {len(mg_all_vs_q2)} |")
    
    if corr_q1_q2 is not None:
        report_lines.append(f"| 1Q vs 2Q | **{corr_q1_q2:.3f}** | {len(mg_q1_vs_q2)} |")
    else:
        report_lines.append(f"| 1Q vs 2Q | データ不足 | {len(mg_q1_vs_q2)} |")
    
    report_lines.extend([
        "",
        "### MG別比較表（全期間・1Q・2Q）",
        "",
        "| 担当MG | 全期間_平均投稿数 | 全期間_人数 | 1Q_平均投稿数 | 1Q_人数 | 2Q_平均投稿数 | 2Q_人数 |",
        "|--------|-----------------|------------|-------------|---------|-------------|---------|",
    ])
    
    for mg, row in mg_comparison.iterrows():
        all_avg = f"{row['全期間_平均投稿数']:.2f}" if pd.notna(row['全期間_平均投稿数']) else "-"
        all_count = int(row['全期間_人数']) if pd.notna(row['全期間_人数']) else "-"
        q1_avg = f"{row['1Q_平均投稿数']:.2f}" if pd.notna(row['1Q_平均投稿数']) else "-"
        q1_count = int(row['1Q_人数']) if pd.notna(row['1Q_人数']) else "-"
        q2_avg = f"{row['2Q_平均投稿数']:.2f}" if pd.notna(row['2Q_平均投稿数']) else "-"
        q2_count = int(row['2Q_人数']) if pd.notna(row['2Q_人数']) else "-"
        report_lines.append(
            f"| {mg} | {all_avg} | {all_count} | {q1_avg} | {q1_count} | {q2_avg} | {q2_count} |"
        )
    
    report_lines.extend([
        "",
        "---",
        "",
        "## データ出所・定義",
        "",
        "- **ファイル**: `コミットプラン (4).xlsx` の「新 月次投稿数」シート",
        "- **卒業時投稿数**: 0-6ヶ月目の合計投稿数",
        "- **0m-6m**: 各月の投稿数（データがない場合は「ー」）",
        "- **4期1Q**: 2025年8月・9月・10月卒業生",
        "- **4期2Q**: 2025年11月・12月・2026年1月卒業生",
        "",
        "---",
        "*出力: 4期1Q_2Q_卒業時平均投稿数比較.py*",
    ])
    
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    
    print(f"出力完了: {OUTPUT_PATH}")
    print(f"分析結果: {REPORT_PATH}")
    print("\n【4期1Q vs 4期2Q 卒業時平均投稿数比較】")
    if isinstance(q1_avg, (int, float)):
        print(f"  4期1Q: {q1_avg}投稿（{q1_total_count}名）")
    else:
        print(f"  4期1Q: {q1_avg}（{q1_total_count}名）")
    if isinstance(q2_avg, (int, float)):
        print(f"  4期2Q: {q2_avg}投稿（{q2_total_count}名）")
    else:
        print(f"  4期2Q: {q2_avg}（{q2_total_count}名）")
    if isinstance(q1_avg, (int, float)) and isinstance(q2_avg, (int, float)):
        print(f"  差（2Q - 1Q）: {round(q2_avg - q1_avg, 2):+.2f}投稿")
    print(f"\n【MG別分析（全期間・全卒業生{len(all_graduates_df)}名）】")
    for mg, row in mg_all.iterrows():
        print(f"  {mg}: {row['平均投稿数']}投稿（{int(row['人数'])}名）")
    
    print(f"\n【相関分析】")
    if corr_all_q1 is not None:
        print(f"  全期間 vs 1Q: {corr_all_q1:.3f}（対象MG数: {len(mg_all_vs_q1)}）")
    if corr_all_q2 is not None:
        print(f"  全期間 vs 2Q: {corr_all_q2:.3f}（対象MG数: {len(mg_all_vs_q2)}）")
    if corr_q1_q2 is not None:
        print(f"  1Q vs 2Q: {corr_q1_q2:.3f}（対象MG数: {len(mg_q1_vs_q2)}）")


if __name__ == "__main__":
    main()
