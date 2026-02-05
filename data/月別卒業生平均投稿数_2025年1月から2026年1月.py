# -*- coding: utf-8 -*-
"""
2025年1月から2026年1月までの各月について、
卒業生の平均投稿数と人数を集計するスクリプト。

【データ出所】コミットプラン (4).xlsx の「新 月次投稿数」シート
【卒業の定義】6回目実施日がある、または6ヶ月目のデータがある
【卒業時投稿数】0-6ヶ月目の合計投稿数
【卒業月】6回目実施日の年月、または6ヶ月目のデータがある月を推定
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
OUTPUT_PATH = Path(__file__).parent / "月別卒業生平均投稿数_2025年1月から2026年1月_結果.xlsx"
REPORT_PATH = Path(__file__).parent.parent / "分析結果" / "月別卒業生平均投稿数_2025年1月から2026年1月.md"

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

# 対象期間
TARGET_MONTHS = []
for year in [2025, 2026]:
    for month in range(1, 13):
        if year == 2025 and month < 1:
            continue
        if year == 2026 and month > 1:
            break
        TARGET_MONTHS.append(f"{year}-{month:02d}")


def normalize_name(name):
    """名前を正規化（全角スペース→半角スペース、小文字化など）"""
    if pd.isna(name):
        return ""
    s = str(name).strip()
    s = s.replace("　", " ")
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
        print("警告: 6回目実施日の列が見つかりません。6ヶ月目のデータで判定します。")
    
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
    
    # 全卒業生データを収集
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
        
        # 対象期間内の卒業生のみ
        if grad_month not in TARGET_MONTHS:
            continue
        
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
    
    if not all_graduates_results:
        print("エラー: 卒業生が見つかりませんでした。")
        return
    
    all_graduates_df = pd.DataFrame(all_graduates_results)
    
    # 月別集計
    monthly_summary = []
    for month in TARGET_MONTHS:
        month_graduates = all_graduates_df[all_graduates_df["卒業月"] == month]
        if len(month_graduates) > 0:
            avg_posts = round(month_graduates["卒業時投稿数"].mean(), 2)
            total_count = len(month_graduates)
        else:
            avg_posts = 0
            total_count = 0
        
        # 年月を表示用に変換
        year, month_num = month.split("-")
        month_label = f"{year}年{int(month_num)}月"
        
        monthly_summary.append({
            "年月": month,
            "月": month_label,
            "卒業生数": total_count,
            "平均投稿数": avg_posts,
        })
    
    monthly_summary_df = pd.DataFrame(monthly_summary)
    
    # Excel出力
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as w:
        monthly_summary_df.to_excel(w, sheet_name="月別集計", index=False)
        all_graduates_df.to_excel(w, sheet_name="卒業生一覧", index=False)
    
    # Markdownレポート
    report_lines = [
        "# 月別卒業生平均投稿数（2025年1月〜2026年1月）",
        "",
        "## 月別集計",
        "",
        "| 月 | 卒業生数 | 平均投稿数 |",
        "|----|---------|-----------|",
    ]
    
    for _, row in monthly_summary_df.iterrows():
        if row["卒業生数"] > 0:
            report_lines.append(f"| {row['月']} | {row['卒業生数']}名 | **{row['平均投稿数']}**投稿 |")
        else:
            report_lines.append(f"| {row['月']} | 0名 | - |")
    
    report_lines.extend([
        "",
        "---",
        "",
        "## データ出所・定義",
        "",
        "- **ファイル**: `コミットプラン (4).xlsx` の「新 月次投稿数」シート",
        "- **卒業の定義**: 6回目実施日がある、または6ヶ月目のデータがある",
        "- **卒業時投稿数**: 0-6ヶ月目の合計投稿数",
        "- **卒業月**: 6回目実施日の年月、または6ヶ月目のデータがある月を推定",
        "",
        "---",
        "*出力: 月別卒業生平均投稿数_2025年1月から2026年1月.py*",
    ])
    
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    
    print(f"出力完了: {OUTPUT_PATH}")
    print(f"分析結果: {REPORT_PATH}")
    print("\n【月別卒業生平均投稿数（2025年1月〜2026年1月）】")
    for _, row in monthly_summary_df.iterrows():
        if row["卒業生数"] > 0:
            print(f"  {row['月']}: {row['平均投稿数']}投稿（{row['卒業生数']}名）")
        else:
            print(f"  {row['月']}: 卒業生なし")


if __name__ == "__main__":
    main()
