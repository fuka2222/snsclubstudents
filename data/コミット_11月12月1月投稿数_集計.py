# -*- coding: utf-8 -*-
"""
KGI「特進コース(コミットコース)全体の卒業時平均投稿数80」に対する、
2025年11月・12月・2026年1月の実績を算出するスクリプト。

【ロジック】
- 「セッション実施状況管理」の初回の通常セッション日(W列=col22)と照らし合わせ、
  入会タイミングを考慮してコホートを絞り込む
- 6ヶ月目(V列)にデータがある人は卒業済み・古いデータのため除外
- 初回セッション 2025/6/1〜2026/1/31 の生徒のみ対象（分析期間中にアクティブな生徒）
"""
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

# 入力パス（202602分析フォルダ or Downloads）
BASE = Path(__file__).parent.parent
INPUT_PATH = BASE / "コミットプラン (4).xlsx"
if not INPUT_PATH.exists():
    INPUT_PATH = Path.home() / "Downloads" / "コミットプラン (4).xlsx"
OUTPUT_PATH = Path(__file__).parent / "コミット_11月12月1月投稿数_集計結果.xlsx"
REPORT_PATH = Path(__file__).parent.parent / "分析結果" / "4期1Q_KGI_コミット_11月12月1月結果.md"

# セッション実施状況管理
SESS_HEADER_ROW = 9
SESS_DATA_START = 10
SESS_COL_NO = 0
SESS_COL_NAME = 7
SESS_COL_FIRST_NORMAL = 22  # W列: 初回の通常セッション日

# 新 月次投稿数
MONTH_HEADER_ROW = 10
MONTH_DATA_START = 11
MONTH_COL_NO = 0
MONTH_COL_NAME = 4
MONTH_COL_0M = 15   # P列: 0ヶ月目
MONTH_COL_6M = 21   # V列: 6ヶ月目

# 分析期間・コホート
TARGET_MONTHS = [(2025, 11), (2025, 12), (2026, 1)]
COHORT_START = datetime(2025, 6, 1)
COHORT_END = datetime(2026, 1, 31)

# KGI目標（月別想定: 0,6,15,16,16,16,11 → 11月=2ヶ月目想定6, 12月=3ヶ月目15, 1月=4ヶ月目16）
TARGET_BY_MONTH = {  # カレンダー月 -> 目標投稿数
    (2025, 11): 6,
    (2025, 12): 15,
    (2026, 1): 16,
}


def to_num(val):
    if pd.isna(val):
        return 0
    s = str(val).strip()
    if s in ("ー", "－", "-", ""):
        return 0
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def main():
    df_sess = pd.read_excel(INPUT_PATH, sheet_name="セッション実施状況管理", header=None)
    df_month = pd.read_excel(INPUT_PATH, sheet_name="新 月次投稿数", header=None)

    # no. -> 初回の通常セッション日
    sess_map = {}
    for i in range(SESS_DATA_START, len(df_sess)):
        no_ = df_sess.iloc[i, SESS_COL_NO]
        if pd.isna(no_):
            continue
        try:
            no_int = int(float(no_))
        except (ValueError, TypeError):
            continue
        dt = df_sess.iloc[i, SESS_COL_FIRST_NORMAL]
        if pd.notna(dt) and isinstance(dt, (pd.Timestamp, datetime)):
            sess_map[no_int] = pd.Timestamp(dt)
        else:
            sess_map[no_int] = None

    rows = []
    for i in range(MONTH_DATA_START, len(df_month)):
        no_ = df_month.iloc[i, MONTH_COL_NO]
        if pd.isna(no_):
            continue
        try:
            no_int = int(float(no_))
        except (ValueError, TypeError):
            continue

        first_sess = sess_map.get(no_int)
        if first_sess is None or pd.isna(first_sess):
            continue

        # コホート絞り込み: 初回セッション 2025/6/1 〜 2026/1/31
        if first_sess < pd.Timestamp(COHORT_START) or first_sess > pd.Timestamp(COHORT_END):
            continue

        # 0ヶ月目 = 初回セッションの月。カレンダー月 -> ヶ月目
        start_year, start_month = first_sess.year, first_sess.month

        def months_diff(y1, m1, y2, m2):
            return (y2 - y1) * 12 + (m2 - m1)

        nov_val = dec_val = jan_val = 0
        for (y, m), val_name in [
            ((2025, 11), "nov_val"),
            ((2025, 12), "dec_val"),
            ((2026, 1), "jan_val"),
        ]:
            diff = months_diff(start_year, start_month, y, m)
            if 0 <= diff <= 6:
                col_idx = MONTH_COL_0M + diff
                val = to_num(df_month.iloc[i, col_idx])
                if val_name == "nov_val":
                    nov_val = val
                elif val_name == "dec_val":
                    dec_val = val
                else:
                    jan_val = val
            # diff < 0: まだ開始前 → 0
            # diff > 6: 6ヶ月目超（卒業後）→ 0

        name = df_month.iloc[i, MONTH_COL_NAME]
        rows.append({
            "no.": no_int,
            "生徒名": name,
            "初回セッション日": first_sess.strftime("%Y-%m-%d"),
            "2025年11月": nov_val,
            "2025年12月": dec_val,
            "2026年1月": jan_val,
            "3ヶ月合計": nov_val + dec_val + jan_val,
        })

    result_df = pd.DataFrame(rows)

    # 全体集計（チーム別ではなく全生徒の平均）
    n = len(result_df)
    total_nov = result_df["2025年11月"].sum()
    total_dec = result_df["2025年12月"].sum()
    total_jan = result_df["2026年1月"].sum()
    avg_nov = round(total_nov / n, 2) if n > 0 else 0
    avg_dec = round(total_dec / n, 2) if n > 0 else 0
    avg_jan = round(total_jan / n, 2) if n > 0 else 0
    avg_3m = round((total_nov + total_dec + total_jan) / n, 2) if n > 0 else 0

    # Excel出力
    summary_df = pd.DataFrame([
        {"項目": "対象生徒数", "値": n},
        {"項目": "2025年11月_合計", "値": total_nov},
        {"項目": "2025年11月_1人あたり平均", "値": avg_nov},
        {"項目": "2025年12月_合計", "値": total_dec},
        {"項目": "2025年12月_1人あたり平均", "値": avg_dec},
        {"項目": "2026年1月_合計", "値": total_jan},
        {"項目": "2026年1月_1人あたり平均", "値": avg_jan},
        {"項目": "3ヶ月合計_1人あたり平均", "値": avg_3m},
    ])

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as w:
        result_df.to_excel(w, sheet_name="生徒別", index=False)
        summary_df.to_excel(w, sheet_name="全体サマリ", index=False)

    # 分析結果をキャプチャ形式でMarkdown出力
    report = f"""# 4期目1Q KGI 特進コース(コミットコース) 11月・12月・1月 結果

## KGI: 特進コース(コミットコース)全体の卒業時平均投稿数 80

※60投稿以上と80投稿以上の万垢率が30%のため、特進コースの目標投稿数を設定

---

## 月別実績（コホート: 初回セッション 2025/6/1〜2026/1/31 の生徒）

| 月 | 1人あたり平均投稿数 | 目標 | GAP |
|----|---------------------|------|-----|
| **2025年11月** | **{avg_nov}** | {TARGET_BY_MONTH.get((2025, 11), "-")} | {avg_nov - TARGET_BY_MONTH.get((2025, 11), 0):+.2f} |
| **2025年12月** | **{avg_dec}** | {TARGET_BY_MONTH.get((2025, 12), "-")} | {avg_dec - TARGET_BY_MONTH.get((2025, 12), 0):+.2f} |
| **2026年1月** | **{avg_jan}** | {TARGET_BY_MONTH.get((2026, 1), "-")} | {avg_jan - TARGET_BY_MONTH.get((2026, 1), 0):+.2f} |

---

## サマリ

- **対象生徒数**: {n}名（初回セッション日で絞り込み済み）
- **3ヶ月合計 1人あたり平均**: {avg_3m}投稿
- **KGI目標 卒業時80投稿** に対する進捗指標として、月別目標(6→15→16)との比較を参照

---
*出力: コミット_11月12月1月投稿数_集計.py*
"""
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"出力完了: {OUTPUT_PATH}")
    print(f"分析結果: {REPORT_PATH}")
    print("\n【KGI 特進コース 11月・12月・1月 結果】")
    print(f"  対象生徒数: {n}名")
    print(f"  2025年11月 1人あたり平均: {avg_nov} (目標{TARGET_BY_MONTH.get((2025, 11))})")
    print(f"  2025年12月 1人あたり平均: {avg_dec} (目標{TARGET_BY_MONTH.get((2025, 12))})")
    print(f"  2026年1月 1人あたり平均: {avg_jan} (目標{TARGET_BY_MONTH.get((2026, 1))})")
    print(f"  3ヶ月合計 1人あたり平均: {avg_3m}")


if __name__ == "__main__":
    main()
