# -*- coding: utf-8 -*-
"""
プレミアムプラス（PP）の卒業した生徒について、
・各月の「全体」（その月までに卒業した全員）の平均卒業時投稿数
・「当月卒業生のみ」の平均卒業時投稿数
を算出するスクリプト。画像と同じ形式で出力。

【データ出所】［最新版］mg_monthly_analysis_results_v1.1.xlsx の PP_Rawdata
【卒業の定義】6回目実施日が入っている＝6回セッション完了＝卒業とする
【卒業時投稿数】合計投稿数列
【卒業月】6回目実施日の年月
"""
import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent
INPUT_PATH = BASE / "［最新版］mg_monthly_analysis_results_v1.1.xlsx"
OUTPUT_PATH = BASE / "プレミアムプラス_卒業生_月次平均卒業時投稿数_集計結果.xlsx"
REPORT_PATH = BASE.parent / "分析結果" / "プレミアムプラス_卒業生_月次平均卒業時投稿数.md"

# 画像形式で出す対象月（年月）
TARGET_MONTHS = [
    ("2025年11月", "2025-11"),
    ("2025年12月", "2025-12"),
    ("2026年1月", "2026-01"),
]


def to_num(val):
    if pd.isna(val):
        return None
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None


def main():
    df = pd.read_excel(INPUT_PATH, sheet_name="PP_Rawdata", header=2)

    # 卒業＝6回目実施日あり
    graduated = df[df["6回目実施日"].notna()].copy()
    graduated["卒業月"] = pd.to_datetime(graduated["6回目実施日"], errors="coerce").dt.to_period("M")
    graduated["卒業時投稿数"] = graduated["合計投稿数"].map(lambda x: to_num(x))

    # 卒業時投稿数が有効な行のみ（NaNは月次平均からは除外）
    valid = graduated[graduated["卒業時投稿数"].notna()].copy()
    valid["卒業月_str"] = valid["卒業月"].astype(str)

    # 月別集計（全期間）
    monthly = (
        valid.groupby("卒業月", as_index=False)
        .agg(
            卒業生数=("名前", "count"),
            卒業時投稿数_合計=("卒業時投稿数", "sum"),
        )
        .assign(
            平均卒業時投稿数=lambda x: (x["卒業時投稿数_合計"] / x["卒業生数"]).round(2)
        )
    )
    monthly["卒業月"] = monthly["卒業月"].astype(str)

    # 画像形式: 各月の「全体」と「当月卒業生のみ」
    rows_display = []
    for label, period in TARGET_MONTHS:
        # 全体＝その月までに卒業した全員（卒業月 <= period）
        up_to = valid[valid["卒業月_str"] <= period]
        n_zen = len(up_to)
        avg_zen = round(up_to["卒業時投稿数"].mean(), 2) if n_zen else 0
        # 当月卒業生のみ
        in_month = valid[valid["卒業月_str"] == period]
        n_tou = len(in_month)
        avg_tou = round(in_month["卒業時投稿数"].mean(), 2) if n_tou else 0
        rows_display.append({
            "月": label,
            "全体_平均投稿数": avg_zen,
            "全体_人数": n_zen,
            "当月卒業生のみ_平均投稿数": avg_tou,
            "当月卒業生のみ_人数": n_tou,
        })
    display_df = pd.DataFrame(rows_display)

    # 卒業生一覧（卒業月・名前・卒業時投稿数）
    detail = (
        valid[["卒業月", "名前", "担当MG", "チーム名", "6回目実施日", "卒業時投稿数"]]
        .copy()
    )
    detail["卒業月"] = detail["卒業月"].astype(str)

    # 全体サマリ
    n_all = len(valid)
    avg_all = round(valid["卒業時投稿数"].mean(), 2) if n_all else 0
    summary_df = pd.DataFrame([
        {"項目": "卒業生総数（6回目実施日あり・投稿数有効）", "値": n_all},
        {"項目": "卒業時投稿数_全体平均", "値": avg_all},
    ])

    # Excel出力
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as w:
        display_df.to_excel(w, sheet_name="月別_全体と当月卒業生のみ", index=False)
        monthly.to_excel(w, sheet_name="月別_平均卒業時投稿数", index=False)
        summary_df.to_excel(w, sheet_name="全体サマリ", index=False)
        detail.to_excel(w, sheet_name="卒業生一覧", index=False)

    # Markdownレポート（画像と同じ形式を先に）
    report_lines = [
        "# プレミアムプラス 卒業生 月次 平均卒業時投稿数",
        "",
        "## 月別（全体 ＋ 当月卒業生のみ）",
        "",
    ]
    for _, row in display_df.iterrows():
        report_lines.append(f"**{row['月']}**")
        report_lines.append(f"- 全体：{row['全体_平均投稿数']}投稿（{row['全体_人数']}人）")
        report_lines.append(f"- ▼当月卒業生のみ：{row['当月卒業生のみ_平均投稿数']}投稿（{row['当月卒業生のみ_人数']}人）")
        report_lines.append("")
    report_lines.extend([
        "---",
        "",
        "## データ出所・定義",
        "- **ファイル**: `［最新版］mg_monthly_analysis_results_v1.1.xlsx` の **PP_Rawdata**",
        "- **卒業**: 6回目実施日が入っている（6回セッション完了）",
        "- **卒業時投稿数**: 合計投稿数列",
        "- **全体**: その月までに卒業した全員の平均卒業時投稿数",
        "- **当月卒業生のみ**: その月に卒業した人のみの平均卒業時投稿数",
        "",
        "## 月別 平均卒業時投稿数（全期間）",
        "",
        "| 卒業月 | 卒業生数 | 平均卒業時投稿数 |",
        "|--------|----------|------------------|",
    ])
    for _, row in monthly.iterrows():
        report_lines.append(f"| {row['卒業月']} | {row['卒業生数']}名 | **{row['平均卒業時投稿数']}** |")
    report_lines.extend([
        "",
        "---",
        "*出力: プレミアムプラス_卒業生_月次平均卒業時投稿数_集計.py*",
    ])
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))

    print(f"出力: {OUTPUT_PATH}")
    print(f"レポート: {REPORT_PATH}")
    print()
    print("【プレミアムプラス 卒業生 月次 平均卒業時投稿数】（画像形式）")
    for _, row in display_df.iterrows():
        print(f"  {row['月']}")
        print(f"    全体：{row['全体_平均投稿数']}投稿（{row['全体_人数']}人）")
        print(f"    ▼当月卒業生のみ：{row['当月卒業生のみ_平均投稿数']}投稿（{row['当月卒業生のみ_人数']}人）")
    print()
    print(f"  卒業生総数（投稿数有効）: {n_all}名 / 全体平均: {avg_all}投稿")


if __name__ == "__main__":
    main()
