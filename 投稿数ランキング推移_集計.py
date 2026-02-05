# -*- coding: utf-8 -*-
"""
投稿数ランキング推移（11月1日〜2026年1月31日）をチーム別に集計し、
見やすいExcelに出力するスクリプト。
"""
import pandas as pd
import numpy as np
from pathlib import Path

EXCEL_PATH = Path(__file__).parent / "［最新版］mg_monthly_analysis_results_v1.1.xlsx"
OUTPUT_PATH = Path(__file__).parent / "投稿数ランキング推移_11月〜1月_チーム別.xlsx"

def clean_team(s):
    if pd.isna(s): return s
    return str(s).strip()

def main():
    xl = pd.ExcelFile(EXCEL_PATH)
    df = pd.read_excel(xl, sheet_name="PP_Rawdata", header=2)

    df["チーム名"] = df["チーム名"].apply(clean_team)
    df = df[df["チーム名"].notna() & (df["チーム名"] != "") & (df["チーム名"] != "全体")].copy()

    session_cols = [
        "1回目実施日", "2回目実施日", "3回目実施日",
        "4回目実施日", "5回目実施日", "6回目実施日",
    ]
    incr_cols = [
        "前回からの増加投稿数", "前回からの増加投稿数.1", "前回からの増加投稿数.2",
        "前回からの増加投稿数.3", "前回からの増加投稿数.4", "前回からの増加投稿数.5",
    ]

    for c in session_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in incr_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    rows = []
    for _, row in df.iterrows():
        team = row["チーム名"]
        for sess, incr in zip(session_cols, incr_cols):
            dt = row[sess]
            val = row[incr]
            if pd.isna(dt) or pd.isna(val):
                continue
            try:
                v = float(val)
            except (TypeError, ValueError):
                continue
            rows.append({
                "チーム名": team,
                "年": dt.year,
                "月": dt.month,
                "投稿増加": v,
            })

    monthly = pd.DataFrame(rows)
    monthly["チーム名"] = monthly["チーム名"].str.strip()
    agg = monthly.groupby(["チーム名", "年", "月"], as_index=False)["投稿増加"].sum()

    target_months = [(2025, 11), (2025, 12), (2026, 1)]
    month_labels = ["2025年11月", "2025年12月", "2026年1月"]

    # 各月の投稿数と順位
    result_list = []
    teams = agg["チーム名"].unique().tolist()
    teams = sorted([t for t in teams if t and str(t).strip()])

    for team in teams:
        row_data = {"チーム名": team}
        for (y, m), label in zip(target_months, month_labels):
            sub = agg[(agg["チーム名"] == team) & (agg["年"] == y) & (agg["月"] == m)]
            count = sub["投稿増加"].sum() if len(sub) else 0
            row_data[f"{label}_投稿数"] = int(count)
        result_list.append(row_data)

    result_df = pd.DataFrame(result_list)

    # 順位を付与（各月でランキング）
    for label in month_labels:
        col = f"{label}_投稿数"
        result_df[f"{label}_順位"] = result_df[col].rank(ascending=False, method="min").astype(int)

    # 3ヶ月合計・平均順位
    result_df["3ヶ月合計投稿数"] = (
        result_df["2025年11月_投稿数"]
        + result_df["2025年12月_投稿数"]
        + result_df["2026年1月_投稿数"]
    )
    result_df["3ヶ月合計順位"] = result_df["3ヶ月合計投稿数"].rank(ascending=False, method="min").astype(int)
    result_df["平均順位"] = (
        result_df["2025年11月_順位"]
        + result_df["2025年12月_順位"]
        + result_df["2026年1月_順位"]
    ).round(1)

    # 列順を整理
    cols = ["チーム名"]
    for label in month_labels:
        cols.append(f"{label}_投稿数")
        cols.append(f"{label}_順位")
    cols.extend(["3ヶ月合計投稿数", "3ヶ月合計順位", "平均順位"])
    result_df = result_df[cols]

    # Excel出力（複数シートで見やすく）
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        # シート1: 推移一覧（メイン）
        result_df.to_excel(writer, sheet_name="投稿数ランキング推移_一覧", index=False)

        # シート2: 月別ランキング（縦持ち）
        rank_rows = []
        for (y, m), label in zip(target_months, month_labels):
            sub = agg[(agg["年"] == y) & (agg["月"] == m)]
            sub = sub.sort_values("投稿増加", ascending=False).reset_index(drop=True)
            sub["順位"] = range(1, len(sub) + 1)
            sub["対象月"] = label
            rank_rows.append(sub[["対象月", "順位", "チーム名", "投稿増加"]])
        rank_df = pd.concat(rank_rows, ignore_index=True)
        rank_df = rank_df.rename(columns={"投稿増加": "投稿数"})
        rank_df.to_excel(writer, sheet_name="月別ランキング詳細", index=False)

        # シート3: チーム別サマリ（チームごとに行で見る）
        summary_rows = []
        for team in teams:
            summary_rows.append({
                "チーム名": team,
                "11月投稿数": result_df.loc[result_df["チーム名"] == team, "2025年11月_投稿数"].values[0],
                "11月順位": result_df.loc[result_df["チーム名"] == team, "2025年11月_順位"].values[0],
                "12月投稿数": result_df.loc[result_df["チーム名"] == team, "2025年12月_投稿数"].values[0],
                "12月順位": result_df.loc[result_df["チーム名"] == team, "2025年12月_順位"].values[0],
                "1月投稿数": result_df.loc[result_df["チーム名"] == team, "2026年1月_投稿数"].values[0],
                "1月順位": result_df.loc[result_df["チーム名"] == team, "2026年1月_順位"].values[0],
                "3ヶ月合計": result_df.loc[result_df["チーム名"] == team, "3ヶ月合計投稿数"].values[0],
                "合計順位": result_df.loc[result_df["チーム名"] == team, "3ヶ月合計順位"].values[0],
            })
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="チーム別サマリ", index=False)

    print(f"出力先: {OUTPUT_PATH}")
    print("\n【投稿数ランキング推移 サマリ】")
    print(result_df.to_string(index=False))
    return OUTPUT_PATH

if __name__ == "__main__":
    main()
