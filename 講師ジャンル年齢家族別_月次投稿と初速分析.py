# -*- coding: utf-8 -*-
"""
「新 月次投稿数」シートの X,Y,Z,AA列（講師・ジャンル・年齢・家族構成）ごとに
・平均月間投稿数の傾向
・投稿までの初速（何ヶ月目に初投稿か）
を分析する。
"""
import pandas as pd
import numpy as np
from pathlib import Path

INPUT_PATH = Path(__file__).parent / "コミットプラン (4).xlsx"
OUTPUT_DIR = Path(__file__).parent / "分析結果"
SHEET = "新 月次投稿数"

# 列インデックス（0始まり）
COL_NO = 0
COL_AVG_MONTHLY = 12   # 個人平均月間投稿数
COL_MONTH_START = 15   # 0ヶ月目
COL_MONTH_END = 21     # 6ヶ月目
COL_INSTRUCTOR = 23    # 講師 (X)
COL_GENRE = 24         # ジャンル (Y)
COL_AGE = 25           # 年齢 (Z)
COL_FAMILY = 26        # 家族構成 (AA)
DATA_START_ROW = 11


def to_num(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ("ー", "－", "-", "", "nan"):
        return None
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def first_post_month(row):
    """初速: 初めて投稿数>0になった月の番号（0=0ヶ月目, 1=1ヶ月目, ...）。投稿なしは np.nan"""
    for m in range(COL_MONTH_START, COL_MONTH_END + 1):
        v = to_num(row[m])
        if v is not None and v > 0:
            return m - COL_MONTH_START  # 0ヶ月目→0, 1ヶ月目→1
    return np.nan


def load_data():
    df = pd.read_excel(INPUT_PATH, sheet_name=SHEET, header=None)
    rows = []
    for i in range(DATA_START_ROW, len(df)):
        no_ = df.iloc[i, COL_NO]
        if pd.isna(no_) or str(no_).strip() == "":
            continue
        try:
            int(float(no_))
        except (ValueError, TypeError):
            continue
        row = df.iloc[i]
        avg = row[COL_AVG_MONTHLY]
        avg_val = None
        if pd.notna(avg):
            try:
                avg_val = float(avg)
            except (ValueError, TypeError):
                pass
        instructor = row[COL_INSTRUCTOR]
        genre = row[COL_GENRE]
        age = row[COL_AGE]
        family = row[COL_FAMILY]
        first_month = first_post_month(row)
        rows.append({
            "no": no_,
            "講師": instructor if pd.notna(instructor) and str(instructor).strip() else "未入力",
            "ジャンル": genre if pd.notna(genre) and str(genre).strip() else "未入力",
            "年齢": age if pd.notna(age) and str(age).strip() else "未入力",
            "家族構成": family if pd.notna(family) and str(family).strip() else "未入力",
            "個人平均月間投稿数": avg_val,
            "初速_初投稿月": first_month,
        })
    return pd.DataFrame(rows)


def aggregate_by(df, group_col, label):
    """group_col ごとに 平均月間投稿数 と 初速 を集計"""
    g = df.groupby(group_col, dropna=False).agg(
        人数=("no", "count"),
        平均月間投稿数=("個人平均月間投稿数", lambda x: x.dropna().mean()),
        平均月間投稿数_件数=("個人平均月間投稿数", lambda x: x.notna().sum()),
        初速_平均ヶ月目=("初速_初投稿月", "mean"),
        初速_中央値=("初速_初投稿月", "median"),
        初速_未投稿数=("初速_初投稿月", lambda x: x.isna().sum()),
    ).reset_index()
    g = g.rename(columns={group_col: label})
    return g


def main():
    df = load_data()
    print(f"総レコード数: {len(df)}")
    print()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUTPUT_DIR / "講師ジャンル年齢家族別_月次投稿と初速分析.xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        # 講師別
        by_instructor = aggregate_by(df, "講師", "講師")
        by_instructor = by_instructor.sort_values("平均月間投稿数", ascending=False)
        by_instructor.to_excel(w, sheet_name="講師別", index=False)
        print("【講師別】平均月間投稿数・初速")
        print(by_instructor.to_string(index=False))
        print()

        # ジャンル別
        by_genre = aggregate_by(df, "ジャンル", "ジャンル")
        by_genre = by_genre.sort_values("平均月間投稿数", ascending=False)
        by_genre.to_excel(w, sheet_name="ジャンル別", index=False)
        print("【ジャンル別】平均月間投稿数・初速")
        print(by_genre.to_string(index=False))
        print()

        # 年齢別
        by_age = aggregate_by(df, "年齢", "年齢")
        # 年齢順に並べる
        age_order = ["10〜19", "20〜29", "30〜39", "40〜49", "50〜59", "60〜", "不明", "未入力"]
        by_age["_order"] = by_age["年齢"].astype(str).map(lambda x: age_order.index(x) if x in age_order else 99)
        by_age = by_age.sort_values("_order").drop(columns=["_order"])
        by_age.to_excel(w, sheet_name="年齢別", index=False)
        print("【年齢別】平均月間投稿数・初速")
        print(by_age.to_string(index=False))
        print()

        # 家族構成別
        by_family = aggregate_by(df, "家族構成", "家族構成")
        by_family = by_family.sort_values("平均月間投稿数", ascending=False)
        by_family.to_excel(w, sheet_name="家族構成別", index=False)
        print("【家族構成別】平均月間投稿数・初速")
        print(by_family.to_string(index=False))
        print()

        # 生データ（サマリ用）
        df.to_excel(w, sheet_name="元データサマリ", index=False)

    print(f"出力: {out_path}")
    return df, by_instructor, by_genre, by_age, by_family


if __name__ == "__main__":
    main()
