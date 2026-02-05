# -*- coding: utf-8 -*-
"""
KGI「生徒一人当たり全期間合計平均投稿数」を、
コミット・プレミアムプラスの両方で算出する。

【データ出所】［最新版］mg_monthly_analysis_results_v1.1.xlsx
- コミット：コミットRawdata の「現在投稿数合計」
- プレミアムプラス：PP_Rawdata の「合計投稿数」
※有効な数値が入っている生徒のみで平均を算出
"""
import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent
INPUT_PATH = BASE / "［最新版］mg_monthly_analysis_results_v1.1.xlsx"
REPORT_PATH = BASE.parent / "分析結果" / "KGI_全期間合計平均投稿数_コミットとPP.md"


def to_num(val):
    if pd.isna(val):
        return None
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None


def main():
    # コミット
    df_cc = pd.read_excel(INPUT_PATH, sheet_name="コミットRawdata", header=0)
    cc_total = df_cc["現在投稿数合計"].map(to_num)
    valid_cc = cc_total.notna()
    n_cc = valid_cc.sum()
    avg_cc = round(cc_total[valid_cc].mean(), 2) if n_cc else 0

    # プレミアムプラス
    df_pp = pd.read_excel(INPUT_PATH, sheet_name="PP_Rawdata", header=2)
    pp_total = df_pp["合計投稿数"].map(to_num)
    valid_pp = pp_total.notna()
    n_pp = valid_pp.sum()
    avg_pp = round(pp_total[valid_pp].mean(), 2) if n_pp else 0

    # レポート
    report = f"""# KGI：生徒一人当たり全期間合計平均投稿数

| コース | 一人当たり全期間合計平均投稿数 | 対象生徒数 |
|--------|------------------------------|-------------|
| **コミット** | **{avg_cc}本** | {n_cc}人 |
| **プレミアムプラス** | **{avg_pp}本** | {n_pp}人 |

---

## 定義・出所

- **指標**: 生徒一人当たりの「全期間合計投稿数」の平均（有効な値がある生徒のみ）
- **コミット**: `［最新版］mg_monthly_analysis_results_v1.1.xlsx` の **コミットRawdata** → 列「現在投稿数合計」
- **プレミアムプラス**: 同上の **PP_Rawdata** → 列「合計投稿数」

---
*出力: KGI_全期間合計平均投稿数_コミットとPP.py*
"""
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"出力: {REPORT_PATH}")
    print()
    print("【KGI 生徒一人当たり全期間合計平均投稿数】")
    print(f"  コミット：{avg_cc}本（{n_cc}人）")
    print(f"  プレミアムプラス：{avg_pp}本（{n_pp}人）")


if __name__ == "__main__":
    main()
