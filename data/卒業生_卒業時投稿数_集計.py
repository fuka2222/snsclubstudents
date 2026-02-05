# -*- coding: utf-8 -*-
"""
卒業生に限定し、新 月次投稿数タブのP列〜V列の合計＝卒業時投稿数として、
累計の結果を算出するスクリプト。

【月別コホート】ユーザー指定の「新たに増えた卒業生」名簿に基づく
"""
import pandas as pd
from pathlib import Path

BASE = Path(__file__).parent.parent
INPUT_PATH = BASE / "コミットプラン (4).xlsx"
if not INPUT_PATH.exists():
    INPUT_PATH = Path.home() / "Downloads" / "コミットプラン (4).xlsx"
OUTPUT_PATH = Path(__file__).parent / "卒業生_卒業時投稿数_集計結果.xlsx"
REPORT_PATH = Path(__file__).parent.parent / "分析結果" / "卒業生_卒業時投稿数_累計結果.md"

# 月別「新たに増えた卒業生」名簿
NOV_GRADUATES = [
    "ひらやまみか", "しむらまなぶ", "やまぐちちづる", "のぶとうまさこ",
    "かげやまこゆき", "こうごたかひろ", "らぶひとみ", "かわらまき",
]
DEC_GRADUATES = [
    "かきのきまゆみ", "ながいけいこ", "ろばーつあゆみ", "じくまるみほ",
    "のざきせいか", "はらだたかよ", "いまむらやすえ", "おおつばきかなこ",
    "すのうちなお", "またよししょうや",
]
JAN_GRADUATES = [
    "よしだえみ", "たむらやすのり", "なかおしょうや", "まえだりお", "みやざとせいぎ",
]


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


def norm(s):
    return str(s).strip().replace(" ", "").replace("　", "").lower()


def name_matches(a, b):
    """名前の部分一致（スペース除去）"""
    na, nb = norm(a), norm(b)
    return na == nb or (len(na) >= 3 and na in nb) or (len(nb) >= 3 and nb in na)


def main():
    df = pd.read_excel(INPUT_PATH, sheet_name="新 月次投稿数", header=None)

    # 卒業生のみ抽出（在学=卒業）。1月卒業は在学中の可能性あり → 名簿にいれば含める
    all_grads = []
    for i in range(11, len(df)):
        status = str(df.iloc[i, 2]) if pd.notna(df.iloc[i, 2]) else ""
        name = str(df.iloc[i, 4]).strip() if pd.notna(df.iloc[i, 4]) else ""

        # 卒業生 または 1月卒業名簿にいる在学生（1月に卒業したばかり）
        is_jan_grad = any(name_matches(name, n) for n in JAN_GRADUATES)
        if "卒業" in status or is_jan_grad:
            pv_sum = sum(to_num(df.iloc[i, c]) for c in range(15, 22))
            all_grads.append({"生徒名": name, "卒業時投稿数": pv_sum, "ステータス": status})

    # 月別コホートにマッチ
    def match_cohort(names, grads):
        matched = []
        for g in grads:
            for n in names:
                if name_matches(g["生徒名"], n):
                    matched.append(g["卒業時投稿数"])
                    break
        return matched

    nov_vals = match_cohort(NOV_GRADUATES, all_grads)
    dec_vals = match_cohort(DEC_GRADUATES, all_grads)
    jan_vals = match_cohort(JAN_GRADUATES, all_grads)

    # 重複除去（1人1回・先にマッチした方）
    def dedup_match(names, grads):
        used = set()
        result = []
        for n in names:
            for g in grads:
                if g["生徒名"] in used:
                    continue
                if name_matches(g["生徒名"], n):
                    result.append(g["卒業時投稿数"])
                    used.add(g["生徒名"])
                    break
        return result

    nov_vals = dedup_match(NOV_GRADUATES, all_grads)
    dec_vals = dedup_match(DEC_GRADUATES, all_grads)
    jan_vals = dedup_match(JAN_GRADUATES, all_grads)

    # 集計
    n_all = len(all_grads)
    avg_all = round(sum(g["卒業時投稿数"] for g in all_grads) / n_all, 2) if n_all else 0

    avg_nov = round(sum(nov_vals) / len(nov_vals), 2) if nov_vals else 0
    avg_dec = round(sum(dec_vals) / len(dec_vals), 2) if dec_vals else 0
    avg_jan = round(sum(jan_vals) / len(jan_vals), 2) if jan_vals else 0

    # Excel出力
    detail = pd.DataFrame(all_grads)
    summary = pd.DataFrame([
        {"項目": "卒業生総数", "値": n_all},
        {"項目": "卒業時全体_平均投稿数", "値": avg_all},
        {"項目": "11月新規卒業_人数", "値": len(nov_vals)},
        {"項目": "11月新規卒業_平均投稿数", "値": avg_nov},
        {"項目": "12月新規卒業_人数", "値": len(dec_vals)},
        {"項目": "12月新規卒業_平均投稿数", "値": avg_dec},
        {"項目": "1月新規卒業_人数", "値": len(jan_vals)},
        {"項目": "1月新規卒業_平均投稿数", "値": avg_jan},
    ])
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as w:
        detail.to_excel(w, sheet_name="卒業生一覧", index=False)
        summary.to_excel(w, sheet_name="サマリ", index=False)

    # レポート出力（ユーザー指定フォーマット）
    report = f"""# 卒業生 卒業時投稿数 累計結果

## 卒業時全体 平均投稿数

**{avg_all}投稿**（卒業生{n_all}名、過去卒業生含む）

---

## 月別「新たに増えた卒業生」の平均投稿数

| 月 | 人数 | 平均投稿数 |
|----|------|------------|
| **11月** | {len(nov_vals)}名 | **{avg_nov}投稿** |
| **12月** | {len(dec_vals)}名 | **{avg_dec}投稿** |
| **1月** | {len(jan_vals)}名 | **{avg_jan}投稿** |

---

### 回答フォーマット（キャプチャ用）

```
卒業時全体 平均投稿数: {avg_all}投稿

11月: {avg_nov}投稿
12月: {avg_dec}投稿
1月: {avg_jan}投稿
```

---
*P列〜V列（0〜6ヶ月目）の合計＝卒業時投稿数として算出*
"""
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"出力: {OUTPUT_PATH}")
    print(f"レポート: {REPORT_PATH}")
    print()
    print("【卒業生 卒業時投稿数 累計結果】")
    print(f"  卒業時全体 平均投稿数: {avg_all}投稿（{n_all}名）")
    print(f"  11月: {avg_nov}投稿（{len(nov_vals)}名）")
    print(f"  12月: {avg_dec}投稿（{len(dec_vals)}名）")
    print(f"  1月: {avg_jan}投稿（{len(jan_vals)}名）")


if __name__ == "__main__":
    main()
