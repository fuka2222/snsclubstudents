# -*- coding: utf-8 -*-
"""
チーム別の3ヶ月（11月・12月・1月）投稿数トレンドと、
チャットログの活発度・注意・訂正系キーワードを集計する。
マネジメント観点の分析用データを出力する。
"""
import pandas as pd
import re
from pathlib import Path

BASE = Path(__file__).parent
RANKING_XLSX = BASE / "投稿数ランキング推移_11月〜1月_チーム別.xlsx"
CHAT_DIR = BASE.parent.parent / "コーチングチーム5チーム分析" / "チーム別チャットログ"

# チーム名の対応（Excel表記 → チャットログファイル名）
TEAM_FILE_MAP = {
    "トミーT": "トミーチーム_4期1Q_チャットログ.md",
    "そたかT": "そたかチーム_4期1Q_チャットログ.md",
    "ちづるT": "ちづるチーム_4期1Q_チャットログ.md",
    "ゆりT": "ゆりチーム_4期1Q_チャットログ.md",
    "なつみT": "なつみチーム_4期1Q_チャットログ.md",
}

# 注意・訂正・指摘など「メンバーへの指揮」に関連しそうなキーワード（部分一致）
KEYWORDS_SUPERVISION = [
    "注意", "訂正", "指摘", "気をつけ", "改善して", "直して", "お願いします",
    "〜ください", "確認お願い", "入力漏れ", "遅くなり", "申し訳",
]

KEYWORDS_FEEDBACK = ["FB", "フィードバック", "振り返り", "良い点", "改善点"]


def load_posting_ranking():
    """投稿数ランキング推移 Excel を読み、チーム別の3ヶ月トレンドを返す。"""
    if not RANKING_XLSX.exists():
        return None
    df = pd.read_excel(RANKING_XLSX, sheet_name="投稿数ランキング推移_一覧")
    return df


def count_chat_blocks_and_keywords(file_path: Path):
    """
    チャットログファイルを読み、
    - 投稿ブロック数（### で始まる行 or 日付行の数）
    - 注意・訂正系キーワード出現回数
    を返す。
    """
    if not file_path.exists():
        return {"blocks": 0, "supervision": 0, "feedback": 0}
    text = file_path.read_text(encoding="utf-8")
    # 形式1: "123. ### 【MG】名前" のブロック数
    blocks_dot = len(re.findall(r"^\d+\.\s+###\s", text, re.MULTILINE))
    # 形式2: "— 2025/11/01" のような日付行（ゆりチーム形式）
    blocks_date = len(re.findall(r"—\s*\d{4}/\d{2}/\d{2}\s", text))
    blocks = blocks_dot if blocks_dot > 0 else blocks_date

    supervision = sum(text.count(kw) for kw in KEYWORDS_SUPERVISION)
    feedback = sum(text.count(kw) for kw in KEYWORDS_FEEDBACK)
    return {"blocks": blocks, "supervision": supervision, "feedback": feedback}


def main():
    df = load_posting_ranking()
    if df is None:
        print("投稿数ランキングExcelが見つかりません。先に 投稿数ランキング推移_集計.py を実行してください。")
        return

    month_cols = ["2025年11月_投稿数", "2025年12月_投稿数", "2026年1月_投稿数"]
    rank_cols = ["2025年11月_順位", "2025年12月_順位", "2026年1月_順位"]

    print("=== チーム別 3ヶ月投稿数・トレンド ===\n")
    rows = []
    for _, row in df.iterrows():
        team = row["チーム名"]
        nov, dec, jan = row[month_cols[0]], row[month_cols[1]], row[month_cols[2]]
        r_nov, r_dec, r_jan = int(row[rank_cols[0]]), int(row[rank_cols[1]]), int(row[rank_cols[2]])
        # 3ヶ月で単調増か
        trend_up = nov <= dec <= jan
        # 順位が改善したか（1月が11月より良いか）
        rank_improved = r_jan < r_nov
        rank_same = r_jan == r_nov
        rank_worse = r_jan > r_nov
        chat_path = CHAT_DIR / TEAM_FILE_MAP.get(team, "")
        chat = count_chat_blocks_and_keywords(chat_path) if chat_path else {}
        rows.append({
            "チーム名": team,
            "11月投稿数": int(nov),
            "12月投稿数": int(dec),
            "1月投稿数": int(jan),
            "投稿数トレンド": "増加" if trend_up else "減少あり",
            "11月順位": r_nov,
            "1月順位": r_jan,
            "順位変化": "改善" if rank_improved else ("維持" if rank_same else "悪化"),
            "チャット投稿ブロック数": chat.get("blocks", 0),
            "注意・依頼系キーワード出現数": chat.get("supervision", 0),
            "FB・振り返り系キーワード出現数": chat.get("feedback", 0),
        })
        print(f"{team}: 11月{int(nov)} → 12月{int(dec)} → 1月{int(jan)} | トレンド: {'増加' if trend_up else '減少あり'} | 順位: {r_nov}→{r_jan} ({'改善' if rank_improved else '維持' if rank_same else '悪化'})")
        print(f"  チャット: ブロック数={chat.get('blocks', 0)}, 注意・依頼系={chat.get('supervision', 0)}, FB・振り返り系={chat.get('feedback', 0)}")

    summary_df = pd.DataFrame(rows)
    out_path = BASE / "チーム別_3ヶ月投稿とチャット集計結果.xlsx"
    summary_df.to_excel(out_path, index=False)
    print(f"\n出力: {out_path}")
    return summary_df


if __name__ == "__main__":
    main()
