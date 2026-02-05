# -*- coding: utf-8 -*-
"""
卒業生投稿数レンジ別の万垢達成率テーブルを図で可視化する。
"""
import matplotlib.pyplot as plt
import matplotlib as mpl

mpl.rcParams["font.family"] = ["Hiragino Sans", "sans-serif"]

# 投稿数レンジ（左から多い順）
ranges = [
    "100\n投稿以上",
    "90\n投稿以上",
    "80\n投稿以上",
    "70\n投稿以上",
    "60\n投稿以上",
    "50\n投稿以上",
    "40\n投稿以上",
    "30\n投稿以上",
    "20\n投稿以上",
    "0〜19\n投稿",
]
x = range(len(ranges))

# 卒業生数・万垢数・達成率（%）
sotsugyosei = [14, 2, 5, 5, 6, 3, 7, 9, 8, 25]
manka = [5, 0, 2, 1, 2, 1, 1, 3, 0, 2]
tassei_rate = [35.7, 0.0, 40.0, 20.0, 33.3, 33.3, 14.3, 33.3, 0.0, 8.0]
zentai_wari = [29.4, 0.0, 11.8, 5.9, 11.8, 5.9, 5.9, 17.6, 0.0, 11.8]

fig, ax1 = plt.subplots(figsize=(12, 6))

w = 0.35
bars1 = ax1.bar([i - w / 2 for i in x], sotsugyosei, width=w, label="卒業生（人数）", color="#1565c0", alpha=0.9)
bars2 = ax1.bar([i + w / 2 for i in x], manka, width=w, label="万垢（人数）", color="#2e7d32", alpha=0.9)

ax1.set_ylabel("人数", fontsize=11)
ax1.set_ylim(0, max(sotsugyosei) * 1.15)
ax1.set_xticks(x)
ax1.set_xticklabels(ranges, fontsize=9)
ax1.set_xlabel("投稿数レンジ", fontsize=11)

# 達成率を右軸で折れ線
ax2 = ax1.twinx()
line = ax2.plot(
    x, tassei_rate, color="#c62828", marker="o", linewidth=2, markersize=7, label="達成率（%）"
)
ax2.set_ylabel("達成率（%）", color="#c62828", fontsize=11)
ax2.tick_params(axis="y", labelcolor="#c62828")
ax2.set_ylim(0, 50)
ax2.axhline(y=0, color="#c62828", linestyle="--", alpha=0.4)

# 凡例をまとめる
lns = list(bars1) + list(bars2) + line
lns = [a for a in lns if not a.get_label().startswith("_")]
labs = [l.get_label() for l in lns]
ax1.legend(lns, labs, loc="upper right", fontsize=9)

plt.title("投稿数レンジ別：卒業生数・万垢数・万垢達成率", fontsize=13)
plt.tight_layout()

out_path = __file__.replace(".py", ".png")
plt.savefig(out_path, dpi=150, bbox_inches="tight")
print(f"Saved: {out_path}")
plt.close()
