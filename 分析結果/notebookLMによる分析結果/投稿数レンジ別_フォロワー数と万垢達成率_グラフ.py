# -*- coding: utf-8 -*-
"""
投稿数レンジ別のフォロワー数・万垢達成率を可視化する。
投稿数レンジ別_フォロワー数と万垢達成率の傾向.md の表に基づく代表値を使用。
"""
import matplotlib.pyplot as plt

# レンジとラベル
ranges = ["0〜20\n投稿", "21〜40", "41〜60", "61〜80", "81〜100", "100\n投稿以上"]
x = range(len(ranges))

# 平均フォロワー数（未達成者目安）の代表値（レンジの中間値など）
follower_mid = [25, 450, 1750, 3000, 5000, 6000]  # 100+ は 3000〜 の代表で6000

# 万垢達成者の割合（%）
manaka_rate = [5, 15, 20, 25, 35, 50]

fig, ax1 = plt.subplots(figsize=(10, 5.5))

color1 = "#2e7d32"
color2 = "#1565c0"
bars = ax1.bar([i - 0.2 for i in x], follower_mid, width=0.4, label="平均フォロワー数（未達成者目安）", color=color1, alpha=0.85)
ax1.set_ylabel("平均フォロワー数（目安）", color=color1, fontsize=11)
ax1.tick_params(axis="y", labelcolor=color1)
ax1.set_ylim(0, 7000)

ax2 = ax1.twinx()
line = ax2.plot(x, manaka_rate, color=color2, marker="o", linewidth=2, markersize=8, label="万垢達成者の割合（目安）")
ax2.set_ylabel("万垢達成者の割合（%）", color=color2, fontsize=11)
ax2.tick_params(axis="y", labelcolor=color2)
ax2.set_ylim(0, 55)

ax1.set_xticks(x)
ax1.set_xticklabels(ranges, fontsize=10)
ax1.set_xlabel("投稿数レンジ", fontsize=11)
ax1.set_title("投稿数レンジ別：フォロワー数と万垢達成率の傾向", fontsize=13)

# 凡例をまとめる
lns = list(bars) + line
labs = [a.get_label() for a in lns if not a.get_label().startswith("_")]
lns = [a for a in lns if not a.get_label().startswith("_")]
if lns:
    ax1.legend(lns, labs, loc="upper left", fontsize=9)

plt.tight_layout()
out_path = __file__.replace(".py", ".png")
plt.savefig(out_path, dpi=150, bbox_inches="tight")
print(f"Saved: {out_path}")
plt.close()
