#!/usr/bin/env python3
"""
build_chart_svg.py — 実データチャートを SVG で構築し PNG 化する参照実装。

⚠️ これは最終fallback（チャートが画像になり Excel の「データの編集」ができなくなる）。
   実データチャートの主経路は「addChart（ネイティブOOXML）＋ネイティブ・オーバーレイ」
   （chart-rules.md §2.5）。この経路を使ってよいのは、ユーザーが編集可能性を明示的に
   放棄した場合のみ（2026-07-10 に全面画像で納品して差し戻された実失敗あり）。

なぜこれが要るか（2026-07-10 実証）:
  実データのチャートは imagegen で描かせると数値が捏造される（ユーザーに「数字が全く違う」と
  reject された実績あり）。chart-rules.md の原則どおり「実データ＝ネイティブ描画」を、追加の重い
  依存なし（soffice は本スキルの既存依存）で成立させるのがこの経路。
  matplotlib が無い環境でも動く。pptxgenjs の addChart では表現できないカスタム
  （軸クリップ・系列末尾に国名/国旗ラベル・凡例の廃止・特定系列だけ強調）を、座標を自分で
  制御できる SVG なら確実に描ける。

使い方:
  1. この雛形をデッキの process/ にコピーし、DATA と描画を対象チャート用に改造する
     （数値は必ず outline.md（唯一の真実源）から直書きする。画像やmockupから読み取らない）。
  2. `python3 build_chart_svg.py` → chart.svg を書き出す
  3. `soffice --headless --convert-to png --outdir . chart.svg` → chart.png
  4. ③のネイティブビルドで、フルスライド画像としてではなく「本文エリアの図」として addImage 埋め込み、
     タイトル/eyebrow/footer はネイティブテキストで別途置く（テキストの編集可能性を確保）。
     ※どうしても軸・ラベルが複雑で全体を1枚にしたい場合のみ全面埋め込みも可（編集元スクリプトを版管理）。

設計メモ:
  - フォントは "Noto Sans CJK JP"（soffice が解決。日本語ラベルが豆腐化しない）。
  - 色は tokens.json 準拠: 強調1系列を primary(#4F4F70) か negative(#C00000)、他は同一グレー(tint)。
  - 外れ値で軸が潰れる場合は y 上限をクリップし、クリップした系列は注記で補う（対数軸も可）。
  - 凡例を廃し、系列末尾（最右点の右横）に「ラベル＋任意アイコン」を置くと視認性が上がる
    （近接する末尾ラベルは y をずらし、リーダー線で結ぶ）。
"""

W, H = 1673, 942  # 16:9。スライド本文図として使うなら任意サイズでよい
PRIMARY, NEG = "#4F4F70", "#C00000"
GRAY_LINE, GRAY_TXT, HAIR = "#A7A7B8", "#59595C", "#DDDDE3"
FONT = "Noto Sans CJK JP"

# ---- データ（outline.md から直書き。ここを対象チャート用に置換する）----
YEARS = [1980, 1990, 2000, 2010, 2020, 2024]
DATA = {  # name: (values, is_emphasis)
    "系列A": ([10, 20, 30, 40, 50, 60], True),
    "系列B": ([60, 50, 40, 30, 20, 10], False),
}
YMAX = 70.0  # y 上限（外れ値クリップ用）

PX0, PX1, PY0, PYB = 200, 1250, 240, 780  # 描画領域（右にラベル余白）

def x(t):  # 年→x
    return PX0 + (t - YEARS[0]) / (YEARS[-1] - YEARS[0]) * (PX1 - PX0)

def y(v):  # 値→y（クリップ）
    v = min(v, YMAX)
    return PYB - (v / YMAX) * (PYB - PY0)

def main():
    s = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
         f'viewBox="0 0 {W} {H}" font-family="{FONT}">',
         f'<rect width="{W}" height="{H}" fill="#FFFFFF"/>']
    # gridlines + y labels
    for gv in range(0, int(YMAX) + 1, max(1, int(YMAX // 5))):
        gy = y(gv)
        s.append(f'<line x1="{PX0}" y1="{gy:.1f}" x2="{PX1}" y2="{gy:.1f}" stroke="{HAIR}" stroke-width="1"/>')
        s.append(f'<text x="{PX0-14}" y="{gy+7:.1f}" font-size="20" fill="{GRAY_TXT}" text-anchor="end">{gv}</text>')
    # x axis + labels
    s.append(f'<line x1="{PX0}" y1="{PYB}" x2="{PX1}" y2="{PYB}" stroke="{GRAY_TXT}" stroke-width="1.5"/>')
    for t in YEARS:
        s.append(f'<text x="{x(t):.1f}" y="{PYB+34:.1f}" font-size="22" fill="{GRAY_TXT}" text-anchor="middle">{t}</text>')
    # series (emphasis last / on top)
    for name, (vals, emph) in sorted(DATA.items(), key=lambda kv: kv[1][1]):
        col = (NEG if emph else GRAY_LINE)
        pts = " ".join(f"{x(t):.1f},{y(v):.1f}" for t, v in zip(YEARS, vals))
        s.append(f'<polyline points="{pts}" fill="none" stroke="{col}" stroke-width="{5 if emph else 3}" stroke-linejoin="round"/>')
        # end-of-line label (右端)
        s.append(f'<text x="{PX1+12}" y="{y(vals[-1])+6:.1f}" font-size="22" '
                 f'fill="{col}" font-weight="{"700" if emph else "400"}">{name} {vals[-1]}</text>')
    s.append('</svg>')
    with open("chart.svg", "w") as f:
        f.write("\n".join(s))
    print("wrote chart.svg  →  soffice --headless --convert-to png --outdir . chart.svg")

if __name__ == "__main__":
    main()
