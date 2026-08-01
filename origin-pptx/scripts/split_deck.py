#!/usr/bin/env python3
"""
split_deck.py — 1枚岩の build_deck.js を slides/sNN.js のスライド別モジュールへ分割する。

なぜ分割するか（2026-07-10 実証）:
  ③で各スライドを「そのmockupを読んだビルダーサブエージェント」に並列で実装させるとき、
  1ファイルの build_deck.js だと全員が同一ファイルを奪い合って衝突する。1スライド1ファイルに
  割ると、ビルダーは自分の sNN.js だけを触れば済み、安全に並列実行できる。

前提の構成:
  - build_deck.js は各スライドを `// ============ SNN ... ============` の見出し＋
    `(() => { ... })();` 即時関数で書いていること（build_slide.example / deck_helpers 準拠）。
  - 分割後は薄いエントリ（build_deck2.js 相当）が共有 ctx（deck_helpers の mk()）を作り、
    `for (i=1..N) require('./slides/sNN.js')(ctx)` で束ねる。各 sNN.js は
    `module.exports = (ctx) => { const {T,box,...}=ctx; <元のIIFE本文> }` の形。

使い方:
  cd <deck>/process && python3 split_deck.py   # build_deck.js を読み slides/sNN.js を書き出す
"""
import re, os

SRC = "build_deck.js"
with open(SRC) as f:
    src = f.read()

pat = re.compile(r"// ={5,} S(\d+)[^\n]*\n\(\(\) => \{\n(.*?)\n\}\)\(\);", re.S)
matches = pat.findall(src)
os.makedirs("slides", exist_ok=True)
CTX = ("  const { pptx, ST, C, IN, T, box, rct, arrow, chevron, homePlate, line, imgFit, imgPath,\n"
       "          chrome, footer, source, badge, newSlide, ASSETS, IMG } = ctx;\n")
for num, body in matches:
    n = int(num)
    mod = (f"// slides/s{n:02d}.js — mockup_{n:02d}.png を仕様とするスライド実装\n"
           f"module.exports = (ctx) => {{\n{CTX}{body}\n}};\n")
    with open(f"slides/s{n:02d}.js", "w") as f:
        f.write(mod)
    print(f"slides/s{n:02d}.js")
print(f"total: {len(matches)}  (次: 薄いエントリから require して束ねる)")
