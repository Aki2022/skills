# 段 3 伝えることを選び、優先順を付け、削る

目的：対象者の `priorities` から、この成果物で伝えるメッセージを選び、順番を決め、**言わないもの**を決める。ここが情報設計の質を一番左右する。

## やること

1. （v5）`infoarch-author compose --audience <aud> --strategy <選んだ原子>` で、選んだ戦略に立つメッセージを、対象者の優先順に並べた候補として受け取る。候補の外のメッセージを使うと IA_COMPOSITION_OFF_STRATEGY の警告になる。足りなければ、メッセージの追加を提案する。
2. 対象者の `priorities` の**先頭（1 位）を必ず入れる**。1 位を入れない成果物は、その対象者向けではない。比べる案などで入れないときは、理由を `strategy.deviation` に書く（書かないと IA_COMPOSITION_DEVIATION_UNEXPLAINED の警告）。
3. 媒体の器（段 4 と `media/` のファイル）に入る数だけ、上から選ぶ。器に入らないものは、無理に詰めずに削る。
4. 削ったメッセージを `brief.cut` に書く。削った理由が依頼にある場合は constraints に添える。
5. 選んだメッセージごとに、支える証拠（`proofs` で `supports` が一致するもの）を 1〜2 個選ぶ（個数は出典なし・目安）。証拠の無い効用は、読み手に信じてもらえない。
6. 結論（1 位のメッセージ）を先に置く。流し読みする読み手は、最初の数行で離れるか決める。

## よくある失敗

- 全部入れる（削らない）。どれも弱くなり、1 位が埋もれる。
- 優先リストに無いメッセージを足す。足すべきなら、SSOT の対象者の `priorities` の変更として別の PR で提案する（検査では IA_OUTLINE_OFF_PRIORITY の警告になる）。
- 証拠を置かずに数字や実績を言う。数字は固定値（`{{fct-NNNN}}`）で参照する。

## アウトラインへの書き方

各セクションの `message` と `proofs`、`brief.cut`。

## 出典

- 結論を先頭に置く（逆ピラミッド）、利用者の多くは流し読みする：Nielsen Norman Group「How Users Read on the Web」（https://www.nngroup.com/articles/how-users-read-on-the-web/）。
- 削る判断（「kill your darlings」）：出典なし・目安（原典未確認の格言）。
