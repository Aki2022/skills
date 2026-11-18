# 段 3 言うことを選び、一番を決め、削る

目的：立っている positioning の提供価値から、この成果物で言うメッセージを選び、一番のメッセージを 1 つ決め、**言わないもの**を決める。ここが情報設計の質を一番左右する。

## やること

1. `infoarch-author positioning` の隙間（違いになる点）と同等であるべき点を見る。違いの提供価値（`kind: difference`）は隙間を、同等の提供価値（`kind: parity`）は同等であるべき点を満たす。
2. メッセージは提供価値をちょうど 1 つ指す（`value`）。言葉ではなく中身で、文案は段 5 で書く。比べる代替があれば `compares` に書く（立っている positioning の承認済みの代替だけ）。
3. **一番のメッセージ（`rank: lead`）をちょうど 1 つ**決める。結論を先に置く。流し読みする読み手は、最初の数行で離れるか決める。
4. 行動喚起は `type: cta` のメッセージで、その行動で何が得られるかとして提供価値を指し、`achieves` で取ってほしい行動を指す。置かないなら `no_cta: true` と宣言する。
5. 媒体の器に入らないものは、メッセージにしない（outline に書かない）。書いたのに置かないメッセージは警告になる（IA_OUTLINE_MESSAGE_UNPLACED）。
6. 一番のメッセージの提供価値の経路に、証拠が支える原子・矢印があるか確かめる（無ければ IA_LEAD_UNSUPPORTED の警告。メッセージは証拠を持たず、経路から受け継ぐ）。

## よくある失敗

- 全部入れる（削らない）。どれも弱くなり、一番が埋もれる。
- 立っていない positioning（別のセグメント）の提供価値を指す（IA_OUTLINE_VALUE_OTHER_POSITIONING）。
- 証拠の無い効用を一番にする。数字は固定値（`{{fct-NNNN}}`）で参照する。

## アウトラインへの書き方

`messages`（`value`・`rank`・`type`・`compares`・`achieves`）、`no_cta`。

## 出典

- 結論を先頭に置く（逆ピラミッド）、利用者の多くは流し読みする：Nielsen Norman Group「How Users Read on the Web」（https://www.nngroup.com/articles/how-users-read-on-the-web/）。
- 削る判断（「kill your darlings」）：出典なし・目安（原典未確認の格言）。
