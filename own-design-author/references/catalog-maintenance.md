# カタログの生成・分類・移行

取得は人間が起動し、手元に置かれた版付きDADS Markdown ZIPを読む。ファイル名だけで版を認めず、生成物のarchive_sha256も確認する。対象はfoundations/components各直下のindex.md、h2/h3。guidanceは参照元であり項目にはしない。

```sh
python3 <skill>/scripts/catalog.py generate --zip <local.zip> --version <new-version> --output <new-catalog.json> --ledger <id-ledger.json> --previous <old-catalog.json>
```

新規発行id、消えた照合キー、明示した改名が差分に出る。台帳のidを振り直さない。改名は `--renames <json>` の旧キー→新キー対応で明示し、推測で確定しない。旧キー履歴を保持する。項目の入力が同一なら既存の分類を持ち越し、変わった入力だけ再分類する。

初回カタログは2026-10-14版、8＋49文書、560見出し＋補遺6件。出典と加工の明示はカタログ自身が持つ。id-ledgerと版付きカタログを同じ変更に含める。

## must分類

`classification-policy.json` の規則・閾値・版固定モデルを使う。分類はカタログ更新時だけのAPI処理で、L0/L1/L2から呼ばない。公開された上流の本文・見出し・出典だけを送り、repoの実トークンや承認情報を送らない。

```sh
python3 <skill>/scripts/classify.py --catalog <catalog.json> --usage-ledger <usage.json> --limit 50 --dry-run
python3 <skill>/scripts/classify.py --catalog <catalog.json> --usage-ledger <usage.json> --limit 50
```

dry-runは件数・リクエスト規模の見積もりであり、APIや書込み経路の検証ではない。最初の50件で実際のinput/output token使用量と使用量由来USDを確認し、その結果を記録してから残りに広げる。通貨上限を強制する機能ではない。請求書の金額を検証したとも扱わない。

既存の認証はvibe-guardと同順で読む。環境変数 `TYPESAFE_API_KEY`、未設定なら `TYPESAFE_API_KEY_FILE`、それも未指定なら既存の `~/.config/typesafe/api-key`。値を引数・ログ・生成物に出さず、認証を新設しない。送信先はTypeSafe公式APIに固定し、redirectでキーを転送しない。

1項目につき、全体mustと3プリセットでの追加mustの組合せをChoice一問で決める。16通りの回答を固定し、confidenceが閾値未満なら保守側のmustと全presetにする。入力ハッシュ、モデル版、policyハッシュ、閾値、生のChoice確率を保存する。後段のL0は、その固定結果と現在の入力の一致を検証する。

台帳はネットワーク送信前に判定件数を予約する。失敗も消費に含め、2,000判定を超えて実行しない。自動再試行はない。成功batchはその都度保存し、再開時は入力・policyが変わっていない既判定を再送しない。HTTP応答を保存できなかった分の使用量は不明として扱い、0と推定しない。同じ原因の失敗が3回続けばWSへ証拠を残して停止する。

分類済みのカタログを既存repoへ自動反映しない。`design-author migrate` によってだけpinを変える。状態差分と旧承認台帳を保持し、新項目・新mustの扱いをL0で確認する。
