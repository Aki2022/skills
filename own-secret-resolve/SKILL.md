---
name: own-secret-resolve
description: |
  sops + age で暗号化した秘密と Bitwarden（bw CLI）の vault を扱う作業を、止まらずに、値を一度も
  表示せずに進める手順。手元の Mac で age 鍵や vault を扱う「local」と、project の repo で
  secrets.enc.json を読み書きする「開発」の 2 場面を持つ。
  必ず使うこと: sops・age・keys.txt・secrets.enc.json・Bitwarden・bw・BW_SESSION・
  MISE_SOPS_AGE_KEY_FILE・COMMON__ が出てきた時、「failed to decrypt」「環境変数が空」「秘密を足したい」
  「API キーを追加/差し替えたい」「鍵を戻したい」「vault に書きたい」と言われた時、mise の [env] _.file や
  check-secrets が赤い時、worktree や新しい clone で test が秘密不足で落ちる時、秘密の値を
  「確かめて」と頼まれた時。名前が出ていなくても、秘密の供給で手が止まりそうなら使う。
  使わない場面: GCP の ADC・impersonation（own-gcp-provision）、gitleaks などの guardrail 設定
  （own-vibeguard-harden）、.env を repo 一式で sops に移す大きな移行（reinstall repo の
  migrate_env_to_mise_sops.py と secret-delivery-rollout guide）。
---

# own-secret-resolve

秘密の作業で AI が止まる場所は、だいたい 4 つに決まっている。人の手が要る所、復号できない原因、秘密の足し方、値を見せない線引き。この skill は、その 4 つそれぞれに次の一手を用意する。

## まず場面を決める

| 場面 | 見分け方 | 読む手順 |
| --- | --- | --- |
| local | 相手が repo の外にある。`~/.config/sops/age/keys.txt`、Bitwarden の vault、新しい Mac、`bw login`/`unlock` | `references/local.md` |
| 開発 | 相手が repo の中にある。`secrets.enc.json`、`.mise.toml` の `_.file`、`mise run check-secrets`、worktree | `references/dev.md` |

両方にまたがる場合は、local から片づける。開発の問題の多くは、鍵が手元に無いことが根にあるから。

## 共通の決まり

### 1. 値は表示しない。確かめるときも値は見ない

一度表示した値は transcript に残り、取り消せない。だから、値が出るコマンドは使わない。確かめたいことは、値を出さない別の方法で確かめられる。

| やりたいこと | 使うもの | 使わないもの |
| --- | --- | --- |
| 秘密が供給されているか | `mise run check-secrets`。`check-<name>-secrets` があればそれも | `mise env`、`mise set`、`printenv`、`echo $X` |
| ある変数が入っているか | `mise exec -- sh -c 'test -n "$NAME" && echo set \|\| echo unset'` | 値を echo すること |
| どんな鍵名があるか | `jq -r 'keys[] \| select(. != "sops")' secrets.enc.json`（鍵名は平文） | `sops -d` |
| 2 か所の値が同じか | reinstall repo の ssot_check（`ssot_check.py`）（digest で比べる）。単発なら `mise exec -- sh -c 'printf %s "$X" \| shasum'` | 値を並べて比べること |
| 鍵が正しいか | `age-keygen -y ~/.config/sops/age/keys.txt`（公開鍵だけが出る） | `cat keys.txt` |

vibe-guard の hook は、`sops -d`/`decrypt`、`mise env`、引数なしの `mise set`、`bw get` の表示、keys.txt の読み出しを止める。止められたときに抜け道を探すのはやめる。それは「上の表の右の列を使った」という合図なので、左の列の方法に切り替える。

### 2. 人の手が要る所では、1 行を渡して結果を読む

AI のシェルには、マスターパスワードも 2FA も `BW_SESSION` も届かない。`export` してもらっても、AI のプロセスには伝わらない。そこで「何をしてほしいか」を説明して頼むのではなく、管理者が Terminal にそのまま貼れば終わる 1 行を、`bash` のコードブロックで渡す。

- その 1 行で完結させる。unlock から作業、`unset BW_SESSION BW_CLI BITWARDENCLI_APPDATA_DIR` までを含める。`BW_SESSION` は期限が無く、vault 全体を開けてしまうから
- 値を入れる作業は `scripts/sops-set-secret.sh` に任せる。値は画面に出ず、空の値、空白を含む値、prefix が違う値は書かれない
- 渡した後は、結果を値を見ずに確かめる（上の表）。Terminal の出力は `read_terminal` か、管理者が貼ったログで読む
- 人の手が要るのは、`bw login`/`unlock`、vault への書き込み、値の入力、force-push、`.env` の削除。それ以外は AI が自分で進める

例（開発の repo に API キーを足す。値は実行後に貼ってもらい、画面には出ない）:

```bash
cd <repo> && EXPECT_PREFIX=sk-ant- bash ~/.agents/skills/own-secret-resolve/scripts/sops-set-secret.sh ANTHROPIC_API_KEY
```

### 3. 原因は症状から引く

推測で設定を足していかない。各 reference の「症状 → 原因 → 次の一手」の表で原因を 1 つに絞ってから直す。特に `_.file` の指す先が無いと、エラーにならず秘密が黙って空になる。「動いたように見える」は確認として扱わない。`check-secrets` が緑になって初めて確認できたことになる。

## 正本

詳しいことは reinstall repo の文書にある。この skill は、その要点を引けるようにした写しにすぎない。

- `RECOVERY.md`: 3 節が bw、4 節が age 鍵、6 節が症状表
- `docs/specs/secret-delivery.md`: 秘密の届き方
- `docs/specs/agent-secret-handling.md`: この skill と扉の仕様
- `docs/guides/secret-delivery-field-notes.md`、`secret-delivery-migrate-tool.md`（COMMON__）、`secret-delivery-ssot-check.md`
