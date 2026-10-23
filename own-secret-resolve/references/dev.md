# 開発: project の repo で secrets.enc.json を扱う

## 前提の事実

- 秘密は repo に commit された `secrets.enc.json`（sops + age）にあり、`.mise.toml` の `[env] _.file` から読み込まれる。名前は `.json` か `.yaml` にする。`*.env` は mise が復号しない
- `.sops.yaml` は無い。暗号化するときは公開鍵をコマンドで渡す（`sops --encrypt --age <pub> --input-type json --output-type json /dev/stdin`）。足すとき・変えるときは `sops set` を使う（つまり `scripts/sops-set-secret.sh`）
- 公開してよい設定値は、`.mise.toml` に平文で書く。アカウントに紐づく設定は `~/.config/mise/env/<name>.json` に置く
- 複数の repo で共有する秘密は、`COMMON__<NAME>` という鍵名にする。正本は Bitwarden の `reinstall/common/<name>`。`.mise.toml` の `_.file` より後ろで `NAME = "{{ env.COMMON__NAME }}"` と書いて写す
- 供給の検査は repo ごとの `mise run check-secrets` で行う（入れ子の config には `check-<name>-secrets` もある）

## 症状 → 原因 → 次の一手

| 症状 | 原因 | 次の一手 |
| --- | --- | --- |
| `failed to decrypt sops file` | 鍵が無い、または `MISE_SOPS_AGE_KEY_FILE` が無い | `references/local.md` の表へ。worktree かどうかは関係ない |
| `mise ... is not trusted` / config が読まれない | worktree や新しい clone で、まだ trust していない | `mise trust .mise.toml`（AI が実行してよい）。worktree ごとに要る |
| test が「環境変数が無い」で落ちるのに、mise はエラーを出さない | `_.file` の指す先が無く、黙って空になっている | `ls` で `_.file` の path を確かめる。`git status` で `secrets.enc.json` が checkout にあるかも見る。その後 `mise run check-secrets` |
| `check-secrets` が特定の鍵だけ赤い | その鍵が `secrets.enc.json` に無い | `jq -r 'keys[]' secrets.enc.json` で鍵名を見る。無ければ下の「秘密を足す」 |
| mise の変数展開でエラー（`env.COMMON__…` が無い） | 写す行が `_.file` より前にある、または `COMMON__` の鍵が無い | 写す行を `_.file` より後ろへ動かす。鍵名を確かめる |
| `check-secrets` task が無い | まだ供給検査を入れていない | mockup repo の `scripts/check_sops_supply.py` を写して task に載せる（reinstall の `secret-delivery-rollout` guide） |
| 隔離した `MISE_DATA_DIR` で動かない | 相対パスで指定した | 絶対パス（scratchpad の下）にする |

## 秘密を足す・変える

1. どの名前で足すかを決める。他の repo でも同じ値を使うなら `COMMON__<NAME>` にし、vault の `reinstall/common/<name>` が正本になる
2. 管理者に 1 行を渡す。値は実行後に貼ってもらうか、vault から管で渡す:

```bash
cd <repo> && EXPECT_PREFIX=<prefixがあれば> bash ~/.agents/skills/own-secret-resolve/scripts/sops-set-secret.sh <KEY_NAME>
```

```bash
cd <repo> && export BW_SESSION="$(bw unlock --raw)" && bw get password '<vault item>' | bash ~/.agents/skills/own-secret-resolve/scripts/sops-set-secret.sh <KEY_NAME>; unset BW_SESSION
```

3. AI が値を見ずに確かめる。`jq -r 'keys[]' secrets.enc.json` に鍵名があること、`mise run check-secrets` が緑になること
4. `secrets.enc.json` の差分を commit する。暗号文なので commit してよい。PR 本文に値やパスは書かない
5. `COMMON__` の鍵を変えたときは、他の repo に写した分も揃える（reinstall の `propagate_vault_to_replicas.py --common`。`BW_SESSION` が要るので 1 行で頼む）。食い違いは `ssot_check.py` で検出する

`EXPECT_PREFIX` の例: Anthropic は `sk-ant-`、OpenAI は `sk-`、GitHub PAT は `github_pat_`、Slack bot は `xoxb-`。分からないときは付けない。

## worktree・新しい clone での手順

```bash
mise trust .mise.toml && mise run check-secrets
```

これが緑なら test に進む。赤なら上の表で原因を引く。
