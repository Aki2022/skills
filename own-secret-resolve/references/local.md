# local: 手元の Mac で age 鍵と Bitwarden を扱う

## 前提の事実

- age 鍵は `~/.config/sops/age/keys.txt`（mode 0600）にある。この 1 本で全 repo の秘密が開く。公開鍵は reinstall repo の `RECOVERY.md` 4 節に載っている
- 鍵の写しは 2 つある。1 つは Bitwarden の item `reinstall/host/environment` の添付 `keys.txt`。もう 1 つは passphrase で暗号化した `keys.txt.age` で、Time Machine と rclone の `reinstall-gdrive:nix_backups-rclone/reinstall/recovery/` に置いてある
- mise は既定の場所から鍵を探さない。`MISE_SOPS_AGE_KEY_FILE` が要る。nix-darwin の home の `shell.nix` が設定するので、Nix の home 環境が入っていないシェル（darwin-rebuild の前、一部の agent のシェル）では未設定になる
- 使うのは `bw` CLI だけ（rbw も bws も使わない）。reinstall のスクリプトは、`BW_CLI`（絶対パス）、`/opt/homebrew/bin/bw`、PATH の順に `bw` を探す

## 症状 → 原因 → 次の一手

| 症状 | 原因 | 次の一手 |
| --- | --- | --- |
| `test -f ~/.config/sops/age/keys.txt` が偽 | 新しい Mac で、鍵をまだ戻していない | 下の「鍵を戻す 1 行」を渡す。Bitwarden が使えないときは `RECOVERY.md` 4.2（passphrase の写し） |
| 鍵はあるのに `failed to decrypt` | `MISE_SOPS_AGE_KEY_FILE` が無い、または鍵が別物 | `echo "${MISE_SOPS_AGE_KEY_FILE:-unset}"` で変数を見る（パスなので表示してよい）。`age-keygen -y` の結果が公開鍵と合うかも見る |
| `bw unlock` が "You are not logged in" | ログインしていない | 1 行に `bw login` を含める |
| unlock は rc=0 なのに vault が空、または復号されない | 別の版の bw が書いた `BITWARDENCLI_APPDATA_DIR` を共有している | 1 行に `export BITWARDENCLI_APPDATA_DIR="$(mktemp -d)"` を含め、ログインからやり直す |
| "declared bw binary is not installed" | Homebrew の bw が無い | 1 行に `export BW_CLI="$(command -v bw)"` を含める |
| bw が固まる、またはすぐ失敗する | オフライン（bw はネットワークが要る） | 通信を確かめる。オフラインで回避する手は無い |
| スクリプトが「BW_SESSION が無い」で止まる | AI のシェルで実行した | AI は実行しない。`--apply` まで含めた 1 行を管理者に渡す |

## 1 行の型

鍵だけを戻す（新しい Mac。鍵は画面に出さず、ファイルに直接書く。既にあれば上書きしない）:

```bash
test -e ~/.config/sops/age/keys.txt && echo "既に鍵があります" || (export BITWARDENCLI_APPDATA_DIR="$(mktemp -d)"; bw login && export BW_SESSION="$(bw unlock --raw)" && mkdir -p -m 700 ~/.config/sops/age && (umask 077; bw get attachment keys.txt --itemid "$(bw get item reinstall/host/environment | jq -r .id)" --output ~/.config/sops/age/keys.txt); bw lock; unset BW_SESSION BITWARDENCLI_APPDATA_DIR)
```

Mac 全体の秘密（SSH 鍵、各 `.env`、`~/.config` 一式）を戻すときは、手順が長く、作業用ディレクトリの後片付けまで要る。1 行には収まらないので、reinstall repo の `RECOVERY.md` 3 節の手順をそのまま案内する（`restore_secrets.py --output-root "$HOME/.recovery-staging"` に書き出してから移す）。

鍵を置いたら、そのすぐ後に（ファイルを置き直すたびに）Time Machine から外す。これは AI が実行してよい:

```bash
chmod 600 ~/.config/sops/age/keys.txt && tmutil addexclusion ~/.config/sops/age/keys.txt && tmutil isexcluded ~/.config/sops/age/keys.txt
```

vault への書き込みは、`register_common_secret.py`、`set_vault_field.py`、`sync_*_to_bitwarden.py`、`propagate_vault_to_replicas.py --apply` で行う。どれも `BW_SESSION` が要り、`--apply` を付けなければ dry-run になる。dry-run にも `BW_SESSION` が要るので、まず dry-run の 1 行を頼み、その出力を読んでから、`--apply` を付けた 1 行を渡す。正確な引数は各スクリプトの `--help` で確かめる（help は値を出さない）。

## 確かめ方（値を見ない）

- 鍵: `age-keygen -y ~/.config/sops/age/keys.txt` の結果が、`RECOVERY.md` の公開鍵と一致する
- mise から使えるか: repo に入って `mise run check-secrets` が緑になる
- 後片付け: 管理者の Terminal で `bw lock`。1 行に `unset` まで含めてあれば、それで済んでいる
