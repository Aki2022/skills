"""refresh_adc.sh の走査対象と書き込みを固定する。

2026-11-09 に分かった不具合: 走査が `CODE_DIR/*/.mise.toml`（ドット付き）だけで、
`mise.toml`（ドットなし）の repo は対象に挙がらず、エラーも出ずに古い per-repo ADC が残った。
これは「成功に見えて誤る」型なので、次を固定する:
  1. `.mise.toml` と `mise.toml` の両方を走査する
  2. `--list` は対象を一覧するだけで、ADC も設定ファイルも書かない（gcloud も source ADC も要らない）
  3. 1 つの repo の 2 つのファイルが別々の SA を指すときは、何も書かずに止まる
  4. 書き込み経路: ドットなしのファイルにも GOOGLE_APPLICATION_CREDENTIALS が [env] に入り、
     値はホームを `{{env.HOME}}` で書く（repo に個人のパスを残さない）
  5. どちらかのファイルに GOOGLE_APPLICATION_CREDENTIALS があれば追記しない
"""
import json
import os
import stat
import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[2] / "refresh_adc.sh"

SA_A = "a-sa@proj-a.iam.gserviceaccount.com"
SA_B = "b-sa@proj-b.iam.gserviceaccount.com"


def _write(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)
    return path


class RefreshAdcTest(unittest.TestCase):
    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.home = self.root / "home"
        self.code = self.root / "code"
        self.gcloud_dir = self.home / ".config" / "gcloud"  # 実運用と同じくホーム配下
        self.bin = self.root / "bin"
        for d in (self.home, self.code, self.gcloud_dir, self.bin):
            d.mkdir(parents=True)
        # 偽の gcloud: print-access-token は成功し、それ以外は何もしない
        fake = _write(self.bin / "gcloud", "#!/bin/sh\nexit 0\n")
        fake.chmod(fake.stat().st_mode | stat.S_IEXEC)

    def tearDown(self):
        self._tmp.cleanup()

    def _run(self, *args, source_adc=True):
        if source_adc:
            _write(self.gcloud_dir / "application_default_credentials.json", json.dumps({
                "type": "authorized_user", "client_id": "x", "client_secret": "y",
                "refresh_token": "z",
            }))
        env = dict(os.environ, HOME=str(self.home), CODE_DIR=str(self.code),
                   CLOUDSDK_CONFIG=str(self.gcloud_dir),
                   PATH=f"{self.bin}:/usr/bin:/bin")
        env.pop("SOURCE_ADC_PATH", None)
        return subprocess.run(["bash", str(SCRIPT), *args], env=env,
                              capture_output=True, text=True, timeout=60)

    def _snapshot(self):
        return {str(p): p.read_bytes() for p in self.root.rglob("*") if p.is_file()}

    # 1 + 2
    def test_list_includes_dotless_and_dotted_and_writes_nothing(self):
        _write(self.code / "dotted" / ".mise.toml",
               f'[env]\nCLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "{SA_A}"\n')
        _write(self.code / "dotless" / "mise.toml",
               f'[tools]\npython = "3.12"\n\n[env]\nCLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "{SA_B}"\n')
        _write(self.code / "no_gcp" / "mise.toml", '[env]\nFOO = "1"\n')
        before = self._snapshot()

        r = self._run("--list", source_adc=False)

        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        rows = [ln.split("\t") for ln in r.stdout.splitlines() if "\t" in ln]
        by_repo = {row[0]: row for row in rows}
        self.assertEqual(set(by_repo), {"dotted", "dotless"}, r.stdout)
        self.assertEqual(by_repo["dotted"][1], SA_A)
        self.assertEqual(by_repo["dotless"][1], SA_B)
        self.assertTrue(by_repo["dotless"][2].endswith("/dotless/mise.toml"))
        self.assertEqual(self._snapshot(), before, "--list が何かを書いた")

    # 3
    def test_conflicting_sa_in_both_files_stops_without_writing(self):
        _write(self.code / "both" / ".mise.toml",
               f'[env]\nCLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "{SA_A}"\n')
        _write(self.code / "both" / "mise.toml",
               f'[env]\nCLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "{SA_B}"\n')
        _write(self.code / "ok" / "mise.toml",
               f'[env]\nCLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "{SA_A}"\n')
        self._run("--list", source_adc=True)  # source ADC を置くためだけ
        before = self._snapshot()

        for args in (("--list",), ()):
            r = self._run(*args)
            self.assertNotEqual(r.returncode, 0, f"{args}: {r.stdout}")
            self.assertIn("both", r.stdout + r.stderr)
            self.assertEqual(self._snapshot(), before, f"{args}: 衝突なのに書いた")

    # 4
    def test_full_run_writes_adc_and_inserts_into_dotless_env(self):
        mise = _write(self.code / "dotless" / "mise.toml",
                      '[tools]\npython = "3.12"\n\n[env]\n'
                      f'CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "{SA_B}"\n\n'
                      '[tasks.x]\nrun = "true"\n')

        r = self._run()

        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        adc = json.loads((self.gcloud_dir / "dotless_adc.json").read_text())
        self.assertEqual(adc["type"], "impersonated_service_account")
        self.assertIn(SA_B, adc["service_account_impersonation_url"])
        text = mise.read_text()
        env_block = text.split("[env]", 1)[1].split("[tasks.x]", 1)[0]
        self.assertIn("GOOGLE_APPLICATION_CREDENTIALS", env_block, text)
        self.assertNotIn(str(self.home), text, "個人のホームパスが設定ファイルに入った")
        self.assertIn('"{{env.HOME}}/', text)

    # 5
    def test_existing_gac_in_either_file_is_not_duplicated(self):
        dotted = _write(self.code / "split" / ".mise.toml",
                        f'[env]\nCLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "{SA_A}"\n')
        dotless = _write(self.code / "split" / "mise.toml",
                         '[env]\nGOOGLE_APPLICATION_CREDENTIALS = "{{env.HOME}}/x_adc.json"\n')
        before = (dotted.read_text(), dotless.read_text())

        r = self._run()

        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertEqual((dotted.read_text(), dotless.read_text()), before)
        self.assertTrue((self.gcloud_dir / "split_adc.json").exists())


    # 4 の補足: [env] が無いファイルでも、末尾の別の表（[tasks.*]）へ入れない
    def test_insert_without_env_table_does_not_land_in_last_table(self):
        mise = _write(self.root / "mise.toml", '[tasks.x]\nrun = "true"\n')
        env = dict(os.environ, HOME=str(self.home))
        r = subprocess.run(["python3", str(SCRIPT.parent / "scripts" / "mise_insert_adc.py"),
                            str(mise), str(self.gcloud_dir / "r_adc.json")],
                           env=env, capture_output=True, text=True, timeout=30)
        self.assertEqual(r.returncode, 0, r.stderr)
        text = mise.read_text()
        tail = text.split("[tasks.x]", 1)[1]
        self.assertIn("[env]", tail, text)
        self.assertIn("GOOGLE_APPLICATION_CREDENTIALS", tail.split("[env]", 1)[1], text)
        self.assertNotIn("GOOGLE_APPLICATION_CREDENTIALS", tail.split("[env]", 1)[0], text)


if __name__ == "__main__":
    unittest.main()
