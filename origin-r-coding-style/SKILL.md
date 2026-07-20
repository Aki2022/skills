---
name: origin-r-coding-style
description: >-
  Write or review R code (.R / .Rmd / .qmd) in the human-authored,
  reproducible tidyverse + targets style used across these climate/GIS
  research repos. Use whenever creating, editing, refactoring, or reviewing
  R functions, {targets} pipelines, or geospatial data code (sf / GeoParquet
  / DuckDB / duckplyr), or when naming functions/variables and choosing which
  libraries and loop constructs to use. Also use when porting legacy
  camelCase R code to the current standard. Do not use for Python-only work
  (use uv-based rules) or non-code docs.
---

# R Coding Style（tidyverse + targets, human-first）

このスキルは、人間（R しか読まない研究者）が保守するコードの一貫性を守るための R スタイル規約。
これまでのプロジェクトの流儀を踏襲しつつ、命名は標準の snake_case に統一する。

対象: `**/*.R`, `**/*.Rmd`, `**/*.qmd`。Python は対象外（`uv` ベースの別規約に従う）。

## 命名とレイアウト（必須）

| 項目       | 規則                                 | 例                                         |
| ---------- | ------------------------------------ | ------------------------------------------ |
| ファイル名 | snake_case                           | `func_make_pmtiles.R`, `params.R`          |
| 関数名     | `func_` + snake_case                 | `func_year_round()`, `func_make_geojson()` |
| 変数名     | snake_case                           | `target_mesh`, `input_dir`, `all_pattern`  |
| 定数       | UPPER_SNAKE_CASE                     | `GCS_DEFAULT_BUCKET`, `TARGET_CRS`         |
| インデント | スペース 2 文字                      | —                                          |
| データ列名 | ドメイン名を尊重（無理に改名しない） | `change_2030`, `MESH3_ID`, `var_eng`       |

- **関数は必ず `func_` プレフィックス**。パイプラインの target 名やヘルパーと区別する既存慣習。
- **旧 camelCase（`func_yearRound`, `targetMesh` 等）は移行対象**。触ったファイルは snake_case へ寄せる（一度に全改名しない、触れた範囲で）。データ列の物理名（`B1980_2000` 等）は外部データ由来なので改名しない。

## 基本原則

- **tidyverse 中心**。データ操作は dplyr / tidyr / purrr / stringr で書く。base の for ループより `purrr::map_*` を優先。
- **パイプは `%>%`（magrittr）**。既存コードが `%>%` で統一されているため踏襲（`|>` に混ぜない）。
- **パス管理は `here::here()`**。マシン固有の絶対パス・`~/` をコードに書かない（config か環境変数へ）。
- **1 関数 1 責務**。小さく分割し、同じ処理を複数箇所に書かない。
- **エラーは `stop()` で明示**。回復不能な前提崩れは早期に落とす。進捗表示は `message()`（旧コードの `print(paste0(...))` は `message()` に寄せる）。
- **roxygen2** で関数をドキュメント化（`@param` / `@return`）。
- **TDD / testthat**。純関数にはテストを先に書く。副作用（IO・GCS・スクレイプ）は薄いラッパに隔離してテスト対象から外す。

## パッケージ管理

ファイル冒頭で `pacman` を使ってロードする（インストール兼用）:

```r
if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")
pacman::p_load(tidyverse, sf, arrow, here)
```

- **依存の固定は renv**（`renv.lock`）。さらに GDAL/GEOS/PROJ 等のシステムライブラリは rix か Docker で固定する（sf/terra の再現性のため）。
- tidyverse・sf 以外の**使用頻度が低い/名前が衝突しうるパッケージは `pkg::fn()` で明示呼び出し**（例: `DBI::dbConnect`, `dbplyr::sql`, `duckdb::duckdb`, `googleCloudStorageR::gcs_upload`, `sf::st_as_sf`）。

## ライブラリ選定（この系統の標準）

| 用途                  | 使うもの                                                                                                                                          |
| --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| パイプライン / 再現性 | `targets`, `tarchetypes`, 並列は `crew`（大量分岐は `tar_map` / dynamic branching）                                                               |
| 並列 map              | `furrr`（`future_map_dfr(..., .options = furrr_options(seed = TRUE))`）。crew と役割が重なる場合は crew を優先                                    |
| ベクタ地理データ      | `sf`、保存は GeoParquet（`arrow` / `sfarrow`、可能なら 1.1 の bbox covering 列つき）                                                              |
| 大規模結合・集計      | **DuckDB spatial**。R からは **`duckplyr` を必須**（dplyr のまま out-of-core。生 SQL は spatial join 等どうしても必要な箇所のみ `dbplyr::sql()`） |
| タイル生成            | tippecanoe（`system2()` でシェル呼び出し）。バージョンは環境固定 + 関数内で `--version` 照合                                                      |
| クラウド              | `googleCloudStorageR`（GCS）, `googledrive`（元データ取込）。認証は鍵ファイルでなく ADC                                                           |
| 表・レポート          | Quarto（`tar_quarto` で DAG の葉に）。地図は `mapgl`（純 R で MapLibre + PMTiles）                                                                |

## ループ / 反復の書き方

- パラメータ全組み合わせは `expand.grid()` → `tibble` 化して 1 行 1 ケースで持つ（既存 `func_make_all_params` 慣習）。
- 反復は `purrr::map_*` / `furrr::future_map_dfr`。**素の `for` は副作用のみの手続き（ディレクトリ作成等）に限定**。
- 乱数を含む並列は `furrr_options(seed = TRUE)`、targets 側は `tar_option_set(seed = ...)` を明示。

## ファイル構成（targets プロジェクト）

- `_targets.R`（DAG 定義）／ `params.R`（設定・定数）／ `functions.R` or `func_*.R`（関数）を分ける。
- セクション区切りコメントで責務を可視化（既存慣習、日本語コメント可）:

```r
# # ##########################################
# # 気候メッシュデータを読み、変化率を付与する
# # ##########################################
func_year_round <- function(df) {
  ...
}
```

## レビュー時チェックリスト

- [ ] 関数は `func_` + snake_case、変数は snake_case、インデント 2 スペース
- [ ] `%>%` で統一（`|>` 混在なし）、パスは `here::here()` / config（`~/`・絶対パスなし）
- [ ] 重い結合は duckplyr / DuckDB で out-of-core（全量メモリ読みしていない）
- [ ] 純関数に testthat、副作用はラッパに隔離
- [ ] roxygen2 コメントあり、エラーは `stop()`、進捗は `message()`
- [ ] 秘密情報（鍵・トークン）をコード/ログに書いていない（ADC 利用）
- [ ] 触ったレガシー camelCase を snake_case へ寄せた（範囲内で）
