# 変更・期限・版更新の維持検査

`design-author maintain --repo .` をセッション終了のStep 2.5から呼ぶ。
`design/`のないrepoは1行のJSONでskipする。モデル・ネットワーク・承認記録を使わない。
`--force` は変更の有無にかかわらずL0を再実行する。

成功した検査のcommitと実装/policy/カタログの指紋を
`design/.resolved/maintenance-stamp.json`へ保存する。これは生成物であり、正典や承認ではない。
初回・不正stamp・commit差分・未commit変更・期限・指紋の変更でL0を実行し、
マニフェストを再生成する。差分はgitの`--diff-filter=ACDMRT`と`status --porcelain`で取得する。
作業ツリーに変更が残っていれば、次の実行でも検査する。
失格時は成功stampを進めず、古いマニフェストはL0の通常契約に従って消す。

期限が必要な項目だけdesign.yamlへ設定する。未設定なら期限によるtriggerはない。

```json
{"maintenance":{"within_days":7,"review_due_dates":{"dads-000001":"2026-10-30"}}}
```

`within_days`の既定は7、0以上の整数。期限はカタログまたはretired_itemsに存在するidと
ISO日付だけを受け付ける。今日から指定日数以内と期限超過を列挙し、差分0でもL0を走らせる。
期限を過ぎただけでは承認を削除せず、人間へ再点検の候補として示す。
日付や上限を検査の都合で自動延長しない。

出力はtriggers、changed_paths、due_items、invalidated_approvals、approval_candidates、
judgment_candidates（hashのみ）、L0のcounts/errors/warnings/5区分を持つ。
候補の本文や端末固有のパスを記録しない。実承認台帳は一切更新しない。
終了コードはL0の0/1/2。0は整合、1は警告、2は欠落・不一致等の失格。
未commitのrepo、読めない履歴、symlink出力先、同時実行、検査中の入力変更も失格にする。
維持検査はL0だけを担当し、実アプリのL1/L2や出荷ゲートの代わりにはしない。
