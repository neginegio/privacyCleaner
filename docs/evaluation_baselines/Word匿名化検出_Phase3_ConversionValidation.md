# Word匿名化検出 Phase3 ConversionValidation

Phase 2.1で凍結済みの検出結果(構造位置一致・文字範囲一致・カテゴリ一致)を再現したうえで、同一の検証用Word・ground truthに対して実際に scan() → convert() を実行し、設計書が定める4つ目の評価軸「置換後残存なし」を初めて検証した記録です。

## 再現性確認

| 指標 | Phase 2.1 凍結値 |
|---|---:|
| ground_truth_count | `24` |
| candidate_count | `24` |
| matched_ground_truth_count | `24` |
| missed_ground_truth_count | `0` |
| false_positive_candidate_count | `0` |
| recall | `1.0` |
| candidate_precision | `1.0` |

再現一致: `True`(検出ロジックは変更されていません)

## Conversion Metrics

| 指標 | 値 |
|---|---:|
| matched_ground_truth_count | `24` |
| converted_matched_candidate_count | `24` |
| converted_run_count | `22` |
| converted_property_count | `0` |
| review_required_count | `7` |
| skipped_hyperlink_target_count | `0` |
| warning_count | `1` |

## Residual Check(置換後残存なし)

| 指標 | 値 |
|---|---:|
| verdict | `pass` |
| convert_internal_residual_warning | `False` |
| matched_truth_readable_text_residual_count | `0` |
| matched_truth_internal_xml_residual_parts | `なし` |

## Notes

- 本記録は開発/検証データのみを対象とします。完全ホールドアウト評価(設計書の手順10)はこの記録の一部として実施していません。
- 候補生成ルールは変更していません(Phase 2.1 frozenのまま、再現性確認で照合済み)。
- 変換承認は全非ハイパーリンク候補を有効化した上で実施しています(レビュー承認を模擬)。
- 変換後の.docxおよびCSV/報告書/監査JSONの実ファイルはコミットしていません(一時ディレクトリで生成・破棄)。
