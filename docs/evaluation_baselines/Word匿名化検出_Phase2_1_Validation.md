# Word匿名化検出 Phase2.1 Validation

Validation Initialを上書きせず、同一validation Wordと固定ground truthでPhase 2.1を評価した記録です。

Phase 2.1の `recall=1.0` / `candidate_precision=1.0` は、Validation Initialの結果を原因分析して1回ルール改善した後のpost-tuning validation値です。未知Wordに対する汎化性能、完全ホールドアウト性能としては扱いません。

このコミット以降、完全ホールドアウト評価が終了するまで候補生成ルールを凍結します。

## Development Baseline

| 指標 | 値 |
|---|---:|
| ground_truth_count | `16` |
| matched_ground_truth_count | `16` |
| false_positive_candidate_count | `0` |
| recall | `1.0` |
| candidate_precision | `1.0` |

## Overall

| 指標 | Validation Initial | Phase 2.1 |
|---|---:|---:|
| ground_truth_count | `24` | `24` |
| candidate_count | `26` | `24` |
| matched_ground_truth_count | `20` | `24` |
| missed_ground_truth_count | `4` | `0` |
| true_positive_candidate_count | `20` | `24` |
| false_positive_candidate_count | `6` | `0` |
| recall | `0.833` | `1.0` |
| candidate_precision | `0.769` | `1.0` |
| exact_span_match_count | `20` | `24` |

## Category Results

| カテゴリ | 指標 | Initial | Phase 2.1 |
|---|---|---:|---:|
| メールアドレス | 正解 | `3` | `3` |
| メールアドレス | 検出 | `3` | `3` |
| メールアドレス | 見逃し | `0` | `0` |
| メールアドレス | TP候補 | `3` | `3` |
| メールアドレス | FP候補 | `0` | `0` |
| 会社名 | 正解 | `7` | `7` |
| 会社名 | 検出 | `8` | `7` |
| 会社名 | 見逃し | `2` | `0` |
| 会社名 | TP候補 | `5` | `7` |
| 会社名 | FP候補 | `3` | `0` |
| 住所 | 正解 | `3` | `3` |
| 住所 | 検出 | `4` | `3` |
| 住所 | 見逃し | `1` | `0` |
| 住所 | TP候補 | `2` | `3` |
| 住所 | FP候補 | `2` | `0` |
| 氏名 | 正解 | `6` | `6` |
| 氏名 | 検出 | `6` | `6` |
| 氏名 | 見逃し | `1` | `0` |
| 氏名 | TP候補 | `5` | `6` |
| 氏名 | FP候補 | `1` | `0` |
| 銀行名 | 正解 | `3` | `3` |
| 銀行名 | 検出 | `3` | `3` |
| 銀行名 | 見逃し | `0` | `0` |
| 銀行名 | TP候補 | `3` | `3` |
| 銀行名 | FP候補 | `0` | `0` |
| 電話番号 | 正解 | `2` | `2` |
| 電話番号 | 検出 | `2` | `2` |
| 電話番号 | 見逃し | `0` | `0` |
| 電話番号 | TP候補 | `2` | `2` |
| 電話番号 | FP候補 | `0` | `0` |

## Notes

- Phase 2.1はValidationを100%にすることを目的にしていません。
- Phase 2.1はValidation Initialの原因分析後に実施した1回の一般ルール改善結果です。
- Validationデータは今後、Phase 2.1までの開発・調整済み評価データとして扱います。
- Phase 2.1の結果は未知Word性能でも完全ホールドアウト性能でもありません。
- 候補生成コードはvalidation ground truth、truth_id、固定位置、dataset splitを参照していません。
- Phase 3 UI、Word書き換え、holdout評価は実施していません。
