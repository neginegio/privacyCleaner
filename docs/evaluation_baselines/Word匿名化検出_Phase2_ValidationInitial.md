# Word匿名化検出 Phase2 Validation Initial

Phase 2 baselineコードを変更せず、validation合成Wordに対して初回候補生成を実行した記録です。

- validation_docx_sha256: `028972970592182d8c8ccc1a48e0e591998c5ac1f6b088ad090f8efdc95fb1a8`
- ground_truth_sha256: `a036fd2711f0fd25c4b8d352db57566720ca14a6c17688952d697904f0b91810`
- 評価方法: 同一location_id、カテゴリ一致、char_start完全一致、char_end完全一致を正式TPとする。
- この結果を見たルール改善はまだ実施していません。
- Development baselineの `recall=1.0` / `candidate_precision=1.0` とは別の、初回validation固定成績です。
- 今後Phase 2の検出ルールを改善しても、このValidation Initial記録は書き換えません。

## Metrics

| 指標 | 値 |
|---|---:|
| ground_truth_count | `24` |
| candidate_count | `26` |
| matched_ground_truth_count | `20` |
| missed_ground_truth_count | `4` |
| true_positive_candidate_count | `20` |
| false_positive_candidate_count | `6` |
| recall | `0.833` |
| candidate_precision | `0.769` |
| exact_span_match_count | `20` |

## カテゴリ別結果

| カテゴリ | 正解 | 検出 | 見逃し | TP候補 | FP候補 |
|---|---:|---:|---:|---:|---:|
| メールアドレス | `3` | `3` | `0` | `3` | `0` |
| 会社名 | `7` | `8` | `2` | `5` | `3` |
| 住所 | `3` | `4` | `1` | `2` | `2` |
| 氏名 | `6` | `6` | `1` | `5` | `1` |
| 銀行名 | `3` | `3` | `0` | `3` | `0` |
| 電話番号 | `2` | `2` | `0` | `2` | `0` |
