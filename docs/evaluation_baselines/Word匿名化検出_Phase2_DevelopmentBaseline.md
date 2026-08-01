# Word匿名化検出 Phase2 Development Baseline

この文書は、Word匿名化v1 Phase 2の候補検出をdevelopment合成Word上のBaselineとして固定する記録です。

Phase 2は、Wordファイルを書き換えず、Phase 1で抽出したWord内部構造と文字列から匿名化候補を列挙する段階です。仮名化、レビューUI、承認/解除、内部XML残存削除、外部共有用出力は含みません。

## Development Baseline

| 指標 | 値 |
|---|---:|
| ground_truth_count | `16` |
| candidate_count | `16` |
| matched_ground_truth_count | `16` |
| missed_ground_truth_count | `0` |
| true_positive_candidate_count | `16` |
| false_positive_candidate_count | `0` |
| recall | `1.0` |
| candidate_precision | `1.0` |
| exact_span_match_count | `16` |

## カテゴリ別結果

| カテゴリ | 正解 | 検出 | 見逃し | TP候補 | FP候補 |
|---|---:|---:|---:|---:|---:|
| 会社名 | `5` | `5` | `0` | `5` | `0` |
| 氏名 | `5` | `5` | `0` | `5` | `0` |
| 住所 | `1` | `1` | `0` | `1` | `0` |
| 電話番号 | `1` | `1` | `0` | `1` | `0` |
| メールアドレス | `3` | `3` | `0` | `3` | `0` |
| 銀行名 | `1` | `1` | `0` | `1` | `0` |

## 注意

この結果はdevelopment用の合成Wordに対する結果です。validation、完全ホールドアウト、未知Word、実運用文書での性能を意味しません。

この記録には、実在する氏名、会社名、住所、顧客データ、正解文字列、秘密情報を含めません。
