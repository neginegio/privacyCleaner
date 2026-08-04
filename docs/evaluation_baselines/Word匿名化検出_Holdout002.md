# Word匿名化検出 Holdout002

Phase 2.1 frozen候補と、独立作成したHoldout002 ground truthを初めて比較した正式ホールドアウト評価記録です。

この公開記録には、実在する会社名、氏名、銀行名、本文、候補文字列、ground truth文字列を含めません。

## 固定対象

- Phase 2.1 commit: `ca7b01e+ginza_ner`
- Word SHA-256: `074ad48cb816f6e93f506fbd3c8b0e776320954aa439bbd35633c4f0d5311b14`
- candidate SHA-256: `56aa29a75539ec674cf3982d2f8323a205a20f3c3352884e5a1fc4680c24fd16`
- ground truth SHA-256: `689f2e30128857d2e373c616f5c0505cc382431654f0ed220f1bc902cb1630d4`

## 評価上の解釈

- Holdout002 is a completely unused real Word document evaluated after the Phase 2.1 frozen commit.
- Initial candidates were generated and fixed before ground truth creation.
- Human ground truth was created independently without viewing candidate contents.
- Candidate and ground truth were compared for the first time only after ground truth completion.
- This is an unknown-real-document evaluation for Phase 2.1.
- Development and post-tuning Validation 1.0/1.0 results must be kept separate from this holdout result.
- A single holdout document is not enough to make a general claim about Word-wide performance.
- Company names and personal names dropped substantially: company matched 0/1, person matched 2/3.
- 1 of 4 missed ground truths had no generated candidate, so the drop is not explained only by strict exact-span matching.

## Overall

| 指標 | 値 |
|---|---:|
| ground_truth_count | `9` |
| candidate_count | `23` |
| matched_ground_truth_count | `5` |
| missed_ground_truth_count | `4` |
| true_positive_candidate_count | `5` |
| false_positive_candidate_count | `18` |
| recall | `0.556` |
| candidate_precision | `0.217` |
| exact_span_match_count | `5` |

## Category Results

| カテゴリ | GT | Candidates | Matched | Missed | TP Candidate | FP Candidate |
|---|---:|---:|---:|---:|---:|---:|
| メールアドレス | `1` | `1` | `1` | `0` | `1` | `0` |
| 会社名 | `1` | `7` | `0` | `1` | `0` | `7` |
| 住所 | `1` | `1` | `0` | `1` | `0` | `1` |
| 氏名 | `3` | `12` | `2` | `1` | `2` | `10` |
| 銀行名 | `1` | `0` | `0` | `1` | `0` | `0` |
| 電話番号 | `2` | `2` | `2` | `0` | `2` | `0` |

## Cause Summary

### Missed

| 原因カテゴリ | 件数 |
|---|---:|
| span過大 | `3` |
| 候補自体が生成されなかった | `1` |

### False Positive

| 原因カテゴリ | 件数 |
|---|---:|
| span過大 | `3` |
| 一般語誤検出 | `15` |

## Conversion Residual Check(置換後残存なし)

- verdict: `pass`
- converted_matched_candidate_count: `5`

## Notes

- metadata ground truth count: `1`
- document propertiesとunsupported領域は本文系Phase 2.1評価のRecall/Precisionへ混ぜていません。
- 評価結果を見た後の候補生成ルール変更は実施していません。
