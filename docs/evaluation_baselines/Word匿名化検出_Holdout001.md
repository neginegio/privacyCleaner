# Word匿名化検出 Holdout001

Phase 2.1 frozen候補と、独立作成したHoldout001 ground truthを初めて比較した正式ホールドアウト評価記録です。

この公開記録には、実在する会社名、氏名、銀行名、本文、候補文字列、ground truth文字列を含めません。

## 固定対象

- Phase 2.1 commit: `ca7b01e`
- Word SHA-256: `de87ba901b268978506b72c5b52c7d5a4a6c07f35c4e924411fc6e02cf06a376`
- candidate SHA-256: `b54b3ceb19a92a53b81ec57317483a8bd49e42c32ffa90af418bbfd18c4f620f`
- ground truth SHA-256: `85da8b98fbdd07c5084d11304f9a2d3813c6204e1100a19deafc58a7727fcf59`

## 評価上の解釈

- Holdout001 is a completely unused real Word document evaluated after the Phase 2.1 frozen commit.
- Initial candidates were generated and fixed before ground truth creation.
- Human ground truth was created independently without viewing candidate contents.
- Candidate and ground truth were compared for the first time only after ground truth completion.
- This is an unknown-real-document evaluation for Phase 2.1.
- Development and post-tuning Validation 1.0/1.0 results must be kept separate from this holdout result.
- A single holdout document is not enough to make a general claim about Word-wide performance.
- Company names and personal names dropped substantially: company matched 0/5, person matched 0/6.
- 9 of 13 missed ground truths had no generated candidate, so the drop is not explained only by strict exact-span matching.

## Overall

| 指標 | 値 |
|---|---:|
| ground_truth_count | `20` |
| candidate_count | `13` |
| matched_ground_truth_count | `7` |
| missed_ground_truth_count | `13` |
| true_positive_candidate_count | `7` |
| false_positive_candidate_count | `6` |
| recall | `0.35` |
| candidate_precision | `0.538` |
| exact_span_match_count | `7` |

## Category Results

| カテゴリ | GT | Candidates | Matched | Missed | TP Candidate | FP Candidate |
|---|---:|---:|---:|---:|---:|---:|
| メールアドレス | `0` | `0` | `0` | `0` | `0` | `0` |
| 会社名 | `5` | `3` | `0` | `5` | `0` | `3` |
| 住所 | `0` | `0` | `0` | `0` | `0` | `0` |
| 氏名 | `6` | `1` | `0` | `6` | `0` | `1` |
| 銀行名 | `9` | `9` | `7` | `2` | `7` | `2` |
| 電話番号 | `0` | `0` | `0` | `0` | `0` | `0` |

## Cause Summary

### Missed

| 原因カテゴリ | 件数 |
|---|---:|
| span過大 | `2` |
| カテゴリ誤分類 | `2` |
| 候補自体が生成されなかった | `9` |

### False Positive

| 原因カテゴリ | 件数 |
|---|---:|
| span過大 | `3` |
| カテゴリ誤分類 | `2` |
| 一般語誤検出 | `1` |

## Conversion Residual Check(置換後残存なし)

- verdict: `blocked`
- blocked_guard: `overlap_guard`
- converted_matched_candidate_count: `7`

## Notes

- metadata ground truth count: `2`
- document propertiesとunsupported領域は本文系Phase 2.1評価のRecall/Precisionへ混ぜていません。
- 評価結果を見た後の候補生成ルール変更は実施していません。
