# PDF汎用検出 Holdout003 正式評価記録

これは、PDF汎用候補生成Generation3に対する、完全未使用PDFによる第3回ホールドアウト評価です。評価対象は16ページ、正解43件の単一文書です。

## 評価の固定条件

- 使用Generation: `pdf-generation3`
- 使用commit: `2ab576ddc75c7493ad9629675ef145dde7a1dc32`
- Generation3は評価前に固定済み。
- 初回候補は、正解データ作成前に固定済み。
- 正解データは、初回候補を表示せず、候補枠を重ねず、独立して人間が作成。
- PDF、初回候補CSV、正解CSVはSHA-256で固定。
- 評価結果を見て、候補生成ルール、OCR設定、評価方法、閾値は変更していない。
- 原因分析とGeneration4検討は実施していない。

## 入力の固定

| 対象 | SHA-256 |
|---|---|
| 対象PDF | `1fb4664482d940990df2c991a1159f2d020361eab8dd5103f5f419f5a28d91dc` |
| 初回候補CSV | `273be3008f8241c29a79ee90dda39cf4bd46c7fc5420f4290dcf95a207e2f849` |
| 正解CSV | `467c3c88130b3acbec8c9d7298326a9f89f638fa00c0e8316a9198f1067366ed` |

この記録には、PDF原本、OCR原文、実際の会社名、実際の氏名、正解座標、候補座標、正解CSVの内容、初回候補CSVの内容は含めない。

## 評価方法の同一性

Holdout003の評価は、Generation3開発時およびHoldout001/002で使用した正式評価方法と同じ定義で実施した。

- 矩形マッチング条件: 同一ページ、正規化後カテゴリ一致、`coverage >= 0.85`
- coverage: `intersection_area / ground_truth_area`
- `matched_ground_truth_count`: 1つ以上の有効候補に一致した正解矩形数
- `true_positive_candidate_count`: 1つ以上の正解矩形と有効に一致した候補数
- `false_positive_candidate_count`: 正解矩形と一致しなかった候補数
- `recall`: `matched_ground_truth_count / ground_truth_count`
- `candidate_precision`: `true_positive_candidate_count / candidate_count`
- カテゴリ正規化は `個人名 -> 氏名` のみ。
- `組織名 -> 会社名` の正規化は行っていない。

## 正式評価結果

| 指標 | 値 |
|---|---:|
| ground_truth_count | `43` |
| candidate_count | `44` |
| matched_ground_truth_count | `0` |
| missed_ground_truth_count | `43` |
| true_positive_candidate_count | `0` |
| false_positive_candidate_count | `44` |
| recall | `0.000` |
| candidate_precision | `0.000` |
| average_coverage | `0.000` |
| minimum_coverage | `0.000` |

## 正解カテゴリ別結果

| カテゴリ | 正解数 | 検出数 | 見逃し数 |
|---|---:|---:|---:|
| 会社名 | `25` | `0` | `25` |
| 住所 | `3` | `0` | `3` |
| 氏名 | `15` | `0` | `15` |
| 銀行名 | `0` | `0` | `0` |

## 候補カテゴリ別結果

| 候補カテゴリ | TP | FP |
|---|---:|---:|
| 住所 | `0` | `25` |
| 郵便番号 | `0` | `2` |
| 電話番号 | `0` | `2` |
| メールアドレス | `0` | `1` |
| 組織名 | `0` | `12` |
| 氏名 | `0` | `2` |

郵便番号、電話番号、メールアドレス、組織名も初回候補として `candidate_count` に含めた。正式正解カテゴリと一致しない候補は、既存評価方法どおりFPとして扱った。

## Holdout参考集計

以下は3文書に対する参考集計であり、原因分析やルール改善はまだ行っていない。

| 評価 | 正解 | 検出 | 候補 | TP | FP |
|---|---:|---:|---:|---:|---:|
| Holdout001 | `10` | `0` | `5` | `0` | `5` |
| Holdout002 | `49` | `0` | `2` | `0` | `2` |
| Holdout003 | `43` | `0` | `44` | `0` | `44` |
| 3文書累計（参考） | `102` | `0` | `51` | `0` | `51` |

## 解釈

Generation3の完全未使用PDFによる第3回ホールドアウト評価では、16ページ・正解43件の単一文書に対してRecall `0.000` / Candidate Precision `0.000`だった。

ただし、これは1つの文書に対する正式評価結果であり、Generation3の未知文書全般の性能が0であると断定するものではない。

## データセットの位置づけ

今回のPDFは、正式ホールドアウト評価を完了したため、`Generation3 完全ホールドアウト評価③（評価済み）`として固定する。この正式評価結果は今後変更しない。

今後このPDFを原因分析やGeneration4開発に利用する場合、その時点からこのPDFは開発・分析データとして扱い、Generation4以降の完全ホールドアウト評価には再利用しない。
