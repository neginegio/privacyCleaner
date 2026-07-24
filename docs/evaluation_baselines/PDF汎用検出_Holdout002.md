# PDF汎用検出 Holdout002 正式評価記録

これは、PDF汎用候補生成Generation3に対する、完全未使用PDFによる第2回ホールドアウト評価です。評価対象は17ページ、正解49件の単一文書です。

## 評価の固定条件

- 使用Generation: `pdf-generation3`
- 使用commit: `2ab576ddc75c7493ad9629675ef145dde7a1dc32`
- Generation3は評価前に固定済み。
- 初回候補は、正解データ作成前に固定済み。
- 正解データは、初回候補を表示せず、候補枠を重ねず、独立して人間が作成。
- PDF、初回候補CSV、正解CSVはSHA-256で固定。
- 評価結果を見て、候補生成ルール、OCR設定、評価方法、閾値は変更していない。
- 原因分析とGeneration4検討は実施していない。
- 混入防止検査は `9/9 PASS`。

## 入力の固定

| 対象 | SHA-256 |
|---|---|
| 対象PDF | `ce8a2f4749381fcd21bbb392190e6fd2f84ee339a955c1a4f229e9b4889d362b` |
| 初回候補CSV | `80fd87127639becc522c3a1bf4cf8046ac204042125415dbc569953bac44668f` |
| 正解CSV | `afaa5fc3887e603e924cfd49cf6da7cba31efacc13a7afdd4b4df77571c8aa5c` |

この記録には、PDF原本、OCR原文、実際の会社名、実際の氏名、正解座標、候補座標、正解CSVの内容、初回候補CSVの内容は含めない。

## 評価方法の同一性

Holdout002の評価は、Generation3開発時に確定した正式評価方法と同じ定義で実施した。

- 矩形マッチング条件: 同一ページ、正規化後カテゴリ一致、`coverage >= 0.85`
- coverage: `intersection_area / ground_truth_area`
- `matched_ground_truth_count`: 1つ以上の有効候補に一致した正解矩形数
- `true_positive_candidate_count`: 1つ以上の正解矩形と有効に一致した候補数
- `false_positive_candidate_count`: 正解矩形と一致しなかった候補数
- `recall`: `matched_ground_truth_count / ground_truth_count`
- `candidate_precision`: `true_positive_candidate_count / candidate_count`

既存の共通評価関数そのものは呼び出していないが、Git管理外のHoldout002評価スクリプトで同じ評価定義を複製して使用した。Holdout002だけに有利または不利になる閾値や特別処理は入れていない。

## OCR状態

| 項目 | 値 |
|---|---:|
| REVIEW_REQUIREDページ数 | `17` |
| FAILEDページ数 | `0` |

## 正式評価結果

| 指標 | 値 |
|---|---:|
| ground_truth_count | `49` |
| candidate_count | `2` |
| matched_ground_truth_count | `0` |
| missed_ground_truth_count | `49` |
| true_positive_candidate_count | `0` |
| false_positive_candidate_count | `2` |
| recall | `0.000` |
| candidate_precision | `0.000` |
| average_coverage | `0.000` |
| minimum_coverage | `0.000` |

## カテゴリ別結果

| カテゴリ | 正解数 | 検出数 | 見逃し数 |
|---|---:|---:|---:|
| 会社名 | `49` | `0` | `49` |

候補側では、氏名候補1件と氏名1件が生成され、いずれもTP 0件、FP 1件だった。

## 正解作成GUIの操作性改善

Holdout002の正解作成支援として、Git管理外の正解レビューGUIに次の操作性改善を行った。

- 枠の端・角のドラッグによるサイズ変更
- `Ctrl + マウスホイール` による拡大縮小
- ウィンドウ最大化
- PDF表示領域最大化

このGUI改善は、正解作成支援だけを目的としたものであり、Generation3候補生成、OCR、評価方法、初回候補には影響していない。また、正解作成GUIでは初回候補CSV、候補座標、候補件数、候補理由を表示・参照していない。

## Holdout参考集計

以下は2文書に対する参考集計であり、原因分析やルール改善はまだ行っていない。

| 評価 | 正解 | 検出 | 候補 | TP | FP |
|---|---:|---:|---:|---:|---:|
| Holdout001 | `10` | `0` | `5` | `0` | `5` |
| Holdout002 | `49` | `0` | `2` | `0` | `2` |
| 2文書累計（参考） | `59` | `0` | `7` | `0` | `7` |

## 解釈

Generation3の完全未使用PDFによる第2回ホールドアウト評価では、17ページ・正解49件の単一文書に対してRecall `0.000` / Candidate Precision `0.000`だった。

ただし、これは1つの文書に対する正式評価結果であり、Generation3の未知文書全般の性能が0であると断定するものではない。

## データセットの位置づけ

今回のPDFは、正式ホールドアウト評価を完了したため、`Generation3 完全ホールドアウト評価②（評価済み）`として固定する。この正式評価結果は今後変更しない。

今後このPDFを原因分析やGeneration4開発に利用する場合、その時点からこのPDFは開発・分析データとして扱い、Generation4以降の完全ホールドアウト評価には再利用しない。
