# PDF汎用検出 Holdout001 正式評価記録

これは、PDF汎用候補生成Generation3に対する完全未使用PDFの初回ホールドアウト評価です。評価対象は1ページ、正解10件の単一文書です。

## 評価の固定条件

- 使用Generation: `pdf-generation3`
- 使用commit: `2ab576ddc75c7493ad9629675ef145dde7a1dc32`
- Generation3は評価前に固定済み。
- 初回候補は、正解データ作成前に固定済み。
- 正解データは、初回候補を表示せず、候補枠を重ねず、独立して人間が作成。
- PDF、初回候補CSV、正解CSVはSHA-256で固定。
- 評価結果を見て、候補生成ルール、OCR設定、評価方法、閾値は変更していない。
- 混入防止検査は `9/9 PASS`。

## 入力の固定

| 対象 | SHA-256 |
|---|---|
| 対象PDF | `d8bf1260a3c17fa9bc7b06fe402c8d1a6af2793308bd0c51de8e5a0219a32559` |
| 初回候補CSV | `47299e98ee066a2d51feec9eb1a69961a1c1c94dddd2dcf8aa05c2e16f4c0684` |
| 正解CSV | `58c62f6780eff6879370fd436980544912549be2e5cf084aea036969d5d3def8` |

この記録には、PDF原本、OCR原文、会社名、氏名、正解座標、候補座標、正解CSVの内容、初回候補CSVの内容は含めない。

## OCR状態

| 項目 | 値 |
|---|---:|
| quality | `REVIEW_REQUIRED` |
| OCR文字数 | `1471` |
| 座標付きOCR word数 | `337` |
| FAILED | `なし` |

## 正式評価結果

| 指標 | 値 |
|---|---:|
| ground_truth_count | `10` |
| candidate_count | `5` |
| matched_ground_truth_count | `0` |
| missed_ground_truth_count | `10` |
| true_positive_candidate_count | `0` |
| false_positive_candidate_count | `5` |
| recall | `0.000` |
| candidate_precision | `0.000` |
| average_coverage | `0.000` |
| minimum_coverage | `0.000` |

## カテゴリ別結果

| カテゴリ | 正解数 | 検出数 | 見逃し数 |
|---|---:|---:|---:|
| 会社名 | `5` | `0` | `5` |
| 氏名 | `5` | `0` | `5` |

候補側では会社名候補が5件生成され、TPは0件、FPは5件だった。

## 解釈

開発データセットv1ではRecall `0.612` / Candidate Precision `0.714`だったが、完全ホールドアウト評価①では双方`0.000`となり、この帳票形式への汎化性能が確認できなかった。

ただし、これは1ページ、正解10件の単一文書に対する結果であり、「Generation3はすべての未知PDFで性能0」という意味ではない。

## データセットの位置づけ

今回のPDFは、正式ホールドアウト評価を完了したため、`Generation3 完全ホールドアウト評価①（評価済み）`として固定する。この正式評価結果は今後変更しない。

今後このPDFを原因分析やGeneration4開発に利用する場合、その時点からこのPDFは開発・分析データとして扱い、Generation4以降の完全ホールドアウト評価には再利用しない。
