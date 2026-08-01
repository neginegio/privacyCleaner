# Word匿名化検出 Phase2 Validation Initial Error Analysis

この文書は、`441d588 Add Word candidate detection Phase 2 baseline` の初回validation結果に対する原因分析です。

候補生成ルールは変更していません。Phase 3、Word書き換え、holdout評価も実施していません。

## 固定済み初回結果

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

Development baselineは `recall=1.0` / `candidate_precision=1.0` ですが、これはdevelopment合成Word上の結果です。Validation Initialは未知寄りの合成validation Wordに対する、ルール改善前の固定成績です。

## Fixtureの扱い

validation用Word、ground truth、initial candidatesはいずれも合成データです。実在顧客情報、実在個人情報、実運用文書由来の機密情報は含みません。

Git管理候補:

- `docs/evaluation_baselines/word_phase2_validation/Word匿名化検出_Phase2_ValidationSynthetic.docx`: 合成validation Word。再現性のためfixtureとして管理可能。
- `docs/evaluation_baselines/word_phase2_validation/Word匿名化検出_Phase2_ValidationGroundTruth.csv`: 固定ground truth。再現性のためfixtureとして管理可能。
- `tools/evaluate_word_phase2_validation.py`: validation生成・評価ツール。検出ルール本体ではないため管理可能。
- `docs/evaluation_baselines/Word匿名化検出_Phase2_ValidationInitial.{json,csv,md}`: 初回評価の正式集計記録として管理対象。
- `docs/evaluation_baselines/Word匿名化検出_Phase2_DevelopmentBaseline.{json,csv,md}`: development baselineの正式集計記録として管理対象。

生成物扱いを検討:

- `docs/evaluation_baselines/word_phase2_validation/Word匿名化検出_Phase2_ValidationInitialCandidates.csv`: ツールから再生成可能。正式集計記録が別にあるため、Git管理は任意。ただし初回candidate SHA-256を固定したい場合は同時に管理してよい。

配置方針案:

- 固定fixture: `tests/fixtures/word/phase2_validation/`
- 評価集計: `docs/evaluation_baselines/`
- 評価ツール: `tools/evaluate_word_phase2_validation.py`

今回は移動・削除していません。

## 見逃し4件

| ID | category | location_id | 正解span | paragraph全体 | run分割 | 基礎検出器 | Word regex | 最終候補にならなかった理由 | 原因分類 | 改善可能性 |
|---|---|---|---|---|---|---|---|---|---|---|
| WORDVAL-002 | 会社名 | `body:None::paragraph:None:None:None:1:None:/word/document.xml:body/p[1]` | `5-13` | `配送担当は北斗物流株式会社です。` | `run0: 0-16` | `北斗物流株式会社` を `氏名` として返した | `配送担当` を会社名、`株式会社` を会社名として返した | 正解span・categoryの会社名候補が出ず、同じspanは氏名扱いになった | D span範囲不一致 + E カテゴリ正規化 | 汎用ルールで安全に改善可能 |
| WORDVAL-004 | 会社名 | `body:None::paragraph:None:None:None:3:None:/word/document.xml:body/p[3]` | `13-17` | `法人格を含まない組織として白樺ラボを記載します。` | `run0: 0-24` | なし | なし | 法人格なし・既存ラベルなしの組織名に対応する候補が出なかった | A 基礎検出器が候補なし + B regex条件不足 | 改善可能だがFP増加リスクあり |
| WORDVAL-006 | 氏名 | `body:None::paragraph:None:None:None:7:None:/word/document.xml:body/p[7]` | `4-8` | `連絡者は鈴木花子です。` | `run0: 0-11` | なし | なし | 空白なし氏名で、`連絡者は` が既存氏名ラベル条件にない | A 基礎検出器が候補なし + B regex条件不足 | 改善可能だがFP増加リスクあり |
| WORDVAL-011 | 住所 | `body:None::paragraph:None:None:None:12:None:/word/document.xml:body/p[12]` | `5-27` | `住所3: 京都府京都市下京区烏丸通1-2 青葉ビル5階` | `run0: 0-27` | `5-20` の `京都府京都市下京区烏丸通1-2` を住所として返した | なし | 建物名部分までspanが伸びず、完全一致条件では見逃しになった | D span範囲不一致 | 改善可能だがFP増加リスクあり |

## FP6件

| candidate | category | location_id | span | detection_rule | source | confidence | 周辺文脈 | paragraph全体 | 誤認理由 | 原因分類 | 改善可能性 |
|---|---|---|---|---|---|---:|---|---|---|---|---|
| `0372044657ca` | 会社名 | `body:None::paragraph:None:None:None:1:None:/word/document.xml:body/p[1]` | `0-4` | `word_company_suffix` | paragraph | `0.840` | `配送担当` + `は北斗物流株式会社です。` | `配送担当は北斗物流株式会社です。` | 後置会社名regexが文頭から `株式会社` までを広く拾い、trimで `は` より前の一般語だけが残った | A regexが広すぎる | 汎用ルールで安全に改善可能 |
| `1ba70c74a3cf` | 氏名 | `body:None::paragraph:None:None:None:1:None:/word/document.xml:body/p[1]` | `5-13` | `JapanesePresidioDetector` | paragraph | `0.830` | `配送担当は` + `です。` | `配送担当は北斗物流株式会社です。` | `担当は...です` の氏名パターンが会社名を氏名として拾った | B Presidio基礎検出の誤認 | 汎用ルールで安全に改善可能 |
| `a63d4a531232` | 会社名 | `body:None::paragraph:None:None:None:1:None:/word/document.xml:body/p[1]` | `9-13` | `word_company_prefix` | paragraph | `0.880` | `配送担当は北斗物流` + `です。` | `配送担当は北斗物流株式会社です。` | 前置会社名regexが後置の `株式会社` だけを会社名として残した | G span過小 | 汎用ルールで安全に改善可能 |
| `7b29f6e4b4d2` | 会社名 | `body:None::paragraph:None:None:None:2:None:/word/document.xml:body/p[2]` | `0-5` | `word_company_suffix` | paragraph | `0.840` | `共同研究先` + `は合同会社青葉研究所です` | `共同研究先は合同会社青葉研究所です。` | 後置会社名regexが文頭から `合同会社` までを広く拾い、trimで一般語だけが残った | A regexが広すぎる | 汎用ルールで安全に改善可能 |
| `04132d9a0968` | 住所 | `body:None::paragraph:None:None:None:12:None:/word/document.xml:body/p[12]` | `5-20` | `JapanesePresidioDetector` | paragraph | `0.820` | `住所3: ` + ` 青葉ビル5階` | `住所3: 京都府京都市下京区烏丸通1-2 青葉ビル5階` | 住所の一部としては正しいが、正式評価では建物名を含む完全spanと一致しない | G span過小 | 改善可能だがFP増加リスクあり |
| `4b74173c7381` | 住所 | `body:None::paragraph:None:None:None:13:None:/word/document.xml:body/p[13]` | `12-19` | `JapanesePresidioDetector` | paragraph | `0.680` | `住所に似た一般文として、` + `について説明します。` | `住所に似た一般文として、市区町村の制度について説明します。` | 市/区/町/村を含む一般説明文を municipality address と誤認した | E 住所らしい一般文章 | 汎用ルールで安全に改善可能 |

## 原因カテゴリ別件数

| 原因カテゴリ | 件数 |
|---|---:|
| A 基礎検出器が候補なし / regexが広すぎる | `4` |
| B regex条件不足 / Presidio基礎検出の誤認 | `3` |
| D span範囲不一致 | `2` |
| E カテゴリ正規化 | `1` |
| G span過小 | `2` |

複数原因を持つケースがあるため、合計は10件を超えます。

## 改善可能性

| 分類 | 件数 |
|---|---:|
| 汎用ルールで安全に改善可能 | `6` |
| 改善可能だがFP増加リスクあり | `4` |
| 既存検出器側の問題 | `0` |
| 追加研究が必要 | `0` |

既存検出器の誤認はありますが、Word側の後処理・競合解決で扱える範囲と判断します。

## 次のPhase 2改良案

### 会社名

現在の弱点:

- 後置の `株式会社` / `合同会社` で、文頭の一般語から法人格までを広く取りすぎる。
- 前置会社名regexが、後置位置の `株式会社` 単体を候補化する。
- 法人格なしの組織名に弱い。
- 会社名らしいspanと氏名spanが競合したとき、会社名優先の判断がない。

改善案:

- 後置法人格regexは、候補開始位置を助詞・句読点・空白・文頭境界の直後に限定し、`は` などをまたがない。
- `株式会社` 等の法人格単体は会社名候補から除外する。
- 法人格を含む文字列が氏名候補になった場合、会社名へ再分類または氏名候補を抑制する。
- 法人格なし組織名は、`組織として` / `取引先` / `共同研究先` などの文脈語と、`ラボ` / `研究所` 等の汎用組織語を組み合わせる。

期待できる改善:

- `北斗物流株式会社` の見逃しと、`配送担当` / `株式会社` / `共同研究先` のFPを減らせる。

想定される副作用:

- 法人格なし組織名の検出を広げると、一般語や部署名のFPが増える可能性がある。

### 住所

現在の弱点:

- 都道府県・市区町村・番地までは取れるが、空白後の建物名・階数を取りこぼす。
- `市区町村の制度` のような一般説明文を住所として拾う。

改善案:

- 住所候補の直後に建物名らしい語尾（ビル、マンション、階、号室など）が続く場合だけspanを拡張する。
- municipality address は番地・数字・丁目・番地・号などの住所強度シグナルを必須化または加点する。
- `制度` / `説明` / `について` など説明文らしい語尾は住所候補から除外する。

期待できる改善:

- 建物名付き住所の完全span一致が改善し、一般説明文のFPを減らせる。

想定される副作用:

- 建物名拡張を広げすぎると、住所後続の説明文まで取り込むspan過大が起きる。

### 氏名

現在の弱点:

- 空白なし日本人氏名は、ラベルや敬称が弱いと検出できない。
- `担当は...です` のような文脈で会社名を氏名として拾う。

改善案:

- 氏名ラベルに `連絡者` / `確認者` などの汎用人物ロールを追加する。
- 空白なし氏名はラベル直後など文脈限定で扱い、単独出現では広げすぎない。
- 法人格・金融機関語・組織語を含むspanは氏名候補から抑制する。

期待できる改善:

- `連絡者は...です` の空白なし氏名を拾いやすくなり、会社名の氏名誤認を減らせる。

想定される副作用:

- 人物ロール語を増やすと、役職名や部署名を氏名として拾うリスクがある。

## 混入防止

原因分析ではground truthを参照していますが、候補生成コードには取り込んでいません。

再確認対象:

- `ground_truth`
- `truth_id`
- validation fixture固有文字列
- 固定paragraph/table/row/cell番号
- dataset split

候補生成本体 `src/excel_privacy_cleaner/word_processor.py` は、これらを参照していません。
