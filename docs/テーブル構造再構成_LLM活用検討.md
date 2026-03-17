# テーブル構造再構成 — LLM活用検討

## 1. 本ドキュメントの位置づけ

Excel/Word のテーブル変換において、現行のヒューリスティクスベースの Markdown 変換を LLM ベースの再構成に切り替える方針について、ディスカッション内容と検討結果を記録する。

---

## 2. 問題の発端

### 2.1 きっかけ: Excel 2列テーブルの縦横ヘッダー誤判定

`change_history.xlsx` の「設計概要」シートに以下のような2列テーブルがある:

| A列 (項目) | B列 (内容) |
|---|---|
| 文書名 | ドキュメント処理パイプライン詳細設計書 |
| アーキテクチャ | マイクロサービス |
| DB | PostgreSQL |
| 帳票 | PDF / Excel |

Excel 上では1行目「項目 / 内容」がカラムヘッダー（オートフィルタ付き）で、A列がラベル、B列が値というキーバリュー構造。

#### 現行の出力（問題あり）

```
[行1]
  項目: 文書名
  内容: ドキュメント処理パイプライン詳細設計書

[行2]
  項目: アーキテクチャ
  内容: マイクロサービス
```

1行目の「項目 / 内容」がカラムヘッダーとして採用され、全データ行が `項目: 文書名 / 内容: ドキュメント処理...` という形式で出力される。

#### 期待する出力

```
文書名: ドキュメント処理パイプライン詳細設計書
アーキテクチャ: マイクロサービス
DB: PostgreSQL
帳票: PDF / Excel
```

このテーブルは縦方向にヘッダー（A列がラベル）を持つキーバリュー型であり、1行目の「項目 / 内容」はメタヘッダー（列の役割を示す記述）として省略可能。

### 2.2 問題の根本原因

表面的には「2列テーブルの縦横判定」の問題だが、ディスカッションを通じて以下のより根本的な設計課題が浮上した:

1. **`to_markdown.py` の構造情報破棄**: 中間 JSON には `row/col/rowspan/colspan` が保持されているが、Markdown 変換時に全て捨てられ、代わりにヒューリスティクスで再推定している
2. **行単位の逐次処理**: テーブル全体を見た判断ができず、行ごとに「ヘッダーか？バナーか？フォームか？」を判定するため、全体の構造（縦横ヘッダーの方向等）を把握できない
3. **参考資料の思想との乖離**: Task.md §12 で参考にした RAG x FINDY スライドの「構造を取ってから再構成する」思想に反して、「構造を捨ててから再推定する」実装になっている

---

## 3. ディスカッションの経緯

### 3.1 初期アプローチ: ヒューリスティクスの改善

最初は `to_markdown.py` に縦方向ヘッダーの検出ロジックを追加する方向で検討。

- `_is_vertical_header_table()` 関数の追加
- 左列の値パターン分析（数値・日付 → データテーブル型、テキストラベル → KV型）
- `_render_vertical_header_table()` でキーバリュー形式出力

しかし **「このドキュメントではそうだが、他のドキュメントでは話が変わる。汎用的にするにはどうするか」** という指摘があり、ヒューリスティクスの限界が認識された。

### 3.2 構造化アプローチへの転換

次に「テーブル全体を構造化した状態（JSON形式）で扱う」方向が提案された。

- 現状は1行ずつ取得して判定しているが、テーブル全体を見て構造分析すべき
- Extractor と Transform の間に Analyzer ステップを挟む案（B案）が有力に

### 3.3 参考資料への立ち返り

ここで Task.md §12 の参考資料（https://speakerdeck.com/harumiweb/rag-findy?slide=14）の思想に立ち返った:

> Office 文書をそのまま Markdown やプレーンテキストに潰すのではなく、先に構造を取ってから再構成する

現状の実装は:

```
Step2: 構造を取る → ✓ row/col/rowspan/colspan を保持した JSON
Step3: 再構成する → ✗ 構造を全部捨ててフラットテキストに潰している
```

**「構造を取ってから再構成する」のではなく「構造を捨ててから再推定している」** という問題が明確になった。

### 3.4 LLM活用の方針決定

この認識を踏まえ、Step3 の「解釈・再構成」を LLM に委ねる方針に至った:

```
現状:  Excel → [Extractor] → JSON → [ヒューリスティクス] → Markdown → Dify
変更後: Excel → [Extractor] → JSON → [LLM再構成モード] → Markdown → Dify
```

これは Task.md §5.4 の設計原則にも合致:

> - 抽出は python-docx / COM で確定的に行う
> - 解釈・再構成は LLM で品質を上げる

**重要な認識**: LLM に渡すのは Dify ではなく、Dify に入れる前の前処理段階。LLM が構造化 JSON を理解し、最適な Markdown を生成し、その Markdown を Dify に投入する。

---

## 4. 設計方針

### 4.1 全体アーキテクチャ

```
Excel/Word → [Extractor] → 中間 JSON → [LLM 再構成モード] → Markdown → Dify
                                     └→ [LLMなしモード] → Markdown → Dify
```

- LLM を使用するモード: 中間 JSON をプロンプトとともに LLM に渡し、最適な Markdown を生成
- LLMなしモード: 既存実装をベースに半構造化 Markdown を生成する（§4.2 参照）

### 4.2 LLMなしでも動く設計

Task.md §5.4 の原則「LLM なしでも完走できる設計」を維持する。

`LLMなし` モードの出力方針として以下を検討:

| 案 | 内容 | 評価 |
|---|---|---|
| 案1 | 現行の `to_markdown.py` をベースに半構造化 Markdown を継続利用 | 既存資産を活かせる。出力の一貫性も保ちやすい |
| 案2 | シンプルに Markdown テーブルとして出力 | 構造保持はしやすいが、既存出力との乖離が大きい |
| 案3 | JSON を整形して出力 | 最もシンプルだが可読性が低い |

**案1を採用**: `LLMなし` モードの標準出力は、既存実装をベースにした半構造化 Markdown とする。
ただし、解釈が難しいテーブルも別形式へ逃がすのではなく、安全側の半構造化表現で出力する。
このとき、強い意味解釈は行わず、行単位・列単位・セル位置ベースで情報を欠損なく保持することを優先する。

### 4.3 プロンプト設計

#### 粒度の選択肢

| アプローチ | 単位 | メリット | デメリット |
|---|---|---|---|
| テーブル単位 | テーブルごとに LLM 呼び出し | 無関係な表を混ぜにくい、比較・再実行しやすい | 呼び出し回数が増える。文脈付与が必要 |
| シート全体 | シートごとに LLM 呼び出し | シート全体の文脈をまとめて扱える | 無関係な表が混ざりやすい。トークン消費も大きい |

**採用**: 原則としてテーブル単位で LLM に渡す。

理由: 1つのシートに複数の無関係なテーブルが存在するケースがあり、シート全体を 1 単位にすると不要な文脈が混ざって解釈を誤るリスクがあるため。
ただし、テーブル単体を裸で渡すのではなく、シート名、近接見出し、シート上の位置などの局所文脈を合わせて渡す。
将来的に、複数テーブルが同一の意味ブロックを構成すると判断できる場合は、テーブル群単位の扱いを追加検討する。

#### プロンプトに含めるべき情報

- シート名
- 対象テーブルの構造化 JSON（row/col/rowspan/colspan/is_header など）
- テーブルのシート上の元位置
- 近接する見出し・ラベル・補足要素などの局所文脈
- 出力形式の指示（Markdown、Dify チャンキングに適した形）

### 4.4 Word との統一

中間表現が同じ `CellData[][]` なので、LLM 再構成ロジックは Excel/Word 共通にできる。

#### 段階的適用計画

| Phase | 対象 | 内容 |
|---|---|---|
| Phase 1 | Excel | LLM 再構成を実装・検証。問題が顕在化しているため優先 |
| Phase 2 | Word | Excel で検証済みのロジックを Word にも適用 |

各 Phase で `LLMなし` モードも同時に実装する。

#### Word と Excel の違い

| | Word | Excel |
|---|---|---|
| テーブル検出 | python-docx が `doc.tables` で直接提供 | BFS で連結領域を検出（境界が曖昧） |
| テーブルの文脈 | 表の前後に段落・見出しがある | シート全体が表のようなもの |
| 複雑さの傾向 | 結合セル、多段ヘッダー | 結合セル + 複数テーブルの境界問題 |
| 現行の問題度 | 比較的安定 | 縦横判定等の問題が顕在化 |

### 4.5 段階的移行の方針

- 既存実装は、特に `LLMなし` モードにおけるベースラインとして活用する
- すでに確認できている出力品質と挙動を尊重し、全面的な作り直しは前提にしない
- 責務分離を進めながら、必要な箇所を段階的に差し替える
- モード別に同じ中間成果物から再実行できるようにし、既存実装との比較・検証を行いながら移行する

---

## 5. 検討が必要な残課題

1. **プロンプトの具体的な設計**: テーブル JSON をどう渡し、どういう指示を出すか。few-shot 例の設計
2. **コスト・速度の見積もり**: テーブル/シートあたりの API 費用と処理時間。特にテーブル単位で LLM を呼ぶ場合、1ファイル内のテーブル数に比例して呼び出し回数が増えるため、実サンプルで呼び出し回数と総処理時間を確認し、必要に応じてバッチ化可否を別途検討する
3. **サイズ上限の閾値**: シート単位 → テーブル単位への切り替え条件
4. **品質評価方法**: LLM 出力の品質をどう評価するか（既存テストケースとの比較等）
5. **Dify チャンキングとの相性検証**: LLM 生成 Markdown が Dify の分割で適切にチャンクされるか

### 5.1 議論テーマチェックリスト

以下は、今後のディスカッションで順に詰めるための論点一覧である。
結論が出た項目は `- [x]` に変更し、直下の `決定事項` に要点を追記する。

- [x] **LLM非依存の完走条件**
  - 決定事項:
    - LLM非依存モードの完走条件は、抽出済みの可視テキスト情報を欠損なく Markdown に出力できることとする。
    - この条件は最低条件であり、構造解釈の正しさや可読性は別の品質論点として扱う。
    - 非LLMモードでは表現の不自然さは許容するが、テキストの脱落・捏造・対応関係の取り違えは許容しない。

- [x] **運用モードと切り替え方針**
  - 決定事項:
    - 本設計の運用モードは `OpenAI API`、`ローカルLLM`、`LLMなし` の 3 つとする。
    - モードの切り替えは自動切り替えではなく、利用者または設定による明示的な選択で行う。
    - 選択したモードが実行不能な場合でも、別モードへ自動切り替えは行わない。
    - モードの切り替えは、その場での自動退避ではなく、中間成果物を起点とした再実行で行う。
    - Step1（正規化）および Step2（構造抽出）はモード非依存とし、Step3（再構成・Markdown化）をモードごとに再実行できる設計とする。

- [x] **責務分担の原則**
  - 決定事項:
    - 正規化はファイル形式の統一のみを責務とし、意味解釈は行わない。
    - 抽出は可視テキストおよび構造情報の取得を責務とし、欠損なく中間表現へ保持する。
    - 分析は再構成のための補助情報や候補分類を付与するが、元データを破壊せず、推定結果は事実ではなくヒントとして扱う。
    - 解釈はモード依存の責務とし、`OpenAI API` モードおよび `ローカルLLM` モードでは意味解釈を行い、`LLMなし` モードでは強い意味解釈を行わない。
    - Markdown の最終組み立てはレンダリングの責務として分離し、可能な限り deterministic に出力する。
    - 今後の修正は責務単位で行い、抽出の問題を解釈やレンダリングで吸収し続けるような修正は避ける。

- [x] **中間表現の要件**
  - 決定事項:
    - 中間表現は、`OpenAI API`、`ローカルLLM`、`LLMなし` の各モードで共通利用する source of truth とする。
    - 中間表現には、可視テキスト、要素順序、表の `row/col/rowspan/colspan`、見出し、段落、シート名など、再構成に必要な観測事実を欠損なく保持する。
    - Excel については、表内相対座標だけでなく、シート上の元の位置も保持できる設計とする。
    - 推定結果や解釈結果は観測事実と分離し、事実フィールドに混在させない。必要な場合は `hint` として別フィールドまたは別成果物で保持する。
    - 同一の中間表現から、モードを切り替えて Step3 を再実行できることを要件とする。
    - 中間表現または付随メタデータには、元ファイルへの追跡や比較・再生成のための識別情報を保持する。

- [x] **LLMなしモードの出力仕様**
  - 決定事項:
    - `LLMなし` モードの標準出力は、既存実装による半構造化 Markdown をベースラインとして継続利用する。
    - 解釈が難しいテーブルについても、別形式へ逃がさず、同じく半構造化 Markdown として出力する。
    - その際は強い意味解釈を行わず、行単位・列単位・セル位置ベースで情報を欠損なく保持することを優先する。
    - `LLMなし` モードでは、可読性向上のための軽い整形は許容するが、情報の省略や強い意味付けは行わない。

- [x] **品質劣化の許容範囲**
  - 決定事項:
    - 品質上の優先順位は、(1) 情報を欠損させないこと、(2) 意味関係を取り違えないこと、(3) 可読性を上げること、の順とする。
    - 可読性の低下や表現の不自然さは一定範囲で許容する。
    - 一方で、テキストの脱落、捏造、行列関係や対応関係の取り違え、キーと値の逆転など、事実を変えてしまう劣化は許容しない。
    - 解釈に迷う場合は、意味を補って自然に見せることよりも、安全側の半構造化表現を維持することを優先する。
    - `OpenAI API` モードおよび `ローカルLLM` モードは可読性と構造解釈の改善を担うが、上記の優先順位自体は全モード共通とする。

- [x] **テーブル類型の整理**
  - 決定事項:
    - テーブル類型の整理は、`input/excel` 配下の実サンプルをベースに行う。
    - ただし、サンプルに現れた形だけに閉じた専用設計にはせず、未知のケースにも対応できるよう抽象化して定義する。
    - 類型は、実サンプルで確認できる代表的な出力戦略を基準に整理する。
    - サンプルに当てはまらないケースや判定しきれないケースは `unknown` として扱い、安全側の半構造化出力を維持する。
    - 類型定義は理論先行ではなく、実サンプルで検証しながら更新可能なものとして運用する。

- [x] **出力契約の統一**
  - 決定事項:
    - 現時点の最終成果物は Markdown に統一する。
    - 再構成フェーズでは、構造保持と比較可能性を優先し、全モードで同一の出力契約を維持する。
    - 説明的な言い換えやチャンキング最適化のための文章化は、再構成とは別段階の将来検討事項として扱う。
    - したがって、現段階では Markdown 再構成を正本とし、その後段に必要に応じて説明化処理を追加できる構成を目指す。

- [x] **LLMのI/O設計**
  - 決定事項:
    - LLM の入力単位は原則としてテーブル単位とする。
    - ただし、テーブル単体を裸で渡すのではなく、シート名、近接見出し、シート上の元位置などの局所文脈を付与する。
    - LLM には Markdown 全文の生成ではなく、テーブル解釈に関する構造化結果を返させ、最終的な Markdown 組み立てはレンダリング層で行う。
    - 将来的に、複数テーブルが同一の意味ブロックを構成すると判断できる場合は、テーブル群単位の扱いを追加検討する。

- [x] **切り替え条件**
  - 決定事項:
    - モード切り替えの条件は自動判定しない。
    - 実行前に選択したモードを固定し、別モードへの切り替えが必要な場合は中間成果物から明示的に再実行する。

- [x] **信頼度と要確認マーカー**
  - 決定事項:
    - 信頼度と要確認マーカーは、中間表現には埋め込まず、レビュー用成果物で管理する。
    - 最終 Markdown には信頼度や要確認マーカーを埋め込まない。
    - 管理単位は原則としてテーブル単位とし、必要に応じて文書単位へ集約する。
    - 信頼度は `high` `medium` `low` の 3 段階、要確認は `review_required` の真偽値で表現する。
    - 要確認理由は自由文ではなく `reason_code` の列挙で保持する。
    - 要確認項目があっても Markdown 生成は継続し、レビュー用成果物で追跡できるようにする。

- [x] **評価方法**
  - 決定事項:
    - 評価方法は、受け入れ評価と比較評価の 2 段階で行う。
    - 受け入れ評価では、可視テキストの欠損がないこと、捏造がないこと、対応関係の取り違えがないこと、Markdown とレビュー用成果物が生成されることを確認する。
    - 比較評価では、同一の中間成果物に対して `既存実装`、`LLMなし`、`OpenAI API`、`ローカルLLM` の各方式を比較する。
    - 評価単位は原則としてテーブル単位とし、`input/excel` 配下の実サンプルを評価セットとして用いる。
    - 評価観点の優先順位は、(1) 情報保持、(2) 意味関係の正しさ、(3) 可読性、(4) Dify での扱いやすさ、とする。
    - 判定が難しいケースは `unknown` を許容し、安全側に倒れているかを確認する。

- [x] **運用モードの整理**
  - 決定事項:
    - 運用モードは `OpenAI API`、`ローカルLLM`、`LLMなし` の 3 モードで確定とする。
    - 各モードは同一の抽出結果に対して比較・再生成できることを前提とする。
    - 出力成果物およびログには、使用モード、モデル、プロンプト版などの追跡情報を残す。

- [x] **Word/Excel共通化の範囲**
  - 決定事項:
    - Word/Excel で共通化するのは、運用モード、中間表現の基本思想、Step3 の責務分離、LLM の I/O 契約、最終 Markdown 出力契約、レビュー用成果物、評価方針とする。
    - Step1（正規化）および Step2（抽出）は形式別実装を維持し、Word と Excel を無理に共通化しない。
    - LLM に渡す単位や構造化結果の契約は共通化するが、局所文脈の集め方や hint 生成は形式別に実装する。
    - したがって、共通化の対象は「再構成の枠組み」であり、「抽出ロジックそのもの」ではない。
    - Phase 1 は Excel を対象に検証し、成立した枠組みを Phase 2 で Word に適用する。
    - ただし Word については既存処理で動作確認とチェックが進んでいるため、既存の動きを大きく変えず、必要最小限の適用に留める。

- [x] **将来拡張の境界**
  - 決定事項:
    - 現段階のスコープは、Excel のテーブル再構成を中心に、`OpenAI API`、`ローカルLLM`、`LLMなし` の 3 モードで共通利用できる再構成基盤を整えるところまでとする。
    - 最終成果物は Markdown を正本とし、説明的な文章化は将来拡張とする。
    - チャンキングは Dify 側の仕組みに寄せる前提とし、事前チャンキング専用成果物の生成は現段階でも将来拡張でも対象外とする。
    - 複数テーブルの高度な意味ブロック化、図形や画像を含む統一解釈、Word の大幅な再設計は現段階の対象外とする。
    - ただし、将来的にテーブル群単位の解釈や説明化フェーズを追加できるよう、入出力契約と中間成果物の設計には拡張余地を持たせる。
    - Word への適用は既存動作を大きく変えない範囲での限定的拡張に留める。

- [x] **既存実装の段階的移行方針**
  - 決定事項:
    - 既存実装は、特に `LLMなし` モードにおけるベースラインとして活用する。
    - 現時点で動作確認できている品質を前提資産として扱い、全面的な作り直しは前提にしない。
    - 今後の改修は、責務分離を進めながら必要箇所を段階的に差し替える方針とする。
    - 新方式は既存実装との比較・検証を行いながら導入し、柔軟に併存・切り替えできる設計を目指す。

### 5.2 先に議論する優先テーマ

- [x] **LLMなしでの成功条件**
  - 参照: [5.1 の `LLM非依存の完走条件`](#51-議論テーマチェックリスト)
  - 優先理由: `LLMなし` モードをベースラインとして固定しないと、以降の品質基準と比較軸がぶれやすいため

- [x] **責務分担の原則**
  - 参照: [5.1 の `責務分担の原則`](#51-議論テーマチェックリスト)
  - 優先理由: 既存実装を大きく壊さず段階的に整理するための土台になるため

- [x] **LLMなしモードの出力仕様**
  - 参照: [5.1 の `LLMなしモードの出力仕様`](#51-議論テーマチェックリスト)
  - 優先理由: `to_markdown.py` の初期改修で、どこまで既存出力互換を維持するかの判断に直結するため

- [x] **LLMのI/O設計**
  - 参照: [5.1 の `LLMのI/O設計`](#51-議論テーマチェックリスト)
  - 優先理由: `ReconstructionUnit` と `TableInterpretationResult` の境界を決める起点になるため

### 5.3 現時点の結論サマリ

- 現段階のスコープは、Excel のテーブル再構成を中心に、`OpenAI API`、`ローカルLLM`、`LLMなし` の 3 モードで共通利用できる再構成基盤を整えるところまでとする
- モードの切り替えは自動ではなく明示的に行い、中間成果物を起点に Step3 を再実行できる設計とする
- `LLMなし` モードは既存実装をベースラインとして活用し、強い意味解釈を避けた半構造化 Markdown を標準出力とする
- 中間表現は各モード共通の source of truth とし、観測事実を欠損なく保持する。推定結果や解釈結果は混在させない
- LLM の入力単位は原則としてテーブル単位とし、シート名、近接見出し、シート上の位置などの局所文脈を付与する
- LLM には Markdown 全文ではなく、テーブル解釈のための構造化結果を返させ、最終 Markdown はレンダリング層で組み立てる
- 最終成果物は Markdown に統一し、信頼度と要確認情報はレビュー用成果物で管理する
- チャンキングは Dify 側に寄せる前提とし、事前チャンキング専用成果物の生成は対象外とする
- Word については既存の動きを大きく変えず、必要最小限の適用に留める

#### 5.3.1 最小理解ポイント

- 最低限覚える点は 5 つだけでよい
- `LLMなし` をベースラインにして、`OpenAI API` / `ローカルLLM` は上位モードとして扱う
- 自動切り替えはせず、中間成果物から明示的に再実行する
- LLM は Markdown 全文ではなく、表の解釈結果だけ返す
- 最終成果物は Markdown、レビュー情報は別成果物、Word は既存挙動優先で進める

### 5.4 実装検討への引き継ぎ項目

- `ReconstructionUnit`（テーブル単位 + 局所文脈）の入出力スキーマを定義する
- LLM が返す構造化結果のスキーマ（類型、confidence、reason_code、位置参照など）を定義する
- `LLMなし` モードの既存実装をどこまでそのまま活用し、どこから責務分離するかを切り分ける
- `input/excel` 配下の代表サンプルを、実装・評価で使う評価セットとして整理する
- レビュー用成果物の形式と出力粒度（テーブル単位/文書集約）を決める
- Excel で成立した枠組みを Word にどう最小適用するかを、既存挙動を維持する前提で整理する

### 5.5 実装タスク分解

以下は、実装時に着手順で進めるための具体タスクである。
結論に基づく実装方針を崩さないため、スキーマ定義 → 評価セット整理 → 既存実装の責務分離 → LLM モード追加 → 評価の順で進める。

#### 依存関係の見取り図

- `Task 1 -> Task 6 -> Task 10`
  - `to_markdown.py` の棚卸しを先に行い、`LLMなし` モードの責務分離を経て、共通 Renderer 契約へつなぐ
- `Task 2 -> Task 3 -> Task 9 -> Task 10`
  - `ReconstructionUnit` と `TableInterpretationResult` の契約を先に固め、その上で LLM アダプタと Renderer 契約を設計する
- `Task 4 -> Task 7`
  - 中間表現の最小拡張点が決まってから、中間成果物からの再実行経路を整理する
- `Task 5 -> Task 11`
  - 評価セット整理は比較評価フローの前提になる
- `Task 3 + Task 7 -> Task 8`
  - レビュー用成果物は解釈結果スキーマと再実行出力先設計の両方を前提にする
- `Task 6` から `Task 10` が固まってから `Task 12`
  - Word への最小適用範囲は、Excel 側の枠組みが見えてから判断する

- [x] **Task 1: 現行 Step3 の責務棚卸し**
  - 対象: `src/transform/to_markdown.py`
  - 内容: 現行ロジックを「分析」「解釈」「レンダリング」に分類し、どこまでを `LLMなし` モードでそのまま活用するかを明確にする
  - 完了条件: 主要関数ごとの責務一覧と、残す処理/切り出す処理の対応表が作成されている

#### Task 1 棚卸し表（たたき台）

対象ファイル: `src/transform/to_markdown.py`

| 関数 | 現在の主分類 | 主な役割 | 初期方針 |
|---|---|---|---|
| `_render_heading` | レンダリング | 見出し要素を Markdown 見出しへ変換 | そのまま維持 |
| `_render_paragraph` | レンダリング | 段落・箇条書きを Markdown 文字列へ変換 | そのまま維持 |
| `_fill_rowspan` | 分析 | rowspan を後続行へ展開し、表の観測事実を正規化 | Analyzer 側へ切り出し候補 |
| `_expand_row_to_positions` | 分析 | `col` / `colspan` を列位置配列へ展開 | Analyzer / Renderer 共通ユーティリティ候補 |
| `_is_empty_row` | 分析 | 空行判定 | 共通ユーティリティとして維持 |
| `_is_form_field_row` | 解釈 | 行がフォーム行かを判定 | `LLMなし` 解釈ロジック候補 |
| `_estimate_total_cols` | 分析 | テーブル全体の展開列数を推定 | Analyzer 側へ切り出し候補 |
| `_render_form_field_row` | レンダリング | フォーム行を `ラベル: 値` 形式へ変換 | Renderer 側に維持 |
| `_is_form_grid_table` | 解釈 | テーブル全体がフォーム型か判定 | `LLMなし` 解釈ロジック候補 |
| `_should_skip_as_header` | 解釈 | ヘッダー候補から除外すべき行か判定 | `LLMなし` 解釈ロジック候補 |
| `_find_header_row` | 解釈 | 先頭有効ヘッダー行を探索 | `LLMなし` 解釈ロジック候補 |
| `_build_column_labels` | 解釈 | ヘッダー行から列ラベルとデータ開始位置を構築 | `LLMなし` 解釈ロジック候補 |
| `_is_banner_row` | 解釈 | 行がバナー行か判定 | `LLMなし` 解釈ロジック候補 |
| `_is_section_header_row` | 解釈 | 行がセクション見出し行か判定 | `LLMなし` 解釈ロジック候補 |
| `_get_section_header_text` | レンダリング | セクション見出し行から表示テキストを抽出 | Renderer 側に維持 |
| `_render_form_grid` | レンダリング | フォーム型テーブル全体を出力 | Renderer 側に維持 |
| `_render_pre_header_rows` | 混在 | ヘッダー前行の判定と出力を同時に行う | 判定と描画の分離候補 |
| `_detect_active_columns` | 分析 | データ行の充填率からアクティブ列を検出 | Analyzer 側へ切り出し候補 |
| `_render_key_value_table` | 混在 | KV 型のキー/値判定結果をもとに出力 | レンダリング中心に整理し直す候補 |
| `_render_data_table` | 混在 | データテーブル型の行出力と一部再判定を担当 | レンダリング中心に整理し直す候補 |
| `_render_table_as_labeled_text` | 混在 | 表全体の型判定・分岐・出力を一括で担当 | 最優先の分離対象 |
| `_render_shape` | レンダリング | 図形情報をテキスト説明へ変換 | 今回は原則維持 |
| `transform_to_markdown` | 混在 | 要素種別ごとに各 renderer を呼び出す Step3 入口 | オーケストレーション層として整理 |
| `transform_file` | I/O | JSON 読み込み、Markdown 書き出し、ログ記録 | I/O 層として維持 |

#### Task 1 の優先分離対象

- `_render_table_as_labeled_text`
  - 表型判定、列ラベル構築、KV 判定、描画分岐が集中しているため、`analyze_table` / `interpret_table_no_llm` / `render_table` 相当へ分離する第一候補
- `_render_data_table`
  - セクション見出し後のヘッダー再検出など、解釈ロジックが残っているため、レンダリング専用に寄せる
- `_render_key_value_table`
  - `active_cols` が与えられた前提での描画に寄せ、列選択ロジックは外へ出す
- `_render_pre_header_rows`
  - 「何を出すかの判定」と「どう出すかの描画」が混在しているため、分離候補とする

#### Task 1 の初期方針メモ

- 初期段階では、`LLMなし` モードの出力互換を優先し、文字列フォーマット自体は大きく変えない
- まずは関数の責務を整理し、既存ロジックを別層へ移すだけで同等出力を維持できる形を目指す
- 図形系 (`_render_shape`) とファイル I/O (`transform_file`) は今回の主戦場ではないため、後回しとする

#### Task 1 残す処理 / 切り出す処理 対応表

| 処理のまとまり | 現行関数 | 当面の扱い | 切り出し先イメージ | 理由 |
|---|---|---|---|---|
| 見出し・段落の基本描画 | `_render_heading`, `_render_paragraph` | 残す | Renderer | 既存挙動が安定しており、今回の主戦場ではない |
| 図形の描画 | `_render_shape` | 残す | Renderer | Excel テーブル再構成の中心論点ではなく、Word 側影響も抑えたい |
| ファイル I/O | `transform_file` | 残す | I/O 層 | JSON 読み書きとログ出力は責務が明確 |
| 構造正規化 | `_fill_rowspan`, `_expand_row_to_positions`, `_estimate_total_cols` | 切り出す | Analyzer 共通ユーティリティ | 観測事実の整形であり、解釈や描画から独立できる |
| 行・表の判定 | `_is_empty_row`, `_is_form_field_row`, `_is_form_grid_table`, `_is_banner_row`, `_is_section_header_row`, `_should_skip_as_header`, `_find_header_row`, `_detect_active_columns` | 切り出す | `LLMなし` 解釈ロジック / Analyzer | 表型判定や行種別判定は解釈責務に寄せたい |
| 列ラベル構築 | `_build_column_labels` | 切り出す | `LLMなし` 解釈ロジック | 多段ヘッダー解釈を含み、レンダリングから独立させたい |
| ヘッダー前行の処理 | `_render_pre_header_rows` | 分離する | 判定は解釈層、文字列化は Renderer | 判定と描画が混在している代表例 |
| フォーム系の文字列化 | `_render_form_field_row`, `_render_form_grid`, `_get_section_header_text` | 残しつつ整理 | Renderer | 解釈結果を受けて描画する責務に寄せやすい |
| KV / データ表の出力 | `_render_key_value_table`, `_render_data_table` | 分離する | Renderer + 解釈結果入力 | 現状は一部再判定を含むため、描画専用に寄せ直す必要がある |
| 表全体の統括 | `_render_table_as_labeled_text` | 最優先で分離 | `analyze_table` / `interpret_table_no_llm` / `render_table` | 表型判定・分岐・描画呼び出しが集中しており、責務分離の中心になる |
| 文書全体の統括 | `transform_to_markdown` | 整理する | Step3 オーケストレーション層 | 要素ごとの呼び出し順は維持しつつ、テーブル処理の境界を明確にしたい |

#### Task 1 の完了イメージ

- `to_markdown.py` の主要処理について、「そのまま残す」「Analyzer 側へ切り出す」「`LLMなし` 解釈ロジックへ切り出す」「Renderer に残す」の 4 区分が定義されている
- `_render_table_as_labeled_text` を中心に、表処理の責務分離後の骨格イメージが文章で説明できる
- 既存出力互換を優先する範囲と、今後の拡張に備えて切り出す範囲が明文化されている

- [x] **Task 2: `ReconstructionUnit` スキーマ定義**
  - 対象: Excel テーブル単位 + 局所文脈
  - 内容: テーブル本体、シート名、近接見出し、シート上の元位置、hint 情報の持ち方を定義する
  - 完了条件: JSON 例を含むスキーマ案があり、`input/excel` の代表サンプルに適用できる

#### Task 2 スキーマ定義（たたき台）

`ReconstructionUnit` は、中間表現そのものを置き換えるものではなく、Step3 開始時に中間表現から生成する「テーブル単位 + 局所文脈」の入力オブジェクトとして定義する。

##### 基本方針

- `IntermediateDocument` / `ExtractedFileRecord` は引き続き source of truth とする
- `ReconstructionUnit` は、LLM モードおよび `LLMなし` モードの解釈処理に渡す共通入力単位とする
- 必須項目は観測事実のみで構成し、推定値は `hints` に分離する。`doc_role_guess` のような推定メタデータも `source` ではなく `hints` に置く
- テーブル本体のセル構造は、既存の `CellData[][]` にできるだけ近い形を維持する

##### 必須フィールド

| フィールド | 型 | 内容 |
|---|---|---|
| `schema_version` | `string` | `ReconstructionUnit` の版 |
| `unit_id` | `string` | テーブル単位の一意 ID |
| `source` | `object` | 元ファイルの追跡情報 |
| `sheet` | `object` | シート名とシート内位置情報 |
| `table` | `object` | 対象テーブルの事実情報 |

##### 任意フィールド

| フィールド | 型 | 内容 |
|---|---|---|
| `context` | `object` | 近接見出し・近接ラベル・補足メモなどの局所文脈 |
| `hints` | `object` | ヘッダー候補や類型候補などの推定情報 |

##### フィールド詳細

```json
{
  "schema_version": "1.0",
  "unit_id": "input/excel/change_history.xlsx::設計概要::table_0001",
  "source": {
    "source_path": "input/excel/change_history.xlsx",
    "source_ext": ".xlsx"
  },
  "sheet": {
    "name": "設計概要",
    "sheet_index": 0
  },
  "table": {
    "table_id": "table_0001",
    "caption": "",
    "bounds": {
      "sheet_row_start": 1,
      "sheet_col_start": 1,
      "sheet_row_end": 5,
      "sheet_col_end": 2
    },
    "has_merged_cells": false,
    "rows": [
      [
        {
          "text": "項目",
          "row": 0,
          "col": 0,
          "rowspan": 1,
          "colspan": 1,
          "is_header": true,
          "source_position": {
            "sheet_row": 1,
            "sheet_col": 1
          }
        },
        {
          "text": "内容",
          "row": 0,
          "col": 1,
          "rowspan": 1,
          "colspan": 1,
          "is_header": true,
          "source_position": {
            "sheet_row": 1,
            "sheet_col": 2
          }
        }
      ]
    ]
  },
  "context": {
    "nearby_headings": [],
    "nearby_labels": [],
    "sheet_notes": []
  },
  "hints": {
    "doc_role_guess": "data_sheet",
    "table_type_candidates": [
      "key_value",
      "data_table"
    ],
    "header_row_candidates": [0],
    "active_column_candidates": [0, 1]
  }
}
```

##### 設計上の注意

- `table.rows` は観測事実を保持する領域であり、セル本文を書き換えない
- `is_header` は抽出時点では事実ではなくヒントに近いため、解釈時には絶対視しない
- `bounds` と `source_position` により、シート上の元位置を追跡できるようにする
- `context` は局所文脈のみを持ち、シート全体を丸ごと抱え込まない
- `hints` は省略可能とし、初期段階では最小限の候補だけ持てればよい

##### 初期実装で最低限必要な項目

- `schema_version`
- `unit_id`
- `source.source_path`
- `source.source_ext`
- `sheet.name`
- `table.table_id`
- `table.bounds`
- `table.rows`

##### 後続タスクへの受け渡し

- Task 3 では、この `ReconstructionUnit` を入力として受け取る LLM 解釈結果スキーマを定義する
- Task 4 では、`source_position` や `bounds` を既存中間表現にどう持たせるかを整理する
- Task 9 では、このスキーマを前提に `OpenAI API` / `ローカルLLM` のアダプタを設計する

- [x] **Task 3: LLM 解釈結果スキーマ定義**
  - 対象: `table_type`、ヘッダー解釈、active_columns、confidence、reason_code など
  - 内容: LLM が返す構造化結果のフィールドと必須/任意を定義する
  - 完了条件: Markdown 全文ではなく構造化結果を返す契約が文書化されている

#### Task 3 スキーマ定義（たたき台）

LLM が返す出力は、Markdown 本文ではなく `TableInterpretationResult` と呼ぶ構造化結果とする。
このスキーマは、将来的に `LLMなし` 解釈ロジックの出力契約にも流用できる形を意識する。

##### 基本方針

- 入力は `ReconstructionUnit`
- 出力は Markdown 全文ではなく、Renderer に渡すための構造化結果
- セル本文の再出力や言い換えは行わず、可能な限り行/列/セル参照で結果を返す
- `confidence` と `reason_code` は LLM 自己評価として返してよいが、最終的なレビュー判定は別途レビュー用成果物で管理する
- 判定不能な場合でも `unknown` を返し、失敗ではなく安全側の出力へつなげる

##### 必須フィールド

| フィールド | 型 | 内容 |
|---|---|---|
| `schema_version` | `string` | `TableInterpretationResult` の版 |
| `unit_id` | `string` | 対応する `ReconstructionUnit.unit_id` |
| `table_type` | `string` | `key_value` / `data_table` / `form` / `unknown` |
| `render_strategy` | `string` | Renderer に渡す出力戦略。初期値は `key_value` / `data_table` / `form` / `safe_unknown` |

##### 任意フィールド

| フィールド | 型 | 内容 |
|---|---|---|
| `header_rows` | `array[int]` | データヘッダーとして扱う行インデックス |
| `meta_header_rows` | `array[int]` | 列の役割を示すだけで本文出力しない行インデックス |
| `data_start_row` | `int` | データ本体の開始行 |
| `column_labels` | `array[object]` | 列ラベルの確定結果 |
| `active_columns` | `array[int]` | 実際に意味を持つ列候補 |
| `key_columns` | `array[int]` | KV 型でキーとして扱う列 |
| `value_columns` | `array[int]` | KV 型で値として扱う列 |
| `row_annotations` | `array[object]` | 行ごとの役割注釈 |
| `section_rows` | `array[object]` | セクション境界行と見出し |
| `notes` | `array[string]` | 補足メモ |
| `self_assessment` | `object` | LLM 自己評価 |

##### フィールド詳細

```json
{
  "schema_version": "1.0",
  "unit_id": "input/excel/change_history.xlsx::設計概要::table_0001",
  "table_type": "key_value",
  "render_strategy": "key_value",
  "header_rows": [],
  "meta_header_rows": [0],
  "data_start_row": 1,
  "column_labels": [
    { "col": 0, "label": "項目" },
    { "col": 1, "label": "内容" }
  ],
  "active_columns": [0, 1],
  "key_columns": [0],
  "value_columns": [1],
  "row_annotations": [
    { "row": 0, "role": "meta_header" }
  ],
  "section_rows": [],
  "notes": [
    "2列のキーバリュー型として扱う",
    "先頭行は列の役割を示すメタヘッダーとみなす"
  ],
  "self_assessment": {
    "confidence": "medium",
    "reason_codes": [
      "kv_two_column_pattern",
      "meta_header_detected"
    ]
  }
}
```

##### 各フィールドの意味

- `table_type`
  - 表の解釈結果そのもの
- `render_strategy`
  - Renderer にどの出力経路を選ばせるか
  - 初期段階では `table_type` と同値でよいが、将来は `unknown -> safe_unknown` のように分けて扱える
- `header_rows`
  - 実データの列ラベルとして採用する行
- `meta_header_rows`
  - 役割説明や補助情報として扱い、データヘッダーには採用しない行
- `column_labels`
  - レンダリング時に使う確定ラベル
  - 多段ヘッダー時は LLM が最終ラベルを組み立てて返してよい
- `row_annotations`
  - `meta_header`, `banner`, `section_header`, `form_row`, `data_row`, `skip` などを想定する
- `self_assessment`
  - LLM 自己評価であり、レビュー用成果物にそのまま採用するとは限らない

##### 設計上の注意

- LLM はセル本文を書き換えず、どの行/列をどう解釈したかを返す
- 判定に迷う場合は `table_type = unknown`、`render_strategy = safe_unknown` を返せばよい
- `notes` は省略可能であり、初期段階では短い補足のみを許容する
- `self_assessment.confidence` は `high` / `medium` / `low` の 3 段階とし、最終的なレビュー判定とは分離する
- `reason_codes` は固定コードの列挙とし、自由文を前提にしない

##### 初期 `reason_codes` 一覧

- `kv_two_column_pattern`
  - 2 列 KV パターンが強く観測された
- `header_row_detected`
  - データヘッダー行を特定した
- `meta_header_detected`
  - メタヘッダー相当の行を検出した
- `form_grid_pattern`
  - フォーム型グリッドの特徴が強い
- `section_header_detected`
  - セクション見出し行を検出した
- `merged_cells_complex`
  - 結合セルが多く、解釈難度が高い
- `context_missing`
  - 局所文脈が不足している
- `unknown_table_type`
  - 類型を確定できず `unknown` に倒した
- `safe_unknown_required`
  - 安全側の出力戦略を選択した

##### 初期実装で最低限必要な項目

- `schema_version`
- `unit_id`
- `table_type`
- `render_strategy`

##### 後続タスクへの受け渡し

- Task 8 では、`self_assessment` の内容をどうレビュー用成果物へ反映するかを設計する
- Task 9 では、このスキーマを返すプロンプトとアダプタを設計する
- Task 10 では、`render_strategy` と `column_labels` などを受けて Markdown を生成する Renderer 契約を定義する

- [x] **Task 4: 中間表現の拡張点整理**
  - 対象: `src/models/intermediate.py`、Excel 抽出結果
  - 内容: シート上の元位置、観測事実と hint の分離など、最小限必要な拡張点を洗い出す
  - 完了条件: 既存中間表現のままで足りる項目と、追加が必要な項目が整理されている

#### Task 4 拡張点整理（たたき台）

Task 4 の目的は、中間表現を全面的に作り直すことではなく、`ReconstructionUnit` を安定して生成するために本当に必要な追加項目だけを特定することにある。

##### 基本方針

- 既存の `IntermediateDocument` / `DocumentElement` / `TableElement` / `CellData` は基本的に維持する
- `ReconstructionUnit` は中間表現から組み立てるものであり、中間表現にそのまま埋め込まない
- 追加は「観測事実の不足」を埋めるものに限定し、解釈結果やレビュー情報は追加しない
- 既存出力との互換を優先し、既存コードが広く参照している項目名は可能なら維持する

##### 現状の中間表現で足りているもの

| 項目 | 現状 | 判断 |
|---|---|---|
| セル本文 | `CellData.text` | そのままで足りる |
| 表内相対位置 | `CellData.row`, `CellData.col`, `rowspan`, `colspan` | そのままで足りる |
| 表の要素順序 | `DocumentElement.source_index` | そのままで足りる |
| 見出し・段落 | `HeadingElement`, `ParagraphElement` | 局所文脈の構築に利用可能 |
| ファイル追跡情報 | `FileMetadata.source_path`, `source_ext` など | そのままで足りる |
| 表の基本状態 | `TableElement.caption`, `has_merged_cells`, `confidence`, `fallback_reason` | 現段階では維持してよい |

##### 追加が必要なもの

| 追加項目 | 追加先候補 | 理由 | 優先度 |
|---|---|---|---|
| テーブルのシート上 bounds | `TableElement` | `ReconstructionUnit.table.bounds` を生成するために必要 | 必須 |
| シート上の元位置を導く情報 | `TableElement.bounds` から導出 | セルごとの絶対位置は bounds + 相対座標で計算できる | 必須 |
| 抽出時点の hint を事実と分離して扱う運用 | 実装規約 | `CellData.is_header` を事実として使わないため | 必須 |

##### 追加しないもの

| 追加しない項目 | 理由 |
|---|---|
| `ReconstructionUnit` 全体を中間表現に埋め込む | Step3 用の派生オブジェクトであり、source of truth と混ぜない |
| レビュー用の `confidence` / `review_required` / `reason_code` | レビュー用成果物で管理する方針のため |
| LLM 解釈結果そのもの | 中間表現は観測事実を保持する領域であり、解釈結果は別成果物とする |
| シート全体の文脈を丸ごと重複保持する項目 | `DocumentElement` の順序と見出し・段落から組み立て可能なため |
| `sheet_name` を各テーブルに重複保持する項目 | `ReconstructionUnit.sheet.name` で保持するため、`TableElement` へは初期段階では追加しない |

##### 最小拡張案

初期実装では、`TableElement` にシート上の矩形 bounds を保持できれば十分とする。

例:

```python
@dataclass
class TableElement:
    rows: list[list[CellData]]
    caption: str = ""
    has_merged_cells: bool = False
    confidence: Confidence = Confidence.HIGH
    fallback_reason: str = ""
    source_row_start: int | None = None
    source_col_start: int | None = None
    source_row_end: int | None = None
    source_col_end: int | None = None
```

この案では、各セルの絶対位置は以下で導出する:

- `sheet_row = source_row_start + cell.row`
- `sheet_col = source_col_start + cell.col`

##### `CellData.is_header` の扱い

- 既存コード互換のため、直ちにフィールド名は変更しない
- ただし意味上は「事実」ではなく「抽出時点の hint」として扱う
- `ReconstructionUnit` を組み立てる段階で、必要に応じて `hints.header_row_candidates` へ写像する
- 将来的に名称を `header_hint` 相当に整理する余地はあるが、初期段階では互換優先とする

##### Excel 抽出側で必要な対応

- `src/extractors/excel.py` で取得済みの `bounds`（連結領域の矩形）を `TableElement` に格納する
- `rows` の構造自体は大きく変えない
- シート名は Excel のシート情報から `ReconstructionUnit.sheet.name` にそのまま格納する
- 近接見出しや近接ラベルは既存の `HeadingElement` / `ParagraphElement` を利用し、抽出器側で重複保持しない

##### Task 4 の結論候補

- 既存中間表現は大枠維持する
- 初期段階の必須拡張は `TableElement` へのシート上 bounds 追加のみとする
- `CellData.is_header` は削除せず、意味上は hint として扱う運用へ変更する
- `ReconstructionUnit`、LLM 解釈結果、レビュー情報は中間表現へ埋め込まない

##### 後続タスクへの受け渡し

- Task 6 では、この前提で `LLMなし` モードの分析・解釈・レンダリング分離を設計する
- Task 7 では、中間成果物から `ReconstructionUnit` を組み立てる入出力経路を定義する
- Task 9 では、この最小拡張で LLM 入力に必要な情報が揃うかを確認する

- [x] **Task 5: 評価セット整理**
  - 対象: `input/excel`
  - 内容: 代表サンプルを類型ごとに選定し、期待する出力の観点を整理する
  - 完了条件: 少なくとも `change_history.xlsx`、`approval_request.xlsx` を含む評価セット一覧と評価観点メモがある

#### Task 5 評価セット整理（たたき台）

##### 基本方針

- 評価単位はファイル単位ではなく、原則としてテーブル単位とする
- 1ファイルに複数類型が含まれる場合は、同一ファイルから複数の評価ケースを切り出してよい
- 代表サンプルは「頻度」だけでなく「設計判断に効く難所」を含めて選定する
- コア評価セットと拡張評価セットに分け、初期実装ではコア評価セットを優先する
- 以下の候補ファイルは `input/excel` 配下に実在することを確認済みとする

##### コア評価セット候補

| ファイル | 主な観点 | 想定類型/論点 |
|---|---|---|
| `input/excel/change_history.xlsx` | 通常データ表、2列 KV、メタヘッダー | `data_table` / `key_value` |
| `input/excel/approval_request.xlsx` | フォーム型、帳票型、表と本文の混在 | `form` / mixed |
| `input/excel/excel_form_grid.xlsx` | 典型的なフォームグリッド | `form` |
| `input/excel/ledger_with_sections.xlsx` | セクション区切り、途中見出し | `data_table` + section |
| `input/excel/multiple_tables_sheet.xlsx` | 1シート複数表、無関係表の混在 | テーブル単位分割の妥当性 |
| `input/excel/merged_cells.xlsx` | 結合セル、rowspan/colspan 解釈 | merged cells |

##### 拡張評価セット候補

| ファイル | 主な観点 |
|---|---|
| `input/excel/invoice_print_layout.xlsx` | 印刷レイアウト寄り、`unknown` 安全側出力 |
| `input/excel/mixed_complex.xlsx` | 複合難所の組み合わせ |
| `input/excel/outline_and_filter.xlsx` | シート機能メモとの共存 |
| `input/excel/protected_master_validation.xlsx` | 保護・入力規則メモとの共存 |
| `input/excel/comments_and_annotations.xlsx` | コメント・注釈の保持 |
| `input/excel/formulas_and_formats.xlsx` | 数式・表示値の扱い |

##### 各評価ケースで記録する観点

- 対象ファイル / シート / テーブル ID
- 想定する出力戦略（`key_value` / `data_table` / `form` / `unknown`）
- 落としてはいけない情報
- 誤解釈してはいけない関係
- 可読性上の期待
- `unknown` を許容するかどうか

##### Task 5 の完了イメージ

- コア評価セットと拡張評価セットが一覧化されている
- 各ケースに対して期待する観点メモがある
- `既存実装` / `LLMなし` / `OpenAI API` / `ローカルLLM` の比較にそのまま使える

- [x] **Task 6: `LLMなし` モードの責務分離設計**
  - 対象: 既存 `to_markdown.py`
  - 内容: 既存の半構造化出力を維持したまま、分析・解釈・レンダリングの境界をどこに置くか設計する
  - 完了条件: 既存出力を大きく変えずに分離できるリファクタ方針が決まっている

#### Task 6 責務分離設計（たたき台）

##### 目標

- `LLMなし` モードの出力互換をなるべく維持しながら、Step3 を `Analyzer` / `Interpreter(no_llm)` / `Renderer` に分離する

##### 想定する内部フロー

1. `analyze_table(unit) -> TableAnalysis`
2. `interpret_table_no_llm(analysis) -> TableInterpretationResult`
3. `render_table(analysis, interpretation) -> str`

##### 初期段階で分離するもの

- rowspan 展開、列位置展開、列数推定
- 行種別判定、ヘッダー候補探索、アクティブ列検出
- 表型判定（`key_value` / `data_table` / `form` / `unknown`）
- Markdown 文字列組み立て

##### 初期段階で分離しすぎないもの

- 見出し・段落・図形の描画
- ファイル I/O
- 文字列フォーマットそのもの

##### Task 6 の完了イメージ

- `to_markdown.py` をどう分割するかの設計図がある
- `LLMなし` モードだけで完走できる構成が説明できる
- 既存出力との互換をどこで担保するかが明記されている

- [x] **Task 7: モード切り替えと再実行の入出力設計**
  - 対象: `OpenAI API` / `ローカルLLM` / `LLMなし`
  - 内容: 中間成果物から Step3 を再実行するための設定値、出力先、ログ項目を定義する
  - 完了条件: モードごとの出力ディレクトリと追跡項目の設計が決まっている

#### Task 7 モード切り替えと再実行の入出力設計（たたき台）

##### 基本方針

- モードは実行時に明示選択する
- Step2 の中間成果物を共通入力にして Step3 を再実行できるようにする
- 自動切り替えはしない

##### 想定設定値

- `llm_mode = openai | local | none`
- `model_name`
- `prompt_version`
- `review_output = true | false`

##### 出力先イメージ

- `intermediate/03_transformed/openai/...`
- `intermediate/03_transformed/local/...`
- `intermediate/03_transformed/none/...`
- `intermediate/03_review/openai/...`
- `intermediate/03_review/local/...`
- `intermediate/03_review/none/...`

##### 最低限残す追跡項目

- `source_path`
- `input_json_path`
- `llm_mode`
- `model_name`
- `prompt_version`
- `output_markdown_path`
- `output_review_path`
- `timestamp`

##### Task 7 の完了イメージ

- モードごとの入出力先が定義されている
- 同一中間成果物から別モード再実行する手順が説明できる
- ログに何を残すかが決まっている

- [x] **Task 8: レビュー用成果物の設計**
  - 対象: テーブル単位のレビュー記録
  - 内容: `confidence`、`review_required`、`reason_code` をどの形式で出力するか決める
  - 完了条件: レビュー用成果物のフォーマット案とサンプル出力例がある

#### Task 8 レビュー用成果物の設計（たたき台）

##### 基本方針

- レビュー用成果物は Markdown 本文とは別に出力する
- 主目的は「危ない箇所を後から見つけられること」
- 初期段階では機械可読性を優先し、主形式は JSON とする

##### 1テーブル分の記録イメージ

```json
{
  "unit_id": "input/excel/change_history.xlsx::設計概要::table_0001",
  "source_path": "input/excel/change_history.xlsx",
  "sheet_name": "設計概要",
  "table_id": "table_0001",
  "llm_mode": "openai",
  "render_strategy": "key_value",
  "confidence": "medium",
  "review_required": true,
  "reason_codes": [
    "meta_header_detected"
  ]
}
```

##### 形式案

- 主形式: JSONL（テーブル単位 1 レコード）
- 任意補助: 文書単位集約 JSON

##### Task 8 の完了イメージ

- レビュー用成果物の主形式が決まっている
- テーブル単位記録と文書集約の関係が整理されている
- `confidence` / `review_required` / `reason_codes` の出力例がある

- [x] **Task 9: LLM プロンプト/アダプタ設計**
  - 対象: `OpenAI API` モード、`ローカルLLM` モード
  - 内容: `ReconstructionUnit` を入力にし、構造化結果を返すプロンプトとアダプタ境界を定義する
  - 完了条件: モデル非依存の I/O 契約と、モードごとの差し替え点が整理されている

#### Task 9 LLM プロンプト/アダプタ設計（たたき台）

##### 基本方針

- モデル非依存の I/O 契約を先に固定する
- `OpenAI API` と `ローカルLLM` の差分は、呼び出し方法とモデル名に閉じ込める
- LLM には JSON だけを返させる
- ただし `ローカルLLM` はモデルによって JSON 出力の安定性が揺れるため、アダプタ側でスキーマ検証と必要最小限の再試行/補正方針を持てるようにする

##### プロンプトで守らせること

- セル本文を書き換えない
- 行/列/セル参照で判断する
- 判定不能なら `unknown` を返す
- 出力は `TableInterpretationResult` の JSON のみ

##### アダプタ境界

- 入力: `ReconstructionUnit`
- 出力: `TableInterpretationResult`
- 失敗時: 明示的にエラーを返し、自動で別モードへ切り替えない

##### Task 9 の完了イメージ

- プロンプトの役割とアダプタの役割が分離されている
- `OpenAI API` / `ローカルLLM` の差し替え点が明文化されている
- JSON 出力を前提にした呼び出し仕様が決まっている

- [x] **Task 10: レンダリング層の共通契約定義**
  - 対象: Markdown 出力
  - 内容: `LLMなし` / `OpenAI API` / `ローカルLLM` のいずれでも同じ出力契約になるよう、レンダリング入力と責務を定義する
  - 完了条件: 「構造化結果 -> Markdown」の共通インターフェースが決まっている

#### Task 10 レンダリング層の共通契約定義（たたき台）

##### 基本方針

- Renderer はモード非依存とする
- 入力差分は `TableInterpretationResult` に閉じ込める
- 出力は常に Markdown

##### 想定インターフェース

- 入力:
  - `ReconstructionUnit`
  - `TableInterpretationResult`
- 出力:
  - `markdown_text`

##### レンダリング戦略

- `key_value`
- `data_table`
- `form`
- `safe_unknown`

##### Task 10 の完了イメージ

- 同じ Renderer で 3 モードを処理できる
- `render_strategy` ごとの出力責務が整理されている
- 既存 `LLMなし` 出力互換をどこで担保するかが書かれている

- [ ] **Task 11: 比較評価フロー設計**
  - 対象: 既存実装、新 `LLMなし`、`OpenAI API`、`ローカルLLM`
  - 内容: 同一サンプルに対して各方式を比較する評価手順を決める
  - 完了条件: 受け入れ評価と比較評価の実施手順、および確認観点が手順化されている

#### Task 11 比較評価フロー設計（たたき台）

##### 比較対象

- 既存実装
- 新 `LLMなし`
- `OpenAI API`
- `ローカルLLM`

##### 評価手順イメージ

1. 同一の中間成果物 JSON を用意する
2. 各モードで Step3 を実行する
3. Markdown とレビュー用成果物を収集する
4. テーブル単位で比較する
5. 受け入れ条件を満たすか確認する

##### 記録する観点

- 情報保持
- 意味関係の正しさ
- 可読性
- `unknown` の妥当性
- Dify 取り込み上の問題の有無

##### Task 11 の完了イメージ

- 比較の実施手順が再現可能な形で書かれている
- 評価観点ごとに何を見るかが決まっている
- 代表サンプルに対する比較表のひな形がある

- [ ] **Task 12: Word 最小適用方針の整理**
  - 対象: Word 系既存処理
  - 内容: Excel で成立した枠組みのうち、Word に適用してよい部分と、既存挙動を維持すべき部分を切り分ける
  - 完了条件: Word にはどこまで適用し、どこは適用しないかの境界が明文化されている

#### Task 12 Word 最小適用方針の整理（たたき台）

##### 基本方針

- Word は既存挙動を大きく変えない
- Excel で成立した「契約」と「枠組み」だけを必要最小限で適用する

##### 適用してよいもの

- モード選択の考え方
- レビュー用成果物の考え方
- Renderer 契約
- LLM 入出力契約の考え方

##### 直ちには適用しないもの

- Word 抽出ロジックの大幅変更
- 図形・画像の扱いの再設計
- 既存 Word 出力フォーマットの大幅変更

##### Task 12 の完了イメージ

- Word に流用する枠組みと、流用しない既存挙動が切り分けられている
- Excel 優先で進めても Word 側に不要な破壊が起きないことが確認できる

### 5.6 想定修正ファイルの分類

実装着手時の見通しを持つため、想定修正ファイルを以下の 3 区分で整理する。

#### 変更必須

- `src/transform/to_markdown.py`
  - Step3 の主戦場。責務分離、`LLMなし` モード維持、レンダリング契約の整理の中心になる
- `src/extractors/excel.py`
  - テーブル単位 + 局所文脈 + シート上の元位置を扱うため、抽出情報の拡張が必要になる可能性が高い
- `src/models/intermediate.py`
  - 中間表現に追加すべき観測事実、hint、位置情報の持ち方を整理する必要がある

#### 変更の可能性あり

- `src/models/metadata.py`
  - モード情報、レビュー用成果物との連携、追跡項目の整理で見直す可能性がある
- `src/llm/base.py`
  - テーブル解釈の構造化結果を返す共通 I/O 契約の定義で変更候補になる
- `src/llm/openai_backend.py`
  - `OpenAI API` モード用のアダプタ調整が必要になる可能性がある
- `src/llm/local_backend.py`
  - `ローカルLLM` モード用のアダプタ調整が必要になる可能性がある
- `src/llm/noop_backend.py`
  - `LLMなし` モードの明示的な位置づけに合わせて役割整理が必要になる可能性がある
- `src/config.py`
  - モード選択、出力先、レビュー用成果物設定などの追加で変更候補になる
- `src/main.py`
  - 実行モードの明示選択や Step3 再実行導線の追加で変更候補になる
- `src/pipeline/folder_processor.py`
  - モード別出力やレビュー用成果物の出力制御で変更候補になる
- `src/pipeline/splitter.py`
  - 最終出力物の扱いに影響が出る場合に確認対象となる
- `src/extractors/registry.py`
  - モード依存ではないが、実行経路整理の影響で確認対象になる

#### なるべく触らない方針

- `src/extractors/word.py`
  - 既存動作の確認が進んでいるため、必要最小限の適用に留める
- `src/pipeline/normalizer.py`
  - 正規化は今回の主戦場ではなく、既存挙動を維持する前提とする
- `src/__init__.py`、`src/transform/__init__.py`、`src/models/__init__.py`、`src/llm/__init__.py`、`src/pipeline/__init__.py`、`src/extractors/__init__.py`
  - 構成上の補助ファイルであり、原則として今回の設計変更の中心にはしない

### 5.7 最初に触るファイルと変更箇所

既存影響を抑えながら実装を進めるため、最初に触るファイルと変更箇所を以下に固定する。

#### 第1優先

- `src/transform/to_markdown.py`
  - 変更箇所:
    - `_render_table_as_labeled_text`
    - `_render_data_table`
    - `_render_key_value_table`
    - `_render_pre_header_rows`
  - 目的:
    - 表処理の責務を `Analyzer` / `Interpreter(no_llm)` / `Renderer` に分離する
    - 既存の `LLMなし` 出力をできるだけ維持したまま整理する

#### 第2優先

- `src/models/intermediate.py`
  - 変更箇所:
    - `TableElement`
  - 目的:
    - シート上 bounds を保持できるようにする
    - 既存の `CellData` / `TableElement` の大枠は維持する

- `src/extractors/excel.py`
  - 変更箇所:
    - `_extract_region_table`
    - `extract_xlsx`
  - 目的:
    - 取得済みの連結領域 `bounds` を `TableElement` に格納する
    - `ReconstructionUnit` 生成に必要な元位置情報を中間表現へ渡す

#### 第3優先

- `src/llm/base.py`
  - 変更箇所:
    - LLM 入出力インターフェース
  - 目的:
    - `ReconstructionUnit -> TableInterpretationResult` の契約を定義する

- `src/llm/openai_backend.py`
- `src/llm/local_backend.py`
- `src/llm/noop_backend.py`
  - 変更箇所:
    - 各バックエンドの呼び出し境界
  - 目的:
    - 同じ I/O 契約で 3 モードを差し替えられるようにする

#### 後から着手

- `src/config.py`
- `src/main.py`
- `src/pipeline/folder_processor.py`
- `src/models/metadata.py`
  - 目的:
    - モード選択、出力先、レビュー用成果物、ログ項目をつなぐ

#### 初期段階では触らない

- `src/extractors/word.py`
  - 既存挙動を大きく変えない方針のため、Excel 側の枠組みが固まるまで原則保留とする

### 5.8 `to_markdown.py` の初期変更方針

最初のコーディング対象は [`src/transform/to_markdown.py`](/c:/Users/マサフミ/Downloads/files/src/transform/to_markdown.py) とする。
ここでは既存出力を大きく変えないことを優先し、内部責務の整理から着手する。

#### 目的

- `LLMなし` モードの出力互換を維持しながら、表処理を `Analyzer` / `Interpreter(no_llm)` / `Renderer` に分離しやすい形へ整える
- 後から `OpenAI API` / `ローカルLLM` を挿し込める境界を作る

#### 初期変更の原則

- Markdown の見た目はできるだけ変えない
- 関数分割を先に行い、ロジックの意味を変える変更は後回しにする
- 見出し・段落・図形・ファイル I/O は初期段階ではなるべく触らない
- 最優先はテーブル処理の境界整理であり、`_render_table_as_labeled_text` を中心に分解する

#### 変更順のたたき台

1. **共通ユーティリティの独立**
   - 対象:
     - `_fill_rowspan`
     - `_expand_row_to_positions`
     - `_estimate_total_cols`
     - `_is_empty_row`
   - 目的:
     - 表構造の正規化と位置展開を、描画ロジックから切り離す

2. **`LLMなし` 解釈ロジックのまとまり化**
   - 対象:
     - `_is_form_field_row`
     - `_is_form_grid_table`
     - `_is_banner_row`
     - `_is_section_header_row`
     - `_should_skip_as_header`
     - `_find_header_row`
     - `_build_column_labels`
     - `_detect_active_columns`
   - 目的:
     - 「どう描画するか」の前に「どう解釈するか」をまとめる

3. **Renderer の整理**
   - 対象:
     - `_render_form_field_row`
     - `_render_form_grid`
     - `_render_pre_header_rows`
     - `_render_key_value_table`
     - `_render_data_table`
   - 目的:
     - 入力として解釈結果を受ける前提に寄せ、内部再判定を減らす

4. **表全体の統括関数の分解**
   - 対象:
     - `_render_table_as_labeled_text`
   - 目的:
     - 1関数に集中している「分析 -> 解釈 -> 分岐 -> 描画」を分離する
   - 目標イメージ:
     - `analyze_table(...)`
     - `interpret_table_no_llm(...)`
     - `render_table(...)`

5. **Step3 入口の整理**
   - 対象:
     - `transform_to_markdown`
   - 目的:
     - 文書全体のオーケストレーションに責務を限定し、表処理は独立した流れにする

#### 初期段階で維持する既存関数

- `_render_heading`
- `_render_paragraph`
- `_render_shape`
- `transform_file`

#### 初期段階で最も注意する点

- `_render_data_table` と `_render_key_value_table` は描画関数に見えるが、内部に解釈ロジックが残っている
- `_render_pre_header_rows` は見た目以上に判定ロジックを持っている
- `_build_column_labels` は多段ヘッダーの解釈を含むため、単純なユーティリティではない
- `_render_table_as_labeled_text` を一気に置き換えるのではなく、内部処理を外へ逃がす形で段階的に小さくする

#### この段階でまだやらないこと

- LLM 呼び出しの実装
- プロンプト最適化
- Renderer の出力形式変更
- Word 用の処理変更

#### 完了イメージ

- `to_markdown.py` を見たときに、テーブル処理が「構造正規化」「解釈」「描画」のどこにあるか追える
- `LLMなし` モードだけで既存と大きく変わらない出力が出せる
- 次の段階で `ReconstructionUnit` / `TableInterpretationResult` を挟み込む場所が明確になっている

### 5.9 実装着手順（初期スライス）

実装は一気に進めず、既存出力への影響を最小化するため、以下のスライスで順に進める。

#### スライス A: 中間表現へシート上 bounds を追加する

- 対象ファイル:
  - `src/models/intermediate.py`
  - `src/extractors/excel.py`
- 目的:
  - `TableElement` にシート上の元位置を保持できるようにし、`ReconstructionUnit.table.bounds` の元情報を確保する
- 具体変更:
  - `TableElement` に以下を追加する
    - `source_row_start`
    - `source_col_start`
    - `source_row_end`
    - `source_col_end`
  - `IntermediateDocument.add_table(...)` に同名の任意引数を追加する
  - `excel.py` の `intermediate.add_table(...)` 呼び出しで `bounds` の値をそのまま渡す
- この段階で変えないもの:
  - `CellData` の構造
  - `_extract_region_table(...)` の戻り値
  - `to_markdown.py` の出力ロジック
- 完了条件:
  - Step2 JSON にテーブル bounds が追加される
  - Step3 の Markdown 出力は既存と変わらない

#### スライス B: `to_markdown.py` の共通ユーティリティを整理する

- 対象ファイル:
  - `src/transform/to_markdown.py`
- 目的:
  - 描画と独立できる表構造処理を先に分離し、後続の責務分離をやりやすくする
- 対象関数:
  - `_fill_rowspan`
  - `_expand_row_to_positions`
  - `_estimate_total_cols`
  - `_is_empty_row`
- この段階で変えないもの:
  - 文字列フォーマット
  - テーブル型判定の意味
  - `transform_to_markdown(...)` の呼び出し構造
- 完了条件:
  - 上記関数群が「構造正規化 / 位置展開」のまとまりとして追える
  - 既存サンプル Markdown の差分が出ない

#### スライス C: `LLMなし` 解釈ロジックの入口を作る

- 対象ファイル:
  - `src/transform/to_markdown.py`
- 目的:
  - `LLMなし` モードでも `TableInterpretationResult` 相当の出力契約に寄せていく
- 対象関数:
  - `_is_form_field_row`
  - `_is_form_grid_table`
  - `_should_skip_as_header`
  - `_find_header_row`
  - `_build_column_labels`
  - `_is_banner_row`
  - `_is_section_header_row`
  - `_detect_active_columns`
- 目標イメージ:
  - `analyze_table(...)`
  - `interpret_table_no_llm(...)`
  - `render_table(...)`
- この段階で変えないもの:
  - `OpenAI API` / `ローカルLLM` の呼び出し
  - 出力形式の変更
- 完了条件:
  - `_render_table_as_labeled_text(...)` の中で判定処理の責務が小さくなっている
  - `LLMなし` だけで既存と大きく変わらない出力が出る

#### スライス D: LLM 入出力境界を追加する

- 対象ファイル:
  - `src/llm/base.py`
  - `src/llm/openai_backend.py`
  - `src/llm/local_backend.py`
  - `src/llm/noop_backend.py`
  - `src/config.py`
  - `src/main.py`
- 目的:
  - `ReconstructionUnit -> TableInterpretationResult` の境界を追加し、3 モードを同一契約で扱えるようにする
- 実装方針:
  - 既存の `LLMBackend.generate(prompt, system)` はすぐに全面置換せず、テーブル解釈用の薄いアダプタ層を追加する方向で進める
  - `ローカルLLM` は JSON 安定性に揺れがある前提で、スキーマ検証と最小限の再試行をアダプタ側で吸収する
- 完了条件:
  - モード選択ごとの差し替え点が `llm/*` と設定層に閉じている
  - `to_markdown.py` 側へモデル依存分岐を持ち込まない

#### スライス E: レビュー用成果物を出力する

- 対象ファイル:
  - `src/pipeline/folder_processor.py`
  - `src/models/metadata.py`
  - 必要に応じて `src/config.py`
- 目的:
  - Markdown 本文と別に、テーブル単位レビュー情報を JSONL で出力できるようにする
- 完了条件:
  - `confidence` / `review_required` / `reason_codes` をテーブル単位で記録できる
  - 出力先がモード別に整理されている

#### 最初の実装開始点

- 最初のコード変更は `スライス A` から入る
- 理由:
  - 出力 Markdown に影響しない
  - 既存実装の安全性を保ったまま、後続タスクの前提データだけを追加できる
  - `intermediate.py` と `excel.py` の変更で閉じやすい

#### 初回確認に使うサンプル

- `input/excel/change_history.xlsx`
- `input/excel/approval_request.xlsx`

#### 初回確認観点

- Step2 JSON に bounds が追加されている
- Step3 の Markdown が既存と大きく変わらない
- `change_history` と `approval_request` の代表テーブルで欠損や崩れが増えていない

---

## 6. 参考情報

### 6.1 関連ファイル

| ファイル | 役割 |
|---|---|
| `src/extractors/excel.py` | Excel 構造抽出（Step2） |
| `src/extractors/word.py` | Word 構造抽出（Step2） |
| `src/models/intermediate.py` | 中間表現データモデル（CellData, TableElement 等） |
| `src/models/metadata.py` | 追跡情報とメタデータの保持 |
| `src/llm/base.py` | LLM バックエンド共通 I/O 契約 |
| `src/llm/openai_backend.py` | `OpenAI API` モードのバックエンド |
| `src/llm/local_backend.py` | `ローカルLLM` モードのバックエンド |
| `src/llm/noop_backend.py` | `LLMなし` モード相当のバックエンド整理候補 |
| `src/config.py` | モード選択、出力先、設定値の起点 |
| `src/main.py` | 実行経路の入口 |
| `src/pipeline/folder_processor.py` | モード別出力とレビュー用成果物の制御候補 |
| `src/transform/to_markdown.py` | 現行の Markdown 変換（Step3、置き換え対象） |
| `docs/Task.md` §5.4 | LLM 活用の設計原則 |
| `docs/Task.md` §12 | 参考リンク（RAG x FINDY スライド） |

### 6.2 参考資料

- RAG x FINDY: https://speakerdeck.com/harumiweb/rag-findy
- 参照スライド: https://speakerdeck.com/harumiweb/rag-findy?slide=14
- 取り込んだ考え方: 「Office 文書をそのまま Markdown やプレーンテキストに潰すのではなく、先に構造を取ってから再構成する」
