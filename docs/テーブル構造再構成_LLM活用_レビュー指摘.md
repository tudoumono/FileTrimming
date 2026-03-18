# テーブル構造再構成 LLM活用 — レビュー指摘事項

レビュー対象:
- `docs/テーブル構造再構成_LLM活用検討.md`
- `src/llm/` 配下の実装一式
- `src/transform/to_markdown.py` の LLM 統合部分
- `src/config.py`
- `tests/test_transform.py`, `tests/test_pipeline.py`

レビュー日: 2026-03-18

---

## 全体評価

設計ドキュメントは非常に丁寧に作り込まれており、問題の発見→議論→方針決定→実装タスク分解の流れが明確に追跡できる。実装もドキュメントの方針に忠実に従っており、整合性は高い。

---

## 良い点

1. **設計思想が一貫している**: 「構造を取ってから再構成する」という原則が全レイヤーに浸透している
2. **安全側への倒し方が明確**: confidence が low なら fallback、LLM エラーでも fallback、パターン不一致でも fallback — 多重のセーフティネットがある
3. **3モード設計が実用的**: noop/openai/local の切り替えが `create_backend()` のファクトリ一箇所に集約されており、差し替えが容易
4. **observation_only モード**: LLM 結果をログに残しつつ既存出力を維持できる — 導入リスクを最小化する良い設計
5. **中間表現を source of truth として維持**: 推定値と観測事実を分離する方針が `hints` フィールドで実現されている
6. **`_sanitize_markdown_lines` / `_sanitize_summary_labels`**: LLM 出力のバリデーションが丁寧で、元テキストとの突合チェックまで行っている

---

## 指摘事項

### [高] RV-001: ドキュメントと実装のスキーマ乖離

- 対象: `docs/テーブル構造再構成_LLM活用検討.md` Task 2, Task 3 / `src/llm/base.py`

ドキュメント (Task 2) の `ReconstructionUnit` は `source` / `sheet` / `table` のネスト構造を想定しているが、実装の `src/llm/base.py:14-30` ではフラット構造（`source_path`, `source_ext`, `sheet_name`, `rows` が直接フィールド）になっている。

ドキュメント Task 3 の `TableInterpretationResult` にも `meta_header_rows`, `key_columns`, `value_columns`, `row_annotations`, `section_rows` 等のフィールドが定義されているが、実装には存在しない。

**推奨**: ドキュメント側を「実装に合わせた最終版」に更新するか、「たたき台」と「採用版」を明示的に区別する記述を追加する。現状だと読者が混乱する。

---

### [高] RV-002: OpenAI / Local バックエンドのコード重複

- 対象: `src/llm/openai_backend.py`, `src/llm/local_backend.py`

`generate()` / `interpret_table()` メソッドが完全に同一。差分は `__init__` のみ。

**推奨**: 共通基底クラス `OpenAICompatibleBackend` を作り、`__init__` だけをサブクラスで実装する形にリファクタする。現在 90 行 x 2 ファイルで、修正漏れのリスクがある。

---

### [中] RV-003: JSON パースのエラーハンドリング

- 対象: `src/llm/table_interpretation.py:107`, `src/transform/to_markdown.py:1493-1510`

`parse_table_interpretation_response` 内の `json.loads()` が `JSONDecodeError` を catch しておらず、呼び出し側の `to_markdown.py` で `except Exception` で一括 catch している。これ自体は動作するが、LLM が不正 JSON を返した場合のエラーメッセージが `json.decoder.JSONDecodeError` の生テキストになり、デバッグ時に「プロンプトの問題か、パースの問題か」の切り分けが難しい。

**推奨**: `parse_table_interpretation_response` 内で `JSONDecodeError` を catch し、元の応答テキスト（先頭100文字程度）を含んだ専用例外に変換する。

---

### [中] RV-004: `_extract_json_text` の不完全 JSON 抽出リスク

- 対象: `src/llm/table_interpretation.py:76-79`

`{` から `}` の最後までを切り出すロジックで、LLM が JSON の後にテキストを付加し、そのテキスト内に `}` があると壊れる可能性がある。現状は `rfind` で最後の `}` を取っているので一般ケースでは動くが、ローカル LLM（出力安定性が低い）を使う場合に問題化するリスクがある。

**推奨**: JSON パースを試みて失敗したら `rfind` にフォールバックする2段階方式を検討。

---

### [中] RV-005: `httpx.Client` のリソースリーク

- 対象: `src/llm/http_client.py`

`build_http_client()` で生成した `httpx.Client` が明示的に `close()` されていない。OpenAI SDK が内部でコネクションプール管理しているはずだが、プロセスのライフタイムが長い場合（並列ワーカー等）にコネクションが溜まる可能性がある。

**推奨**: `PipelineConfig` のライフサイクルに合わせてクリーンアップするか、少なくとも `atexit` フックで close する仕組みを検討。

---

### [低] RV-006: 元々の動機となった問題が LLM パスでカバーされていない

- 対象: `src/transform/to_markdown.py:564-585`

`_should_request_llm_interpretation` の判定条件が「結合セルがあり、かつ小さなフォーム型に見える」テーブルに限定されている。ドキュメントの元々の動機であった「2列 KV テーブルの縦横誤判定」(SS2.1) は、結合セルがないため LLM に送られない。

**推奨**: これが意図的な段階的導入であれば、ドキュメント内で明記する（または今後の Phase として記録する）。

---

### [低] RV-007: Task 11, 12 が未完了

- 対象: `docs/テーブル構造再構成_LLM活用検討.md` SS5.5

Task 11（比較評価フロー設計）と Task 12（Word 最小適用方針）が `[ ]` のまま。他は全て `[x]`。

**推奨**: 完了状態の更新か、残タスクとしての扱いを明確にする。

---

### [低] RV-008: プロンプトのバージョン管理

- 対象: `src/llm/table_interpretation.py`, `docs/テーブル構造再構成_LLM活用検討.md` SS5.7

ドキュメント SS5.7 で「出力成果物にプロンプト版を残す」としているが、実装の `TABLE_INTERPRETATION_SYSTEM_PROMPT` にはバージョン情報が含まれていない。

**推奨**: プロンプト定数にバージョン文字列を持たせ、レビュー用成果物に含める。

---

## まとめ

| 重要度 | 件数 | 概要 |
|--------|------|------|
| 高     | 2    | スキーマ乖離の解消、バックエンドのコード重複排除 |
| 中     | 3    | JSON エラーハンドリング、JSON 抽出ロジック、httpx リソースリーク |
| 低     | 3    | LLM パスのカバー範囲、未完了タスク、プロンプトバージョン管理 |

主な改善点は **ドキュメントと実装のスキーマ乖離の解消** と **バックエンドのコード重複排除**。機能面では、元々の動機（2列 KV 問題）が現在の LLM パスで実際にカバーされるかの確認が重要。
