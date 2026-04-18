# CLAUDE.md (monorepo root)

## Worktree 起動時チェックリスト

この CLAUDE.md を読んだら、以下を実行してロールを確定せよ:

1. `pwd` → パスから worktree ロールを判定
   - `.claude/worktrees/pptx` または `ooxml-pptx` → PPTX session（packages/pptx のみ編集可）
2. `packages/pptx/CLAUDE.md` を必ず読む（PPTX 固有の詳細ルール）
3. 他 package のファイルは読み取り OK、編集は禁止

## プロジェクト概要

OOXML (pptx/docx/xlsx) をブラウザ Canvas に描画するライブラリ群。
Rust/WASM parser + TypeScript Canvas renderer 構成。

## ディレクトリ

- `packages/core/` — 共有レンダリングプリミティブ + 共有型
- `packages/pptx/` — PPTX 固有（Session A 所有）
- `packages/docx/` — DOCX 固有（Session B、未作成）
- `packages/xlsx/` — XLSX 固有（Session C、未作成）

## 自律作業の原則

- AM1時〜AM9時はユーザー確認不要。破壊的操作以外はすべて自律的に進めること。
- 確認なしで進めてよい作業: コード修正・WASM ビルド・テスト実行・commit/push・Python/npm スクリプト実行。
- `git push` は `http.postBuffer 524288000` を設定してから実行。
- 参照画像（`packages/*/tests/visual/references/`）はユーザー指示のみ更新。絶対に自動更新しない。
- pptx ファイルは git にコミットしない。

## WASM ビルド手順

```bash
cd packages/pptx/parser && wasm-pack build --target web
cp pkg/pptx_parser_bg.wasm pkg/pptx_parser.js ../src/wasm/
```

## テスト実行

```bash
npx playwright test --reporter=list
# または
pnpm vrt
```
