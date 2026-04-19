# CLAUDE.md (monorepo root)

## Worktree 起動時チェックリスト

この CLAUDE.md を読んだら、以下を実行してロールを確定せよ:

1. `pwd` → パスから worktree ロールを判定
   - `.claude/worktrees/pptx` または `ooxml-pptx` → PPTX session（packages/pptx のみ編集可）
2. 各パッケージの `CLAUDE.md` を必ず読む（パッケージ固有の詳細ルール）
3. 他 package のファイルは読み取り OK、編集は禁止

## プロジェクト概要

OOXML (pptx/docx/xlsx) をブラウザ Canvas に描画するライブラリ群。
Rust/WASM parser + TypeScript Canvas renderer 構成。

## ディレクトリ

- `packages/core/` — 共有レンダリングプリミティブ + 共有型
- `packages/pptx/` — PPTX 固有（Session A 所有）
- `packages/docx/` — DOCX 固有（Session B 所有）
- `packages/xlsx/` — XLSX 固有（Session C 所有）

## Git ワークフロー

**複数セッションが並列で作業するため、main への直接 push は禁止。**

- 作業は必ず feature branch で行う（例: `feature/xlsx-xxx`、`feature/pptx-xxx`）
- `git push origin <branch>` して PR を作成し、main へマージする
- `git push origin main` は絶対に行わない
- `git push` 前に `git config http.postBuffer 524288000` を設定すること

## 自律作業の原則

- AM1時〜AM9時はユーザー確認不要。破壊的操作以外はすべて自律的に進めること。
- 確認なしで進めてよい作業: コード修正・WASM ビルド・テスト実行・commit/push（feature branch のみ）・Python/npm スクリプト実行。
- 参照画像（`packages/*/tests/visual/references/`）はユーザー指示のみ更新。絶対に自動更新しない。
- pptx/xlsx/docx ファイルは git にコミットしない。

## WASM ビルド手順

```bash
# パッケージ別
cd packages/pptx/parser  && wasm-pack build --target web && cp pkg/pptx_parser_bg.wasm pkg/pptx_parser.js ../src/wasm/
cd packages/xlsx/parser  && wasm-pack build --target web && cp pkg/xlsx_parser_bg.wasm  pkg/xlsx_parser.js  ../src/wasm/
cd packages/docx/parser  && wasm-pack build --target web && cp pkg/docx_parser_bg.wasm  pkg/docx_parser.js  ../src/wasm/

# 全パッケージ一括
pnpm build:wasm
```

## Storybook

Storybook はルートに一本化（port 6006）。各パッケージのストーリーは `packages/*/src/*.stories.ts` に置く。

静的ファイルのパスプレフィックス（`.storybook/main.ts` の `staticDirs` で定義）:
- `packages/pptx/public/` → `/pptx/`
- `packages/xlsx/public/` → `/xlsx/`
- `packages/docx/public/` → `/docx/`

サンプルファイルを fetch する際は必ずプレフィックスを付ける（例: `/pptx/sample-1.pptx`, `/xlsx/sample-1.xlsx`）。

ローカル専用のサンプルストーリーは各パッケージの `Samples.sample.stories.ts` に置き、title は `<Viewer>/Samples` でネストさせる（例: `PptxViewer/Samples`）。`.gitignore` 済みなのでコミット対象外。

```bash
pnpm storybook        # dev server (port 6006)
pnpm build-storybook  # storybook-static/ にビルド
pnpm build:wasm       # 全パッケージの WASM をビルド（Storybook ビルド前に必要）
```

## テスト実行

```bash
npx playwright test --reporter=list
# または
pnpm vrt
```
