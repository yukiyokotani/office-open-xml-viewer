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
- `git push origin <branch>` して GitHub で PR を作成し、PR 経由で main へマージする
- `git push origin main` は絶対に行わない（直接 push 禁止）
- **squash merge は使わない。** merge commit（`--no-ff`）または rebase merge を使うこと。squash すると feature branch の commit 粒度が失われ、bisect / revert の単位が粗くなる
- `git push` 前に `git config http.postBuffer 524288000` を設定すること

## 自律作業の原則

- AM1時〜AM9時はユーザー確認不要。破壊的操作以外はすべて自律的に進めること。
- 確認なしで進めてよい作業: コード修正・WASM ビルド・テスト実行・commit/push（feature branch のみ）・Python/npm スクリプト実行。
- 参照画像（`packages/*/tests/visual/references/`）はユーザー指示のみ更新。絶対に自動更新しない。
- pptx/xlsx/docx ファイルは git にコミットしない。

## 実装方針: ヒューリスティックより仕様忠実を優先

- VRT を一時的に良くするためだけのヒューリスティック（「M > 2 なら grid snap」「auto > 720 は atLeast と見なす」「body は natural × M で heading は max(natural, pitch × M)」等）を**入れない**。
  短期的に数字が上がっても別サンプルで後退し、理由を書けない挙動が積み重なる。
- まず ECMA-376 / ISO-29500 の該当節を読み、Word が実際にどう解釈しているか（docGrid の snap ルール、line rule の各意味、paragraph mark sz の扱い、spacing 継承の各属性、compat フラグなど）を突き止める。
- 仕様との差の原因が分からないときは、parser 側で情報を捨てていないか（inherit / merge で潰れていないか）を先に疑う。情報が足りなければ parser を拡張するのが正道。
- 工数が増えても spec に忠実な実装を選ぶ。empirical な定数（1.15、0.25、ceiling 付きの条件分岐など）を入れそうになったら、いったん手を止めて「どの §x.x.x の挙動なのか」を書き出す。書き出せないなら実装しない。
- Excel / PowerPoint / Word の UI 挙動（spec に書かれていないランタイム autofit 等）を reverse-engineering して合わせる場合は、事前にユーザー承認を得ること。迷ったら spec 通りを選ぶ。
- 既存コードに上の原則に反するコードが残っている場合は、触る機会があったら正道に寄せる。

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
