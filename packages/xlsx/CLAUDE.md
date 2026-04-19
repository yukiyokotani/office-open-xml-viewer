# @silurus/ooxml-xlsx — XLSX セッション用

## あなたのロール
このディレクトリ（`packages/xlsx/`）の編集責任者。XLSX の parser・renderer を 0 から実装する。

## MVP スコープ（Phase 2）
- ZIP 解凍 + `xl/workbook.xml`, `xl/sharedStrings.xml`, `xl/worksheets/sheet1.xml` パース
- セル値（文字列・数値）、`xl/styles.xml` から bold/color/背景色
- 列幅・行高
- `XlsxWorkbook.renderViewport(target, sheetIdx, {row, col, rows, cols})` 公開 API
- 非対応（後続）: 数式計算、マージセル、チャート、条件付き書式、freeze panes

## ディレクトリ構成

```
packages/xlsx/
├── package.json
├── tsconfig.json
├── vite.config.ts
├── CLAUDE.md
├── public/
│   └── sample.xlsx          ← テスト用 XLSX（git 未追跡）
├── src/
│   ├── index.ts             ← 公開 API
│   ├── types.ts             ← Rust JSON 出力と 1:1 対応
│   ├── workbook.ts          ← XlsxWorkbook クラス
│   ├── renderer.ts          ← Canvas 2D レンダラー
│   ├── worker.ts            ← Web Worker (WASM 呼び出し)
│   └── wasm/                ← wasm-pack ビルド出力（git 未追跡）
├── parser/
│   ├── Cargo.toml
│   └── src/lib.rs           ← XLSX パーサー（Rust）
└── tests/visual/
    ├── visual.spec.ts       ← Playwright VRT
    ├── fixture.html
    ├── references/          ← 正解画像（ユーザー指示のみ更新）
    ├── screenshots/
    └── diffs/
```

## WASM ビルド手順

```bash
cd packages/xlsx
npm run wasm
# または
cd packages/xlsx/parser && wasm-pack build --target web --out-dir ../src/wasm
```

## Storybook

Storybook はルート一本化のため、パッケージ単体では起動しない。
ルートから `pnpm storybook` で全パッケージのストーリーが参照できる。

## 編集してよいもの
- `packages/xlsx/**` すべて
- `packages/xlsx/parser/src/**`（Rust）

## 絶対に編集してはいけないもの
- `packages/core/**` ← 変更が必要なら main ブランチへ PR
- `packages/pptx/**` / `packages/docx/**` ← 他セッションの領域
- root の config 類 ← main で管理

## 参照画像
`packages/xlsx/tests/visual/references/` は Excel export PNG のみ配置。自動更新禁止。

## サンプルデータ
- `packages/xlsx/public/sample.xlsx` は git にコミットしない（pptx と同様のルール）
- Sales シート（10行×4列）と Summary シート（5行×2列）を含む
