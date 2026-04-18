# @ooxml/docx — DOCX セッション用

## あなたのロール
このディレクトリ（`packages/docx/`）の編集責任者。DOCX の parser・renderer を 0 から実装する。

## MVP スコープ（Phase 2）
- ZIP 解凍 + `word/document.xml` パース（Rust）
- パラグラフ + ラン（bold/italic/color/font-size）
- 明示改ページ `<w:br w:type="page"/>` のみ
- `DocxDocument.renderPage(target, pageIndex)` 公開 API
- 非対応（後続）: テーブル、画像、ヘッダー/フッター、自動改ページ、脚注

## 編集してよいもの
- `packages/docx/**` すべて
- `packages/docx/parser/src/**`（Rust）

## 絶対に編集してはいけないもの
- `packages/core/**` ← 共有コード。変更が必要なら main ブランチへ PR
- `packages/pptx/**` / `packages/xlsx/**` ← 他セッションの領域
- root の config 類 ← main で管理

## 参考にしてよいもの（読み取り専用）
- `src/renderer.ts` — Canvas 描画プリミティブの参考
- `pptx-parser/src/lib.rs` — wasm_bindgen + roxmltree の使い方
- `src/types.ts` — 型定義パターンの参考

## 参照画像
`packages/docx/tests/visual/references/` は Word export PNG のみ配置。自動更新禁止。

## ディレクトリ構成

```
packages/docx/
├── CLAUDE.md
├── package.json
├── tsconfig.json
├── vite.config.ts
├── public/              ← sample .docx ファイル置き場（git 管理対象外）
├── src/
│   ├── index.ts         ← 公開 API
│   ├── types.ts         ← Document モデル型（Rust JSON 出力と1:1）
│   ├── document.ts      ← DocxDocument クラス
│   ├── viewer.ts        ← DocxViewer クラス
│   ├── renderer.ts      ← renderPage 実装
│   └── wasm/            ← wasm-pack 出力（git 管理対象外）
├── parser/
│   ├── Cargo.toml
│   └── src/
│       ├── lib.rs       ← WASM エントリポイント
│       ├── types.rs     ← Rust 型定義
│       └── parser.rs    ← DOCX パーサー実装
└── tests/visual/
    ├── visual.spec.ts
    ├── fixture.html
    ├── references/      ← Word export PNG（git 管理対象外）
    ├── screenshots/
    └── diffs/
```

## WASM ビルド手順

```bash
cd packages/docx && npm run wasm
# または
cd packages/docx/parser && wasm-pack build --target web --out-dir ../src/wasm
```

## テスト実行

```bash
npx playwright test packages/docx/tests/visual
```
