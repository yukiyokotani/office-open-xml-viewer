# @silurus/ooxml-xlsx — XLSX セッション用

ルートの `AGENTS.md`（正）と `CLAUDE.md` を先に読むこと。ワークフロー、
ブランチ、VRT、参照画像のルールはそちらに従う。このファイルは XLSX 固有の
事実だけを記録する。

## 編集範囲
- 主に `packages/xlsx/**`（`parser/`、`model/` の Rust を含む）。
- `packages/core` / `packages/ooxml-common` や他フォーマットへの変更は、
  `AGENTS.md` の Cross-Package Integration に沿った一貫した修正の場合のみ。

## 数式の扱い（方針）
- セルの数式（`<f>`）は一切計算しない。表示は常にキャッシュ値（`<v>`）。
  `TODAY()` / `NOW()` も例外ではない（1 セルだけ現在時刻にすると、そこから
  派生するセルはキャッシュ値のままになり、表示が不整合になるため）。
  キャッシュ値のない数式セルは、関数にかかわらず値のない（空の）セルとして
  扱い、その値に表示形式を適用する（表示形式によっては文字が出る）。
- 条件付き書式の数式（`expression` 等）は、キャッシュ値を入力とする部分的な
  評価器（`src/formula.ts`）で評価する。この評価器の中の `TODAY()` / `NOW()`
  は現在時刻を参照する。完全な評価は #1547 で追跡。
- 例外: `feature/spreadsheet-formulas-support` ブランチは上記方針の意図的な
  フォークである。`LoadOptions.recalculateFormulas: true` のオプトイン時のみ
  `src/recalc.ts` が Formualizer（MIT / Apache-2.0）で `<f>` を再計算し、
  結果を `cell.value` に書き戻す（`cell.formula` は不変）。デフォルト（未指定）
  は従来通りバイト同一。`formula.ts`・`number-format.ts`・`renderer.ts` の呼び出し
  側は変更禁止。再計算の失敗はキャッシュ値にフォールバックし、描画を壊さない。

## ディレクトリ構成（主要なもの）

```
packages/xlsx/
├── parser/src/          Rust パーサー（lib.rs, styles.rs, drawing.rs, chart.rs,
│                        pivot.rs, worksheet_*.rs, …）
├── model/src/           Rust モデル crate
├── src/
│   ├── index.ts         公開 API
│   ├── types.ts         Rust 出力と 1:1 対応
│   ├── workbook.ts      XlsxWorkbook
│   ├── renderer.ts      Canvas 2D レンダラー
│   ├── worker.ts / render-worker.ts  パースワーカー / レンダーワーカー
│   └── wasm/            wasm-pack 出力（生成物・git 管理対象外）
├── public/demo/         追跡対象の demo サンプル（sample-1.xlsx）
├── public/private/xlsx/ ローカル専用の private コーパス（git 管理対象外）
└── tests/visual/        Playwright spec（visual.spec.ts, worker-equivalence.spec.ts, …）
    └── references/demo/<name>/sheet-N.png + scores.json
                         採用済みレンダラー出力（Excel export ではない）。自動更新禁止。
```

## WASM ビルド

```bash
pnpm --filter @silurus/ooxml-xlsx wasm
```

`parser/` で `wasm-pack build --target web --out-dir ../src/wasm` を実行し、
続けて `scripts/append-wasm-reinit.mjs` を適用する。

## テスト
- ユニットテスト: ルートで `pnpm test`（vitest）。
- VRT: `pnpm --filter @silurus/ooxml-xlsx vrt` / `vrt:private` / `vrt:fidelity`。
  ベースライン用の環境変数はルート `AGENTS.md` を参照。
- Storybook はルートで `pnpm storybook`。
- サンプルのコミット禁止: private コーパスと生成物はコミットしない。
