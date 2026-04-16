# CLAUDE.md

## プロジェクト概要

OOXML (PowerPoint .pptx) をブラウザ上の Canvas に描画するライブラリ。
Rust/WASM パーサー + TypeScript Canvas レンダラー構成。

## ディレクトリ構成

```
pptx-parser/          ← Rust/wasm-pack パーサー
  src/lib.rs          ← OOXML解析コア。serde で camelCase JSON を出力
  pkg/                ← wasm-pack build の出力先

src/
  wasm/               ← ★ pkg/ の手動コピー先。アプリはここを読む
  types.ts            ← Rust の JSON 出力と1:1対応する TypeScript 型
  renderer.ts         ← Canvas 2D API でスライドを描画
  index.ts            ← PptxViewer 公開 API
  worker.ts           ← Web Worker で WASM を呼び出す

public/sample.pptx    ← テスト用 PPTX (5スライド)

tests/visual/
  visual.spec.ts      ← Playwright ビジュアルリグレッションテスト
  fixture.html        ← テスト用 HTML (width=1920)
  references/         ← 正解画像 slide-1.png〜slide-5.png
  screenshots/        ← 実行ごとに更新されるスクリーンショット
  diffs/              ← ピクセル差分画像
```

## WASM ビルド手順（重要）

```bash
cd pptx-parser && wasm-pack build --target web

# ★ 必ず src/wasm/ にコピーする（別物なので自動同期されない）
cp pptx-parser/pkg/pptx_parser_bg.wasm pptx-parser/pkg/pptx_parser.js src/wasm/
```

コピー忘れると古い WASM が使われ続ける。

## テスト実行

```bash
npx playwright test --reporter=list
# 例: slide 4: match=93.8%  diff=6.2%  (127,638 / 2,073,600 px)
```

## 現在のテスト結果 (session 3, 2026-04-16)

| スライド | match% | 備考 |
|---------|--------|------|
| 1 | 99.6% | |
| 2 | 100.0% | |
| 3 | 99.4% | |
| 4 | 99.0% | |
| 5 | 98.8% | |

## 修正済みバグ (session 2)

### タブストップ（スライド4「22%」の右揃えズレ）
- `pPr > tabLst > tab` をパース → `Paragraph.tabStops: TabStop[]` に格納
- `layoutParagraph` で `\t` トークン検出時に `tabStop.segments` に後続テキストを蓄積
- 描画時 `tabAbsX - totalTabW` で右揃えレンダリング

### grpFill 継承（flipオブジェクトが塗りつぶされない）
- `parse_sp_tree_node` / `parse_shape` に `group_fill: Option<&Fill>` を追加
- `spPr > grpFill` の図形が親グループの solidFill を継承するように
- スライド5のアワードバッジ（金賞等）のリース葉が accent4 ゴールド (#EBC83C) で塗られるように

## 修正済みバグ (session 3)

### グループ回転が子シェイプに未適用（リース葉の回転ズレ）
- `GroupTransform` に `rot: f64` フィールドを追加
- grpSp の `xfrm` から `rot / 60000` を読み取るように
- `apply_to_transform`: 子の中心をグループ中心周りに回転（clockwise screen coords）
- 子の rot の正しい公式: `child.rot = group.rot + (group.flipH XOR group.flipV ? -t.rot : t.rot)`
  - グループにネットflip（flipH XOR flipV）がある場合、子の回転方向が反転するため t.rot を負にする
  - 単純な `t.rot + group.rot` は誤り（flipH時に方向が逆になる）
- `apply_group_transform_to_element`: `s.rotation = nt.rot` / `p.rotation = nt.rot` を追加（以前は破棄していた）

### レイアウトプレースホルダーの枠線誤継承（タイトルの黒枠線）
- slideLayout の `spPr > ln` は編集モード用インジケーターで、描画時は不要
- `by_type_stroke` フィールドと `lookup_stroke()` を削除
- `parse_shape` での `lph.lookup_stroke()` 呼び出しを削除

### trapezoid adj（スライド5 角デコレーションが三角形になる）
- OOXML 仕様: `ss = min(w, h)`, `inset = adj / 100000 * ss`
- 修正: `const ss = Math.min(w, h); const inset = Math.min(w/2, adj/100000 * ss)`
- adj=99828, w=159, h=31 → inset=30.95px（正しい台形）

## 残課題

### スライド3・5 のフォントサイズが小さい
- プレースホルダーシェイプのフォントサイズがスライドレイアウト/マスターから継承されていない
- 調査先: `ppt/slideLayouts/` と `ppt/slideMasters/` の `lstStyle > lvl1pPr > defRPr sz`
- `parse_text_body` でレイアウト/マスターのデフォルトを読む処理が未実装

### autofit 未実装
- `bodyPr > spAutoFit` でテキストが収まらない場合にフォントサイズ縮小する仕様
- 現在はクリッピングのみ

### lumMod/lumOff の色変換精度
- `tx2 + lumMod=50000` などのスキームカラー修飾の近似精度に課題あり

### leftBracket などのプリセット形状
- 未実装の prstGeom は `rect` にフォールバック中（スライド5で使用あり）

## 重要な技術メモ

### OOXML 単位
- 回転: `rot / 60000` → 度
- フォントサイズ: `sz / 100` → pt  (例: 2400 → 24pt)
- スペース: `spaceBefore/spaceAfter` は hundredths of pt → `/ 100 * PT_TO_EMU * scale` で px

### テーマカラー (sample.pptx)
| name | hex |
|------|-----|
| dk2 / tx2 | #196ECA |
| accent1 | #E46970 |
| accent4 | #EBC83C (gold) |
| accent5 | #00A08C |

### layoutParagraph のシグネチャ
```typescript
layoutParagraph(ctx, para, maxWidthPx, defaultFontSizePx, defaultColor, scale, marLPx)
//                                                                                ↑ タブストップ計算用
```
