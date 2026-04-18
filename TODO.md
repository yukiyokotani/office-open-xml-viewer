# 残課題一覧

最終更新: 2026-04-18 (session 4)

---

## 高優先度（スコアへの影響大）

### 1. sample-4 slide 13「28%」縦位置ズレ（match=87.1%）

**症状**: テキストキャップトップが参照画像 y≈175px に対しスクリーンショット y≈249px（約74px下）

**根本原因**:
- `renderTextBody` の baseline 計算: `cursorY + lineHeight * 0.8`
  - デフォルト1.2×行間だと `fontSizePx * 0.96` になり、PowerPoint実際の ascent（`fontSizePx * 0.77`前後）より大きくなる → テキストが過度に下へずれる
- レイアウト `lstStyle > lvl1pPr > lnSpc`（この場合90%）がスライドのパラグラフに継承されていない（Rustパーサー未実装）

**試した修正**（net マイナスで没）:
- Rust: `LayoutPlaceholders` に `by_idx_space_line` 追加、`parse_paragraph` に継承 spaceLine フォールバック追加
- Renderer: `baseline = cursorY + maxSizePx * 0.8`（lineHeight ではなく fontSizePx ベース）
- 結果: slide 13 +1.1% だが slide 7/12 が -0.4〜1%、sample-2 全体も小幅退行 → 全体では net マイナス

**正しい修正方針**:
- `ctx.measureText()` の `actualBoundingBoxAscent` を使い、実際のフォントメトリクスで baseline を決定する
- spaceLine 継承（Rust）と組み合わせれば精度向上の見込み

---

### 2. sample-3 slide 5（match=88.9%）

**症状**: プリセット図形グリッドが rect フォールバック中

**対象図形**（現在 rect 表示）:
- `noSmoking` / `noSmokingSign`（ケース名不一致：renderer は `nosmokingsign` を期待）
- `ellipseRibbon`, `ellipseRibbon2`
- `stripedRightArrow`
- flowchart 系15種類（後述）
- `smileyFace`（輪郭のみ、顔パーツなし）
- `uturnArrow`（形状不正確）

---

### 3. sample-4 slide 9（match=89.0%）

**症状**: "OUR CUSTOMERS" スライド、フォント・レイアウトの差異
未調査のため詳細不明。

---

### 4. sample-2 slide 8（match=89.1%）

未調査。

---

## 中優先度（図形レンダリング）

### 未実装プリセット図形（rect フォールバック中）

| 図形名 | 備考 |
|--------|------|
| `flowChartPreparation` | 六角形（左右に斜め辺） |
| `flowChartCollate` | 砂時計形 |
| `flowChartMagneticDisk` | 横向き円柱 |
| `flowChartInternalStorage` | 矩形＋内部縦横線 |
| `flowChartMagneticDrum` | 矩形＋左半円 |
| `flowChartSummingJunction` | 丸に X |
| `flowChartMagneticTape` | 波底面の矩形 |
| `flowChartPunchedTape` | 波底面の矩形（別形） |
| `flowChartManualOperation` | 逆台形 |
| `flowChartMultidocument` | 複数文書重ね |
| `ellipseRibbon` | 下部楕円弧リボン |
| `ellipseRibbon2` | 上部楕円弧リボン |
| `stripedRightArrow` | 縦縞＋右矢印 |
| `smileyFace` | 輪郭のみ実装、目・口が未実装 |
| `noSmoking` | ケース名不一致（`nosmoking` vs `nosmokingsign`）+ パス未実装 |
| `uturnArrow` | 形状不正確 |

---

## 低優先度 / 仕様上の制限

### lumMod/lumOff の色変換精度
- `tx2 + lumMod=50000` などのスキームカラー修飾の近似精度に課題

### sample-4 縦書きテキスト（slide 2付近）
- レイアウト形状に `rot="16200000"` (270°) + 負の x 座標が組み合わさっている
- 現状おおよそ表示できているが精度確認が必要

---

## 全体スコア現状（session 4 終了時点）

| サンプル | 最低 | 最高 | 備考 |
|---------|------|------|------|
| sample-1 (5枚) | 90.6% | 95.2% | slide 3 が最低 |
| sample-2 (17枚) | 89.1% | 99.2% | slide 8 が最低 |
| sample-3 (8枚) | 88.9% | 99.1% | slide 5 が最低 |
| sample-4 (15枚) | 87.1% | 98.4% | slide 13 が最低 |

---

## session 4 で完了した作業

- テーブルスタイル（背景色・枠線）対応: `tableStyles.xml` パース、firstRow/bandRow 塗り、内部枠線
- テーブルセル縦アンカー（top/center/bottom）対応: `tcPr > anchor` を `verticalAnchor` に格納
- ページ番号表示: `<a:fld type="slidenum">` を `fieldType: "slidenum"` としてパース、レンダラーで実スライド番号に置換
- `Slide.slideNumber`（1-based）追加と renderSlide → renderShape/renderTable への伝播
