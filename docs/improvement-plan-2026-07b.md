# プロジェクト総合レビューと改善計画(2026-07 第2次)— 「比類なき存在」ロードマップ

> **履歴資料 (2026-08-17 更新):** この文書は 2026-07 時点の候補と見積りを
> 保存したもので、現行バックログではない。機能の現状はルート README、実装、
> テスト、open issue を参照すること。関連資料の位置付けは [`README.md`](README.md)
> にまとめている。

対象: monorepo 全体。前次計画(2026-07、全57項目、PR #658–#703)完了後の状態(main = 1610253)を起点に、
**設計・品質・パフォーマンスで他ライブラリの追随を許さない存在にする**ことをゴールとして再レビューした。

## レビュー手法

13観点(core TS / docx / pptx / xlsx / Rust アーキテクチャ / WASM 境界 / 描画性能 / 仕様フィデリティ / API・DX / テスト / 堅牢性・セキュリティ / 競合分析 / ユーザー UX)の並列レビュー →
全指摘を「反証を既定とする」独立検証エージェントがコード上で裏取り → 全体俯瞰の網羅性チェック、の3段で実施。
**検証対象 claim 152 件中 REFUTED 0(CONFIRMED 145 / PARTIAL 7)**、網羅性チェックで8件追加。計161指摘を重複統合し、本計画の約105項目に整理した。
PARTIAL 7 件は本文の該当項目に補正済み(例: clrChange recolor は docx のみ実装済み、renderer.ts の行番号ずれ等)。

## 総評

前次計画の成果は全観点で「本物」と確認された: worker bridge、UAX#9 bidi(Unicode 公式データ駆動の conformance テスト)、
スクロールビューアの epoch-guarded zoom、パニック監査(全4クレートで production `unwrap()` 4箇所のみ・全て構造的ガード付き)、
測定/描画統一、ooxml-common への共有化はいずれも健全。**負債は「コードの質」ではなく「カバレッジの穴」と「ビューアの上物の薄さ」に集中している。**

戦略テーマは6つ:

1. **チャートが看板機能なのに最大の弱点** — 正負・積み上げの正しさバグ、軸モデルの構造的欠落、ファミリー間の機能非対称、
   docx 完全未対応、pptx/xlsx でパーサー二重実装(そこから実際に色解決の仕様乖離が発生済み)。
2. **「core / ooxml-common にあるのに1フォーマットしか使わない」横断ギャップ** — effectLst・gradFill/pattFill・fillRef・rPr・
   縦書き・下線スタイル。AGENTS.md の横断統合原則の適用対象が今も残っている。
3. **インタラクション層の欠落** — ハイパーリンクが3フォーマットとも「塗るだけで踏めない」、検索なし、キーボードなし、
   スクリーンリーダー不可視(alt text はパーサーが落としている)、タッチピンチなし。競合対比で最も目立つ穴。
4. **文書提供リソースの不使用** — 埋め込みフォント(docx ODTTF / pptx fntdata)を両パーサーが完全に無視。企業文書の忠実度に直撃。
5. **xlsx のスケール天井** — 200k 行で JSON 199MB / メインスレッド 4.5 秒フリーズ / WASM 1.17GB 常駐(実測)。
   sharedStrings の dedup をワイヤ上で自ら破壊している。
6. **未出荷の資産** — markdown 変換は全ユーザーの WASM に既に入っているのに公開 API なし。node レンダリングは private。
   PDF 出力は「本物のレイアウトエンジンからベクター PDF を吐ける」唯一のアーキテクチャなのに存在しない。
   競合数値(per-format WASM 0.8–0.9MB vs Apryse 展開後 265MB、MIT vs 営業見積もり)も未使用。

このほか横断で: OOXML **Strict** 文書(ISO/IEC 29500 Strict 名前空間)が3フォーマットとも真っ白になる、
CI が Chromium のみ、パーサーのファジング 0 件、MathJax (Apache-2.0) の帰属表記が欠落、という基盤の穴が確認された。

**既知の継続スレッド**(本計画と別トラック): issue #698(SPACE_SHRINK_RATIO)、A2 残り120プリセットの spec エンジン移行(VRT 参照更新承認とセット)。

---

## 領域別指摘(全件コード裏取り済み)

工数: S=~半日 / M=~2-3日 / L=~1-2週 / XL=それ以上。場所は代表点のみ(詳細は各項目の実装時に再確認)。

### CH — チャートエンジン

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| CH1 | 棒/縦棒チャートが**負の値を正の上向きバーで描画**(§21.2.2) | `core/src/chart/renderer.ts:567` | S |
| CH2 | **stackedLine / stackedLinePct が積み上げずに描画**(area の stacked 分岐と同型で直せる) | `core/src/chart/renderer.ts:2429` | S |
| CH3 | xlsx チャート色解決が古い私的コピー経由: **bg2/tx2 が入れ替わり、lumMod/lumOff を HSL でなく RGB 空間で適用**。`parse_solid_fill`(ooxml-common 委譲済み)へ一本化 | `xlsx/parser/src/chart.rs:1585-1655` | S |
| CH4 | 日付カテゴリ軸 `c:dateAx`(§21.2.2.39)未認識 — 時系列チャートが生シリアル値表示になり軸書式も全損 | `xlsx/parser/src/chart.rs:654`, `ooxml-common/src/chart.rs:470` | S |
| CH5 | **チャート XML パースが pptx(1,130行)/xlsx(2,807行)で二重実装**。`parse_chart_part(xml, theme) → ChartModel` を ooxml-common に一本化(CH3 の乖離の根本原因。docx チャート・xlsx chartEx の enabling step) | 両 `chart.rs` | L |
| CH6 | 軸モデルの構造的欠落: majorGridlines 有無・majorUnit/minorUnit・logBase・orientation maxMin・minor gridlines・目盛りラベル rot/skip・トレンドライン(§21.2.2) | `core/src/chart/renderer.ts:687`, `axis-scale.ts:53` | M |
| CH7 | 第2値軸が bar ファミリーのみ — line/area/scatter のコンボが単一スケールに潰れる。bar の実装(:581-592, :951-984)を axes ヘルパーへ抽出し全ファミリーから呼ぶ | `core/src/chart/renderer.ts:512` | M |
| CH8 | pie/doughnut: holeSize・firstSliceAng・explosion・多重リング・dLbls 全部無視、ドーナツ穴が不透明白塗り | `core/src/chart/renderer.ts:1341` | M |
| CH9 | line ファミリーがマーカー・エラーバー・per-point ラベル・c:smooth・dispBlanksAs を無視(モデルと scatter 実装は既にある — 配線のみ) | `core/src/chart/renderer.ts:1888` | M |
| CH10 | タイトル以外の全チャートテキストが `sans-serif` 固定 — テーマフォント(fontScheme は ooxml-common で解析済み)と `c:txPr` が canvas に届かない | `core/src/chart/renderer.ts` | M |
| CH11 | チャートの目盛り/データラベル数値が **`toLocaleString()` で閲覧者の OS ロケール依存**(同じファイルが 1,000 / 1.000 / 1 000 に化ける)。セルと同じ §18.8.30 エンジンへ | `core/src/chart/renderer.ts:2259` | S |
| CH12 | **docx はチャートを一切描画しない** — graphicData chart URI → rels → chartN.xml のパース(CH5 後は共有パーサー呼ぶだけ)+ ChartModel 要素の emit | `docx/parser/src/parser.rs`(graphicData 分岐) | M(CH5 後) |
| CH13 | 3D 系(bar3D/pie3D/line3D/area3D)は 2D へ平坦化写像(Google Slides/Keynote 同等)、stock は hi-lo-close(エラーバー描画部品を再利用)、ofPie は連結2パイ。現状は「Chart: unknown」の文字列描画 | `core/src/chart/renderer.ts:2448`, `pptx/parser/src/chart.rs:40-91` | M |
| CH14 | xlsx が chartEx(2014 namespace)を完全スキップ — Excel 2016+ の waterfall/treemap/sunburst/histogram/funnel が消える(pptx にはパーサーあり → CH5 で共有化) | `xlsx/parser/src/chart.rs:379` | M |
| CH15 | chartEx レンダラー群: funnel/histogram(bar 変種で小)→ treemap(squarify)→ sunburst(階層 pie)→ boxWhisker(エラーバー部品再利用)。waterfall で確立したパターンの反復 | `core/src/chart/renderer.ts:2421-2453` | L |

### XF — DrawingML 横断パリティ(共有層)

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| XF1 | **OOXML Strict(ISO/IEC 29500 Strict)文書が3フォーマットとも真っ白** — パーサーが Transitional 名前空間しかマッチしない。`is_w_ns()/is_a_ns()/is_r_ns()` 述語を ooxml-common に集約し両 URI を受理 | `docx/parser/src/parser.rs:795`, `xlsx/parser/src/lib.rs:482` 他全マッチ点 | M |
| XF2 | `<a:effectLst>`(影/グロー/softEdge/reflection)を **pptx しかパースしない** — docx/xlsx は図形・画像エフェクト全損。parse_effect_lst を ooxml-common 化し3フォーマットで emit、core の既存エフェクト描画へ | `pptx/parser/src/fill.rs:358`, docx/xlsx 各所 | M |
| XF3 | xlsx/docx の図形塗りが solid のみ — **gradFill/pattFill を silently drop**(ooxml_common::fill のパーサーは既に export 済み。消費するだけ) | `xlsx/parser/src/drawing.rs:918`, `docx/parser/src/parser.rs:4388` | M |
| XF4 | style-matrix **fillRef が pptx/xlsx で単色に平坦化**(fillStyleLst レシピ実装は docx のみ)。resolve_fill_ref を ooxml-common へ | `pptx/parser/src/shape.rs:274`, `xlsx/parser/src/drawing.rs:977` | M |
| XF5 | DrawingML rPr パース三重実装で忠実度損失 — xlsx/docx の図形テキストで **underline/strike/baseline/hyperlink が drop**。ooxml_common::text へ(BodyPr の前例踏襲) | `xlsx/parser/src/drawing.rs:615`, `docx/parser/src/parser.rs:4065` | M |
| XF6 | グラデーション: `lin@scaled`・`path`(radial/rect)・`fillToRect`/`tileRect` 未対応 — 非正方形図形で角度が狂い、オフセンター radial が中央寄せになる(§20.1.8.30/.36/.41) | `ooxml-common/src/fill.rs:119`, `core/src/shape/paint.ts:68` | M |
| XF7 | 画像 recolor: duotone/grayscl/biLevel は3フォーマットとも未実装、lum も未対応(clrChange は docx のみ実装済み)。共有 BlipEffects(types+parse)+ core のピクセルパスへ | `docx/parser/src/parser.rs:3382`(clrChange 実装例) | M |
| XF8 | `scrgbClr`/`hslClr` が None 解決 — **これらを使う塗り/線が消える**。prstClr の transform 未適用、ST transform 数種(comp/gray/gamma 等)無視 | `ooxml-common/src/color.rs:268-297` | S |
| XF9 | **縦書きの非対称**: docx は皆無(w:textDirection・bodyPr@vert・w:eastAsianLayout 全部未パース)、pptx も eaVert 近似で wordArtVert/mongolianVert・表セル vert が drop。段階実装: (1) docx textbox bodyPr@vert(common BodyPr は既に保持)(2) docx tcPr w:textDirection(3) pptx 残り | `pptx/src/renderer.ts:2007`, docx 各所 | L |
| XF10 | pptx **themeOverride part(§14.2.7)不使用** — オーバーライド持ちスライド/レイアウトがマスターテーマの色/フォントで描かれる | `pptx/parser/src/master.rs:1521` | M |
| XF11 | docx が **17種の w:u 下線スタイル(+w:u@color)をパース時に boolean へ潰す**(pptx は全種描画済み)。ST_Underline を model に通し、pptx の drawUnderline を core へ hoist して docx/xlsx rich-text から使う | `docx/parser/src/styles.rs:1081` | S |
| XF12 | `cNvPr@hidden="1"` の図形を描画してしまう(作者が隠したものが見える)。§20.1.2.2.8。**3フォーマット同時に**読む | `pptx/parser/src/shape.rs:96-107` 他 | S |
| XF13 | EMF プレイヤー完成: window/viewport mapping(SETMAPMODE 系 [MS-EMF] 2.3.11)欠落で MM_ANISOTROPIC 系がスケール狂い/空描画、pptx パスでは EMF が null に落ちて消える。併せて README の「EMF not yet rendered」誤記も是正(実態は player 出荷済み) | `core/src/image/emf.ts:49-93`, `core/src/image/bitmap-image-by-path.ts:52` | M/L |
| XF14 | **埋め込みフォント不使用(全フォーマット)**: docx `w:embedRegular`+ODTTF(GUID XOR 32byte、§15.2.13)、pptx `p:embeddedFontLst`(.fntdata=素の TTF、§19.2.1)。Archive handle に extract_font を足し、FontFace 登録をレイアウト前に(node/skia パスにも同等品) | `docx/parser/src/parser.rs:782`, `pptx/parser/src/lib.rs:769` | L |

### WD — docx 固有フィデリティ

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| WD1 | `w:noBreakHyphen`/`w:cr`/`w:ptab`(+softHyphen)が**黙って drop = 可視テキスト喪失**(§17.3.3.18/.4/.23/.29) | `docx/parser/src/parser.rs:2398-2695` | S |
| WD2 | PAGE フィールドが `w:pgNumType`(restart/format §17.6.12)とフィールド書式スイッチを無視 — 前付のある文書でページ番号が狂う | `docx/src/line-layout.ts:826`, parser | M |
| WD3 | `w:object`(OLE)と VML `v:imagedata`/`v:textpath` が何も描かない — **埋め込み Excel/数式のプレビューと透かし(watermark)が不可視**。parse_vml_pict の拡張で既存 blip パイプラインへ | `docx/parser/src/parser.rs:2398-2695` | M |
| WD4 | run レベル `w:spacing`/`w:w`/`w:position`/`w:kern` 未パース — 文字間隔付き見出しの測定が狂い、**折返し・改ページ位置が Word とズレる**(measure==paint を保って ctx.letterSpacing で) | `docx/parser/src/styles.rs:10-95` | M |
| WD5 | `w:em` 圏点(§17.3.2.12)未パース — JP 特化レンダラーとして標準機能。ruby の ascent 拡張配管を共用 | `docx/parser/src/styles.rs` | S |
| WD6 | sectPr 装飾: `w:pgBorders`/`w:lnNumType`/セクション `w:vAlign` 未パース(§17.6.10/.8/.23) | `docx/parser/src/parser.rs:5232` | M |
| WD7 | 隣接セル境界の競合解決(§17.4.66 weight アルゴリズム)未実装 — 内部グリッド線で後塗り勝ちの恣意的結果 | `docx/src/renderer.ts:7463` | M |
| WD8 | `jc=kashida`/`thaiDistribute` が左寄せに落ちる — まず既存 'both' justify へマップ(ST_Jc の正当な解釈)、真の kashida(U+0640 挿入)は issue 化 | `docx/src/types.ts:364` | S |
| WD9 | OMML: `m:phant` の中身が**見えてしまう**(不可視スペーシングが正)、`m:sPre`/`m:box`/`m:borderBox` が線形 run に平坦化。MathML の mphantom/mmultiscripts/menclose へ | `ooxml-common/src/math.rs:127-249` | S |

### PP — pptx 固有フィデリティ

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| PP1 | OLE オブジェクトが**空穴**で描かれる — `p:oleObj` 配下の `<p:pic>` プレビュー画像(§19.3.2.4)を通常 Picture として emit するだけで直る(README の Not Planned を撤回) | `pptx/parser/src/shape.rs:1662-1727` | M |
| PP2 | SmartArt の data↔drawing 対応を**ファイル名末尾数字のヒューリスティック**で推測 — `dsp:dataModelExt relId` の正規バインディングへ(ヒューリスティック排除原則) | `pptx/parser/src/lib.rs:353-375` | S |
| PP3 | 保存済み drawing の無い SmartArt が**何も描かない** — 段階: (S) 枠のプレースホルダ (M) data1.xml の `dgm:ptLst` から全ノードの txBody を取り出しテキストリスト描画 (XL) hierarchy/cycle/process の native layout エンジン | `pptx/parser/src/shape.rs:1677` | M→XL |
| PP4 | WordArt(`prstTxWarp` §20.1.9.19)完全未対応 — warp プリセットは preset-geometry と同じ guide-formula 言語なので既存エンジンで評価可能 | `pptx/parser/src/text.rs:688` | L |

### XL — xlsx 固有フィデリティ

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| XL1 | **date1904 未パース — 1904 エポックのブックで全日付が 1,462 日(4年)ズレる**(§18.2.28)。併せて 1900-02-29 バグ互換(§18.17.4.1)も正す | `xlsx/src/number-format.ts:171` | S |
| XL2 | General 書式が f64 の生 round-trip 精度を表示 — 計算値が `0.30000000000000004` に。Excel の 11 有効桁規則を実装(ground-truth ループで桁数確定) | `xlsx/src/number-format.ts:375` | S |
| XL3 | 数値書式エンジン: `[$€-407]` ロケール通貨記号 drop・分数書式壊滅・セクション色 `[Red]` 無視・§18.8.31 条件セクション未対応・カンマスケーリングなし(テスト12件のみ)。文法は完全仕様定義 — corpus テーブル駆動で | `xlsx/src/number-format.ts:447` | M |
| XL4 | アウトライングルーピング(行/列グループ §18.3.1.73/.13/.71)完全未対応 — gutter・±ボタン・collapse | `xlsx/parser/src/lib.rs:955` | L |
| XL5 | ふりがな(`rPh`/`phoneticPr` §18.4.6/.3)未対応 — JP 投資との整合 | sharedStrings パス | M |
| XL6 | **C12 未解決の実害**: グリッドジオメトリが viewer/renderer 5箇所に手動ミラー、丸め規約2系統 — 小数ズームで**別セルヒット**の潜在バグ + 毎フレーム O(n) walk。閉形式の共有 ScaledAxis モジュール(純関数)に一本化 | `xlsx/src/renderer.ts:3134`, `viewer.ts:204` | M |

### IX — インタラクション & アクセシビリティ

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| IX1 | **ハイパーリンクが3フォーマット全てで inert**(塗るだけ)。内部アンカー(docx w:anchor+bookmark / xlsx location / pptx 内部スライド)はパーサーが**収集すらしない**。text layer に real `<a>`(rel=noopener)+ onLinkClick + 内部ジャンプ + **javascript:/data: サニタイズ必須** | 3 viewer + 3 parser | M |
| IX2 | **文書内検索が存在しない** — findText API(page/slide/sheet + rects + context を既存 run 幾何から構築)+ ハイライトオーバーレイ + next/prev。商用ビューア全てが持つ table stakes | 3 engine + text-layer | L |
| IX3 | キーボード操作ほぼゼロ: docx/pptx はゼロ、xlsx は矢印キーセル移動なし、フォーカス可能要素なし。host に tabindex + PageUp/Down/Arrow/Home/End、xlsx は role=grid + Ctrl+Arrow ジャンプ | 3 viewer | M |
| IX4 | **スクリーンリーダーに完全不可視**。既存 text layer を PDF.js 型のアクセシビリティ層へ昇格(常時 DOM 出力、段落/行グループ、reading order、role=document、見出し→aria-level) | `docx/pptx/src/text-layer.ts`, xlsx | L |
| IX5 | **alt text(wp:docPr / cNvPr @descr/@title)をパーサーが drop** — 3フォーマット追加フィールドとして emit し IX4 が消費 | 3 parser | S |
| IX6 | worker モードで**テキスト選択が黙って消える**(docx/pptx。run 幾何は structured-clone 可能 — bitmap と一緒に返すだけ) | `pptx/src/render-worker.ts:133`, `docx/src/worker-protocol.ts:19` | M |
| IX7 | worker モードの OMML 数式 — MathJax LiteDOM による in-worker SVG 化と共通 Canvas path ラスタライズを実装。main/worker の画像差分検証を継続 | `core/src/math`, 各 `render-worker.ts` | M |
| IX8 | **タッチピンチズームなし**(ctrl+wheel はトラックパッド合成イベントのみ)— 2-pointer トラッカーを core/interaction に足し3ビューアで共用。モバイル対応の第一歩 | `core/src/interaction/zoom.ts` + 3 viewer | M |
| IX9 | ズーム契約の不統一: 単canvas DocxViewer/PptxViewer はズーム API 自体なし、スクロールビューアは UI なし、fit-width/fit-page モードなし、xlsx だけスライダー。setScale/zoomIn/Out/fitWidth/fitPage を4ビューア共通契約に | 全 viewer | M |
| IX10 | `load()` に進捗報告なし(100MB ファイルで first paint まで無音)。`onProgress({phase, loaded, total})` を共有 LoadOptions へ(fetch は body reader 計数) | `core/src/types/load-options.ts` | M |
| IX11 | 破損ファイル/描画失敗で**何も表示されない** — デフォルト(上書き可)のエラーステートパネル + スクロールビューアの per-page 失敗カード | 全 viewer | S |
| IX12 | xlsx の document-level Ctrl+C が**ホストページ全体のクリップボードを乗っ取る** — フォーカス所有権でゲート(IX3 とセット) | `xlsx/src/viewer.ts:1853` | S |
| IX13 | UI 文字列が英語ハードコード・ライトテーマ固定 — `strings` オプション + 小さな theme token セット | xlsx viewer, scroll-viewer chrome | S |

### SC — スケール性能・メモリ

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| SC1 | **xlsx ワイヤモデルが sharedStrings dedup を自壊**: セル毎に full string+rich-text runs をクローン(string table も別送で未使用)+ 死にフィールド `colRef`/`row`/未 skip デフォルトで **シート JSON の約半分が冗長**。`CellValue::Shared{si}` 参照化 + serde skip + TS 側 index 解決 | `xlsx/parser/src/lib.rs:1955`, `types.rs:1333` | M |
| SC2 | worksheet XML の**ストリーミングパスなし**(full String + full DOM、実測 9.4× メモリ倍率)— `<sheetData>` のみ quick-xml/pull 走査で直接 Row/Cell へ、前後は roxmltree 維持 | `xlsx/parser/src/lib.rs:152` | L |
| SC3 | **200k 行 xlsx(6MB)が JSON 199MB で境界通過: wasm parse 6.5s + main JSON.parse 4.5s フリーズ + wasm 1.17GB 常駐**(実測)。XlsxArchive 保持を活かし `sheet_meta`(寸法/metrics/merges)+ `sheet_rows(from,to)` の窓配送へ | `xlsx/src/workbook.ts:221`, `parser/src/lib.rs:2127` | L |
| SC4 | xlsx がスクロール毎フレームで全可視セルの数値書式(toLocaleString+Date 構築)・wrap レイアウト・フォント文字列を再計算 — SheetRenderCache に per-cell memo 3種 | `xlsx/src/renderer.ts:1835,2213` | M |
| SC5 | xlsx が毎フレーム**全可視チャートを full renderChart** — (anchor, cw, ch, cs, dpr) キーの ImageBitmap キャッシュで blit 化 | `xlsx/src/renderer.ts:4167` | S |
| SC6 | アンカー配置が毎フレーム O(row-index) 線形 walk + 全 merges 走査 — viewer に既にある AxisMetrics(prefix-sum)を共有モジュール化して O(log n) に(XL6 と同一 PR 圏) | `xlsx/src/renderer.ts:3134`, `2754` | S |
| SC7 | xlsx スクロールが毎フレーム全ビューポート再描画 — 同一 (cs,dpr,size,sheet) なら前フレームを blit し露出帯のみ描画 | `xlsx/src/renderer.ts:2710` | L |
| SC8 | docx renderPage が**文書全体の全画像 decode を await してから**1ページを塗る(canvas を白紙化した後に)— per-page の画像だけ待ち、近傍ページは background warm | `docx/src/renderer.ts:769-786` | M |
| SC9 | docx first paint が直列: フォント preload → **全文書 pagination(同期・main блок)** → 全画像 decode。ページ毎 yield の incremental pagination + 先頭ページ即描画 | `docx/src/document.ts:99-118,244` | L |
| SC10 | **テキスト advance キャッシュが皆無** — docx page paint が mount/settle 毎に全 segment を再計測、justified CJK は O(pieces) の prefix 計測。(fontString, text)→width の bounded LRU を core に(FontFaceSet loadingdone で無効化) | `docx/src/line-layout.ts:2093`, `core/src/text/justify-positions.ts:69` | M |
| SC11 | スクロールビューアが slot recycle で**描画済みピクセルを破棄** — 戻りスクロールは常に full re-layout+repaint。(page, scale, dpr) キーの byte-budget ImageBitmap LRU | `docx/src/scroll-viewer.ts:577` | M |
| SC12 | docx/pptx の scroll handler が rAF 非結合(xlsx は結合済み)+ computeVisibleRange が毎回 O(n) offsets 再構築 | `docx/src/scroll-viewer.ts:491` | S |
| SC13 | 共有画像 LRU が**件数上限(256)のみでバイト無制限** — 画像重い文書で数 GB の GPU bitmap を pin し得る。decoded byte コストで課金し byte budget で evict | `core/src/image/bitmap-image-by-path.ts:42` | S |
| SC14 | devicePixelRatio 変化検知なし — 別 DPI モニタへ移動後もボケたまま。matchMedia(resolution) チェーンの observeDpr(cb) を core に | 3 viewer + 2 scroll-viewer | S |
| SC15 | pptx メディアオーバーレイが**全メディア停止中も無条件 60fps rAF 全スライド再合成** — dirty-driven 化(再生中/hover 中のみ) | `pptx/src/presentation-handle.ts:170` | S |
| SC16 | B2-T2: 表セル段落が**paint 毎に full line layout 再実行**(vAlign セルは毎レンダー2回)— セル段落にも stamp 機構拡張 | `docx/src/renderer.ts:4688` | M |
| SC17 | Archive コンストラクタが WASM 内でファイルを**二重コピーし 2× を恒久 pin**(メモリは縮まない)— 所有権渡し `new(data: Vec<u8>)` へ(wasm-bindgen owned-vector ABI でコピー1回) | 3 parser `lib.rs` | S |
| SC18 | extract_media/extract_image 応答が転送前に**冗長 full copy 2回**(60MB 動画で +55ms/+120MB churn 実測)— glue buffer 直 transfer | 3 `worker.ts` | S |
| SC19 | parse_sheet 毎にシート .rels を最大7回・drawing XML を3回 re-inflate+re-parse — parse_sheet_with へ hoist | `xlsx/parser/src/lib.rs:126-174` | S |
| SC20 | 5箇所の viewer `load()` が**旧エンジンの worker(+pinned WASM メモリ)を orphan** — 旧 engine destroy をスワップに組み込む | 3 viewer + 2 scroll-viewer | S |
| SC21 | (opt-in)N ビューア = N worker + N WASM ヒープ — サムネイルグリッド用に pooled worker / SharedWorker の共有実行コンテキスト | 新規 | M |

### RB — 堅牢性・セキュリティ

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| RB1 | ラスタ画像 decode に**ピクセル寸法キャップなし** — PNG/JPEG 解凍爆弾が zip 512MiB キャップを迂回。ヘッダ sniff(PNG IHDR/JPEG SOFn/GIF/WEBP)で W×H 予算超過を reject/downscale(createImageBitmap resize) | `core/src/image/` 各 decode 入口 | M |
| RB2 | `a:grpSp`・入れ子 `w:tbl`・OMML tree の**再帰深度無制限** — 深い入れ子で WASM trap。共有 DepthGuard(~50-100、Office 自身の上限準拠)で partial tree 返却 | `pptx/parser/src/shape.rs:1770`, `docx/parser/src/parser.rs:5290`, `ooxml-common/src/math.rs:179` | M |
| RB3 | sparkline `<xm:f>` レンジが `vec![None; rows*cols]` を**無検証 dense 確保** — 巨大レンジで multi-GB alloc→abort。セル数キャップ + sparse 構築 | `xlsx/parser/src/lib.rs:1717` | S |
| RB4 | DIB デコーダが**検証前に full RGBA 確保**、W×H 積は ~16GB まで無制限 — 検証を確保の前へ + megapixel 予算 | `core/src/image/dib.ts:45` | S |
| RB5 | canvas backing 寸法をブラウザ上限(~32767px)に**クランプしていない** — 異常アスペクト比で blank canvas。共有 clampCanvasSize + 一度きり warn | `pptx/src/viewer.ts:248` 他 | M |
| RB6 | Rust panic(=trap)後も worker が**汚染済み WASM インスタンスを使い回す** — RuntimeError 検知で self-poison + main 側 auto-respawn。panic メッセージの保存 hook で 'unreachable' を actionable に | 3 `worker.ts`, `core/src/worker/bridge.ts:122` | M |
| RB7 | XML パースエラーに part 名が乗らない + **1枚の破損スライドが deck 全体を fail** — part path 付与 + 破損スライドは空スライド+警告リストに degrade(79/80 枚見えるビューアが正) | `pptx/parser/src/lib.rs:667` 他 | S |
| RB8 | CI に依存監査ゲートなし — cargo audit(または cargo-deny)+ pnpm audit --prod を report-only → gating | `.github/workflows/ci.yml` | S |
| RB9 | ホストページの **CSP 要件が undocumented**(wasm-unsafe-eval / worker-src blob: / img-src blob: / fonts) — README に節を追加 | README | S |
| RB10 | Google Fonts の **base-URL/ミラー hook なし**(エアギャップ/社内環境で all-or-nothing)— fontBaseUrl 変換 hook | `core/src/fonts/google-fonts.ts:49` | S |
| RB11 | zip read が**ファイル申告の非圧縮サイズで事前 reserve**(512MiB まで)— min(size, 1MiB) 初期確保 + read_to_end 成長へ | `ooxml-common/src/zip.rs:106` | S |
| RB12 | 暗号化(MS-OFFCRYPTO)/レガシー .doc/.xls を **CFB シグネチャ(D0 CF 11 E0)で検知**し typed error(`encrypted` / `legacy-binary-format`)へ — 現状は不可解な zip エラー(PD7 とセット。復号対応自体は PD8) | load パス | S |

### PD — プロダクト & エコシステム

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| PD1 | **markdown 変換が全ユーザーの WASM に入っているのに公開 API なし**(`*_to_markdown` export 済み・packages/markdown は private:true)— `toMarkdown()` を3 headless API に追加 + パッケージ公開 + CLI bin。LLM-ingestion 市場(MarkItDown 150k★)への最安・最速の一手 | `packages/markdown`, 3 worker protocol | S |
| PD2 | **@silurus/ooxml-node が private のまま** — CLI は v0 の pptx のみ。docx ページ/xlsx レンジのサムネイル完成 + 公開 + 「20行でプレビューサービス」レシピ。GroupDocs/Aspose が課金している市場 | `packages/node` | M |
| PD3 | **印刷/PDF 出力が存在しない — 最大の table-stakes 欠落にして最大の差別化機会**(本物のレイアウトエンジンからベクター PDF を replay できる唯一のアーキテクチャ)。Phase 1: 高 DPI ラスタ印刷(S)。Phase 2: core に PDF 記録コンテキスト(Canvas2D サブセット: path/fill/gradient/image/text+フォント埋め込み)、node は skia-canvas native PDF | 新規 core モジュール | L |
| PD4 | typed error surface: 全失敗が `new Error(string)`。`OoxmlError{code}`('encrypted'/'legacy-binary-format'/'not-ooxml'/'zip-entry-cap'/'worker-timeout' 等)+ onError/load() 型付け(RB12 を包含) | core + 3 format | M |
| PD5 | 公式フレームワークラッパー不在 — `@silurus/ooxml-react`(components + useViewer hook、StrictMode double-mount 対応)を最初に。npm 検索で競合に流れている | 新パッケージ | M |
| PD6 | 宣言的エントリポイント不在 — `<ooxml-viewer src type>` Web Component + iframe 埋め込み用静的ホストページ | 新規(既存 API の shim) | M |
| PD7 | サイトが**定量的競合比較ゼロ**(実測: per-format WASM 0.77-0.94MB vs Apryse 展開後 265MB/lib 175MB、MIT vs 営業見積もり、client-only vs iframe/公開URL 必須)。出典・日付付き Why ページ + 競合優位でない行(print/PDF/検索/a11y — 塞ぐまで)も正直に載せる | `site/src/components/Pitch.astro` | S |
| PD8 | 暗号化 OOXML の実対応([MS-OFFCRYPTO] §2.3.4 Agile Encryption、SHA-512 spin+AES — WebCrypto で可)— RB12 の検知の次段 | 新規 | L |
| PD9 | **fidelity スコアカード公開** — README 対応表を ECMA-376 § キーの versioned scorecard としてサイトへ(検証可能な spec-faithfulness は競合が偽造できない資産) | site + README | M |
| PD10 | サイト API リファレンスが手書き 245 行・リリース時同期 — rolled-up .d.ts から生成(ts-morph / api-extractor)して drift クラスを殺す | `site/src/lib/api-reference.ts` | M |
| PD11 | STABILITY.md(安定 API / experimental / deprecation 規則 / 1.0 への条件)— 週次 0.x minor が続く今こそ | 新規 | S |
| PD12 | viewer API パリティ: docx `currentPage` vs pptx `slideIndex`、onReady が xlsx のみ、単canvas viewer の zoom/width API 欠如(IX9 とセットで 1.0 前に) | 3 viewer | S |
| PD13 | dev-profile wasm ビルドが **dist に ~6MB の死んだ base64 wasm を再混入**(再現済み)— 3 Cargo.toml に dev プロファイルの omit-default-module-path + publish.yml にインライン検出ゲート | 3 `Cargo.toml`, `publish.yml` | S |
| PD14 | render エラーの運命が3様(docx rethrow / pptx 黙殺 / xlsx unhandled rejection)— scroll viewer の契約(onError → rethrow/console.error)へ統一 | 3 viewer | S |

### QA — テスト & 品質基盤

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| QA1 | **xlsx の CI 描画カバレッジがゼロ**(smoke は pptx/docx のみ)— Layouts 型ストーリー + タブ切替 + ink assertion | `tests/smoke/layouts.spec.ts:9` | S |
| QA2 | worker-vs-main ピクセル等価 spec が **CI 未接続**(同一マシン相対比較なので E9 のフォント可搬性問題と無関係)— smoke job へ | `packages/*/tests/visual/worker-equivalence.spec.ts` | S |
| QA3 | **CI が Chromium のみ** — Playwright 同梱の webkit/firefox を smoke に追加(font 非依存 assertion なので E9 と両立) | `tests/smoke/playwright.config.ts` | S |
| QA4 | CI がレイアウト回帰を検出不能(assertion は「400 サンプル中非白 ≥3 画素」のみ)— **同一 runner 内 merge-base vs HEAD の差分 VRT**(両側を同環境でラスタライズ = E9 尊重) | 新 CI job | L |
| QA5 | VRT fidelity スコアが console 出力のみ・20% 失敗天井 — **数値 ratchet を commit**(~0.5pt 低下で fail、更新は明示的に) | `packages/*/tests/visual/visual.spec.ts:15` | S |
| QA6 | **ファジング 0 件**(untrusted zip+XML × panic=abort)— cargo-fuzz workspace: parse_*_native / handle-lifecycle(open→parse_sheet/extract_image)/ rels::resolve_target。既存サンプルで corpus seed | 新規 `fuzz/` | M |
| QA7 | property-based テストなし — proptest: resolve_target(archive root 脱出不能・冪等)、fast-check: preset evaluator(全プリセット×ランダム adj)・number-format | ooxml-common, core | M |
| QA8 | カバレッジ計測が皆無 — vitest --coverage(v8)+ cargo llvm-cov を report-only で CI へ、1ヶ月後に pure module へ floor 設定 | CI | S |
| QA9 | bench スクリプトの結果を**誰も消費していない** — merge-base vs HEAD 同一 runner の perf job(nightly or label-gated)+ 閾値 | `packages/node/src/bench-*.mjs` | M |
| QA10 | **Word ground-truth ループが maintainer のマシンにしか存在しない**(#698 も contributor 検証もブロック)— 自作 fixture 生成スクリプト(再配布可)+ Word PDF を一度 commit + bbox-diff runner を CI へ | `packages/docx/src/*.test.ts` の定数群 | L |
| QA11 | **再配布可能な公開 conformance corpus なし**(fidelity 網が private サンプル依存 — Strict 未対応が出荷された遠因)— 自作+ライセンスクリーンな corpus を tests/corpus/ に整備し smoke を通す | 新規 | L |
| QA12 | ast-grep ルールが1件のみ — (1) parser/common での `.unwrap()` 禁止 tripwire (2) `as unknown as` の allowlist 外禁止 (3) 実測定数への spec § コメント必須、を rule 化 | `rules/` | S |
| QA13 | Rust↔TS ワイヤ契約が手書き同期(field drift は実行時にしか割れない)— ts-rs で .d.ts 生成 or 最低限 golden-JSON round-trip テスト | 全 parser types | M |
| QA14 | **MathJax(Apache-2.0)の帰属表記が strip されたまま MIT-only リポジトリで出荷** — esbuild legalComments + THIRD_PARTY_NOTICES(STIX フォント・主要 Rust crate 含む) | `packages/core/build/build-mathjax.mjs` | S |

### AR — アーキテクチャ衛生

| # | 指摘 | 場所 | 工数 |
|---|------|------|------|
| AR1 | docx `parser.rs` 10,601 行(うち ~5,030 行 inline test)— pptx 8-module house pattern で section/paragraph/drawing/shape/table + test 随伴の機械的分割(verbatim + golden sha256) | `docx/parser/src/parser.rs` | M |
| AR2 | docx `renderer.ts` 7,777 行に4エンジン同居 — paginate.ts / table-render.ts / notes.ts / anchor-render.ts へ物理分割(#700 の手法踏襲、B2 相境界に沿う) | `docx/src/renderer.ts` | M |
| AR3 | TS strictness が `strict: true` 止まり(noUncheckedIndexedAccess off、`as unknown as` 90箇所)— base config で有効化しパッケージ毎 burn-down(xlsx の sparse Record と docx line 配列が実バグの温床) | `tsconfig.base.json` | L |
| AR4 | pptx worker の init 失敗で **load() が永久 hang**(ready-handshake が request timeout の外)— docx/xlsx パターン(initPromise 後置キュー)へ統一 | `pptx/src/worker.ts:31` | S |
| AR5 | worker エラー伝播が message string に平坦化(JS stack 喪失、Rust panic は 'unreachable')— {message, stack, name} + bridge 側 Error 再構築 + panic hook でメッセージ保存(RB6 とセット) | 3 `worker.ts`, `bridge.ts` | S |

---

## 改善計画(フェーズ構成)

原則は前次計画を踏襲: 1関心=1commit、レンダリングに触る修正はローカル `pnpm build:wasm && pnpm vrt` 必須、
参照画像更新はユーザー承認のみ、perf 項目は PR にベンチ前後比較添付、横断修正は core/ooxml-common を含む協調 PR 可。

### Phase 1 — 即効の正しさ・安全網・クイックウィン(全 S/M、~25項目)

ユーザー可視の正しさバグと、以降の全フェーズを守る網を先に張る。

- **正しさ(S)**: CH1 CH2 CH3 CH4 / XL1 XL2 / XF8 XF11 XF12 / WD1 WD8 WD9 / PP2 / CH11
- **Strict 名前空間(M)**: XF1(3フォーマット横断・共通述語)
- **セキュリティ(S)**: RB3 RB4 RB11 RB12 / RB8(audit ゲート)
- **worker/ライフサイクル(S)**: SC17 SC18 SC20 / AR4 / PD14
- **CI 網(S)**: QA1 QA2 QA3 QA5 QA12 / PD13
- **コンプライアンス(S)**: QA14(MathJax NOTICE)
- **プロダクト即効(S)**: PD1(markdown 公開)/ PD7(Why ページ)/ PD11 / README EMF 行是正(XF13 の註)

### Phase 2 — チャートエンジンの制覇

看板機能を「業界最良」に。**CH5(パース統一)を最初に**やり、以降の全修正を1箇所に効かせる。

- CH5 → CH6 CH7 CH8 CH9 CH10 → CH12(docx チャート)→ CH13 CH14 → CH15(chartEx 段階導入)
- 検証: chart 系 VRT + 数値 oracle(既存 preset-parity の手法)。CH12/CH15 は新規参照画像 = ユーザー承認とセット

### Phase 3 — DrawingML 横断パリティと埋め込みフォント

「1フォーマットにしかない実装」を共有層に揃える。AGENTS.md 横断原則の総仕上げ。

- 共有層: XF2 XF3 XF4 XF5 XF6 XF7 XF10 XF13(EMF 完成)
- フォント: XF14(埋め込みフォント docx+pptx — 企業文書忠実度の最大レバー)
- 縦書き: XF9(段階実装、JA 需要)
- 検証: 各項目とも「pptx の既存出力 byte-stable + docx/xlsx が新獲得」を oracle に

### Phase 4 — フォーマット固有フィデリティ深化

- docx: WD2 WD3 WD4 WD5 WD6 WD7(w:kern/spacing は折返し変化 = VRT 承認前提)
- pptx: PP1 PP3(S/M 段階まで)PP4
- xlsx: XL3(+QA の corpus テーブル)XL4 XL5
- OMML: WD9 残り
- 検証: fixture → Word/Excel/PowerPoint ground-truth ループ(QA10 と並走)

### Phase 5 — インタラクション層とアクセシビリティ

競合対比で最も見える投資。IX1(リンク)と IX2(検索)がヘッドライン。

- IX1(sanitizer 必須)→ IX2 → IX3 IX4 IX5(a11y 三点セット)→ IX6 IX7(worker パリティ)→ IX8 IX9(ズーム統一)→ IX10 IX11 IX12 IX13
- 検証: smoke へのインタラクションテスト追加(クリック/検索/キーボード)、axe-core による a11y 自動チェック導入

### Phase 6 — スケール性能(100万行 xlsx / 500ページ docx)

- xlsx 三段ロケット: SC1(ワイヤ正規化)→ SC2(streaming parse)→ SC3(窓配送)。各段でベンチ添付(bench-handle 拡張)
- xlsx 描画: SC4 SC5 SC6 SC7 / XL6(C12 一本化 — 正しさも兼ねる)
- docx: SC8 → SC9(incremental pagination)→ SC10 SC11 SC16
- 共通: SC12 SC13 SC14 SC15 SC19 / SC21(opt-in worker pool)
- 検証: 合成 200k 行 fixture + 実測目標(first paint < 1s / シート切替 < 100ms / メモリ < 3×file)を PR に明記

### Phase 7 — 堅牢性: 悪意ファイルに沈まない

- RB1 RB2 RB5 RB6 RB7 / AR5
- QA6(ファジング)→ 出た crash を corpus 化 → QA7(property tests)
- 目標: 「任意バイト列で panic/hang/OOM しない」を fuzz 24h クリーンで裏付け

### Phase 8 — プロダクト展開 【全項目見送り(2026-07-03 ユーザー決定)】

ニーズが実証されていない段階でのプロダクト展開は行わない(PD7 も同判断で見送り)。需要が出た時点で個別に再検討。PD1 / PD11 は Phase 1 に存続。

- ~~PD2(node 公開)→ PD3(PDF: ラスタ印刷 → ベクター記録コンテキスト)→ PD5(react)→ PD6(Web Component)~~
- ~~PD4(typed errors)→ PD8(暗号化対応)~~
- ~~PD9(スコアカード)PD10(API ref 生成)~~

### 継続トラック(フェーズ非依存・機会があれば)

- AR1 AR2(モジュール分割 — 触る PR のついでに境界を切る、#700 方式)
- AR3(strictness burn-down)
- QA4(差分 VRT)QA8 QA9 QA10 QA11 QA13
- SC21 / PD12

### 推奨 PR 分割(Phase 1 の目安)

1. `fix/chart-correctness` — CH1 + CH2 + CH3 + CH4 + CH11(VRT 必須)
2. `fix/xlsx-date-general` — XL1 + XL2
3. `fix/strict-namespaces` — XF1(3パーサー + 共通述語 + Strict fixture)
4. `fix/drawingml-small` — XF8 + XF11 + XF12 + PP2
5. `fix/docx-dropped-runs` — WD1 + WD8 + WD9
6. `sec/parser-hardening-s` — RB3 + RB4 + RB11 + RB12
7. `perf/worker-lifecycle` — SC17 + SC18 + SC20 + AR4 + PD14
8. `ci/coverage-gaps` — QA1 + QA2 + QA3 + QA5 + QA12 + PD13 + RB8
9. `chore/mathjax-notice` — QA14
10. `feat/markdown-public` — PD1(+ site の Why ページ PD7 は別 PR)

---

## 検証成果物の所在

レビューの生データ(161指摘の全文 evidence/proposal/検証ノート)はセッション成果物として保存済み。
本文書の各項目は実装着手時に現物(行番号・定数)を必ず再確認すること(レビュー時点 = main 1610253)。
