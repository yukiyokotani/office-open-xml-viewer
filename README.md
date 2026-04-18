# office-open-xml-viewer

A browser-based viewer for Office Open XML documents that renders to an HTML Canvas element.
The parser is written in Rust and compiled to WebAssembly; the renderer uses the Canvas 2D API.

## Feature Support

### PowerPoint (.pptx)

| Category | Feature | Status |
|----------|---------|--------|
| **Slides** | Slide rendering | ✅ |
| | Slide layout / master inheritance | ✅ |
| | Slide size (custom dimensions) | ✅ |
| | Slide background (solid, gradient, image) | ✅ |
| | Slide numbers | ✅ |
| | Notes pages | ❌ |
| | Animations / transitions | ❌ |
| **Element types** | Shapes (`sp`) | ✅ |
| | Pictures (`pic`) | ✅ |
| | Groups (`grpSp`) with nested transforms | ✅ |
| | Connectors (`cxnSp`) | ✅ |
| | Tables (`tbl` in `graphicFrame`) | ✅ |
| | Charts (`c:chart` — bar, waterfall) | ✅ |
| | Charts (line, pie, area, radar, scatter, bubble) | ❌ |
| | SmartArt | ❌ |
| | OLE objects | ❌ |
| | Video / audio | ❌ |
| **Shape geometry** | 130+ preset shapes (`prstGeom`) | ✅ |
| | Custom geometry (`custGeom`) | ✅ |
| | Rotation and flip (flipH / flipV) | ✅ |
| | 3D preset shapes | ❌ |
| **Fills** | Solid fill (`solidFill`) | ✅ |
| | Linear / radial gradient (`gradFill`) | ✅ |
| | No fill (`noFill`) | ✅ |
| | Pattern fill (`pattFill`) | ❌ |
| | Image fill on shapes (`blipFill` in `sp`) | ❌ |
| **Strokes** | Solid line color and width | ✅ |
| | Dash / dot styles | ❌ |
| | Arrow heads | ❌ |
| | Compound / double lines | ❌ |
| **Shape effects** | Drop shadow (`outerShdw`) | ✅ |
| | Inner shadow (`innerShdw`) | ❌ |
| | Glow | ❌ |
| | Reflection | ❌ |
| | Soft edge | ❌ |
| | Bevel / 3D extrusion | ❌ |
| **Text — characters** | Bold, italic | ✅ |
| | Underline | ✅ |
| | Strikethrough | ✅ |
| | Font family, size, color | ✅ |
| | Superscript / subscript | ❌ |
| | Hyperlinks | ❌ |
| | Text shadow / outline effects | ❌ |
| **Text — paragraphs** | Horizontal alignment (left / center / right / justify) | ✅ |
| | Vertical anchor (top / center / bottom) | ✅ |
| | Line spacing (`spcPct`, `spcPts`) | ✅ |
| | Space before / after paragraph | ✅ |
| | Bullet points (character and auto-numbered) | ✅ |
| | Tab stops | ✅ |
| | Indent / margin | ✅ |
| | Vertical text direction | ❌ |
| | Right-to-left text | ❌ |
| **Text — body** | Text padding (insets) | ✅ |
| | normAutoFit (shrink to fit) | ✅ |
| | spAutoFit (expand box) | ✅ |
| | Word wrap / no wrap | ✅ |
| | Text overflow clipping | ✅ |
| **Tables** | Cells, rows, columns | ✅ |
| | Cell merges (horizontal / vertical) | ✅ |
| | Cell borders | ✅ |
| | Cell fills (solid / gradient) | ✅ |
| | Table theme styles | ❌ |
| | Cell diagonal lines | ❌ |
| **Theme** | Scheme colors (dk1/lt1/accent1–6 etc.) | ✅ |
| | Font scheme (`+mj-lt`, `+mn-lt`) | ✅ |
| | lumMod / lumOff color transforms | ✅ (approx) |
| | alpha transparency | ✅ |

---

### Word (.docx) — Planned

| Category | Feature | Status |
|----------|---------|--------|
| **Document** | Page rendering | 🔜 Planned |
| | Page size and margins | 🔜 Planned |
| | Section breaks | 🔜 Planned |
| | Headers / footers | 🔜 Planned |
| **Text** | Paragraphs | 🔜 Planned |
| | Bold, italic, underline, strikethrough | 🔜 Planned |
| | Font family, size, color | 🔜 Planned |
| | Superscript / subscript | 🔜 Planned |
| | Hyperlinks | 🔜 Planned |
| **Formatting** | Paragraph alignment | 🔜 Planned |
| | Line spacing | 🔜 Planned |
| | Indents and tab stops | 🔜 Planned |
| | Paragraph styles (Heading 1–6, Normal, etc.) | 🔜 Planned |
| | Lists (bullet and numbered) | 🔜 Planned |
| **Elements** | Tables | 🔜 Planned |
| | Images | 🔜 Planned |
| | Text boxes | 🔜 Planned |
| | Drawing shapes | 🔜 Planned |
| **Advanced** | Track changes | ❌ Not planned |
| | Comments | ❌ Not planned |
| | Footnotes / endnotes | 🔜 Planned |
| | Mail merge fields | ❌ Not planned |

---

### Excel (.xlsx) — Planned

| Category | Feature | Status |
|----------|---------|--------|
| **Workbook** | Multiple sheets | 🔜 Planned |
| | Sheet names | 🔜 Planned |
| **Cells** | Text values | 🔜 Planned |
| | Number values | 🔜 Planned |
| | Boolean values | 🔜 Planned |
| | Formula results (display only) | 🔜 Planned |
| | Dates | 🔜 Planned |
| **Formatting** | Bold, italic, underline, strikethrough | 🔜 Planned |
| | Font family, size, color | 🔜 Planned |
| | Cell background color | 🔜 Planned |
| | Borders | 🔜 Planned |
| | Horizontal / vertical alignment | 🔜 Planned |
| | Text wrapping | 🔜 Planned |
| | Number formats | 🔜 Planned |
| **Structure** | Merged cells | 🔜 Planned |
| | Frozen panes | 🔜 Planned |
| | Row / column sizing | 🔜 Planned |
| | Hidden rows / columns | 🔜 Planned |
| **Elements** | Images | 🔜 Planned |
| | Charts | 🔜 Planned |
| | Sparklines | ❌ Not planned |
| **Advanced** | Pivot tables | ❌ Not planned |
| | Conditional formatting | 🔜 Planned |
| | Data validation dropdowns | ❌ Not planned |
| | Comments / notes | ❌ Not planned |
