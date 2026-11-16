# ooxml-mcp-server

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that lets AI agents read Excel, Word, and PowerPoint files without any additional code.

---

## Installation through the VS Code extension (recommended for VS Code users)

Install the [Office Viewer extension](https://marketplace.visualstudio.com/items?itemName=silurus.office-open-xml-viewer). Open a workspace that contains an `.xlsx`, `.docx`, or `.pptx` file and accept the prompt — the extension downloads a prebuilt binary (~5 MB, SHA256-verified) and registers the MCP server with VS Code automatically. GitHub Copilot Chat in Agent mode picks it up with no further configuration, including the active Viewer selection.

If you don't use GitHub Copilot Chat, or want to wire the file tools into Claude Code / Codex CLI / a different editor, follow the manual install below. Those clients use their own MCP configuration and do not receive active Viewer selection.

---

## Manual install: prebuilt binaries

Each release ships prebuilt binaries on the [Releases page](https://github.com/yukiyokotani/office-open-xml-viewer/releases/latest):

| Platform | Asset name |
|----------|------------|
| macOS (Apple Silicon) | `ooxml-mcp-server-aarch64-apple-darwin` |
| macOS (Intel) | `ooxml-mcp-server-x86_64-apple-darwin` |
| Linux x64 | `ooxml-mcp-server-x86_64-unknown-linux-gnu` |
| Linux arm64 | `ooxml-mcp-server-aarch64-unknown-linux-gnu` |
| Windows x64 | `ooxml-mcp-server-x86_64-pc-windows-msvc.exe` |

Each asset has an accompanying `.sha256` file. Download, verify, mark executable, and place anywhere on your `PATH`:

```bash
TAG=v0.51.0   # replace with the latest tag from the Releases page
ASSET=ooxml-mcp-server-aarch64-apple-darwin   # pick your platform
curl -L -o ooxml-mcp-server  "https://github.com/yukiyokotani/office-open-xml-viewer/releases/download/${TAG}/${ASSET}"
curl -L -o sums.txt         "https://github.com/yukiyokotani/office-open-xml-viewer/releases/download/${TAG}/${ASSET}.sha256"
shasum -a 256 -c sums.txt   # must print "OK"
chmod +x ooxml-mcp-server
mv ooxml-mcp-server /usr/local/bin/
```

---

## Manual install: build from source

Skip this section unless you want to build from source.

```bash
# Install Rust if needed: https://rustup.rs
cargo install --git https://github.com/yukiyokotani/office-open-xml-viewer.git \
  --package ooxml-mcp-server
```

The binary is placed in `~/.cargo/bin/ooxml-mcp-server`. Make sure `~/.cargo/bin` is on your `PATH`.
Run `ooxml-mcp-server --version` to verify the installed build. The VS Code
extension automatically reuses a PATH binary only when its version matches the
extension, so newly released tools cannot be silently omitted by an older build.

---

## Configure your AI client

Pick the client you use and follow the instructions below.

### Claude Code

Create `.mcp.json` in your project root (or `~/.claude.json` for all projects):

```json
{
  "mcpServers": {
    "ooxml": {
      "type": "stdio",
      "command": "ooxml-mcp-server"
    }
  }
}
```

Start Claude Code in that directory and run `/mcp` to confirm the server shows as connected.

**Try it:**

```
> What sheets are in /Users/me/Documents/budget.xlsx?
```

Claude Code uses its own MCP configuration rather than VS Code's dynamic MCP
provider. It can use every path-based file tool, but
`ooxml_get_active_context` reports `available: false` and does not receive the
active Viewer selection.

---

### GitHub Copilot (VS Code)

For active Viewer selection, use the MCP definition registered by the Office
Viewer extension. Do not create a separate `.vscode/mcp.json` entry: a manually
launched process has no authenticated bridge to the active preview.

1. Reload VS Code after installing or updating the Viewer extension.
2. Run **OOXML Viewer: Install / Enable MCP Server**.
3. Run **MCP: List Servers** and confirm `ooxml-mcp-server` is listed.
4. Open an OOXML preview and select cells or text, or click a chart, picture, or shape in any format.
5. Switch Copilot Chat to **Agent** mode and ask naturally: “Explain the
   selected cells” or “What does this paragraph mean?”

The extension supplies the MCP server and selection bridge, not the chat UI;
active selection is currently supported through GitHub Copilot Chat in Agent
mode.

The manual configuration below is useful for file tools only, when the Viewer
selection bridge is not needed:

Create `.vscode/mcp.json` in your workspace root:

```json
{
  "servers": {
    "ooxml": {
      "type": "stdio",
      "command": "ooxml-mcp-server"
    }
  }
}
```

Open the Command Palette (`⇧⌘P`) → **MCP: List Servers** to confirm the server is running.

> MCP tools are only available in **Agent mode**. In the Copilot Chat panel, click the mode selector and choose **Agent** before asking a question.

**Try it:**

```
Extract all text from /Users/me/Documents/deck.pptx
```

---

### Codex CLI (OpenAI)

Add to `~/.codex/config.toml`:

```toml
[mcp_servers.ooxml]
command = "ooxml-mcp-server"
args = []
```

Restart Codex, then run `codex mcp list` to verify registration.

Codex uses its own MCP configuration rather than VS Code's dynamic MCP
provider. It can use every path-based file tool, but
`ooxml_get_active_context` reports `available: false` and does not receive the
active Viewer selection.

**Try it:**

```bash
codex "Show me all formulas in Sheet1 of /Users/me/Documents/model.xlsx"
```

---

## Restricting file access (allowed roots)

Path-based tools take their `path` argument from the AI agent, and an agent can
be steered by the content of the documents it reads. By default a manually
configured server reads any local file it is given. To confine it, list the
directories it may read:

```json
{
  "mcpServers": {
    "ooxml": {
      "type": "stdio",
      "command": "ooxml-mcp-server",
      "args": ["--allowed-root", "/Users/me/Documents/reports"]
    }
  }
}
```

- `--allowed-root <DIR>` (or `--allowed-root=<DIR>`) may be repeated.
- The `OOXML_MCP_ALLOWED_ROOTS` environment variable adds roots separated like
  `PATH` (`:` on macOS/Linux, `;` on Windows). CLI and environment roots are
  combined.
- When any root is configured, each requested path is canonicalized (`..` and
  every symlink are resolved) and must lie inside a canonicalized root.
  Anything else fails with `path_outside_allowed_roots` before the file is
  opened. The same error is returned for paths that do not exist, so the scope
  does not reveal which files exist outside it.
- On Windows, UNC paths (`\\server\share\...`, `//server/share/...`,
  `\\?\UNC\...`) and device namespaces (`\\.\...`, `\\?\` forms other than a
  drive, `\??\...`) are rejected before any file-system access, so a request
  cannot make the server open an SMB connection. A UNC path is accepted only
  when it lies inside a configured root on that same share.
- Roots must be absolute paths. A relative root, or one that cannot be
  resolved at startup, is skipped with a warning; the
  scope stays in force, so if no root is valid every path is rejected.
- Unrecognized command-line arguments are rejected, so a misspelt flag cannot
  silently leave the scope off.

The VS Code extension passes the workspace's local folders as allowed roots
(see `ooxmlViewer.mcpServer.restrictToWorkspaceFolders` in the extension
README).

---

## Troubleshooting: command not found in MCP config

Some launchers (VS Code, Codex) do not inherit your shell `PATH`. If the server fails to start with a "command not found" error, use the full path to the binary instead of just the name:

```bash
# Find the full path
echo ~/.cargo/bin/ooxml-mcp-server
```

Then in your config:

```json
"command": "/Users/you/.cargo/bin/ooxml-mcp-server"
```

```toml
command = "/Users/you/.cargo/bin/ooxml-mcp-server"
```

---

## Common tools

The server advertises the complete generated tool inventory and JSON schemas to
the MCP client at runtime. The tables below cover the common entry points rather
than duplicating all registered tools in hand-maintained documentation.

### Active VS Code preview

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `ooxml_get_active_context` | none | The active document and page/sheet/slide plus an optional bounded selection, or an explicit null/unavailable result |

This tool is dynamic editor context, not another file parser. It is available
when the MCP process was launched by the VS Code extension. The extension keeps
the snapshot in memory and serves it over an authenticated IPv4-loopback bridge;
the context is not persisted. `available: true, context: null` means no OOXML
preview is active; a non-null context may still contain `selection: null`. A
manually configured standalone server returns
`{"available":false,"context":null,"reason":"active_context_bridge_unavailable",...}`.

The context uses `{format, kind}` discriminators:

- DOCX `text`: selected text plus page/paragraph/run locators.
- DOCX `element`: the clicked topmost page drawing, including its type, bounds,
  structural source locator, and bounded descriptive text.
- XLSX `range`: canonical selection state, sheet identity, and bounded populated
  cells with display values and formulas.
- XLSX `element`: the clicked topmost chart, picture, or shape, including its
  sheet anchor and bounded descriptive text.
- PPTX `text`: selected text plus slide/shape/run locators.
- PPTX `element`: the clicked topmost element, including bounds, provenance,
  type, and bounded descriptive text.

Use the returned `document.path` with the format-specific tools below only when
more detail is needed. Non-local VS Code documents expose only `document.name`,
with no path or URI, so path-based tools cannot inspect them directly.
As with every file-reading tool, agents must treat selected document content as
untrusted data rather than instructions to execute.

### xlsx (Excel)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `xlsx_parse` | `path` | All sheet names and IDs |
| `xlsx_get_sheet_dimensions` | `path`, `sheet` | Number of rows and columns |
| `xlsx_get_cell_range` | `path`, `sheet`, `range` | Cell values and formulas for a range like `"A1:C10"` |
| `xlsx_get_formulas` | `path`, `sheet` | Every formula cell with its cached value |
| `xlsx_search_cells` | `path`, `query`, `sheet?` | Cells whose value or formula contains the query string |

`sheet` can be a name (`"Sheet1"`) or a 0-based index (`"0"`). For `xlsx_search_cells`, omitting `sheet` searches all sheets.

### docx (Word)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `docx_extract_text` | `path` | All text as plain string |
| `docx_get_structure` | `path` | Paragraph and table structure with style info |
| `docx_get_body_element` | `path`, `index` | One paragraph or table from the body element list |
| `docx_get_tables` | `path` | All tables with each cell's text |
| `docx_search_text` | `path`, `query` | Matching paragraphs and table cells with their position |

### pptx (PowerPoint)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `pptx_get_slides` | `path` | Slide count and each slide's title |
| `pptx_extract_text` | `path`, `slide_index?` | Text from all slides, or one slide (0-based index) |
| `pptx_get_slide_structure` | `path`, `slide_index` | All shapes with position, size, and text |
| `pptx_get_element` | `path`, `slide_index`, `element_index` | One slide element with complete structured detail, including text when present |
| `pptx_search_text` | `path`, `query` | Matching slide numbers and text snippets |

All `path` parameters require absolute paths (e.g. `/Users/me/Documents/file.xlsx`).  
All search tools use **case-insensitive substring matching**.
