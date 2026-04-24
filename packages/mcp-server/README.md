# ooxml-mcp-server

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that lets AI agents read Excel, Word, and PowerPoint files without any additional code.

---

## Quick start

### Step 1 — Install Rust

Skip this if `rustc --version` already prints a version number.

```bash
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
# Follow the on-screen prompt, then open a new terminal tab.
```

### Step 2 — Install the MCP server

No need to clone the repository. Run this single command:

```bash
cargo install --git https://github.com/yukiyokotani/office-open-xml-viewer.git \
  --package ooxml-mcp-server
```

The first run downloads and compiles everything (~2 minutes). The binary is placed in `~/.cargo/bin/ooxml-mcp-server`.

**Check it is on your PATH:**

```bash
which ooxml-mcp-server
# expected: /Users/you/.cargo/bin/ooxml-mcp-server
```

If the command is not found, add this line to your shell profile (`~/.zshrc` or `~/.bashrc`) and open a new terminal:

```bash
export PATH="$HOME/.cargo/bin:$PATH"
```

### Step 3 — Configure your AI client

Pick the client you use and follow the instructions in the section below.

---

## Configuration

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

---

### GitHub Copilot (VS Code)

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

**Try it:**

```bash
codex "Show me all formulas in Sheet1 of /Users/me/Documents/model.xlsx"
```

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

## Available tools

### xlsx (Excel)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `xlsx_parse` | `path` | All sheet names and IDs |
| `xlsx_get_sheet_names` | `path` | Sheet name list |
| `xlsx_get_sheet_dimensions` | `path`, `sheet` | Number of rows and columns |
| `xlsx_get_cell_range` | `path`, `sheet`, `range` | Cell values and formulas for a range like `"A1:C10"` |
| `xlsx_get_formulas` | `path`, `sheet` | Every formula cell with its cached value |

`sheet` can be a name (`"Sheet1"`) or a 0-based index (`"0"`).

### docx (Word)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `docx_extract_text` | `path` | All text as plain string |
| `docx_get_structure` | `path` | Paragraph and table structure with style info |
| `docx_get_tables` | `path` | All tables with each cell's text |

### pptx (PowerPoint)

| Tool | Parameters | What it returns |
|------|-----------|-----------------|
| `pptx_get_slides` | `path` | Slide count and each slide's title |
| `pptx_extract_text` | `path`, `slide_index?` | Text from all slides, or one slide (0-based index) |
| `pptx_get_slide_structure` | `path`, `slide_index` | All shapes with position, size, and text |

All `path` parameters require absolute paths (e.g. `/Users/me/Documents/file.xlsx`).
