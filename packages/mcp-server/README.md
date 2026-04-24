# ooxml-mcp-server

A [Model Context Protocol](https://modelcontextprotocol.io/) (MCP) server that exposes the OOXML Rust parsers as structured query tools for AI agents.

Agents can ask questions like *"What formulas are in Sheet1?"* or *"Extract all text from this presentation"* without writing any parsing code themselves.

## Installation

**Prerequisites:** [Rust toolchain](https://rustup.rs/) (stable, 1.75+)

> **Note:** This is a pure-Rust binary. You do **not** need Node.js, npm, or pnpm — those are only required for the browser renderer packages in the rest of the monorepo.

```bash
# 1. Clone the repository
git clone https://github.com/yukiyokotani/office-open-xml-viewer.git
cd office-open-xml-viewer

# 2. Install the binary (builds the Rust parsers automatically)
cargo install --path packages/mcp-server
```

The binary lands in `~/.cargo/bin/ooxml-mcp-server`. Confirm it is on your `PATH`:

```bash
which ooxml-mcp-server   # → /Users/you/.cargo/bin/ooxml-mcp-server
```

If `~/.cargo/bin` is not on `PATH`, add this to your shell profile (`~/.zshrc`, `~/.bashrc`, …):

```bash
export PATH="$HOME/.cargo/bin:$PATH"
```

---

## Available Tools

### xlsx

| Tool | Parameters | Description |
|------|-----------|-------------|
| `xlsx_parse` | `path` | Workbook overview: sheet names, IDs |
| `xlsx_get_sheet_names` | `path` | Array of sheet names only |
| `xlsx_get_sheet_dimensions` | `path`, `sheet` | Max row and column in a sheet |
| `xlsx_get_cell_range` | `path`, `sheet`, `range` | Cell values and formulas for a range (e.g. `"A1:C10"`) |
| `xlsx_get_formulas` | `path`, `sheet` | All formula cells with their cached values |

`sheet` accepts either a sheet name (`"Sheet1"`) or a 0-based index (`"0"`).

### docx

| Tool | Parameters | Description |
|------|-----------|-------------|
| `docx_extract_text` | `path` | Plain text from the entire document |
| `docx_get_structure` | `path` | Paragraph/table structure with style info |
| `docx_get_tables` | `path` | All tables with cell-by-cell contents |

### pptx

| Tool | Parameters | Description |
|------|-----------|-------------|
| `pptx_get_slides` | `path` | Slide count and each slide's title |
| `pptx_extract_text` | `path`, `slide_index?` | Text from all slides or a specific one (0-based) |
| `pptx_get_slide_structure` | `path`, `slide_index` | All elements with position, size, and text |

All `path` parameters accept absolute paths.

---

## Configuration

### Claude Code

Add to your project's `.mcp.json` (version-controlled, shared with the team):

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

Or to your personal global config `~/.claude.json` (applies to every project):

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

After adding the config, run `/mcp` inside Claude Code to confirm the server is connected.

**Usage example:**

```
> What sheets are in /path/to/budget.xlsx?

[Claude calls xlsx_get_sheet_names with {"path": "/path/to/budget.xlsx"}]
→ ["Summary", "Q1", "Q2", "Q3", "Q4"]
```

---

### GitHub Copilot (VS Code)

Add `.vscode/mcp.json` to your workspace root:

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

Then in VS Code, open the Command Palette (`⇧⌘P`) → **"MCP: List Servers"** to confirm the server appears as *running*.

> **Note:** MCP tools are only available in **Agent mode** (`@agent` in the Copilot Chat panel). They are not accessible in Ask or Edit modes.

**Usage example (Copilot Chat, Agent mode):**

```
@agent Summarise the formulas in Sheet1 of @/path/to/model.xlsx
```

Copilot will call `xlsx_get_formulas` and present the results.

---

### Codex CLI (OpenAI)

Add to `~/.codex/config.toml` (global, all projects):

```toml
[mcp_servers.ooxml]
command = "ooxml-mcp-server"
args = []
```

Or to `.codex/config.toml` in the project root (project-scoped):

```toml
[mcp_servers.ooxml]
command = "ooxml-mcp-server"
args = []
```

Restart Codex after editing the config. Use `codex mcp list` to verify the server is registered.

**Usage example:**

```bash
codex "Extract all text from slides 1–3 of /path/to/deck.pptx"
```

Codex will call `pptx_extract_text` with `slide_index` set appropriately.

---

## Full path (if not on PATH)

If `~/.cargo/bin` is not on your `PATH`, use the absolute path to the binary:

```bash
# Find the binary
which ooxml-mcp-server         # or:
echo ~/.cargo/bin/ooxml-mcp-server
```

Then replace `"ooxml-mcp-server"` with the full path in any config above:

```json
"command": "/Users/you/.cargo/bin/ooxml-mcp-server"
```

```toml
command = "/Users/you/.cargo/bin/ooxml-mcp-server"
```

---

## Building from source (without installing)

```bash
cargo build --release -p ooxml-mcp-server
# binary at: target/release/ooxml-mcp-server
```

Use the path `target/release/ooxml-mcp-server` in your MCP config.
