# Word Master MCP Server (docx-master)

> **Auto-Generated Explanation by Pickle Rick**
> *Date: 2026-01-21*

## 1. Project Overview
**Word Master** is a Model Context Protocol (MCP) server designed for surgical manipulation of Microsoft Word (`.docx`) files. Unlike standard libraries that simply "write" files, this server includes a "Surgical Suite" capable of injecting XML directly into existing documents, preserving their original formatting, styles, and quirks.

**Tech Stack:**
-   **Runtime:** Bun (v1.x)
-   **Language:** TypeScript
-   **Transport:** Stdio (Standard Input/Output)

## 2. Capabilities & Tools

The server exposes 7 distinct tools divided into two categories:

### A. Core Operations
These tools handle standard document lifecycle events.
-   **`read_doc_structure`**: Converts `.docx` content to Markdown for easy reading by LLMs. Uses `mammoth`.
-   **`create_styled_doc`**: Generates *new* documents from scratch with support for:
    -   Headers & Footers
    -   Page Layouts (Margins, Orientation)
    -   Tables (with styling)
    -   Images
    -   Paragraph Styles (Headings, Alignment)

### B. The Surgical Suite
These tools modify *existing* documents in-place by manipulating the internal OOXML structure.
-   **`inspect_doc_formatting`**: Ultra-fidelity inspection. Extracts images and returns a detailed JSON map of the document's internal structure (styles, runs, indentation).
-   **`surgical_insert_paragraph`**: Inserts a new paragraph by cloning the styles (`w:pPr`) of an existing template paragraph.
-   **`surgical_text_replace`**: Performs global search-and-replace across text runs while attempting to preserve paragraph-level formatting.
-   **`surgical_add_image`**: Injects a raw XML drawing object (`w:drawing`) into the document and updates relationships.
-   **`cleanup_media`**: Utility to remove temporary media files extracted during inspection.

## 3. Architecture
-   **Entry Point**: `src/index.ts` (Registers the MCP server and tools).
-   **Logic**: `src/tools/*.ts` (Individual tool handlers).
-   **State Management**: `src/utils/EditSession.ts` (Manages the Zip/XML state for surgical operations).

## 4. Usage

### Setup
```bash
bun install
```

### Build
```bash
bun run build
```
Output: `dist/index.js`

### Run (MCP)
Add this to your MCP Client configuration (e.g., Claude Desktop, Gemini CLI):

```json
{
  "mcpServers": {
    "word-master": {
      "command": "bun",
      "args": ["run", "src/index.ts"]
    }
  }
}
```

## 5. Development Notes
-   **Version**: The server identifies itself as v2.2.0.
-   **Limitations**: Recursive table nesting in `create_styled_doc` is currently limited.
