# Word Master MCP Server (v1.1.0)

A high-fidelity Model Context Protocol (MCP) server for local Microsoft Word (.docx) manipulation.

## Features

*   **Read Documents**: Extract text and headings as Markdown.
*   **High-Fidelity Inspection**: Dissect `.docx` files into JSON blueprints, capturing:
    *   **Fonts**: Family, size, color, bold, italic, underline.
    *   **Paragraphs**: Alignment and Spacing (before/after/line).
    *   **Images**: Automatic extraction of embedded media to local storage.
*   **Advanced Creation**: Build documents from JSON blueprints with support for all the above, plus Tables and Page Layouts (margins/orientation).
*   **Media Management**: Dedicated tools for cleaning up extracted media.

## Tools

### `read_doc_structure`
Simple text extraction.

### `inspect_doc_formatting`
Deep-XML inspection. Generates a blueprint for cloning.
*   **path**: Absolute path to `.docx`.
*   *Note*: Extracts images to an `mcp_media` folder in the same directory.

### `create_styled_doc`
Generates a `.docx` from a section/block blueprint.
*   **path**: Absolute output path.
*   **sections**: JSON structure (compatible with `inspect_doc_formatting` output).

### `cleanup_media`
Deletes the `mcp_media` folder.
*   **directory**: Parent directory of the media folder.

## Installation

1.  **Prerequisites**: [Bun](https://bun.sh/).
2.  **Setup**:
    ```bash
    bun install
    ```

## Usage (Claude/Gemini)

```json
"word-master": {
  "command": "bun",
  "args": ["run", "D:/porto/MCP/src/index.ts"]
}
```
