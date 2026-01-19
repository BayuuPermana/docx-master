# Product Requirements Document (PRD): Word Master MCP Server

## 1. Project Overview

**Project Name:** Word Master MCP Server

**Technology Stack:** TypeScript, Bun, Model Context Protocol (MCP)

**Goal:** Create a bridge between the Gemini CLI and local Microsoft Word (.docx) files, allowing users to read, create, and precisely format content (including styles, tables, and layouts) using natural language commands.

## 2. Problem Statement

Users of the Gemini CLI currently lack a native, high-performance way to manipulate professional documents (.docx) on their local machines. Existing solutions are either cloud-based or lack fine-grained control over Word's core formatting features like styling, tables, and page layouts. There is a need for an "agentic" tool that understands the full Word object model.

## 3. Target Audience

*   **Developers:** Automating technical documentation with specific company branding/styling.
*   **Content Creators:** Drafting formatted articles, whitepapers, or reports from AI sessions.
*   **Power Users:** Managing complex document structures and layouts through a terminal.

## 4. Technical Requirements

### 4.1. Core Stack

*   **Runtime:** Bun (Native TS support and high-speed execution).
*   **Protocol:** MCP SDK (Tool-use communication standard).
*   **Document Processing:**
    *   `mammoth`: High-level text extraction.
    *   `docx`: Full-feature document generation engine (Handles styles, tables, headers/footers).
*   **Validation:** `zod` for strict runtime typing of complex formatting objects.

### 4.2. Infrastructure

*   **Transport:** Standard Input/Output (stdio).
*   **File System:** `node:fs/promises`.

## 5. Functional Requirements (The "Tools")

The server must expose tools that support the following formatting categories:

### 5.1. Text & Font Styling

*   **Headings:** Support for hierarchical levels (Heading 1-6).
*   **Fonts:** Control over font family, size, color, bold, italic, and underline.
*   **Alignment:** Left, center, right, and justified text.

### 5.2. Structural Elements

*   **Tables:** Creation of N-column tables with header rows, borders, and cell shading.
*   **Lists:** Support for nested bulleted and numbered lists.
*   **Images:** Ability to insert local images with specified dimensions and wrapping styles.

### 5.3. Layout & Referencing

*   **Margins & Orientation:** Control over page margins (top, bottom, left, right) and orientation (portrait/landscape).
*   **Headers & Footers:** Support for page numbers and recurring text.
*   **References:** Support for hyperlinks, bookmarks, and basic footnotes.

## 6. Detailed Tool Specs (Expanded)

### 6.1. `create_styled_doc`

*   **Input:** `path`, `title`, `sections[]`.
*   **Section Object:** Includes type (paragraph, table, image), content, and style (font, size, alignment).

### 6.2. `update_layout`

*   **Input:** `path`, `margins: {top, bottom, left, right}`, `orientation`.

### 6.3. `read_doc_structure`

*   **Input:** `path`.
*   **Output:** A JSON-like map of the document's structure (identifying headings, tables, and sections).

## 7. User Stories

*   **Branded Reporting:** "Create a report at ./Q4.docx with a 2cm margin, Arial 12pt font, and a table summarizing our sales data."
*   **Academic Drafting:** "Draft my research paper with a title page, Heading 1 for 'Abstract', and insert a placeholder for a 3x3 table."
*   **Template Filling:** "Read the layout of template.docx and create a new doc using the same margins and header, but change the body text."

## 8. Roadmap & Milestones

### Phase 1: MVP (Completed)
*   [x] Scaffolding and basic Read/Create.

### Phase 2: Core Formatting (Current)
*   [ ] Styles: Implementation of custom font and color schemas.
*   [ ] Tables: dynamic table generation from JSON input.
*   [ ] Images: local image path embedding logic.

### Phase 3: Layout & Advanced Features
*   [ ] Page Setup: Margins, Orientation, and Columns.
*   [ ] Referencing: Hyperlinks and Footnotes.
*   [ ] PDF Export: Integration with conversion tools.

## 9. Success Metrics

*   **Formatting Accuracy:** 100% adherence to specified font/margin parameters in generated files.
*   **CLI Latency:** Execution < 1s for documents under 50 pages.
*   **Documentation:** Clear documentation of the JSON schema required for tables and complex styles.