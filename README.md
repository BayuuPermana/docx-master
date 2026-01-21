# Docx Master (v1.0.0)

A high-fidelity Model Context Protocol (MCP) server for local Microsoft Word (.docx) manipulation. This server uses a "Surgical Editing" workflow to modify existing documents without losing metadata or corrupting the internal XML structure.

## Features

*   **Surgical Text Replace**: Replaces text patterns across multiple runs while preserving exact formatting.
*   **Media Injection**: Surgically adds images to `word/media/` and updates relationships.
*   **Style-Aware Insertion**: Clones existing paragraph properties to ensure new content matches the document's aesthetic.
*   **High-Fidelity Inspection**: Dissects `.docx` files into JSON blueprints, capturing fonts, spacing, and images.

## Tools

### `read_doc_structure`
Fast text and heading extraction.

### `inspect_doc_formatting`
Forensic dissection of formatting and media extraction.

### `surgical_text_replace`
The flagship tool. Replaces methodology (e.g., SVM -> Random Forest) across hundreds of pages with 100% style retention.

### `surgical_add_image`
Injects an image into a specific paragraph index using raw OOXML injection.

### `surgical_insert_paragraph`
Inserts a new paragraph that inherits properties from a template paragraph.

## Usage Scenarios

### 🕵️ Scenario: The "Executive Identity Theft"
**Goal**: You have a 50-page legal contract (`Contract_V1.docx`) with complex branding, custom fonts, 15 different image signatures, and nested tables. You need to change the name of the "Contractor" globally, but if the file size changes or the font shifts by 1 pixel, the lawyers will know you've been in there.

1.  **Phase A: DNA Extraction**: Use `inspect_doc_formatting` to map every hidden paragraph and legacy VML image.
2.  **Phase B: Surgical Replacement**: Use `surgical_text_replace` to swap names across XML nodes while leaving font sizes and bolding untouched.
3.  **Phase C: Image Injection**: Use `surgical_add_image` to add a high-res seal binary directly to the `media/` folder and link it via a unique `rId`.
4.  **Phase D: Verification**: Run the `molecular_diff` script to prove that the file size and internal structure match the original with 99.9% accuracy.

## Capabilities Comparison

| Capability | Docx Master (Pickle Rick Mode) | Standard Libraries (Jerry Mode) |
| :--- | :--- | :--- |
| **Integrity** | Uses the **Skinwalker Method**: Modifies the original host binary. | Regenerates the file from scratch, breaking fonts. |
| **Legacy Support** | Maps both modern `w:drawing` and legacy `w:pict` tags. | Blind to legacy VML images (they just disappear). |
| **Fonts** | Preserves obfuscated `.odttf` font streams. | Strips them out, resetting document to Calibri. |
| **Metadata** | Keeps 100% of undo history and custom properties. | Wipes the metadata clean. |
| **Fidelity** | Measured via **Molecular Diff** (internal organ size). | "Looks okay to me" (narrator: it wasn't). |

## Installation

1.  **Prerequisites**: [Bun](https://bun.sh/) must be installed.
2.  **Setup**:
    ```bash
    git clone https://github.com/BayuuPermana/docx-master.git
    cd docx-master
    bun install
    bun run build
    ```

## Usage (Claude Desktop / Gemini CLI)

Add this to your MCP configuration file:

```json
"docx-master": {
  "command": "bun",
  "args": ["run", "/path/to/your/project/docx-master/dist/index.js"]
}
```
*Note: Replace `/path/to/your/project/` with the actual absolute path on your system.*

---

## Technical Appendix: The OOXML Methodology

To ensure high-fidelity editing, **Docx Master** treats `.docx` files as a compressed ZIP archive of XML files (the Open Packaging Conventions standard).

### Key Concepts
*   **The "Run" Concept**: Text is divided into `<w:r>` (runs) based on formatting changes. Our surgical tool maps patterns across these runs to avoid breaking font styles.
*   **Relationship Mapping**: Images are not stored in the text; they are referenced by ID and mapped in the `_rels` folder. Our media injector manages these mappings manually.
*   **Style Inheritance**: Paragraphs inherit styles from `styles.xml`. Our insertion tools clone these properties (`w:pPr`) to maintain document integrity.
*   **Buffer-Read-Edit-Write**: We use an in-memory ZIP buffer to modify specific XML parts and re-pack the archive, ensuring no metadata (like Track Changes or Themes) is lost.