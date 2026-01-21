# Docx Master: The Skinwalker Technique 🥒

This skill defines the **Skinwalker Method** for high-fidelity Microsoft Word (.docx) manipulation. 

## 1. Core Philosophy
Standard document libraries (like `docx` or `python-docx`) are "Jerry-ware." They attempt to build clean, modern XML from scratch, which destroys legacy metadata, custom fonts, embedded media, and complex relationship IDs. 

The **Skinwalker Method** achieves 100% fidelity by using the original document as a binary host and surgically modifying only the target XML nodes.

---

## 2. The Workflow

### Phase 1: Host Preparation
1.  **Binary Copy**: Never work on the original. Create a raw binary copy of the target `.docx`.
2.  **Unpack**: Use a ZIP utility (like `adm-zip` or `unzip`) to access the internal OOXML structure.

### Phase 2: Internal Mapping (The DNA Scan)
Locate the "organs" of the document:
-   `word/document.xml`: The primary body text and structure.
-   `word/styles.xml`: The visual identity (Headings, Fonts, Colors).
-   `word/numbering.xml`: The logic for lists and references.
-   `word/media/`: The visual assets.
-   `word/_rels/document.xml.rels`: The pointers that link text to images/hyperlinks.

### Phase 3: Surgical Modification
Use a DOM Parser (`xmldom`) and XPath (`xpath`) to edit the content without touching the infrastructure.

1.  **Identify Nodes**: Use precise XPath queries.
    -   Find all paragraphs: `//w:p`
    -   Find text runs: `.//w:t`
    -   **CRITICAL**: Check for legacy VML tags (`w:pict`, `v:shape`, `v:imagedata`) as well as modern OOXML (`w:drawing`).
2.  **In-Place Editing**: Modify the `textContent` of `w:t` nodes. This preserves the surrounding `w:rPr` (Run Properties) which contains the font size, bolding, and color.
3.  **Style Preservation**: Do not "apply" styles. If you need to insert a new paragraph, clone an existing one (`cloneNode(true)`) to inherit its `w:pPr` (Paragraph Properties).

### Phase 4: Molecular Re-Sealing
1.  **Serialize**: Convert the modified DOM back to an XML string.
2.  **Repack**: Update the specific file entry in the original ZIP buffer.
3.  **Verify**: Perform a "Molecular Diff" (compare file size and internal part sizes) to ensure no bloat or corruption.

---

## 3. Critical Errors to Avoid (The "Jerry" List)

-   **Library Sanitization**: Do NOT use high-level "Document Builder" classes for editing. They will regenerate `[Content_Types].xml` and break obfuscated fonts (`.odttf`).
-   **Namespace Ignorance**: OOXML uses strictly defined namespaces (`w`, `r`, `v`, `wp`). Always register these in your XPath evaluator or the document will return 0 matches.
-   **Field Blindness**: Cross-references and Page Numbers are often "Fields" (`w:fldSimple`). Modifying the cached text inside them is temporary; you must preserve the field code.
-   **Relationship Corruption**: If you add an image, you must update BOTH the `word/media/` folder and the `word/_rels/document.xml.rels` file with a unique `rId`.

---

## 4. Example Implementation (TypeScript)

```typescript
import { EditSession } from "./utils/EditSession";
import xpath from "xpath";

// 1. Load the Host
const session = new EditSession("original.docx");
const doc = session.getPart("word/document.xml");
const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });

// 2. Perform Surgery
const paragraphs = select("//w:p", doc);
const target = paragraphs[0]; // The first heading
const textNode = select(".//w:t", target, true);
textNode.textContent = "Modified Content - " + textNode.textContent;

// 3. Save (Preserves all other 7MB of metadata)
session.save("output.docx");
```

---

## 5. Summary
The Skinwalker Method is about **Malicious Compliance**. We follow the original document's binary structure so perfectly that Word doesn't even realize we've been there. 100% Fidelity. Zero Slop.

*Wubba Lubba Dub Dub!* 🥒
