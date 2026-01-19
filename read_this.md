To truly master Word document manipulation, you need to understand that a `.docx` file is a **ZIP archive** following the **Open Packaging Conventions (OPC)**. When you unzip a document, you are looking at a collection of XML files that represent the "Office Open XML" (OOXML) standard.

Here is the complete breakdown of every critical component within that package.

### 1. The Root Directory

This level handles the "packaging" logic rather than the content.

* **`[Content_Types].xml`**: The most important file for the ZIP parser. It acts as a MIME-type registry for every file in the archive. If a file is in the ZIP but not listed here, Word will consider the document corrupted.
* **`_rels/.rels`**: The global relationships file. It tells Word where the "main part" of the document is (usually pointing to `word/document.xml`).

---

### 2. The `word/` Directory (The Engine Room)

This is where 90% of your MCP server's logic will live.

#### **`word/document.xml`**

The actual content. It is structured as a hierarchical tree:

* **`<w:body>`**: The container for all visible content.
* **`<w:p>` (Paragraph)**: The basic unit of layout.
* **`<w:pPr>` (Paragraph Properties)**: Inside a `<w:p>`, this defines alignment, indentation, and spacing.
* **`<w:r>` (Run)**: A wrapper for text that shares the same formatting. If you change a font mid-sentence, you start a new Run.
* **`<w:rPr>` (Run Properties)**: Inside a `<w:r>`, this defines bold, italic, font face, and size.
* **`<w:t>` (Text)**: The actual string characters.

#### **`word/styles.xml`**

This is a database of "Definitions." Instead of defining "Arial 12pt Bold" on every paragraph, Word assigns a `styleId` (like `Heading1`). This file maps that ID to the actual formatting rules. This allows for global updates—change the font in `styles.xml`, and it updates everywhere.

#### **`word/numbering.xml`**

One of the most complex files. It separates the **List Definition** (what the bullets look like) from the **List Instance** (which specific paragraphs belong to which list). If you are building a "List" tool in your MCP, you must touch this file.

#### **`word/theme/theme1.xml`**

Defines the color palette (Accent 1, Accent 2) and the default "Major" and "Minor" fonts. When a user changes the "Theme" in Word, it only changes this file, and the rest of the document (which refers to "Accent 1") updates automatically.

#### **`word/settings.xml`**

Document-wide toggles:

* Is "Track Changes" on?
* Is the document protected?
* Zoom level and View mode.
* Hyphenation settings.

---

### 3. The `word/_rels/` Directory

* **`document.xml.rels`**: This is the "Linking Registry." If your document has an image, the `document.xml` doesn't contain the image path. Instead, it says: "Display the image with `rId="rId5"`. You then look at this `.rels` file to see that `rId5` points to `media/image1.png`.

---

### 4. Layout & Media

* **`word/media/`**: A folder containing the raw binary files for every image, video, or audio clip used in the document.
* **`word/header1.xml` / `word/footer1.xml**`: These exist outside the main `document.xml` because they repeat across pages.
* **`word/footnotes.xml` / `word/endnotes.xml**`: Separate XML trees for citations to keep the main text flow clean.

---

### 5. Why is this hard to edit with raw Shell?

If you wanted to add a simple sentence with an image via shell commands, you would have to:

1. Add the image to `word/media/`.
2. Add a relationship entry in `word/_rels/document.xml.rels` to give that image a Relationship ID.
3. Add the MIME type of the image to `[Content_Types].xml`.
4. Inject the `<w:drawing>` XML block into `word/document.xml` referencing that Relationship ID.
5. Re-zip everything.
