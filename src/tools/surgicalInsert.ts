import { z } from "zod";
import { EditSession } from "../utils/EditSession.js";
import xpath from "xpath";

export const surgicalInsertTool = {
    name: "surgical_insert_paragraph",
    description: "Inserts a new paragraph into a .docx file, inheriting styles from a template paragraph.",
    inputSchema: z.object({
        inputPath: z.string(),
        outputPath: z.string(),
        templateParagraphIndex: z.number().describe("Index of the paragraph to clone styles from"),
        text: z.string().describe("Text content for the new paragraph"),
        insertAfter: z.boolean().default(true).describe("If true, inserts after the template; otherwise before."),
    }),
    handler: async (args: { inputPath: string, outputPath: string, templateParagraphIndex: number, text: string, insertAfter: boolean }) => {
        try {
            const session = new EditSession(args.inputPath);
            const doc = session.getPart("word/document.xml");
            const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
            
            const paragraphs = select("//w:p", doc) as any[];
            const templateP = paragraphs[args.templateParagraphIndex];

            if (!templateP) throw new Error("Template paragraph not found.");

            // 1. Create new Paragraph
            const newP = doc.createElement("w:p");

            // 2. Clone Properties (w:pPr)
            const pPr = select("w:pPr", templateP, true) as Node;
            if (pPr) {
                newP.appendChild(pPr.cloneNode(true));
            }

            // 3. Create Run and Text
            const run = doc.createElement("w:r");
            const textNode = doc.createElement("w:t");
            textNode.textContent = args.text;
            run.appendChild(textNode);
            newP.appendChild(run);

            // 4. Insert into DOM
            if (args.insertAfter) {
                if (templateP.nextSibling) {
                    templateP.parentNode?.insertBefore(newP, templateP.nextSibling);
                } else {
                    templateP.parentNode?.appendChild(newP);
                }
            } else {
                templateP.parentNode?.insertBefore(newP, templateP);
            }

            session.save(args.outputPath);
            return { content: [{ type: "text", text: `Success! Paragraph inserted at index ${args.templateParagraphIndex + (args.insertAfter ? 1 : 0)}` }], isError: false };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};
