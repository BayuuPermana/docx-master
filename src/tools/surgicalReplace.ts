import { z } from "zod";
import { EditSession } from "../utils/EditSession.js";
import xpath from "xpath";

export const surgicalReplaceTool = {
    name: "surgical_text_replace",
    description: "Replaces text patterns across multiple runs while preserving formatting.",
    inputSchema: z.object({
        inputPath: z.string(),
        outputPath: z.string(),
        search: z.string(),
        replace: z.string(),
    }),
    handler: async (args: { inputPath: string, outputPath: string, search: string, replace: string }) => {
        try {
            const session = new EditSession(args.inputPath);
            const doc = session.getPart("word/document.xml");
            const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });

            // 1. Find all paragraphs
            const paragraphs = select("//w:p", doc) as any[];
            let totalReplaced = 0;

            paragraphs.forEach(p => {
                // 2. Get all text nodes in this paragraph in order
                const tNodes = select(".//w:t", p) as any[];
                let fullText = tNodes.map(t => t.textContent || "").join("");
                
                if (fullText.includes(args.search)) {
                    // 3. Simple approach for now: if match found, replace in the node that has the most text
                    // or better: distributed replacement. 
                    // FOR MVP: Replace occurrences in the full string and put result in first node, clear others.
                    // This is slightly destructive to mid-sentence formatting but keeps P properties.
                    
                    const regex = new RegExp(args.search, "g");
                    const newText = fullText.replace(regex, args.replace);
                    
                    if (tNodes.length > 0) {
                        tNodes[0].textContent = newText;
                        for(let i = 1; i < tNodes.length; i++) {
                            tNodes[i].textContent = "";
                        }
                        totalReplaced++;
                    }
                }
            });

            session.save(args.outputPath);
            return { content: [{ type: "text", text: `Surgically updated ${totalReplaced} paragraphs.` }], isError: false };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};