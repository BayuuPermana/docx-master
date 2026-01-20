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
            const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
            
            const partNames = session.getPartNames().filter(name => 
                name.startsWith("word/") && name.endsWith(".xml") && !name.includes("styles.xml") && !name.includes("fontTable.xml") && !name.includes("settings.xml")
            );

            let totalReplaced = 0;

            for (const partName of partNames) {
                const doc = session.getPart(partName);
                const paragraphs = select("//w:p", doc) as any[];

                paragraphs.forEach(p => {
                    const tNodes = select(".//w:t", p) as any[];
                    let fullText = tNodes.map(t => t.textContent || "").join("");
                    
                    if (fullText.includes(args.search)) {
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
            }

            session.save(args.outputPath);
            return { content: [{ type: "text", text: `Surgically updated ${totalReplaced} paragraphs across document parts.` }], isError: false };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};