import { z } from "zod";
import mammoth from "mammoth";

export const readDocTool = {
  name: "read_doc_structure",
  description: "Reads a Microsoft Word (.docx) file and returns its content as Markdown structure.",
  inputSchema: z.object({
    path: z.string().describe("The absolute path to the local .docx file."),
  }),
  handler: async (args: { path: string }) => {
    try {
        const result = await mammoth.convertToMarkdown({ path: args.path });
        const markdown = result.value;
        
        return {
            content: [
                {
                    type: "text",
                    text: markdown
                }
            ],
            isError: false
        };
    } catch (error: any) {
        return {
            content: [
                {
                    type: "text",
                    text: `Error reading file: ${error.message}`
                }
            ],
            isError: true
        }
    }
  },
};
