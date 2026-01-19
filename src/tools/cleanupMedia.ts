import { z } from "zod";
import * as fs from "fs";
import path from "path";

export const cleanupMediaTool = {
    name: "cleanup_media",
    description: "Deletes the temporary mcp_media folder to save space.",
    inputSchema: z.object({
        directory: z.string().describe("The directory containing the mcp_media folder"),
    }),
    handler: async (args: { directory: string }) => {
        try {
            const mediaDir = path.join(args.directory, "mcp_media");
            if (fs.existsSync(mediaDir)) {
                fs.rmSync(mediaDir, { recursive: true, force: true });
                return {
                    content: [{ type: "text", text: "Media folder cleaned up successfully." }],
                    isError: false
                };
            }
            return {
                content: [{ type: "text", text: "No media folder found." }],
                isError: false
            };
        } catch (error: any) {
            return {
                content: [{ type: "text", text: `Error: ${error.message}` }],
                isError: true
            };
        }
    }
};
