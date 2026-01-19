import { z } from "zod";
import AdmZip from "adm-zip";
import { XMLParser } from "fast-xml-parser";
import * as fs from "fs";
import path from "path";

export const inspectDocTool = {
    name: "inspect_doc_formatting",
    description: "FORENSIC inspection of a .docx file. Captures content, styles, headers, footers, and media.",
    inputSchema: z.object({
        path: z.string().describe("Absolute path to the .docx file"),
    }),
    handler: async (args: { path: string }) => {
        try {
            if (!fs.existsSync(args.path)) throw new Error(`File not found: ${args.path}`);

            const zip = new AdmZip(args.path);
            const parser = new XMLParser({
                ignoreAttributes: false,
                attributeNamePrefix: "@_",
                trimValues: false
            });

            // --- Helper: Safe XML Read ---
            const readXml = (fileName: string) => {
                try {
                    const content = zip.readAsText(fileName);
                    return content ? parser.parse(content) : null;
                } catch (e) { return null; }
            };

            const docXml = readXml("word/document.xml");
            const relsXml = readXml("word/_rels/document.xml.rels");
            const stylesXml = readXml("word/styles.xml");
            const numberingXml = readXml("word/numbering.xml");

            // --- Map Relationships ---
            const relsMap: Record<string, string> = {};
            if (relsXml?.Relationships?.Relationship) {
                const relList = Array.isArray(relsXml.Relationships.Relationship) 
                    ? relsXml.Relationships.Relationship : [relsXml.Relationships.Relationship];
                relList.forEach((rel: any) => relsMap[rel["@_Id"]] = rel["@_Target"]);
            }

            // --- Extract Media ---
            const mediaDir = path.join(path.dirname(args.path), "mcp_media");
            if (!fs.existsSync(mediaDir)) fs.mkdirSync(mediaDir);

            // --- Parser: Paragraph to JSON ---
            const parseParagraph = (p: any) => {
                const block: any = { type: "paragraph" };
                const pPr = p["w:pPr"];
                if (pPr) {
                    if (pPr["w:jc"]) block.alignment = pPr["w:jc"]["@_w:val"];
                    if (pPr["w:pStyle"]) block.style = pPr["w:pStyle"]["@_w:val"];
                    const spacing = pPr["w:spacing"];
                    if (spacing) {
                        block.spacing = {
                            before: spacing["@_w:before"] ? parseInt(spacing["@_w:before"]) : undefined,
                            after: spacing["@_w:after"] ? parseInt(spacing["@_w:after"]) : undefined,
                            line: spacing["@_w:line"] ? parseInt(spacing["@_w:line"]) : undefined
                        };
                    }
                }

                const children: any[] = [];
                const runs = p["w:r"] ? (Array.isArray(p["w:r"]) ? p["w:r"] : [p["w:r"]]) : [];
                
                runs.forEach((r: any) => {
                    if (r["w:t"]) {
                        const run: any = { type: "text_run" };
                        const t = r["w:t"];
                        run.text = (typeof t === "object") ? t["#text"] || "" : String(t);
                        const rPr = r["w:rPr"];
                        if (rPr) {
                            if (rPr["w:b"] !== undefined) run.bold = true;
                            if (rPr["w:i"] !== undefined) run.italic = true;
                            if (rPr["w:u"] !== undefined) run.underline = true;
                            if (rPr["w:sz"]) run.size = parseInt(rPr["w:sz"]["@_w:val"]) / 2;
                            if (rPr["w:color"]) run.color = rPr["w:color"]["@_w:val"];
                            if (rPr["w:rFonts"]) run.font = rPr["w:rFonts"]["@_w:ascii"];
                        }
                        children.push(run);
                    }
                    if (r["w:drawing"]) {
                        const inline = r["w:drawing"]["wp:inline"] || r["w:drawing"]["wp:anchor"];
                        const blip = inline?.["a:graphic"]?.["a:graphicData"]?.["pic:pic"]?.["pic:blipFill"]?.["a:blip"];
                        const relId = blip?.["@_r:embed"];
                        if (relId && relsMap[relId]) {
                            const mediaPath = relsMap[relId];
                            const zipPath = mediaPath.startsWith("media/") ? `word/${mediaPath}` : mediaPath;
                            const buffer = zip.readFile(zipPath);
                            if (buffer) {
                                const localPath = path.join(mediaDir, path.basename(mediaPath));
                                fs.writeFileSync(localPath, buffer);
                                children.push({ type: "image", path: localPath, width: 300, height: 200 }); // Default size if extent missing
                            }
                        }
                    }
                });
                if (children.length === 0) children.push({ type: "text_run", text: "" });
                block.children = children;
                return block;
            };

            // --- Process Body ---
            const body = docXml?.["w:document"]?.["w:body"];
            const pList = body?.["w:p"] ? (Array.isArray(body["w:p"]) ? body["w:p"] : [body["w:p"]]) : [];
            const resultBlocks = pList.map(parseParagraph);

            // --- Process Headers/Footers ---
            const headers = zip.getEntries().filter(e => e.entryName.startsWith("word/header")).map(e => {
                const xml = parser.parse(zip.readAsText(e.entryName));
                const hdrP = xml["w:hdr"]?.["w:p"];
                const hdrPList = hdrP ? (Array.isArray(hdrP) ? hdrP : [hdrP]) : [];
                return { name: e.entryName, content: hdrPList.map(parseParagraph) };
            });

            const footers = zip.getEntries().filter(e => e.entryName.startsWith("word/footer")).map(e => {
                const xml = parser.parse(zip.readAsText(e.entryName));
                const ftrP = xml["w:ftr"]?.["w:p"];
                const ftrPList = ftrP ? (Array.isArray(ftrP) ? ftrP : [ftrP]) : [];
                return { name: e.entryName, content: ftrPList.map(parseParagraph) };
            });

            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        metadata: {
                            styles: stylesXml?.["w:styles"]?.["w:style"],
                            numbering: numberingXml?.["w:numbering"]
                        },
                        headers,
                        footers,
                        sections: [{ children: resultBlocks }]
                    }, null, 2)
                }],
                isError: false
            };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};