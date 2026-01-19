import { z } from "zod";
import AdmZip from "adm-zip";
import { DOMParser } from "xmldom";
import xpath from "xpath";
import * as fs from "fs";
import path from "path";

export const inspectDocTool = {
    name: "inspect_doc_formatting",
    description: "ULTRA-FIDELITY context-aware inspection of a .docx file.",
    inputSchema: z.object({
        path: z.string().describe("Absolute path to the .docx file"),
    }),
    handler: async (args: { path: string }) => {
        try {
            if (!fs.existsSync(args.path)) throw new Error(`File not found: ${args.path}`);

            const zip = new AdmZip(args.path);
            const mediaDir = path.join(path.dirname(args.path), "mcp_media");
            if (!fs.existsSync(mediaDir)) fs.mkdirSync(mediaDir);

            const select = xpath.useNamespaces({
                "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture"
            });

            const getRelsMap = (xmlPath: string) => {
                const dir = path.dirname(xmlPath);
                const base = path.basename(xmlPath);
                const relsFile = `${dir}/_rels/${base}.rels`;
                const relsMap: Record<string, string> = {};
                try {
                    const relsXml = zip.readAsText(relsFile);
                    if (relsXml) {
                        const relsDoc = new DOMParser().parseFromString(relsXml);
                        const relNodes = select("//r:Relationship", relsDoc) as any[];
                        if (relNodes.length === 0) {
                            // Fallback for namespace-less selection
                            const allRels = relsDoc.getElementsByTagName("Relationship");
                            for (let i = 0; i < allRels.length; i++) {
                                const n = allRels[i];
                                relsMap[n.getAttribute("Id")!] = n.getAttribute("Target")!;
                            }
                        } else {
                            relNodes.forEach(n => {
                                relsMap[n.getAttribute("Id")] = n.getAttribute("Target");
                            });
                        }
                    }
                } catch (e) {}
                return relsMap;
            };

            const parseParagraph = (pNode: Node, relsMap: Record<string, string>) => {
                const block: any = { type: "paragraph", children: [] };
                
                // Alignment / Style / Spacing / Indent
                const jc = select("w:pPr/w:jc/@w:val", pNode, true) as any;
                if (jc) block.alignment = jc.value;
                const style = select("w:pPr/w:pStyle/@w:val", pNode, true) as any;
                if (style) block.style = style.value;
                const spacing = select("w:pPr/w:spacing", pNode, true) as any;
                if (spacing) {
                    block.spacing = {
                        before: spacing.getAttribute("w:before") ? parseInt(spacing.getAttribute("w:before")) : undefined,
                        after: spacing.getAttribute("w:after") ? parseInt(spacing.getAttribute("w:after")) : undefined,
                        line: spacing.getAttribute("w:line") ? parseInt(spacing.getAttribute("w:line")) : undefined
                    };
                }
                const ind = select("w:pPr/w:ind", pNode, true) as any;
                if (ind) {
                    block.indent = {
                        left: ind.getAttribute("w:left") ? parseInt(ind.getAttribute("w:left")) : undefined,
                        hanging: ind.getAttribute("w:hanging") ? parseInt(ind.getAttribute("w:hanging")) : undefined,
                        firstLine: ind.getAttribute("w:firstLine") ? parseInt(ind.getAttribute("w:firstLine")) : undefined
                    };
                }

                // Walk all children of the paragraph to find Runs
                const children = pNode.childNodes;
                for (let i = 0; i < children.length; i++) {
                    const child = children[i];
                    if (child.nodeName === "w:r") {
                        // Process Run
                        const rChildren = child.childNodes;
                        for (let j = 0; j < rChildren.length; j++) {
                            const rChild = rChildren[j];
                            
                            // Text
                            if (rChild.nodeName === "w:t") {
                                const run: any = { type: "text_run", text: rChild.textContent || "" };
                                // Look for rPr in siblings
                                const rPr = select("w:rPr", child, true) as any;
                                if (rPr) {
                                    if (select("w:b", rPr, true)) run.bold = true;
                                    if (select("w:i", rPr, true)) run.italic = true;
                                    const sz = select("w:sz/@w:val", rPr, true) as any;
                                    if (sz) run.size = parseInt(sz.value) / 2;
                                    const font = select("w:rFonts/@w:ascii", rPr, true) as any;
                                    if (font) run.font = font.value;
                                }
                                block.children.push(run);
                            }

                            // Drawing / Image
                            if (rChild.nodeName === "w:drawing" || rChild.nodeName === "w:pict") {
                                // Find r:embed or r:id
                                const blip = select(".//@r:embed | .//@r:id", rChild, true) as any;
                                if (blip && relsMap[blip.value]) {
                                    const mediaPath = relsMap[blip.value];
                                    const zipPath = mediaPath.startsWith("media/") ? `word/${mediaPath}` : (mediaPath.startsWith("../media/") ? `word/${mediaPath.substring(3)}` : mediaPath);
                                    try {
                                        const buffer = zip.readFile(zipPath);
                                        if (buffer) {
                                            const fileName = path.basename(mediaPath);
                                            const localPath = path.join(mediaDir, fileName);
                                            fs.writeFileSync(localPath, buffer);
                                            
                                            const extent = select(".//wp:extent", rChild, true) as any;
                                            block.children.push({
                                                type: "image",
                                                path: localPath,
                                                width: extent ? Math.round(parseInt(extent.getAttribute("cx")) / 9525) : 300,
                                                height: extent ? Math.round(parseInt(extent.getAttribute("cy")) / 9525) : 200
                                            });
                                        }
                                    } catch (e) {}
                                }
                            }
                        }
                    }
                }

                if (block.children.length === 0) block.children.push({ type: "text_run", text: "" });
                return block;
            };

            const parseTable = (tblNode: Node, relsMap: Record<string, string>) => {
                const rows: any[] = [];
                const trs = select("w:tr", tblNode) as any[];
                trs.forEach(tr => {
                    const cells: any[] = [];
                    const tcs = select("w:tc", tr) as any[];
                    tcs.forEach(tc => {
                        const cellContent: any[] = [];
                        const contentNodes = tc.childNodes;
                        for (let i = 0; i < contentNodes.length; i++) {
                            const node = contentNodes[i];
                            if (node.nodeName === "w:p") cellContent.push(parseParagraph(node, relsMap));
                            if (node.nodeName === "w:tbl") cellContent.push(parseTable(node, relsMap));
                        }
                        cells.push({ content: cellContent });
                    });
                    rows.push({ cells });
                });
                return { type: "table", rows };
            };

            const docRels = getRelsMap("word/document.xml");
            const doc = new DOMParser().parseFromString(zip.readAsText("word/document.xml"));
            const body = select("/w:document/w:body", doc, true) as any;
            
            const resultBlocks: any[] = [];
            const bodyChildren = body.childNodes;
            for (let i = 0; i < bodyChildren.length; i++) {
                const node = bodyChildren[i];
                if (node.nodeName === "w:p") resultBlocks.push(parseParagraph(node, docRels));
                if (node.nodeName === "w:tbl") resultBlocks.push(parseTable(node, docRels));
            }

            const headers = zip.getEntries().filter(e => e.entryName.startsWith("word/header")).map(e => {
                const rels = getRelsMap(e.entryName);
                const xml = new DOMParser().parseFromString(zip.readAsText(e.entryName));
                const ps = select("//w:p", xml) as any[];
                return { name: e.entryName, content: ps.map(p => parseParagraph(p, rels)) };
            });

            const footers = zip.getEntries().filter(e => e.entryName.startsWith("word/footer")).map(e => {
                const rels = getRelsMap(e.entryName);
                const xml = new DOMParser().parseFromString(zip.readAsText(e.entryName));
                const ps = select("//w:p", xml) as any[];
                return { name: e.entryName, content: ps.map(p => parseParagraph(p, rels)) };
            });

            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({ headers, footers, sections: [{ children: resultBlocks }] }, null, 2)
                }],
                isError: false
            };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};
