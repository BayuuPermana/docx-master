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
                "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
                "v": "urn:schemas-microsoft-com:vml"
            });

            const getRelsMap = (xmlPath: string) => {
                const dir = path.dirname(xmlPath);
                const base = path.basename(xmlPath);
                const relsFile = (dir === "." || dir === "") ? `_rels/${base}.rels` : `${dir}/_rels/${base}.rels`;
                const relsMap: Record<string, string> = {};
                try {
                    const content = zip.readAsText(relsFile);
                    if (content) {
                        const relsDoc = new DOMParser().parseFromString(content);
                        const relNodes = select("//r:Relationship", relsDoc) as any[];
                        relNodes.forEach(n => {
                            relsMap[n.getAttribute("Id")] = n.getAttribute("Target");
                        });
                    }
                } catch (e) {}
                return relsMap;
            };

            const stylesXml = zip.readAsText("word/styles.xml") || "";
            let numberingXml = "";
            try { numberingXml = zip.readAsText("word/numbering.xml"); } catch (e) {}

            const parseRun = (rNode: Node, relsMap: Record<string, string>) => {
                const rPr = select("w:rPr", rNode, true) as any;
                
                const tNode = select("w:t", rNode, true) as any;
                if (tNode) {
                    const run: any = { type: "text_run", text: tNode.textContent || "" };
                    if (rPr) {
                        if (select("w:b", rPr, true)) run.bold = true;
                        if (select("w:i", rPr, true)) run.italic = true;
                        if (select("w:u", rPr, true)) run.underline = true;
                        const sz = select("w:sz/@w:val", rPr, true) as any;
                        if (sz) run.size = parseInt(sz.value) / 2;
                        const color = select("w:color/@w:val", rPr, true) as any;
                        if (color) run.color = color.value;
                        const font = select("w:rFonts/@w:ascii", rPr, true) as any;
                        if (font) run.font = font.value;
                    }
                    return run;
                }

                // IMPROVED IMAGE DETECTION (Modern + Legacy VML)
                const imageId = select(".//@r:embed | .//@r:id | .//@v:relid", rNode, true) as any;
                if (imageId && relsMap[imageId.value]) {
                    let mediaTarget = relsMap[imageId.value];
                    const zipPath = mediaTarget.startsWith("media/") ? `word/${mediaTarget}` : (mediaTarget.startsWith("../media/") ? `word/${mediaTarget.substring(3)}` : mediaTarget);
                    try {
                        const buffer = zip.readFile(zipPath);
                        if (buffer) {
                            const ext = path.extname(mediaTarget) || ".png";
                            const localPath = path.join(mediaDir, `extracted_${imageId.value}${ext}`);
                            fs.writeFileSync(localPath, buffer);
                            const extent = select(".//wp:extent", rNode, true) as any;
                            return {
                                type: "image",
                                path: localPath,
                                width: extent ? Math.round(parseInt(extent.getAttribute("cx")) / 9525) : 300,
                                height: extent ? Math.round(parseInt(extent.getAttribute("cy")) / 9525) : 200
                            };
                        }
                    } catch (e) {}
                }

                const instrText = select(".//w:instrText", rNode, true) as any;
                if (instrText && instrText.textContent?.includes("PAGE")) return { type: "page_number" };

                return null;
            };

            const parseParagraph = (pNode: Node, relsMap: Record<string, string>) => {
                const block: any = { type: "paragraph", children: [] };
                const pPr = select("w:pPr", pNode, true) as any;
                if (pPr) {
                    const jc = select("w:jc/@w:val", pPr, true) as any;
                    if (jc) block.alignment = jc.value;
                    const style = select("w:pStyle/@w:val", pPr, true) as any;
                    if (style) block.style = style.value;
                    const spacing = select("w:spacing", pPr, true) as any;
                    if (spacing) {
                        block.spacing = {
                            before: spacing.getAttribute("w:before") ? parseInt(spacing.getAttribute("w:before")) : undefined,
                            after: spacing.getAttribute("w:after") ? parseInt(spacing.getAttribute("w:after")) : undefined,
                            line: spacing.getAttribute("w:line") ? parseInt(spacing.getAttribute("w:line")) : undefined,
                            lineRule: spacing.getAttribute("w:lineRule") || undefined
                        };
                    }
                    const ind = select("w:ind", pPr, true) as any;
                    if (ind) {
                        block.indent = {
                            left: ind.getAttribute("w:left") ? parseInt(ind.getAttribute("w:left")) : undefined,
                            right: ind.getAttribute("w:right") ? parseInt(ind.getAttribute("w:right")) : undefined,
                            hanging: ind.getAttribute("w:hanging") ? parseInt(ind.getAttribute("w:hanging")) : undefined,
                            firstLine: ind.getAttribute("w:firstLine") ? parseInt(ind.getAttribute("w:firstLine")) : undefined
                        };
                    }
                }

                const walk = (node: Node) => {
                    const children = node.childNodes;
                    for (let i = 0; i < children.length; i++) {
                        const child = children[i];
                        if (child.nodeName === "w:r") {
                            const res = parseRun(child, relsMap);
                            if (res) block.children.push(res);
                        } else if (child.nodeName === "w:hyperlink" || child.nodeName === "w:fldSimple") {
                            walk(child);
                        }
                    }
                };
                walk(pNode);

                if (block.children.length === 0) block.children.push({ type: "text_run", text: "" });
                return block;
            };

            const parseTable = (tblNode: Node, relsMap: Record<string, string>) => {
                const rows: any[] = [];
                const trs = select(".//w:tr", tblNode) as any[];
                trs.forEach(tr => {
                    const cells: any[] = [];
                    const tcs = select("./w:tc", tr) as any[];
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
            const sectPr = select("//w:sectPr", doc, true) as any;
            const margins: any = {};
            if (sectPr) {
                const pgMar = select("w:pgMar", sectPr, true) as any;
                if (pgMar) {
                    margins.top = parseInt(pgMar.getAttribute("w:top"));
                    margins.bottom = parseInt(pgMar.getAttribute("w:bottom"));
                    margins.left = parseInt(pgMar.getAttribute("w:left"));
                    margins.right = parseInt(pgMar.getAttribute("w:right"));
                    margins.header = parseInt(pgMar.getAttribute("w:header"));
                    margins.footer = parseInt(pgMar.getAttribute("w:footer"));
                }
            }

            const body = select("/w:document/w:body", doc, true) as any;
            const resultBlocks: any[] = [];
            const bodyChildren = body.childNodes;
            for (let i = 0; i < bodyChildren.length; i++) {
                const node = bodyChildren[i];
                if (node.nodeName === "w:p") resultBlocks.push(parseParagraph(node, docRels));
                if (node.nodeName === "w:tbl") resultBlocks.push(parseTable(node, docRels));
            }

            const extractParts = (prefix: string) => zip.getEntries().filter(e => e.entryName.startsWith(prefix)).map(e => {
                const rels = getRelsMap(e.entryName);
                const xml = new DOMParser().parseFromString(zip.readAsText(e.entryName));
                const ps = select("//w:p", xml) as any[];
                return { name: e.entryName, content: ps.map(p => parseParagraph(p, rels)) };
            });

            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({ 
                        stylesXml, numberingXml,
                        headers: extractParts("word/header"), 
                        footers: extractParts("word/footer"), 
                        sections: [{ properties: { margins }, children: resultBlocks }] 
                    }, null, 2)
                }],
                isError: false
            };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};