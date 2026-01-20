import { z } from "zod";
import { EditSession } from "../utils/EditSession.js";
import xpath from "xpath";
import { DOMParser } from "xmldom";
import * as fs from "fs";
import path from "path";

export const surgicalAddImageTool = {
    name: "surgical_add_image",
    description: "Surgically injects an image into a .docx file using raw OOXML injection.",
    inputSchema: z.object({
        inputPath: z.string(),
        outputPath: z.string(),
        imagePath: z.string(),
        targetParagraphIndex: z.number().default(0),
        width: z.number().default(100),
        height: z.number().default(100),
    }),
    handler: async (args: { inputPath: string, outputPath: string, imagePath: string, targetParagraphIndex: number, width: number, height: number }) => {
        try {
            const session = new EditSession(args.inputPath);       

            // 1. Add Media
            const ext = path.extname(args.imagePath).substring(1); 
            const mediaName = `injected_${Date.now()}.${ext}`;     
            session.addMedia(args.imagePath, mediaName);

            // 2. Update Rels
            const relsDoc = session.getPart("word/_rels/document.xml.rels");
            const relsRoot = relsDoc.documentElement;
            const rId = `rIdInject${Date.now()}_${Math.floor(Math.random() * 1000)}`;

            const newRel = relsDoc.createElement("Relationship");  
            newRel.setAttribute("Id", rId);
            newRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
            newRel.setAttribute("Target", `media/${mediaName}`);   
            relsRoot.appendChild(newRel);

            // 3. Inject XML into Document
            const doc = session.getPart("word/document.xml");      
            const select = xpath.useNamespaces({ "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" });
            const paragraphs = select("//w:p", doc) as any[];      
            const p = paragraphs[args.targetParagraphIndex];       

            if (!p) throw new Error("Target paragraph not found.");

            // Construct drawing XML (Highly simplified for injection)
            const cx = args.width * 9525;
            const cy = args.height * 9525;

            const drawingXml = `
                <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                    <w:drawing>
                        <wp:inline distT="0" distB="0" distL="0" distR="0">
                            <wp:extent cx="${cx}" cy="${cy}"/>     
                            <wp:docPr id="1" name="Injected Image"/>
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                            </wp:cNvGraphicFramePr>
                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:nvPicPr>
                                            <pic:cNvPr id="0" name="${mediaName}"/>
                                            <pic:cNvPicPr/>        
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                            <a:blip r:embed="${rId}"/>
                                            <a:stretch><a:fillRect/></a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing>
                </w:r>`;

            const drawingDom = new DOMParser().parseFromString(drawingXml);
            p.appendChild(doc.importNode(drawingDom.documentElement, true));
            session.save(args.outputPath);
            return { content: [{ type: "text", text: `Image injected with rId: ${rId}. File saved to ${args.outputPath}` }], isError: false };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};
