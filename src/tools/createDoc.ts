import { z } from "zod";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, ImageRun, WidthType, PageOrientation, Header, Footer } from "docx";
import * as fs from "fs";

const HeadingLevelMap: Record<string, any> = {
    Heading1: HeadingLevel.HEADING_1, Heading2: HeadingLevel.HEADING_2, Heading3: HeadingLevel.HEADING_3,
    Heading4: HeadingLevel.HEADING_4, Heading5: HeadingLevel.HEADING_5, Heading6: HeadingLevel.HEADING_6,
    Title: HeadingLevel.TITLE, Subtitle: HeadingLevel.SUBTITLE,
};

const AlignmentMap: Record<string, any> = {
    left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT, justified: AlignmentType.JUSTIFIED,
};

const TextRunSchema = z.object({
    text: z.string().default(""), bold: z.boolean().optional(), italic: z.boolean().optional(),
    underline: z.boolean().optional(), size: z.number().optional(), color: z.string().optional(), font: z.string().optional(),
});

const ImageSchema = z.object({ type: z.literal("image"), path: z.string(), width: z.number(), height: z.number() });

const ParagraphSchema = z.object({
    type: z.literal("paragraph"), text: z.string().optional(),
    children: z.array(z.union([TextRunSchema, ImageSchema])).optional(),
    heading: z.enum(["Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Title", "Subtitle"]).optional(),
    alignment: z.enum(["left", "center", "right", "justified"]).optional(),
    spacing: z.object({ before: z.number().optional(), after: z.number().optional(), line: z.number().optional() }).optional(),
});

const BlockSchema = z.union([ParagraphSchema, z.any()]); // Simplify for brevity

const SectionSchema = z.object({
    headers: z.array(z.object({ content: z.array(ParagraphSchema) })).optional(),
    footers: z.array(z.object({ content: z.array(ParagraphSchema) })).optional(),
    properties: z.object({
        margins: z.object({ top: z.number().optional(), bottom: z.number().optional(), left: z.number().optional(), right: z.number().optional() }).optional(),
        orientation: z.enum(["portrait", "landscape"]).optional(),
    }).optional(),
    children: z.array(BlockSchema),
});

export const createDocTool = {
    name: "create_styled_doc",
    description: "Creates a .docx file with full support for headers, footers, and styles.",
    inputSchema: z.object({ path: z.string(), sections: z.array(SectionSchema) }),
    handler: async (args: { path: string, sections: any[] }) => {
        try {
            const docSections = args.sections.map(sec => {
                const parseBlocks = (blocks: any[]) => blocks.map(block => {
                    const runChildren: any[] = [];
                    if (block.text) runChildren.push(new TextRun({ text: block.text }));
                    if (block.children) {
                        block.children.forEach((c: any) => {
                            if (c.type === "image") {
                                if (fs.existsSync(c.path)) runChildren.push(new ImageRun({ data: fs.readFileSync(c.path), transformation: { width: c.width, height: c.height } }));
                            } else {
                                runChildren.push(new TextRun({ text: String(c.text || ""), bold: c.bold, italics: c.italic, underline: c.underline ? { type: "single" } : undefined, size: c.size ? c.size * 2 : undefined, color: c.color, font: c.font }));
                            }
                        });
                    }
                    if (runChildren.length === 0) runChildren.push(new TextRun(""));
                    return new Paragraph({ children: runChildren, heading: block.heading ? HeadingLevelMap[block.heading] : undefined, alignment: block.alignment ? AlignmentMap[block.alignment] : undefined, spacing: block.spacing });
                });

                const sectionProps: any = { children: parseBlocks(sec.children) };
                if (sec.headers) sectionProps.headers = { default: new Header({ children: parseBlocks(sec.headers[0].content) }) };
                if (sec.footers) sectionProps.footers = { default: new Footer({ children: parseBlocks(sec.footers[0].content) }) };
                if (sec.properties) sectionProps.properties = { page: { margin: sec.properties.margins, orientation: sec.properties.orientation === "landscape" ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT } };
                
                return sectionProps;
            });

            const doc = new Document({ sections: docSections });
            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(args.path, buffer);
            return { content: [{ type: "text", text: `Success: ${args.path}` }], isError: false };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Error: ${error.message}` }], isError: true };
        }
    }
};