import { z } from "zod";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, ImageRun, WidthType, PageOrientation, Header, Footer, BorderStyle } from "docx";
import * as fs from "fs";

const HeadingLevelMap: Record<string, any> = {
    Heading1: HeadingLevel.HEADING_1, Heading2: HeadingLevel.HEADING_2, Heading3: HeadingLevel.HEADING_3,
    Heading4: HeadingLevel.HEADING_4, Heading5: HeadingLevel.HEADING_5, Heading6: HeadingLevel.HEADING_6,
    Title: HeadingLevel.TITLE, Subtitle: HeadingLevel.SUBTITLE,
};

const AlignmentMap: Record<string, any> = {
    left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT, 
    justified: AlignmentType.JUSTIFIED, justify: AlignmentType.JUSTIFIED,
    both: AlignmentType.JUSTIFIED
};

export const createDocTool = {
    name: "create_styled_doc",
    description: "Creates a .docx file with ultra-high fidelity images, headers, and footers.",
    inputSchema: z.object({
        path: z.string(),
        sections: z.array(z.any()),
        headers: z.array(z.any()).optional(),
        footers: z.array(z.any()).optional(),
    }),
    handler: async (args: { path: string, sections: any[], headers?: any[], footers?: any[] }) => {
        try {
            const parseParagraph = (p: any) => {
                const children: any[] = [];
                (p.children || []).forEach((c: any) => {
                    if (c.type === "text_run") {
                        children.push(new TextRun({
                            text: String(c.text || ""),
                            bold: !!c.bold,
                            italics: !!c.italic,
                            underline: c.underline ? { type: "single" } : undefined,
                            size: (typeof c.size === 'number') ? Math.round(c.size * 2) : undefined,
                            color: (typeof c.color === 'string' && c.color.length === 6) ? c.color : undefined,
                            font: c.font
                        }));
                    } else if (c.type === "image") {
                        if (fs.existsSync(c.path)) {
                            children.push(new ImageRun({
                                data: fs.readFileSync(c.path),
                                transformation: { width: c.width || 100, height: c.height || 100 }
                            }));
                        }
                    }
                });

                if (children.length === 0) children.push(new TextRun(""));

                return new Paragraph({
                    children: children,
                    heading: HeadingLevelMap[p.style] || undefined,
                    alignment: (AlignmentMap as any)[p.alignment] || undefined,
                    spacing: p.spacing ? {
                        before: p.spacing.before,
                        after: p.spacing.after,
                        line: p.spacing.line
                    } : undefined,
                    indent: p.indent ? {
                        left: p.indent.left,
                        hanging: p.indent.hanging,
                        firstLine: p.indent.firstLine
                    } : undefined
                });
            };

            const parseBlocks = (blocks: any[]) => (blocks || []).map(block => {
                if (block.type === "paragraph") return parseParagraph(block);
                if (block.type === "table") {
                    return new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
                            bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
                            left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
                            right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
                            insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
                        },
                        rows: (block.rows || []).map((row: any) => new TableRow({
                            children: (row.cells || []).map((cell: any) => new TableCell({
                                children: (cell.content || []).map((b: any) => {
                                    if (b.type === "paragraph") return parseParagraph(b);
                                    if (b.type === "table") return new Table({ rows: [] }); // Recursive nesting too deep for now
                                    return new Paragraph("");
                                })
                            }))
                        }))
                    });
                }
                return new Paragraph("");
            });

            const docSections = args.sections.map(sec => {
                const sectionProps: any = {
                    children: parseBlocks(sec.children || []),
                    properties: {
                        page: {
                            margin: sec.properties?.margins,
                            orientation: sec.properties?.orientation === "landscape" ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT
                        }
                    }
                };

                if (args.headers?.[0]?.content) {
                    sectionProps.headers = { default: new Header({ children: parseBlocks(args.headers[0].content) }) };
                }
                if (args.footers?.[0]?.content) {
                    sectionProps.footers = { default: new Footer({ children: parseBlocks(args.footers[0].content) }) };
                }

                return sectionProps;
            });

            const doc = new Document({ sections: docSections });
            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(args.path, buffer);
            return { content: [{ type: "text", text: `Success: ${args.path}` }], isError: false };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Fatal Error: ${error.message}` }], isError: true };
        }
    }
};
