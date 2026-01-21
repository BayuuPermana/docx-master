import { z } from "zod";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, ImageRun, WidthType, PageOrientation, Header, Footer, BorderStyle, PageNumber, IStylesOptions } from "docx";
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
        externalStyles: z.string().optional()
    }),
    handler: async (args: { path: string, sections: any[], headers?: any[], footers?: any[], externalStyles?: string }) => {
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
                            color: (typeof c.color === 'string' && (c.color.length === 6 || c.color.length === 8)) ? c.color : undefined,
                            font: c.font
                        }));
                    } else if (c.type === "image") {
                        if (fs.existsSync(c.path)) {
                            children.push(new ImageRun({
                                data: fs.readFileSync(c.path),
                                transformation: { width: c.width || 100, height: c.height || 100 }
                            }));
                        }
                    } else if (c.type === "page_number") {
                        children.push(PageNumber.CURRENT);
                    }
                });

                if (children.length === 0) children.push(new TextRun(""));

                const pProps: any = { children };
                if (p.style) pProps.style = p.style; 
                if (AlignmentMap[p.alignment]) pProps.alignment = AlignmentMap[p.alignment];
                
                if (p.spacing) {
                    pProps.spacing = {};
                    if (typeof p.spacing.before === 'number') pProps.spacing.before = p.spacing.before;
                    if (typeof p.spacing.after === 'number') pProps.spacing.after = p.spacing.after;
                    if (typeof p.spacing.line === 'number') pProps.spacing.line = p.spacing.line;
                    if (p.spacing.lineRule) pProps.spacing.lineRule = p.spacing.lineRule;
                }

                if (p.indent) {
                    pProps.indent = {};
                    if (typeof p.indent.left === 'number') pProps.indent.left = p.indent.left;
                    if (typeof p.indent.right === 'number') pProps.indent.right = p.indent.right;
                    if (typeof p.indent.hanging === 'number') pProps.indent.hanging = p.indent.hanging;
                    if (typeof p.indent.firstLine === 'number') pProps.indent.firstLine = p.indent.firstLine;
                }

                return new Paragraph(pProps);
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
                                    if (b.type === "table") return new Table({ rows: [] }); 
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

                if (args.headers?.length > 0) {
                    sectionProps.headers = { default: new Header({ children: parseBlocks(args.headers[0].content) }) };
                }
                if (args.footers?.length > 0) {
                    const pageFooter = args.footers.find((f:any) => f.content.some((p:any) => p.children.some((c:any) => c.type === "page_number")));
                    const footerToUse = pageFooter || args.footers[0];
                    sectionProps.footers = { default: new Footer({ children: parseBlocks(footerToUse.content) }) };
                }

                return sectionProps;
            });

            const docConfig: any = { sections: docSections };
            if (args.externalStyles) {
                // docx library supports importing external styles XML
                docConfig.externalStyles = args.externalStyles;
            }

            const doc = new Document(docConfig);
            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(args.path, buffer);
            return { content: [{ type: "text", text: `Success: ${args.path}` }], isError: false };
        } catch (error: any) {
            return { content: [{ type: "text", text: `Fatal Error: ${error.message}` }], isError: true };
        }
    }
};