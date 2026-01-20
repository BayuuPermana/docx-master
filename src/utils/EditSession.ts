import AdmZip from "adm-zip";
import { DOMParser, XMLSerializer } from "xmldom";
import * as fs from "fs";

export class EditSession {
    private zip: AdmZip;
    private parts: Record<string, Document> = {};
    private serializer = new XMLSerializer();

    constructor(private inputPath: string) {
        if (!fs.existsSync(inputPath)) throw new Error(`File not found: ${inputPath}`);
        this.zip = new AdmZip(inputPath);
    }

    public getPart(partPath: string): Document {
        if (this.parts[partPath]) return this.parts[partPath];
        
        const content = this.zip.readAsText(partPath);
        if (!content) throw new Error(`Part not found in ZIP: ${partPath}`);
        
        const dom = new DOMParser().parseFromString(content, "application/xml");
        this.parts[partPath] = dom;
        return dom;
    }

    public getPartNames(): string[] {
        return this.zip.getEntries().map(entry => entry.entryName);
    }

    public setPart(partPath: string, dom: Document) {
        this.parts[partPath] = dom;
    }

    public addMedia(mediaPath: string, targetName: string): string {
        const buffer = fs.readFileSync(mediaPath);
        this.zip.addFile(`word/media/${targetName}`, buffer);
        return `media/${targetName}`;
    }

    public save(outputPath: string) {
        // Sync parts back to ZIP
        for (const partPath in this.parts) {
            const xml = this.serializer.serializeToString(this.parts[partPath]);
            // Ensure no undefined or null strings get in
            if (xml) {
                this.zip.updateFile(partPath, Buffer.from(xml, 'utf-8'));
            }
        }
        this.zip.writeZip(outputPath);
    }
}
