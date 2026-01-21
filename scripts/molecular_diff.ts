import AdmZip from "adm-zip";
import * as fs from "fs";
import path from "path";

function getFolderSize(zip: AdmZip, folder: string): number {
    return zip.getEntries()
        .filter(e => e.entryName.startsWith(folder))
        .reduce((acc, e) => acc + e.header.size, 0);
}

async function compare(fileA: string, fileB: string) {
    console.log(`🥒 MOLECULAR DIFF: ${path.basename(fileA)} vs ${path.basename(fileB)}`);
    console.log("--------------------------------------------------");

    const statsA = fs.statSync(fileA);
    const statsB = fs.statSync(fileB);
    console.log(`Total File Size:`);
    console.log(`  Original: ${(statsA.size / 1024 / 1024).toFixed(2)} MB`);
    console.log(`  Clone:    ${(statsB.size / 1024 / 1024).toFixed(2)} MB`);
    console.log(`  Difference: ${(((statsB.size - statsA.size) / statsA.size) * 100).toFixed(2)}%`);

    const zipA = new AdmZip(fileA);
    const zipB = new AdmZip(fileB);

    const getFileInfo = (zip: AdmZip, name: string) => {
        const entry = zip.getEntry(name);
        return entry ? `${(entry.header.size / 1024).toFixed(2)} KB` : "MISSING";
    };

    console.log(`
Internal Organ Sizes:`);
    console.log(`  word/document.xml:`);
    console.log(`    Orig:  ${getFileInfo(zipA, "word/document.xml")}`);
    console.log(`    Clone: ${getFileInfo(zipB, "word/document.xml")}`);
    
    console.log(`  word/styles.xml:`);
    console.log(`    Orig:  ${getFileInfo(zipA, "word/styles.xml")}`);
    console.log(`    Clone: ${getFileInfo(zipB, "word/styles.xml")}`);

    const mediaA = getFolderSize(zipA, "word/media/");
    const mediaB = getFolderSize(zipB, "word/media/");
    console.log(`  word/media/ (Images):`);
    console.log(`    Orig:  ${(mediaA / 1024).toFixed(2)} KB (${zipA.getEntries().filter(e => e.entryName.startsWith("word/media/")).length} files)`);
    console.log(`    Clone: ${(mediaB / 1024).toFixed(2)} KB (${zipB.getEntries().filter(e => e.entryName.startsWith("word/media/")).length} files)`);

    console.log("\nTop 5 Largest Files in Original:");
    const topA = zipA.getEntries().sort((a, b) => b.header.size - a.header.size).slice(0, 5);
    topA.forEach(e => console.log(`  - ${e.entryName}: ${(e.header.size / 1024).toFixed(2)} KB`));
}

const [,, fileA, fileB] = process.argv;
if (!fileA || !fileB) {
    console.error("Usage: bun run scripts/molecular_diff.ts <fileA> <fileB>");
} else {
    compare(fileA, fileB).catch(console.error);
}
