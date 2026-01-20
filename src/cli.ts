#!/usr/bin/env bun
import { readDocTool } from "./tools/readDoc.js";
import { createDocTool } from "./tools/createDoc.js";
import { inspectDocTool } from "./tools/inspectDoc.js";
import { cleanupMediaTool } from "./tools/cleanupMedia.js";
import { surgicalReplaceTool } from "./tools/surgicalReplace.js";
import { surgicalAddImageTool } from "./tools/surgicalAddImage.js";
import { surgicalInsertTool } from "./tools/surgicalInsert.js";

const TOOLS: Record<string, any> = {
  read: readDocTool,
  create: createDocTool,
  inspect: inspectDocTool,
  cleanup: cleanupMediaTool,
  replace: surgicalReplaceTool,
  addImage: surgicalAddImageTool,
  insert: surgicalInsertTool,
};

async function main() {
  const [,, toolName, ...argsList] = process.argv;

  if (!toolName || !TOOLS[toolName]) {
    console.error(`Unknown tool: ${toolName}. Available: ${Object.keys(TOOLS).join(", ")}`);
    process.exit(1);
  }

  const tool = TOOLS[toolName];
  const args: Record<string, any> = {};

  // Simple key=value parser for CLI args
  for (const arg of argsList) {
    const [key, ...valueParts] = arg.split("=");
    if (key && valueParts.length > 0) {
      args[key.replace(/^--/, "")] = valueParts.join("=");
    }
  }

  try {
    const result = await tool.handler(args);
    if (result.isError) {
      console.error(JSON.stringify(result.content, null, 2));
      process.exit(1);
    } else {
      console.log(result.content.map((c: any) => c.text).join("\n"));
    }
  } catch (error: any) {
    console.error(`Error executing ${toolName}: ${error.message}`);
    process.exit(1);
  }
}

main();
