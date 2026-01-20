#!/usr/bin/env bun
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { zodToJsonSchema } from "zod-to-json-schema";
import { readDocTool } from "./tools/readDoc.js";
import { createDocTool } from "./tools/createDoc.js";
import { inspectDocTool } from "./tools/inspectDoc.js";
import { cleanupMediaTool } from "./tools/cleanupMedia.js";
import { surgicalReplaceTool } from "./tools/surgicalReplace.js";
import { surgicalAddImageTool } from "./tools/surgicalAddImage.js";
import { surgicalInsertTool } from "./tools/surgicalInsert.js";

const TOOLS = [readDocTool, createDocTool, inspectDocTool, cleanupMediaTool, surgicalReplaceTool, surgicalAddImageTool, surgicalInsertTool];

const server = new Server(
  {
    name: "docx-master",
    version: "2.2.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: TOOLS.map((tool) => {
      const jsonSchema = zodToJsonSchema(tool.inputSchema) as any;
      const { $schema, ...cleanSchema } = jsonSchema;
      return {
        name: tool.name,
        description: tool.description,
        inputSchema: cleanSchema,
      };
    }),
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const toolName = request.params.name;
  const tool = TOOLS.find((t) => t.name === toolName);

  if (!tool) {
    throw new Error(`Tool not found: ${toolName}`);
  }

  const args = request.params.arguments as any;
  const parsed = tool.inputSchema.safeParse(args);
  if (!parsed.success) {
      throw new Error(`Invalid arguments: ${JSON.stringify(parsed.error.format())}`);
  }

  return tool.handler(parsed.data);
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  
  // Keep the process alive
  setInterval(() => {}, 1000);
}

main().catch((error) => {
  console.error("Server error:", error);
  process.exit(1);
});