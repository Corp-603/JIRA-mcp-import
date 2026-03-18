#!/usr/bin/env node

import * as dotenv from 'dotenv';
dotenv.config();

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import JiraClient from "jira-client";
import type { Request } from "@modelcontextprotocol/sdk/types.js";
import * as fs from "fs";
import * as path from "path";
import ExcelJS from "exceljs";

const DEFAULT_PROJECT = {
  KEY: "CCS",
  NAME: "Chat System",
};

// Environment variables with validation
const JIRA_HOST = process.env.JIRA_HOST ?? "";
const JIRA_EMAIL = process.env.JIRA_EMAIL ?? "";
const JIRA_API_TOKEN = process.env.JIRA_API_TOKEN ?? "";

if (!JIRA_HOST || !JIRA_EMAIL || !JIRA_API_TOKEN) {
  throw new Error(
    "Missing required environment variables: JIRA_HOST, JIRA_EMAIL, and JIRA_API_TOKEN"
  );
}

interface GetIssuesArgs {
  projectKey: string;
  jql?: string;
}

interface CreateIssuesBulkArgs {
  issues: Array<{
    summary: string;
    issueType: string;
    projectKey: string;
    description?: string;
    assignee?: string;
    priority?: string;
    labels?: string[];
    components?: string[];
    parent?: string;
    platform?: string;
  }>;
}

interface GetIssueArgs {
  issueKey: string;
}

interface SearchIssuesArgs {
  jql: string;
  maxResults?: number;
}

interface UpdateIssueArgs {
  issueKey: string;
  summary?: string;
  description?: string;
  assignee?: string;
  priority?: string;
  labels?: string[];
  components?: string[];
}

interface TransitionIssueArgs {
  issueKey: string;
  transitionId: string;
  comment?: string;
}

interface AddCommentArgs {
  issueKey: string;
  comment: string;
}

interface AddAttachmentArgs {
  issueKey: string;
  filePath: string;
}

interface GenerateExcelArgs {
  fileName: string;
  sheetName?: string;
  columns: Array<{ header: string; key: string; width?: number }>;
  rows: Array<Record<string, string>>;
  outputDir?: string;
}

interface ToolDefinition {
  description: string;
  inputSchema: object;
}

class JiraServer {
  private readonly server: Server;
  private readonly jira: JiraClient;
  private readonly toolDefinitions: Record<string, ToolDefinition>;

  constructor() {
    this.toolDefinitions = {
      get_projects: {
        description: "List all Jira projects",
        inputSchema: {
          type: "object",
          properties: {},
          additionalProperties: false
        }
      },
      get_issues: {
        description: "Get project issues",
        inputSchema: {
          type: "object",
          properties: {
            projectKey: { type: "string" },
            jql: { type: "string" }
          },
          required: ["projectKey"]
        }
      },
      create_issues_bulk: {
        description: "Create multiple Jira issues at once",
        inputSchema: {
          type: "object",
          properties: {
            issues: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  summary: { type: "string" },
                  issueType: { type: "string" },
                  projectKey: { type: "string" },
                  description: { type: "string" },
                  assignee: { type: "string" },
                  priority: { type: "string" },
                  labels: { type: "array", items: { type: "string" } },
                  components: { type: "array", items: { type: "string" } },
                  parent: { type: "string" },
                  platform: { type: "string" }
                },
                required: ["summary", "issueType", "projectKey"]
              }
            }
          },
          required: ["issues"]
        }
      },
      jira_get_issue: {
        description: "Get details of a specific issue",
        inputSchema: {
          type: "object",
          properties: {
            issueKey: { type: "string" }
          },
          required: ["issueKey"]
        }
      },
      jira_search: {
        description: "Search issues using JQL",
        inputSchema: {
          type: "object",
          properties: {
            jql: { type: "string" },
            maxResults: { type: "number" }
          },
          required: ["jql"]
        }
      },
      jira_update_issue: {
        description: "Update an existing issue",
        inputSchema: {
          type: "object",
          properties: {
            issueKey: { type: "string" },
            summary: { type: "string" },
            description: { type: "string" },
            assignee: { type: "string" },
            priority: { type: "string" },
            labels: { type: "array", items: { type: "string" } },
            components: { type: "array", items: { type: "string" } }
          },
          required: ["issueKey"]
        }
      },
      jira_transition_issue: {
        description: "Transition an issue to a new status",
        inputSchema: {
          type: "object",
          properties: {
            issueKey: { type: "string" },
            transitionId: { type: "string" },
            comment: { type: "string" }
          },
          required: ["issueKey", "transitionId"]
        }
      },
      jira_add_comment: {
        description: "Add a comment to an issue",
        inputSchema: {
          type: "object",
          properties: {
            issueKey: { type: "string" },
            comment: { type: "string" }
          },
          required: ["issueKey", "comment"]
        }
      },
      jira_add_attachment: {
        description: "Add a file attachment to a Jira issue",
        inputSchema: {
          type: "object",
          properties: {
            issueKey: { type: "string" },
            filePath: { type: "string" }
          },
          required: ["issueKey", "filePath"]
        }
      },
      generate_excel: {
        description: "Generate an Excel (.xlsx) file with custom columns and rows. Returns the file path of the generated file.",
        inputSchema: {
          type: "object",
          properties: {
            fileName: { type: "string", description: "Name of the Excel file (without extension)" },
            sheetName: { type: "string", description: "Name of the worksheet (default: Sheet1)" },
            columns: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  header: { type: "string" },
                  key: { type: "string" },
                  width: { type: "number" }
                },
                required: ["header", "key"]
              },
              description: "Column definitions with header, key, and optional width"
            },
            rows: {
              type: "array",
              items: { type: "object" },
              description: "Array of row objects where keys match column keys"
            },
            outputDir: { type: "string", description: "Output directory path (default: current working directory)" }
          },
          required: ["fileName", "columns", "rows"]
        }
      }
    };

    this.server = new Server(
      { name: "jira-server", version: "0.1.0" },
      { 
        capabilities: { tools: this.toolDefinitions }
      }
    );

    this.jira = new JiraClient({
      protocol: "https",
      host: JIRA_HOST,
      username: JIRA_EMAIL,
      password: JIRA_API_TOKEN,
      apiVersion: "3",
      strictSSL: false,
    });

    this.setupToolHandlers();
    
    this.server.onerror = (error) => console.error("[MCP Error]", error);
    process.on("SIGINT", async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private async jiraFetch(endpoint: string, options: RequestInit = {}): Promise<any> {
    const url = `https://${JIRA_HOST}${endpoint}`;
    const response = await fetch(url, {
      ...options,
      headers: {
        'Authorization': `Basic ${Buffer.from(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`).toString('base64')}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        ...options.headers
      }
    });
    return response.json();
  }

  private setupToolHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async (request: Request) => ({
      tools: Object.entries(this.toolDefinitions).map(([name, def]) => ({
        name,
        description: def.description,
        inputSchema: def.inputSchema
      })),
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request: Request) => {
      try {
        if (!request.params?.name) {
          throw new McpError(ErrorCode.InvalidParams, "Tool name is required");
        }

        switch (request.params.name) {
          case "get_projects":
            const projects = await this.jira.listProjects();
            return {
              content: [{
                type: "text",
                text: JSON.stringify(projects.map(p => ({ key: p.key, name: p.name })))
              }]
            };

          case "get_issues":
            const args = request.params.arguments as GetIssuesArgs;
            if (!args?.projectKey) {
              throw new McpError(ErrorCode.InvalidParams, "projectKey is required");
            }
            
            const jql = `project = ${args.projectKey}${args.jql ? ` AND ${args.jql}` : ''}`;
            const issuesResult = await this.jiraFetch('/rest/api/3/search/jql', {
              method: 'POST',
              body: JSON.stringify({ jql, maxResults: 100, fields: ["summary", "status", "assignee", "priority", "issuetype", "created", "updated", "sprint", "description", "reporter", "duedate", "customfield_10015"] })
            });
            return {
              content: [{
                type: "text",
                text: JSON.stringify(issuesResult.issues || issuesResult)
              }]
            };

          case "create_issues_bulk":
            const bulkArgs = request.params.arguments as CreateIssuesBulkArgs;
            if (!bulkArgs?.issues || !Array.isArray(bulkArgs.issues)) {
              throw new McpError(ErrorCode.InvalidParams, "issues array is required");
            }

            const results = await Promise.all(
              bulkArgs.issues.map(async (issue) => {
                try {
                  const issueData: any = {
                    fields: {
                      project: { key: issue.projectKey },
                      summary: issue.summary,
                      issuetype: { name: issue.issueType },
                      description: {
                        type: "doc",
                        version: 1,
                        content: [
                          {
                            type: "paragraph",
                            content: [
                              {
                                type: "text",
                                text: issue.description || ""
                              }
                            ]
                          }
                        ]
                      }
                    }
                  };

                  if (issue.assignee) {
                    issueData.fields.assignee = { accountId: issue.assignee };
                  }
                  if (issue.priority) {
                    issueData.fields.priority = { name: issue.priority };
                  }
                  if (issue.labels && issue.labels.length > 0) {
                    issueData.fields.labels = issue.labels;
                  }
                  if (issue.components && issue.components.length > 0) {
                    issueData.fields.components = issue.components.map(c => ({ name: c }));
                  }
                  if (issue.parent) {
                    issueData.fields.parent = { key: issue.parent };
                  }
                  if (issue.platform) {
                    const platformMap: Record<string, string> = {
                      'Android': '10131',
                      'IOS': '10132',
                      'WEB': '10145',
                      'API or Service': '10146',
                      'Flutter': '10272',
                      'Testing': '10282'
                    };
                    const platformId = platformMap[issue.platform] || issue.platform;
                    issueData.fields.customfield_10072 = { id: platformId };
                  }

                  const createdIssue = await this.jira.addNewIssue(issueData);
                  return {
                    success: true,
                    issue: {
                      key: createdIssue.key,
                      id: createdIssue.id,
                      summary: issue.summary
                    }
                  };
                } catch (error) {
                  return {
                    success: false,
                    error: error instanceof Error ? error.message : 'Unknown error',
                    summary: issue.summary
                  };
                }
              })
            );

            return {
              content: [{
                type: "text",
                text: JSON.stringify({ message: "Bulk issue creation completed", results }, null, 2)
              }]
            };

          case "jira_get_issue":
            const getIssueArgs = request.params.arguments as GetIssueArgs;
            if (!getIssueArgs?.issueKey) {
              throw new McpError(ErrorCode.InvalidParams, "issueKey is required");
            }
            
            const issue = await this.jira.findIssue(getIssueArgs.issueKey);
            return {
              content: [{
                type: "text",
                text: JSON.stringify(issue, null, 2)
              }]
            };

          case "jira_search":
            const searchArgs = request.params.arguments as SearchIssuesArgs;
            if (!searchArgs?.jql) {
              throw new McpError(ErrorCode.InvalidParams, "jql is required");
            }
            
            const maxResults = searchArgs.maxResults || 50;
            const searchResults = await this.jiraFetch('/rest/api/3/search/jql', {
              method: 'POST',
              body: JSON.stringify({ jql: searchArgs.jql, maxResults, fields: ["summary", "status", "assignee", "priority", "issuetype", "created", "updated", "labels", "components", "sprint", "description", "reporter", "duedate", "customfield_10015"] })
            });
            return {
              content: [{
                type: "text",
                text: JSON.stringify(searchResults, null, 2)
              }]
            };

          case "jira_update_issue":
            const updateArgs = request.params.arguments as UpdateIssueArgs;
            if (!updateArgs?.issueKey) {
              throw new McpError(ErrorCode.InvalidParams, "issueKey is required");
            }

            const updateData: any = { fields: {} };
            if (updateArgs.summary) updateData.fields.summary = updateArgs.summary;
            if (updateArgs.description) {
              updateData.fields.description = {
                type: "doc",
                version: 1,
                content: [{
                  type: "paragraph",
                  content: [{
                    type: "text",
                    text: updateArgs.description
                  }]
                }]
              };
            }
            if (updateArgs.assignee) updateData.fields.assignee = { accountId: updateArgs.assignee };
            if (updateArgs.priority) updateData.fields.priority = { name: updateArgs.priority };
            if (updateArgs.labels) updateData.fields.labels = updateArgs.labels;
            if (updateArgs.components) updateData.fields.components = updateArgs.components.map(c => ({ name: c }));

            await this.jira.updateIssue(updateArgs.issueKey, updateData);
            return {
              content: [{
                type: "text",
                text: JSON.stringify({ message: "Issue updated successfully", issueKey: updateArgs.issueKey })
              }]
            };

          case "jira_transition_issue":
            const transitionArgs = request.params.arguments as TransitionIssueArgs;
            if (!transitionArgs?.issueKey || !transitionArgs?.transitionId) {
              throw new McpError(ErrorCode.InvalidParams, "issueKey and transitionId are required");
            }

            const transitionData: any = {
              transition: { id: transitionArgs.transitionId }
            };
            
            if (transitionArgs.comment) {
              transitionData.update = {
                comment: [{
                  add: {
                    body: {
                      type: "doc",
                      version: 1,
                      content: [{
                        type: "paragraph",
                        content: [{
                          type: "text",
                          text: transitionArgs.comment
                        }]
                      }]
                    }
                  }
                }]
              };
            }

            await this.jira.transitionIssue(transitionArgs.issueKey, transitionData);
            return {
              content: [{
                type: "text",
                text: JSON.stringify({ 
                  message: "Issue transitioned successfully", 
                  issueKey: transitionArgs.issueKey,
                  transitionId: transitionArgs.transitionId
                })
              }]
            };

          case "jira_add_comment":
            const commentArgs = request.params.arguments as AddCommentArgs;
            if (!commentArgs?.issueKey || !commentArgs?.comment) {
              throw new McpError(ErrorCode.InvalidParams, "issueKey and comment are required");
            }

            // Use jira-client's built-in addComment method
            const addedComment = await this.jira.addComment(commentArgs.issueKey, commentArgs.comment);
            
            return {
              content: [{
                type: "text",
                text: JSON.stringify({ 
                  message: "Comment added successfully", 
                  issueKey: commentArgs.issueKey,
                  commentId: addedComment.id,
                  comment: commentArgs.comment
                }, null, 2)
              }]
            };

          case "jira_add_attachment":
            const attachArgs = request.params.arguments as AddAttachmentArgs;
            if (!attachArgs?.issueKey || !attachArgs?.filePath) {
              throw new McpError(ErrorCode.InvalidParams, "issueKey and filePath are required");
            }

            const resolvedPath = path.resolve(attachArgs.filePath);
            if (!fs.existsSync(resolvedPath)) {
              throw new McpError(ErrorCode.InvalidParams, `File not found: ${resolvedPath}`);
            }

            const readStream = fs.createReadStream(resolvedPath);
            const attachResult = await this.jira.addAttachmentOnIssue(attachArgs.issueKey, readStream);

            return {
              content: [{
                type: "text",
                text: JSON.stringify({
                  message: "Attachment added successfully",
                  issueKey: attachArgs.issueKey,
                  filePath: resolvedPath,
                  attachment: attachResult
                }, null, 2)
              }]
            };

          case "generate_excel":
            const excelArgs = request.params.arguments as GenerateExcelArgs;
            if (!excelArgs?.fileName || !excelArgs?.columns || !excelArgs?.rows) {
              throw new McpError(ErrorCode.InvalidParams, "fileName, columns, and rows are required");
            }

            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet(excelArgs.sheetName || 'Sheet1');

            ws.columns = excelArgs.columns.map(col => ({
              header: col.header,
              key: col.key,
              width: col.width || 30
            }));

            // Style header row
            ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };
            ws.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

            for (const row of excelArgs.rows) {
              ws.addRow(row);
            }

            // Auto-wrap text for all rows
            ws.eachRow((row) => { row.alignment = { ...row.alignment, wrapText: true }; });

            const outDir = excelArgs.outputDir ? path.resolve(excelArgs.outputDir) : process.cwd();
            if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

            const excelFilePath = path.join(outDir, `${excelArgs.fileName}.xlsx`);
            await wb.xlsx.writeFile(excelFilePath);

            return {
              content: [{
                type: "text",
                text: JSON.stringify({
                  message: "Excel file generated successfully",
                  filePath: excelFilePath,
                  fileName: `${excelArgs.fileName}.xlsx`,
                  rowCount: excelArgs.rows.length,
                  columnCount: excelArgs.columns.length
                }, null, 2)
              }]
            };

          default:
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
        }
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error: ${error instanceof Error ? error.message : 'Unknown error'}`
          }],
          isError: true
        };
      }
    });
  }

  public async run(): Promise<void> {
    await this.server.connect(new StdioServerTransport());
    console.error("Jira MCP server running on stdio");
  }
}

const jiraServer = new JiraServer();
jiraServer.run().catch((error: Error) => console.error(error));
