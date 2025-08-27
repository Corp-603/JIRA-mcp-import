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
                  parent: { type: "string" }
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
      strictSSL: true,
    });

    this.setupToolHandlers();
    
    this.server.onerror = (error) => console.error("[MCP Error]", error);
    process.on("SIGINT", async () => {
      await this.server.close();
      process.exit(0);
    });
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
            const issues = await this.jira.searchJira(jql, { maxResults: 100 });
            return {
              content: [{
                type: "text",
                text: JSON.stringify(issues.issues)
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
            
            const searchResults = await this.jira.searchJira(searchArgs.jql, { 
              maxResults: searchArgs.maxResults || 50 
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
