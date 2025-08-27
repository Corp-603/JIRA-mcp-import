# Jira MCP Server Max

[![MCP](https://badge.mcpx.dev/default)](https://modelcontextprotocol.io)
[![Claude](https://img.shields.io/badge/Claude-AI-D79943)](https://claude.ai)
[![Cursor](https://img.shields.io/badge/Cursor-Editor-009EFF)](https://cursor.sh)

A powerful Model Context Protocol (MCP) server for Jira integration, providing tools to manage Jira projects, issues, and bulk operations. Speak to Jira in natural language to get information on and modify your project.

Built using the [Model Context Protocol](https://github.com/modelcontextprotocol).

## Steps to Connect Your Jira-MCP-Server-Max in VS Code with Copilot (MCP Integration)

### 1. Clone the Repository
```bash
git clone https://github.com/Nipun-706/Jira-MCP-Server-Max.git
cd Jira-MCP-Server-Max
npm install
npm run build
```

### 2. Open the Project in VS Code
Launch VS Code from the project's root directory.

### 3. Trigger the MCP Setup
- Press `Ctrl + Shift + P` (`Cmd + Shift + P` on Mac) to open the Command Palette
- Search for and select **MCP: Add Server**

### 4. Add Your Server in .vscode/mcp.json
VS Code will open (or create) `.vscode/mcp.json`. Add this configuration:

```json
{
  "servers": {
    "my-mcp-server-2b4beb1a": {
      "name": "jira-server",
      "command": "node",
      "args": ["${workspaceFolder}/build/index.js"],
      "cwd": "${workspaceFolder}",
      "env": {
        "JIRA_HOST": "xxxx",
        "JIRA_EMAIL": "xxx",
        "JIRA_API_TOKEN": "xxx"
      },
      "type": "stdio"
    }
  }
}
```

**Key notes:**
- Use `"node"` as the command, so it relies on your installed Node.js version
- Use `${workspaceFolder}` to avoid hard-coded paths and make it portable
- Keep sensitive values like API tokens out of source control—consider using environment variable placeholders or .env files

### 5. Save the File and Add the Server
After saving, you should see a prompt like "Add Server" or an indicator that VS Code has detected your MCP server. Click to connect.

### 6. MCP Server Should Be Listed
In the Copilot / MCP panel, you should now see `jira-server` with its available tools (e.g., `get_projects`, `get_issues`).

### 7. Interact with the Server
Now you can ask Copilot to, for example:
- "List all projects in Jira"
- "Get issues in project CCS"

and it will leverage your MCP server to execute those commands.

## Steps to Connect with Cline (Alternative MCP Client)

### 1. Clone and Setup Repository
```bash
git clone https://github.com/Nipun-706/Jira-MCP-Server-Max.git
cd Jira-MCP-Server-Max
npm install
npm run build
```

### 2. Open Cline in VS Code
- Install the Cline extension in VS Code
- Open your project in VS Code

### 3. Configure Cline MCP Settings
Create or edit `cline_mcp_settings.json` in your VS Code workspace with the following configuration:

```json
{
  "mcpServers": {
    "jira-server": {
      "timeout": 60,
      "command": "D:/node/node.exe",
      "args": [
        "D:/coding/jira/Jira-MCP-Server/build/index.js"
      ],
      "cwd": "D:/coding/jira/Jira-MCP-Server",
      "env": {
        "JIRA_HOST": "your-domain.atlassian.net",
        "JIRA_EMAIL": "your-email@example.com",
        "JIRA_API_TOKEN": "your-api-token"
      },
      "type": "stdio",
      "disabled": false
    }
  }
}
```

**Configuration Notes:**
- Replace `D:/node/node.exe` with your actual Node.js executable path
- Update the `args` and `cwd` paths to match your project location
- Replace placeholder values in `env` with your actual Jira credentials
- Set `timeout` to 60 seconds to handle longer operations

### 4. Apply Configuration
1. Save the `cline_mcp_settings.json` file
2. Restart VS Code (or reload window: `Ctrl+Shift+P` → **Reload Window**)
3. In Cline's MCP panel, verify that `jira-server` appears as enabled
4. Check logs in Cline Output if the server doesn't connect

### 5. Test Integration
Once connected, you can interact with Cline using natural language:
- "Show me all issues in the CCS project"
- "Create a new bug report for the login issue"
- "List all high-priority tasks assigned to me"

## Steps to Connect with Windsurf

### 1. Clone and Setup Repository
```bash
git clone https://github.com/Nipun-706/Jira-MCP-Server-Max.git
cd Jira-MCP-Server-Max
npm install
npm run build
```

### 2. Configure Windsurf MCP Settings
Add the following configuration to your Windsurf MCP settings under the `"mcpServers"` object:

```json
{
  "mcpServers": {
    "jira-server": {
      "type": "stdio",
      "command": "D:/node/node.exe",
      "args": [
        "D:/coding/jira/Jira-MCP-Server/build/index.js"
      ],
      "cwd": "D:/coding/jira/Jira-MCP-Server",
      "timeout": 60,
      "env": {
        "JIRA_HOST": "your-domain.atlassian.net",
        "JIRA_EMAIL": "your-email@example.com",
        "JIRA_API_TOKEN": "your-api-token-here"
      },
      "disabledTools": [],
      "disabled": false
    }
  }
}
```

**Configuration Notes:**
- Replace `D:/node/node.exe` with your actual Node.js executable path
- Update the `args` and `cwd` paths to match your project location
- Replace placeholder values in `env` with your actual Jira credentials
- Set `timeout` to 60 seconds to handle longer operations
- Keep `disabledTools` empty to enable all available tools

### 3. Apply Configuration and Test
1. Save your Windsurf MCP configuration
2. Restart Windsurf to apply the changes
3. Verify that `jira-server` appears in your MCP servers list
4. Test with natural language commands like:
   - "List all projects in my Jira instance"
   - "Show me open issues in project CCS"
   - "Create a new task for API integration"

## Available Tools

The Jira MCP Server provides the following tools:

### Core Tools
- **get_projects** - List all Jira projects
- **get_issues** - Get project issues with JQL filtering  
- **create_issues_bulk** - Create multiple Jira issues at once

## Configuration

Required environment variables:

- `JIRA_HOST`: Your Jira instance hostname
- `JIRA_EMAIL`: Your Jira account email
- `JIRA_API_TOKEN`: API token from https://id.atlassian.com/manage-profile/security/api-tokens

### 1. User Management

```typescript
// Get user's account ID by email
{
  email: "user@example.com";
}
```

### 2. Issue Type Management

```typescript
// List all available issue types
// Returns: id, name, description, subtask status
// No parameters required
```

### 3. Issue Link Types

```typescript
// List all available issue link types
// Returns: id, name, inward/outward descriptions
// No parameters required
```

### 4. Issue Management

#### Retrieving Issues

```typescript
// Get all issues in a project
{
  projectKey: "PROJECT"
}

// Get issues with JQL filtering
{
  projectKey: "PROJECT",
  jql: "status = 'In Progress' AND assignee = currentUser()"
}

// Get issues assigned to user
{
  projectKey: "PROJECT",
  jql: "assignee = 'user@example.com' ORDER BY created DESC"
}
```

#### Creating Issues

```typescript
// Create a standard issue
{
  projectKey: "PROJECT",
  summary: "Issue title",
  issueType: "Task",  // or "Story", "Bug", etc.
  description: "Detailed description",
  assignee: "accountId",  // from get_user tool
  labels: ["frontend", "urgent"],
  components: ["ui", "api"],
  priority: "High"
}

// Create a subtask
{
  parent: "PROJECT-123",
  projectKey: "PROJECT",
  summary: "Subtask title",
  issueType: "Subtask",
  description: "Subtask details",
  assignee: "accountId"
}
```

#### Updating Issues

```typescript
// Update issue fields
{
  issueKey: "PROJECT-123",
  summary: "Updated title",
  description: "New description",
  assignee: "accountId",
  status: "In Progress",
  priority: "High"
}
```

#### Issue Dependencies

```typescript
// Create issue link
{
  linkType: "Blocks",  // from list_link_types
  inwardIssueKey: "PROJECT-124",  // blocked issue
  outwardIssueKey: "PROJECT-123"  // blocking issue
}
```

#### Deleting Issues

```typescript
// Delete single issue
{
  issueKey: "PROJECT-123"
}

// Delete issue with subtasks
{
  issueKey: "PROJECT-123",
  deleteSubtasks: true
}

// Delete multiple issues
{
  issueKeys: ["PROJECT-123", "PROJECT-124"]
}
```

## Field Formatting

### Description Field

The description field supports markdown-style formatting:

- Use blank lines between paragraphs
- Use "- " for bullet points
- Use "1. " for numbered lists
- Use headers ending with ":" (followed by blank line)

Example:

```
Task Overview:

This task involves implementing new features:
- Feature A implementation
- Feature B testing

Steps:
1. Design component
2. Implement logic
3. Add tests

Acceptance Criteria:
- All tests passing
- Documentation updated
```

## Error Handling

The server provides detailed error messages for:

- Invalid issue keys
- Missing required fields
- Permission issues
- API rate limits

## Setup Instructions

1. Clone the repository:

   ```bash
   git clone https://github.com/George5562/Jira-MCP-Server.git
   cd Jira-MCP-Server
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Configure environment variables:
   Create a `.env` file in the root directory:

   ```bash
   JIRA_HOST=your-instance.atlassian.net
   JIRA_EMAIL=your-email@example.com
   JIRA_API_TOKEN=your-api-token
   ```

4. Build the project:

   ```bash
   npm run build
   ```

5. Start the server:
   ```bash
   npm start
   ```

## Configuring Claude Desktop

To use this MCP server with Claude Desktop:

1. Locate your Claude Desktop configuration file:

   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%/Claude/claude_desktop_config.json`
   - Linux: `~/.config/Claude/claude_desktop_config.json`

2. Add the Jira MCP server to your configuration:

   ```json
   {
     "mcpServers": {
       "jira-server": {
         "name": "jira-server",
         "command": "/path/to/node",
         "args": ["/path/to/jira-server/build/index.js"],
         "cwd": "/path/to/jira-server",
         "env": {
           "JIRA_HOST": "your-jira-instance.atlassian.net",
           "JIRA_EMAIL": "your-email@example.com",
           "JIRA_API_TOKEN": "your-api-token"
         }
       }
     }
   }
   ```

   Replace `/path/to/jira-server` with the absolute path to your cloned repository.
   Replace `/path/to/node` with the absolute path to your Node.js executable (you can usually find this by running `which node` or `where node` in your terminal).
   Using the direct path to the Node.js executable and the built JavaScript file (`build/index.js` after running `npm run build`) is recommended for reliability.

3. Restart Claude Desktop to apply the changes.

## Configuring Cursor

To use this Jira MCP server with Cursor:

1.  **Ensure the server is built:** Run `npm run build` in the `Jira-MCP-Server` directory to create the necessary `build/index.js` file.
2.  **Locate or create Cursor's MCP configuration file:**
    - For project-specific configuration: `.cursor/mcp.json` in your project's root directory.
    - For global configuration (all projects): `~/.cursor/mcp.json` in your home directory.
3.  **Add the Jira MCP server configuration to `mcp.json`:**

    ```json
    {
      "mcpServers": {
        "jira-mcp-server": {
          "command": "node", // Or provide the absolute path to your Node.js executable
          "args": [
            "/path/to/your/Jira-MCP-Server/build/index.js" // Absolute path to the server's built index.js
          ],
          "cwd": "/path/to/your/Jira-MCP-Server", // Absolute path to the Jira-MCP-Server directory
          "env": {
            "JIRA_HOST": "your-jira-instance.atlassian.net",
            "JIRA_EMAIL": "your-email@example.com", // Your Jira email
            "JIRA_API_TOKEN": "your-api-token" // Your Jira API token
          }
        }
        // You can add other MCP server configurations here
      }
    }
    ```

    - Replace `/path/to/your/Jira-MCP-Server` with the correct absolute path to where you cloned the `Jira-MCP-Server` repository.
    - If `node` is not in your system's PATH or you prefer an absolute path, replace `"node"` with the full path to your Node.js executable (e.g., `/usr/local/bin/node` or `C:\Program Files\nodejs\node.exe`).
    - Ensure your Jira instance details and API token are correctly filled in the `env` section.

4.  **Restart Cursor** to apply the changes.

### Using Cursor Rules for Jira Context

To make interacting with Jira smoother, you can define your default Jira project and user identifier in Cursor's rules. This helps Cursor's AI understand your context without you needing to specify it in every prompt.

Create or edit your Cursor Rules file (e.g., in your project `.cursor/rules.json` or global `~/.cursor/rules.json` (the exact file and method for rules might vary, check Cursor documentation for "Rules" or "Context Management")). Add entries like:

```
As an AI assistant, when I am asked about Jira tasks:
- Assume the primary Jira project key is 'YOUR_PROJECT_KEY_HERE'.
- Assume 'my assigned tasks' or tasks assigned to 'me' refer to the Jira user with the email 'your_jira_email@example.com' (or your Jira Account ID).
You can then use these in your JQL queries, for example: project = YOUR_PROJECT_KEY_HERE AND assignee = 'your_jira_email@example.com'.
```

Replace `YOUR_PROJECT_KEY_HERE` and `your_jira_email@example.com` with your actual details.

### Example Usage in Cursor Chat

Once configured (especially with Cursor Rules for context), you can ask Cursor:

`"Using Jira MCP, list my assigned tasks. Then, based on these tasks, come up with an implementation plan and work schedule."`

If you haven't set up rules, or need to specify a different project or user, you'd be more explicit:

`"Using Jira MCP, list tasks assigned to 'user@example.com' in project 'PROJECT_KEY'. Then, based on these tasks, come up with an implementation plan and work schedule."`

Cursor's AI will use the Jira MCP server to fetch the tasks, and then proceed with the planning and scheduling request.

### setup for Amazon Q
one time step 
git clone https://github.com/Nipun-706/Jira-MCP-Server-Max.git
cd Jira-MCP-Server-Max
npm install
npm run build

use below settings as per screenshot

<img width="428" height="617" alt="amzaon q" src="https://github.com/user-attachments/assets/ef1d3c6c-9b11-46f6-b128-1dd21326e74e" />

change red higlighted values as mentioned below

keep name, transport, arguments same as screenshot
Command -> add nodes exe location path of your systm
 

## References

- [Model Context Protocol](https://github.com/modelcontextprotocol)
- [Jira REST API Documentation](https://docs.atlassian.com/software/jira/docs/api/REST/7.12.0)
- [Jira REST API Examples](https://developer.atlassian.com/server/jira/platform/jira-rest-api-examples/)
