/**
 * Folder management module for Outlook MCP server
 */
const handleListFolders = require('./list');
const handleCreateFolder = require('./create');
const handleMoveEmails = require('./move');

// Shared mailbox description used across all folder tools
const MAILBOX_DESCRIPTION = "Email address of a shared mailbox to access. Leave empty to use your primary mailbox.";

// Folder management tool definitions
const folderTools = [
  {
    name: "list-folders",
    description: "Lists mail folders in your Outlook account or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        includeItemCounts: {
          type: "boolean",
          description: "Include counts of total and unread items"
        },
        includeChildren: {
          type: "boolean",
          description: "Include child folders in hierarchy"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: []
    },
    handler: handleListFolders
  },
  {
    name: "create-folder",
    description: "Creates a new mail folder in your account or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Name of the folder to create"
        },
        parentFolder: {
          type: "string",
          description: "Optional parent folder name (default is root)"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "move-emails",
    description: "Moves emails from one folder to another in your account or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        emailIds: {
          type: "string",
          description: "Comma-separated list of email IDs to move"
        },
        targetFolder: {
          type: "string",
          description: "Name of the folder to move emails to"
        },
        sourceFolder: {
          type: "string",
          description: "Optional name of the source folder (default is inbox)"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["emailIds", "targetFolder"]
    },
    handler: handleMoveEmails
  }
];

module.exports = {
  folderTools,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails
};
