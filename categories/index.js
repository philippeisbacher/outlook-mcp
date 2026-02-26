/**
 * Categories module for Outlook MCP server
 */
const handleListCategories = require('./list');
const handleAddCategory = require('./add');
const handleRemoveCategory = require('./remove');
const handleCreateCategory = require('./create');

// Shared mailbox description
const MAILBOX_DESCRIPTION = "Email address of a shared mailbox to access. Leave empty to use your primary mailbox.";

// Category tool definitions
const categoryTools = [
  {
    name: "create-category",
    description: "Creates a new Outlook category (label) in the master category list",
    inputSchema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Display name of the new category"
        },
        color: {
          type: "string",
          description: "Color preset (e.g. 'preset0' to 'preset24', or 'none'). Default: 'none'",
          enum: ["none","preset0","preset1","preset2","preset3","preset4","preset5","preset6","preset7","preset8","preset9","preset10","preset11","preset12","preset13","preset14","preset15","preset16","preset17","preset18","preset19","preset20","preset21","preset22","preset23","preset24"]
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["name"]
    },
    handler: handleCreateCategory
  },
  {
    name: "list-categories",
    description: "Lists all available Outlook categories (labels) that can be applied to emails",
    inputSchema: {
      type: "object",
      properties: {
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: []
    },
    handler: handleListCategories
  },
  {
    name: "add-category",
    description: "Adds a category (label) to an email. Use 'list-categories' to see available categories.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to add the category to"
        },
        category: {
          type: "string",
          description: "Name of the category to add (must match an existing category name)"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["id", "category"]
    },
    handler: handleAddCategory
  },
  {
    name: "remove-category",
    description: "Removes a category (label) from an email",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to remove the category from"
        },
        category: {
          type: "string",
          description: "Name of the category to remove"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["id", "category"]
    },
    handler: handleRemoveCategory
  }
];

module.exports = {
  categoryTools,
  handleListCategories,
  handleAddCategory,
  handleRemoveCategory,
  handleCreateCategory
};
