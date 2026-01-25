/**
 * Categories module for Outlook MCP server
 */
const handleListCategories = require('./list');
const handleAddCategory = require('./add');
const handleRemoveCategory = require('./remove');

// Shared mailbox description
const MAILBOX_DESCRIPTION = "Email address of a shared mailbox to access. Leave empty to use your primary mailbox.";

// Category tool definitions
const categoryTools = [
  {
    name: "list-categories",
    description: "Lists all available Outlook categories (labels) that can be applied to emails",
    inputSchema: {
      type: "object",
      properties: {},
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
  handleRemoveCategory
};
