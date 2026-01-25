/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleMarkAsRead = require('./mark-as-read');
const handleDeleteEmail = require('./delete');

// Shared mailbox description used across all email tools
const MAILBOX_DESCRIPTION = "Email address of a shared mailbox to access. Leave empty to use your primary mailbox.";

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50)"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria in your inbox or a shared mailbox. Supports date filtering and sort order.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        category: {
          type: "string",
          description: "Filter by category/label name (use 'list-categories' to see available categories)"
        },
        before: {
          type: "string",
          description: "Filter emails received before this date. Supports ISO dates (2024-01-15), or relative: 'today', 'yesterday', '7 days ago', '2 weeks ago', '1 month ago'"
        },
        after: {
          type: "string",
          description: "Filter emails received after this date. Supports ISO dates (2024-01-15), or relative: 'today', 'yesterday', '7 days ago', '2 weeks ago', '1 month ago'"
        },
        sortOrder: {
          type: "string",
          enum: ["asc", "desc"],
          description: "Sort order: 'asc' for oldest first, 'desc' for newest first (default: 'desc')"
        },
        skip: {
          type: "number",
          description: "Number of emails to skip for pagination (e.g., skip=50 to get emails 51-100)"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email from your inbox or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email from your account or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread in your inbox or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "delete-email",
    description: "Deletes an email (moves it to Deleted Items) from your inbox or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to delete"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["id"]
    },
    handler: handleDeleteEmail
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleMarkAsRead,
  handleDeleteEmail
};
