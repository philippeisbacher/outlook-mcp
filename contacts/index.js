/**
 * Contacts module for Outlook MCP server
 */
const handleSearchPeople = require('./search');
const {
  handleListContacts,
  handleGetContact,
  handleCreateContact,
  handleUpdateContact,
  handleDeleteContact
} = require('./crud');

const contactsTools = [
  {
    name: "search-people",
    description: "Searches for people (colleagues, contacts) by name or email fragment. Returns email addresses, job titles, and departments. Useful for finding someone's email address.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Name or email fragment to search for (e.g. 'Alice' or 'smith@')"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        }
      },
      required: ["query"]
    },
    handler: handleSearchPeople
  },
  {
    name: "list-contacts",
    description: "Lists contacts from your personal contacts folder, sorted alphabetically.",
    inputSchema: {
      type: "object",
      properties: {
        count: { type: "number", description: "Number of contacts to return (default: 25, max: 50)" }
      },
      required: []
    },
    handler: handleListContacts
  },
  {
    name: "get-contact",
    description: "Retrieves full details for a specific contact by ID.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Contact ID (from list-contacts or search-people)" }
      },
      required: ["id"]
    },
    handler: handleGetContact
  },
  {
    name: "create-contact",
    description: "Creates a new contact in your personal contacts folder.",
    inputSchema: {
      type: "object",
      properties: {
        displayName: { type: "string", description: "Full name of the contact" },
        email: { type: "string", description: "Email address" },
        jobTitle: { type: "string", description: "Job title" },
        department: { type: "string", description: "Department" },
        mobilePhone: { type: "string", description: "Mobile phone number" }
      },
      required: ["displayName"]
    },
    handler: handleCreateContact
  },
  {
    name: "update-contact",
    description: "Updates an existing contact. Only the provided fields are changed.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Contact ID to update" },
        displayName: { type: "string", description: "New full name" },
        email: { type: "string", description: "New email address" },
        jobTitle: { type: "string", description: "New job title" },
        department: { type: "string", description: "New department" },
        mobilePhone: { type: "string", description: "New mobile phone number" }
      },
      required: ["id"]
    },
    handler: handleUpdateContact
  },
  {
    name: "delete-contact",
    description: "Permanently deletes a contact from your contacts folder.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Contact ID to delete" }
      },
      required: ["id"]
    },
    handler: handleDeleteContact
  }
];

module.exports = {
  contactsTools,
  handleSearchPeople,
  handleListContacts,
  handleGetContact,
  handleCreateContact,
  handleUpdateContact,
  handleDeleteContact
};
