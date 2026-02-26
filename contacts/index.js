/**
 * Contacts module for Outlook MCP server
 */
const handleSearchPeople = require('./search');

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
  }
];

module.exports = {
  contactsTools,
  handleSearchPeople
};
