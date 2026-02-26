/**
 * Create a new Outlook master category
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Create category handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateCategory(args) {
  const { name, color = 'none', mailbox = null } = args;

  if (!name) {
    return {
      content: [{ type: "text", text: "Category name is required." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);

    const category = await callGraphAPI(
      accessToken, 'POST', `${basePath}/outlook/masterCategories`,
      { displayName: name, color }
    );

    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
    return {
      content: [{
        type: "text",
        text: `Category '${category.displayName}' created successfully${mailboxInfo} (color: ${category.color}).`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }],
        isError: true
      };
    }

    return {
      content: [{ type: "text", text: `Error creating category: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleCreateCategory;
