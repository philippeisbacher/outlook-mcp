/**
 * Add category to an email
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Add category to email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAddCategory(args) {
  const emailId = args.id || '';
  const category = args.category || '';
  const mailbox = args.mailbox || null;

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required. Please provide the ID of the email."
      }]
    };
  }

  if (!category) {
    return {
      content: [{
        type: "text",
        text: "Category name is required. Use 'list-categories' to see available categories."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);
    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';

    console.error(`[ADD-CATEGORY] Adding category "${category}" to email ${emailId}${mailboxInfo}`);

    // First, get the current categories of the email
    const email = await callGraphAPI(
      accessToken,
      'GET',
      `${basePath}/messages/${emailId}`,
      { $select: 'categories' }
    );

    const currentCategories = email.categories || [];

    // Check if category already exists
    if (currentCategories.includes(category)) {
      return {
        content: [{
          type: "text",
          text: `Email already has the category "${category}".`
        }]
      };
    }

    // Add the new category
    const updatedCategories = [...currentCategories, category];

    await callGraphAPI(
      accessToken,
      'PATCH',
      `${basePath}/messages/${emailId}`,
      null,
      { categories: updatedCategories }
    );

    return {
      content: [{
        type: "text",
        text: `Category "${category}" added to email${mailboxInfo}. Email now has categories: ${updatedCategories.join(', ')}`
      }]
    };
  } catch (error) {
    console.error(`[ADD-CATEGORY] Error: ${error.message}`);

    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    // Handle category not found error
    if (error.message.includes('does not exist')) {
      return {
        content: [{
          type: "text",
          text: `Category "${category}" does not exist. Use 'list-categories' to see available categories.`
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error adding category: ${error.message}`
      }]
    };
  }
}

module.exports = handleAddCategory;
