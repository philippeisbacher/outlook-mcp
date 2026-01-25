/**
 * Remove category from an email
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Remove category from email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveCategory(args) {
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
        text: "Category name is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);
    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';

    console.error(`[REMOVE-CATEGORY] Removing category "${category}" from email ${emailId}${mailboxInfo}`);

    // First, get the current categories of the email
    const email = await callGraphAPI(
      accessToken,
      'GET',
      `${basePath}/messages/${emailId}`,
      { $select: 'categories' }
    );

    const currentCategories = email.categories || [];

    // Check if category exists on the email
    if (!currentCategories.includes(category)) {
      return {
        content: [{
          type: "text",
          text: `Email does not have the category "${category}". Current categories: ${currentCategories.length > 0 ? currentCategories.join(', ') : 'none'}`
        }]
      };
    }

    // Remove the category
    const updatedCategories = currentCategories.filter(c => c !== category);

    await callGraphAPI(
      accessToken,
      'PATCH',
      `${basePath}/messages/${emailId}`,
      null,
      { categories: updatedCategories }
    );

    const remainingText = updatedCategories.length > 0
      ? `Remaining categories: ${updatedCategories.join(', ')}`
      : 'Email has no more categories.';

    return {
      content: [{
        type: "text",
        text: `Category "${category}" removed from email${mailboxInfo}. ${remainingText}`
      }]
    };
  } catch (error) {
    console.error(`[REMOVE-CATEGORY] Error: ${error.message}`);

    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error removing category: ${error.message}`
      }]
    };
  }
}

module.exports = handleRemoveCategory;
