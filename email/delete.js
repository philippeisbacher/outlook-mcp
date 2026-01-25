/**
 * Delete email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Delete email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteEmail(args) {
  const emailId = args.id || '';
  const mailbox = args.mailbox || null;

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required. Please provide the ID of the email to delete."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);
    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';

    console.error(`[DELETE-EMAIL] Deleting email ${emailId}${mailboxInfo}`);

    // Delete the email - this moves it to Deleted Items
    await callGraphAPI(
      accessToken,
      'DELETE',
      `${basePath}/messages/${emailId}`,
      null
    );

    return {
      content: [{
        type: "text",
        text: `Email deleted successfully${mailboxInfo}. The email has been moved to Deleted Items.`
      }]
    };
  } catch (error) {
    console.error(`[DELETE-EMAIL] Error: ${error.message}`);

    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    // Handle specific error cases
    if (error.message.includes('404') || error.message.includes('not found')) {
      return {
        content: [{
          type: "text",
          text: `Email not found. The email with ID "${emailId}" may have already been deleted or does not exist.`
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error deleting email: ${error.message}`
      }]
    };
  }
}

module.exports = handleDeleteEmail;
