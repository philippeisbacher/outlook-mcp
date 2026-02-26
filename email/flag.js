/**
 * Flag email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

const VALID_STATUSES = ['flagged', 'notFlagged', 'complete'];

/**
 * Flag/unflag an email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleFlagEmail(args) {
  const { id, status = 'flagged', mailbox = null } = args;

  if (!id) {
    return {
      content: [{ type: "text", text: "Email id is required." }],
      isError: true
    };
  }

  if (!VALID_STATUSES.includes(status)) {
    return {
      content: [{ type: "text", text: `Invalid status '${status}'. Allowed values: ${VALID_STATUSES.join(', ')}.` }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);

    await callGraphAPI(accessToken, 'PATCH', `${basePath}/messages/${id}`, {
      flag: { flagStatus: status }
    });

    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
    return {
      content: [{
        type: "text",
        text: `Email successfully marked as ${status}${mailboxInfo}.`
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
      content: [{ type: "text", text: `Error flagging email: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleFlagEmail;
