/**
 * Move a single email to a folder
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getFolderIdByName } = require('../email/folder-utils');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Move a single email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveEmail(args) {
  const { id, targetFolder, mailbox = null } = args;

  if (!id) {
    return {
      content: [{ type: "text", text: "Email ID is required." }],
      isError: true
    };
  }

  if (!targetFolder) {
    return {
      content: [{ type: "text", text: "Target folder name is required." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const destinationId = await getFolderIdByName(accessToken, targetFolder, mailbox);
    if (!destinationId) {
      return {
        content: [{ type: "text", text: `Folder "${targetFolder}" not found. Please specify a valid folder name.` }],
        isError: true
      };
    }

    const basePath = getMailboxBasePath(mailbox);
    await callGraphAPI(accessToken, 'POST', `${basePath}/messages/${id}/move`, { destinationId });

    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
    return {
      content: [{ type: "text", text: `Email moved to "${targetFolder}"${mailboxInfo}.` }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }],
        isError: true
      };
    }

    return {
      content: [{ type: "text", text: `Error moving email: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleMoveEmail;
