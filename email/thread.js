/**
 * Get email thread functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');
const { escapeODataString } = require('../utils/odata-helpers');

/**
 * Get all messages in an email thread
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetEmailThread(args) {
  const { id, mailbox = null } = args;

  if (!id) {
    return {
      content: [{ type: "text", text: "Email id is required." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);

    // Step 1: fetch the anchor email to get its conversationId
    const anchor = await callGraphAPI(
      accessToken, 'GET', `${basePath}/messages/${id}`,
      null, { $select: 'id,conversationId' }
    );

    const conversationId = anchor.conversationId;

    // Step 2: fetch all messages with the same conversationId
    const response = await callGraphAPI(
      accessToken, 'GET', `${basePath}/messages`,
      null,
      {
        $filter: `conversationId eq '${escapeODataString(conversationId)}'`,
        $orderby: 'receivedDateTime asc',
        $select: 'id,subject,from,receivedDateTime,bodyPreview,isRead'
      }
    );

    const messages = response.value || [];

    if (messages.length === 0) {
      return {
        content: [{ type: "text", text: "No messages found in this thread." }]
      };
    }

    const threadText = messages.map((msg, index) => {
      const sender = msg.from?.emailAddress
        ? `${msg.from.emailAddress.name} (${msg.from.emailAddress.address})`
        : 'Unknown';
      const date = new Date(msg.receivedDateTime).toLocaleString();
      const readStatus = msg.isRead ? '' : '[UNREAD] ';

      return `${index + 1}. ${readStatus}${date}\nFrom: ${sender}\nSubject: ${msg.subject}\nPreview: ${msg.bodyPreview}\nID: ${msg.id}`;
    }).join('\n\n');

    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
    return {
      content: [{
        type: "text",
        text: `Thread with ${messages.length} messages${mailboxInfo}:\n\n${threadText}`
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
      content: [{ type: "text", text: `Error getting email thread: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleGetEmailThread;
