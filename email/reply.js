/**
 * Reply and forward email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Reply to an email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleReplyEmail(args) {
  const { id, comment, replyAll = false, mailbox = null } = args;

  if (!id) {
    return {
      content: [{ type: "text", text: "Email id is required." }],
      isError: true
    };
  }

  if (!comment) {
    return {
      content: [{ type: "text", text: "Reply comment (body) is required." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);
    const action = replyAll ? 'replyAll' : 'reply';

    await callGraphAPI(accessToken, 'POST', `${basePath}/messages/${id}/${action}`, { comment });

    const mailboxInfo = mailbox ? ` (from shared mailbox: ${mailbox})` : '';
    return {
      content: [{
        type: "text",
        text: `Reply sent successfully${mailboxInfo}!`
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
      content: [{ type: "text", text: `Error sending reply: ${error.message}` }],
      isError: true
    };
  }
}

/**
 * Forward an email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleForwardEmail(args) {
  const { id, to, comment = '', mailbox = null } = args;

  if (!id) {
    return {
      content: [{ type: "text", text: "Email id is required." }],
      isError: true
    };
  }

  if (!to) {
    return {
      content: [{ type: "text", text: "Recipient (to) is required." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);

    const toRecipients = to.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    }));

    await callGraphAPI(accessToken, 'POST', `${basePath}/messages/${id}/forward`, {
      toRecipients,
      comment
    });

    const mailboxInfo = mailbox ? ` (from shared mailbox: ${mailbox})` : '';
    return {
      content: [{
        type: "text",
        text: `Email forwarded successfully to ${toRecipients.length} recipient(s)${mailboxInfo}!`
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
      content: [{ type: "text", text: `Error forwarding email: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = { handleReplyEmail, handleForwardEmail };
