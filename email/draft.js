/**
 * Create email draft functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Create draft handler — saves email as draft without sending
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateDraft(args) {
  const { to, subject, body, cc, bcc, mailbox = null } = args;

  if (!to) {
    return {
      content: [{ type: "text", text: "Recipient (to) is required." }],
      isError: true
    };
  }

  if (!subject) {
    return {
      content: [{ type: "text", text: "Subject is required." }],
      isError: true
    };
  }

  if (!body) {
    return {
      content: [{ type: "text", text: "Body is required." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const toAddresses = to.split(',').map(addr => ({
      emailAddress: { address: addr.trim() }
    }));

    const message = {
      subject,
      body: { contentType: 'HTML', content: body },
      toRecipients: toAddresses
    };

    if (cc) {
      message.ccRecipients = cc.split(',').map(addr => ({
        emailAddress: { address: addr.trim() }
      }));
    }

    if (bcc) {
      message.bccRecipients = bcc.split(',').map(addr => ({
        emailAddress: { address: addr.trim() }
      }));
    }

    const basePath = getMailboxBasePath(mailbox);
    const response = await callGraphAPI(accessToken, 'POST', `${basePath}/messages`, message);

    const mailboxInfo = mailbox ? ` in shared mailbox ${mailbox}` : '';
    return {
      content: [{
        type: "text",
        text: `Draft created${mailboxInfo}. Draft ID: ${response.id}`
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
      content: [{ type: "text", text: `Error creating draft: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleCreateDraft;
