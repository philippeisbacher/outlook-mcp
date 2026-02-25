/**
 * Attachment functionality for emails
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * List attachments handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListAttachments(args) {
  const emailId = args.id;
  const mailbox = args.mailbox || null;

  if (!emailId) {
    return {
      content: [{ type: "text", text: "Email ID is required." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);
    const endpoint = `${basePath}/messages/${encodeURIComponent(emailId)}/attachments`;

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, {
      $select: 'id,name,contentType,size'
    });

    const attachments = response.value || [];

    if (attachments.length === 0) {
      return {
        content: [{ type: "text", text: "No attachments found." }]
      };
    }

    const list = attachments.map(a =>
      `- ${a.name} (${a.contentType}, ${(a.size / 1024).toFixed(1)} KB) | ID: ${a.id}`
    ).join('\n');

    return {
      content: [{
        type: "text",
        text: `Found ${attachments.length} attachment(s):\n\n${list}`
      }]
    };
  } catch (error) {
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
        text: `Error listing attachments: ${error.message}`
      }]
    };
  }
}

/**
 * Get attachment handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetAttachment(args) {
  const emailId = args.emailId;
  const attachmentId = args.attachmentId;
  const mailbox = args.mailbox || null;

  if (!emailId || !attachmentId) {
    return {
      content: [{ type: "text", text: "Email ID and Attachment ID are required." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);
    const endpoint = `${basePath}/messages/${encodeURIComponent(emailId)}/attachments/${encodeURIComponent(attachmentId)}`;

    const attachment = await callGraphAPI(accessToken, 'GET', endpoint, null, {});

    const { name, contentType, contentBytes, size } = attachment;

    // Images: return as image content block so Claude can see them
    if (contentType && contentType.startsWith('image/')) {
      return {
        content: [
          {
            type: "text",
            text: `Attachment: ${name} (${contentType}, ${(size / 1024).toFixed(1)} KB)`
          },
          {
            type: "image",
            data: contentBytes,
            mimeType: contentType
          }
        ]
      };
    }

    // Text files: decode base64 and return as text
    if (contentType && contentType.startsWith('text/')) {
      const decoded = Buffer.from(contentBytes, 'base64').toString('utf-8');
      return {
        content: [{
          type: "text",
          text: `Attachment: ${name}\n\n${decoded}`
        }]
      };
    }

    // Binary files (PDF, Excel, etc.): return base64 with metadata
    return {
      content: [{
        type: "text",
        text: `Attachment: ${name} (${contentType}, ${(size / 1024).toFixed(1)} KB)\nBase64 content:\n${contentBytes}`
      }]
    };
  } catch (error) {
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
        text: `Error getting attachment: ${error.message}`
      }]
    };
  }
}

module.exports = { handleListAttachments, handleGetAttachment };
