/**
 * List emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const mailbox = args.mailbox || null;

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Resolve the folder path (with optional shared mailbox support)
    const endpoint = await resolveFolderPath(accessToken, folder, mailbox);
    
    // Add query parameters
    const queryParams = {
      $top: Math.min(50, requestedCount), // Use 50 per page for efficiency
      $orderby: 'receivedDateTime desc',
      $select: config.EMAIL_SELECT_FIELDS
    };
    
    // Make API call with pagination support
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, queryParams, requestedCount);
    
    if (!response.value || response.value.length === 0) {
      const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
      return {
        content: [{
          type: "text",
          text: `No emails found in ${folder}${mailboxInfo}.`
        }]
      };
    }

    // Format results
    const emailList = response.value.map((email, index) => {
      const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
      const date = new Date(email.receivedDateTime).toLocaleString();
      const readStatus = email.isRead ? '' : '[UNREAD] ';
      const categories = email.categories && email.categories.length > 0
        ? `[${email.categories.join(', ')}] `
        : '';

      return `${index + 1}. ${readStatus}${categories}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
    }).join("\n");

    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} emails in ${folder}${mailboxInfo}:\n\n${emailList}`
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
        text: `Error listing emails: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEmails;
