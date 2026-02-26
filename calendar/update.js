/**
 * Update event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { DEFAULT_TIMEZONE } = require('../config');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Update event handler — patches only provided fields
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateEvent(args) {
  const { eventId, subject, start, end, body, attendees, location, mailbox = null } = args;

  if (!eventId) {
    return {
      content: [{ type: "text", text: "Event ID is required." }],
      isError: true
    };
  }

  const patch = {};

  if (subject !== undefined) patch.subject = subject;
  if (start !== undefined) patch.start = { dateTime: start, timeZone: DEFAULT_TIMEZONE };
  if (end !== undefined) patch.end = { dateTime: end, timeZone: DEFAULT_TIMEZONE };
  if (body !== undefined) patch.body = { contentType: 'HTML', content: body };
  if (location !== undefined) patch.location = { displayName: location };
  if (attendees !== undefined) {
    patch.attendees = attendees.map(email => ({
      emailAddress: { address: email },
      type: 'required'
    }));
  }

  if (Object.keys(patch).length === 0) {
    return {
      content: [{ type: "text", text: "No fields to update. Provide at least one of: subject, start, end, body, attendees, location." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const basePath = getMailboxBasePath(mailbox);
    await callGraphAPI(accessToken, 'PATCH', `${basePath}/events/${eventId}`, patch);

    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
    return {
      content: [{ type: "text", text: `Event ${eventId} updated successfully${mailboxInfo}.` }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }],
        isError: true
      };
    }

    return {
      content: [{ type: "text", text: `Error updating event: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleUpdateEvent;
