/**
 * List events functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * List events handler — uses calendarView to correctly expand recurring events
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);
  const mailbox = args.mailbox || null;

  const startDateTime = args.after
    ? new Date(args.after).toISOString()
    : new Date().toISOString();

  // Default window: 30 days from start
  const defaultEnd = new Date(startDateTime);
  defaultEnd.setDate(defaultEnd.getDate() + 30);
  const endDateTime = args.before
    ? new Date(args.before).toISOString()
    : defaultEnd.toISOString();

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);

    // calendarView expands recurring event instances — /events does not
    const endpoint = `${basePath}/calendarView`;

    const queryParams = {
      startDateTime,
      endDateTime,
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
      return {
        content: [{
          type: "text",
          text: `No calendar events found${mailboxInfo}.`
        }]
      };
    }

    const eventList = response.value.map((event, index) => {
      const startDate = new Date(event.start.dateTime).toLocaleString(event.start.timeZone);
      const endDate = new Date(event.end.dateTime).toLocaleString(event.end.timeZone);
      const location = event.location?.displayName || 'No location';

      return `${index + 1}. ${event.subject} - Location: ${location}\nStart: ${startDate}\nEnd: ${endDate}\nSummary: ${event.bodyPreview}\nID: ${event.id}\n`;
    }).join("\n");

    const mailboxInfo = mailbox ? ` (shared mailbox: ${mailbox})` : '';
    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} events${mailboxInfo}:\n\n${eventList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }],
        isError: true
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing events: ${error.message}`
      }],
      isError: true
    };
  }
}

module.exports = handleListEvents;
