/**
 * Search calendar events functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getMailboxBasePath } = require('../utils/mailbox-path');
const { escapeODataString } = require('../utils/odata-helpers');

/**
 * Search calendar events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);
  const mailbox = args.mailbox || null;
  const query = args.query || null;
  const after = args.after || null;
  const before = args.before || null;

  try {
    const accessToken = await ensureAuthenticated();
    const basePath = getMailboxBasePath(mailbox);
    const endpoint = `${basePath}/events`;

    const filterConditions = [];

    if (query) {
      filterConditions.push(`contains(subject, '${escapeODataString(query)}')`);
    }

    if (after) {
      const afterDate = new Date(after).toISOString();
      filterConditions.push(`start/dateTime ge '${afterDate}'`);
    }

    if (before) {
      const beforeDate = new Date(before).toISOString();
      filterConditions.push(`start/dateTime lt '${beforeDate}'`);
    }

    const queryParams = {
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS
    };

    if (filterConditions.length > 0) {
      queryParams.$filter = filterConditions.join(' and ');
    }

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
        text: `Error searching events: ${error.message}`
      }],
      isError: true
    };
  }
}

module.exports = handleSearchEvents;
