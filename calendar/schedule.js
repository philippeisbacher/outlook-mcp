/**
 * Get schedule / free-busy availability for multiple people
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Get schedule handler — checks availability for one or more people
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetSchedule(args) {
  const { attendees, startTime, endTime, timezone = 'UTC' } = args;

  if (!attendees) {
    return {
      content: [{ type: "text", text: "attendees is required (comma-separated email addresses)." }],
      isError: true
    };
  }

  if (!startTime) {
    return {
      content: [{ type: "text", text: "startTime is required (ISO format, e.g. '2026-03-01T09:00:00')." }],
      isError: true
    };
  }

  if (!endTime) {
    return {
      content: [{ type: "text", text: "endTime is required (ISO format, e.g. '2026-03-01T17:00:00')." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const schedules = attendees.split(',').map(e => e.trim());

    const response = await callGraphAPI(accessToken, 'POST', 'me/calendar/getSchedule', {
      schedules,
      startTime: { dateTime: startTime, timeZone: timezone },
      endTime: { dateTime: endTime, timeZone: timezone },
      availabilityViewInterval: 30
    });

    const results = (response.value || []).map(person => {
      const items = person.scheduleItems || [];
      if (items.length === 0) {
        return `${person.scheduleId}: free during this period`;
      }

      const busyList = items.map(item => {
        const start = new Date(item.start.dateTime).toLocaleString();
        const end = new Date(item.end.dateTime).toLocaleString();
        return `  - [${item.status}] ${item.subject || '(no title)'}: ${start} → ${end}`;
      }).join('\n');

      return `${person.scheduleId}:\n${busyList}`;
    }).join('\n\n');

    return {
      content: [{
        type: "text",
        text: `Schedule for ${schedules.length} attendee(s) from ${startTime} to ${endTime}:\n\n${results}`
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
      content: [{ type: "text", text: `Error getting schedule: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleGetSchedule;
