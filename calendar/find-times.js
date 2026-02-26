/**
 * Find meeting times functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Find meeting times handler — POST /me/calendar/findMeetingTimes
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleFindMeetingTimes(args) {
  const { attendees, duration, startTime, endTime, timezone = 'UTC' } = args;

  if (!attendees) {
    return { content: [{ type: "text", text: "attendees is required (comma-separated email addresses)." }], isError: true };
  }

  if (!duration) {
    return { content: [{ type: "text", text: "duration is required (meeting length in minutes)." }], isError: true };
  }

  if (!startTime) {
    return { content: [{ type: "text", text: "startTime is required (ISO 8601 format)." }], isError: true };
  }

  if (!endTime) {
    return { content: [{ type: "text", text: "endTime is required (ISO 8601 format)." }], isError: true };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const attendeeList = attendees.split(',').map(addr => ({
      emailAddress: { address: addr.trim() },
      type: 'required'
    }));

    const body = {
      attendees: attendeeList,
      meetingDuration: `PT${duration}M`,
      timeConstraint: {
        timeslots: [{
          start: { dateTime: startTime, timeZone: timezone },
          end: { dateTime: endTime, timeZone: timezone }
        }]
      },
      maxCandidates: 10,
      isOrganizerOptional: false
    };

    const response = await callGraphAPI(accessToken, 'POST', 'me/calendar/findMeetingTimes', body);
    const suggestions = response.meetingTimeSuggestions || [];

    if (suggestions.length === 0) {
      return {
        content: [{ type: "text", text: `No available meeting slots found for ${duration}-minute meeting in the given time range.` }]
      };
    }

    const slots = suggestions.map((s, i) => {
      const slot = s.meetingTimeSlot;
      return `${i + 1}. ${slot.start.dateTime} → ${slot.end.dateTime} (${slot.start.timeZone})`;
    }).join('\n');

    return {
      content: [{
        type: "text",
        text: `Found ${suggestions.length} available slot(s) for a ${duration}-minute meeting:\n\n${slots}`
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
      content: [{ type: "text", text: `Error finding meeting times: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleFindMeetingTimes;
