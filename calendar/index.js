/**
 * Calendar module for Outlook MCP server
 */
const handleListEvents = require('./list');
const handleSearchEvents = require('./search');
const handleGetSchedule = require('./schedule');
const handleAcceptEvent = require('./accept');
const handleDeclineEvent = require('./decline');
const handleCreateEvent = require('./create');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');
const handleUpdateEvent = require('./update');

// Shared mailbox description used across all calendar tools
const MAILBOX_DESCRIPTION = "Email address of a shared mailbox to access. Leave empty to use your primary calendar.";

// Calendar tool definitions
const calendarTools = [
  {
    name: "list-events",
    description: "Lists upcoming events from your calendar or a shared mailbox calendar",
    inputSchema: {
      type: "object",
      properties: {
        count: {
          type: "number",
          description: "Number of events to retrieve (default: 10, max: 50)"
        },
        after: {
          type: "string",
          description: "Show events starting after this date (ISO format, e.g. '2026-03-01'). Defaults to now."
        },
        before: {
          type: "string",
          description: "Show events starting before this date (ISO format, e.g. '2026-03-31')"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: []
    },
    handler: handleListEvents
  },
  {
    name: "search-events",
    description: "Search calendar events by subject text, attendee, location, and/or date range. Uses calendarView to correctly expand recurring events.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Text to search for in event subjects"
        },
        attendee: {
          type: "string",
          description: "Filter events by attendee email address (e.g. 'alice@example.com')"
        },
        location: {
          type: "string",
          description: "Filter events by location name (e.g. 'Konferenzraum')"
        },
        after: {
          type: "string",
          description: "Filter events starting after this date (ISO format, e.g. '2026-03-01'). Defaults to now."
        },
        before: {
          type: "string",
          description: "Filter events starting before this date (ISO format, e.g. '2026-03-31'). Defaults to 30 days from after."
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: []
    },
    handler: handleSearchEvents
  },
  {
    name: "get-schedule",
    description: "Checks free/busy availability for one or more people over a time range. Useful for finding meeting slots.",
    inputSchema: {
      type: "object",
      properties: {
        attendees: {
          type: "string",
          description: "Comma-separated email addresses to check (e.g. 'alice@example.com, bob@example.com')"
        },
        startTime: {
          type: "string",
          description: "Start of the time range (ISO format, e.g. '2026-03-01T09:00:00')"
        },
        endTime: {
          type: "string",
          description: "End of the time range (ISO format, e.g. '2026-03-01T17:00:00')"
        },
        timezone: {
          type: "string",
          description: "IANA timezone for interpreting times (default: 'UTC', e.g. 'Europe/Berlin')"
        }
      },
      required: ["attendees", "startTime", "endTime"]
    },
    handler: handleGetSchedule
  },
  {
    name: "accept-event",
    description: "Accepts a calendar event invitation in your calendar or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to accept"
        },
        comment: {
          type: "string",
          description: "Optional comment to send with the acceptance"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["eventId"]
    },
    handler: handleAcceptEvent
  },
  {
    name: "decline-event",
    description: "Declines a calendar event from your calendar or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to decline"
        },
        comment: {
          type: "string",
          description: "Optional comment for declining the event"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["eventId"]
    },
    handler: handleDeclineEvent
  },
  {
    name: "create-event",
    description: "Creates a new calendar event in your calendar or a shared mailbox calendar",
    inputSchema: {
      type: "object",
      properties: {
        subject: {
          type: "string",
          description: "The subject of the event"
        },
        start: {
          type: "string",
          description: "The start time of the event in ISO 8601 format"
        },
        end: {
          type: "string",
          description: "The end time of the event in ISO 8601 format"
        },
        attendees: {
          type: "array",
          items: {
            type: "string"
          },
          description: "List of attendee email addresses"
        },
        body: {
          type: "string",
          description: "Optional body content for the event"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
  },
  {
    name: "cancel-event",
    description: "Cancels a calendar event from your calendar or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to cancel"
        },
        comment: {
          type: "string",
          description: "Optional comment for cancelling the event"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["eventId"]
    },
    handler: handleCancelEvent
  },
  {
    name: "update-event",
    description: "Updates an existing calendar event. Only the provided fields are changed.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to update"
        },
        subject: {
          type: "string",
          description: "New subject/title for the event"
        },
        start: {
          type: "string",
          description: "New start time (ISO 8601, e.g. '2026-03-01T10:00:00')"
        },
        end: {
          type: "string",
          description: "New end time (ISO 8601, e.g. '2026-03-01T11:00:00')"
        },
        body: {
          type: "string",
          description: "New body content (HTML or plain text)"
        },
        location: {
          type: "string",
          description: "New location name (e.g. 'Konferenzraum A')"
        },
        attendees: {
          type: "array",
          items: { type: "string" },
          description: "New list of attendee email addresses (replaces existing attendees)"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["eventId"]
    },
    handler: handleUpdateEvent
  },
  {
    name: "delete-event",
    description: "Deletes a calendar event from your calendar or a shared mailbox",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to delete"
        },
        mailbox: {
          type: "string",
          description: MAILBOX_DESCRIPTION
        }
      },
      required: ["eventId"]
    },
    handler: handleDeleteEvent
  }
];

module.exports = {
  calendarTools,
  handleListEvents,
  handleSearchEvents,
  handleGetSchedule,
  handleAcceptEvent,
  handleDeclineEvent,
  handleCreateEvent,
  handleCancelEvent,
  handleDeleteEvent,
  handleUpdateEvent
};
