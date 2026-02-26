const handleListEvents = require('../../calendar/list');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleListEvents', () => {
  const mockAccessToken = 'dummy_access_token';

  const makeEvent = (overrides = {}) => ({
    id: 'event-1',
    subject: 'Team Meeting',
    bodyPreview: 'Discuss Q1 results',
    start: { dateTime: '2024-03-10T10:00:00', timeZone: 'UTC' },
    end:   { dateTime: '2024-03-10T11:00:00', timeZone: 'UTC' },
    location: { displayName: 'Conference Room A' },
    ...overrides
  });

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('calendarView endpoint', () => {
    test('should use calendarView endpoint (not /events) to expand recurring events', async () => {
      callGraphAPI.mockResolvedValue({ value: [makeEvent()] });

      await handleListEvents({});

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toContain('calendarView');
      expect(endpoint).not.toContain('/events');
    });

    test('should pass startDateTime and endDateTime as query params (not $filter)', async () => {
      callGraphAPI.mockResolvedValue({ value: [makeEvent()] });

      await handleListEvents({ after: '2026-03-01', before: '2026-03-31' });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.startDateTime).toContain('2026-03-01');
      expect(params.endDateTime).toContain('2026-03-31');
      expect(params.$filter).toBeUndefined();
    });

    test('should default endDateTime to ~30 days from startDateTime when before is not provided', async () => {
      callGraphAPI.mockResolvedValue({ value: [makeEvent()] });
      const before = new Date();

      await handleListEvents({});

      const after = new Date();
      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.startDateTime).toBeDefined();
      expect(params.endDateTime).toBeDefined();

      const start = new Date(params.startDateTime);
      const end = new Date(params.endDateTime);
      const diffDays = (end - start) / (1000 * 60 * 60 * 24);
      expect(diffDays).toBeCloseTo(30, 0);
    });
  });

  describe('event rendering', () => {
    test('should list events with location', async () => {
      callGraphAPI.mockResolvedValue({ value: [makeEvent()] });

      const result = await handleListEvents({ count: 5 });

      expect(result.content[0].text).toContain('Team Meeting');
      expect(result.content[0].text).toContain('Conference Room A');
    });

    test('Bug: should not crash when event has no location (location is null)', async () => {
      callGraphAPI.mockResolvedValue({
        value: [makeEvent({ location: null })]
      });

      const result = await handleListEvents({ count: 5 });

      expect(result.content[0].text).toContain('Team Meeting');
      expect(result.content[0].text).toContain('No location');
    });

    test('Bug: should not crash when event location has no displayName', async () => {
      callGraphAPI.mockResolvedValue({
        value: [makeEvent({ location: {} })]
      });

      const result = await handleListEvents({ count: 5 });

      expect(result.content[0].text).toContain('No location');
    });

    test('should return "no events" message when API returns empty list', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('No calendar events found');
    });
  });
});
