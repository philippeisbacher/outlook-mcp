const handleSearchEvents = require('../../calendar/search');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const mockAccessToken = 'dummy_access_token';

const makeEvent = (subject, id = 'evt-1') => ({
  id,
  subject,
  start: { dateTime: '2026-03-01T10:00:00', timeZone: 'UTC' },
  end: { dateTime: '2026-03-01T11:00:00', timeZone: 'UTC' },
  location: { displayName: 'Room A' },
  bodyPreview: 'preview text'
});

beforeEach(() => {
  callGraphAPI.mockClear();
  ensureAuthenticated.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

describe('handleSearchEvents', () => {
  describe('calendarView endpoint', () => {
    test('should use calendarView endpoint to expand recurring events', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchEvents({ query: 'test' });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toContain('calendarView');
      expect(endpoint).not.toContain('/events');
    });

    test('should pass after as startDateTime query param (not $filter)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchEvents({ after: '2026-03-01' });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.startDateTime).toContain('2026-03-01');
      // $filter should not contain date conditions (may be undefined or a subject filter)
      expect(params.$filter || '').not.toContain('start/dateTime ge');
    });

    test('should pass before as endDateTime query param (not $filter)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchEvents({ before: '2026-03-31' });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.endDateTime).toContain('2026-03-31');
      expect(params.$filter || '').not.toContain('start/dateTime lt');
    });

    test('text query should still use $filter=contains() alongside calendarView date params', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchEvents({ query: 'Sync', after: '2026-03-01', before: '2026-03-31' });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.startDateTime).toContain('2026-03-01');
      expect(params.endDateTime).toContain('2026-03-31');
      expect(params.$filter).toContain("contains(subject, 'Sync')");
      // Date must NOT be in $filter
      expect(params.$filter).not.toContain('start/dateTime');
    });
  });

  describe('text search', () => {
    test('should search events by subject using $filter contains()', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [makeEvent('Team Meeting')] });

      await handleSearchEvents({ query: 'Team' });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.$filter).toContain("contains(subject, 'Team')");
    });

    test('should escape single quotes in query to prevent OData injection', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchEvents({ query: "O'Brien" });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.$filter).toContain("O''Brien");
    });

    test('should return formatted event list on success', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [makeEvent('Budget Review')] });

      const result = await handleSearchEvents({ query: 'Budget' });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('Budget Review');
      expect(result.content[0].text).toContain('evt-1');
    });

    test('should return no-results message when list is empty', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleSearchEvents({ query: 'nonexistent' });

      expect(result.content[0].text).toContain('No calendar events found');
    });
  });

  describe('count and mailbox', () => {
    test('should respect count parameter', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchEvents({ count: 5 });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.$top).toBe(5);
    });

    test('should use shared mailbox endpoint when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchEvents({ query: 'test', mailbox: 'shared@company.com' });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toContain('users/shared@company.com');
      expect(endpoint).toContain('calendarView');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleSearchEvents({ query: 'test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle generic API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('500 Server Error'));

      const result = await handleSearchEvents({ query: 'test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('500 Server Error');
    });
  });
});
