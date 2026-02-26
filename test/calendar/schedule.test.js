const handleGetSchedule = require('../../calendar/schedule');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const mockAccessToken = 'dummy_access_token';

beforeEach(() => {
  callGraphAPI.mockClear();
  ensureAuthenticated.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

const makeScheduleResponse = (email, freeSlotsOnly = false) => ({
  scheduleId: email,
  availabilityView: '0002000',  // 0=free, 2=busy
  scheduleItems: freeSlotsOnly ? [] : [
    {
      status: 'busy',
      start: { dateTime: '2026-03-01T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-03-01T11:00:00', timeZone: 'UTC' },
      subject: 'Team Meeting'
    }
  ]
});

describe('handleGetSchedule', () => {
  describe('successful schedule query', () => {
    test('should POST to /me/calendar/getSchedule with correct payload', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        value: [makeScheduleResponse('alice@example.com')]
      });

      await handleGetSchedule({
        attendees: 'alice@example.com',
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/calendar/getSchedule',
        expect.objectContaining({
          schedules: ['alice@example.com'],
          startTime: expect.objectContaining({ dateTime: '2026-03-01T09:00:00' }),
          endTime: expect.objectContaining({ dateTime: '2026-03-01T17:00:00' })
        })
      );
    });

    test('should support multiple attendees (comma-separated)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        value: [
          makeScheduleResponse('alice@example.com'),
          makeScheduleResponse('bob@example.com')
        ]
      });

      await handleGetSchedule({
        attendees: 'alice@example.com, bob@example.com',
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.schedules).toEqual(['alice@example.com', 'bob@example.com']);
    });

    test('should return formatted availability for each attendee', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        value: [makeScheduleResponse('alice@example.com')]
      });

      const result = await handleGetSchedule({
        attendees: 'alice@example.com',
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('alice@example.com');
      expect(result.content[0].text).toContain('Team Meeting');
    });

    test('should show "free" when attendee has no schedule items', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        value: [makeScheduleResponse('alice@example.com', true)]
      });

      const result = await handleGetSchedule({
        attendees: 'alice@example.com',
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.content[0].text).toContain('free');
    });
  });

  describe('validation', () => {
    test('should return error when attendees is missing', async () => {
      const result = await handleGetSchedule({
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('attendees');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when startTime is missing', async () => {
      const result = await handleGetSchedule({
        attendees: 'alice@example.com',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('startTime');
    });

    test('should return error when endTime is missing', async () => {
      const result = await handleGetSchedule({
        attendees: 'alice@example.com',
        startTime: '2026-03-01T09:00:00'
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endTime');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleGetSchedule({
        attendees: 'alice@example.com',
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('403 Forbidden'));

      const result = await handleGetSchedule({
        attendees: 'alice@example.com',
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('403 Forbidden');
    });
  });
});
