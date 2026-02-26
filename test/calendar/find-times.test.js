const handleFindMeetingTimes = require('../../calendar/find-times');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const mockAccessToken = 'dummy_access_token';

const makeSuggestion = (start, end) => ({
  meetingTimeSlot: {
    start: { dateTime: start, timeZone: 'UTC' },
    end: { dateTime: end, timeZone: 'UTC' }
  },
  confidence: 100
});

beforeEach(() => {
  callGraphAPI.mockClear();
  ensureAuthenticated.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

describe('handleFindMeetingTimes', () => {
  describe('successful find', () => {
    test('should POST to /me/calendar/findMeetingTimes', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        meetingTimeSuggestions: [makeSuggestion('2026-03-01T10:00:00', '2026-03-01T11:00:00')]
      });

      await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        duration: 60,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/calendar/findMeetingTimes',
        expect.objectContaining({
          attendees: expect.any(Array),
          meetingDuration: 'PT60M'
        })
      );
    });

    test('should include all attendees in request', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ meetingTimeSuggestions: [] });

      await handleFindMeetingTimes({
        attendees: 'alice@example.com, bob@example.com',
        duration: 30,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.attendees).toHaveLength(2);
      expect(body.attendees[0].emailAddress.address).toBe('alice@example.com');
      expect(body.attendees[1].emailAddress.address).toBe('bob@example.com');
    });

    test('should return formatted suggestions', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        meetingTimeSuggestions: [
          makeSuggestion('2026-03-01T10:00:00', '2026-03-01T11:00:00'),
          makeSuggestion('2026-03-01T14:00:00', '2026-03-01T15:00:00')
        ]
      });

      const result = await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        duration: 60,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('2026-03-01T10:00:00');
      expect(result.content[0].text).toContain('2026-03-01T14:00:00');
    });

    test('should return "no slots" when suggestions list is empty', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ meetingTimeSuggestions: [] });

      const result = await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        duration: 60,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.content[0].text).toContain('No available');
    });

    test('should convert duration minutes to ISO 8601 duration', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ meetingTimeSuggestions: [] });

      await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        duration: 90,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.meetingDuration).toBe('PT90M');
    });
  });

  describe('validation', () => {
    test('should return error when attendees is missing', async () => {
      const result = await handleFindMeetingTimes({
        duration: 60,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when duration is missing', async () => {
      const result = await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when startTime is missing', async () => {
      const result = await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        duration: 60,
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        duration: 60,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('authenticate');
    });

    test('should handle API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('400 Bad Request'));

      const result = await handleFindMeetingTimes({
        attendees: 'alice@example.com',
        duration: 60,
        startTime: '2026-03-01T09:00:00',
        endTime: '2026-03-01T17:00:00'
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('400 Bad Request');
    });
  });
});
