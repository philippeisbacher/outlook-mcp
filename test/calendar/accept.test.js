const handleAcceptEvent = require('../../calendar/accept');
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

describe('handleAcceptEvent', () => {
  describe('successful accept', () => {
    test('should POST to /events/{id}/accept', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleAcceptEvent({ eventId: 'event-123' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-123/accept',
        expect.any(Object)
      );
    });

    test('should return success message', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleAcceptEvent({ eventId: 'event-123' });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('event-123');
      expect(result.content[0].text).toContain('accepted');
    });

    test('should include comment in request body', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleAcceptEvent({ eventId: 'event-123', comment: 'Looking forward to it!' });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.comment).toBe('Looking forward to it!');
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleAcceptEvent({ eventId: 'event-123', mailbox: 'shared@example.com' });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toContain('shared@example.com');
    });
  });

  describe('validation', () => {
    test('should return error when eventId is missing', async () => {
      const result = await handleAcceptEvent({});

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleAcceptEvent({ eventId: 'event-123' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('authenticate');
    });

    test('should handle API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('404 Event not found'));

      const result = await handleAcceptEvent({ eventId: 'event-123' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('404 Event not found');
    });
  });
});
