const handleUpdateEvent = require('../../calendar/update');
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

describe('handleUpdateEvent', () => {
  describe('successful update', () => {
    test('should PATCH /me/events/{id}', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleUpdateEvent({ eventId: 'event-123', subject: 'New Subject' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'PATCH',
        'me/events/event-123',
        expect.any(Object)
      );
    });

    test('should include subject in patch body when provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleUpdateEvent({ eventId: 'event-123', subject: 'Updated Title' });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.subject).toBe('Updated Title');
    });

    test('should only include provided fields in patch body', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleUpdateEvent({ eventId: 'event-123', subject: 'Only Subject' });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.subject).toBe('Only Subject');
      expect(body.start).toBeUndefined();
      expect(body.end).toBeUndefined();
      expect(body.location).toBeUndefined();
    });

    test('should include start and end when provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleUpdateEvent({
        eventId: 'event-123',
        start: '2026-03-01T10:00:00',
        end: '2026-03-01T11:00:00'
      });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.start).toBeDefined();
      expect(body.start.dateTime).toBe('2026-03-01T10:00:00');
      expect(body.end).toBeDefined();
      expect(body.end.dateTime).toBe('2026-03-01T11:00:00');
    });

    test('should include location when provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleUpdateEvent({ eventId: 'event-123', location: 'Konferenzraum A' });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.location).toEqual({ displayName: 'Konferenzraum A' });
    });

    test('should include attendees when provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleUpdateEvent({
        eventId: 'event-123',
        attendees: ['alice@example.com', 'bob@example.com']
      });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.attendees).toHaveLength(2);
      expect(body.attendees[0].emailAddress.address).toBe('alice@example.com');
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleUpdateEvent({ eventId: 'event-123', subject: 'Test', mailbox: 'shared@example.com' });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toContain('shared@example.com');
    });

    test('should return success message', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleUpdateEvent({ eventId: 'event-123', subject: 'New Title' });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('updated');
    });
  });

  describe('validation', () => {
    test('should return error when eventId is missing', async () => {
      const result = await handleUpdateEvent({ subject: 'Test' });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when no fields to update are provided', async () => {
      const result = await handleUpdateEvent({ eventId: 'event-123' });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleUpdateEvent({ eventId: 'event-123', subject: 'Test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('authenticate');
    });

    test('should handle API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('404 Event not found'));

      const result = await handleUpdateEvent({ eventId: 'event-123', subject: 'Test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('404 Event not found');
    });
  });
});
