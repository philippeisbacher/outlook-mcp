const handleGetEmailThread = require('../../email/thread');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const mockAccessToken = 'dummy_access_token';

const makeEmail = (id, subject, date, fromName = 'Alice') => ({
  id,
  subject,
  conversationId: 'conv-abc',
  receivedDateTime: date,
  from: { emailAddress: { name: fromName, address: `${fromName.toLowerCase()}@example.com` } },
  bodyPreview: `Preview of ${subject}`,
  isRead: true
});

beforeEach(() => {
  callGraphAPI.mockClear();
  ensureAuthenticated.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

describe('handleGetEmailThread', () => {
  describe('successful thread retrieval', () => {
    test('should fetch conversationId from email, then fetch all thread messages', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      // First call: fetch the anchor email
      callGraphAPI.mockResolvedValueOnce({ id: 'msg-1', conversationId: 'conv-abc' });
      // Second call: fetch thread
      callGraphAPI.mockResolvedValueOnce({
        value: [
          makeEmail('msg-1', 'Hello', '2026-01-01T10:00:00Z', 'Alice'),
          makeEmail('msg-2', 'Re: Hello', '2026-01-01T11:00:00Z', 'Bob')
        ]
      });

      const result = await handleGetEmailThread({ id: 'msg-1' });

      // First call: get the anchor email for conversationId
      expect(callGraphAPI.mock.calls[0][2]).toContain('me/messages/msg-1');
      // Second call: filter by conversationId
      expect(callGraphAPI.mock.calls[1][4].$filter).toContain("conversationId eq 'conv-abc'");
      expect(callGraphAPI.mock.calls[1][4].$orderby).toBe('receivedDateTime asc');

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('2 messages');
      expect(result.content[0].text).toContain('Hello');
      expect(result.content[0].text).toContain('Re: Hello');
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValueOnce({ id: 'msg-1', conversationId: 'conv-xyz' });
      callGraphAPI.mockResolvedValueOnce({ value: [makeEmail('msg-1', 'Test', '2026-01-01T10:00:00Z')] });

      await handleGetEmailThread({ id: 'msg-1', mailbox: 'shared@company.com' });

      expect(callGraphAPI.mock.calls[0][2]).toContain('users/shared@company.com/messages/msg-1');
      expect(callGraphAPI.mock.calls[1][2]).toContain('users/shared@company.com/messages');
    });

    test('should show "no messages" when thread is empty', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValueOnce({ id: 'msg-1', conversationId: 'conv-empty' });
      callGraphAPI.mockResolvedValueOnce({ value: [] });

      const result = await handleGetEmailThread({ id: 'msg-1' });

      expect(result.content[0].text).toContain('No messages');
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleGetEmailThread({});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('id');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleGetEmailThread({ id: 'msg-1' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle API error on first call (fetch anchor email)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValueOnce(new Error('404 Not Found'));

      const result = await handleGetEmailThread({ id: 'msg-1' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('404 Not Found');
    });
  });
});
