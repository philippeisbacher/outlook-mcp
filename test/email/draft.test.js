const handleCreateDraft = require('../../email/draft');
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

describe('handleCreateDraft', () => {
  describe('successful draft creation', () => {
    test('should POST to /me/messages (not /sendMail)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'draft-abc' });

      await handleCreateDraft({ to: 'alice@example.com', subject: 'Test', body: 'Hello' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages',
        expect.any(Object)
      );
    });

    test('should NOT call /send — draft must stay unsent', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'draft-abc' });

      await handleCreateDraft({ to: 'alice@example.com', subject: 'Test', body: 'Hello' });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).not.toContain('send');
    });

    test('should return success with draft id', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'draft-abc' });

      const result = await handleCreateDraft({ to: 'alice@example.com', subject: 'Test', body: 'Hello' });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('draft-abc');
    });

    test('should include cc and bcc when provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'draft-abc' });

      await handleCreateDraft({
        to: 'alice@example.com',
        subject: 'Test',
        body: 'Hello',
        cc: 'bob@example.com',
        bcc: 'carol@example.com'
      });

      const [, , , body] = callGraphAPI.mock.calls[0];
      expect(body.ccRecipients).toBeDefined();
      expect(body.bccRecipients).toBeDefined();
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'draft-abc' });

      await handleCreateDraft({
        to: 'alice@example.com',
        subject: 'Test',
        body: 'Hello',
        mailbox: 'shared@example.com'
      });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toContain('shared@example.com');
    });
  });

  describe('validation', () => {
    test('should return error when to is missing', async () => {
      const result = await handleCreateDraft({ subject: 'Test', body: 'Hello' });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when subject is missing', async () => {
      const result = await handleCreateDraft({ to: 'alice@example.com', body: 'Hello' });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when body is missing', async () => {
      const result = await handleCreateDraft({ to: 'alice@example.com', subject: 'Test' });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleCreateDraft({ to: 'alice@example.com', subject: 'Test', body: 'Hello' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('authenticate');
    });

    test('should handle API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('400 Bad Request'));

      const result = await handleCreateDraft({ to: 'alice@example.com', subject: 'Test', body: 'Hello' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('400 Bad Request');
    });
  });
});
