const { handleReplyEmail, handleForwardEmail } = require('../../email/reply');
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

// ---------------------------------------------------------------------------
// reply-email
// ---------------------------------------------------------------------------
describe('handleReplyEmail', () => {
  describe('successful reply', () => {
    test('should reply to a primary mailbox email', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleReplyEmail({ id: 'msg-1', comment: 'Thanks!' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages/msg-1/reply',
        { comment: 'Thanks!' }
      );
      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('Reply sent successfully');
    });

    test('should reply-all when replyAll flag is true', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleReplyEmail({ id: 'msg-2', comment: 'Noted.', replyAll: true });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages/msg-2/replyAll',
        { comment: 'Noted.' }
      );
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleReplyEmail({ id: 'msg-3', comment: 'Hi', mailbox: 'shared@company.com' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'users/shared@company.com/messages/msg-3/reply',
        { comment: 'Hi' }
      );
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleReplyEmail({ comment: 'test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('id');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when comment is missing', async () => {
      const result = await handleReplyEmail({ id: 'msg-1' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('comment');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleReplyEmail({ id: 'msg-1', comment: 'test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle generic API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('500 Server Error'));

      const result = await handleReplyEmail({ id: 'msg-1', comment: 'test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('500 Server Error');
    });
  });
});

// ---------------------------------------------------------------------------
// forward-email
// ---------------------------------------------------------------------------
describe('handleForwardEmail', () => {
  describe('successful forward', () => {
    test('should forward email to a single recipient', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleForwardEmail({ id: 'msg-4', to: 'bob@example.com' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages/msg-4/forward',
        {
          toRecipients: [{ emailAddress: { address: 'bob@example.com' } }],
          comment: ''
        }
      );
      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('forwarded successfully');
    });

    test('should forward to multiple recipients', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleForwardEmail({ id: 'msg-5', to: 'alice@example.com, bob@example.com' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages/msg-5/forward',
        {
          toRecipients: [
            { emailAddress: { address: 'alice@example.com' } },
            { emailAddress: { address: 'bob@example.com' } }
          ],
          comment: ''
        }
      );
    });

    test('should include comment when provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleForwardEmail({ id: 'msg-6', to: 'carol@example.com', comment: 'FYI' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages/msg-6/forward',
        {
          toRecipients: [{ emailAddress: { address: 'carol@example.com' } }],
          comment: 'FYI'
        }
      );
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleForwardEmail({ id: 'msg-7', to: 'dave@example.com', mailbox: 'shared@company.com' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'users/shared@company.com/messages/msg-7/forward',
        expect.any(Object)
      );
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleForwardEmail({ to: 'bob@example.com' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('id');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when to is missing', async () => {
      const result = await handleForwardEmail({ id: 'msg-1' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('to');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleForwardEmail({ id: 'msg-1', to: 'bob@example.com' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle generic API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('404 Not Found'));

      const result = await handleForwardEmail({ id: 'msg-1', to: 'bob@example.com' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('404 Not Found');
    });
  });
});
