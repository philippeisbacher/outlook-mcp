const handleFlagEmail = require('../../email/flag');
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

describe('handleFlagEmail', () => {
  describe('successful flag operations', () => {
    test('should flag an email (default)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleFlagEmail({ id: 'msg-1' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'PATCH',
        'me/messages/msg-1',
        { flag: { flagStatus: 'flagged' } }
      );
      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('flagged');
    });

    test('should unflag an email when status is "notFlagged"', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleFlagEmail({ id: 'msg-2', status: 'notFlagged' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'PATCH',
        'me/messages/msg-2',
        { flag: { flagStatus: 'notFlagged' } }
      );
      expect(result.content[0].text).toContain('notFlagged');
    });

    test('should mark email as complete when status is "complete"', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleFlagEmail({ id: 'msg-3', status: 'complete' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'PATCH',
        'me/messages/msg-3',
        { flag: { flagStatus: 'complete' } }
      );
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      await handleFlagEmail({ id: 'msg-4', mailbox: 'shared@company.com' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'PATCH',
        'users/shared@company.com/messages/msg-4',
        expect.any(Object)
      );
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleFlagEmail({});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('id');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error for invalid status value', async () => {
      const result = await handleFlagEmail({ id: 'msg-1', status: 'invalid' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('status');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleFlagEmail({ id: 'msg-1' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle generic API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('404 Not Found'));

      const result = await handleFlagEmail({ id: 'msg-1' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('404 Not Found');
    });
  });
});
