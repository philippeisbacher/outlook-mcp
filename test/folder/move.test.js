const handleMoveEmail = require('../../folder/move-single');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { getFolderIdByName } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

const mockAccessToken = 'dummy_access_token';

beforeEach(() => {
  callGraphAPI.mockClear();
  ensureAuthenticated.mockClear();
  getFolderIdByName.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

describe('handleMoveEmail', () => {
  describe('successful move', () => {
    test('should POST to /messages/{id}/move with destinationId', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      getFolderIdByName.mockResolvedValue('folder-abc');
      callGraphAPI.mockResolvedValue({});

      await handleMoveEmail({ id: 'msg-123', targetFolder: 'Archive' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages/msg-123/move',
        { destinationId: 'folder-abc' }
      );
    });

    test('should return success message with folder name', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      getFolderIdByName.mockResolvedValue('folder-abc');
      callGraphAPI.mockResolvedValue({});

      const result = await handleMoveEmail({ id: 'msg-123', targetFolder: 'Archive' });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('Archive');
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      getFolderIdByName.mockResolvedValue('folder-abc');
      callGraphAPI.mockResolvedValue({});

      await handleMoveEmail({ id: 'msg-123', targetFolder: 'Archive', mailbox: 'shared@example.com' });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toContain('shared@example.com');
    });

    test('should return error when folder not found', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      getFolderIdByName.mockResolvedValue(null);

      const result = await handleMoveEmail({ id: 'msg-123', targetFolder: 'NoSuchFolder' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('NoSuchFolder');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleMoveEmail({ targetFolder: 'Archive' });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when targetFolder is missing', async () => {
      const result = await handleMoveEmail({ id: 'msg-123' });

      expect(result.isError).toBe(true);
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleMoveEmail({ id: 'msg-123', targetFolder: 'Archive' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('authenticate');
    });

    test('should handle API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      getFolderIdByName.mockResolvedValue('folder-abc');
      callGraphAPI.mockRejectedValue(new Error('404 Message not found'));

      const result = await handleMoveEmail({ id: 'msg-123', targetFolder: 'Archive' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('404 Message not found');
    });
  });
});
