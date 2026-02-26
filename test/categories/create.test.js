const handleCreateCategory = require('../../categories/create');
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

describe('handleCreateCategory', () => {
  describe('successful creation', () => {
    test('should create a category with default color (none)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'cat-1', displayName: 'Work', color: 'none' });

      const result = await handleCreateCategory({ name: 'Work' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/outlook/masterCategories',
        { displayName: 'Work', color: 'none' }
      );
      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('Work');
      expect(result.content[0].text).toContain('created');
    });

    test('should create a category with specified color preset', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'cat-2', displayName: 'Urgent', color: 'preset2' });

      await handleCreateCategory({ name: 'Urgent', color: 'preset2' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/outlook/masterCategories',
        { displayName: 'Urgent', color: 'preset2' }
      );
    });

    test('should use shared mailbox path when mailbox is provided', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ id: 'cat-3', displayName: 'Shared Cat', color: 'none' });

      await handleCreateCategory({ name: 'Shared Cat', mailbox: 'shared@company.com' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'users/shared@company.com/outlook/masterCategories',
        expect.any(Object)
      );
    });
  });

  describe('validation', () => {
    test('should return error when name is missing', async () => {
      const result = await handleCreateCategory({});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('name');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleCreateCategory({ name: 'Test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle API error (e.g. duplicate category)', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('409 Conflict'));

      const result = await handleCreateCategory({ name: 'Existing' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('409 Conflict');
    });
  });
});
