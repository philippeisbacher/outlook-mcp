const handleListCategories = require('../../categories/list');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { getMailboxBasePath } = require('../../utils/mailbox-path');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../utils/mailbox-path');

describe('handleListCategories', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    getMailboxBasePath.mockClear();
    getMailboxBasePath.mockImplementation((mailbox) => mailbox ? `users/${mailbox}` : 'me');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should list categories successfully', async () => {
    const mockCategories = {
      value: [
        { displayName: 'Red Category', color: 'preset0' },
        { displayName: 'Blue Category', color: 'preset1' },
        { displayName: 'Green Category', color: 'preset2' }
      ]
    };

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(mockCategories);

    const result = await handleListCategories({});

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/outlook/masterCategories',
      { $top: 50 }
    );
    expect(result.content[0].text).toContain('Found 3 categories');
    expect(result.content[0].text).toContain('Red Category');
    expect(result.content[0].text).toContain('Blue Category');
  });

  test('should handle empty category list', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListCategories({});

    expect(result.content[0].text).toContain('No categories found');
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleListCategories({});

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });

  test('should handle API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('API Error'));

    const result = await handleListCategories({});

    expect(result.content[0].text).toBe('Error listing categories: API Error');
  });

  test('should list categories from shared mailbox', async () => {
    const sharedMailbox = 'shared@company.com';
    const mockCategories = {
      value: [
        { displayName: 'Shared Category', color: 'preset0' }
      ]
    };

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(mockCategories);

    const result = await handleListCategories({ mailbox: sharedMailbox });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      `users/${sharedMailbox}/outlook/masterCategories`,
      { $top: 50 }
    );
    expect(result.content[0].text).toContain('shared mailbox');
    expect(result.content[0].text).toContain('Shared Category');
  });
});
