const handleAddCategory = require('../../categories/add');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleAddCategory', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should add category to email successfully', async () => {
    const emailId = 'email-123';
    const category = 'Red Category';

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    // First call: GET to get current categories
    callGraphAPI.mockResolvedValueOnce({ categories: ['Blue Category'] });
    // Second call: PATCH to update categories
    callGraphAPI.mockResolvedValueOnce({});

    const result = await handleAddCategory({ id: emailId, category });

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      mockAccessToken,
      'GET',
      `me/messages/${emailId}`,
      { $select: 'categories' }
    );
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      2,
      mockAccessToken,
      'PATCH',
      `me/messages/${emailId}`,
      null,
      { categories: ['Blue Category', 'Red Category'] }
    );
    expect(result.content[0].text).toContain('Category "Red Category" added');
  });

  test('should add category to email in shared mailbox', async () => {
    const emailId = 'email-123';
    const category = 'Red Category';
    const mailbox = 'shared@company.com';

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValueOnce({ categories: [] });
    callGraphAPI.mockResolvedValueOnce({});

    const result = await handleAddCategory({ id: emailId, category, mailbox });

    expect(callGraphAPI).toHaveBeenNthCalledWith(
      1,
      mockAccessToken,
      'GET',
      `users/${mailbox}/messages/${emailId}`,
      { $select: 'categories' }
    );
    expect(result.content[0].text).toContain('shared mailbox');
  });

  test('should not add duplicate category', async () => {
    const emailId = 'email-123';
    const category = 'Red Category';

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValueOnce({ categories: ['Red Category'] });

    const result = await handleAddCategory({ id: emailId, category });

    expect(callGraphAPI).toHaveBeenCalledTimes(1); // Only GET, no PATCH
    expect(result.content[0].text).toContain('already has the category');
  });

  test('should return error when email ID is missing', async () => {
    const result = await handleAddCategory({ category: 'Red Category' });

    expect(result.content[0].text).toBe(
      'Email ID is required. Please provide the ID of the email.'
    );
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });

  test('should return error when category is missing', async () => {
    const result = await handleAddCategory({ id: 'email-123' });

    expect(result.content[0].text).toContain('Category name is required');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleAddCategory({ id: 'email-123', category: 'Red' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });
});
