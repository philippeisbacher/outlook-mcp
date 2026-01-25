const handleRemoveCategory = require('../../categories/remove');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleRemoveCategory', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should remove category from email successfully', async () => {
    const emailId = 'email-123';
    const category = 'Red Category';

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    // First call: GET to get current categories
    callGraphAPI.mockResolvedValueOnce({ categories: ['Red Category', 'Blue Category'] });
    // Second call: PATCH to update categories
    callGraphAPI.mockResolvedValueOnce({});

    const result = await handleRemoveCategory({ id: emailId, category });

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
    expect(callGraphAPI).toHaveBeenNthCalledWith(
      2,
      mockAccessToken,
      'PATCH',
      `me/messages/${emailId}`,
      null,
      { categories: ['Blue Category'] }
    );
    expect(result.content[0].text).toContain('Category "Red Category" removed');
    expect(result.content[0].text).toContain('Blue Category');
  });

  test('should handle removing last category', async () => {
    const emailId = 'email-123';
    const category = 'Red Category';

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValueOnce({ categories: ['Red Category'] });
    callGraphAPI.mockResolvedValueOnce({});

    const result = await handleRemoveCategory({ id: emailId, category });

    expect(callGraphAPI).toHaveBeenNthCalledWith(
      2,
      mockAccessToken,
      'PATCH',
      `me/messages/${emailId}`,
      null,
      { categories: [] }
    );
    expect(result.content[0].text).toContain('no more categories');
  });

  test('should return error when category not on email', async () => {
    const emailId = 'email-123';
    const category = 'Red Category';

    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValueOnce({ categories: ['Blue Category'] });

    const result = await handleRemoveCategory({ id: emailId, category });

    expect(callGraphAPI).toHaveBeenCalledTimes(1); // Only GET, no PATCH
    expect(result.content[0].text).toContain('does not have the category');
  });

  test('should return error when email ID is missing', async () => {
    const result = await handleRemoveCategory({ category: 'Red Category' });

    expect(result.content[0].text).toBe(
      'Email ID is required. Please provide the ID of the email.'
    );
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });

  test('should return error when category is missing', async () => {
    const result = await handleRemoveCategory({ id: 'email-123' });

    expect(result.content[0].text).toBe('Category name is required.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleRemoveCategory({ id: 'email-123', category: 'Red' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });
});
