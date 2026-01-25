const handleDeleteEmail = require('../../email/delete');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleDeleteEmail', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful deletion', () => {
    test('should delete email from primary mailbox', async () => {
      const emailId = 'email-123';
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleDeleteEmail({ id: emailId });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'DELETE',
        `me/messages/${emailId}`,
        null
      );
      expect(result.content[0].text).toContain('Email deleted successfully');
      expect(result.content[0].text).toContain('Deleted Items');
    });

    test('should delete email from shared mailbox', async () => {
      const emailId = 'email-456';
      const sharedMailbox = 'shared@company.com';
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({});

      const result = await handleDeleteEmail({ id: emailId, mailbox: sharedMailbox });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'DELETE',
        `users/${sharedMailbox}/messages/${emailId}`,
        null
      );
      expect(result.content[0].text).toContain('shared mailbox: shared@company.com');
    });
  });

  describe('validation', () => {
    test('should return error when no email ID is provided', async () => {
      const result = await handleDeleteEmail({});

      expect(result.content[0].text).toBe(
        'Email ID is required. Please provide the ID of the email to delete.'
      );
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });

    test('should return error when email ID is empty string', async () => {
      const result = await handleDeleteEmail({ id: '' });

      expect(result.content[0].text).toBe(
        'Email ID is required. Please provide the ID of the email to delete.'
      );
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleDeleteEmail({ id: 'email-123' });

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle not found error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('404 Resource not found'));

      const result = await handleDeleteEmail({ id: 'email-123' });

      expect(result.content[0].text).toContain('Email not found');
      expect(result.content[0].text).toContain('email-123');
    });

    test('should handle generic API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('Network error'));

      const result = await handleDeleteEmail({ id: 'email-123' });

      expect(result.content[0].text).toBe('Error deleting email: Network error');
    });
  });
});
