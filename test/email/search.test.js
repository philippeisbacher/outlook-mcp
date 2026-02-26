const handleSearchEmails = require('../../email/search');
const { callGraphAPI, callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveFolderPath } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

describe('handleSearchEmails', () => {
  const mockAccessToken = 'dummy_access_token';
  const mockEmails = [
    {
      id: 'email-1',
      subject: 'Quarterly Report',
      from: { emailAddress: { name: 'John Doe', address: 'john@example.com' } },
      receivedDateTime: '2024-01-15T10:30:00Z',
      isRead: true,
      categories: []
    }
  ];

  beforeEach(() => {
    callGraphAPI.mockClear();
    callGraphAPIPaginated.mockClear();
    ensureAuthenticated.mockClear();
    resolveFolderPath.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('Bug 1+2: $search parameter must use correct KQL format', () => {
    test('from-filter should be passed as $search with outer quotes', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'john@example.com' });

      // The first call should be the combined search
      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3]; // queryParams

      expect(params.$search).toBe('"from:john@example.com"');
    });

    test('subject-filter should be passed as $search with outer quotes', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ subject: 'Quarterly Report' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toBe('"subject:Quarterly Report"');
    });

    test('combined from+subject should use correct KQL format', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'john@example.com', subject: 'Report' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toBe('"subject:Report from:john@example.com"');
    });

    test('$orderby must NOT be set when $search is used', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'john@example.com' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toBeDefined();
      expect(params.$orderby).toBeUndefined();
    });

    test('$orderby should be set when no $search terms are used', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ hasAttachments: true });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toBeUndefined();
      expect(params.$orderby).toBe('receivedDateTime desc');
    });

    test('general query should be passed as $search with outer quotes', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ query: 'budget meeting' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toBe('"budget meeting"');
    });
  });

  describe('Bug 3: $search and $filter must not be combined', () => {
    test('from + after must NOT send $filter to the API', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', after: '2025-12-01' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toBeDefined();
      expect(params.$filter).toBeUndefined();
    });

    test('from + after: results must be filtered client-side by date', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      const emailsFromApi = [
        {
          id: 'old-email',
          subject: 'Old GoDaddy invoice',
          from: { emailAddress: { name: 'GoDaddy', address: 'billing@godaddy.com' } },
          receivedDateTime: '2025-11-15T10:00:00Z', // before after-date → should be filtered out
          isRead: true,
          categories: []
        },
        {
          id: 'new-email',
          subject: 'New GoDaddy invoice',
          from: { emailAddress: { name: 'GoDaddy', address: 'billing@godaddy.com' } },
          receivedDateTime: '2026-01-10T10:00:00Z', // after after-date → should be kept
          isRead: true,
          categories: []
        }
      ];
      callGraphAPIPaginated.mockResolvedValue({ value: emailsFromApi });

      const result = await handleSearchEmails({ from: 'godaddy', after: '2025-12-01' });

      expect(result.content[0].text).toContain('new-email');
      expect(result.content[0].text).not.toContain('old-email');
    });

    test('when $search returns 0 results, should NOT fall back to basic listing', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: [] });

      const result = await handleSearchEmails({ from: 'nobody@example.com' });

      // Should only make ONE API call (no fallbacks)
      expect(callGraphAPIPaginated).toHaveBeenCalledTimes(1);
      expect(result.content[0].text).toContain('No emails found');
    });
  });

  describe('Bug 4: global search when no folder specified', () => {
    test('should use me/messages when no folder is specified', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ query: 'test' });

      // resolveFolderPath should be called with null folder
      expect(resolveFolderPath).toHaveBeenCalledWith(mockAccessToken, null, null);
    });

    test('should use specific folder path when folder is specified', async () => {
      resolveFolderPath.mockResolvedValue('me/mailFolders/inbox/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ folder: 'inbox', query: 'test' });

      expect(resolveFolderPath).toHaveBeenCalledWith(mockAccessToken, 'inbox', null);
    });
  });
});
