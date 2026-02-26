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

      // Multi-word subjects get inner KQL quotes for exact phrase search,
      // escaped within the outer double-quote wrapper
      expect(params.$search).toBe('"subject:\\"Quarterly Report\\""');
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

    test('from + after: date filtering is delegated to server-side KQL, not client-side', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      const emailsFromApi = [
        {
          id: 'old-email',
          subject: 'Old GoDaddy invoice',
          from: { emailAddress: { name: 'GoDaddy', address: 'billing@godaddy.com' } },
          receivedDateTime: '2025-11-15T10:00:00Z',
          isRead: true,
          categories: []
        },
        {
          id: 'new-email',
          subject: 'New GoDaddy invoice',
          from: { emailAddress: { name: 'GoDaddy', address: 'billing@godaddy.com' } },
          receivedDateTime: '2026-01-10T10:00:00Z',
          isRead: true,
          categories: []
        }
      ];
      callGraphAPIPaginated.mockResolvedValue({ value: emailsFromApi });

      const result = await handleSearchEmails({ from: 'godaddy', after: '2025-12-01' });

      // Date filtering is now server-side via KQL — mock API returns both emails,
      // client-side filtering by date has been removed.
      expect(result.content[0].text).toContain('new-email');
      expect(result.content[0].text).toContain('old-email');
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

  describe('Improvement 1+2: client-side sort after $search', () => {
    const makeEmail = (id, date) => ({
      id,
      subject: `Email ${id}`,
      from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
      receivedDateTime: date,
      isRead: true,
      categories: []
    });

    test('$search results should be sorted by date descending by default', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({
        value: [
          makeEmail('oldest', '2024-01-01T10:00:00Z'),
          makeEmail('newest', '2024-03-01T10:00:00Z'),
          makeEmail('middle', '2024-02-01T10:00:00Z'),
        ]
      });

      const result = await handleSearchEmails({ from: 'sender@example.com' });
      const text = result.content[0].text;

      expect(text.indexOf('ID: newest')).toBeLessThan(text.indexOf('ID: middle'));
      expect(text.indexOf('ID: middle')).toBeLessThan(text.indexOf('ID: oldest'));
    });

    test('$search results should be sorted by date ascending when sortOrder is asc', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({
        value: [
          makeEmail('oldest', '2024-01-01T10:00:00Z'),
          makeEmail('newest', '2024-03-01T10:00:00Z'),
          makeEmail('middle', '2024-02-01T10:00:00Z'),
        ]
      });

      const result = await handleSearchEmails({ from: 'sender@example.com', sortOrder: 'asc' });
      const text = result.content[0].text;

      expect(text.indexOf('ID: oldest')).toBeLessThan(text.indexOf('ID: middle'));
      expect(text.indexOf('ID: middle')).toBeLessThan(text.indexOf('ID: newest'));
    });

    test('with after-filter active, should request only requestedCount results (KQL handles server-side filtering)', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', after: '2025-12-01', count: 10 });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$top).toBe(10);
    });
  });

  describe('Bug 2: KQL values with spaces need inner quotes', () => {
    test('multi-word subject: $search should contain escaped inner quotes', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ subject: 'budget meeting' });

      const params = callGraphAPIPaginated.mock.calls[0][3];
      expect(params.$search).toContain('subject:\\"budget meeting\\"');
    });

    test('multi-word from: $search should contain escaped inner quotes', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'John Doe' });

      const params = callGraphAPIPaginated.mock.calls[0][3];
      expect(params.$search).toContain('from:\\"John Doe\\"');
    });

    test('single-word value: $search should NOT contain escaped inner quotes', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy' });

      const params = callGraphAPIPaginated.mock.calls[0][3];
      expect(params.$search).toBe('"from:godaddy"');
    });
  });

  describe('Bug 3: toKqlDate null-fallback produces invalid KQL', () => {
    test('invalid after date: received: term should be omitted from $search', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', after: 'not-a-date' });

      const params = callGraphAPIPaginated.mock.calls[0][3];
      expect(params.$search).not.toContain('received:');
    });

    test('invalid after + valid before: $search should produce received:..MM/DD/YYYY', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', after: 'not-a-date', before: '2025-12-31' });

      const params = callGraphAPIPaginated.mock.calls[0][3];
      expect(params.$search).toContain('received:..12/31/2025');
    });
  });

  describe('Improvement 3: KQL date ranges in $search', () => {
    test('from + after: $search should contain received:MM/DD/YYYY..', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', after: '2025-12-01' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toContain('received:12/01/2025..');
    });

    test('from + before: $search should contain received:..MM/DD/YYYY', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', before: '2025-12-31' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toContain('received:..12/31/2025');
    });

    test('from + after + before: $search should contain received:MM/DD/YYYY..MM/DD/YYYY', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', after: '2025-12-01', before: '2025-12-31' });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$search).toContain('received:12/01/2025..12/31/2025');
    });

    test('from + after: $top should be requestedCount, not 250', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', after: '2025-12-01', count: 10 });

      const firstCall = callGraphAPIPaginated.mock.calls[0];
      const params = firstCall[3];

      expect(params.$top).toBe(10);
    });
  });

  describe('Bug 1: $skip is incompatible with $search', () => {
    const makeEmail = (id, date) => ({
      id,
      subject: `Email ${id}`,
      from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
      receivedDateTime: date,
      isRead: true,
      categories: []
    });

    test('text search with skip: $skip should NOT be in API request params', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ from: 'godaddy', skip: 20 });

      const params = callGraphAPIPaginated.mock.calls[0][3];
      expect(params.$search).toBeDefined();
      expect(params.$skip).toBeUndefined();
    });

    test('text search with skip + count: $top should be skip + count', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({ value: [] });

      await handleSearchEmails({ from: 'godaddy', skip: 20, count: 5 });

      const params = callGraphAPIPaginated.mock.calls[0][3];
      expect(params.$top).toBe(25); // 20 + 5
    });

    test('text search with skip: results should be sliced correctly after client-side sort', async () => {
      resolveFolderPath.mockResolvedValue('me/messages');
      callGraphAPIPaginated.mockResolvedValue({
        value: [
          makeEmail('oldest', '2024-01-01T10:00:00Z'),
          makeEmail('newest', '2024-04-01T10:00:00Z'),
          makeEmail('second', '2024-03-01T10:00:00Z'),
          makeEmail('third',  '2024-02-01T10:00:00Z'),
        ]
      });

      // sort desc: [newest(Apr), second(Mar), third(Feb), oldest(Jan)]
      // skip=2, count=1 → slice(2, 3) → [third]
      const result = await handleSearchEmails({ from: 'sender@example.com', skip: 2, count: 1 });

      expect(result.content[0].text).toContain('ID: third');
      expect(result.content[0].text).not.toContain('ID: newest');
      expect(result.content[0].text).not.toContain('ID: second');
      expect(result.content[0].text).not.toContain('ID: oldest');
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
