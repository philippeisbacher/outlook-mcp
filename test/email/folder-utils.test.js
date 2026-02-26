const {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName
} = require('../../email/folder-utils');
const { callGraphAPI } = require('../../utils/graph-api');

jest.mock('../../utils/graph-api');

describe('resolveFolderPath', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    // Mock console.error to avoid cluttering test output
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('well-known folders', () => {
    test('Bug 4: should return me/messages (global) when no folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, null);
      expect(result).toBe('me/messages');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('Bug 4: should return me/messages (global) when undefined folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, undefined);
      expect(result).toBe('me/messages');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('Bug 4: should return me/messages (global) when empty string is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, '');
      expect(result).toBe('me/messages');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return correct endpoint for well-known folders', async () => {
      const result = await resolveFolderPath(mockAccessToken, 'drafts');
      expect(result).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle case-insensitive well-known folder names', async () => {
      const result1 = await resolveFolderPath(mockAccessToken, 'INBOX');
      const result2 = await resolveFolderPath(mockAccessToken, 'Drafts');
      const result3 = await resolveFolderPath(mockAccessToken, 'SENT');

      expect(result1).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(result2).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(result3).toBe(WELL_KNOWN_FOLDERS['sent']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('custom folders', () => {
    test('should resolve custom folder by ID when found', async () => {
      const customFolderId = 'custom-folder-id-123';
      const customFolderName = 'MyCustomFolder';

      callGraphAPI.mockResolvedValueOnce({
        value: [{ id: customFolderId, displayName: customFolderName }]
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        { $filter: `displayName eq '${customFolderName}'` }
      );
    });

    test('should try case-insensitive search when exact match fails', async () => {
      const customFolderId = 'custom-folder-id-456';
      const customFolderName = 'ProjectAlpha';

      // First call returns empty (exact match fails)
      callGraphAPI.mockResolvedValueOnce({ value: [] });

      // Second call returns all folders for case-insensitive match
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'other-id', displayName: 'OtherFolder' },
          { id: customFolderId, displayName: 'projectalpha' }
        ]
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('should fall back to inbox when custom folder is not found', async () => {
      const nonExistentFolder = 'NonExistentFolder';

      // First call returns empty (exact match fails)
      callGraphAPI.mockResolvedValueOnce({ value: [] });

      // Second call returns folders without a match
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'id1', displayName: 'Folder1' },
          { id: 'id2', displayName: 'Folder2' }
        ]
      });

      const result = await resolveFolderPath(mockAccessToken, nonExistentFolder);

      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('should fall back to inbox when API call fails', async () => {
      const customFolderName = 'CustomFolder';

      callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
    });
  });
});

describe('getFolderIdByName', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return folder ID when exact match is found', async () => {
    const folderId = 'folder-id-123';
    const folderName = 'TestFolder';

    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: folderId, displayName: folderName }]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${folderName}'` }
    );
  });

  test('should return folder ID when case-insensitive match is found', async () => {
    const folderId = 'folder-id-456';
    const folderName = 'TestFolder';

    // First call returns empty (exact match fails)
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call returns folders with case-insensitive match
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: folderId, displayName: 'testfolder' }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });

  test('should return null when folder is not found', async () => {
    const folderName = 'NonExistentFolder';

    // First call returns empty
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call returns top-level folders without a match (but with children)
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: 'id1', displayName: 'OtherFolder', childFolderCount: 0 }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
  });

  test('Bug 3: should find child folders (e.g. 01_OPEN under Inbox)', async () => {
    const childFolderId = 'child-folder-id-789';
    const childFolderName = '01_OPEN';

    // First call: exact match filter returns empty
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call: top-level folders (no match, but Inbox has children)
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: 'inbox-id', displayName: 'Inbox', childFolderCount: 3 },
        { id: 'sent-id', displayName: 'Sent Items', childFolderCount: 0 }
      ]
    });

    // Third call: child folders of Inbox
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: 'child-1', displayName: '70_internal_c4l' },
        { id: childFolderId, displayName: '01_OPEN' },
        { id: 'child-3', displayName: 'Archiv' }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, childFolderName);

    expect(result).toBe(childFolderId);
  });

  test('Bug 3: should find child folders case-insensitively', async () => {
    const childFolderId = 'child-folder-id-abc';

    // First call: exact match filter returns empty
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call: top-level folders
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: 'inbox-id', displayName: 'Inbox', childFolderCount: 2 }
      ]
    });

    // Third call: child folders of Inbox
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: childFolderId, displayName: 'archiv' }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, 'Archiv');

    expect(result).toBe(childFolderId);
  });

  test('should return null when API call fails', async () => {
    const folderName = 'TestFolder';

    callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  test('Security: folder name with apostrophe should be escaped in $filter', async () => {
    callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'folder-id', displayName: "John's Inbox" }] });

    await getFolderIdByName(mockAccessToken, "John's Inbox");

    const queryParams = callGraphAPI.mock.calls[0][4];
    // OData escaping: single quote → two single quotes
    expect(queryParams.$filter).toBe("displayName eq 'John''s Inbox'");
  });
});

describe('shared mailbox support', () => {
  const mockAccessToken = 'dummy_access_token';
  const sharedMailbox = 'shared@company.com';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('resolveFolderPath should use shared mailbox path for well-known folders', async () => {
    const result = await resolveFolderPath(mockAccessToken, 'inbox', sharedMailbox);
    expect(result).toBe(`users/${sharedMailbox}/mailFolders/inbox/messages`);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('getFolderIdByName should use shared mailbox path', async () => {
    const folderId = 'folder-id-789';
    const folderName = 'CustomFolder';

    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: folderId, displayName: folderName }]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName, sharedMailbox);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      `users/${sharedMailbox}/mailFolders`,
      null,
      { $filter: `displayName eq '${folderName}'` }
    );
  });

  test('resolveFolderPath should resolve custom folder with shared mailbox', async () => {
    const customFolderId = 'custom-folder-shared';
    const customFolderName = 'SharedCustomFolder';

    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: customFolderId, displayName: customFolderName }]
    });

    const result = await resolveFolderPath(mockAccessToken, customFolderName, sharedMailbox);

    expect(result).toBe(`users/${sharedMailbox}/mailFolders/${customFolderId}/messages`);
  });
});
