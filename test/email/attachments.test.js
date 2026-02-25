const { handleListAttachments, handleGetAttachment } = require('../../email/attachments');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleListAttachments', () => {
  const mockAccessToken = 'dummy_access_token';
  const mockAttachments = [
    {
      id: 'att-1',
      name: 'screenshot.png',
      contentType: 'image/png',
      size: 102400
    },
    {
      id: 'att-2',
      name: 'report.pdf',
      contentType: 'application/pdf',
      size: 2048000
    }
  ];

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should list attachments for a given email', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: mockAttachments });

    const result = await handleListAttachments({ id: 'email-1' });

    expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/email-1/attachments',
      null,
      { $select: 'id,name,contentType,size' }
    );
    expect(result.content[0].text).toContain('2 attachment(s)');
    expect(result.content[0].text).toContain('screenshot.png');
    expect(result.content[0].text).toContain('report.pdf');
  });

  test('should return message when no attachments found', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListAttachments({ id: 'email-1' });

    expect(result.content[0].text).toBe('No attachments found.');
  });

  test('should require email ID', async () => {
    const result = await handleListAttachments({});

    expect(result.content[0].text).toBe('Email ID is required.');
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });

  test('should support shared mailbox', async () => {
    const sharedMailbox = 'shared@company.com';
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: mockAttachments });

    await handleListAttachments({ id: 'email-1', mailbox: sharedMailbox });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      `users/${sharedMailbox}/messages/email-1/attachments`,
      null,
      { $select: 'id,name,contentType,size' }
    );
  });

  test('should include attachment IDs in output', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: mockAttachments });

    const result = await handleListAttachments({ id: 'email-1' });

    expect(result.content[0].text).toContain('ID: att-1');
    expect(result.content[0].text).toContain('ID: att-2');
  });

  test('should show size in KB', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: mockAttachments });

    const result = await handleListAttachments({ id: 'email-1' });

    expect(result.content[0].text).toContain('100.0 KB');
    expect(result.content[0].text).toContain('2000.0 KB');
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleListAttachments({ id: 'email-1' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });

  test('should handle Graph API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

    const result = await handleListAttachments({ id: 'email-1' });

    expect(result.content[0].text).toBe('Error listing attachments: Graph API Error');
  });
});

describe('handleGetAttachment', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return image attachment as image content block', async () => {
    const mockAttachment = {
      name: 'screenshot.png',
      contentType: 'image/png',
      size: 102400,
      contentBytes: 'iVBORw0KGgoAAAANSUhEUg=='
    };
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(mockAttachment);

    const result = await handleGetAttachment({ emailId: 'email-1', attachmentId: 'att-1' });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/email-1/attachments/att-1',
      null,
      {}
    );
    expect(result.content).toHaveLength(2);
    expect(result.content[0].type).toBe('text');
    expect(result.content[0].text).toContain('screenshot.png');
    expect(result.content[1].type).toBe('image');
    expect(result.content[1].data).toBe('iVBORw0KGgoAAAANSUhEUg==');
    expect(result.content[1].mimeType).toBe('image/png');
  });

  test('should return text attachment as decoded text', async () => {
    const textContent = 'Hello, this is a text file.';
    const mockAttachment = {
      name: 'notes.txt',
      contentType: 'text/plain',
      size: 27,
      contentBytes: Buffer.from(textContent).toString('base64')
    };
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(mockAttachment);

    const result = await handleGetAttachment({ emailId: 'email-1', attachmentId: 'att-2' });

    expect(result.content).toHaveLength(1);
    expect(result.content[0].text).toContain('notes.txt');
    expect(result.content[0].text).toContain(textContent);
  });

  test('should return binary attachment as base64 with metadata', async () => {
    const mockAttachment = {
      name: 'report.pdf',
      contentType: 'application/pdf',
      size: 2048000,
      contentBytes: 'JVBERi0xLjQK'
    };
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(mockAttachment);

    const result = await handleGetAttachment({ emailId: 'email-1', attachmentId: 'att-3' });

    expect(result.content[0].text).toContain('report.pdf');
    expect(result.content[0].text).toContain('application/pdf');
    expect(result.content[0].text).toContain('2000.0 KB');
    expect(result.content[0].text).toContain('JVBERi0xLjQK');
  });

  test('should require emailId', async () => {
    const result = await handleGetAttachment({ attachmentId: 'att-1' });

    expect(result.content[0].text).toBe('Email ID and Attachment ID are required.');
  });

  test('should require attachmentId', async () => {
    const result = await handleGetAttachment({ emailId: 'email-1' });

    expect(result.content[0].text).toBe('Email ID and Attachment ID are required.');
  });

  test('should support shared mailbox', async () => {
    const sharedMailbox = 'shared@company.com';
    const mockAttachment = {
      name: 'file.pdf',
      contentType: 'application/pdf',
      size: 1024,
      contentBytes: 'abc123'
    };
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(mockAttachment);

    await handleGetAttachment({ emailId: 'email-1', attachmentId: 'att-1', mailbox: sharedMailbox });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      `users/${sharedMailbox}/messages/email-1/attachments/att-1`,
      null,
      {}
    );
  });

  test('should handle authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleGetAttachment({ emailId: 'email-1', attachmentId: 'att-1' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
  });

  test('should handle Graph API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Not Found'));

    const result = await handleGetAttachment({ emailId: 'email-1', attachmentId: 'att-1' });

    expect(result.content[0].text).toBe('Error getting attachment: Not Found');
  });
});
