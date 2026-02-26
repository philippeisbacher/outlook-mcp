const {
  handleListContacts,
  handleGetContact,
  handleCreateContact,
  handleUpdateContact,
  handleDeleteContact
} = require('../../contacts/crud');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const mockAccessToken = 'dummy_access_token';

const makeContact = (id, name, email) => ({
  id,
  displayName: name,
  emailAddresses: [{ address: email, name }],
  jobTitle: 'Engineer',
  mobilePhone: '+49123456'
});

beforeEach(() => {
  callGraphAPI.mockClear();
  ensureAuthenticated.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

// ── list-contacts ─────────────────────────────────────────────────────────────
describe('handleListContacts', () => {
  test('should GET /me/contacts', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [makeContact('c1', 'Alice', 'alice@example.com')] });

    await handleListContacts({});

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken, 'GET', 'me/contacts', null,
      expect.objectContaining({ $top: expect.any(Number) })
    );
  });

  test('should return formatted list', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [makeContact('c1', 'Alice', 'alice@example.com')] });

    const result = await handleListContacts({});

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Alice');
    expect(result.content[0].text).toContain('alice@example.com');
  });

  test('should return "no contacts" when list is empty', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListContacts({});

    expect(result.content[0].text).toContain('No contacts');
  });
});

// ── get-contact ───────────────────────────────────────────────────────────────
describe('handleGetContact', () => {
  test('should GET /me/contacts/{id}', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(makeContact('c1', 'Alice', 'alice@example.com'));

    await handleGetContact({ id: 'c1' });

    expect(callGraphAPI).toHaveBeenCalledWith(mockAccessToken, 'GET', 'me/contacts/c1', null, expect.any(Object));
  });

  test('should return contact details', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue(makeContact('c1', 'Alice', 'alice@example.com'));

    const result = await handleGetContact({ id: 'c1' });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Alice');
    expect(result.content[0].text).toContain('alice@example.com');
  });

  test('should return error when id is missing', async () => {
    const result = await handleGetContact({});

    expect(result.isError).toBe(true);
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });
});

// ── create-contact ────────────────────────────────────────────────────────────
describe('handleCreateContact', () => {
  test('should POST /me/contacts with displayName and email', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ id: 'new-c1', displayName: 'Bob' });

    await handleCreateContact({ displayName: 'Bob', email: 'bob@example.com' });

    const [, method, endpoint, body] = callGraphAPI.mock.calls[0];
    expect(method).toBe('POST');
    expect(endpoint).toBe('me/contacts');
    expect(body.displayName).toBe('Bob');
    expect(body.emailAddresses[0].address).toBe('bob@example.com');
  });

  test('should return success with contact id', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ id: 'new-c1', displayName: 'Bob' });

    const result = await handleCreateContact({ displayName: 'Bob', email: 'bob@example.com' });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('new-c1');
  });

  test('should return error when displayName is missing', async () => {
    const result = await handleCreateContact({ email: 'bob@example.com' });

    expect(result.isError).toBe(true);
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });
});

// ── update-contact ────────────────────────────────────────────────────────────
describe('handleUpdateContact', () => {
  test('should PATCH /me/contacts/{id} with only provided fields', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({});

    await handleUpdateContact({ id: 'c1', jobTitle: 'Manager' });

    const [, method, endpoint, body] = callGraphAPI.mock.calls[0];
    expect(method).toBe('PATCH');
    expect(endpoint).toBe('me/contacts/c1');
    expect(body.jobTitle).toBe('Manager');
    expect(body.displayName).toBeUndefined();
  });

  test('should return error when id is missing', async () => {
    const result = await handleUpdateContact({ jobTitle: 'Manager' });

    expect(result.isError).toBe(true);
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });

  test('should return error when no fields to update', async () => {
    const result = await handleUpdateContact({ id: 'c1' });

    expect(result.isError).toBe(true);
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });
});

// ── delete-contact ────────────────────────────────────────────────────────────
describe('handleDeleteContact', () => {
  test('should DELETE /me/contacts/{id}', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({});

    await handleDeleteContact({ id: 'c1' });

    expect(callGraphAPI).toHaveBeenCalledWith(mockAccessToken, 'DELETE', 'me/contacts/c1', null);
  });

  test('should return success message', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({});

    const result = await handleDeleteContact({ id: 'c1' });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('deleted');
  });

  test('should return error when id is missing', async () => {
    const result = await handleDeleteContact({});

    expect(result.isError).toBe(true);
    expect(ensureAuthenticated).not.toHaveBeenCalled();
  });
});
