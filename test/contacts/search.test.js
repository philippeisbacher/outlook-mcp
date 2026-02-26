const handleSearchPeople = require('../../contacts/search');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const mockAccessToken = 'dummy_access_token';

const makePerson = (name, email, title = 'Engineer') => ({
  id: `person-${name}`,
  displayName: name,
  jobTitle: title,
  department: 'Engineering',
  emailAddresses: [{ address: email, name }]
});

beforeEach(() => {
  callGraphAPI.mockClear();
  ensureAuthenticated.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

describe('handleSearchPeople', () => {
  describe('successful search', () => {
    test('should call GET /me/people with $search parameter', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [makePerson('Alice Smith', 'alice@example.com')] });

      await handleSearchPeople({ query: 'Alice' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/people',
        null,
        expect.objectContaining({ $search: '"Alice"' })
      );
    });

    test('should return formatted list with name, email, title', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        value: [makePerson('Alice Smith', 'alice@example.com', 'Senior Engineer')]
      });

      const result = await handleSearchPeople({ query: 'Alice' });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].text).toContain('Alice Smith');
      expect(result.content[0].text).toContain('alice@example.com');
      expect(result.content[0].text).toContain('Senior Engineer');
    });

    test('should return "no results" when list is empty', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleSearchPeople({ query: 'nobody' });

      expect(result.content[0].text).toContain('No people found');
    });

    test('should respect count parameter', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleSearchPeople({ query: 'test', count: 5 });

      const [, , , , params] = callGraphAPI.mock.calls[0];
      expect(params.$top).toBe(5);
    });
  });

  describe('validation', () => {
    test('should return error when query is missing', async () => {
      const result = await handleSearchPeople({});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('query');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleSearchPeople({ query: 'test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain("authenticate");
    });

    test('should handle API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('500 Server Error'));

      const result = await handleSearchPeople({ query: 'test' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('500 Server Error');
    });
  });
});
