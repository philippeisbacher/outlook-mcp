/**
 * Tests for callGraphAPI URL construction — specifically the encoding of OData params
 */
const https = require('https');

jest.mock('https');
jest.mock('../../config', () => ({
  USE_TEST_MODE: false,
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/'
}));

const { callGraphAPI } = require('../../utils/graph-api');

function mockHttpsRequest(statusCode = 200, body = '{"value":[]}') {
  const mockReq = {
    write: jest.fn(),
    end: jest.fn(),
    on: jest.fn()
  };

  https.request.mockImplementation((url, options, callback) => {
    const mockRes = {
      statusCode,
      on: jest.fn((event, cb) => {
        if (event === 'data') cb(body);
        if (event === 'end') cb();
      })
    };
    callback(mockRes);
    return mockReq;
  });

  return () => https.request.mock.calls[0][0]; // returns captured URL getter
}

beforeEach(() => {
  https.request.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  console.error.mockRestore();
});

describe('callGraphAPI — $search URL encoding', () => {
  test('$search value must NOT be percent-encoded — KQL requires literal : and "', async () => {
    const getUrl = mockHttpsRequest();

    await callGraphAPI('token', 'GET', 'me/messages', null, {
      $search: '"from:paddle"',
      $top: 10
    });

    const url = getUrl();
    // The URL must contain the literal KQL syntax
    expect(url).toContain('$search="from:paddle"');
    // Must NOT have percent-encoded quotes or colons
    expect(url).not.toContain('%22');
    expect(url).not.toContain('%3A');
  });

  test('from: KQL term must survive URL construction unchanged', async () => {
    const getUrl = mockHttpsRequest();

    await callGraphAPI('token', 'GET', 'me/messages', null, {
      $search: '"from:cursor subject:invoice"',
      $top: 10
    });

    const url = getUrl();
    expect(url).toContain('from:cursor');
    expect(url).toContain('subject:invoice');
  });

  test('received: date range must survive URL construction unchanged', async () => {
    const getUrl = mockHttpsRequest();

    await callGraphAPI('token', 'GET', 'me/messages', null, {
      $search: '"from:paddle received:12/01/2025.."',
      $top: 10
    });

    const url = getUrl();
    expect(url).toContain('received:12/01/2025..');
  });
});

describe('callGraphAPI — $filter URL encoding', () => {
  test('$filter value must be encodeURIComponent encoded', async () => {
    const getUrl = mockHttpsRequest();

    await callGraphAPI('token', 'GET', 'me/messages', null, {
      $filter: "hasAttachments eq true",
      $top: 10
    });

    const url = getUrl();
    expect(url).toContain('$filter=hasAttachments%20eq%20true');
  });

  test('$filter and $search can coexist in URL', async () => {
    const getUrl = mockHttpsRequest();

    await callGraphAPI('token', 'GET', 'me/messages', null, {
      $search: '"from:paddle"',
      $filter: "isRead eq false",
      $top: 10
    });

    const url = getUrl();
    expect(url).toContain('$search="from:paddle"');
    expect(url).toContain('$filter=isRead%20eq%20false');
  });
});

describe('callGraphAPI — regular params', () => {
  test('$top and $select are encoded via URLSearchParams ($ becomes %24, commas encoded)', async () => {
    const getUrl = mockHttpsRequest();

    await callGraphAPI('token', 'GET', 'me/messages', null, {
      $top: 25,
      $select: 'id,subject,from'
    });

    const url = getUrl();
    // URLSearchParams encodes $ → %24; Graph API accepts this transparently
    expect(url).toContain('%24top=25');
    expect(url).toContain('%24select=id%2Csubject%2Cfrom');
  });
});
