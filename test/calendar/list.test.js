const handleListEvents = require('../../calendar/list');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleListEvents', () => {
  const mockAccessToken = 'dummy_access_token';

  const makeEvent = (overrides = {}) => ({
    id: 'event-1',
    subject: 'Team Meeting',
    bodyPreview: 'Discuss Q1 results',
    start: { dateTime: '2024-03-10T10:00:00', timeZone: 'UTC' },
    end:   { dateTime: '2024-03-10T11:00:00', timeZone: 'UTC' },
    location: { displayName: 'Conference Room A' },
    ...overrides
  });

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should list events with location', async () => {
    callGraphAPI.mockResolvedValue({ value: [makeEvent()] });

    const result = await handleListEvents({ count: 5 });

    expect(result.content[0].text).toContain('Team Meeting');
    expect(result.content[0].text).toContain('Conference Room A');
  });

  test('Bug: should not crash when event has no location (location is null)', async () => {
    callGraphAPI.mockResolvedValue({
      value: [makeEvent({ location: null })]
    });

    // Should NOT throw TypeError: Cannot read properties of null
    const result = await handleListEvents({ count: 5 });

    expect(result.content[0].text).toContain('Team Meeting');
    expect(result.content[0].text).toContain('No location');
  });

  test('Bug: should not crash when event location has no displayName', async () => {
    callGraphAPI.mockResolvedValue({
      value: [makeEvent({ location: {} })]
    });

    const result = await handleListEvents({ count: 5 });

    expect(result.content[0].text).toContain('No location');
  });

  test('should return "no events" message when API returns empty list', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListEvents({});

    expect(result.content[0].text).toContain('No calendar events found');
  });
});
