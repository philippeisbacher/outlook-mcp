const { escapeODataString } = require('../../utils/odata-helpers');

jest.spyOn(console, 'error').mockImplementation(() => {});

describe('escapeODataString', () => {
  test('should escape single quotes (OData injection vector)', () => {
    expect(escapeODataString("O'Brien")).toBe("O''Brien");
  });

  test('should preserve @ in email addresses', () => {
    expect(escapeODataString('alice@example.com')).toBe('alice@example.com');
  });

  test('should preserve dots in email addresses', () => {
    expect(escapeODataString('first.last@example.com')).toBe('first.last@example.com');
  });

  test('should preserve hyphens in strings', () => {
    expect(escapeODataString('my-folder')).toBe('my-folder');
  });

  test('should handle null/undefined gracefully', () => {
    expect(escapeODataString(null)).toBe(null);
    expect(escapeODataString(undefined)).toBe(undefined);
  });

  test('should escape single quote in email-like string', () => {
    expect(escapeODataString("john's.boss@example.com")).toBe("john''s.boss@example.com");
  });
});
