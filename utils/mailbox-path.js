/**
 * Utility functions for building mailbox-aware API paths
 * Supports both primary mailbox (me) and shared mailboxes (users/{email})
 */

/**
 * Get the base path for a mailbox
 * @param {string|null} mailbox - Email address of shared mailbox, or null/undefined for primary mailbox
 * @returns {string} - Base path ('me' or 'users/{email}')
 */
function getMailboxBasePath(mailbox) {
  if (!mailbox || mailbox.trim() === '') {
    return 'me';
  }
  // For shared mailboxes, use users/{email} format
  return `users/${mailbox.trim()}`;
}

/**
 * Build a mailbox-aware API path
 * @param {string|null} mailbox - Email address of shared mailbox, or null for primary
 * @param {string} resourcePath - The resource path (e.g., 'mailFolders/inbox/messages')
 * @returns {string} - Full API path
 */
function buildMailboxPath(mailbox, resourcePath) {
  const basePath = getMailboxBasePath(mailbox);
  // Remove leading slash if present
  const cleanResourcePath = resourcePath.startsWith('/') ? resourcePath.slice(1) : resourcePath;
  return `${basePath}/${cleanResourcePath}`;
}

/**
 * Get well-known folder paths for a mailbox
 * @param {string|null} mailbox - Email address of shared mailbox, or null for primary
 * @returns {object} - Object mapping folder names to their paths
 */
function getWellKnownFolders(mailbox) {
  const basePath = getMailboxBasePath(mailbox);
  return {
    'inbox': `${basePath}/mailFolders/inbox/messages`,
    'drafts': `${basePath}/mailFolders/drafts/messages`,
    'sent': `${basePath}/mailFolders/sentItems/messages`,
    'deleted': `${basePath}/mailFolders/deletedItems/messages`,
    'junk': `${basePath}/mailFolders/junkemail/messages`,
    'archive': `${basePath}/mailFolders/archive/messages`
  };
}

/**
 * Schema property for mailbox parameter (to be used in tool definitions)
 */
const mailboxSchemaProperty = {
  mailbox: {
    type: 'string',
    description: 'Email address of a shared mailbox to access. Leave empty to use your primary mailbox.'
  }
};

module.exports = {
  getMailboxBasePath,
  buildMailboxPath,
  getWellKnownFolders,
  mailboxSchemaProperty
};
