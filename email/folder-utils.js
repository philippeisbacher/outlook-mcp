/**
 * Email folder utilities
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getMailboxBasePath, getWellKnownFolders, buildMailboxPath } = require('../utils/mailbox-path');

/**
 * Cache of folder information to reduce API calls
 * Format: { mailbox: { folderName: { id, path } } }
 */
const folderCache = {};

/**
 * Well-known folder names and their endpoints (for primary mailbox - legacy support)
 * @deprecated Use getWellKnownFolders(mailbox) instead for shared mailbox support
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

/**
 * Resolve a folder name to its endpoint path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name to resolve
 * @param {string|null} mailbox - Email address of shared mailbox, or null for primary
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName, mailbox = null) {
  const wellKnownFolders = getWellKnownFolders(mailbox);
  const basePath = getMailboxBasePath(mailbox);

  // Default to inbox if no folder specified
  if (!folderName) {
    return wellKnownFolders['inbox'];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (wellKnownFolders[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}" in mailbox "${mailbox || 'primary'}"`);
    return wellKnownFolders[lowerFolderName];
  }

  try {
    // Try to find the folder by name
    const folderId = await getFolderIdByName(accessToken, folderName, mailbox);
    if (folderId) {
      const path = `${basePath}/mailFolders/${folderId}/messages`;
      console.error(`Resolved folder "${folderName}" to path: ${path}`);
      return path;
    }

    // If not found, fall back to inbox
    console.error(`Couldn't find folder "${folderName}", falling back to inbox`);
    return wellKnownFolders['inbox'];
  } catch (error) {
    console.error(`Error resolving folder "${folderName}": ${error.message}`);
    return wellKnownFolders['inbox'];
  }
}

/**
 * Get the ID of a mail folder by its name
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find
 * @param {string|null} mailbox - Email address of shared mailbox, or null for primary
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName, mailbox = null) {
  const basePath = getMailboxBasePath(mailbox);

  try {
    // First try with exact match filter
    console.error(`Looking for folder with name "${folderName}" in mailbox "${mailbox || 'primary'}"`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `${basePath}/mailFolders`,
      null,
      { $filter: `displayName eq '${folderName}'` }
    );

    if (response.value && response.value.length > 0) {
      console.error(`Found folder "${folderName}" with ID: ${response.value[0].id}`);
      return response.value[0].id;
    }

    // If exact match fails, try to get all folders and do a case-insensitive comparison
    console.error(`No exact match found for "${folderName}", trying case-insensitive search`);
    const allFoldersResponse = await callGraphAPI(
      accessToken,
      'GET',
      `${basePath}/mailFolders`,
      null,
      { $top: 100 }
    );

    if (allFoldersResponse.value) {
      const lowerFolderName = folderName.toLowerCase();
      const matchingFolder = allFoldersResponse.value.find(
        folder => folder.displayName.toLowerCase() === lowerFolderName
      );

      if (matchingFolder) {
        console.error(`Found case-insensitive match for "${folderName}" with ID: ${matchingFolder.id}`);
        return matchingFolder.id;
      }
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Get all mail folders
 * @param {string} accessToken - Access token
 * @param {string|null} mailbox - Email address of shared mailbox, or null for primary
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken, mailbox = null) {
  const basePath = getMailboxBasePath(mailbox);

  try {
    // Get top-level folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `${basePath}/mailFolders`,
      null,
      {
        $top: 100,
        $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      }
    );

    if (!response.value) {
      return [];
    }

    // Get child folders for folders with children
    const foldersWithChildren = response.value.filter(f => f.childFolderCount > 0);

    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `${basePath}/mailFolders/${folder.id}/childFolders`,
          null,
          {
            $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
          }
        );

        return childResponse.value || [];
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    });

    const childFolders = await Promise.all(childFolderPromises);

    // Combine top-level folders and all child folders
    return [...response.value, ...childFolders.flat()];
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  getAllFolders
};
