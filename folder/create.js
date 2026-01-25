/**
 * Create folder functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getFolderIdByName } = require('../email/folder-utils');
const { getMailboxBasePath } = require('../utils/mailbox-path');

/**
 * Create folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolder(args) {
  const folderName = args.name;
  const parentFolder = args.parentFolder || '';
  const mailbox = args.mailbox || null;

  if (!folderName) {
    return {
      content: [{
        type: "text",
        text: "Folder name is required."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Create folder with appropriate parent
    const result = await createMailFolder(accessToken, folderName, parentFolder, mailbox);

    return {
      content: [{
        type: "text",
        text: result.message
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error creating folder: ${error.message}`
      }]
    };
  }
}

/**
 * Create a new mail folder
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to create
 * @param {string} parentFolderName - Name of the parent folder (optional)
 * @param {string|null} mailbox - Email address of shared mailbox, or null for primary
 * @returns {Promise<object>} - Result object with status and message
 */
async function createMailFolder(accessToken, folderName, parentFolderName, mailbox = null) {
  const basePath = getMailboxBasePath(mailbox);
  const mailboxInfo = mailbox ? ` in shared mailbox ${mailbox}` : '';

  try {
    // Check if a folder with this name already exists
    const existingFolder = await getFolderIdByName(accessToken, folderName, mailbox);
    if (existingFolder) {
      return {
        success: false,
        message: `A folder named "${folderName}" already exists${mailboxInfo}.`
      };
    }

    // If parent folder specified, find its ID
    let endpoint = `${basePath}/mailFolders`;
    if (parentFolderName) {
      const parentId = await getFolderIdByName(accessToken, parentFolderName, mailbox);
      if (!parentId) {
        return {
          success: false,
          message: `Parent folder "${parentFolderName}" not found${mailboxInfo}. Please specify a valid parent folder or leave it blank to create at the root level.`
        };
      }

      endpoint = `${basePath}/mailFolders/${parentId}/childFolders`;
    }

    // Create the folder
    const folderData = {
      displayName: folderName
    };

    const response = await callGraphAPI(
      accessToken,
      'POST',
      endpoint,
      folderData
    );

    if (response && response.id) {
      const locationInfo = parentFolderName
        ? `inside "${parentFolderName}"`
        : "at the root level";

      return {
        success: true,
        message: `Successfully created folder "${folderName}" ${locationInfo}${mailboxInfo}.`,
        folderId: response.id
      };
    } else {
      return {
        success: false,
        message: "Failed to create folder. The server didn't return a folder ID."
      };
    }
  } catch (error) {
    console.error(`Error creating folder "${folderName}": ${error.message}`);
    throw error;
  }
}

module.exports = handleCreateFolder;
