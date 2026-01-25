/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `📧 MODULAR Outlook Assistant MCP Server v${config.SERVER_VERSION} 📧

Provides access to Microsoft Outlook email, calendar, and contacts through Microsoft Graph API.
Implemented with a modular architecture for improved maintainability.

Features:
- Email: Read, search, send, and manage emails
- Calendar: View, create, and manage calendar events
- Folders: List and manage mail folders
- Rules: Manage inbox rules
- Shared Mailboxes: Access shared mailboxes and Microsoft 365 Group mailboxes

To access a shared mailbox, use the 'mailbox' parameter with the shared mailbox email address.
Example: list-emails with mailbox="shared@company.com"

Use 'list-shared-mailboxes' to discover available shared mailboxes.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;
  
  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    tokenManager.createTestTokens();
    
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }
  
  // For real authentication, generate an auth URL and instruct the user to visit it
  const authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;
  
  return {
    content: [{
      type: "text",
      text: `Authentication required. Please visit the following URL to authenticate with Microsoft: ${authUrl}\n\nAfter authentication, you will be redirected back to this application.`
    }]
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');

  const tokens = tokenManager.loadTokenCache();

  console.error(`[CHECK-AUTH-STATUS] Tokens loaded: ${tokens ? 'YES' : 'NO'}`);

  if (!tokens || !tokens.access_token) {
    console.error('[CHECK-AUTH-STATUS] No valid access token found');
    return {
      content: [{ type: "text", text: "Not authenticated. Please use the 'authenticate' tool to sign in." }]
    };
  }

  console.error('[CHECK-AUTH-STATUS] Access token present');
  console.error(`[CHECK-AUTH-STATUS] Token expires at: ${tokens.expires_at}`);
  console.error(`[CHECK-AUTH-STATUS] Current time: ${Date.now()}`);

  // Check if token is expired
  const isExpired = tokens.expires_at && Date.now() > tokens.expires_at;

  // Parse scopes from token response
  const grantedScopes = tokens.scope ? tokens.scope.split(' ') : [];
  const hasSharedMailboxAccess = grantedScopes.some(s =>
    s.toLowerCase().includes('.shared')
  );

  let statusMessage = isExpired
    ? "Token expired. Please re-authenticate using the 'authenticate' tool."
    : "Authenticated and ready";

  // Add scope information
  if (!isExpired && grantedScopes.length > 0) {
    statusMessage += "\n\nGranted permissions:";
    statusMessage += `\n- Primary mailbox: Yes`;
    statusMessage += `\n- Shared mailboxes: ${hasSharedMailboxAccess ? 'Yes' : 'No (re-authenticate to enable)'}`;

    if (!hasSharedMailboxAccess) {
      statusMessage += "\n\nNote: To access shared mailboxes, please re-authenticate with 'authenticate' tool (force=true).";
      statusMessage += "\nThe new authentication will request shared mailbox permissions.";
    }
  }

  return {
    content: [{ type: "text", text: statusMessage }]
  };
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this Outlook Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
