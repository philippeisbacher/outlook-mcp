/**
 * Shared mailboxes discovery functionality
 */
const { callGraphAPI } = require('./graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List shared mailboxes that the user has access to
 * Note: This uses the Graph API to find mailboxes the user has been granted access to.
 * The exact method depends on how shared mailboxes are configured in the organization.
 *
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListSharedMailboxes(args) {
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Try multiple methods to discover shared mailboxes
    const sharedMailboxes = [];
    const errors = [];

    // Method 1: Try to get shared mailboxes via mailboxSettings
    // This gets the user's own mailbox info first
    try {
      const userResponse = await callGraphAPI(accessToken, 'GET', 'me', null, {
        $select: 'displayName,mail,userPrincipalName'
      });

      if (userResponse && userResponse.mail) {
        // Add the user's primary mailbox info for reference
        sharedMailboxes.push({
          type: 'primary',
          displayName: userResponse.displayName || 'Primary Mailbox',
          email: userResponse.mail || userResponse.userPrincipalName,
          note: 'Your primary mailbox'
        });
      }
    } catch (error) {
      console.error('Error getting user info:', error.message);
    }

    // Method 2: Try to find shared mailboxes via people API (contacts with shared mailbox type)
    // This may work in some organizations
    try {
      const peopleResponse = await callGraphAPI(accessToken, 'GET', 'me/people', null, {
        $filter: "personType/subclass eq 'SharedMailbox'",
        $select: 'displayName,scoredEmailAddresses',
        $top: 50
      });

      if (peopleResponse && peopleResponse.value) {
        for (const person of peopleResponse.value) {
          if (person.scoredEmailAddresses && person.scoredEmailAddresses.length > 0) {
            sharedMailboxes.push({
              type: 'shared',
              displayName: person.displayName,
              email: person.scoredEmailAddresses[0].address,
              note: 'Shared mailbox'
            });
          }
        }
      }
    } catch (error) {
      // This is expected to fail in many configurations
      console.error('People API shared mailbox search failed:', error.message);
    }

    // Method 3: Try group mailboxes (Microsoft 365 Groups)
    try {
      const groupsResponse = await callGraphAPI(accessToken, 'GET', 'me/memberOf', null, {
        $filter: "groupTypes/any(c:c eq 'Unified')",
        $select: 'displayName,mail,id',
        $top: 50
      });

      if (groupsResponse && groupsResponse.value) {
        for (const group of groupsResponse.value) {
          if (group.mail) {
            sharedMailboxes.push({
              type: 'group',
              displayName: group.displayName,
              email: group.mail,
              note: 'Microsoft 365 Group mailbox'
            });
          }
        }
      }
    } catch (error) {
      console.error('Groups API failed:', error.message);
      errors.push('Could not retrieve Microsoft 365 Group mailboxes');
    }

    // Format the response
    if (sharedMailboxes.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No shared mailboxes found.\n\nNote: Shared mailbox discovery depends on your organization's configuration. If you know the email address of a shared mailbox you have access to, you can use it directly with the 'mailbox' parameter in other tools (e.g., list-emails, send-email, list-events).\n\nExample: Use mailbox="shared@company.com" to access that shared mailbox.`
        }]
      };
    }

    // Format the list
    const mailboxList = sharedMailboxes.map((mb, index) => {
      return `${index + 1}. ${mb.displayName}\n   Email: ${mb.email}\n   Type: ${mb.note}`;
    }).join('\n\n');

    let response = `Found ${sharedMailboxes.length} mailbox(es):\n\n${mailboxList}`;

    response += `\n\n---\nTo access a shared mailbox, use the 'mailbox' parameter with the email address.\nExample: list-emails with mailbox="${sharedMailboxes.find(m => m.type === 'shared')?.email || sharedMailboxes[0].email}"`;

    if (errors.length > 0) {
      response += `\n\nNote: Some discovery methods failed. If you know of additional shared mailboxes, you can try accessing them directly.`;
    }

    return {
      content: [{
        type: "text",
        text: response
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
        text: `Error discovering shared mailboxes: ${error.message}\n\nNote: If you know the email address of a shared mailbox you have access to, you can use it directly with the 'mailbox' parameter in other tools.`
      }]
    };
  }
}

// Tool definition for list-shared-mailboxes
const sharedMailboxTool = {
  name: "list-shared-mailboxes",
  description: "Discovers shared mailboxes and Microsoft 365 Group mailboxes that you have access to. Use the returned email addresses with the 'mailbox' parameter in other tools.",
  inputSchema: {
    type: "object",
    properties: {},
    required: []
  },
  handler: handleListSharedMailboxes
};

module.exports = {
  handleListSharedMailboxes,
  sharedMailboxTool
};
