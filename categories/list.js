/**
 * List available Outlook categories (master list)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List categories handler - returns the master category list
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListCategories(args) {
  try {
    const accessToken = await ensureAuthenticated();

    console.error('[LIST-CATEGORIES] Fetching master category list');

    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/outlook/masterCategories',
      { $top: 50 }
    );

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No categories found. You can create categories in Outlook settings."
        }]
      };
    }

    const categoryList = response.value.map((cat, index) => {
      return `${index + 1}. ${cat.displayName} (Color: ${cat.color || 'none'})`;
    }).join('\n');

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} categories:\n\n${categoryList}`
      }]
    };
  } catch (error) {
    console.error(`[LIST-CATEGORIES] Error: ${error.message}`);

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
        text: `Error listing categories: ${error.message}`
      }]
    };
  }
}

module.exports = handleListCategories;
