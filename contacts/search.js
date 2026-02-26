/**
 * Search people (contacts + org chart) functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Search people handler — finds colleagues by name or email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchPeople(args) {
  const { query, count = 10 } = args;

  if (!query) {
    return {
      content: [{ type: "text", text: "query is required (name or email fragment to search for)." }],
      isError: true
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const response = await callGraphAPI(accessToken, 'GET', 'me/people', null, {
      $search: `"${query}"`,
      $top: Math.min(count, config.MAX_RESULT_COUNT),
      $select: 'id,displayName,emailAddresses,jobTitle,department,officeLocation'
    });

    const people = response.value || [];

    if (people.length === 0) {
      return {
        content: [{ type: "text", text: `No people found matching '${query}'.` }]
      };
    }

    const peopleList = people.map((person, index) => {
      const email = person.emailAddresses?.[0]?.address || 'no email';
      const title = person.jobTitle || '';
      const dept = person.department ? ` — ${person.department}` : '';
      const office = person.officeLocation ? ` (${person.officeLocation})` : '';

      const titleLine = title ? `\n   ${title}${dept}${office}` : '';
      return `${index + 1}. ${person.displayName}\n   ${email}${titleLine}`;
    }).join('\n\n');

    return {
      content: [{
        type: "text",
        text: `Found ${people.length} people matching '${query}':\n\n${peopleList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }],
        isError: true
      };
    }

    return {
      content: [{ type: "text", text: `Error searching people: ${error.message}` }],
      isError: true
    };
  }
}

module.exports = handleSearchPeople;
