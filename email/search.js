/**
 * Improved search emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

/**
 * Search emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;
  const mailbox = args.mailbox || null;

  // New parameters for advanced filtering
  const before = args.before || null;  // ISO date string or relative like "2024-01-01"
  const after = args.after || null;    // ISO date string or relative like "2024-01-01"
  const sortOrder = args.sortOrder || 'desc';  // 'asc' for oldest first, 'desc' for newest first
  const skip = args.skip || 0;  // Number of emails to skip (for pagination)

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Resolve the folder path (with optional shared mailbox support)
    const endpoint = await resolveFolderPath(accessToken, folder, mailbox);
    console.error(`Using endpoint: ${endpoint} for folder: ${folder} in mailbox: ${mailbox || 'primary'}`);

    // Execute progressive search with pagination
    const response = await progressiveSearch(
      endpoint,
      accessToken,
      { query, from, to, subject },
      { hasAttachments, unreadOnly, before, after },
      requestedCount,
      sortOrder,
      skip
    );

    return formatSearchResults(response, sortOrder, skip);
  } catch (error) {
    // Handle authentication errors
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    // General error response
    return {
      content: [{
        type: "text",
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

/**
 * Execute a search with progressively simpler fallback strategies
 * @param {string} endpoint - API endpoint
 * @param {string} accessToken - Access token
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly, before, after)
 * @param {number} maxCount - Maximum number of results to retrieve
 * @param {string} sortOrder - Sort order ('asc' or 'desc')
 * @param {number} skip - Number of results to skip
 * @returns {Promise<object>} - Search results
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount, sortOrder = 'desc', skip = 0) {
  // Track search strategies attempted
  const searchAttempts = [];
  const orderBy = `receivedDateTime ${sortOrder}`;

  // 1. Try combined search (most specific)
  try {
    const params = buildSearchParams(searchTerms, filterTerms, Math.min(50, maxCount), orderBy, skip);
    console.error("Attempting combined search with params:", params);
    searchAttempts.push("combined-search");

    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, maxCount);
    if (response.value && response.value.length > 0) {
      console.error(`Combined search successful: found ${response.value.length} results`);
      return response;
    }
  } catch (error) {
    console.error(`Combined search failed: ${error.message}`);
  }

  // 2. Try each search term individually, starting with most specific
  const searchPriority = ['subject', 'from', 'to', 'query'];

  for (const term of searchPriority) {
    if (searchTerms[term]) {
      try {
        console.error(`Attempting search with only ${term}: "${searchTerms[term]}"`);
        searchAttempts.push(`single-term-${term}`);

        // For single term search, only use $search with that term
        const simplifiedParams = {
          $top: Math.min(50, maxCount),
          $select: config.EMAIL_SELECT_FIELDS,
          $orderby: orderBy
        };

        // Add skip if specified
        if (skip > 0) {
          simplifiedParams.$skip = skip;
        }

        // Add the search term in the appropriate KQL syntax
        if (term === 'query') {
          // General query doesn't need a prefix
          simplifiedParams.$search = `"${searchTerms[term]}"`;
        } else {
          // Specific field searches use field:value syntax
          simplifiedParams.$search = `${term}:"${searchTerms[term]}"`;
        }

        // Add boolean and date filters if applicable
        addFilters(simplifiedParams, filterTerms);

        const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, simplifiedParams, maxCount);
        if (response.value && response.value.length > 0) {
          console.error(`Search with ${term} successful: found ${response.value.length} results`);
          return response;
        }
      } catch (error) {
        console.error(`Search with ${term} failed: ${error.message}`);
      }
    }
  }

  // 3. Try with only filters (boolean + date)
  const hasFilters = filterTerms.hasAttachments === true ||
                     filterTerms.unreadOnly === true ||
                     filterTerms.before ||
                     filterTerms.after;

  if (hasFilters) {
    try {
      console.error("Attempting search with only filters");
      searchAttempts.push("filters-only");

      const filterOnlyParams = {
        $top: Math.min(50, maxCount),
        $select: config.EMAIL_SELECT_FIELDS,
        $orderby: orderBy
      };

      // Add skip if specified
      if (skip > 0) {
        filterOnlyParams.$skip = skip;
      }

      // Add the filters
      addFilters(filterOnlyParams, filterTerms);

      const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterOnlyParams, maxCount);
      console.error(`Filter search found ${response.value?.length || 0} results`);
      return response;
    } catch (error) {
      console.error(`Filter search failed: ${error.message}`);
    }
  }

  // 4. Final fallback: just get emails with pagination and sort order
  console.error("All search strategies failed, falling back to basic listing");
  searchAttempts.push("basic-listing");

  const basicParams = {
    $top: Math.min(50, maxCount),
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: orderBy
  };

  // Add skip if specified
  if (skip > 0) {
    basicParams.$skip = skip;
  }

  const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
  console.error(`Basic listing found ${response.value?.length || 0} results`);

  // Add a note to the response about the search attempts
  response._searchInfo = {
    attemptsCount: searchAttempts.length,
    strategies: searchAttempts,
    originalTerms: searchTerms,
    filterTerms: filterTerms
  };

  return response;
}

/**
 * Build search parameters from search terms and filter terms
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly, before, after)
 * @param {number} count - Maximum number of results
 * @param {string} orderBy - Order by clause
 * @param {number} skip - Number of results to skip
 * @returns {object} - Query parameters
 */
function buildSearchParams(searchTerms, filterTerms, count, orderBy = 'receivedDateTime desc', skip = 0) {
  const params = {
    $top: count,
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: orderBy
  };

  // Add skip if specified
  if (skip > 0) {
    params.$skip = skip;
  }

  // Handle search terms
  const kqlTerms = [];

  if (searchTerms.query) {
    // General query doesn't need a prefix
    kqlTerms.push(searchTerms.query);
  }

  if (searchTerms.subject) {
    kqlTerms.push(`subject:"${searchTerms.subject}"`);
  }

  if (searchTerms.from) {
    kqlTerms.push(`from:"${searchTerms.from}"`);
  }

  if (searchTerms.to) {
    kqlTerms.push(`to:"${searchTerms.to}"`);
  }

  // Add $search if we have any search terms
  if (kqlTerms.length > 0) {
    params.$search = kqlTerms.join(' ');
  }

  // Add filters (boolean + date)
  addFilters(params, filterTerms);

  return params;
}

/**
 * Parse a date string into ISO format for OData filtering
 * Supports: ISO dates, "YYYY-MM-DD", relative dates like "today", "yesterday", "7 days ago"
 * @param {string} dateStr - Date string to parse
 * @returns {string|null} - ISO date string or null if invalid
 */
function parseDate(dateStr) {
  if (!dateStr) return null;

  const str = dateStr.toLowerCase().trim();

  // Handle relative dates
  const now = new Date();

  if (str === 'today') {
    now.setHours(0, 0, 0, 0);
    return now.toISOString();
  }

  if (str === 'yesterday') {
    now.setDate(now.getDate() - 1);
    now.setHours(0, 0, 0, 0);
    return now.toISOString();
  }

  // Match patterns like "7 days ago", "1 week ago", "2 months ago"
  const relativeMatch = str.match(/^(\d+)\s*(day|days|week|weeks|month|months)\s*ago$/);
  if (relativeMatch) {
    const amount = parseInt(relativeMatch[1], 10);
    const unit = relativeMatch[2];

    if (unit.startsWith('day')) {
      now.setDate(now.getDate() - amount);
    } else if (unit.startsWith('week')) {
      now.setDate(now.getDate() - (amount * 7));
    } else if (unit.startsWith('month')) {
      now.setMonth(now.getMonth() - amount);
    }
    now.setHours(0, 0, 0, 0);
    return now.toISOString();
  }

  // Try parsing as a date directly
  const parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) {
    return parsed.toISOString();
  }

  return null;
}

/**
 * Add filters to query parameters (boolean + date filters)
 * @param {object} params - Query parameters
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly, before, after)
 */
function addFilters(params, filterTerms) {
  const filterConditions = [];

  if (filterTerms.hasAttachments === true) {
    filterConditions.push('hasAttachments eq true');
  }

  if (filterTerms.unreadOnly === true) {
    filterConditions.push('isRead eq false');
  }

  // Add date filters
  if (filterTerms.before) {
    const beforeDate = parseDate(filterTerms.before);
    if (beforeDate) {
      filterConditions.push(`receivedDateTime lt ${beforeDate}`);
    }
  }

  if (filterTerms.after) {
    const afterDate = parseDate(filterTerms.after);
    if (afterDate) {
      filterConditions.push(`receivedDateTime ge ${afterDate}`);
    }
  }

  // Add $filter parameter if we have any filter conditions
  if (filterConditions.length > 0) {
    params.$filter = filterConditions.join(' and ');
  }
}

/**
 * Format search results into a readable text format
 * @param {object} response - The API response object
 * @param {string} sortOrder - Sort order used ('asc' or 'desc')
 * @param {number} skip - Number of results skipped
 * @returns {object} - MCP response object
 */
function formatSearchResults(response, sortOrder = 'desc', skip = 0) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{
        type: "text",
        text: `No emails found matching your search criteria.`
      }]
    };
  }

  // Format results
  const emailList = response.value.map((email, index) => {
    const sender = email.from?.emailAddress || { name: 'Unknown', address: 'unknown' };
    const date = new Date(email.receivedDateTime).toLocaleString();
    const readStatus = email.isRead ? '' : '[UNREAD] ';
    const displayIndex = skip + index + 1;

    return `${displayIndex}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
  }).join("\n");

  // Build info string
  let infoStr = '';
  const sortInfo = sortOrder === 'asc' ? 'oldest first' : 'newest first';

  if (skip > 0) {
    infoStr = ` (showing ${skip + 1}-${skip + response.value.length}, ${sortInfo})`;
  } else {
    infoStr = ` (${sortInfo})`;
  }

  // Add search strategy info if available
  let additionalInfo = '';
  if (response._searchInfo) {
    additionalInfo = `\n(Search used ${response._searchInfo.strategies[response._searchInfo.strategies.length - 1]} strategy)`;
  }

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails${infoStr}:${additionalInfo}\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;
