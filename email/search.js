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
  const folder = args.folder || null;
  const requestedCount = args.count || 10;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;
  const mailbox = args.mailbox || null;
  const category = args.category || null;  // Filter by category/label

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
      { hasAttachments, unreadOnly, before, after, category },
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
 * Sort emails by receivedDateTime client-side (used after $search, which ignores $orderby)
 * @param {Array} emails - Email list
 * @param {string} sortOrder - 'asc' or 'desc'
 * @returns {Array} - Sorted email list (new array, original not mutated)
 */
function sortByDate(emails, sortOrder) {
  return [...emails].sort((a, b) => {
    const dateA = new Date(a.receivedDateTime).getTime();
    const dateB = new Date(b.receivedDateTime).getTime();
    return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
  });
}

/**
 * Apply date/boolean/category filters client-side (used after $search results)
 * @param {Array} emails - Email list from API
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly, before, after, category)
 * @returns {Array} - Filtered email list
 */
function applyClientSideFilters(emails, filterTerms) {
  let result = emails;

  // after/before are handled server-side via KQL received: in $search

  if (filterTerms.hasAttachments === true) {
    result = result.filter(e => e.hasAttachments === true);
  }

  if (filterTerms.unreadOnly === true) {
    result = result.filter(e => e.isRead === false);
  }

  if (filterTerms.category) {
    result = result.filter(e => e.categories && e.categories.includes(filterTerms.category));
  }

  return result;
}

/**
 * Execute a search against the Graph API
 * - Text terms (from, subject, to, query): use $search; date/boolean filters applied client-side
 *   ($search and $filter are mutually exclusive in Graph API)
 * - Boolean/date filters only: use $filter with $orderby
 * - No terms, no filters: basic listing
 * @param {string} endpoint - API endpoint
 * @param {string} accessToken - Access token
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly, before, after, category)
 * @param {number} maxCount - Maximum number of results to retrieve
 * @param {string} sortOrder - Sort order ('asc' or 'desc')
 * @param {number} skip - Number of results to skip
 * @returns {Promise<object>} - Search results
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount, sortOrder = 'desc', skip = 0) {
  const orderBy = `receivedDateTime ${sortOrder}`;
  const hasTextTerms = !!(searchTerms.query || searchTerms.from || searchTerms.to || searchTerms.subject);

  // Path 1: Text search → use $search only (no $filter allowed with $search in Graph API)
  // Date filters are embedded as KQL received: ranges in $search (server-side).
  // Boolean/category filters are still applied client-side on the results.
  if (hasTextTerms) {
    const apiFetchCount = Math.min(50, maxCount);

    const params = buildSearchParams(searchTerms, filterTerms, apiFetchCount, skip);
    console.error("Executing $search with params:", params);

    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, apiFetchCount);
    let filtered = applyClientSideFilters(response.value || [], filterTerms);
    filtered = sortByDate(filtered, sortOrder);
    filtered = filtered.slice(0, maxCount);
    return { value: filtered };
  }

  // Path 2: No text terms, only boolean/date/category filters → use $filter with $orderby
  const hasFilters = filterTerms.hasAttachments === true ||
                     filterTerms.unreadOnly === true ||
                     filterTerms.before ||
                     filterTerms.after ||
                     filterTerms.category;

  if (hasFilters) {
    const filterParams = {
      $top: Math.min(50, maxCount),
      $select: config.EMAIL_SELECT_FIELDS,
      $orderby: orderBy
    };
    if (skip > 0) filterParams.$skip = skip;
    addFilters(filterParams, filterTerms);
    console.error("Executing $filter search with params:", filterParams);

    return await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterParams, maxCount);
  }

  // Path 3: No terms, no filters → basic listing
  console.error("No search terms or filters, returning basic listing");
  const basicParams = {
    $top: Math.min(50, maxCount),
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: orderBy
  };
  if (skip > 0) basicParams.$skip = skip;

  return await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
}

/**
 * Build $search query parameters from text search terms.
 * Note: $filter and $orderby must NOT be included — they are incompatible with $search in Graph API.
 * Date filters (after/before) are included as KQL received: ranges.
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (after, before handled here; rest applied client-side)
 * @param {number} count - Maximum number of results
 * @param {number} skip - Number of results to skip
 * @returns {object} - Query parameters with $search set
 */
function buildSearchParams(searchTerms, filterTerms, count, skip = 0) {
  const params = {
    $top: count,
    $select: config.EMAIL_SELECT_FIELDS
  };

  if (skip > 0) {
    params.$skip = skip;
  }

  const kqlTerms = [];

  if (searchTerms.query) {
    kqlTerms.push(searchTerms.query);
  }

  if (searchTerms.subject) {
    kqlTerms.push(`subject:${kqlPhrase(searchTerms.subject)}`);
  }

  if (searchTerms.from) {
    kqlTerms.push(`from:${kqlPhrase(searchTerms.from)}`);
  }

  if (searchTerms.to) {
    kqlTerms.push(`to:${kqlPhrase(searchTerms.to)}`);
  }

  // Add KQL date range filter — server-side filtering via $search
  // Only add if at least one date is parseable (Bug 3: avoid received:.. with empty sides)
  if (filterTerms.after || filterTerms.before) {
    const afterKql = filterTerms.after ? (toKqlDate(filterTerms.after) || '') : '';
    const beforeKql = filterTerms.before ? (toKqlDate(filterTerms.before) || '') : '';
    if (afterKql || beforeKql) {
      kqlTerms.push(`received:${afterKql}..${beforeKql}`);
    }
  }

  // Wrap KQL terms in outer double quotes as required by Graph API.
  // Inner double quotes (from kqlPhrase) are escaped to keep the outer wrapper intact.
  if (kqlTerms.length > 0) {
    const kqlExpression = kqlTerms.join(' ').replace(/"/g, '\\"');
    params.$search = `"${kqlExpression}"`;
  }

  return params;
}

/**
 * Wrap a KQL property value in double quotes if it contains spaces (phrase search).
 * The outer $search wrapper requires inner quotes to be escaped separately.
 * @param {string} value - The property value (e.g. "John Doe" or "godaddy")
 * @returns {string} - Unquoted for single words, "..." wrapped for phrases
 */
function kqlPhrase(value) {
  if (!value.includes(' ')) return value;
  return `"${value}"`;
}

/**
 * Convert a date string to KQL date format (MM/DD/YYYY) for use in $search received: filters
 * @param {string} dateStr - Date string (ISO, "YYYY-MM-DD", relative)
 * @returns {string|null} - Date in MM/DD/YYYY format or null if invalid
 */
function toKqlDate(dateStr) {
  const iso = parseDate(dateStr);
  if (!iso) return null;
  const d = new Date(iso);
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const yyyy = d.getUTCFullYear();
  return `${mm}/${dd}/${yyyy}`;
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
 * Add filters to query parameters (boolean + date + category filters)
 * @param {object} params - Query parameters
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly, before, after, category)
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

  // Add category filter
  if (filterTerms.category) {
    // OData filter for categories array contains the specified category
    filterConditions.push(`categories/any(c:c eq '${filterTerms.category}')`);
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
    const categories = email.categories && email.categories.length > 0
      ? `[${email.categories.join(', ')}] `
      : '';
    const displayIndex = skip + index + 1;

    return `${displayIndex}. ${readStatus}${categories}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
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
