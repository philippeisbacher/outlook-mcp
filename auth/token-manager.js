/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const https = require('https');
const querystring = require('querystring');
require('dotenv').config();
const config = require('../config');

// Global variable to store tokens
let cachedTokens = null;
let refreshPromise = null;

// Token refresh buffer (5 minutes before expiry)
const REFRESH_BUFFER_MS = 5 * 60 * 1000;

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`[DEBUG] Attempting to load tokens from: ${tokenPath}`);
    console.error(`[DEBUG] HOME directory: ${process.env.HOME}`);
    console.error(`[DEBUG] Full resolved path: ${tokenPath}`);
    
    // Log file existence and details
    if (!fs.existsSync(tokenPath)) {
      console.error('[DEBUG] Token file does not exist');
      return null;
    }
    
    const stats = fs.statSync(tokenPath);
    console.error(`[DEBUG] Token file stats:
      Size: ${stats.size} bytes
      Created: ${stats.birthtime}
      Modified: ${stats.mtime}`);
    
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    console.error('[DEBUG] Token file contents length:', tokenData.length);
    console.error('[DEBUG] Token file first 200 characters:', tokenData.slice(0, 200));
    
    try {
      const tokens = JSON.parse(tokenData);
      console.error('[DEBUG] Parsed tokens keys:', Object.keys(tokens));
      
      // Log each key's value to see what's present
      Object.keys(tokens).forEach(key => {
        console.error(`[DEBUG] ${key}: ${typeof tokens[key]}`);
      });
      
      // Check for access token presence
      if (!tokens.access_token) {
        console.error('[DEBUG] No access_token found in tokens');
        return null;
      }
      
      // Update the cache (even if expired, we need it for refresh)
      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('[DEBUG] Error parsing token JSON:', parseError);
      return null;
    }
  } catch (error) {
    console.error('[DEBUG] Error loading token cache:', error);
    return null;
  }
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`Saving tokens to: ${tokenPath}`);
    
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    console.error('Tokens saved successfully');
    
    // Update the cache
    cachedTokens = tokens;
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error);
    return false;
  }
}

/**
 * Checks if the token needs refresh (expired or about to expire)
 * @returns {boolean} - Whether the token needs refresh
 */
function needsRefresh() {
  if (!cachedTokens || !cachedTokens.expires_at) {
    return true;
  }
  return Date.now() >= (cachedTokens.expires_at - REFRESH_BUFFER_MS);
}

/**
 * Refreshes the access token using the refresh token
 * @returns {Promise<object>} - The new tokens
 */
async function refreshAccessToken() {
  if (!cachedTokens || !cachedTokens.refresh_token) {
    throw new Error('No refresh token available');
  }

  // Prevent multiple concurrent refresh attempts
  if (refreshPromise) {
    console.error('[TOKEN-MANAGER] Refresh already in progress, waiting...');
    return refreshPromise;
  }

  console.error('[TOKEN-MANAGER] Refreshing access token...');

  const postData = querystring.stringify({
    client_id: process.env.MS_CLIENT_ID || process.env.OUTLOOK_CLIENT_ID || config.AUTH_CONFIG.clientId,
    client_secret: process.env.MS_CLIENT_SECRET || process.env.OUTLOOK_CLIENT_SECRET || config.AUTH_CONFIG.clientSecret,
    grant_type: 'refresh_token',
    refresh_token: cachedTokens.refresh_token,
    scope: config.AUTH_CONFIG.scopes.join(' ')
  });

  refreshPromise = new Promise((resolve, reject) => {
    const requestOptions = {
      hostname: 'login.microsoftonline.com',
      path: '/common/oauth2/v2.0/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(requestOptions, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        refreshPromise = null;
        try {
          const responseBody = JSON.parse(data);
          if (res.statusCode >= 200 && res.statusCode < 300) {
            // Update tokens
            cachedTokens.access_token = responseBody.access_token;
            if (responseBody.refresh_token) {
              cachedTokens.refresh_token = responseBody.refresh_token;
            }
            cachedTokens.expires_in = responseBody.expires_in;
            cachedTokens.expires_at = Date.now() + (responseBody.expires_in * 1000);
            cachedTokens.scope = responseBody.scope || cachedTokens.scope;

            // Save to file
            saveTokenCache(cachedTokens);
            console.error('[TOKEN-MANAGER] Access token refreshed successfully');
            resolve(cachedTokens);
          } else {
            console.error('[TOKEN-MANAGER] Token refresh failed:', responseBody);
            reject(new Error(responseBody.error_description || `Token refresh failed: ${res.statusCode}`));
          }
        } catch (e) {
          console.error('[TOKEN-MANAGER] Error parsing refresh response:', e);
          reject(e);
        }
      });
    });

    req.on('error', (error) => {
      refreshPromise = null;
      console.error('[TOKEN-MANAGER] Refresh request error:', error);
      reject(error);
    });

    req.write(postData);
    req.end();
  });

  return refreshPromise;
}

/**
 * Gets the current access token, refreshing if necessary
 * @returns {Promise<string|null>} - The access token or null if not available
 */
async function getAccessToken() {
  // Load tokens if not cached
  if (!cachedTokens) {
    loadTokenCache();
  }

  if (!cachedTokens || !cachedTokens.access_token) {
    return null;
  }

  // Check if we need to refresh
  if (needsRefresh()) {
    if (cachedTokens.refresh_token) {
      try {
        await refreshAccessToken();
      } catch (error) {
        console.error('[TOKEN-MANAGER] Failed to refresh token:', error.message);
        return null;
      }
    } else {
      console.error('[TOKEN-MANAGER] Token expired and no refresh token available');
      return null;
    }
  }

  return cachedTokens.access_token;
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000) // 1 hour
  };
  
  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  refreshAccessToken,
  needsRefresh,
  createTestTokens
};
