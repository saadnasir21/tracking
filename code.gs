const SHIPNOC_CONFIG = {
  propertyKeys: {
    userId: 'SHIPNOC_USER_ID',
    password: 'SHIPNOC_PASSWORD'
  },
  signatureSuffix: 'noc@#',
  endpoints: [
    { url: 'https://shipnoc.com/TrackingAPI/api/Tracking/GetTrackingDetails', methods: ['get', 'post'] },
    { url: 'https://shipnoc.com/TrackingAPI/api/Tracking/GetStatusHistory', methods: ['get', 'post'] },
    { url: 'https://shipnoc.com/TrackingAPI/api/Tracking/GetTrackingStatus', methods: ['get', 'post'] },
    { url: 'https://api.shipnoc.com/TrackingAPI/api/Tracking/GetTrackingDetails', methods: ['get', 'post'] }
  ],
  trackingParamNames: [
    'trackingNo',
    'trackingNumber',
    'TrackingNo',
    'TrackingNumber',
    'trackingID',
    'TrackingID'
  ],
  signatureParamNames: ['signature', 'Signature'],
  baseRequestOptions: {
    followRedirects: true,
    muteHttpExceptions: true,
    contentType: 'application/json; charset=utf-8',
    headers: {
      Accept: 'application/json, text/plain, */*'
    }
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ShipNoc')
    .addItem('Update parcel statuses', 'updateParcelStatuses')
    .addToUi();
}

function updateParcelStatuses() {
  const signature = buildShipnocSignature_();
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const urlRange = sheet.getRange(2, 11, lastRow - 1, 1); // Column K
  const urls = urlRange.getValues();
  const results = [];

  urls.forEach((row, index) => {
    const url = row[0];
    if (!url) {
      results.push(['']);
      return;
    }

    const trackingId = extractTrackingId_(url);
    if (!trackingId) {
      results.push(['Invalid tracking link']);
      return;
    }

    try {
      const status = fetchLatestStatus_(trackingId, signature);
      results.push([status || 'No tracking data available']);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      Logger.log('Failed to update %s: %s', trackingId, message);
      results.push(['Error: ' + message]);
    }

    // Friendly rate limiting to avoid throttling on the API side.
    Utilities.sleep(400);
  });

  sheet.getRange(2, 12, results.length, 1).setValues(results); // Column L
}

function buildShipnocSignature_() {
  const properties = PropertiesService.getScriptProperties();
  const userId = properties.getProperty(SHIPNOC_CONFIG.propertyKeys.userId);
  const password = properties.getProperty(SHIPNOC_CONFIG.propertyKeys.password);

  if (!userId || !password) {
    throw new Error('Missing ShipNoc credentials. Set the SHIPNOC_USER_ID and SHIPNOC_PASSWORD script properties.');
  }

  return userId + password + SHIPNOC_CONFIG.signatureSuffix;
}

function fetchLatestStatus_(trackingId, signature) {
  let lastError;

  for (let i = 0; i < SHIPNOC_CONFIG.endpoints.length; i++) {
    const endpoint = SHIPNOC_CONFIG.endpoints[i];
    const attempts = buildShipnocRequestAttempts_(endpoint, trackingId, signature);

    for (let j = 0; j < attempts.length; j++) {
      const attempt = attempts[j];
      const response = executeShipnocRequest_(attempt);
      if (!response) {
        continue;
      }

      const { code, text } = response;
      if (code === 200 && text) {
        const parsed = parseTrackingPayload_(text);
        if (parsed) {
          return parsed;
        }

        lastError = new Error('Unable to parse response from ' + endpoint.url);
        continue;
      }

      if (code === 404) {
        lastError = new Error('ShipNoc returned 404 for ' + endpoint.url + '. Confirm the tracking number and signature.');
        continue;
      }

      lastError = new Error('HTTP ' + code + ' calling ' + endpoint.url);
    }
  }

  if (lastError) {
    throw lastError;
  }

  throw new Error('Unable to reach ShipNoc tracking endpoints.');
}

function buildShipnocRequestAttempts_(endpoint, trackingId, signature) {
  const attempts = [];
  if (!endpoint || !endpoint.url) {
    return attempts;
  }

  const methods = endpoint.methods && endpoint.methods.length ? endpoint.methods : ['get'];
  const trackingNames = SHIPNOC_CONFIG.trackingParamNames;
  const signatureNames = SHIPNOC_CONFIG.signatureParamNames;
  const seen = new Set();

  if (methods.indexOf('get') !== -1) {
    for (let i = 0; i < trackingNames.length; i++) {
      for (let j = 0; j < signatureNames.length; j++) {
        const trackingParam = trackingNames[i];
        const signatureParam = signatureNames[j];
        const url = endpoint.url +
          '?' + encodeURIComponent(trackingParam) + '=' + encodeURIComponent(trackingId) +
          '&' + encodeURIComponent(signatureParam) + '=' + encodeURIComponent(signature);
        const key = 'GET:' + url;
        if (seen.has(key)) {
          continue;
        }
        seen.add(key);
        attempts.push({
          method: 'get',
          url: url
        });
      }
    }
  }

  if (methods.indexOf('post') !== -1) {
    for (let i = 0; i < trackingNames.length; i++) {
      for (let j = 0; j < signatureNames.length; j++) {
        const trackingParam = trackingNames[i];
        const signatureParam = signatureNames[j];
        const payload = {};
        payload[trackingParam] = trackingId;
        payload[signatureParam] = signature;
        const body = JSON.stringify(payload);
        const key = 'POST:' + endpoint.url + ':' + body;
        if (seen.has(key)) {
          continue;
        }
        seen.add(key);
        attempts.push({
          method: 'post',
          url: endpoint.url,
          body: body
        });
      }
    }
  }

  return attempts;
}

function executeShipnocRequest_(attempt) {
  if (!attempt) {
    return null;
  }

  const options = Object.assign({}, SHIPNOC_CONFIG.baseRequestOptions);
  options.method = attempt.method || 'get';

  if (options.method.toLowerCase() === 'post') {
    options.payload = attempt.body || '';
  } else {
    delete options.payload;
  }

  try {
    const response = UrlFetchApp.fetch(attempt.url, options);
    return {
      code: response.getResponseCode(),
      text: response.getContentText()
    };
  } catch (error) {
    Logger.log('Error calling %s %s: %s', (attempt.method || 'GET').toUpperCase(), attempt.url, error);
    return null;
  }
}

function extractTrackingId_(value) {
  if (!value) {
    return '';
  }

  const text = String(value).trim();
  if (!text) {
    return '';
  }

  const searchText = text;
  for (let i = 0; i < SHIPNOC_CONFIG.trackingParamNames.length; i++) {
    const paramName = SHIPNOC_CONFIG.trackingParamNames[i];
    const regex = new RegExp('[?&]' + escapeRegex_(paramName) + '=([^&#]+)', 'i');
    const match = searchText.match(regex);
    if (match && match[1]) {
      return safeDecodeURIComponent_(match[1]).trim();
    }
  }

  const genericIdMatch = searchText.match(/[?&]id=([^&#]+)/i);
  if (genericIdMatch && genericIdMatch[1]) {
    return safeDecodeURIComponent_(genericIdMatch[1]).trim();
  }

  const alphanumericMatches = searchText.match(/\b[0-9A-Za-z]{6,}\b/g);
  if (alphanumericMatches) {
    const candidate = alphanumericMatches.find((value) => /\d/.test(value));
    if (candidate) {
      return candidate;
    }
  }

  return '';
}

function escapeRegex_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function safeDecodeURIComponent_(value) {
  try {
    return decodeURIComponent(value);
  } catch (error) {
    return value;
  }
}

function parseTrackingPayload_(payload) {
  if (!payload) {
    return '';
  }

  const trimmed = payload.trim();
  if (!trimmed) {
    return '';
  }

  const jsonResult = tryParseTrackingJson_(trimmed);
  if (jsonResult) {
    return jsonResult;
  }

  return parseTrackingHtml_(trimmed);
}

function tryParseTrackingJson_(text) {
  try {
    const data = JSON.parse(text);
    const candidate = findLatestStatusEntry_(data);
    if (candidate) {
      return formatStatusEntry_(candidate);
    }

    if (typeof data === 'string') {
      return data;
    }
  } catch (error) {
    // Ignore JSON parse errors; fall back to HTML parsing.
  }

  return '';
}

function findLatestStatusEntry_(root) {
  const visited = new Set();
  const candidates = [];

  function walk(node) {
    if (!node || typeof node !== 'object') {
      return;
    }

    if (visited.has(node)) {
      return;
    }
    visited.add(node);

    if (Array.isArray(node)) {
      node.forEach(walk);
      return;
    }

    const keys = Object.keys(node);
    const lowerKeys = keys.map((key) => key.toLowerCase());
    const hasStatusKey = lowerKeys.some((key) => key.includes('status'));

    if (hasStatusKey) {
      candidates.push(node);
    }

    keys.forEach((key) => walk(node[key]));
  }

  walk(root);

  if (!candidates.length) {
    return null;
  }

  candidates.sort((a, b) => {
    const timeA = parseStatusDate_(a);
    const timeB = parseStatusDate_(b);
    if (!timeA && !timeB) {
      return 0;
    }
    if (!timeA) {
      return 1;
    }
    if (!timeB) {
      return -1;
    }
    return timeB.getTime() - timeA.getTime();
  });

  return candidates[0];
}

function parseStatusDate_(entry) {
  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const dateKeys = Object.keys(entry).filter((key) => {
    const lower = key.toLowerCase();
    return lower.includes('date') || lower.includes('time') || lower.includes('updated');
  });

  for (let i = 0; i < dateKeys.length; i++) {
    const value = entry[dateKeys[i]];
    if (!value) {
      continue;
    }

    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }

  return null;
}

function formatStatusEntry_(entry) {
  if (!entry || typeof entry !== 'object') {
    return '';
  }

  const statusKeys = Object.keys(entry).filter((key) => key.toLowerCase().includes('status'));
  const locationKeys = Object.keys(entry).filter((key) => /location|city|station/i.test(key));
  const remarksKeys = Object.keys(entry).filter((key) => /remark|message|details|description/i.test(key));

  const statusText = statusKeys
    .map((key) => entry[key])
    .find((value) => typeof value === 'string' && value.trim());

  const locationText = locationKeys
    .map((key) => entry[key])
    .find((value) => typeof value === 'string' && value.trim());

  const remarksText = remarksKeys
    .map((key) => entry[key])
    .find((value) => typeof value === 'string' && value.trim());

  const date = parseStatusDate_(entry);
  const formattedDate = date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '';

  const parts = [];

  if (statusText) {
    parts.push(statusText.trim());
  }

  if (remarksText && (!statusText || !remarksText.includes(statusText))) {
    parts.push(remarksText.trim());
  }

  if (locationText) {
    parts.push(locationText.trim());
  }

  if (formattedDate) {
    parts.push(formattedDate);
  }

  return parts.join(' | ');
}

function parseTrackingHtml_(html) {
  const withoutScripts = html.replace(/<script[\s\S]*?<\/script>/gi, '');
  const withoutStyles = withoutScripts.replace(/<style[\s\S]*?<\/style>/gi, '');
  const text = withoutStyles.replace(/<[^>]+>/g, '\n');
  const lines = text
    .split(/\n+/)
    .map((line) => line.trim())
    .filter((line) => line);

  if (!lines.length) {
    return '';
  }

  for (let i = 0; i < lines.length; i++) {
    if (/status/i.test(lines[i])) {
      const next = lines[i + 1] || '';
      if (next && !/status/i.test(next)) {
        return (lines[i] + ' - ' + next).trim();
      }
      return lines[i];
    }
  }

  return lines[0];
}

function setShipnocCredentials(userId, password) {
  if (!userId || !password) {
    throw new Error('Both userId and password are required.');
  }

  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(SHIPNOC_CONFIG.propertyKeys.userId, String(userId));
  properties.setProperty(SHIPNOC_CONFIG.propertyKeys.password, String(password));
}
