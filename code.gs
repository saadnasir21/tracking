const SHIPNOC_CONFIG = {
  propertyKeys: {
    userId: 'SHIPNOC_USER_ID',
    password: 'SHIPNOC_PASSWORD'
  },
  signatureSuffix: 'noc@#',
  apiCandidates: [
    'https://shipnoc.com/TrackingAPI/api/Tracking/GetTrackingDetails',
    'https://shipnoc.com/TrackingAPI/api/Tracking/GetStatusHistory',
    'https://shipnoc.com/TrackingAPI/api/Tracking/GetTrackingStatus',
    'https://api.shipnoc.com/TrackingAPI/api/Tracking/GetTrackingDetails'
  ],
  requestOptions: {
    method: 'get',
    followRedirects: true,
    muteHttpExceptions: true,
    contentType: 'application/json; charset=utf-8'
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

  for (let i = 0; i < SHIPNOC_CONFIG.apiCandidates.length; i++) {
    const endpoint = SHIPNOC_CONFIG.apiCandidates[i];

    const response = callShipnocEndpoint_(endpoint, trackingId, signature);
    if (!response) {
      continue;
    }

    const { code, text } = response;
    if (code === 200 && text) {
      const parsed = parseTrackingPayload_(text);
      if (parsed) {
        return parsed;
      }

      lastError = new Error('Unable to parse response from ' + endpoint);
      continue;
    }

    lastError = new Error('HTTP ' + code + ' calling ' + endpoint);
  }

  if (lastError) {
    throw lastError;
  }

  throw new Error('Unable to reach ShipNoc tracking endpoints.');
}

function callShipnocEndpoint_(endpoint, trackingId, signature) {
  if (!endpoint) {
    return null;
  }

  const params = Object.assign({}, SHIPNOC_CONFIG.requestOptions);
  const queryStrings = [
    'trackingNumber=' + encodeURIComponent(trackingId),
    'trackingNo=' + encodeURIComponent(trackingId),
    'trackingnumber=' + encodeURIComponent(trackingId)
  ];

  const signatureParam = 'signature=' + encodeURIComponent(signature);
  const urlsToTry = queryStrings.map((qs) => endpoint + '?' + qs + '&' + signatureParam);

  for (let i = 0; i < urlsToTry.length; i++) {
    const url = urlsToTry[i];
    try {
      const response = UrlFetchApp.fetch(url, params);
      return {
        code: response.getResponseCode(),
        text: response.getContentText()
      };
    } catch (error) {
      Logger.log('Error calling %s: %s', url, error);
    }
  }

  return null;
}

function extractTrackingId_(value) {
  if (!value) {
    return '';
  }

  const text = String(value).trim();
  if (!text) {
    return '';
  }

  const urlMatch = text.match(/[?&]ID=(\w+)/i);
  if (urlMatch) {
    return urlMatch[1];
  }

  const digitsMatch = text.match(/\d{6,}/);
  if (digitsMatch) {
    return digitsMatch[0];
  }

  return '';
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
