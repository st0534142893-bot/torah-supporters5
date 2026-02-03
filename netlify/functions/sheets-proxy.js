/**
 * Netlify Function - Proxy לגישה ל-Google Apps Script
 * פותר בעיית CORS - הבקשות יוצאות מהשרת ולא מהדפדפן
 * דורש Node.js 18+ (יש fetch מובנה)
 */
const https = require('https');
const http = require('http');

function fetchWithFallback(url) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const client = urlObj.protocol === 'https:' ? https : http;
    
    const options = {
      hostname: urlObj.hostname,
      port: urlObj.port || (urlObj.protocol === 'https:' ? 443 : 80),
      path: urlObj.pathname + urlObj.search,
      method: 'GET',
      headers: {
        'User-Agent': 'Netlify-Function/1.0'
      }
    };

    const req = client.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        try {
          resolve({ ok: res.statusCode === 200, status: res.statusCode, json: () => Promise.resolve(JSON.parse(data)) });
        } catch (e) {
          reject(new Error('Invalid JSON: ' + e.message));
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(10000, () => { req.destroy(); reject(new Error('Timeout')); });
    req.end();
  });
}

exports.handler = async (event, context) => {
  const headers = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers, body: '' };
  }

  try {
    const url = event.queryStringParameters?.url;
    const action = event.queryStringParameters?.action || 'getAll';

    if (!url) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ success: false, error: 'Missing url parameter' })
      };
    }

    const baseUrl = decodeURIComponent(url);
    const sep = baseUrl.includes('?') ? '&' : '?';
    const targetUrl = baseUrl + sep + 'action=' + action;
    
    // ניסיון עם fetch מובנה (Node 18+), אם לא - fallback ל-https
    let res;
    let data;
    try {
      // Node 18+ has native fetch
      if (typeof fetch !== 'undefined') {
        res = await fetch(targetUrl, { 
          method: 'GET',
          headers: { 'User-Agent': 'Netlify-Function/1.0' }
        });
        if (!res.ok) {
          throw new Error(`HTTP ${res.status}: ${res.statusText || 'Unknown error'}`);
        }
        data = await res.json();
      } else {
        throw new Error('fetch not available');
      }
    } catch (fetchError) {
      console.log('fetch לא זמין, משתמש ב-https:', fetchError.message);
      res = await fetchWithFallback(targetUrl);
      if (!res.ok) {
        throw new Error(`HTTP ${res.status}: Unknown error`);
      }
      data = await res.json();
    }

    return {
      statusCode: 200,
      headers,
      body: typeof data === 'string' ? data : JSON.stringify(data)
    };
  } catch (err) {
    console.error('sheets-proxy error:', err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ success: false, error: err.message || 'Unknown error' })
    };
  }
};
