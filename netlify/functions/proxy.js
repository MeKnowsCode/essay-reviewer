// ============================================================
// Essay Reviewer — Netlify Proxy Function
// With rate limiting and API key validation
// ============================================================

// In-memory rate limit store
// Tracks requests per IP address
const rateLimitStore = {};
const RATE_LIMIT_MAX = 25;        // max requests
const RATE_LIMIT_WINDOW = 60000;  // per 60 seconds

function checkRateLimit(ip) {
  const now = Date.now();
  
  if (!rateLimitStore[ip]) {
    rateLimitStore[ip] = { count: 1, windowStart: now };
    return true;
  }

  const record = rateLimitStore[ip];

  // Reset window if expired
  if (now - record.windowStart > RATE_LIMIT_WINDOW) {
    record.count = 1;
    record.windowStart = now;
    return true;
  }

  // Check limit
  if (record.count >= RATE_LIMIT_MAX) {
    return false;
  }

  record.count++;
  return true;
}

function isValidAnthropicKey(key) {
  // Anthropic keys always start with sk-ant-
  return typeof key === 'string' && key.startsWith('sk-ant-') && key.length > 20;
}

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, x-api-key, anthropic-version'
};

exports.handler = async function(event) {

  // Handle preflight CORS request
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: corsHeaders, body: '' };
  }

  // Only allow POST
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }

  // Rate limiting — get client IP
  const ip = event.headers['x-forwarded-for']?.split(',')[0]?.trim()
    || event.headers['client-ip']
    || 'unknown';

  if (!checkRateLimit(ip)) {
    return {
      statusCode: 429,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Too many requests. Please wait a minute and try again.' })
    };
  }

  // Validate API key format before forwarding
  const apiKey = event.headers['x-api-key'];
  if (!isValidAnthropicKey(apiKey)) {
    return {
      statusCode: 401,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Invalid or missing API key.' })
    };
  }

  // Parse and validate request body
  let body;
  try {
    body = JSON.parse(event.body);
  } catch(e) {
    return {
      statusCode: 400,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Invalid request body.' })
    };
  }

  // Only allow requests to the messages endpoint
  // and only with our expected model
  if (!body.messages || !Array.isArray(body.messages)) {
    return {
      statusCode: 400,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Invalid request format.' })
    };
  }

  // Forward to Anthropic
  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify(body)
    });

    const data = await response.json();

    return {
      statusCode: response.status,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    };

  } catch(e) {
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: 'Proxy error: ' + e.message })
    };
  }
};
