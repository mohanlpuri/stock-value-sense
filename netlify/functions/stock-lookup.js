const SHEET_ID = process.env.GOOGLE_SHEET_ID_TRACKER
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL
const PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n')
const FINNHUB_KEY = process.env.FINNHUB_API_KEY

// Generate JWT for Google Sheets API auth
async function getAccessToken() {
  const header = { alg: 'RS256', typ: 'JWT' }
  const now = Math.floor(Date.now() / 1000)
  const claim = {
    iss: CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  }

  const encode = obj => Buffer.from(JSON.stringify(obj)).toString('base64url')
  const headerB64 = encode(header)
  const claimB64 = encode(claim)
  const unsigned = headerB64 + '.' + claimB64

  const crypto = require('crypto')
  const sign = crypto.createSign('RSA-SHA256')
  sign.update(unsigned)
  const signature = sign.sign(PRIVATE_KEY, 'base64url')
  const jwt = unsigned + '.' + signature

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + jwt
  })
  const data = await res.json()
  console.log('Access token:', data.access_token ? 'ok' : 'failed')
  return data.access_token
}

// Fetch quote data from Finnhub
async function fetchQuote(ticker) {
  const url = `https://finnhub.io/api/v1/quote?symbol=${ticker}&token=${FINNHUB_KEY}`
  const res = await fetch(url)
  const data = await res.json()
  console.log('Finnhub quote status:', res.status)
  return data
}

// Fetch basic financials from Finnhub (PE, Book Value, 52W High/Low)
async function fetchMetrics(ticker) {
  const url = `https://finnhub.io/api/v1/stock/metric?symbol=${ticker}&metric=all&token=${FINNHUB_KEY}`
  const res = await fetch(url)
  const data = await res.json()
  console.log('Finnhub metrics status:', res.status)
  return data.metric || {}
}

// Fetch company profile from Finnhub (sector, name)
async function fetchProfile(ticker) {
  const url = `https://finnhub.io/api/v1/stock/profile2?symbol=${ticker}&token=${FINNHUB_KEY}`
  const res = await fetch(url)
  const data = await res.json()
  console.log('Finnhub profile status:', res.status)
  return data
}

// Fetch analyst recommendations from Finnhub
async function fetchRecommendations(ticker) {
  const url = `https://finnhub.io/api/v1/stock/recommendation?symbol=${ticker}&token=${FINNHUB_KEY}`
  const res = await fetch(url)
  const data = await res.json()
  console.log('Finnhub recommendations status:', res.status)
  // Return most recent recommendation
  return data && data.length > 0 ? data[0] : null
}

// Summarize analyst recommendation
function summarizeRating(rec) {
  if (!rec) return null
  const { strongBuy, buy, hold, sell, strongSell } = rec
  const total = strongBuy + buy + hold + sell + strongSell
  if (total === 0) return null
  const bullish = strongBuy + buy
  const bearish = sell + strongSell
  if (strongBuy > buy && strongBuy > hold) return `Strong Buy (${strongBuy}/${total})`
  if (bullish > hold && bullish > bearish) return `Buy (${bullish}/${total})`
  if (hold >= bullish && hold >= bearish) return `Hold (${hold}/${total})`
  if (bearish > bullish) return `Sell (${bearish}/${total})`
  return `Hold (${hold}/${total})`
}

// Read all rows from Google Sheet
async function readSheet(token) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1`
  const res = await fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  })
  const data = await res.json()
  return data.values || []
}

// Find row index of existing ticker
function findTickerRow(rows, ticker) {
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] && rows[i][1].toUpperCase() === ticker.toUpperCase()) {
      return i + 1
    }
  }
  return -1
}

// Write or update a row in Google Sheet
async function saveToSheet(token, stock) {
  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })
  const row = [
    today,
    stock.ticker,
    stock.price || '',
    stock.week52High || '',
    stock.week52Low || '',
    stock.peRatio || '',
    stock.bookValue || '',
    stock.pbRatio || '',
    stock.sector || '',
    stock.analystRating || ''
  ]

  const rows = await readSheet(token)
  const existingRow = findTickerRow(rows, stock.ticker)

  let url, method, body

  if (existingRow > 0) {
    url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A${existingRow}:J${existingRow}?valueInputOption=RAW`
    method = 'PUT'
    body = JSON.stringify({ values: [row] })
    console.log('Updating existing row:', existingRow)
  } else {
    url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A:J:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`
    method = 'POST'
    body = JSON.stringify({ values: [row] })
    console.log('Appending new row for:', stock.ticker)
  }

  const res = await fetch(url, {
    method,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    body
  })

  const result = await res.json()
  console.log('Sheet write result:', JSON.stringify(result).slice(0, 200))
  return result
}

exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' }
  }

  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  }

  try {
    const body = JSON.parse(event.body)

    // Get Google access token
    const token = await getAccessToken()

    // History request
    if (body.action === 'history') {
      const rows = await readSheet(token)
      return { statusCode: 200, headers, body: JSON.stringify({ rows }) }
    }

    // Stock lookup request
    const ticker = (body.ticker || '').trim().toUpperCase()
    if (!ticker) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: 'Ticker is required' }) }
    }

    console.log('Looking up ticker:', ticker)

    // Fetch all data in parallel
    const [quote, metrics, profile, recommendation] = await Promise.all([
      fetchQuote(ticker),
      fetchMetrics(ticker),
      fetchProfile(ticker),
      fetchRecommendations(ticker)
    ])

    if (!quote || !quote.c) {
      return { statusCode: 404, headers, body: JSON.stringify({ error: 'Stock not found: ' + ticker }) }
    }

    const price = quote.c
    const bookValue = metrics.bookValuePerShareAnnual || null
    const peRatio = metrics.peAnnual || metrics.peTTM || null
    const pbRatio = metrics.pb || null

    const stock = {
      ticker: ticker,
      name: profile.name || ticker,
      price: price ? parseFloat(price.toFixed(2)) : null,
      week52High: metrics['52WeekHigh'] ? parseFloat(metrics['52WeekHigh'].toFixed(2)) : null,
      week52Low: metrics['52WeekLow'] ? parseFloat(metrics['52WeekLow'].toFixed(2)) : null,
      peRatio: peRatio ? parseFloat(peRatio.toFixed(2)) : null,
      bookValue: bookValue ? parseFloat(bookValue.toFixed(2)) : null,
      pbRatio: pbRatio ? parseFloat(pbRatio.toFixed(2)) : null,
      sector: profile.finnhubIndustry || null,
      analystRating: summarizeRating(recommendation),
      change: quote.d ? parseFloat(quote.d.toFixed(2)) : null,
      changePercent: quote.dp ? parseFloat(quote.dp.toFixed(2)) : null,
      dividendYield: metrics.currentDividendYieldTTM || null
    }

    console.log('Stock data:', JSON.stringify(stock))

    // Save to Google Sheet
    await saveToSheet(token, stock)

    return { statusCode: 200, headers, body: JSON.stringify({ stock }) }

  } catch(e) {
    console.log('Error:', e.message)
    return { statusCode: 500, headers, body: JSON.stringify({ error: e.message }) }
  }
}
