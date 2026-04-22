const SHEET_ID = process.env.GOOGLE_SHEET_ID_TRACKER
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL
const PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n')

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

// Fetch price data from Twelve Data
async function fetchPriceData(ticker) {
  const apiKey = process.env.TWELVE_DATA_API_KEY
  const url = `https://api.twelvedata.com/quote?symbol=${ticker}&apikey=${apiKey}`
  const res = await fetch(url)
  const q = await res.json()
  console.log('Twelve Data status:', res.status, q.code || 'ok')
  if (!q || q.code || !q.symbol) return null
  return {
    ticker: q.symbol,
    name: q.name || q.symbol,
    price: q.close ? parseFloat(parseFloat(q.close).toFixed(2)) : null,
    week52High: q.fifty_two_week ? parseFloat(parseFloat(q.fifty_two_week.high).toFixed(2)) : null,
    week52Low: q.fifty_two_week ? parseFloat(parseFloat(q.fifty_two_week.low).toFixed(2)) : null,
    volume: q.volume ? parseInt(q.volume) : null,
    avgVolume: q.average_volume ? parseInt(q.average_volume) : null,
    exchange: q.exchange || null,
    change: q.change ? parseFloat(parseFloat(q.change).toFixed(2)) : null,
    changePercent: q.percent_change ? parseFloat(parseFloat(q.percent_change).toFixed(2)) : null
  }
}

// Fetch fundamentals from FMP
async function fetchFundamentals(ticker) {
  const apiKey = process.env.FMP_API_KEY

  // Fetch ratios TTM
  const ratiosUrl = `https://financialmodelingprep.com/stable/ratios-ttm?symbol=${ticker}&apikey=${apiKey}`
  const ratiosRes = await fetch(ratiosUrl)
  const ratiosData = await ratiosRes.json()
  console.log('FMP ratios status:', ratiosRes.status)

  // Fetch profile for sector
  const profileUrl = `https://financialmodelingprep.com/stable/profile?symbol=${ticker}&apikey=${apiKey}`
  const profileRes = await fetch(profileUrl)
  const profileData = await profileRes.json()
  console.log('FMP profile status:', profileRes.status)

  const ratios = ratiosData && ratiosData[0] ? ratiosData[0] : null
  const profile = profileData && profileData[0] ? profileData[0] : null

  return {
    peRatio: ratios ? parseFloat(parseFloat(ratios.priceToEarningsRatioTTM).toFixed(2)) : null,
    pbRatio: ratios ? parseFloat(parseFloat(ratios.priceToBookRatioTTM).toFixed(2)) : null,
    bookValue: ratios ? parseFloat(parseFloat(ratios.bookValuePerShareTTM).toFixed(2)) : null,
    dividendYield: ratios ? parseFloat(parseFloat(ratios.dividendYieldTTM * 100).toFixed(2)) : null,
    sector: profile ? profile.sector : null,
    industry: profile ? profile.industry : null,
    marketCap: profile ? profile.marketCap : null
  }
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
    stock.dividendYield ? stock.dividendYield + '%' : ''
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

    // Fetch both data sources in parallel
    const [priceData, fundamentals] = await Promise.all([
      fetchPriceData(ticker),
      fetchFundamentals(ticker)
    ])

    if (!priceData) {
      return { statusCode: 404, headers, body: JSON.stringify({ error: 'Stock not found: ' + ticker }) }
    }

    // Merge data
    const stock = {
      ...priceData,
      peRatio: fundamentals.peRatio,
      pbRatio: fundamentals.pbRatio,
      bookValue: fundamentals.bookValue,
      dividendYield: fundamentals.dividendYield,
      sector: fundamentals.sector,
      industry: fundamentals.industry,
      marketCap: fundamentals.marketCap,
      sectorPE: null,
      analystRating: null
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
