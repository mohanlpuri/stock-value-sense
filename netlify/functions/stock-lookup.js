const SHEET_ID     = process.env.GOOGLE_SHEET_ID_TRACKER
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL
const PRIVATE_KEY  = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n')
const FINNHUB_KEY  = process.env.FINNHUB_API_KEY

// ─── Google JWT auth ───────────────────────────────────────────────────────────
async function getAccessToken() {
  const header = { alg: 'RS256', typ: 'JWT' }
  const now    = Math.floor(Date.now() / 1000)
  const claim  = {
    iss: CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  }
  const encode   = obj => Buffer.from(JSON.stringify(obj)).toString('base64url')
  const unsigned = encode(header) + '.' + encode(claim)
  const crypto   = require('crypto')
  const sign     = crypto.createSign('RSA-SHA256')
  sign.update(unsigned)
  const jwt = unsigned + '.' + sign.sign(PRIVATE_KEY, 'base64url')

  const res  = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + jwt
  })
  const data = await res.json()
  return data.access_token
}

// ─── Read all rows from Google Sheet ──────────────────────────────────────────
async function readSheet(token) {
  const url  = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A:M`
  const res  = await fetch(url, { headers: { 'Authorization': 'Bearer ' + token } })
  const data = await res.json()
  return data.values || []
}

// ─── Save / update row in Google Sheet ────────────────────────────────────────
// Columns: A=Date B=Ticker C=Price D=52WH E=52WL F=PE G=BV H=PB I=Analyst J=DivYield K=TipRanks L=Morningstar M=AI Summary
async function saveToSheet(token, stock) {
  const date = new Date().toLocaleDateString('en-US')
  const row  = [
    date,
    stock.ticker,
    stock.price         != null ? stock.price.toFixed(2)      : '',
    stock.week52High    != null ? stock.week52High.toFixed(2)  : '',
    stock.week52Low     != null ? stock.week52Low.toFixed(2)   : '',
    stock.peRatio       != null ? stock.peRatio.toFixed(2)     : '',
    stock.bookValue     != null ? stock.bookValue.toFixed(2)   : '',
    stock.pbRatio       != null ? stock.pbRatio.toFixed(2)     : '',
    stock.analystRating != null ? stock.analystRating          : '',
    stock.dividendYield != null ? (stock.dividendYield * 100).toFixed(2) + '%' : '',
    stock.tipRanks      != null ? stock.tipRanks               : '',
    stock.morningstar   != null ? stock.morningstar            : '',
    stock.aiSummary     != null ? stock.aiSummary              : ''
  ]

  const existing    = await readSheet(token)
  const existingRow = existing.findIndex((r, i) => i > 0 && r[1] === stock.ticker)

  let url, method
  if (existingRow > 0) {
    const rowNum = existingRow + 1
    url    = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A${rowNum}:M${rowNum}?valueInputOption=RAW`
    method = 'PUT'
  } else {
    url    = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A:M:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`
    method = 'POST'
  }

  await fetch(url, {
    method,
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [row] })
  })
}

// ─── Fetch stock data from Finnhub ────────────────────────────────────────────
async function fetchStock(ticker) {
  const base = 'https://finnhub.io/api/v1'

  const [quoteRes, metricsRes, profileRes, recRes] = await Promise.all([
    fetch(`${base}/quote?symbol=${ticker}&token=${FINNHUB_KEY}`),
    fetch(`${base}/stock/metric?symbol=${ticker}&metric=all&token=${FINNHUB_KEY}`),
    fetch(`${base}/stock/profile2?symbol=${ticker}&token=${FINNHUB_KEY}`),
    fetch(`${base}/stock/recommendation?symbol=${ticker}&token=${FINNHUB_KEY}`)
  ])

  const [quote, metricsData, profile, recData] = await Promise.all([
    quoteRes.json(),
    metricsRes.json(),
    profileRes.json(),
    recRes.json()
  ])

  const metrics = metricsData.metric || {}
  const price   = quote.c || quote.pc || 0
  if (!price || price <= 0) return null

  const bookValue = metrics['bookValuePerShareAnnual'] || metrics['bookValuePerShareQuarterly'] || null
  const pbRatio   = (price && bookValue && bookValue > 0) ? price / bookValue : null

  let analystRating = null
  if (Array.isArray(recData) && recData.length > 0) {
    const latest     = recData[0]
    const totalCount = (latest.strongBuy || 0) + (latest.buy || 0) + (latest.hold || 0) + (latest.sell || 0) + (latest.strongSell || 0)
    const buyScore   = ((latest.strongBuy || 0) * 2 + (latest.buy || 0)) / Math.max(totalCount, 1)
    let label = '3 - Hold'
    if (buyScore >= 1.2)      label = '1 - Strong Buy'
    else if (buyScore >= 0.8) label = '2 - Buy'
    else if (buyScore < 0.4)  label = '4 - Sell'
    analystRating = `${label} (${totalCount})`
  }

  return {
    ticker:        ticker,
    name:          profile.name || ticker,
    sector:        profile.finnhubIndustry || '',
    price:         price,
    week52High:    metrics['52WeekHigh']            || null,
    week52Low:     metrics['52WeekLow']             || null,
    peRatio:       metrics['peBasicExclExtraTTM']   || metrics['peAnnual'] || null,
    bookValue:     bookValue,
    pbRatio:       pbRatio,
    analystRating: analystRating,
    dividendYield: metrics['currentDividendYieldTTM'] || null
  }
}

// ─── Generate AI summary via Anthropic API ────────────────────────────────────
async function generateAISummary(stock) {
  const prompt = `You are a concise stock analyst. Given the following data for ${stock.ticker} (${stock.name}), write a 3-4 sentence summary covering: what the company does, whether the valuation looks attractive or expensive, and any notable positives or concerns. Be factual and balanced. Do not give buy/sell advice.

Stock data:
- Ticker: ${stock.ticker}
- Company: ${stock.name}
- Sector: ${stock.sector || 'Unknown'}
- Price: $${stock.price != null ? stock.price.toFixed(2) : 'N/A'}
- 52W High: $${stock.week52High != null ? stock.week52High.toFixed(2) : 'N/A'}
- 52W Low: $${stock.week52Low != null ? stock.week52Low.toFixed(2) : 'N/A'}
- P/E Ratio: ${stock.peRatio != null ? stock.peRatio.toFixed(1) : 'N/A'}
- Book Value/Share: $${stock.bookValue != null ? stock.bookValue.toFixed(2) : 'N/A'}
- P/B Ratio: ${stock.pbRatio != null ? stock.pbRatio.toFixed(2) : 'N/A'}
- Analyst Rating: ${stock.analystRating || 'N/A'}
- Dividend Yield: ${stock.dividendYield != null ? (stock.dividendYield * 100).toFixed(2) + '%' : 'N/A'}

Write only the summary paragraph, no headings or bullet points.`

  const res  = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type':      'application/json',
      'x-api-key':         process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 300,
      messages:   [{ role: 'user', content: prompt }]
    })
  })

  const data = await res.json()
  return data.content?.[0]?.text || 'Summary unavailable.'
}

// ─── Handler ──────────────────────────────────────────────────────────────────
exports.handler = async function(event) {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' }
  }

  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  }

  try {
    const body  = JSON.parse(event.body)
    const token = await getAccessToken()

    // ── History read ──────────────────────────────────────────────────────────
    if (body.action === 'history') {
      const rows = await readSheet(token)
      return { statusCode: 200, headers, body: JSON.stringify({ rows }) }
    }

    // ── Save manual ratings (Tip Ranks + Morningstar) ─────────────────────────
    if (body.action === 'saveRatings') {
      const ticker      = (body.ticker || '').trim().toUpperCase()
      const existing    = await readSheet(token)
      const existingRow = existing.findIndex((r, i) => i > 0 && r[1] === ticker)

      if (existingRow < 1) {
        return { statusCode: 404, headers, body: JSON.stringify({ error: 'Ticker not found in sheet. Evaluate first.' }) }
      }

      const rowNum  = existingRow + 1
      const current = existing[existingRow]
      while (current.length < 13) current.push('')
      current[10] = body.tipRanks    != null ? String(body.tipRanks)    : (current[10] || '')
      current[11] = body.morningstar != null ? String(body.morningstar) : (current[11] || '')

      await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A${rowNum}:M${rowNum}?valueInputOption=RAW`,
        {
          method: 'PUT',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ values: [current] })
        }
      )
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) }
    }

    // ── Save AI summary to sheet ──────────────────────────────────────────────
    if (body.action === 'saveAISummary') {
      const ticker      = (body.ticker || '').trim().toUpperCase()
      const existing    = await readSheet(token)
      const existingRow = existing.findIndex((r, i) => i > 0 && r[1] === ticker)

      if (existingRow < 1) {
        return { statusCode: 404, headers, body: JSON.stringify({ error: 'Ticker not found in sheet. Evaluate first.' }) }
      }

      const rowNum  = existingRow + 1
      const current = existing[existingRow]
      while (current.length < 13) current.push('')
      current[12] = body.aiSummary || ''

      await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A${rowNum}:M${rowNum}?valueInputOption=RAW`,
        {
          method: 'PUT',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ values: [current] })
        }
      )
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) }
    }

    // ── AI summary generation ─────────────────────────────────────────────────
    if (body.action === 'aiSummary') {
      const stock   = body.stock
      const summary = await generateAISummary(stock)
      return { statusCode: 200, headers, body: JSON.stringify({ summary }) }
    }

    // ── Stock lookup ──────────────────────────────────────────────────────────
    const ticker = (body.ticker || '').trim().toUpperCase()
    if (!ticker) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: 'Ticker is required' }) }
    }

    const stock = await fetchStock(ticker)
    if (!stock) {
      return { statusCode: 404, headers, body: JSON.stringify({ error: 'Stock not found: ' + ticker }) }
    }

    await saveToSheet(token, stock)
    return { statusCode: 200, headers, body: JSON.stringify({ stock }) }

  } catch(e) {
    console.log('Error:', e.message)
    return { statusCode: 500, headers, body: JSON.stringify({ error: e.message }) }
  }
}
