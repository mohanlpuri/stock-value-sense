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

// ─── Get the real sheetId (gid) of Sheet1 ────────────────────────────────────
// The sheetId in batchUpdate is NOT always 0 — fetch it from the spreadsheet metadata
async function getSheetId(token) {
  const url  = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}?fields=sheets.properties`
  const res  = await fetch(url, { headers: { 'Authorization': 'Bearer ' + token } })
  const data = await res.json()
  // Return the sheetId of the first sheet (Sheet1)
  return data.sheets?.[0]?.properties?.sheetId ?? 0
}

// ─── Read all rows from Google Sheet ──────────────────────────────────────────
async function readSheet(token) {
  const url  = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A:M`
  const res  = await fetch(url, { headers: { 'Authorization': 'Bearer ' + token } })
  const data = await res.json()
  return data.values || []
}

// ─── Color helpers ─────────────────────────────────────────────────────────────
const GREEN  = { red: 0.78, green: 0.93, blue: 0.78 }  // soft green
const YELLOW = { red: 1.00, green: 0.95, blue: 0.70 }  // soft yellow
const NONE   = { red: 1.00, green: 1.00, blue: 1.00 }  // white (clear)

function makeColorRequest(sheetId, rowIndex, colIndex, color) {
  return {
    repeatCell: {
      range: {
        sheetId:          sheetId,
        startRowIndex:    rowIndex,      // 0-based
        endRowIndex:      rowIndex + 1,
        startColumnIndex: colIndex,      // 0-based: A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12
        endColumnIndex:   colIndex + 1
      },
      cell: {
        userEnteredFormat: {
          backgroundColor: color
        }
      },
      fields: 'userEnteredFormat.backgroundColor'
    }
  }
}

// ─── Apply cell highlighting via batchUpdate ───────────────────────────────────
async function applyHighlighting(token, sheetId, rowIndex, stock) {
  const requests = []

  // F (col 5) — P/E ratio
  if (stock.peRatio != null && stock.peRatio !== '') {
    const pe    = parseFloat(stock.peRatio)
    if (!isNaN(pe)) {
      const color = pe < 20 ? GREEN : pe > 40 ? YELLOW : NONE
      requests.push(makeColorRequest(sheetId, rowIndex, 5, color))
    }
  }

  // H (col 7) — P/B ratio
  if (stock.pbRatio != null && stock.pbRatio !== '') {
    const pb    = parseFloat(stock.pbRatio)
    if (!isNaN(pb)) {
      const color = pb < 3 ? GREEN : pb > 10 ? YELLOW : NONE
      requests.push(makeColorRequest(sheetId, rowIndex, 7, color))
    }
  }

  // K (col 10) — Tip Ranks
  if (stock.tipRanks != null && stock.tipRanks !== '') {
    const tr    = parseFloat(stock.tipRanks)
    if (!isNaN(tr)) {
      const color = tr >= 8 ? GREEN : NONE
      requests.push(makeColorRequest(sheetId, rowIndex, 10, color))
    }
  }

  // L (col 11) — Morningstar
  if (stock.morningstar != null && stock.morningstar !== '') {
    const ms    = parseInt(stock.morningstar)
    if (!isNaN(ms)) {
      const color = ms >= 4 ? GREEN : NONE
      requests.push(makeColorRequest(sheetId, rowIndex, 11, color))
    }
  }

  if (requests.length === 0) {
    console.log('No highlight requests to apply')
    return
  }

  console.log(`Applying ${requests.length} highlights to sheetId=${sheetId} rowIndex=${rowIndex}`)

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}:batchUpdate`
  const res  = await fetch(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests })
  })
  const result = await res.json()
  console.log('batchUpdate result:', JSON.stringify(result).slice(0, 300))
}

// ─── Save / update row in Google Sheet ────────────────────────────────────────
// Columns: A=Date B=Ticker C=Price D=52WH E=52WL F=PE G=BV H=PB I=Analyst J=DivYield K=TipRanks L=Morningstar M=AISummary
async function saveToSheet(token, sheetId, stock) {
  const date = new Date().toLocaleDateString('en-US')
  const row  = [
    date,
    stock.ticker,
    stock.price         != null ? stock.price.toFixed(2)     : '',
    stock.week52High    != null ? stock.week52High.toFixed(2) : '',
    stock.week52Low     != null ? stock.week52Low.toFixed(2)  : '',
    stock.peRatio       != null ? stock.peRatio.toFixed(2)    : '',
    stock.bookValue     != null ? stock.bookValue.toFixed(2)  : '',
    stock.pbRatio       != null ? stock.pbRatio.toFixed(2)    : '',
    stock.analystRating != null ? stock.analystRating         : '',
    stock.dividendYield != null ? stock.dividendYield.toFixed(2) + '%' : '',
    '',   // Tip Ranks — filled in via saveRatings
    '',   // Morningstar — filled in via saveRatings
    ''    // AI Summary — filled in via saveAISummary
  ]

  const existing    = await readSheet(token)
  const existingIdx = existing.findIndex((r, i) => i > 0 && r[1] === stock.ticker)

  let rowIndex  // 0-based row index for highlighting

  if (existingIdx > 0) {
    // Row exists — update it (preserve existing Tip Ranks, Morningstar, AI Summary)
    const current = existing[existingIdx]
    row[10] = current[10] || ''
    row[11] = current[11] || ''
    row[12] = current[12] || ''

    rowIndex     = existingIdx   // 0-based (row 0 = header, so existingIdx is correct)
    const rowNum = existingIdx + 1
    const url    = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A${rowNum}:M${rowNum}?valueInputOption=RAW`
    await fetch(url, {
      method: 'PUT',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [row] })
    })
    console.log(`Updated row ${rowNum} (0-based index: ${rowIndex}) for ${stock.ticker}`)

  } else {
    // New row — append
    // rowIndex = number of existing rows (header + data rows)
    rowIndex  = existing.length
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1!A:M:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`
    await fetch(url, {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: [row] })
    })
    console.log(`Appended new row at 0-based index: ${rowIndex} for ${stock.ticker}`)
  }

  // Apply highlighting for P/E and P/B
  await applyHighlighting(token, sheetId, rowIndex, stock)
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
    week52High:    metrics['52WeekHigh']              || null,
    week52Low:     metrics['52WeekLow']               || null,
    peRatio:       metrics['peBasicExclExtraTTM']     || metrics['peAnnual'] || null,
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
- Dividend Yield: ${stock.dividendYield != null ? stock.dividendYield.toFixed(2) + '%' : 'N/A'}

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
    const body    = JSON.parse(event.body)
    const token   = await getAccessToken()
    const sheetId = await getSheetId(token)   // fetch real sheetId every time

    console.log('Real sheetId:', sheetId)

    // ── History read ──────────────────────────────────────────────────────────
    if (body.action === 'history') {
      const rows = await readSheet(token)
      return { statusCode: 200, headers, body: JSON.stringify({ rows }) }
    }

    // ── Save manual ratings (Tip Ranks + Morningstar) + re-highlight ──────────
    if (body.action === 'saveRatings') {
      const ticker      = (body.ticker || '').trim().toUpperCase()
      const existing    = await readSheet(token)
      const existingIdx = existing.findIndex((r, i) => i > 0 && r[1] === ticker)

      if (existingIdx < 1) {
        return { statusCode: 404, headers, body: JSON.stringify({ error: 'Ticker not found in sheet. Evaluate first.' }) }
      }

      const rowNum  = existingIdx + 1
      const current = existing[existingIdx]
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

      // Re-apply all highlights for this row including new Tip Ranks + Morningstar
      await applyHighlighting(token, sheetId, existingIdx, {
        peRatio:    current[5],
        pbRatio:    current[7],
        tipRanks:   current[10],
        morningstar: current[11]
      })

      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) }
    }

    // ── Save AI summary to sheet ──────────────────────────────────────────────
    if (body.action === 'saveAISummary') {
      const ticker      = (body.ticker || '').trim().toUpperCase()
      const existing    = await readSheet(token)
      const existingIdx = existing.findIndex((r, i) => i > 0 && r[1] === ticker)

      if (existingIdx < 1) {
        return { statusCode: 404, headers, body: JSON.stringify({ error: 'Ticker not found in sheet. Evaluate first.' }) }
      }

      const rowNum  = existingIdx + 1
      const current = existing[existingIdx]
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
      const summary = await generateAISummary(body.stock)
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

    await saveToSheet(token, sheetId, stock)
    return { statusCode: 200, headers, body: JSON.stringify({ stock }) }

  } catch(e) {
    console.log('Error:', e.message)
    return { statusCode: 500, headers, body: JSON.stringify({ error: e.message }) }
  }
}
