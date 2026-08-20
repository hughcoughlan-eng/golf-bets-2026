import { google } from 'googleapis';

const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const CREDS = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

function getAuth() {
  return new google.auth.GoogleAuth({
    credentials: CREDS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

// Parse a price value from the sheet — handles €4.00, N/A, empty
function parsePrice(val) {
  if (!val || val.toString().trim() === '' || val.toString().trim().toUpperCase() === 'N/A') return null;
  const n = parseFloat(val.toString().replace(/[€,]/g, ''));
  return isNaN(n) ? null : n;
}

// Generate a simple key from a market name e.g. "Winner - Round 1 - Portsalon" -> "m0"
function makeKey(index) {
  return `m${index}`;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  if (req.method === 'GET') {
    const action = req.query.action;

    // ── Config: markets + players from Inputs tab ──────────────
    if (action === 'config') {
      try {
        // Read markets (E2:F20), players (H2:H30) and competition name (C2) from Inputs
        const [marketsResult, playersResult, nameResult] = await Promise.all([
          sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Inputs!E2:F20' }),
          sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Inputs!H2:H30' }),
          sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Inputs!C2' }),
        ]);

        const marketRows = (marketsResult.data.values || []).filter(r => (r[0] || '').trim());
        const markets = marketRows.map((row, i) => ({
          id: makeKey(i),
          label: (row[0] || '').trim(),
          status: (row[1] || 'open').trim().toLowerCase(),
        }));

        const players = (playersResult.data.values || [])
          .map(r => (r[0] || '').trim())
          .filter(Boolean);

        const competitionName = ((nameResult.data.values || [[]])[0] || [])[0] || 'Golf 2026';

        return res.status(200).json({ markets, players, competitionName });
      } catch (err) {
        console.error(err);
        return res.status(500).json({ error: 'Failed to read config', detail: err.message });
      }
    }

    // ── Prices from Price Matrix ───────────────────────────────
    if (action === 'prices') {
      try {
        // Read markets and statuses first
        const inputsResult = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Inputs!E2:F20',
        });
        const marketRows = (inputsResult.data.values || []).filter(r => (r[0] || '').trim());
        const markets = marketRows.map((row, i) => ({
          id: makeKey(i),
          label: (row[0] || '').trim(),
          status: (row[1] || 'open').trim().toLowerCase(),
        }));
        const marketStatus = {};
        markets.forEach(m => { marketStatus[m.id] = m.status; });

        // Read Price Matrix — row 4 onwards, col A = player, B onwards = market prices
        const priceResult = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Price Matrix!A4:Z50',
        });
        const rows = priceResult.data.values || [];

        // Build empty prices object keyed by market id
        const prices = {};
        markets.forEach(m => { prices[m.id] = {}; });

        // Row 0 is the header row — skip it, data starts at row 1
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          const player = (row[0] || '').trim();
          if (!player) continue;
          markets.forEach((m, idx) => {
            const price = parsePrice(row[idx + 1]);
            if (price !== null) prices[m.id][player] = price;
          });
        }

        return res.status(200).json({ prices, marketStatus, markets });
      } catch (err) {
        console.error(err);
        return res.status(500).json({ error: 'Failed to read prices', detail: err.message });
      }
    }

    // ── Scores from Players and scores ────────────────────────
    if (action === 'scores') {
      try {
        // Read how many markets there are to know how many round columns to expect
        const inputsResult = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Inputs!E2:E20',
        });
        const numMarkets = (inputsResult.data.values || []).filter(r => (r[0] || '').trim()).length;
        // Rounds = numMarkets - 1 (last market is Overall), capped at number of score columns
        const numRounds = Math.max(numMarkets - 1, 1);

        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Players and scores!A2:Z30',
        });
        const rows = result.data.values || [];
        const scores = {};
        for (const row of rows) {
          const player = (row[0] || '').trim();
          if (!player || player.toLowerCase().startsWith('stake') || player.startsWith('€')) continue;
          // Columns B onwards = R1, R2, R3... then Total in last column
          const roundScores = [];
          for (let i = 0; i < numRounds; i++) {
            roundScores.push(row[i + 1] ? parseFloat(row[i + 1]) || null : null);
          }
          // Total is in the column after the rounds
          const total = row[numRounds + 1] ? parseFloat(row[numRounds + 1]) || null : null;
          roundScores.push(total);
          scores[player] = roundScores;
        }
        return res.status(200).json({ scores, numRounds });
      } catch (err) {
        console.error(err);
        return res.status(500).json({ error: 'Failed to read scores', detail: err.message });
      }
    }

    // ── Market status for admin page ─────────────────────────
    if (action === 'marketStatus') {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Inputs!E2:F20',
        });
        const rows = (result.data.values || []).filter(r => (r[0] || '').trim());
        const markets = rows.map(row => ({
          name: (row[0] || '').trim(),
          status: (row[1] || 'open').trim().toLowerCase(),
        }));
        return res.status(200).json({ markets });
      } catch (err) {
        console.error(err);
        return res.status(500).json({ error: 'Failed to read market status', detail: err.message });
      }
    }

    return res.status(400).json({ error: 'Unknown action' });
  }

  if (req.method === 'POST') {
    try {
      const { action } = req.query;

      // ── Update market status ────────────────────────────────
      if (action === 'marketStatus') {
        const { marketName, status } = req.body;
        if (!marketName || !status) return res.status(400).json({ error: 'Missing fields' });

        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Inputs!E2:F20',
        });
        const rows = result.data.values || [];
        const rowIndex = rows.findIndex(r => (r[0] || '').trim() === marketName);
        if (rowIndex === -1) return res.status(404).json({ error: 'Market not found' });

        const sheetRow = rowIndex + 2;
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `Inputs!F${sheetRow}`,
          valueInputOption: 'RAW',
          requestBody: { values: [[status]] },
        });
        return res.status(200).json({ success: true });
      }

      // ── Log a bet ───────────────────────────────────────────
      const { bettor, player, market, stake, price } = req.body;
      if (!bettor || !player || !market || !stake || !price) {
        return res.status(400).json({ error: 'Missing fields' });
      }

      const now = new Date();
      const time = now.toTimeString().slice(0, 5);
      const date = `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()}`;

      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: 'Bet_Logs!A:G',
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[time, date, bettor, market, player, stake, price]],
        },
      });

      return res.status(200).json({ success: true });
    } catch (err) {
      console.error(err);
      return res.status(500).json({ error: 'Failed', detail: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
