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

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // GET /api/sheets?action=prices  — read Price Matrix
  // GET /api/sheets?action=scores  — read Players and scores
  // POST /api/sheets               — write a bet to Bet_Logs

  if (req.method === 'GET') {
    const action = req.query.action;

    if (action === 'prices') {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Price Matrix!A4:F17',
        });
        const rows = result.data.values || [];

        // Row 0 is headers: [Player/Market, Round1, Round2, Round3, Round4, Overall]
        const marketKeys = ['r1', 'r2', 'r3', 'r4', 'overall'];
        const prices = { r1: {}, r2: {}, r3: {}, r4: {}, overall: {} };

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          const player = (row[0] || '').trim();
          if (!player) continue;
          marketKeys.forEach((key, idx) => {
            const price = parsePrice(row[idx + 1]);
            if (price !== null) prices[key][player] = price;
          });
        }

        return res.status(200).json({ prices });
      } catch (err) {
        console.error(err);
        return res.status(500).json({ error: 'Failed to read prices', detail: err.message });
      }
    }

    if (action === 'scores') {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Players and scores!A2:G30',
        });
        const rows = result.data.values || [];
        const scores = {};
        for (const row of rows) {
          const player = (row[0] || '').trim();
          if (!player) continue;
          // Columns D-G = index 3-6 for R1-R4
          scores[player] = [
            row[3] ? parseFloat(row[3]) || null : null,
            row[8] ? parseFloat(row[8]) || null : null,
            row[4] ? parseFloat(row[4]) || null : null,
            row[5] ? parseFloat(row[5]) || null : null,
            row[6] ? parseFloat(row[6]) || null : null,
          ];
        }
        return res.status(200).json({ scores });
      } catch (err) {
        console.error(err);
        return res.status(500).json({ error: 'Failed to read scores', detail: err.message });
      }
    }

    return res.status(400).json({ error: 'Unknown action' });
  }

  if (req.method === 'POST') {
    try {
      const { bettor, player, market, stake, price } = req.body;
      if (!bettor || !player || !market || !stake || !price) {
        return res.status(400).json({ error: 'Missing fields' });
      }

      const now = new Date();
      const time = now.toTimeString().slice(0, 5); // HH:MM
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
      return res.status(500).json({ error: 'Failed to log bet', detail: err.message });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
