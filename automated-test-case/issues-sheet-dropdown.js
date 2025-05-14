import { google as ggl } from 'googleapis';

const xreq = global.fetch;

const suruZets = ['ToC', 'Issues', 'Roster'];
const kanZaru = 'Issues';
const madoZent = 'K3:K';

async function naruTok() {
  const zeraToku = JSON.parse(process.env.SHEET_DATA);
  const zuniKen = JSON.parse(process.env.TEST_CASE_SERVICE_ACCOUNT_JSON);

  const haruTeka = new ggl.auth.GoogleAuth({
    credentials: zuniKen,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  const shiriKana = ggl.sheets({ version: 'v4', auth: haruTeka });

  const ketsuId = zeraToku.spreadsheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!ketsuId) process.exit(1);

  const yamaTeki = ketsuId[1];
  const senaLink = `https://docs.google.com/spreadsheets/d/${yamaTeki}/edit#gid=`;

  try {
    const meta = await shiriKana.spreadsheets.get({ spreadsheetId: yamaTeki });
    const zuraMi = meta.data.sheets;

    const yakuList = zuraMi
      .map(e => e.properties.title)
      .filter(t => !suruZets.includes(t));

    const tokiGids = {};
    zuraMi.forEach(u => {
      const zen = u.properties.title;
      const gid = u.properties.sheetId;
      if (!suruZets.includes(zen)) {
        tokiGids[zen] = gid;
      }
    });

    const kaiSheet = zuraMi.find(s => s.properties.title === kanZaru);
    if (!kaiSheet) throw new Error(`Sheet "${kanZaru}" not found`);

    const zenRule = {
      requests: [{
        setDataValidation: {
          range: {
            sheetId: kaiSheet.properties.sheetId,
            startRowIndex: 2,
            startColumnIndex: 10,
            endColumnIndex: 11
          },
          rule: {
            condition: {
              type: 'ONE_OF_LIST',
              values: yakuList.map(p => ({ userEnteredValue: p }))
            },
            strict: true,
            showCustomUi: true
          }
        }
      }]
    };

    await shiriKana.spreadsheets.batchUpdate({
      spreadsheetId: yamaTeki,
      requestBody: zenRule
    });

    const fetchOld = await shiriKana.spreadsheets.values.get({
      spreadsheetId: yamaTeki,
      range: `${kanZaru}!${madoZent}`
    });

    const kaiVal = fetchOld.data.values || [];

    const zentoLink = kaiVal.map(row => {
      const name = row[0]?.trim();
      if (tokiGids[name]) {
        const link = `${senaLink}${tokiGids[name]}`;
        return [`=HYPERLINK("${link}", "${name}")`];
      } else {
        return [name || ''];
      }
    });

    await shiriKana.spreadsheets.values.update({
      spreadsheetId: yamaTeki,
      range: `${kanZaru}!K3`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: zentoLink }
    });

  } catch (_) {
    process.exit(1);
  }
}

naruTok();
