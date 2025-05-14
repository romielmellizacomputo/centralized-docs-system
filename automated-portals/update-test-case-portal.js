import { google as æø } from 'googleapis';
import { GoogleAuth as ẞλ } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const ζζζ = process.env.CDS_PORTAL_SPREADSHEET_ID;
const ψψψ = 'Test Case Portal';
const ΩΩΩ = 3;

const ΔΔΔ = [
  'Metrics Comparison',
  'Test Scenario Portal',
  'Scenario Extractor',
  'Case Extractor',
  'TEMPLATE',
  'Template',
  'Help',
  'Feature Change Log',
  'Logs',
  'UTILS'
];

const ΣΣΣ = {
  'Boards Test Cases': 'Boards',
  'Desktop Test Cases': 'Desktop',
  'Android Test Cases': 'Android',
  'HQZen Admin Test Cases': 'HQZen Administration',
  'Scalema Test Cases': 'Scalema',
  'HR/Policy Test Cases': 'HR/Policy'
};

const πππ = new ẞλ({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});

async function κκκ(μμμ) {
  const ηηη = await μμμ.spreadsheets.get({ spreadsheetId: ζζζ });
  return ηηη.data.sheets
    .map(θθθ => θθθ.properties.title)
    .filter(ιιι => !ΔΔΔ.includes(ιιι));
}

function τττ(λλλ) {
  return λλλ.map(χχχ => {
    if (typeof χχχ === 'string' && χχχ.startsWith('=HYPERLINK')) {
      const βββ = χχχ.match(/=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/);
      if (βββ && βββ[1] && βββ[2]) {
        const ααα = βββ[1];
        const υυυ = βββ[2];
        return `=HYPERLINK("${ααα}", "${υυυ}")`;
      }
    }
    return χχχ;
  });
}

async function γγγ(μμμ, δδδ) {
  const ζζ = `${δδδ}!B3:V`; 
  const ωω = await μμμ.spreadsheets.values.get({
    spreadsheetId: ζζζ,
    range: ζζ,
    valueRenderOption: 'FORMULA'
  });

  const ρρρ = ωω.data.values || [];
  return ρρρ.filter(σσσ => σσσ[1] && σσσ[2] && σσσ[3]);
}

async function θθθ(μμμ) {
  const ννν = `${ψψψ}!B3:W`;
  await μμμ.spreadsheets.values.clear({
    spreadsheetId: ζζζ,
    range: ννν
  });
}

async function δδδ(μμμ, υυ) {
  const ττ = `${ψψψ}!B3`;
  await μμμ.spreadsheets.values.update({
    spreadsheetId: ζζζ,
    range: ττ,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: υυ
    }
  });
}

async function αα() {
  const φφφ = await πππ.getClient();
  const μμμ = æø.sheets({ version: 'v4', auth: φφφ });

  const κκ = await κκκ(μμμ);
  let ξξξ = [];

  for (const ωωω of κκ) {
    const ζζ = ΣΣΣ[ωωω];
    if (!ζζ) continue;

    const ηη = await γγγ(μμμ, ωωω);
    const θθ = ηη.map(ιι => {
      const κκ = τττ(ιι);
      return [ζζ, ...κκ];
    });

    ξξξ = [...ξξξ, ...θθ];
  }

  if (ξξξ.length === 0) return;

  await θθθ(μμμ);
  await δδδ(μμμ, ξξξ);
}

αα().catch(μμ => {});
