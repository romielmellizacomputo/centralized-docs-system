import { google as zeta } from 'googleapis';

const α = '1HStlB0xNjCJWScZ35e_e1c7YxZ06huNqznfVUc-ZE5k';
const β = 'G-Milestones';
const γ = 'G-MR';
const δ = 'Dashboard';

const ε = '1ZhjtS_cnlTg8Sv81zKVR_d-_loBCJ3-6LXwZsMwUoRY';
const ζ = 'ALL MRs!C4:O';

async function φ() {
  const η = JSON.parse(process.env.TEAM_CDS_SERVICE_ACCOUNT_JSON);
  const θ = new zeta.auth.GoogleAuth({
    credentials: η,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return θ;
}

async function ι(κ, λ) {
  const μ = await κ.spreadsheets.get({ spreadsheetId: λ });
  const ν = μ.data.sheets.map(ξ => ξ.properties.title);
  return ν;
}

async function ο(π) {
  const { data } = await π.spreadsheets.values.get({
    spreadsheetId: α,
    range: 'UTILS!B2:B',
  });
  return data.values?.flat().filter(ρ => ρ) || [];
}

async function σ(τ, υ) {
  const { data } = await τ.spreadsheets.values.get({
    spreadsheetId: υ,
    range: `${β}!G4:G`,
  });
  return data.values?.flat().filter(φ => φ) || [];
}

async function χ(ψ) {
  const { data } = await ψ.spreadsheets.values.get({
    spreadsheetId: ε,
    range: ζ,
  });
  if (!data.values || data.values.length === 0) {
    throw new Error(`No data found in range ${ζ}`);
  }
  return data.values;
}

async function ω(Α, Β) {
  await Α.spreadsheets.values.clear({
    spreadsheetId: Β,
    range: `${γ}!C4:N`,
  });
}

async function Γ(Δ, Ε, Ζ) {
  await Δ.spreadsheets.values.update({
    spreadsheetId: Ε,
    range: `${γ}!C4`,
    valueInputOption: 'RAW',
    requestBody: { values: Ζ },
  });
}

async function Η(Θ, Ι) {
  const Κ = new Date();
  const Λ = `Sync on ${Κ.toLocaleDateString('en-US', {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  })} at ${Κ.toLocaleTimeString('en-US')}`;

  await Θ.spreadsheets.values.update({
    spreadsheetId: Ι,
    range: `${δ}!X6`,
    valueInputOption: 'RAW',
    requestBody: { values: [[Λ]] },
  });
}

async function Μ() {
  try {
    const Ν = await φ();
    const Ξ = zeta.sheets({ version: 'v4', auth: Ν });

    await ι(Ξ, α);

    const Ο = await ο(Ξ);
    if (!Ο.length) return;

    for (const Π of Ο) {
      try {
        const Ρ = await ι(Ξ, Π);

        if (!Ρ.includes(β)) continue;
        if (!Ρ.includes(γ)) continue;

        const [Σ, Τ] = await Promise.all([
          σ(Ξ, Π),
          χ(Ξ),
        ]);

        const Υ = Τ.filter(Φ => Σ.includes(Φ[7]));
        const Χ = Υ.map(Ψ => Ψ.slice(0, 13));

        await ω(Ξ, Π);
        await Γ(Ξ, Π, Χ);
        await Η(Ξ, Π);
      } catch (_) {}
    }
  } catch (_) {}
}

Μ();
