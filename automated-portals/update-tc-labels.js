import dotenv from 'dotenv';
import { google } from 'googleapis';
import axios from 'axios';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const _fxxl_ = fileURLToPath(import.meta.url);
const _dxzl_ = path.dirname(_fxxl_);

dotenv.config();

const _zreq_ = [
  'GITLAB_URL',
  'GITLAB_TOKEN',
  'CDS_PORTAL_SPREADSHEET_ID',
  'CDS_PORTALS_SERVICE_ACCOUNT_JSON',
];
_zreq_.forEach((k) => {
  if (!process.env[k]) {
    process.exit(1);
  }
});

const _gurl_ = process.env.GITLAB_URL;
const _gtok_ = process.env.GITLAB_TOKEN;
const _ssid_ = process.env.CDS_PORTAL_SPREADSHEET_ID;
const _svcjson_ = JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON);

const _labx_ = [
  'To Do',
  'Doing',
  'Changes Requested',
  'Manual QA For Review',
  'QA Lead For Review',
  'Automation QA For Review',
  'Done',
  'On Hold',
  'Deprecated',
  'Automation Team For Review',
];

const _pidmap_ = {
  155: 'HQZen',
  23: 'Backend',
  124: 'Android',
  123: 'Desktop',
  88: 'ApplyBPO',
  141: 'Ministry',
  147: 'Scalema',
  89: 'BPOSeats.com',
};

const _exsheets_ = [
  'Metrics Comparison',
  'Test Scenario Portal',
  'Test Case Portal',
  'Scenario Extractor',
  'TEMPLATE',
  'Template',
  'Help',
  'Feature Change Log',
  'Logs',
  'UTILS',
];

async function _gcreds_() {
  const _auth_ = new google.auth.GoogleAuth({
    credentials: _svcjson_,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return _auth_.getClient();
}

function _pxid_(u) {
  const _u_ = u.split('/');
  const _pname_ = _u_[4];

  for (const [i, n] of Object.entries(_pidmap_)) {
    if (_pname_.toLowerCase().includes(n.toLowerCase())) {
      return { id: i, name: n };
    }
  }
  return null;
}

async function _fissue_(pid, iid) {
  const _url_ = `${_gurl_}api/v4/projects/${pid}/issues/${iid}`;
  const _res_ = await axios.get(_url_, {
    headers: {
      'PRIVATE-TOKEN': _gtok_,
    },
  });
  return _res_.data;
}

async function _fnotes_(pid, iid) {
  const _url_ = `${_gurl_}api/v4/projects/${pid}/issues/${iid}/notes`;
  const _res_ = await axios.get(_url_, {
    headers: {
      'PRIVATE-TOKEN': _gtok_,
    },
  });
  return _res_.data;
}

async function _revmet_() {
  const _authc_ = await _gcreds_();
  const _sheets_ = google.sheets({ version: 'v4', auth: _authc_ });

  const _ss_ = await _sheets_.spreadsheets.get({
    spreadsheetId: _ssid_,
  });
  const _titles_ = _ss_.data.sheets
    .map((s) => s.properties.title)
    .filter((t) => !_exsheets_.includes(t));

  for (const _t_ of _titles_) {
    const _rng_ = `'${_t_}'!E3:E`;
    const _res_ = await _sheets_.spreadsheets.values.get({
      spreadsheetId: _ssid_,
      range: _rng_,
    });

    const _rows_ = _res_.data.values || [];
    const _updates_ = [];

    for (let i = 0; i < _rows_.length; i++) {
      const _ridx_ = i + 3;
      const _url_ = _rows_[i][0] || '';

      if (!_url_ || !/^https:\/\/forge\.bposeats\.com\/[^\/]+\/[^\/]+\/-\/issues\/\d+$/.test(_url_)) {
        continue;
      }

      const _iid_ = _url_.split('/').pop();
      const _proj_ = _pxid_(_url_);

      if (!_proj_) {
        continue;
      }

      try {
        const _iss_ = await _fissue_(_proj_.id, _iid_);
        const _lab_ = _labx_.find((l) => _iss_.labels.includes(l));

        if (_lab_) {
          const _notes_ = await _fnotes_(_proj_.id, _iid_);
          const _last_ = _notes_.length > 0 ? _notes_[_notes_.length - 1] : null;

          let _note_ = '';
          if (_iss_.state === 'opened') {
            _note_ += `The ticket was created by ${_iss_.author?.name || 'Unknown'}\n`;
            _note_ += `The ticket is still open\n`;
            _note_ += _iss_.assignee ? `${_iss_.assignee.name} is the current assignee\n` : `No assignee on the ticket\n`;
          } else if (_iss_.state === 'closed') {
            _note_ += `The ticket was created by ${_iss_.author?.name || 'Unknown'}\n`;
            _note_ += `The ticket was closed\n`;
            _note_ += _iss_.assignee ? `${_iss_.assignee.name} was the assignee\n` : `No assignee on the ticket\n`;
          }

          if (_last_) {
            _note_ += `Last activity: ${_last_.body}\n`;
            _note_ += `Commented by: ${_last_.author.name}\n`;
          }

          _updates_.push({
            range: `'${_t_}'!I${_ridx_}`,
            values: [[_lab_]],
          });

          _updates_.push({
            range: `'${_t_}'!E${_ridx_}`,
            values: [[_url_]],
            note: _note_,
          });
        }
      } catch {}
    }

    if (_updates_.length > 0) {
      const _vals_ = _updates_.map(({ range, values }) => ({ range, values }));
      await _sheets_.spreadsheets.values.batchUpdate({
        spreadsheetId: _ssid_,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: _vals_,
        },
      });
    }
  }
}

_revmet_().catch(() => {});
