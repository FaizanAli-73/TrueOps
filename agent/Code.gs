/**
 * WhatsApp → Google Sheets bot for the Business Data Workbook.
 *
 * Text (or send a receipt photo to) the business WhatsApp number in plain English.
 * Claude reads the message, decides which tabs and columns to fill, this script
 * writes ONLY the blue input columns, logs every cell it touched on the
 * "WhatsApp Log" tab, and replies with exactly what it wrote. Reply UNDO to reverse.
 *
 * Setup: see SETUP.md in the same folder. Script Properties used:
 *   ANTHROPIC_API_KEY          (required)
 *   WHATSAPP_TOKEN             (required)  permanent System User token from Meta
 *   WHATSAPP_PHONE_NUMBER_ID   (required)  from WhatsApp > API Setup in the Meta app
 *   RELAY_KEY                  (required)  long random string, same value as in the Cloudflare Worker
 *   CLAUDE_MODEL               (optional)  default claude-opus-5
 *   CLAUDE_EFFORT              (optional)  default medium
 *   GRAPH_VERSION              (optional)  default v23.0
 */

var HDR = 4;        // header row on every data tab
var FIRST = 5;      // first data row
var DATA_SHEETS = ['Clients', 'Properties', 'Quotes', 'Projects', 'Change Orders', 'Time Log', 'Expenses',
                   'Invoices', 'Team', 'Vendors', 'Callbacks', 'Market Rates'];
var ID_FORMAT = {   // sheets whose first column is an ID the bot can generate
  'Clients': ['C-', 4], 'Properties': ['P-', 4], 'Quotes': ['Q-', 4], 'Projects': ['J-', 4],
  'Change Orders': ['CO-', 4], 'Invoices': ['INV-', 4], 'Team': ['W-', 2], 'Vendors': ['V-', 3], 'Callbacks': ['CB-', 3]
};
var LOG_SHEET = 'WhatsApp Log';

function prop_(k, d) {
  var v = PropertiesService.getScriptProperties().getProperty(k);
  return (v === null || v === '') ? d : v;
}

// =====================================================================================
// Web app entry points (called by the Cloudflare relay, never directly by Meta)
// =====================================================================================
function doGet() {
  return ContentService.createTextOutput('WhatsApp bot is running.');
}

function doPost(e) {
  if (!e || !e.parameter || e.parameter.key !== prop_('RELAY_KEY', '__unset__')) {
    return ContentService.createTextOutput('forbidden');
  }
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return ContentService.createTextOutput('bad json'); }
  var msgs = extractMessages_(body);
  msgs.forEach(function (m) {
    try { handleMessage_(m); } catch (err) { console.error(err && err.stack || err); }
  });
  return ContentService.createTextOutput('ok');
}

function extractMessages_(body) {
  var out = [];
  (body.entry || []).forEach(function (en) {
    (en.changes || []).forEach(function (ch) {
      var v = ch.value || {};
      var names = {};
      (v.contacts || []).forEach(function (c) { names[c.wa_id] = c.profile && c.profile.name; });
      (v.messages || []).forEach(function (m) { m._name = names[m.from] || ''; out.push(m); });
    });
  });
  return out;
}

// =====================================================================================
// Message handling
// =====================================================================================
function handleMessage_(m, opts) {
  opts = opts || {};
  var cache = CacheService.getScriptCache();
  if (m.id && !opts.dryRun) {
    if (cache.get('m:' + m.id)) return null;        // Meta re-delivered it
    cache.put('m:' + m.id, '1', 21600);
  }
  var ss = opts.ss || SpreadsheetApp.getActiveSpreadsheet();
  var text = messageText_(m);
  var worker = opts.worker || findWorker_(ss, m.from);
  if (!worker) {
    var msg = 'This number isn\'t on the Team tab, so I can\'t write to the workbook. Add it to the Phone column on the Team tab (with country code) and try again.';
    if (!opts.dryRun) { sendWhatsApp_(m.from, msg); logRow_(ss, m, '', text, '', '', 'Rejected', msg); }
    return { reply: msg };
  }

  var cmd = text.trim().toUpperCase();
  if (cmd === 'HELP' || cmd === '?') return finish_(ss, m, worker, text, helpText_(), '', 'Help', opts);
  if (cmd === 'STATS' || cmd === 'SUMMARY') return finish_(ss, m, worker, text, statsText_(ss), '', 'Stats', opts);
  if (m.type === 'audio') return finish_(ss, m, worker, text, 'Voice notes aren\'t supported yet — please type it out, or send a photo.', '', 'Skipped', opts);

  var lock = LockService.getScriptLock(), locked = false;
  try {
    if (!opts.dryRun) { lock.waitLock(30000); locked = true; }
    if (cmd === 'UNDO') return finish_(ss, m, worker, text, undoLast_(ss, m.from), '', 'Undo', opts);
    var schema = readSchema_(ss);
    var media = opts.media || (m.type === 'image' || m.type === 'document' ? fetchMedia_(m) : null);
    var system = buildSystemPrompt_(schema);
    var context = buildContext_(ss, schema, worker, m.from);
    var result = opts.fakeClaude ? opts.fakeClaude(system, context, text) : askClaude_(system, context, text, media);

    if (result.question && (!result.actions || !result.actions.length)) {
      return finish_(ss, m, worker, text, result.question, '', 'Question', opts);
    }
    var applied = applyActions_(ss, schema, result.actions || [], opts.dryRun);
    var reply = composeReply_(result, applied);
    return finish_(ss, m, worker, text, reply, applied.undo.length ? JSON.stringify(applied.undo) : '',
                   applied.errors.length ? (applied.writes.length ? 'Partial' : 'Error') : (applied.writes.length ? 'Done' : 'No changes'),
                   opts, applied.summary);
  } catch (err) {
    var emsg = 'Sorry, I couldn\'t process that — nothing was written. (' + (err && err.message || err) + ')';
    return finish_(ss, m, worker, text, emsg, '', 'Error', opts);
  } finally {
    if (locked) lock.releaseLock();
  }
}

function finish_(ss, m, worker, text, reply, undoJson, status, opts, summary) {
  if (!opts || !opts.dryRun) {
    sendWhatsApp_(m.from, reply);
    logRow_(ss, m, worker ? worker.id : '', text, summary || '', undoJson, status, reply);
  }
  return { reply: reply, status: status, undo: undoJson };
}

function messageText_(m) {
  if (m.type === 'text') return (m.text && m.text.body) || '';
  if (m.type === 'image') return (m.image && m.image.caption) || '(photo)';
  if (m.type === 'document') return (m.document && m.document.caption) || ('(file: ' + ((m.document && m.document.filename) || 'document') + ')');
  if (m.type === 'audio') return '(voice note)';
  if (m.type === 'button') return (m.button && m.button.text) || '';
  return '(' + m.type + ' message)';
}

function digits_(s) { return String(s || '').replace(/\D/g, ''); }

function findWorker_(ss, from) {
  var sh = ss.getSheetByName('Team');
  if (!sh) return null;
  var info = sheetInfo_(sh);
  var cId = info.col('Worker ID'), cName = info.col('Name'), cPhone = info.colStarts('Phone');
  var rows = dataRows_(sh, info.lastCol);
  var f = digits_(from);
  for (var i = 0; i < rows.length; i++) {
    var p = digits_(rows[i][cPhone]);
    if (p && f && (p === f || (p.length >= 10 && f.slice(-10) === p.slice(-10)))) {
      return { id: String(rows[i][cId]), name: String(rows[i][cName]) };
    }
  }
  return null;
}

// =====================================================================================
// Reading the workbook's structure (so the bot follows whatever the tabs look like today)
// =====================================================================================
function sheetInfo_(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(HDR, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  return {
    headers: headers, lastCol: lastCol,
    col: function (h) { var i = headers.indexOf(h); if (i < 0) throw new Error('Column "' + h + '" not found on ' + sh.getName()); return i; },
    colStarts: function (h) { for (var i = 0; i < headers.length; i++) if (headers[i].indexOf(h) === 0) return i; throw new Error('Column "' + h + '…" not found on ' + sh.getName()); }
  };
}

function dataRows_(sh, lastCol) {
  var last = sh.getLastRow();
  if (last < FIRST) return [];
  var vals = sh.getRange(FIRST, 1, last - FIRST + 1, lastCol).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] !== '' && vals[i][0] !== null) { vals[i]._row = FIRST + i; out.push(vals[i]); }
  }
  return out;
}

function readSchema_(ss) {
  var out = {};
  DATA_SHEETS.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var info = sheetInfo_(sh);
    var r5 = sh.getRange(FIRST, 1, 1, info.lastCol);
    var formulas = r5.getFormulas()[0], formats = r5.getNumberFormats()[0], dvs = r5.getDataValidations()[0];
    var cols = [];
    info.headers.forEach(function (h, i) {
      if (!h) return;
      var input = !formulas[i];
      var options = null, ref = null;
      if (input && dvs[i]) {
        var type = dvs[i].getCriteriaType();
        var vals = dvs[i].getCriteriaValues();
        if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          var rng = vals[0];
          if (rng.getSheet().getName() === 'Settings') {
            options = uniq_(rng.getValues().map(function (r) { return String(r[0]); }).filter(function (v) { return v !== ''; }));
          } else {
            ref = rng.getSheet().getName();
          }
        } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          options = vals[0].map(String);
        }
      }
      cols.push({ h: h, i: i, input: input, fmt: String(formats[i] || ''), options: options, ref: ref });
    });
    out[name] = { name: name, sheet: sh, info: info, cols: cols };
  });
  return out;
}

function uniq_(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }

function kindOf_(fmt) {
  var f = String(fmt || '').toLowerCase();
  if (f === '@') return 'text';
  if (/y/.test(f) && /d/.test(f)) return 'date';
  if (/h:mm|am\/pm/.test(f)) return 'time';
  if (f.indexOf('%') >= 0) return 'percent';
  if (/[0#]/.test(f)) return 'number';
  return 'text';
}

// =====================================================================================
// Prompts
// =====================================================================================
function buildSystemPrompt_(schema) {
  var lines = [];
  lines.push('You are the data-entry assistant for a small home-services contracting business (HVAC, remodeling & renovations, flooring, carpentry). The owners text you from job sites over WhatsApp. You turn each message into rows in their Google Sheets workbook, or answer a question about the business from the data provided.');
  lines.push('');
  lines.push('HOW TO RESPOND');
  lines.push('- Return JSON only, matching the schema: reply, question, actions.');
  lines.push('- actions: each one adds a row ("add") or changes an existing row ("update") on one tab. fields lists column/value pairs.');
  lines.push('- Use column names EXACTLY as listed below, and only the columns listed (the rest are formulas and must never be written).');
  lines.push('- Values are strings. Dates: YYYY-MM-DD. Times: HH:MM (24-hour). Money and numbers: plain digits, no $ or commas (e.g. 3650.5). Percentages: decimals (25% = 0.25).');
  lines.push('- Columns marked "one of" must use one of those options exactly as written. If nothing fits, use the closest option and say so in reply.');
  lines.push('- For "add" on a tab with an ID column, set row_id to a placeholder like "$client1" or "$job1" and leave the ID column out of fields; the system assigns the real ID. Use the same placeholder as a field value in later actions to link rows created in the same message (e.g. a new Property\'s "Client ID (owner)" = "$client1"). For tabs without an ID column (Time Log, Expenses, Market Rates) row_id is "".');
  lines.push('- For "update", row_id is the existing ID (e.g. "J-0004", "Q-0012", "INV-0007"), and fields contains only the columns that change. Time Log, Expenses and Market Rates rows cannot be updated.');
  lines.push('- Never invent facts. Leave a column out rather than guess. Missing details are fine; the owners fill gaps later.');
  lines.push('- If you cannot tell which job/client/quote a message is about, or it would be risky to guess, return no actions and ask ONE short question in "question". Otherwise question is "".');
  lines.push('- If the message is a question about the business, answer it in reply from the data provided and return no actions.');
  lines.push('- reply: one or two short plain-text sentences for WhatsApp (no markdown). The system appends the exact list of what was written, so don\'t repeat every field.');
  lines.push('');
  lines.push('BUSINESS RULES');
  lines.push('- Time Log: one row per person, per job, per task, per day. "Me" = the sender. "Me and Sam", "both of us", "the crew" = one row for each person named or implied (look them up on the Team list). If start and end times are given, fill Start Time and End Time (and Unpaid Break (min) if a lunch/break is mentioned); otherwise fill "Hours (if no start/end)". Project ID (or OVERHEAD) = the job\'s Project ID, or OVERHEAD for admin, shop, training, or estimates for jobs not yet won. Pick the closest Task / Phase; if a day covered several tasks without a split, use the main task.');
  lines.push('- Expenses: Project ID (or OVERHEAD) = the job the purchase was for, or OVERHEAD for general business costs. Job categories need a job; overhead categories use OVERHEAD. From a receipt photo, read the vendor, date, pre-tax amount and sales tax separately (Amount (before tax) $ and Sales Tax $), and put a short item list in Description. Set Receipt Link / # to "WhatsApp photo". Use a vendor name from the Vendors list when it matches.');
  lines.push('- New lead / new customer: check the client list first (match by phone or name). If new, add a Clients row (First Contact Date = today, Original Lead Source if mentioned), a Properties row if an address is given, and a Quotes row (Lead Date = today, Status "Draft" unless a price was already sent). Link them with placeholders.');
  lines.push('- Quote sent: update the quote with Quote Sent Date and Quoted Price $, Status "Sent". Quote won: Status "Won", Decision Date; also add a Projects row (Status "Scheduled", same Client ID / Property ID / Project Type, Quote ID, Original Contract $ = quoted price) and set the quote\'s "Project ID (if Won)" to the new project\'s placeholder. Quote lost: Status "Lost", Decision Date, Lost Reason, and competitor name/price if mentioned (also add a Market Rates row when a competitor price is given).');
  lines.push('- Job started: Actual Start and Status "In Progress". Job finished: Actual Finish and Status "Completed".');
  lines.push('- Payment received: if it matches an open invoice, update that invoice\'s Amount Paid $ (new total paid), Date Paid and Payment Method. If there is no invoice, add an Invoices row with Invoice Date = Date Paid = today, Amount Invoiced $ = Amount Paid $ = the amount, and the right Invoice Type.');
  lines.push('- Money you are told about without a job (insurance, truck payment, ads) is an overhead expense.');
  lines.push('- Use today\'s date when the message says "today" or gives no date; resolve "yesterday", "Monday", etc. from today\'s date.');
  lines.push('');
  lines.push('THE TABS YOU CAN WRITE (input columns only)');
  Object.keys(schema).forEach(function (name) {
    var s = schema[name];
    var idNote = ID_FORMAT[name] ? ' — ID column "' + s.cols[0].h + '", format ' + ID_FORMAT[name][0] + new Array(ID_FORMAT[name][1] + 1).join('0') : ' — no ID column';
    lines.push('');
    lines.push('## ' + name + idNote);
    s.cols.forEach(function (c) {
      if (!c.input) return;
      var t = kindOf_(c.fmt);
      var desc = '- ' + c.h + ' [' + t + ']';
      if (c.options) desc += ' one of: ' + c.options.join(' | ');
      else if (c.ref) desc += ' (an ID from ' + c.ref + ')';
      lines.push(desc);
    });
  });
  return lines.join('\n');
}

function buildContext_(ss, schema, worker, from) {
  var tz = ss.getSpreadsheetTimeZone();
  var now = new Date();
  var L = [];
  L.push('Today: ' + Utilities.formatDate(now, tz, 'yyyy-MM-dd (EEEE)') + ', time ' + Utilities.formatDate(now, tz, 'HH:mm') + ' (' + tz + ')');
  L.push('Sender: ' + worker.name + ' (Worker ID ' + worker.id + ')');
  L.push('');
  L.push('TEAM: ' + table_(schema['Team'], ['Worker ID', 'Name', 'Role', 'Status'], 50));
  L.push('');
  L.push('PROJECTS (open, plus completed in the last 120 days):');
  L.push(table_(schema['Projects'], ['Project ID', 'Project Name', 'Client Name', 'Property ID', 'City', 'Project Type', 'Status',
    'Actual Start', 'Actual Finish', 'Total Revenue $', 'Gross Profit $', 'Gross Profit / Labor Hr', 'Actual Labor Hours', 'Balance Owed $'], 80, function (r, get) {
      var st = get('Status');
      if (st !== 'Completed' && st !== 'Cancelled') return true;
      var end = get('Actual Finish');
      return end instanceof Date && (now - end) / 86400000 <= 120;
    }));
  L.push('');
  L.push('QUOTES (recent 60):');
  L.push(table_(schema['Quotes'], ['Quote ID', 'Lead Date', 'Client Name', 'Property ID', 'Project Type', 'Quoted Price $', 'Status', 'Project ID (if Won)'], 60));
  L.push('');
  L.push('CLIENTS (recent 150):');
  L.push(table_(schema['Clients'], ['Client ID', 'Client Name', 'Phone', 'Client Type'], 150));
  L.push('');
  L.push('PROPERTIES (recent 150):');
  L.push(table_(schema['Properties'], ['Property ID', 'Client ID (owner)', 'Street Address', 'City'], 150));
  L.push('');
  L.push('OPEN INVOICES:');
  L.push(table_(schema['Invoices'], ['Invoice #', 'Project ID', 'Client Name', 'Amount Invoiced $', 'Amount Paid $', 'Balance $', 'Status'], 60,
    function (r, get) { return get('Status') !== 'Paid'; }));
  L.push('');
  L.push('VENDORS: ' + table_(schema['Vendors'], ['Vendor Name'], 150).replace(/\n/g, '; '));
  L.push('');
  L.push('DASHBOARD: ' + statsText_(ss).replace(/\n/g, '; '));
  var hist = recentConversation_(ss, from, 4);
  if (hist) { L.push(''); L.push('RECENT CONVERSATION WITH THIS SENDER (oldest first):'); L.push(hist); }
  return L.join('\n');
}

// Compact pipe-separated table of the last `limit` rows of a tab.
function table_(s, cols, limit, filter) {
  if (!s) return '(tab missing)';
  var tz = s.sheet.getParent().getSpreadsheetTimeZone();
  var idx = cols.map(function (h) { return s.info.headers.indexOf(h); });
  var rows = dataRows_(s.sheet, s.info.lastCol);
  var out = [];
  for (var i = rows.length - 1; i >= 0 && out.length < limit; i--) {
    var r = rows[i];
    var get = function (h) { var k = s.info.headers.indexOf(h); return k < 0 ? '' : r[k]; };
    if (filter && !filter(r, get)) continue;
    out.push(idx.map(function (k) { return k < 0 ? '' : fmtVal_(r[k], tz); }).join(' | '));
  }
  if (!out.length) return '(none)';
  return cols.join(' | ') + '\n' + out.reverse().join('\n');
}

function fmtVal_(v, tz) {
  if (v instanceof Date) {
    if (v.getFullYear() < 1901) return Utilities.formatDate(v, tz, 'HH:mm');
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  if (typeof v === 'number') return String(Math.round(v * 100) / 100);
  return String(v).replace(/\s+/g, ' ').slice(0, 60);
}

function recentConversation_(ss, from, n) {
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < FIRST) return '';
  var vals = sh.getRange(FIRST, 1, sh.getLastRow() - FIRST + 1, 9).getValues();
  var f = digits_(from), out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < n; i--) {
    if (digits_(vals[i][1]) === f) out.push('Them: ' + vals[i][3] + '\nYou: ' + String(vals[i][8]).split('\n')[0]);
  }
  return out.reverse().join('\n');
}

// =====================================================================================
// Claude
// =====================================================================================
var RESPONSE_SCHEMA = {
  type: 'object',
  properties: {
    reply: { type: 'string' },
    question: { type: 'string' },
    actions: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          sheet: { type: 'string', enum: DATA_SHEETS },
          operation: { type: 'string', enum: ['add', 'update'] },
          row_id: { type: 'string' },
          fields: {
            type: 'array',
            items: {
              type: 'object',
              properties: { column: { type: 'string' }, value: { type: 'string' } },
              required: ['column', 'value'],
              additionalProperties: false
            }
          }
        },
        required: ['sheet', 'operation', 'row_id', 'fields'],
        additionalProperties: false
      }
    }
  },
  required: ['reply', 'question', 'actions'],
  additionalProperties: false
};

function askClaude_(system, context, text, media) {
  var key = prop_('ANTHROPIC_API_KEY', '');
  if (!key) throw new Error('ANTHROPIC_API_KEY is not set in Script Properties');
  var model = prop_('CLAUDE_MODEL', 'claude-opus-5');

  var content = [];
  if (media) {
    if (media.mime === 'application/pdf') {
      content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: media.b64 } });
    } else if (/^image\/(jpeg|png|gif|webp)$/.test(media.mime)) {
      content.push({ type: 'image', source: { type: 'base64', media_type: media.mime, data: media.b64 } });
    }
  }
  content.push({ type: 'text', text: context + '\n\nNEW MESSAGE FROM THE SENDER:\n' + text });

  var body = {
    model: model,
    max_tokens: 16000,
    output_config: { effort: prop_('CLAUDE_EFFORT', 'medium'), format: { type: 'json_schema', schema: RESPONSE_SCHEMA } },
    system: [{ type: 'text', text: system, cache_control: { type: 'ephemeral' } }],
    messages: [{ role: 'user', content: content }]
  };
  var headers = { 'x-api-key': key, 'anthropic-version': '2023-06-01' };
  if (/^claude-(opus-5|fable)/.test(model)) {          // server-side refusal fallback (routes a declined request to another model)
    body.fallbacks = 'default';
    headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
  }
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', headers: headers,
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var data = JSON.parse(res.getContentText() || '{}');
  if (code !== 200) {
    var em = (data.error && data.error.message) || res.getContentText().slice(0, 200);
    if (code === 429 || code === 529 || code >= 500) throw new Error('Claude is busy right now, try again in a minute (' + code + ')');
    throw new Error('Claude API error ' + code + ': ' + em);
  }
  if (data.stop_reason === 'refusal') throw new Error('Claude declined this message');
  if (data.stop_reason === 'max_tokens') throw new Error('the reply was cut off — try a shorter message');
  var textBlock = (data.content || []).filter(function (b) { return b.type === 'text'; }).pop();
  if (!textBlock) throw new Error('empty reply from Claude');
  return JSON.parse(textBlock.text);
}

// =====================================================================================
// Writing to the sheet
// =====================================================================================
function applyActions_(ss, schema, actions, dryRun) {
  var placeholders = {};
  var writes = [], errors = [], undo = [], summary = [];
  var tz = ss.getSpreadsheetTimeZone();

  actions.forEach(function (a, n) {
    var s = schema[a.sheet];
    if (!s) { errors.push('Tab "' + a.sheet + '" not found'); return; }
    var sh = s.sheet;
    var byHeader = {};
    s.cols.forEach(function (c) { byHeader[c.h.toLowerCase()] = c; });
    var idCol = ID_FORMAT[a.sheet] ? s.cols[0] : null;
    var row, isNew = a.operation === 'add', label;

    if (isNew) {
      row = firstEmptyRow_(sh, s.info.lastCol, s.cols, dryRun);
    } else {
      if (!idCol) { errors.push(a.sheet + ' rows can\'t be updated from WhatsApp — edit them in the sheet'); return; }
      var target = resolve_(a.row_id, placeholders);
      row = findRowById_(sh, target);
      if (!row) { errors.push(a.sheet + ': no row with ID ' + target); return; }
    }

    var rowWrites = [];
    if (isNew && idCol) {
      var newId = nextId_(sh, a.sheet);
      if (a.row_id) placeholders[a.row_id] = newId;
      rowWrites.push({ c: idCol, v: newId });
      label = newId;
    }
    (a.fields || []).forEach(function (f) {
      var c = byHeader[String(f.column).trim().toLowerCase()];
      if (!c) { errors.push(a.sheet + ': unknown column "' + f.column + '"'); return; }
      if (!c.input) { errors.push(a.sheet + ': "' + c.h + '" is calculated automatically, skipped'); return; }
      if (idCol && c === idCol) return;             // IDs are assigned by the system
      var raw = resolve_(String(f.value), placeholders);
      var conv = convert_(raw, c);
      if (conv.warn) errors.push(a.sheet + ' · ' + c.h + ': ' + conv.warn);
      rowWrites.push({ c: c, v: conv.v });
    });
    if (!rowWrites.length) return;
    if (!label) label = idCol ? resolve_(a.row_id, placeholders) : 'row ' + row;

    rowWrites.forEach(function (w) {
      var cell = sh.getRange(row, w.c.i + 1);
      var old = cell.getValue();
      undo.push({ s: a.sheet, r: row, c: w.c.i + 1, o: serial_(old, w.c, tz), n: isNew ? 1 : 0 });
      if (!dryRun) cell.setValue(w.v);
    });
    writes.push({ sheet: a.sheet, row: row, isNew: isNew, label: label, fields: rowWrites });
    summary.push((isNew ? 'Added ' : 'Updated ') + a.sheet + ' ' + label + ': ' +
      rowWrites.filter(function (w) { return w.c !== idCol; }).map(function (w) { return w.c.h + '=' + w.v; }).join('; '));
  });
  if (!dryRun) SpreadsheetApp.flush();
  return { writes: writes, errors: errors, undo: undo, summary: summary.join('\n') };
}

function resolve_(v, placeholders) {
  v = String(v == null ? '' : v);
  return Object.prototype.hasOwnProperty.call(placeholders, v) ? placeholders[v] : v;
}

// Value in the form Sheets will parse the same way a person typing it would.
function convert_(raw, c) {
  var s = String(raw).trim();
  if (s === '') return { v: '' };
  var kind = kindOf_(c.fmt);
  var warn = null;
  if (c.options && c.options.indexOf(s) < 0) {
    var hit = c.options.filter(function (o) { return o.toLowerCase() === s.toLowerCase(); })[0];
    if (hit) s = hit; else warn = '"' + s + '" isn\'t in the dropdown list (kept anyway)';
  }
  if (kind === 'date') {
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return { v: m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]), warn: warn };
    return { v: s, warn: warn || 'date not in YYYY-MM-DD form' };
  }
  if (kind === 'time') {
    var t = s.match(/^(\d{1,2}):(\d{2})$/);
    if (t) return { v: Number(t[1]) + ':' + t[2], warn: warn };
    return { v: s, warn: warn || 'time not in HH:MM form' };
  }
  if (kind === 'number' || kind === 'percent') {
    var num = Number(s.replace(/[$,\s]/g, '').replace(/%$/, ''));
    if (isNaN(num)) return { v: s, warn: warn || 'expected a number' };
    if (kind === 'percent' && (/%$/.test(s) || num > 1)) num = num / 100;
    return { v: num, warn: warn };
  }
  if (/^[=+]/.test(s) && c.fmt !== '@') s = "'" + s;   // never let a message become a formula
  return { v: s, warn: warn };
}

function pad_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

function serial_(v, c, tz) {
  if (v instanceof Date) {
    return kindOf_(c.fmt) === 'time' ? Utilities.formatDate(v, tz, 'H:mm') : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return v;
}

function firstEmptyRow_(sh, lastCol, cols, dryRun) {
  var last = Math.max(sh.getLastRow(), FIRST);
  var colA = sh.getRange(FIRST, 1, last - FIRST + 1, 1).getValues();
  for (var i = 0; i < colA.length; i++) {
    if (colA[i][0] === '' || colA[i][0] === null) {
      if (!firstEmptyRow_.taken) firstEmptyRow_.taken = {};
      var key = sh.getName() + ':' + (FIRST + i);
      if (dryRun && firstEmptyRow_.taken[key]) continue;   // dry runs don't write, so skip rows already "used"
      if (dryRun) firstEmptyRow_.taken[key] = 1;
      return FIRST + i;
    }
  }
  // All pre-filled rows are used: add a row and copy the formulas down from the last one.
  var target = last + 1;
  if (!dryRun) {
    if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    sh.getRange(last, 1, 1, lastCol).copyTo(sh.getRange(target, 1, 1, lastCol));
    cols.forEach(function (c) { if (c.input) sh.getRange(target, c.i + 1).clearContent(); });
  }
  return target;
}

function findRowById_(sh, id) {
  if (!id) return null;
  var last = sh.getLastRow();
  if (last < FIRST) return null;
  var colA = sh.getRange(FIRST, 1, last - FIRST + 1, 1).getValues();
  var want = String(id).trim().toUpperCase();
  for (var i = 0; i < colA.length; i++) if (String(colA[i][0]).trim().toUpperCase() === want) return FIRST + i;
  return null;
}

function nextId_(sh, name) {
  var fmt = ID_FORMAT[name];
  var last = sh.getLastRow();
  var max = 0;
  if (last >= FIRST) {
    sh.getRange(FIRST, 1, last - FIRST + 1, 1).getValues().forEach(function (r) {
      var m = String(r[0]).match(/(\d+)\s*$/);
      if (m && String(r[0]).toUpperCase().indexOf(fmt[0]) === 0) max = Math.max(max, Number(m[1]));
    });
  }
  if (!nextId_.issued) nextId_.issued = {};
  var n = Math.max(max, nextId_.issued[name] || 0) + 1;
  nextId_.issued[name] = n;
  var s = String(n);
  while (s.length < fmt[1]) s = '0' + s;
  return fmt[0] + s;
}

// =====================================================================================
// Replies, undo, stats, log
// =====================================================================================
function composeReply_(result, applied) {
  var lines = [];
  if (result.reply) lines.push(result.reply);
  applied.writes.forEach(function (w) {
    var shown = w.fields.filter(function (f) { return String(f.v) !== w.label; }).slice(0, 6)
      .map(function (f) { return f.c.h.replace(/ \(.*\)$/, '') + ': ' + f.v; });
    lines.push((w.isNew ? '✅ Added ' : '✏️ Updated ') + w.sheet + ' ' + w.label + ' — ' + shown.join(' · '));
  });
  if (applied.errors.length) lines.push('⚠️ ' + applied.errors.join('\n⚠️ '));
  if (applied.writes.length) lines.push('Reply UNDO to reverse this.');
  var out = lines.join('\n');
  return out.length > 3900 ? out.slice(0, 3900) + '…' : out;
}

function undoLast_(ss, from) {
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < FIRST) return 'Nothing to undo.';
  var vals = sh.getRange(FIRST, 1, sh.getLastRow() - FIRST + 1, 8).getValues();
  var f = digits_(from);
  for (var i = vals.length - 1; i >= 0; i--) {
    var st = vals[i][7];
    if (digits_(vals[i][1]) !== f || !vals[i][6] || (st !== 'Done' && st !== 'Partial')) continue;
    var cells = JSON.parse(vals[i][6]);
    for (var k = cells.length - 1; k >= 0; k--) {
      var x = cells[k];
      var t = ss.getSheetByName(x.s);
      if (t) t.getRange(x.r, x.c).setValue(x.o === null ? '' : x.o);
    }
    sh.getRange(FIRST + i, 8).setValue('Undone');
    SpreadsheetApp.flush();
    return '↩️ Undone: ' + String(vals[i][3]).slice(0, 80) + ' (' + cells.length + ' cells restored).';
  }
  return 'Nothing to undo — your last change was already undone or wasn\'t a write.';
}

function helpText_() {
  return [
    'Text me what happened and I\'ll log it. Examples:',
    '• Me and Sam 8 to 4:30 on the Johnson AC job, install, 30 min lunch',
    '• (photo of a receipt) Johnson job',
    '• New lead: Maria Lopez 312-555-0199, kitchen remodel at 44 Oak St, found us on Google',
    '• Sent Maria the quote, 18,500',
    '• Maria said yes, starting Monday',
    '• Johnson job done today',
    '• Got paid 4,600 by check from Johnson',
    '• How much did we make on the Johnson job?',
    'Commands: UNDO (reverse my last change) · STATS (key numbers) · HELP'
  ].join('\n');
}

function statsText_(ss) {
  var sh = ss.getSheetByName('Dashboard');
  if (!sh) return 'Dashboard tab not found.';
  var want = ['Revenue', 'Gross Profit', 'Gross Margin', 'Net Profit', 'Cash Collected', 'Owed to You Now',
              'Gross Profit / Job Hour', 'Win Rate', 'Projects Completed', 'Overdue Invoices'];
  var vals = sh.getRange(1, 1, 40, 9).getDisplayValues();
  var found = {};
  for (var r = 0; r < vals.length - 1; r++) {
    for (var c = 0; c < 9; c++) {
      var lbl = vals[r][c];
      if (want.indexOf(lbl) >= 0 && !found[lbl]) found[lbl] = vals[r + 1][c];
    }
  }
  var period = sh.getRange('C4').getDisplayValue();
  return ['📊 ' + period].concat(want.filter(function (w) { return found[w] !== undefined; })
    .map(function (w) { return w + ': ' + (found[w] || '—'); })).join('\n');
}

function logRow_(ss, m, workerId, text, summary, undoJson, status, reply) {
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) return;
  var row = FIRST;
  var last = sh.getLastRow();
  if (last >= FIRST) {
    var a = sh.getRange(FIRST, 1, last - FIRST + 1, 1).getValues();
    row = last + 1;
    for (var i = 0; i < a.length; i++) if (a[i][0] === '') { row = FIRST + i; break; }
  }
  sh.getRange(row, 1, 1, 10).setValues([[new Date(), String(m.from), workerId, text, (m.type === 'image' || m.type === 'document') ? 'Yes' : '',
    summary, undoJson, status, reply, m.id || '']]);
}

// =====================================================================================
// WhatsApp Cloud API
// =====================================================================================
function graph_(path) { return 'https://graph.facebook.com/' + prop_('GRAPH_VERSION', 'v23.0') + '/' + path; }

function sendWhatsApp_(to, body) {
  var token = prop_('WHATSAPP_TOKEN', ''), pid = prop_('WHATSAPP_PHONE_NUMBER_ID', '');
  if (!token || !pid) { console.warn('WhatsApp not configured; reply was: ' + body); return; }
  var res = UrlFetchApp.fetch(graph_(pid + '/messages'), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ messaging_product: 'whatsapp', to: String(to), type: 'text', text: { body: body, preview_url: false } })
  });
  if (res.getResponseCode() >= 300) console.error('WhatsApp send failed: ' + res.getContentText());
}

function fetchMedia_(m) {
  var obj = m.type === 'image' ? m.image : m.document;
  if (!obj || !obj.id) return null;
  var token = prop_('WHATSAPP_TOKEN', '');
  var meta = JSON.parse(UrlFetchApp.fetch(graph_(obj.id), { headers: { Authorization: 'Bearer ' + token } }).getContentText());
  var blob = UrlFetchApp.fetch(meta.url, { headers: { Authorization: 'Bearer ' + token } }).getBlob();
  var mime = (meta.mime_type || obj.mime_type || blob.getContentType() || '').split(';')[0];
  if (blob.getBytes().length > 4.5 * 1024 * 1024) throw new Error('that file is too large (over 4.5 MB)');
  return { mime: mime, b64: Utilities.base64Encode(blob.getBytes()) };
}

// =====================================================================================
// Run these from the Apps Script editor
// =====================================================================================

/** 1) Run once after pasting the code: checks the workbook and settings, and grants permissions. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var problems = [];
  DATA_SHEETS.concat(['Dashboard', LOG_SHEET]).forEach(function (n) { if (!ss.getSheetByName(n)) problems.push('Missing tab: ' + n); });
  ['ANTHROPIC_API_KEY', 'WHATSAPP_TOKEN', 'WHATSAPP_PHONE_NUMBER_ID', 'RELAY_KEY'].forEach(function (k) {
    if (!prop_(k, '')) problems.push('Script Property not set yet: ' + k);
  });
  var schema = readSchema_(ss);
  var nInputs = 0;
  Object.keys(schema).forEach(function (k) { nInputs += schema[k].cols.filter(function (c) { return c.input; }).length; });
  var team = dataRows_(ss.getSheetByName('Team'), sheetInfo_(ss.getSheetByName('Team')).lastCol).length;
  console.log('Tabs read: ' + Object.keys(schema).length + ', input columns: ' + nInputs + ', team members: ' + team);
  console.log('Spreadsheet time zone: ' + ss.getSpreadsheetTimeZone());
  console.log(problems.length ? 'TO DO:\n' + problems.join('\n') : 'All set.');
}

/** Makes a long random RELAY_KEY for you. Run it, copy the value from the log. */
function makeKey() {
  console.log((Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
}

/** 2) Test Claude without WhatsApp and without writing anything. Edit the message, then Run. */
function testMessage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var team = ss.getSheetByName('Team');
  var info = sheetInfo_(team);
  var first = team.getRange(FIRST, 1, 1, info.lastCol).getValues()[0];
  var worker = { id: String(first[info.col('Worker ID')]), name: String(first[info.col('Name')]) };
  var r = handleMessage_({ id: 'test-' + Date.now(), from: '10000000000', type: 'text',
    text: { body: 'Me and Sam worked 8 to 4:30 today on J-0001, install, 30 min lunch' } }, { dryRun: true, worker: worker });
  console.log(r.reply);
}
