/**
 * EMB ORDER PROCESS — GOOGLE APPS SCRIPT BACKEND
 *
 * PURPOSE
 * -------
 * Central order-processing backend for an EMB order workflow.
 *
 * IMPORTANT
 * ---------
 * 1. The Orders sheet is already populated by an existing API/import process.
 *    This script does NOT fetch/import Orders data and does NOT modify that
 *    integration.
 * 2. The godown spreadsheets already have their own processing/callback script.
 *    This script only sends the selected order to the selected godown sheet.
 * 3. The central spreadsheet is the database for this application.
 *
 * EXPECTED SHEETS
 * ---------------
 * Orders
 * Users
 * G Apv
 * Production Plan   (one row per order - plan header)
 * Plan Steps        (one row per order x step - analysis table, auto-built)

 * Notifications
 * Log
 *
 * ORDERS BASE COLUMNS
 * -------------------
 * UUID | Date | Order No | Customer | EMB Sales Coordinator | Design No |
 * Colour | Fabric | Yarn | Order Type | Order QTY | Qty Unit | Rate |
 * Width | Delivery Date
 *
 * APP WORKFLOW COLUMNS ADDED TO ORDERS IF MISSING
 * ------------------------------------------------
 * Status | External Inv Qty | Godown | Sent Timestamp |
 * Remaining Qty | Production Decision
 *
 * G APV ACTUAL STRUCTURE
 * ----------------------
 * UUID | Date | Order No | Customer | EMB Sales Coordinator | Design No |
 * Colour | Fabric | Yarn | Order Type | Order QTY | Qty Unit | Rate |
 * Width | Delivery Date | Timestamp | Greige Mending | Finish Mending |
 * Process | Lace cutting | QC & Dispatch
 *
 * G APV IS SOURCE OF TRUTH FOR THE RETURNED GODOWN BREAKDOWN.
 *
 * GODOWN ROUTING
 * --------------
 * All Over -> 1nSeJ7ek_FxWBVroRVKTfFQFKZung6HVtQ6pxrFoPX_s
 * Lace     -> 1jVt66G_GCBX8-M4pmxA-p9Q47AG6pfo-ODRczqUwkHE
 * Solid    -> 1FrCv9lrSQPZzxwPr2ZqvgvvzTTPgLCmT_mReLTH8TV8
 *
 * The callback/processing logic inside those godown spreadsheets is NOT
 * implemented here.
 *
 * JSON-SAFE API
 * -------------
 * Every *_Json endpoint returns a JSON STRING. A generated Index.html can
 * call these endpoints with google.script.run and JSON.parse the response.
 *
 * This avoids Apps Script HTML-service serialization problems with Date values.
 */


/* ============================================================================
   1. CONFIG
   ========================================================================== */

const APP = {
  SHEETS: {
    USERS: 'Users',
    ORDERS: 'Orders',
    GAPV: 'G Apv',
    PLAN: 'Production Plan',
    PLAN_STEPS: 'Plan Steps',
    ALLOCATIONS: 'Production Allocation',

    NOTIFICATIONS: 'Notifications',
    LOG: 'Log'
  },

  GODOWNS: {
    'All Over': '1nSeJ7ek_FxWBVroRVKTfFQFKZung6HVtQ6pxrFoPX_s',
    'Lace': '1jVt66G_GCBX8-M4pmxA-p9Q47AG6pfo-ODRczqUwkHE',
    'Solid': '1FrCv9lrSQPZzxwPr2ZqvgvvzTTPgLCmT_mReLTH8TV8'
  },

  GODOWN_SHEET_NAME: 'Incoming Orders',

  /*
   * Where the production is physically made. Each production unit has its
   * own spreadsheet; allocations are pushed into that spreadsheet.
   */
  PRODUCTION_UNITS: {
    'Multi': '1txK5bnFh4_ZHdeWDTMueOLrd7YQ04w-acrbBHJZhaQs',
    'Schiffli': '1DULOkvFeEofPfMtYnoe92S0JpVtrdGSNJp0eQgjEeMA',
    'Outside': '10_hZgYPdzitprLeK7G2Iz76owXbAcdOHrQNc8t6_Xv4'
  },

  PRODUCTION_UNIT_SHEET_NAME: 'Incoming Production',

  /* Allowed production place selections and the units each one covers. */
  PRODUCTION_PLACES: {
    'Schiffli': ['Schiffli'],
    'Multi': ['Multi'],
    'Outside': ['Outside'],
    'Schiffli + Multi': ['Schiffli', 'Multi'],
    'Schiffli + Multi + Outside': ['Schiffli', 'Multi', 'Outside']
  },

  /*
   * Production Allocation = ONE row per order x machine/party split.
   * Schiffli / Multi rows carry a patti size + production qty.
   * Outside rows additionally carry the party name (one row per party,
   * each with its own patti size and split qty).
   */
  ALLOCATION_COLUMNS: [
    'Allocation ID',
    'Plan ID',
    'UUID',
    'Order No',
    'Order Date',
    'Customer',
    'EMB Sales Coordinator',
    'Design No',
    'Colour',
    'Fabric',
    'Order Type',
    'Delivery Date',
    'Grid',
    'Revo',
    'Desc',
    'Greige/Finish',
    'Production Place',
    'Production Unit',
    'Party Name',
    'Patti Size',
    'Production Qty',
    'Allocation Index',
    'Total Allocations',
    'Dispatched',
    'Dispatched Timestamp',
    'Created Timestamp',
    'Updated Timestamp'
  ],

  SESSION_TTL_SECONDS: 8 * 60 * 60,

  /*
   * Department roles map 1:1 onto the production flow steps. A user with
   * Role = 'department' handles the step whose key equals their Department.
   */
  DEPARTMENTS: [
    { key: 'production', label: 'Production' },
    { key: 'mending', label: 'Mending' },
    { key: 'process', label: 'Process (Mill)' },
    { key: 'qc', label: 'Final QC' },
    { key: 'lace', label: 'Lace cutting' },
    { key: 'dispatch', label: 'Dispatch/Stock' }
  ],

  /*
   * Dispatch instructions are written to a separate external spreadsheet.
   * Column L (index 12) is "Done" — the sales person flips it to Yes when
   * the dispatch request has been fulfilled. One order may be dispatched in
   * parts across multiple rows; the order cannot be completed until every
   * dispatch request for it is Done.
   */
  DISPATCH_SHEET_ID: '1ZOQSM8u1xckso0BUgbYcYJSn9i9jDY7onxIzIHi4g0I',
  DISPATCH_SHEET_NAME: 'Master_Data',
  DISPATCH_COLUMNS: [
    'UUID',
    'Order No',
    'Customer',
    'Design No',
    'Colour',
    'Fabric',
    'Yarn',
    'QTY',
    'Qty Unit',
    'Timestamp',
    'Owner',
    'Done'
  ],

  ORDER_BASE_COLUMNS: [
    'UUID',
    'Date',
    'Order No',
    'Customer',
    'EMB Sales Coordinator',
    'Design No',
    'Colour',
    'Fabric',
    'Yarn',
    'Order Type',
    'Order QTY',
    'Qty Unit',
    'Rate',
    'Width',
    'Delivery Date'
  ],

  ORDER_WORKFLOW_COLUMNS: [
    'Status',
    'External Inv Qty',
    'Godown',
    'Sent Timestamp',
    'Remaining Qty',
    'Production Decision',
    'Order Closed',
    'Closed Timestamp',
    'Closed By'
  ],


  GAPV_STAGE_COLUMNS: [
    'Greige Mending',
    'Finish Mending',
    'Process',
    'Lace cutting',
    'QC & Dispatch'
  ],

  /*
   * Production Plan = ONE row per order (the plan header / summary).
   * Every column is a clean, single-value field so the sheet can be used
   * directly in pivot tables and charts.
   */
  PLAN_COLUMNS: [
    'Plan ID',
    'UUID',
    'Order No',
    'Order Date',
    'Customer',
    'EMB Sales Coordinator',
    'Design No',
    'Colour',
    'Fabric',
    'Order Type',
    'Order QTY',
    'Delivery Date',
    'Grid',
    'Revo',
    'Desc',
    'Greige/Finish',
    'Remaining Qty',
    'Stock Current Stage',
    'Mending Type',
    'Total Steps',
    'Flow Steps',
    'Flow Sequence',
    'Production Date',
    'Mending Date',
    'Process Date',
    'Final QC Date',
    'Lace Cutting Date',
    'Dispatch Date',
    'Plan Start Date',
    'Plan End Date',
    'Plan Duration Days',
    'Production Place',
    'Production Units',
    'Allocation Count',
    'Allocated Qty',
    'Outside Parties',
    'Current Step Index',
    'Current Step',
    'Current Step Date',
    'Completed Steps',
    'Pending Steps',
    'Lifecycle Status',
    'Created Timestamp',
    'Updated Timestamp'
  ],

  /*
   * Plan Steps = ONE row per order x step. This is the analysis table:
   * every selected step of every order is a separate, fully indexed record.
   */
  PLAN_STEP_COLUMNS: [
    'Step ID',
    'Plan ID',
    'UUID',
    'Order No',
    'Order Date',
    'Customer',
    'EMB Sales Coordinator',
    'Design No',
    'Colour',
    'Fabric',
    'Order Type',
    'Delivery Date',
    'Grid',
    'Revo',
    'Desc',
    'Greige/Finish',
    'Remaining Qty',
    'Stock Current Stage',
    'Mending Type',
    'Step Index',
    'Step Key',
    'Step Name',
    'Planned Date',
    'Previous Step',
    'Next Step',
    'Is First Step',
    'Is Last Step',
    'Total Steps',
    'Days From Plan Start',
    'Production Place',
    'Step Status',
    'Is Current Step',
    'Completed Date',
    'Completed By',
    'Created Timestamp',
    'Updated Timestamp'
  ],


  /*
   * Ordered production flow. The user selects which of these steps will
   * actually be performed for an order, then supplies a completion date
   * for each selected step.
   */
  FLOW_STEPS: [
    { key: 'production', label: 'Production', column: 'Production Date' },
    { key: 'mending', label: 'Mending', column: 'Mending Date' },
    { key: 'process', label: 'Process (Mill)', column: 'Process Date' },
    { key: 'qc', label: 'Final QC', column: 'Final QC Date' },
    { key: 'lace', label: 'Lace cutting', column: 'Lace Cutting Date' },
    { key: 'dispatch', label: 'Dispatch/Stock', column: 'Dispatch Date' }
  ],

  STATUS: {
    NEW: 'New',
    SENT_TO_GODOWN: 'Sent To Godown',
    G_APV_RECEIVED: 'G Apv Received',
    SENT_TO_PRODUCTION: 'Sent To Production',
    STOCK_ONLY: 'Stock Only (No Production)'
  },

  LIFECYCLE: {
    PLANNED: 'Planned',
    IN_PRODUCTION: 'In Production',
    COMPLETED: 'Completed',
    CLOSED: 'Closed'
  },


  /* A step is Pending until it is marked done; then the order moves on. */
  STEP_STATUS: {
    PENDING: 'Pending',
    CURRENT: 'Current',
    COMPLETED: 'Completed'
  },

  // Default lead times between planned completion milestones.
  // These can be changed here without changing the rest of the workflow.
  LEAD_DAYS: {
    PRODUCTION: 7,
    MENDING: 3,
    PROCESS: 5,
    QC: 2,
    LACE: 2,
    DISPATCH: 2
  }
};


/* ============================================================================
   2. BASIC HELPERS
   ========================================================================== */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) {
    throw new Error('Required sheet not found: ' + name);
  }
  return sh;
}

function normalize_(value) {
  return String(value == null ? '' : value).trim();
}

function number_(value) {
  if (value === '' || value == null) return 0;
  const n = Number(String(value).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
}

function now_() {
  return new Date();
}

function timestamp_() {
  return Utilities.formatDate(
    now_(),
    Session.getScriptTimeZone() || 'Asia/Kolkata',
    'yyyy-MM-dd HH:mm:ss'
  );
}

function dateString_(value) {
  if (value == null || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' &&
      !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'yyyy-MM-dd'
    );
  }

  return String(value);
}

function jsonSafe_(value) {
  if (value === null || value === undefined) return null;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return dateString_(value);
  }

  if (Array.isArray(value)) {
    return value.map(jsonSafe_);
  }

  if (typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(function(k) {
      out[k] = jsonSafe_(value[k]);
    });
    return out;
  }

  return value;
}

function jsonString_(value) {
  return JSON.stringify(jsonSafe_(value));
}

function ok_(extra) {
  return Object.assign({ ok: true }, extra || {});
}

function fail_(message, extra) {
  return Object.assign({
    ok: false,
    error: String(message || 'Unknown error')
  }, extra || {});
}

function withError_(fn) {
  try {
    return fn();
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    throw new Error(err && err.message ? err.message : String(err));
  }
}


/* ============================================================================
   3. SHEET / HEADER HELPERS
   ========================================================================== */

function getHeader_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];

  return sheet
    .getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(normalize_);
}

function headerIndex_(header, name) {
  return header.indexOf(name);
}

function columnNumber_(header, name) {
  const index = headerIndex_(header, name);
  if (index === -1) {
    throw new Error('Column not found: ' + name);
  }
  return index + 1;
}

function ensureSheet_(name, headers) {
  let sh = ss_().getSheetByName(name);

  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return sh;
  }

  const current = getHeader_(sh);

  // Empty sheet.
  if (!current.length || current.every(function(v) { return !v; })) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }

  return sh;
}

function ensureOrderWorkflowColumns_() {
  const sh = getSheet_(APP.SHEETS.ORDERS);
  const header = getHeader_(sh);

  if (!header.length) {
    throw new Error('Orders sheet has no header row.');
  }

  const missing = APP.ORDER_WORKFLOW_COLUMNS.filter(function(col) {
    return headerIndex_(header, col) === -1;
  });

  if (missing.length) {
    sh.getRange(1, sh.getLastColumn() + 1, 1, missing.length)
      .setValues([missing]);
  }

  return getHeader_(sh);
}

function ensureUsersSheet_() {
  const sh = ensureSheet_(APP.SHEETS.USERS, [
    'Username',
    'Password',
    'Sales Coordinator Name',
    'Role',
    'Department',
    'Active'
  ]);
  ensureColumns_(sh, [
    'Username',
    'Password',
    'Sales Coordinator Name',
    'Role',
    'Department',
    'Active'
  ]);
  return sh;
}

function ensureNotificationsSheet_() {
  return ensureSheet_(APP.SHEETS.NOTIFICATIONS, [
    'Timestamp',
    'UUID',
    'Order No',
    'Message',
    'Resolved'
  ]);
}

function ensureLogSheet_() {
  return ensureSheet_(APP.SHEETS.LOG, [
    'Timestamp',
    'Action',
    'Detail'
  ]);
}

function ensureColumns_(sh, columns) {
  const header = getHeader_(sh);

  const missing = columns.filter(function(col) {
    return headerIndex_(header, col) === -1;
  });

  if (missing.length) {
    sh.getRange(1, header.length + 1, 1, missing.length)
      .setValues([missing]);
  }

  return sh;
}

function ensurePlanSheet_() {
  return ensureColumns_(
    ensureSheet_(APP.SHEETS.PLAN, APP.PLAN_COLUMNS),
    APP.PLAN_COLUMNS
  );
}

function ensurePlanStepsSheet_() {
  const sh = ensureColumns_(
    ensureSheet_(APP.SHEETS.PLAN_STEPS, APP.PLAN_STEP_COLUMNS),
    APP.PLAN_STEP_COLUMNS
  );

  sh.setFrozenRows(1);

  return sh;
}

function ensureAllocationSheet_() {
  const sh = ensureColumns_(
    ensureSheet_(APP.SHEETS.ALLOCATIONS, APP.ALLOCATION_COLUMNS),
    APP.ALLOCATION_COLUMNS
  );

  sh.setFrozenRows(1);

  return sh;
}


/* ---------------------------------------------------------------------------
   PRODUCTION PLACE / ALLOCATION HELPERS
   ------------------------------------------------------------------------ */

/* Accepts any casing / spacing variant of the allowed place names. */
function sanitizeProductionPlace_(place) {
  const wanted = normalize_(place).toLowerCase().replace(/\s+/g, '');

  const keys = Object.keys(APP.PRODUCTION_PLACES);

  for (let i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase().replace(/\s+/g, '') === wanted) return keys[i];
  }

  return '';
}

function unitsForPlace_(place) {
  const clean = sanitizeProductionPlace_(place);
  return clean ? APP.PRODUCTION_PLACES[clean].slice() : [];
}

function sanitizeProductionUnit_(unit) {
  const wanted = normalize_(unit).toLowerCase();

  const keys = Object.keys(APP.PRODUCTION_UNITS);

  for (let i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === wanted) return keys[i];
  }

  return '';
}

/*
 * Validates the machine / party split.
 *
 * Rules:
 * - Every unit belonging to the selected place must have at least one row.
 * - Schiffli / Multi: patti size + production qty (usually one row each).
 * - Outside: one row per party, each with party name, patti size and qty.
 * - Total allocated qty must match the production qty of the order.
 */
function sanitizeAllocations_(place, allocations, productionQty) {
  const cleanPlace = sanitizeProductionPlace_(place);

  if (!cleanPlace) {
    throw new Error(
      'Select where the production will be made (' +
      Object.keys(APP.PRODUCTION_PLACES).join(', ') + ').'
    );
  }

  const units = unitsForPlace_(cleanPlace);
  const seen = {};
  const out = [];

  (allocations || []).forEach(function(raw) {
    const item = raw || {};
    const unit = sanitizeProductionUnit_(item.unit || item['Production Unit']);

    if (!unit) return;

    if (units.indexOf(unit) === -1) {
      throw new Error(
        unit + ' is not part of the selected production place "' +
        cleanPlace + '".'
      );
    }

    const qty = number_(item.qty !== undefined ? item.qty : item['Production Qty']);
    const patti = normalize_(item.pattiSize || item['Patti Size']);
    const party = normalize_(item.partyName || item['Party Name']);

    if (!patti) {
      throw new Error('Patti size is required for ' + unit + '.');
    }

    if (!(qty > 0)) {
      throw new Error('Production qty must be greater than 0 for ' + unit + '.');
    }

    if (unit === 'Outside' && !party) {
      throw new Error('Party name is required for every Outside allocation.');
    }

    seen[unit] = true;

    out.push({
      unit: unit,
      partyName: unit === 'Outside' ? party : '',
      pattiSize: patti,
      qty: qty
    });
  });

  if (!out.length) {
    throw new Error('Add at least one patti / production split.');
  }

  const missing = units.filter(function(u) { return !seen[u]; });

  if (missing.length) {
    throw new Error(
      'Add patti and production qty for: ' + missing.join(', ') + '.'
    );
  }

  const total = out.reduce(function(sum, a) { return sum + a.qty; }, 0);
  const target = number_(productionQty);

  if (target > 0 && Math.abs(total - target) > 0.5) {
    throw new Error(
      'Split qty (' + total + ') must equal the production qty (' +
      target + ').'
    );
  }

  /* Keep unit order stable: Schiffli, Multi, then Outside parties. */
  const order = ['Schiffli', 'Multi', 'Outside'];

  out.sort(function(a, b) {
    return order.indexOf(a.unit) - order.indexOf(b.unit);
  });

  return { place: cleanPlace, allocations: out, totalQty: total };
}

function allocationIdFor_(planId, index, unit) {
  return planId + '-A' + pad_(index, 2) + '-' +
    String(unit).toUpperCase().substring(0, 3);
}

/*
 * Rewrites the Production Allocation rows for one order. Old rows for the
 * same UUID are removed first so the sheet always holds exactly one clean
 * record per order x machine / party.
 */
function syncAllocations_(context) {
  const sh = ensureAllocationSheet_();
  const header = getHeader_(sh);
  const uuidCol = headerIndex_(header, 'UUID');
  const idCol = headerIndex_(header, 'Allocation ID');
  const createdCol = headerIndex_(header, 'Created Timestamp');
  const dispatchedCol = headerIndex_(header, 'Dispatched');
  const dispatchedAtCol = headerIndex_(header, 'Dispatched Timestamp');

  const existing = sh.getLastRow() < 2
    ? []
    : sh.getRange(2, 1, sh.getLastRow() - 1, header.length).getDisplayValues();

  const prev = {};

  for (let i = existing.length - 1; i >= 0; i--) {
    if (uuidCol === -1) break;
    if (normalize_(existing[i][uuidCol]) !== normalize_(context.uuid)) continue;

    if (idCol !== -1) {
      prev[normalize_(existing[i][idCol])] = {
        created: createdCol === -1 ? '' : existing[i][createdCol],
        dispatched: dispatchedCol === -1 ? '' : existing[i][dispatchedCol],
        dispatchedAt: dispatchedAtCol === -1 ? '' : existing[i][dispatchedAtCol]
      };
    }

    sh.deleteRow(i + 2);
  }

  const list = context.allocations || [];
  if (!list.length) return 0;

  const now = timestamp_();

  const rows = list.map(function(a, i) {
    const values = new Array(header.length).fill('');
    const id = allocationIdFor_(context.planId, i + 1, a.unit);
    const old = prev[id] || {};

    setArrayValue_(values, header, 'Allocation ID', id);
    setArrayValue_(values, header, 'Plan ID', context.planId);
    setArrayValue_(values, header, 'UUID', context.uuid);
    setArrayValue_(values, header, 'Order No', context.orderNo);
    setArrayValue_(values, header, 'Order Date', dateOnly_(context.orderDate));
    setArrayValue_(values, header, 'Customer', context.customer);
    setArrayValue_(
      values, header, 'EMB Sales Coordinator', context.coordinator
    );
    setArrayValue_(values, header, 'Design No', context.designNo);
    setArrayValue_(values, header, 'Colour', context.colour);
    setArrayValue_(values, header, 'Fabric', context.fabric);
    setArrayValue_(values, header, 'Order Type', context.orderType);
    setArrayValue_(
      values, header, 'Delivery Date', dateOnly_(context.deliveryDate)
    );
    setArrayValue_(values, header, 'Grid', context.grid);
    setArrayValue_(values, header, 'Revo', context.revo);
    setArrayValue_(values, header, 'Desc', context.desc);
    setArrayValue_(values, header, 'Greige/Finish', context.greigeFinish);
    setArrayValue_(values, header, 'Production Place', context.place);
    setArrayValue_(values, header, 'Production Unit', a.unit);
    setArrayValue_(values, header, 'Party Name', a.partyName || '');
    setArrayValue_(values, header, 'Patti Size', a.pattiSize);
    setArrayValue_(values, header, 'Production Qty', a.qty);
    setArrayValue_(values, header, 'Allocation Index', i + 1);
    setArrayValue_(values, header, 'Total Allocations', list.length);
    setArrayValue_(values, header, 'Dispatched', old.dispatched || 'No');
    setArrayValue_(
      values, header, 'Dispatched Timestamp', old.dispatchedAt || ''
    );
    setArrayValue_(values, header, 'Created Timestamp', old.created || now);
    setArrayValue_(values, header, 'Updated Timestamp', now);

    return values;
  });

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, header.length)
    .setValues(rows);

  return rows.length;
}

/* Columns written into each production unit's own spreadsheet. */
const PRODUCTION_UNIT_COLUMNS = [
  'Allocation ID',
  'Plan ID',
  'UUID',
  'Order No',
  'Order Date',
  'Customer',
  'EMB Sales Coordinator',
  'Design No',
  'Colour',
  'Fabric',
  'Order Type',
  'Delivery Date',
  'Grid',
  'Revo',
  'Desc',
  'Greige/Finish',
  'Production Place',
  'Production Unit',
  'Party Name',
  'Patti Size',
  'Production Qty',
  'Mending Type',
  'Flow Sequence',
  'Production Date',
  'Mending Date',
  'Process Date',
  'Final QC Date',
  'Lace Cutting Date',
  'Dispatch Date',
  'Timestamp'
];

function getOrCreateProductionUnitSheet_(spreadsheet) {
  let sh = spreadsheet.getSheetByName(APP.PRODUCTION_UNIT_SHEET_NAME);

  if (!sh) {
    sh = spreadsheet.insertSheet(APP.PRODUCTION_UNIT_SHEET_NAME);
    sh.getRange(1, 1, 1, PRODUCTION_UNIT_COLUMNS.length)
      .setValues([PRODUCTION_UNIT_COLUMNS]);
    sh.setFrozenRows(1);
    return sh;
  }

  const header = getHeader_(sh);

  if (!header.length || header.every(function(v) { return !v; })) {
    sh.getRange(1, 1, 1, PRODUCTION_UNIT_COLUMNS.length)
      .setValues([PRODUCTION_UNIT_COLUMNS]);
    sh.setFrozenRows(1);
  }

  return sh;
}

/*
 * Pushes every allocation into the spreadsheet of its production unit
 * (Schiffli / Multi / Outside). Rows for the same order are replaced so a
 * re-send never duplicates work orders.
 */
function dispatchAllocations_(context) {
  const byUnit = {};

  (context.allocations || []).forEach(function(a, i) {
    if (!byUnit[a.unit]) byUnit[a.unit] = [];
    byUnit[a.unit].push({ item: a, index: i + 1 });
  });

  const sent = [];

  Object.keys(byUnit).forEach(function(unit) {
    const fileId = APP.PRODUCTION_UNITS[unit];

    if (!fileId) {
      throw new Error('No spreadsheet configured for production unit ' + unit);
    }

    const target = SpreadsheetApp.openById(fileId);
    const sh = getOrCreateProductionUnitSheet_(target);
    const header = getHeader_(sh);
    const uuidCol = headerIndex_(header, 'UUID');

    /* Remove earlier rows for this order so the unit sees only the latest. */
    if (uuidCol !== -1 && sh.getLastRow() > 1) {
      const values = sh
        .getRange(2, 1, sh.getLastRow() - 1, header.length)
        .getDisplayValues();

      for (let i = values.length - 1; i >= 0; i--) {
        if (normalize_(values[i][uuidCol]) === normalize_(context.uuid)) {
          sh.deleteRow(i + 2);
        }
      }
    }

    const now = timestamp_();

    const rows = byUnit[unit].map(function(entry) {
      const a = entry.item;
      const values = new Array(header.length).fill('');

      setArrayValue_(
        values, header, 'Allocation ID',
        allocationIdFor_(context.planId, entry.index, a.unit)
      );
      setArrayValue_(values, header, 'Plan ID', context.planId);
      setArrayValue_(values, header, 'UUID', context.uuid);
      setArrayValue_(values, header, 'Order No', context.orderNo);
      setArrayValue_(values, header, 'Order Date', dateOnly_(context.orderDate));
      setArrayValue_(values, header, 'Customer', context.customer);
      setArrayValue_(
        values, header, 'EMB Sales Coordinator', context.coordinator
      );
      setArrayValue_(values, header, 'Design No', context.designNo);
      setArrayValue_(values, header, 'Colour', context.colour);
      setArrayValue_(values, header, 'Fabric', context.fabric);
      setArrayValue_(values, header, 'Order Type', context.orderType);
      setArrayValue_(
        values, header, 'Delivery Date', dateOnly_(context.deliveryDate)
      );
      setArrayValue_(values, header, 'Grid', context.grid);
      setArrayValue_(values, header, 'Revo', context.revo);
      setArrayValue_(values, header, 'Desc', context.desc);
      setArrayValue_(values, header, 'Greige/Finish', context.greigeFinish);
      setArrayValue_(values, header, 'Production Place', context.place);
      setArrayValue_(values, header, 'Production Unit', a.unit);
      setArrayValue_(values, header, 'Party Name', a.partyName || '');
      setArrayValue_(values, header, 'Patti Size', a.pattiSize);
      setArrayValue_(values, header, 'Production Qty', a.qty);
      setArrayValue_(values, header, 'Mending Type', context.mendingType || '');
      setArrayValue_(values, header, 'Flow Sequence', context.flowSequence || '');

      APP.FLOW_STEPS.forEach(function(step) {
        setArrayValue_(
          values,
          header,
          step.column,
          dateOnly_((context.dates || {})[step.column] || '')
        );
      });

      setArrayValue_(values, header, 'Timestamp', now);

      return values;
    });

    sh.getRange(sh.getLastRow() + 1, 1, rows.length, header.length)
      .setValues(rows);

    sent.push({ unit: unit, rows: rows.length });
  });

  markAllocationsDispatched_(context.uuid);

  return sent;
}

function markAllocationsDispatched_(uuid) {
  const sh = ensureAllocationSheet_();
  const header = getHeader_(sh);
  const uuidCol = headerIndex_(header, 'UUID');
  const dCol = headerIndex_(header, 'Dispatched');
  const dAtCol = headerIndex_(header, 'Dispatched Timestamp');

  if (uuidCol === -1 || dCol === -1 || sh.getLastRow() < 2) return;

  const values = sh
    .getRange(2, 1, sh.getLastRow() - 1, header.length)
    .getDisplayValues();

  const now = timestamp_();

  values.forEach(function(row, i) {
    if (normalize_(row[uuidCol]) !== normalize_(uuid)) return;

    sh.getRange(i + 2, dCol + 1).setValue('Yes');
    if (dAtCol !== -1) sh.getRange(i + 2, dAtCol + 1).setValue(now);
  });
}

function readAllocationsForUUID_(uuid) {
  ensureAllocationSheet_();

  const data = readSheet_(APP.SHEETS.ALLOCATIONS);

  return data.rows
    .filter(function(r) {
      return normalize_(r['UUID']) === normalize_(uuid);
    })
    .map(function(r) {
      return {
        unit: normalize_(r['Production Unit']),
        partyName: normalize_(r['Party Name']),
        pattiSize: normalize_(r['Patti Size']),
        qty: number_(r['Production Qty'])
      };
    });
}




/* Flow step helpers. */
function flowStepByKey_(key) {
  key = normalize_(key).toLowerCase();
  for (let i = 0; i < APP.FLOW_STEPS.length; i++) {
    if (APP.FLOW_STEPS[i].key === key) return APP.FLOW_STEPS[i];
  }
  return null;
}

/*
 * Ordered step keys. Mending sits BEFORE Process (Mill) for Pre Mending
 * and AFTER Process (Mill) for Post Mending.
 */
function flowOrderedKeys_(mendingType) {
  const keys = APP.FLOW_STEPS.map(function(step) { return step.key; });
  const isPost = normalize_(mendingType).toLowerCase().indexOf('post') === 0;

  if (!isPost) return keys;

  const without = keys.filter(function(k) { return k !== 'mending'; });
  const at = without.indexOf('process');

  if (at === -1) return keys;

  without.splice(at + 1, 0, 'mending');
  return without;
}

function sanitizeFlowSteps_(steps, mendingType) {
  const wanted = {};

  (steps || []).forEach(function(s) {
    const step = flowStepByKey_(s);
    if (step) wanted[step.key] = true;
  });

  return flowOrderedKeys_(mendingType)
    .filter(function(key) { return wanted[key]; });
}


/* Steps that still make sense given the stock already available. */
function defaultFlowStepsForStage_(stockStage) {
  let skip = 0;

  if (stockStage === 'Greige Mending' || stockStage === 'Finish Mending') {
    skip = 1;
  } else if (stockStage === 'Process') {
    skip = 2;
  } else if (stockStage === 'Lace cutting') {
    skip = 3;
  } else if (stockStage === 'QC & Dispatch') {
    skip = 5;
  }

  return APP.FLOW_STEPS
    .slice(skip)
    .map(function(step) { return step.key; });
}


/* ============================================================================
   4. READ ORDERS / G APV
   ========================================================================== */

/**
 * Reads a sheet into objects.
 *
 * Uses getDisplayValues() intentionally so Date values become strings and
 * never break google.script.run serialization.
 */
function readSheet_(sheetName) {
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    return { header: [], rows: [] };
  }

  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const header = values[0].map(normalize_);

  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};

    let hasAny = false;

    header.forEach(function(h, c) {
      if (!h) return;
      const value = row[c];
      obj[h] = value;
      if (value !== '') hasAny = true;
    });

    if (hasAny) {
      obj.__row = i + 1;
      rows.push(obj);
    }
  }

  return {
    header: header,
    rows: rows
  };
}

function findOrderByUUID_(uuid) {
  const target = normalize_(uuid);
  if (!target) return null;

  const data = readSheet_(APP.SHEETS.ORDERS);

  for (let i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i]['UUID']) === target) {
      return {
        row: data.rows[i],
        header: data.header
      };
    }
  }

  return null;
}

function findGApvByUUID_(uuid) {
  const target = normalize_(uuid);
  if (!target) return [];

  const data = readSheet_(APP.SHEETS.GAPV);

  return data.rows.filter(function(row) {
    return normalize_(row['UUID']) === target;
  });
}


/* ============================================================================
   5. LOGGING / NOTIFICATIONS
   ========================================================================== */

function log_(action, detail) {
  const sh = ensureLogSheet_();

  let detailText;
  try {
    detailText = typeof detail === 'string'
      ? detail
      : JSON.stringify(jsonSafe_(detail));
  } catch (e) {
    detailText = String(detail);
  }

  sh.appendRow([
    timestamp_(),
    action,
    detailText
  ]);
}

function notify_(uuid, orderNo, message) {
  const sh = ensureNotificationsSheet_();

  sh.appendRow([
    timestamp_(),
    uuid || '',
    orderNo || '',
    message || '',
    false
  ]);
}

function highlightRow_(sheetName, rowNumber, color) {
  const sh = getSheet_(sheetName);
  const lastCol = Math.max(sh.getLastColumn(), 1);

  sh.getRange(rowNumber, 1, 1, lastCol)
    .setBackground(color || '#f4cccc');
}


/* ============================================================================
   6. AUTHENTICATION
   ========================================================================== */

function hashPassword_(password) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password) + '::emb-order-process'
  );

  return digest.map(function(byte) {
    const b = byte < 0 ? byte + 256 : byte;
    return b.toString(16).padStart(2, '0');
  }).join('');
}

function createSession_(user) {
  const token = Utilities.getUuid();

  const session = {
    username: normalize_(user['Username']),
    name: normalize_(user['Sales Coordinator Name']),
    role: normalize_(user['Role']) || 'sales',
    department: normalize_(user['Department']),
    createdAt: Date.now()
  };

  CacheService.getScriptCache().put(
    'emb_session_' + token,
    JSON.stringify(session),
    APP.SESSION_TTL_SECONDS
  );

  return {
    token: token,
    session: session
  };
}

function requireSession_(token) {
  const cleanToken = normalize_(token);

  if (!cleanToken) {
    throw new Error('Not logged in.');
  }

  const raw = CacheService.getScriptCache()
    .get('emb_session_' + cleanToken);

  if (!raw) {
    throw new Error('Session expired. Please login again.');
  }

  return JSON.parse(raw);
}

function apiLogin(username, password) {
  return withError_(function() {
    username = normalize_(username);
    password = String(password == null ? '' : password);

    if (!username || !password) {
      return fail_('Username and password are required.');
    }

    const data = readSheet_(APP.SHEETS.USERS);

    const user = data.rows.find(function(row) {
      return normalize_(row['Username']).toLowerCase() === username.toLowerCase();
    });

    if (!user) {
      log_('LOGIN_FAILED', { username: username, reason: 'User not found' });
      return fail_('Invalid username or password.');
    }

    const storedPassword = normalize_(user['Password']);

    // Users normally contain SHA-256 hashes, but plain text is also accepted
    // so a manually typed password in the Users sheet still works.
    const matches =
      storedPassword === hashPassword_(password) ||
      storedPassword === password;

    if (!matches) {
      log_('LOGIN_FAILED', { username: username, reason: 'Wrong password' });
      return fail_('Invalid username or password.');
    }

    const activeFlag = normalize_(user['Active']);
    if (activeFlag && /^(no|n|0|false|disabled)$/i.test(activeFlag)) {
      log_('LOGIN_FAILED', { username: username, reason: 'Inactive account' });
      return fail_('This account is inactive. Contact an administrator.');
    }

    const created = createSession_(user);

    log_('LOGIN', {
      username: created.session.username,
      coordinator: created.session.name
    });

    return ok_({
      token: created.token,
      username: created.session.username,
      name: created.session.name,
      role: created.session.role,
      department: created.session.department,
      godowns: Object.keys(APP.GODOWNS)
    });
  });
}

function apiLoginJson(username, password) {
  return jsonString_(apiLogin(username, password));
}

function apiLogout(token) {
  CacheService.getScriptCache().remove('emb_session_' + normalize_(token));
  return ok_();
}

function apiLogoutJson(token) {
  return jsonString_(apiLogout(token));
}


/* ============================================================================
   7. WEB APP ENTRY POINT
   ========================================================================== */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('EMB Order Process')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/* ============================================================================
   8. SALES PERSON — NEW ORDERS
   ========================================================================== */

/**
 * Returns only orders belonging to the logged-in Sales Coordinator.
 *
 * New order criteria:
 * - Coordinator matches the logged-in user.
 * - Status is blank or "New".
 *
 * Orders that have already been sent to a godown are not shown as new.
 */
function apiGetMyNewOrders(token) {
  return withError_(function() {
    const session = requireSession_(token);

    const sh = getSheet_(APP.SHEETS.ORDERS);
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 2 || lastCol < 1) {
      return ok_({
        orders: [],
        godowns: Object.keys(APP.GODOWNS),
        user: session.name
      });
    }

    const header = getHeader_(sh);
    const coordinatorCol = headerIndex_(header, 'EMB Sales Coordinator');
    const statusCol = headerIndex_(header, 'Status');

    if (coordinatorCol === -1) {
      throw new Error(
        'Orders sheet is missing required column: EMB Sales Coordinator'
      );
    }

    const values = sh.getRange(2, 1, lastRow - 1, lastCol)
      .getDisplayValues();

    const wanted = APP.ORDER_BASE_COLUMNS.concat(
      APP.ORDER_WORKFLOW_COLUMNS
    );

    const orders = [];
    const myName = normalize_(session.name);

    const closedCol = headerIndex_(header, 'Order Closed');

    values.forEach(function(row, index) {
      const coordinator = normalize_(row[coordinatorCol]);
      const status = statusCol === -1 ? '' : normalize_(row[statusCol]);

      if (session.role !== 'admin' && coordinator !== myName) return;

      /* Closed orders never show up anywhere in the UI again. */
      if (closedCol !== -1 && isYesValue_(row[closedCol])) return;


      // Blank status is treated as a new order because Orders is imported
      // from an external API and the workflow columns may initially be blank.
      if (status && status !== APP.STATUS.NEW) return;

      const obj = {};

      wanted.forEach(function(col) {
        const idx = headerIndex_(header, col);
        if (idx !== -1) obj[col] = row[idx];
      });

      obj.__row = index + 2;
      orders.push(obj);
    });

    return ok_({
      orders: orders,
      godowns: Object.keys(APP.GODOWNS),
      user: session.name
    });
  });
}

function apiGetMyNewOrdersJson(token) {
  return jsonString_(apiGetMyNewOrders(token));
}


/* ============================================================================
   9. SALES PERSON — SEND ORDER TO GODOWN
   ========================================================================== */

/**
 * Sends the selected order to the selected godown spreadsheet.
 *
 * IMPORTANT:
 * - This function does NOT implement the godown callback/processing logic.
 * - It only writes the order to the godown's "Incoming Orders" sheet.
 * - Qty Unit is replaced by the entered External Inv Qty.
 * - A timestamp is appended.
 *
 * The godown's existing script is responsible for processing and sending
 * the resulting breakdown back into the central "G Apv" sheet.
 */
function apiSendToGodown(token, uuid, externalInvQty, godown) {
  return withError_(function() {
    const session = requireSession_(token);

    uuid = normalize_(uuid);
    godown = normalize_(godown);

    if (!uuid) throw new Error('UUID is required.');

    if (!Object.prototype.hasOwnProperty.call(APP.GODOWNS, godown)) {
      throw new Error('Invalid godown: ' + godown);
    }

    const qty = number_(externalInvQty);

    if (qty <= 0) {
      throw new Error('External Inv Qty must be greater than 0.');
    }

    const found = findOrderByUUID_(uuid);

    if (!found) {
      throw new Error('Order not found for UUID: ' + uuid);
    }

    const order = found.row;

    if (
      normalize_(order['EMB Sales Coordinator']) !==
      normalize_(session.name)
    ) {
      throw new Error('This order does not belong to the logged-in sales coordinator.');
    }

    const currentStatus = normalize_(order['Status']);

    if (
      currentStatus === APP.STATUS.SENT_TO_GODOWN ||
      currentStatus === APP.STATUS.G_APV_RECEIVED ||
      currentStatus === APP.STATUS.SENT_TO_PRODUCTION ||
      currentStatus === APP.STATUS.STOCK_ONLY
    ) {
      throw new Error(
        'This order has already moved beyond the new-order stage.'
      );
    }

    const godownSpreadsheet = SpreadsheetApp.openById(APP.GODOWNS[godown]);
    const godownSheet = getOrCreateGodownIncomingSheet_(godownSpreadsheet);

    const godownHeader = getHeader_(godownSheet);

    /*
     * Send the same 15 order fields in the same logical order.
     * Qty Unit is intentionally replaced with External Inv Qty.
     * Timestamp is appended.
     */
    const outgoing = APP.ORDER_BASE_COLUMNS.map(function(col) {
      if (col === 'Qty Unit') return qty;
      return order[col] !== undefined ? order[col] : '';
    });

    outgoing.push(timestamp_());

    // If the existing godown sheet has exactly these expected columns,
    // write the complete row. If it has extra columns, write only the
    // columns we control.
    const expectedWidth = APP.ORDER_BASE_COLUMNS.length + 1;

    if (godownHeader.length < expectedWidth) {
      throw new Error(
        'Godown sheet "' + APP.GODOWN_SHEET_NAME +
        '" does not have the expected order columns.'
      );
    }

    godownSheet.getRange(
      godownSheet.getLastRow() + 1,
      1,
      1,
      expectedWidth
    ).setValues([outgoing]);

    const ordersSheet = getSheet_(APP.SHEETS.ORDERS);
    const ordersHeader = found.header;
    const rowNumber = order.__row;

    setIfColumnExists_(
      ordersSheet,
      ordersHeader,
      rowNumber,
      'Status',
      APP.STATUS.SENT_TO_GODOWN
    );

    setIfColumnExists_(
      ordersSheet,
      ordersHeader,
      rowNumber,
      'External Inv Qty',
      qty
    );

    setIfColumnExists_(
      ordersSheet,
      ordersHeader,
      rowNumber,
      'Godown',
      godown
    );

    setIfColumnExists_(
      ordersSheet,
      ordersHeader,
      rowNumber,
      'Sent Timestamp',
      timestamp_()
    );

    // Remaining quantity is provisional until G Apv confirms the stage
    // breakdown. It is not calculated from the sales person's entered
    // external quantity here.
    setIfColumnExists_(
      ordersSheet,
      ordersHeader,
      rowNumber,
      'Remaining Qty',
      number_(order['Order QTY']) - qty
    );

    log_('SEND_TO_GODOWN', {
      uuid: uuid,
      orderNo: order['Order No'],
      coordinator: session.name,
      godown: godown,
      externalInvQty: qty
    });

    return ok_({
      uuid: uuid,
      orderNo: order['Order No'],
      godown: godown,
      externalInvQty: qty,
      status: APP.STATUS.SENT_TO_GODOWN
    });
  });
}

function apiSendToGodownJson(token, uuid, externalInvQty, godown) {
  return jsonString_(apiSendToGodown(
    token,
    uuid,
    externalInvQty,
    godown
  ));
}

function getOrCreateGodownIncomingSheet_(godownSpreadsheet) {
  let sh = godownSpreadsheet.getSheetByName(APP.GODOWN_SHEET_NAME);

  if (!sh) {
    sh = godownSpreadsheet.insertSheet(APP.GODOWN_SHEET_NAME);

    sh.getRange(
      1,
      1,
      1,
      APP.ORDER_BASE_COLUMNS.length + 1
    ).setValues([
      APP.ORDER_BASE_COLUMNS.concat(['Timestamp'])
    ]);

    sh.setFrozenRows(1);
  }

  return sh;
}


/* ============================================================================
   10. GODOWN CALLBACK — RECEIVE G APV
   ========================================================================== */

/**
 * OPTIONAL HTTP CALLBACK ENTRY.
 *
 * The user's godown spreadsheets already have their own script. If that
 * existing script POSTs back to this web app, it can use:
 *
 * {
 *   "action": "receiveGodownData",
 *   "secret": "...",
 *   "godown": "All Over",
 *   "rows": [
 *      {
 *        "uuid": "...",
 *        "orderNo": "...",
 *        "greigeMending": 50,
 *        "finishMending": 100,
 *        "process": 50,
 *        "laceCutting": 100,
 *        "qcDispatch": 100,
 *        "timestamp": "..."
 *      }
 *   ]
 * }
 *
 * HOWEVER:
 * The central "G Apv" sheet may also already be populated by the existing
 * godown system. The review APIs below read that sheet directly.
 */
function doPost(e) {
  try {
    const payload = JSON.parse(
      e && e.postData && e.postData.contents
        ? e.postData.contents
        : '{}'
    );

    if (payload.action === 'receiveGodownData') {
      return ContentService
        .createTextOutput(jsonString_(receiveGodownData_(payload)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(jsonString_(
        fail_('Unknown action: ' + normalize_(payload.action))
      ))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(jsonString_(
        fail_(err && err.message ? err.message : String(err))
      ))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles a callback only when the godown system actually posts to this
 * endpoint. It is not required for reading an already-populated G Apv sheet.
 */
function receiveGodownData_(payload) {
  const secret = getSharedSecret_();

  if (!secret || normalize_(payload.secret) !== secret) {
    return fail_('Unauthorized callback.');
  }

  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const godown = normalize_(payload.godown);

  if (!Object.prototype.hasOwnProperty.call(APP.GODOWNS, godown)) {
    return fail_('Unknown godown: ' + godown);
  }

  const results = [];

  rows.forEach(function(row) {
    try {
      results.push(upsertGApvRowFromCallback_(godown, row));
    } catch (err) {
      results.push({
        uuid: normalize_(row && row.uuid),
        ok: false,
        error: err.message || String(err)
      });
    }
  });

  return ok_({ results: results });
}

function upsertGApvRowFromCallback_(godown, row) {
  if (!row || !normalize_(row.uuid)) {
    throw new Error('UUID is required.');
  }

  const uuid = normalize_(row.uuid);
  const orderFound = findOrderByUUID_(uuid);

  if (!orderFound) {
    throw new Error('Order not found for UUID: ' + uuid);
  }

  const order = orderFound.row;

  const externalInvQty = number_(row.externalInvQty);

  const stage = {
    'Greige Mending': number_(row.greigeMending),
    'Finish Mending': number_(row.finishMending),
    'Process': number_(row.process),
    'Lace cutting': number_(row.laceCutting),
    'QC & Dispatch': number_(row.qcDispatch)
  };

  const sum =
    stage['Greige Mending'] +
    stage['Finish Mending'] +
    stage['Process'] +
    stage['Lace cutting'] +
    stage['QC & Dispatch'];

  const match = sum === externalInvQty;

  const gapv = getSheet_(APP.SHEETS.GAPV);
  const header = getHeader_(gapv);

  ensureGApvColumns_(gapv);

  const refreshedHeader = getHeader_(gapv);
  const existing = findGApvRowsByUUID_(gapv, uuid);

  const rowValues = buildGApvRow_(
    refreshedHeader,
    order,
    godown,
    row,
    stage,
    externalInvQty,
    sum,
    match
  );

  let rowNumber;

  if (existing.length) {
    rowNumber = existing[0];
    gapv.getRange(rowNumber, 1, 1, refreshedHeader.length)
      .setValues([rowValues]);
  } else {
    rowNumber = gapv.getLastRow() + 1;
    gapv.getRange(rowNumber, 1, 1, refreshedHeader.length)
      .setValues([rowValues]);
  }

  finalizeGApvResult_(
    uuid,
    order,
    orderFound.header,
    rowNumber,
    sum,
    externalInvQty,
    match
  );

  return {
    uuid: uuid,
    ok: true,
    match: match,
    sum: sum,
    externalInvQty: match ? externalInvQty : sum
  };
}


/* ============================================================================
   11. G APV — ACTUAL SHEET STRUCTURE
   ========================================================================== */

/**
 * IMPORTANT:
 * We DO NOT rewrite the user's existing G Apv headers.
 *
 * We only add missing callback/helper columns if needed.
 *
 * The expected current structure from the provided workbook is:
 * 15 order columns + Timestamp + 5 stage columns.
 */
function ensureGApvColumns_(sh) {
  const header = getHeader_(sh);

  if (!header.length) {
    sh.getRange(
      1,
      1,
      1,
      APP.ORDER_BASE_COLUMNS.length + 1 + APP.GAPV_STAGE_COLUMNS.length
    ).setValues([[
      ...APP.ORDER_BASE_COLUMNS,
      'Timestamp',
      ...APP.GAPV_STAGE_COLUMNS
    ]]);
    return;
  }

  // Never clear/rewrite existing G Apv data.
  const missing = [];

  APP.ORDER_BASE_COLUMNS.concat([
    'Timestamp'
  ], APP.GAPV_STAGE_COLUMNS).forEach(function(col) {
    if (headerIndex_(header, col) === -1) {
      missing.push(col);
    }
  });

  if (missing.length) {
    sh.getRange(
      1,
      sh.getLastColumn() + 1,
      1,
      missing.length
    ).setValues([missing]);
  }
}

function buildGApvRow_(
  header,
  order,
  godown,
  callbackRow,
  stage,
  externalInvQty,
  sum,
  match
) {
  const row = new Array(header.length).fill('');

  APP.ORDER_BASE_COLUMNS.forEach(function(col) {
    const idx = headerIndex_(header, col);
    if (idx !== -1) {
      row[idx] = order[col] !== undefined ? order[col] : '';
    }
  });

  const timestampIndex = headerIndex_(header, 'Timestamp');

  if (timestampIndex !== -1) {
    row[timestampIndex] =
      callbackRow.timestamp || timestamp_();
  }

  APP.GAPV_STAGE_COLUMNS.forEach(function(col) {
    const idx = headerIndex_(header, col);

    if (idx !== -1) {
      row[idx] = stage[col];
    }
  });

  // These helper columns are supported if they exist in the G Apv sheet.
  setArrayValue_(row, header, 'Godown', godown);
  setArrayValue_(row, header, 'External Inv Qty', match ? externalInvQty : sum);
  setArrayValue_(row, header, 'Sum', sum);
  setArrayValue_(row, header, 'Match', match ? 'MATCH' : 'MISMATCH');

  return row;
}

function setArrayValue_(array, header, column, value) {
  const idx = headerIndex_(header, column);
  if (idx !== -1) array[idx] = value;
}

function findGApvRowsByUUID_(sh, uuid) {
  const header = getHeader_(sh);
  const uuidCol = headerIndex_(header, 'UUID');

  if (uuidCol === -1 || sh.getLastRow() < 2) return [];

  const values = sh.getRange(
    2,
    uuidCol + 1,
    sh.getLastRow() - 1,
    1
  ).getDisplayValues();

  const rows = [];

  values.forEach(function(v, i) {
    if (normalize_(v[0]) === normalize_(uuid)) {
      rows.push(i + 2);
    }
  });

  return rows;
}

function finalizeGApvResult_(
  uuid,
  order,
  orderHeader,
  gapvRowNumber,
  sum,
  externalInvQty,
  match
) {
  const effectiveExternalQty = match ? externalInvQty : sum;

  const ordersSheet = getSheet_(APP.SHEETS.ORDERS);
  const orderRow = order.__row;

  setIfColumnExists_(
    ordersSheet,
    orderHeader,
    orderRow,
    'Status',
    APP.STATUS.G_APV_RECEIVED
  );

  setIfColumnExists_(
    ordersSheet,
    orderHeader,
    orderRow,
    'External Inv Qty',
    effectiveExternalQty
  );

  setIfColumnExists_(
    ordersSheet,
    orderHeader,
    orderRow,
    'Remaining Qty',
    Math.max(
      number_(order['Order QTY']) - effectiveExternalQty,
      0
    )
  );

  if (!match) {
    highlightRow_(
      APP.SHEETS.GAPV,
      gapvRowNumber,
      '#f4cccc'
    );

    notify_(
      uuid,
      order['Order No'],
      'External inventory mismatch. Entered External Inv Qty: ' +
      externalInvQty +
      ', stage total: ' +
      sum +
      '. External Inv Qty was changed to ' +
      sum +
      '.'
    );

    log_('G_APV_MISMATCH', {
      uuid: uuid,
      orderNo: order['Order No'],
      enteredExternalInvQty: externalInvQty,
      stageSum: sum,
      correctedExternalInvQty: sum
    });
  } else {
    log_('G_APV_MATCH', {
      uuid: uuid,
      orderNo: order['Order No'],
      externalInvQty: externalInvQty,
      stageSum: sum
    });
  }
}


/* ============================================================================
   12. REVIEW / REMAINING QTY FOR SALES PERSON
   ========================================================================== */

/**
 * This is the main Review API.
 *
 * It reads:
 * - Orders for customer/order master data
 * - G Apv for returned godown stage data
 *
 * It joins the two using UUID.
 *
 * No dependency on a Status value in Orders is required. This is important
 * because G Apv itself is the authoritative indication that godown data has
 * returned.
 */
function apiGetOrdersAwaitingDecision(token) {
  return withError_(function() {
    const session = requireSession_(token);

    const ordersData = readSheet_(APP.SHEETS.ORDERS);
    const gapvData = readSheet_(APP.SHEETS.GAPV);

    const ordersByUuid = {};

    ordersData.rows.forEach(function(order) {
      const uuid = normalize_(order['UUID']);
      if (uuid) {
        ordersByUuid[uuid] = order;
      }
    });

    const reviewOrders = [];

    gapvData.rows.forEach(function(gapv) {
      const uuid = normalize_(gapv['UUID']);
      if (!uuid) return;

      const order = ordersByUuid[uuid];
      if (!order) return;

      /* Closed orders are hidden everywhere. */
      if (isOrderClosed_(order)) return;

      /*
       * Once a decision (stock only / send to production) has been made for
       * this order, it is no longer "awaiting decision" — it belongs in the
       * Production Plan / Process Board / Review Orders views instead, not
       * back here. Without this check a decided order kept reappearing in
       * this list forever because its G Apv data still exists.
       */
      if (normalize_(order['Production Decision'])) return;
      const decidedStatus =
        normalize_(order['Status']) === APP.STATUS.STOCK_ONLY ||
        normalize_(order['Status']) === APP.STATUS.SENT_TO_PRODUCTION;
      if (decidedStatus) return;


      // Security: a salesperson sees only their own orders.
      if (
        session.role !== 'admin' &&
        normalize_(order['EMB Sales Coordinator']) !==
        normalize_(session.name)
      ) {
        return;
      }

      const stages = {};
      let stageSum = 0;

      APP.GAPV_STAGE_COLUMNS.forEach(function(stage) {
        stages[stage] = number_(gapv[stage]);
        stageSum += stages[stage];
      });

      /*
       * User's G Apv structure has no separate External Inv Qty column.
       * Therefore the returned godown inventory is the sum of the 5 stages.
       */
      const externalInvQty = stageSum;

      const orderQty = number_(order['Order QTY']);
      const remainingQty = Math.max(
        orderQty - externalInvQty,
        0
      );

      const review = {};

      APP.ORDER_BASE_COLUMNS.concat(APP.ORDER_WORKFLOW_COLUMNS)
        .forEach(function(col) {
          if (order[col] !== undefined) {
            review[col] = order[col];
          }
        });

      review['Status'] = APP.STATUS.G_APV_RECEIVED;
      review['External Inv Qty'] = externalInvQty;
      review['Remaining Qty'] = remainingQty;

      review['Godown'] =
        normalize_(order['Godown']) ||
        normalize_(gapv['Godown']) ||
        normalize_(gapv['Godown Name']);

      review['gapv'] = Object.assign({}, gapv, {
        'External Inv Qty': externalInvQty,
        'Sum': stageSum,
        'Match': 'MATCH'
      });

      review['Stock Current Stage'] =
        determineStockStageFromRows_([gapv]);

      reviewOrders.push(review);
    });

    return ok_({
      orders: reviewOrders,
      user: session.name
    });
  });
}

function apiGetOrdersAwaitingDecisionJson(token) {
  return jsonString_(apiGetOrdersAwaitingDecision(token));
}


/* ============================================================================
   13. STOCK STAGE
   ========================================================================== */

/**
 * Determines the furthest stage reached by existing stock.
 *
 * Example:
 * Mending > 0
 * Process = 0
 * Lace cutting = 0
 * QC & Dispatch = 0
 *
 * => "Mending"
 *
 * If Process > 0:
 * => "Process"
 *
 * If QC & Dispatch > 0:
 * => "QC & Dispatch"
 */
function determineStockStage_(uuid) {
  const rows = findGApvByUUID_(uuid);
  return determineStockStageFromRows_(rows);
}

function determineStockStageFromRows_(rows) {
  if (!rows || !rows.length) return '';

  let furthest = -1;
  let stageName = '';

  rows.forEach(function(row) {
    APP.GAPV_STAGE_COLUMNS.forEach(function(stage, index) {
      if (number_(row[stage]) > 0 && index > furthest) {
        furthest = index;
        stageName = stage;
      }
    });
  });

  return stageName;
}


/* ============================================================================
   14. SALES PERSON — PRODUCTION DECISION
   ========================================================================== */

function apiDecideProduction(
  token,
  uuid,
  decision,
  productionInfo
) {
  return withError_(function() {
    const session = requireSession_(token);

    uuid = normalize_(uuid);
    decision = normalize_(decision).toLowerCase();
    productionInfo = productionInfo || {};

    if (!uuid) {
      throw new Error('UUID is required.');
    }

    if (
      decision !== 'stock_only' &&
      decision !== 'send_to_production'
    ) {
      throw new Error(
        'Decision must be stock_only or send_to_production.'
      );
    }

    const found = findOrderByUUID_(uuid);

    if (!found) {
      throw new Error('Order not found for UUID: ' + uuid);
    }

    const order = found.row;

    if (
      normalize_(order['EMB Sales Coordinator']) !==
      normalize_(session.name)
    ) {
      throw new Error('This order does not belong to you.');
    }

    const gapvRows = findGApvByUUID_(uuid);

    if (!gapvRows.length) {
      throw new Error(
        'G Apv data has not returned for this order yet.'
      );
    }

    const stageSum = gapvRows.reduce(function(total, row) {
      return total + APP.GAPV_STAGE_COLUMNS.reduce(function(sum, stage) {
        return sum + number_(row[stage]);
      }, 0);
    }, 0);

    const remainingQty = Math.max(
      number_(order['Order QTY']) - stageSum,
      0
    );

    const ordersSheet = getSheet_(APP.SHEETS.ORDERS);
    const header = found.header;
    const rowNumber = order.__row;

    if (decision === 'stock_only') {
      setIfColumnExists_(
        ordersSheet,
        header,
        rowNumber,
        'Status',
        APP.STATUS.STOCK_ONLY
      );

      setIfColumnExists_(
        ordersSheet,
        header,
        rowNumber,
        'Remaining Qty',
        remainingQty
      );

      setIfColumnExists_(
        ordersSheet,
        header,
        rowNumber,
        'Production Decision',
        'No production needed'
      );

      log_('PRODUCTION_DECISION', {
        uuid: uuid,
        orderNo: order['Order No'],
        decision: 'stock_only',
        remainingQty: remainingQty,
        by: session.username
      });

      return ok_({
        uuid: uuid,
        decision: decision,
        remainingQty: remainingQty
      });
    }

    /*
     * Send to production.
     *
     * Required information:
     * Grid
     * Revo
     * Desc
     * Greige/Finish
     */
    const grid = normalize_(productionInfo.grid);
    const revo = normalize_(productionInfo.revo);
    const desc = normalize_(productionInfo.desc);
    const greigeFinish = normalize_(productionInfo.greigeFinish);

    if (!grid || !revo || !desc || !greigeFinish) {
      throw new Error(
        'Grid, Revo, Desc and Greige/Finish are required.'
      );
    }

    const stockStage = determineStockStageFromRows_(gapvRows);

    const mendingType = normalize_(productionInfo.mendingType);

    const flowSteps = sanitizeFlowSteps_(
      productionInfo.flowSteps,
      mendingType
    );

    if (!flowSteps.length) {
      throw new Error('Select at least one production flow step.');
    }

    if (flowSteps.indexOf('mending') !== -1 && !mendingType) {
      throw new Error('Select Pre Mending or Post Mending.');
    }

    const dates = calculateLifecycleDates_(
      stockStage,
      flowSteps,
      mendingType
    );


    /* Dates supplied by the user win over the calculated defaults. */
    const givenDates = productionInfo.dates || {};

    APP.FLOW_STEPS.forEach(function(step) {
      if (flowSteps.indexOf(step.key) === -1) {
        dates[step.column] = '';
        return;
      }
      if (givenDates[step.column]) {
        dates[step.column] = normalize_(givenDates[step.column]);
      }
    });

    /*
     * Where will the production be made, and how is the qty split across
     * machines / outside parties (each with its own patti size)?
     */
    const split = sanitizeAllocations_(
      productionInfo.productionPlace,
      productionInfo.allocations,
      remainingQty
    );

    setIfColumnExists_(
      ordersSheet,
      header,
      rowNumber,
      'Status',
      APP.STATUS.SENT_TO_PRODUCTION
    );

    setIfColumnExists_(
      ordersSheet,
      header,
      rowNumber,
      'Remaining Qty',
      remainingQty
    );

    setIfColumnExists_(
      ordersSheet,
      header,
      rowNumber,
      'Production Decision',
      'Sent to production — ' + split.place
    );

    const planRow = upsertProductionPlan_({
      uuid: uuid,
      orderNo: order['Order No'],
      grid: grid,
      revo: revo,
      desc: desc,
      greigeFinish: greigeFinish,
      remainingQty: remainingQty,
      stockStage: stockStage,
      dates: dates,
      flowSteps: flowSteps,
      mendingType: mendingType,
      place: split.place,
      allocations: split.allocations,
      order: order

    });

    /* Push the work orders into each production unit's own spreadsheet. */
    const dispatched = dispatchAllocations_({
      planId: planIdFor_(uuid, order['Order No']),
      uuid: uuid,
      orderNo: order['Order No'],
      orderDate: order['Date'] || '',
      customer: order['Customer'] || '',
      coordinator: order['EMB Sales Coordinator'] || '',
      designNo: order['Design No'] || '',
      colour: order['Colour'] || '',
      fabric: order['Fabric'] || '',
      orderType: order['Order Type'] || '',
      deliveryDate: order['Delivery Date'] || '',
      grid: grid,
      revo: revo,
      desc: desc,
      greigeFinish: greigeFinish,
      place: split.place,
      mendingType: mendingType,
      flowSequence: flowSteps.map(function(k) {
        const s = flowStepByKey_(k);
        return s ? s.label : k;
      }).join(' -> '),
      dates: dates,
      allocations: split.allocations
    });

    log_('PRODUCTION_DECISION', {
      uuid: uuid,
      orderNo: order['Order No'],
      decision: 'send_to_production',
      remainingQty: remainingQty,
      stockStage: stockStage,
      grid: grid,
      revo: revo,
      desc: desc,
      greigeFinish: greigeFinish,
      flowSteps: flowSteps,
      mendingType: mendingType,
      productionPlace: split.place,
      allocations: split.allocations,
      dispatched: dispatched,
      dates: dates,
      by: session.username
    });

    return ok_({
      uuid: uuid,
      decision: decision,
      remainingQty: remainingQty,
      stockStage: stockStage,
      flowSteps: flowSteps,
      mendingType: mendingType,
      productionPlace: split.place,
      allocations: split.allocations,
      dispatched: dispatched,
      dates: dates,
      planRow: planRow
    });

  });
}

function apiDecideProductionJson(
  token,
  uuid,
  decision,
  productionInfo
) {
  return jsonString_(apiDecideProduction(
    token,
    uuid,
    decision,
    productionInfo
  ));
}


/* ============================================================================
   15. PRODUCTION LIFECYCLE PLANNING
   ========================================================================== */

/**
 * Lifecycle:
 *
 * Production
 * Pre Mending / Post Mending
 * Process (Mill)
 * Lace cutting
 * Final QC
 * Dispatch / Stock
 *
 * Planning API stores milestone completion dates as:
 *
 * Production Date
 * Mending Date
 * Process Date
 * Dispatch Date
 *
 * Existing external stock determines where planning starts.
 *
 * No stock:
 *   Production -> Mending -> Process -> Dispatch
 *
 * Stock at Mending:
 *   Mending -> Process -> Dispatch
 *
 * Stock at Process:
 *   Process -> Dispatch
 *
 * Stock at Lace cutting / QC:
 *   Dispatch
 *
 * The exact user-facing lifecycle labels can be rendered by Index.html.
 */
function calculateLifecycleDates_(stockStage, flowSteps, mendingType) {
  const dates = {};

  APP.FLOW_STEPS.forEach(function(step) {
    dates[step.column] = '';
  });

  const steps = (flowSteps && flowSteps.length)
    ? sanitizeFlowSteps_(flowSteps, mendingType)
    : sanitizeFlowSteps_(
        defaultFlowStepsForStage_(stockStage),
        mendingType
      );


  const leadDays = {
    production: APP.LEAD_DAYS.PRODUCTION,
    mending: APP.LEAD_DAYS.MENDING,
    process: APP.LEAD_DAYS.PROCESS,
    qc: APP.LEAD_DAYS.QC,
    lace: APP.LEAD_DAYS.LACE,
    dispatch: APP.LEAD_DAYS.DISPATCH
  };

  let cursor = new Date();

  steps.forEach(function(key) {
    const step = flowStepByKey_(key);
    if (!step) return;

    cursor = new Date(
      cursor.getTime() + (leadDays[key] || 1) * 24 * 60 * 60 * 1000
    );

    dates[step.column] = Utilities.formatDate(
      cursor,
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'yyyy-MM-dd'
    );
  });

  return dates;
}

/* ---------------------------------------------------------------------------
   PLAN INDEXING HELPERS
   -------------------------------------------------------------------------- */

function pad_(n, size) {
  let s = String(n);
  while (s.length < size) s = '0' + s;
  return s;
}

function planIdFor_(uuid, orderNo) {
  const clean = normalize_(uuid).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  return 'PLN-' +
    (normalize_(orderNo) || 'NA') + '-' +
    (clean ? clean.substring(0, 8) : 'NOUUID');
}

function stepIdFor_(planId, index, stepKey) {
  return planId + '-S' + pad_(index, 2) + '-' + String(stepKey).toUpperCase();
}

function dateOnly_(value) {
  const v = normalize_(value);
  if (!v) return '';

  const d = new Date(v);
  if (isNaN(d.getTime())) return v;

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone() || 'Asia/Kolkata',
    'yyyy-MM-dd'
  );
}

function daysBetween_(from, to) {
  if (!from || !to) return '';

  const a = new Date(from);
  const b = new Date(to);

  if (isNaN(a.getTime()) || isNaN(b.getTime())) return '';

  return Math.round((b.getTime() - a.getTime()) / (24 * 60 * 60 * 1000));
}

/*
 * Builds the fully indexed step list for a plan: every selected step in the
 * order it will actually be performed, with its planned date.
 */
function buildPlanStepRecords_(flowSteps, mendingType, dates) {
  const ordered = sanitizeFlowSteps_(flowSteps, mendingType);

  return ordered.map(function(key, i) {
    const step = flowStepByKey_(key);

    return {
      index: i + 1,
      key: key,
      name: step ? step.label : key,
      column: step ? step.column : '',
      date: dateOnly_((dates || {})[step ? step.column : ''] || ''),
      prev: i === 0 ? '' : (flowStepByKey_(ordered[i - 1]) || {}).label || '',
      next: i === ordered.length - 1
        ? ''
        : (flowStepByKey_(ordered[i + 1]) || {}).label || '',
      isFirst: i === 0,
      isLast: i === ordered.length - 1,
      total: ordered.length
    };
  });
}

/*
 * Rewrites the Plan Steps rows for one order. Old rows for the same UUID are
 * deleted first, so the sheet always holds exactly one clean record per
 * order x step - ready for pivots / charts.
 */
function syncPlanSteps_(context) {
  const sh = ensurePlanStepsSheet_();
  const header = getHeader_(sh);
  const uuidCol = headerIndex_(header, 'UUID');
  const createdCol = headerIndex_(header, 'Created Timestamp');
  const stepIdCol = headerIndex_(header, 'Step ID');

  const existing = sh.getLastRow() < 2
    ? []
    : sh.getRange(2, 1, sh.getLastRow() - 1, header.length).getDisplayValues();

  const createdMap = {};
  const statusMap = {};

  /* Delete previous rows for this UUID (bottom-up keeps indexes valid). */
  for (let i = existing.length - 1; i >= 0; i--) {
    if (uuidCol === -1) break;
    if (normalize_(existing[i][uuidCol]) !== normalize_(context.uuid)) continue;

    if (stepIdCol !== -1) {
      const id = normalize_(existing[i][stepIdCol]);

      if (createdCol !== -1) createdMap[id] = existing[i][createdCol];

      statusMap[id] = {
        key: normalize_(readCell_(existing[i], header, 'Step Key')),
        status: normalize_(readCell_(existing[i], header, 'Step Status')),
        completedDate: normalize_(readCell_(existing[i], header, 'Completed Date')),
        completedBy: normalize_(readCell_(existing[i], header, 'Completed By'))
      };
    }

    sh.deleteRow(i + 2);
  }

  /* Completion carries over by step key, so re-planning never loses it. */
  const doneByKey = {};

  Object.keys(statusMap).forEach(function(id) {
    const rec = statusMap[id];
    if (/completed/i.test(rec.status) && rec.key) doneByKey[rec.key] = rec;
  });

  if (context.completedKeys) {
    context.completedKeys.forEach(function(k) {
      if (!doneByKey[k]) {
        doneByKey[k] = { status: APP.STEP_STATUS.COMPLETED, completedDate: '', completedBy: '' };
      }
    });
  }

  const steps = context.steps || [];
  if (!steps.length) {
    return { count: 0, currentIndex: '', currentKey: '', currentName: '',
      currentDate: '', completed: [], pending: [], allDone: false };
  }

  const now = timestamp_();
  const start = steps[0].date;

  /* First step that is not completed is the CURRENT step. */
  let currentIndex = 0;

  for (let i = 0; i < steps.length; i++) {
    if (!doneByKey[steps[i].key]) { currentIndex = i + 1; break; }
  }

  const completed = [];
  const pending = [];

  const rows = steps.map(function(step) {
    const values = new Array(header.length).fill('');
    const stepId = stepIdFor_(context.planId, step.index, step.key);
    const done = doneByKey[step.key];
    const isCurrent = step.index === currentIndex;

    const status = done
      ? APP.STEP_STATUS.COMPLETED
      : (isCurrent ? APP.STEP_STATUS.CURRENT : APP.STEP_STATUS.PENDING);

    if (done) completed.push(step.name); else pending.push(step.name);

    setArrayValue_(values, header, 'Step ID', stepId);
    setArrayValue_(values, header, 'Plan ID', context.planId);
    setArrayValue_(values, header, 'UUID', context.uuid);
    setArrayValue_(values, header, 'Order No', context.orderNo);
    setArrayValue_(values, header, 'Order Date', dateOnly_(context.orderDate));
    setArrayValue_(values, header, 'Customer', context.customer);
    setArrayValue_(
      values, header, 'EMB Sales Coordinator', context.coordinator
    );
    setArrayValue_(values, header, 'Design No', context.designNo);
    setArrayValue_(values, header, 'Colour', context.colour);
    setArrayValue_(values, header, 'Fabric', context.fabric);
    setArrayValue_(values, header, 'Order Type', context.orderType);
    setArrayValue_(
      values, header, 'Delivery Date', dateOnly_(context.deliveryDate)
    );
    setArrayValue_(values, header, 'Grid', context.grid);
    setArrayValue_(values, header, 'Revo', context.revo);
    setArrayValue_(values, header, 'Desc', context.desc);
    setArrayValue_(values, header, 'Greige/Finish', context.greigeFinish);
    setArrayValue_(values, header, 'Remaining Qty', context.remainingQty);
    setArrayValue_(values, header, 'Stock Current Stage', context.stockStage);
    setArrayValue_(values, header, 'Mending Type', context.mendingType);
    setArrayValue_(values, header, 'Step Index', step.index);
    setArrayValue_(values, header, 'Step Key', step.key);
    setArrayValue_(values, header, 'Step Name', step.name);
    setArrayValue_(values, header, 'Planned Date', step.date);
    setArrayValue_(values, header, 'Previous Step', step.prev);
    setArrayValue_(values, header, 'Next Step', step.next);
    setArrayValue_(values, header, 'Is First Step', step.isFirst ? 'Yes' : 'No');
    setArrayValue_(values, header, 'Is Last Step', step.isLast ? 'Yes' : 'No');
    setArrayValue_(values, header, 'Total Steps', step.total);
    setArrayValue_(
      values, header, 'Days From Plan Start', daysBetween_(start, step.date)
    );
    setArrayValue_(values, header, 'Production Place', context.place || '');
    setArrayValue_(values, header, 'Step Status', status);
    setArrayValue_(values, header, 'Is Current Step', isCurrent ? 'Yes' : 'No');
    setArrayValue_(
      values, header, 'Completed Date', done ? (done.completedDate || dateOnly_(now)) : ''
    );
    setArrayValue_(
      values, header, 'Completed By',
      done ? (done.completedBy || context.completedBy || '') : ''

    );
    setArrayValue_(
      values, header, 'Created Timestamp', createdMap[stepId] || now
    );
    setArrayValue_(values, header, 'Updated Timestamp', now);

    return values;
  });

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, header.length)
    .setValues(rows);

  const current = currentIndex
    ? steps[currentIndex - 1]
    : null;

  return {
    count: rows.length,
    currentIndex: currentIndex || '',
    currentKey: current ? current.key : '',
    currentName: current ? current.name : '',
    currentDate: current ? current.date : '',
    completed: completed,
    pending: pending,
    allDone: !currentIndex
  };
}

/* Reads a value from a raw display-values row by column name. */
function readCell_(row, header, column) {
  const idx = headerIndex_(header, column);
  return idx === -1 ? '' : row[idx];
}


function upsertProductionPlan_(data) {
  const sh = ensurePlanSheet_();
  const header = getHeader_(sh);

  const rows = sh.getLastRow() < 2
    ? []
    : sh.getRange(
        2,
        1,
        sh.getLastRow() - 1,
        sh.getLastColumn()
      ).getDisplayValues();

  const uuidCol = headerIndex_(header, 'UUID');
  const createdCol = headerIndex_(header, 'Created Timestamp');
  const planIdCol = headerIndex_(header, 'Plan ID');
  const lifeCol = headerIndex_(header, 'Lifecycle Status');

  let targetRow = -1;
  let createdAt = '';
  let planId = '';
  let lifecycle = '';

  if (uuidCol !== -1) {
    for (let i = 0; i < rows.length; i++) {
      if (normalize_(rows[i][uuidCol]) === normalize_(data.uuid)) {
        targetRow = i + 2;
        if (createdCol !== -1) createdAt = rows[i][createdCol];
        if (planIdCol !== -1) planId = normalize_(rows[i][planIdCol]);
        if (lifeCol !== -1) lifecycle = normalize_(rows[i][lifeCol]);
        break;
      }
    }
  }

  lifecycle = normalize_(data.lifecycle) || lifecycle || APP.LIFECYCLE.PLANNED;


  if (!planId) planId = planIdFor_(data.uuid, data.orderNo);

  const order = data.order || {};

  const steps = buildPlanStepRecords_(
    data.flowSteps,
    data.mendingType,
    data.dates
  );

  const planStart = steps.length ? steps[0].date : '';
  const planEnd = steps.length ? steps[steps.length - 1].date : '';

  const context = {
    planId: planId,
    uuid: data.uuid,
    orderNo: data.orderNo,
    orderDate: order['Date'] || '',
    customer: order['Customer'] || '',
    coordinator: order['EMB Sales Coordinator'] || '',
    designNo: order['Design No'] || '',
    colour: order['Colour'] || '',
    fabric: order['Fabric'] || '',
    orderType: order['Order Type'] || '',
    deliveryDate: order['Delivery Date'] || '',
    grid: data.grid,
    revo: data.revo,
    desc: data.desc,
    greigeFinish: data.greigeFinish,
    remainingQty: data.remainingQty,
    stockStage: data.stockStage || '',
    mendingType: data.mendingType || '',
    place: normalize_(data.place || ''),
    lifecycle: lifecycle,
    completedKeys: data.completedKeys || null,
    completedBy: data.completedBy || '',

    steps: steps
  };

  /* Step index first: it tells us which step the order is currently on. */
  const stepInfo = syncPlanSteps_(context);

  /* Machine / party split. */
  const allocations = data.allocations || [];

  if (allocations.length) {
    syncAllocations_(
      Object.assign({}, context, { allocations: allocations })
    );
  }

  const allocList = allocations.length
    ? allocations
    : readAllocationsForUUID_(data.uuid);

  const allocatedQty = allocList.reduce(function(sum, a) {
    return sum + number_(a.qty);
  }, 0);

  const parties = allocList
    .filter(function(a) { return a.unit === 'Outside' && a.partyName; })
    .map(function(a) { return a.partyName + ' (' + a.pattiSize + ')'; });

  const units = {};
  allocList.forEach(function(a) { units[a.unit] = true; });

  if (stepInfo.allDone && steps.length) {
    lifecycle = APP.LIFECYCLE.COMPLETED;
  } else if (stepInfo.completed.length) {
    lifecycle = APP.LIFECYCLE.IN_PRODUCTION;
  }

  const values = new Array(header.length).fill('');

  setArrayValue_(values, header, 'Plan ID', planId);
  setArrayValue_(values, header, 'UUID', data.uuid);
  setArrayValue_(values, header, 'Order No', data.orderNo);
  setArrayValue_(values, header, 'Order Date', dateOnly_(order['Date']));
  setArrayValue_(values, header, 'Customer', order['Customer'] || '');
  setArrayValue_(
    values,
    header,
    'EMB Sales Coordinator',
    order['EMB Sales Coordinator'] || ''
  );
  setArrayValue_(values, header, 'Design No', order['Design No'] || '');
  setArrayValue_(values, header, 'Colour', order['Colour'] || '');
  setArrayValue_(values, header, 'Fabric', order['Fabric'] || '');
  setArrayValue_(values, header, 'Order Type', order['Order Type'] || '');
  setArrayValue_(values, header, 'Order QTY', order['Order QTY'] || '');
  setArrayValue_(
    values, header, 'Delivery Date', dateOnly_(order['Delivery Date'])
  );
  setArrayValue_(values, header, 'Grid', data.grid);
  setArrayValue_(values, header, 'Revo', data.revo);
  setArrayValue_(values, header, 'Desc', data.desc);
  setArrayValue_(values, header, 'Greige/Finish', data.greigeFinish);
  setArrayValue_(values, header, 'Remaining Qty', data.remainingQty);
  setArrayValue_(values, header, 'Stock Current Stage', data.stockStage || '');
  setArrayValue_(values, header, 'Mending Type', data.mendingType || '');
  setArrayValue_(values, header, 'Total Steps', steps.length);
  setArrayValue_(
    values,
    header,
    'Flow Steps',
    steps.map(function(s) { return s.key; }).join(', ')
  );
  setArrayValue_(
    values,
    header,
    'Flow Sequence',
    steps.map(function(s) { return s.name; }).join(' -> ')
  );

  APP.FLOW_STEPS.forEach(function(step) {
    setArrayValue_(
      values,
      header,
      step.column,
      dateOnly_((data.dates && data.dates[step.column]) || '')
    );
  });

  setArrayValue_(values, header, 'Plan Start Date', planStart);
  setArrayValue_(values, header, 'Plan End Date', planEnd);
  setArrayValue_(
    values, header, 'Plan Duration Days', daysBetween_(planStart, planEnd)
  );
  setArrayValue_(values, header, 'Production Place', context.place);
  setArrayValue_(
    values, header, 'Production Units', Object.keys(units).join(', ')
  );
  setArrayValue_(values, header, 'Allocation Count', allocList.length);
  setArrayValue_(values, header, 'Allocated Qty', allocatedQty || '');
  setArrayValue_(values, header, 'Outside Parties', parties.join(', '));
  setArrayValue_(values, header, 'Current Step Index', stepInfo.currentIndex);
  setArrayValue_(
    values,
    header,
    'Current Step',
    stepInfo.allDone && steps.length ? 'All steps completed' : stepInfo.currentName
  );
  setArrayValue_(values, header, 'Current Step Date', stepInfo.currentDate);
  setArrayValue_(
    values, header, 'Completed Steps', stepInfo.completed.join(' -> ')
  );
  setArrayValue_(
    values, header, 'Pending Steps', stepInfo.pending.join(' -> ')
  );
  setArrayValue_(values, header, 'Lifecycle Status', lifecycle);
  setArrayValue_(
    values, header, 'Created Timestamp', createdAt || timestamp_()
  );
  setArrayValue_(values, header, 'Updated Timestamp', timestamp_());

  if (targetRow === -1) {
    targetRow = sh.getLastRow() + 1;
  }

  sh.getRange(
    targetRow,
    1,
    1,
    header.length
  ).setValues([values]);

  return targetRow;
}




/* ============================================================================
   16. PRODUCTION PLAN API
   ========================================================================== */

/*
 * FRESH INDEXING
 * --------------
 * Rebuilds the whole "Plan Steps" analysis table from the Production Plan
 * sheet. Safe to run any time: it clears the sheet and writes one clean,
 * fully indexed row per order x step.
 *
 * Run manually from the Apps Script editor: rebuildPlanIndex()
 */
function rebuildPlanIndex() {
  const planSheet = ensurePlanSheet_();
  const stepsSheet = ensurePlanStepsSheet_();

  if (stepsSheet.getLastRow() > 1) {
    stepsSheet.getRange(
      2,
      1,
      stepsSheet.getLastRow() - 1,
      stepsSheet.getLastColumn()
    ).clearContent();
  }

  const planHeader = getHeader_(planSheet);
  const data = readSheet_(APP.SHEETS.PLAN);
  const orders = readSheet_(APP.SHEETS.ORDERS);

  const orderByUuid = {};
  orders.rows.forEach(function(row) {
    orderByUuid[normalize_(row['UUID'])] = row;
  });

  let planCount = 0;
  let stepCount = 0;

  data.rows.forEach(function(plan) {
    const uuid = normalize_(plan['UUID']);
    if (!uuid) return;

    const order = orderByUuid[uuid] || {};
    const mendingType = normalize_(plan['Mending Type']);

    let keys = sanitizeFlowSteps_(
      String(plan['Flow Steps'] || '').split(','),
      mendingType
    );

    /* Legacy rows without a Flow Steps value: infer from filled dates. */
    if (!keys.length) {
      keys = sanitizeFlowSteps_(
        APP.FLOW_STEPS
          .filter(function(s) { return normalize_(plan[s.column]); })
          .map(function(s) { return s.key; }),
        mendingType
      );
    }

    if (!keys.length) return;

    const dates = {};
    APP.FLOW_STEPS.forEach(function(step) {
      dates[step.column] = plan[step.column] || '';
    });

    const steps = buildPlanStepRecords_(keys, mendingType, dates);
    const planId = normalize_(plan['Plan ID']) ||
      planIdFor_(uuid, plan['Order No']);

    /* Refresh the plan header row with the derived index columns. */
    setIfColumnExists_(planSheet, planHeader, plan.__row, 'Plan ID', planId);
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Total Steps', steps.length
    );
    setIfColumnExists_(
      planSheet,
      planHeader,
      plan.__row,
      'Flow Steps',
      steps.map(function(s) { return s.key; }).join(', ')
    );
    setIfColumnExists_(
      planSheet,
      planHeader,
      plan.__row,
      'Flow Sequence',
      steps.map(function(s) { return s.name; }).join(' -> ')
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Plan Start Date', steps[0].date
    );
    setIfColumnExists_(
      planSheet,
      planHeader,
      plan.__row,
      'Plan End Date',
      steps[steps.length - 1].date
    );
    setIfColumnExists_(
      planSheet,
      planHeader,
      plan.__row,
      'Plan Duration Days',
      daysBetween_(steps[0].date, steps[steps.length - 1].date)
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Order Date', dateOnly_(order['Date'])
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Customer', order['Customer'] || ''
    );
    setIfColumnExists_(
      planSheet,
      planHeader,
      plan.__row,
      'EMB Sales Coordinator',
      order['EMB Sales Coordinator'] || ''
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Design No', order['Design No'] || ''
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Colour', order['Colour'] || ''
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Fabric', order['Fabric'] || ''
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Order Type', order['Order Type'] || ''
    );
    setIfColumnExists_(
      planSheet, planHeader, plan.__row, 'Order QTY', order['Order QTY'] || ''
    );
    setIfColumnExists_(
      planSheet,
      planHeader,
      plan.__row,
      'Delivery Date',
      dateOnly_(order['Delivery Date'])
    );

    stepCount += syncPlanSteps_({
      planId: planId,
      uuid: uuid,
      orderNo: plan['Order No'],
      orderDate: order['Date'] || '',
      customer: order['Customer'] || '',
      coordinator: order['EMB Sales Coordinator'] || '',
      designNo: order['Design No'] || '',
      colour: order['Colour'] || '',
      fabric: order['Fabric'] || '',
      orderType: order['Order Type'] || '',
      deliveryDate: order['Delivery Date'] || '',
      grid: plan['Grid'],
      revo: plan['Revo'],
      desc: plan['Desc'],
      greigeFinish: plan['Greige/Finish'],
      remainingQty: plan['Remaining Qty'],
      stockStage: plan['Stock Current Stage'],
      mendingType: mendingType,
      place: plan['Production Place'] || '',
      lifecycle: plan['Lifecycle Status'] || APP.LIFECYCLE.PLANNED,
      steps: steps
    }).count;


    planCount++;
  });

  log_('PLAN_INDEX_REBUILT', { plans: planCount, steps: stepCount });

  return { plans: planCount, steps: stepCount };
}

/* Admin-callable version for the UI. */
function apiRebuildPlanIndex(token) {
  return withError_(function() {
    const session = requireSession_(token);

    if (session.role !== 'admin') {
      throw new Error('Only an admin can rebuild the plan index.');
    }

    return ok_(rebuildPlanIndex());
  });
}

function apiRebuildPlanIndexJson(token) {
  return jsonString_(apiRebuildPlanIndex(token));
}

/* Full step-level dataset for analysis screens. */
function apiGetPlanSteps(token, uuid) {
  return withError_(function() {
    requireSession_(token);
    ensurePlanStepsSheet_();

    uuid = normalize_(uuid);
    const data = readSheet_(APP.SHEETS.PLAN_STEPS);

    const rows = uuid
      ? data.rows.filter(function(r) {
          return normalize_(r['UUID']) === uuid;
        })
      : data.rows;

    return ok_({ steps: rows });
  });
}

function apiGetPlanStepsJson(token, uuid) {
  return jsonString_(apiGetPlanSteps(token, uuid));
}



function apiGetProductionPlan(token, uuid) {
  return withError_(function() {
    const session = requireSession_(token);
    uuid = normalize_(uuid);

    if (!uuid) throw new Error('UUID is required.');

    const data = readSheet_(APP.SHEETS.PLAN);

    const plan = data.rows.find(function(row) {
      return normalize_(row['UUID']) === uuid;
    });

    if (!plan) {
      return ok_({
        plan: null
      });
    }

    const orderFound = findOrderByUUID_(uuid);

    if (
      orderFound &&
      normalize_(orderFound.row['EMB Sales Coordinator']) !==
      normalize_(session.name) &&
      session.role !== 'admin'
    ) {
      throw new Error('You do not have access to this production plan.');
    }

    return ok_({
      plan: plan
    });
  });
}

function apiGetProductionPlanJson(token, uuid) {
  return jsonString_(apiGetProductionPlan(token, uuid));
}

function apiGetMyProductionPlans(token) {
  return withError_(function() {
    const session = requireSession_(token);

    const ordersData = readSheet_(APP.SHEETS.ORDERS);
    const plansData = readSheet_(APP.SHEETS.PLAN);

    const myUUIDs = {};

    ordersData.rows.forEach(function(order) {
      /* Closed orders are hidden from the plan list too. */
      if (isOrderClosed_(order)) return;

      if (
        session.role === 'admin' ||
        normalize_(order['EMB Sales Coordinator']) ===
        normalize_(session.name)
      ) {
        const uuid = normalize_(order['UUID']);
        if (uuid) myUUIDs[uuid] = true;
      }
    });


    const plans = plansData.rows.filter(function(plan) {
      return !!myUUIDs[normalize_(plan['UUID'])];
    });

    /*
     * "Active" plans only — once a plan finishes every step it is marked
     * Completed and belongs in Review Orders (ready to close), not in the
     * Production Plan list, which is meant for plans still in progress.
     * The full, unfiltered list (including completed) is still returned
     * separately so the Process Board can keep its "Completed" bucket.
     */
    const activePlans = plans.filter(function(plan) {
      return normalize_(plan['Lifecycle Status']) !== APP.LIFECYCLE.COMPLETED;
    });

    return ok_({
      plans: plans,
      activePlans: activePlans
    });
  });
}

function apiGetMyProductionPlansJson(token) {
  return jsonString_(apiGetMyProductionPlans(token));
}

function apiUpdatePlanDates(token, uuid, dates) {
  return withError_(function() {
    const session = requireSession_(token);
    uuid = normalize_(uuid);

    if (!uuid) throw new Error('UUID is required.');

    dates = dates || {};

    const found = findOrderByUUID_(uuid);

    if (!found) {
      throw new Error('Order not found.');
    }

    if (
      normalize_(found.row['EMB Sales Coordinator']) !==
      normalize_(session.name) &&
      session.role !== 'admin'
    ) {
      throw new Error('You do not have access to this order.');
    }

    ensurePlanSheet_();
    const data = readSheet_(APP.SHEETS.PLAN);

    const plan = data.rows.find(function(row) {
      return normalize_(row['UUID']) === uuid;
    });

    if (!plan) {
      throw new Error('Production plan not found.');
    }

    const mendingType = dates.mendingType !== undefined
      ? normalize_(dates.mendingType)
      : normalize_(plan['Mending Type']);

    const currentSteps = dates.flowSteps !== undefined
      ? dates.flowSteps
      : String(plan['Flow Steps'] || '').split(',');

    const flowSteps = sanitizeFlowSteps_(currentSteps, mendingType);

    if (!flowSteps.length) {
      throw new Error('Select at least one production flow step.');
    }

    if (flowSteps.indexOf('mending') !== -1 && !mendingType) {
      throw new Error('Select Pre Mending or Post Mending.');
    }

    /* Merge: keep stored dates, override with anything the user supplied. */
    const merged = {};

    APP.FLOW_STEPS.forEach(function(step) {
      merged[step.column] = flowSteps.indexOf(step.key) === -1
        ? ''
        : normalize_(
            dates[step.column] !== undefined
              ? dates[step.column]
              : plan[step.column]
          );
    });

    /* Production place / patti split can also be edited from the plan. */
    let place = normalize_(plan['Production Place']);
    let allocations = null;

    if (dates.productionPlace !== undefined || dates.allocations !== undefined) {
      const split = sanitizeAllocations_(
        dates.productionPlace !== undefined
          ? dates.productionPlace
          : place,
        dates.allocations !== undefined
          ? dates.allocations
          : readAllocationsForUUID_(uuid),
        plan['Remaining Qty']
      );

      place = split.place;
      allocations = split.allocations;
    }

    upsertProductionPlan_({
      uuid: uuid,
      orderNo: plan['Order No'] || found.row['Order No'],
      grid: plan['Grid'],
      revo: plan['Revo'],
      desc: plan['Desc'],
      greigeFinish: plan['Greige/Finish'],
      remainingQty: plan['Remaining Qty'],
      stockStage: plan['Stock Current Stage'],
      dates: merged,
      flowSteps: flowSteps,
      mendingType: mendingType,
      place: place,
      allocations: allocations,
      order: found.row
    });


    log_('PLAN_DATES_UPDATED', {
      uuid: uuid,
      by: session.username,
      place: place,
      allocations: allocations,
      dates: dates
    });

    return ok_();
  });
}

function apiUpdatePlanDatesJson(token, uuid, dates) {
  return jsonString_(apiUpdatePlanDates(token, uuid, dates));
}

/*
 * Marks the CURRENT step of an order as done and moves the order to the
 * next step. Completed steps keep their record, so an order never stays
 * sitting in a step it has already finished.
 */
function apiCompleteStep(token, uuid, stepKey) {
  return withError_(function() {
    const session = requireSession_(token);
    uuid = normalize_(uuid);

    if (!uuid) throw new Error('UUID is required.');

    const found = findOrderByUUID_(uuid);
    if (!found) throw new Error('Order not found.');

    if (
      normalize_(found.row['EMB Sales Coordinator']) !==
      normalize_(session.name) &&
      session.role !== 'admin' &&
      session.role !== 'department'
    ) {
      throw new Error('You do not have access to this order.');
    }

    ensurePlanSheet_();
    const data = readSheet_(APP.SHEETS.PLAN);

    const plan = data.rows.find(function(row) {
      return normalize_(row['UUID']) === uuid;
    });

    if (!plan) throw new Error('Production plan not found.');

    const mendingType = normalize_(plan['Mending Type']);

    const flowSteps = sanitizeFlowSteps_(
      String(plan['Flow Steps'] || '').split(','),
      mendingType
    );

    if (!flowSteps.length) throw new Error('This plan has no flow steps.');

    /* Already-completed steps come from the normalized Plan Steps index. */
    ensurePlanStepsSheet_();
    const stepData = readSheet_(APP.SHEETS.PLAN_STEPS);

    const completedKeys = stepData.rows
      .filter(function(r) {
        return normalize_(r['UUID']) === uuid &&
          /completed/i.test(normalize_(r['Step Status']));
      })
      .map(function(r) { return normalize_(r['Step Key']); });

    const target = normalize_(stepKey).toLowerCase() ||
      flowSteps.filter(function(k) {
        return completedKeys.indexOf(k) === -1;
      })[0];

    if (!target) throw new Error('All steps are already completed.');

    if (flowSteps.indexOf(target) === -1) {
      throw new Error('That step is not part of this order flow.');
    }

    if (session.role === 'department' && normalize_(session.department) !== target) {
      throw new Error('This step does not belong to your department.');
    }

    if (completedKeys.indexOf(target) === -1) completedKeys.push(target);

    const dates = {};

    APP.FLOW_STEPS.forEach(function(step) {
      dates[step.column] = flowSteps.indexOf(step.key) === -1
        ? ''
        : normalize_(plan[step.column]);
    });

    upsertProductionPlan_({
      uuid: uuid,
      orderNo: plan['Order No'] || found.row['Order No'],
      grid: plan['Grid'],
      revo: plan['Revo'],
      desc: plan['Desc'],
      greigeFinish: plan['Greige/Finish'],
      remainingQty: plan['Remaining Qty'],
      stockStage: plan['Stock Current Stage'],
      dates: dates,
      flowSteps: flowSteps,
      mendingType: mendingType,
      place: normalize_(plan['Production Place']),
      completedKeys: completedKeys,
      completedBy: session.name || session.username,
      order: found.row
    });

    const next = flowSteps.filter(function(k) {
      return completedKeys.indexOf(k) === -1;
    })[0];

    const nextStep = next ? flowStepByKey_(next) : null;

    log_('PLAN_STEP_COMPLETED', {
      uuid: uuid,
      step: target,
      by: session.username
    });

    return ok_({
      uuid: uuid,
      completedStep: target,
      nextStep: nextStep ? nextStep.key : '',
      nextStepName: nextStep ? nextStep.label : '',
      allDone: !next
    });
  });
}

function apiCompleteStepJson(token, uuid, stepKey) {
  return jsonString_(apiCompleteStep(token, uuid, stepKey));
}



/* ============================================================================
   17. ADMIN / NOTIFICATIONS / LOG
   ========================================================================== */

function apiGetNotifications(token) {
  return withError_(function() {
    requireSession_(token);

    const data = readSheet_(APP.SHEETS.NOTIFICATIONS);

    return ok_({
      notifications: data.rows
    });
  });
}

function apiGetNotificationsJson(token) {
  return jsonString_(apiGetNotifications(token));
}

function apiResolveNotification(token, rowNumber) {
  return withError_(function() {
    const session = requireSession_(token);

    if (session.role !== 'admin') {
      throw new Error('Admin access required.');
    }

    const sh = getSheet_(APP.SHEETS.NOTIFICATIONS);
    const header = getHeader_(sh);

    setIfColumnExists_(
      sh,
      header,
      Number(rowNumber),
      'Resolved',
      true
    );

    log_('NOTIFICATION_RESOLVED', {
      row: Number(rowNumber),
      by: session.username
    });

    return ok_();
  });
}

function apiResolveNotificationJson(token, rowNumber) {
  return jsonString_(apiResolveNotification(token, rowNumber));
}

function apiGetLogs(token) {
  return withError_(function() {
    const session = requireSession_(token);

    if (session.role !== 'admin') {
      throw new Error('Admin access required.');
    }

    return ok_({
      logs: readSheet_(APP.SHEETS.LOG).rows
    });
  });
}

function apiGetLogsJson(token) {
  return jsonString_(apiGetLogs(token));
}


/* ============================================================================
   18. GENERIC COLUMN WRITE
   ========================================================================== */

function setIfColumnExists_(
  sheet,
  header,
  rowNumber,
  columnName,
  value
) {
  const idx = headerIndex_(header, columnName);

  if (idx === -1) return false;

  sheet
    .getRange(Number(rowNumber), idx + 1)
    .setValue(value);

  return true;
}


/* ============================================================================
   19. SETUP / MIGRATION
   ========================================================================== */

/*
 * The one-time setup routine (create missing sheets/columns, seed the first
 * admin account, prepare the dispatch Master_Data sheet) lives in the
 * separate _SETUP.gs file so that it can be deleted after the spreadsheet is
 * configured. Run setupApp() there once, then delete _SETUP.gs.
 *
 * The helpers used below (ensureUsersSheet_, ensureOrderWorkflowColumns_,
 * ensurePlanSheet_, ensureNotificationsSheet_, ensureLogSheet_,
 * ensureGApvColumns_, ensurePlanStepsSheet_, ensureAllocationSheet_) are
 * part of this runtime backend and are kept here.
 */


/* ============================================================================
   20. SHARED SECRET
   ========================================================================== */

function getSharedSecret_() {
  return PropertiesService
    .getScriptProperties()
    .getProperty('SHARED_SECRET') || '';
}


/* ============================================================================
   21. OPTIONAL TEST / DIAGNOSTIC FUNCTIONS
   ========================================================================== */

function testOrderLookup_(uuid) {
  const result = findOrderByUUID_(uuid);

  if (!result) {
    return fail_('Order not found.');
  }

  return ok_({
    order: result.row
  });
}

function testGApvLookup_(uuid) {
  const rows = findGApvByUUID_(uuid);

  return ok_({
    rows: rows,
    stockStage: determineStockStageFromRows_(rows)
  });
}

function testReview_(uuid) {
  const rows = findGApvByUUID_(uuid);

  if (!rows.length) {
    return fail_('No G Apv row found.');
  }

  const found = findOrderByUUID_(uuid);

  if (!found) {
    return fail_('No Orders row found.');
  }

  const total = APP.GAPV_STAGE_COLUMNS.reduce(function(sum, stage) {
    return sum + rows.reduce(function(rowSum, row) {
      return rowSum + number_(row[stage]);
    }, 0);
  }, 0);

  return ok_({
    uuid: uuid,
    orderNo: found.row['Order No'],
    orderQty: number_(found.row['Order QTY']),
    externalInvQty: total,
    remainingQty: Math.max(
      number_(found.row['Order QTY']) - total,
      0
    ),
    stockStage: determineStockStageFromRows_(rows)
  });
}


/* ============================================================================
   20. PREVIEW ORDERS / CLOSE ORDER
   --------------------------------------------------------------------------
   Every order stays visible in "Preview Orders" until it is explicitly
   closed. A closed order disappears from every screen but stays forever in
   the spreadsheet database.
   ========================================================================== */

function isYesValue_(value) {
  return /^(yes|y|true|1|closed|done|completed)$/i.test(
    normalize_(value)
  );
}

function isOrderClosed_(order) {
  return isYesValue_(order && order['Order Closed']);
}

/* Map of UUID -> true for every closed order. */
function closedUuidMap_() {
  const map = {};
  const data = readSheet_(APP.SHEETS.ORDERS);

  data.rows.forEach(function(row) {
    if (isOrderClosed_(row)) {
      const uuid = normalize_(row['UUID']);
      if (uuid) map[uuid] = true;
    }
  });

  return map;
}

/*
 * PREVIEW ORDERS
 * --------------
 * One merged record per order: order details + plan progress + allocation
 * summary. Returns everything that is still open (not closed).
 */
function apiGetPreviewOrders(token) {
  return withError_(function() {
    const session = requireSession_(token);
    const isAdmin = session.role === 'admin';

    const ordersData = readSheet_(APP.SHEETS.ORDERS);

    ensurePlanSheet_();
    const plansData = readSheet_(APP.SHEETS.PLAN);

    const planByUuid = {};

    plansData.rows.forEach(function(plan) {
      const uuid = normalize_(plan['UUID']);
      if (uuid) planByUuid[uuid] = plan;
    });

    const orders = [];

    ordersData.rows.forEach(function(order) {
      const uuid = normalize_(order['UUID']);
      if (!uuid) return;

      if (isOrderClosed_(order)) return;

      /*
       * Stock Only orders never go through production steps, so they have
       * nothing to "review" here — Review Orders is for production orders
       * that finished every planned step.
       */
      if (normalize_(order['Status']) === APP.STATUS.STOCK_ONLY) return;

      if (
        !isAdmin &&
        normalize_(order['EMB Sales Coordinator']) !== normalize_(session.name)
      ) {
        return;
      }

      const plan = planByUuid[uuid] || {};
      const totalSteps = number_(plan['Total Steps']);
      const currentIndex = number_(plan['Current Step Index']);
      const lifecycle = normalize_(plan['Lifecycle Status']);

      const completedSteps = normalize_(plan['Completed Steps']);

      const doneCount = completedSteps
        ? completedSteps.split('->').length
        : 0;

      const readyToClose =
        lifecycle === APP.LIFECYCLE.COMPLETED ||
        (!!totalSteps && doneCount >= totalSteps);

      /*
       * Review Orders only shows orders once every planned step is done —
       * i.e. it's a final-review list before closing, not a general
       * "all open orders" list (that's what the Process Board is for).
       */
      if (!readyToClose) return;

      orders.push({
        'UUID': uuid,
        'Order No': order['Order No'] || '',
        'Date': order['Date'] || '',
        'Customer': order['Customer'] || '',
        'EMB Sales Coordinator': order['EMB Sales Coordinator'] || '',
        'Design No': order['Design No'] || '',
        'Colour': order['Colour'] || '',
        'Fabric': order['Fabric'] || '',
        'Order Type': order['Order Type'] || '',
        'Order QTY': order['Order QTY'] || '',
        'Qty Unit': order['Qty Unit'] || '',
        'Delivery Date': order['Delivery Date'] || '',
        'Status': order['Status'] || APP.STATUS.NEW,
        'Godown': order['Godown'] || '',
        'External Inv Qty': order['External Inv Qty'] || '',
        'Remaining Qty':
          plan['Remaining Qty'] || order['Remaining Qty'] || '',
        'Production Decision': order['Production Decision'] || '',
        'Plan ID': plan['Plan ID'] || '',
        'Production Place': plan['Production Place'] || '',
        'Production Units': plan['Production Units'] || '',
        'Outside Parties': plan['Outside Parties'] || '',
        'Allocated Qty': plan['Allocated Qty'] || '',
        'Mending Type': plan['Mending Type'] || '',
        'Flow Sequence': plan['Flow Sequence'] || '',
        'Total Steps': totalSteps || '',
        'Completed Steps': completedSteps,
        'Pending Steps': plan['Pending Steps'] || '',
        'Current Step': plan['Current Step'] || '',
        'Current Step Index': currentIndex || '',
        'Current Step Date': plan['Current Step Date'] || '',
        'Plan Start Date': plan['Plan Start Date'] || '',
        'Plan End Date': plan['Plan End Date'] || '',
        'Lifecycle Status': lifecycle || '',
        'Progress': totalSteps
          ? Math.round((doneCount / totalSteps) * 100)
          : 0,
        'Steps Done': doneCount,
        'Ready To Close': readyToClose ? 'Yes' : 'No',
        'Updated Timestamp': plan['Updated Timestamp'] || ''
      });
    });

    return ok_({
      orders: orders,
      user: session.name
    });
  });
}

function apiGetPreviewOrdersJson(token) {
  return jsonString_(apiGetPreviewOrders(token));
}

/*
 * CLOSE ORDER
 * -----------
 * Marks the order closed. It stays in the sheets (database) but is filtered
 * out of every UI list from then on.
 */
function apiCloseOrder(token, uuid) {
  return withError_(function() {
    const session = requireSession_(token);
    uuid = normalize_(uuid);

    if (!uuid) throw new Error('UUID is required.');

    ensureOrderWorkflowColumns_();

    const sh = getSheet_(APP.SHEETS.ORDERS);
    const header = getHeader_(sh);
    const lastRow = sh.getLastRow();

    if (lastRow < 2) throw new Error('Order not found.');

    const uuidCol = headerIndex_(header, 'UUID');
    const coordCol = headerIndex_(header, 'EMB Sales Coordinator');

    if (uuidCol === -1) throw new Error('Orders sheet has no UUID column.');

    const values = sh
      .getRange(2, 1, lastRow - 1, sh.getLastColumn())
      .getDisplayValues();

    let targetRow = -1;
    let orderNo = '';

    for (let i = 0; i < values.length; i++) {
      if (normalize_(values[i][uuidCol]) === uuid) {
        if (
          session.role !== 'admin' &&
          coordCol !== -1 &&
          normalize_(values[i][coordCol]) !== normalize_(session.name)
        ) {
          throw new Error('You do not have access to this order.');
        }

        targetRow = i + 2;
        const noCol = headerIndex_(header, 'Order No');
        orderNo = noCol === -1 ? '' : values[i][noCol];
        break;
      }
    }

    if (targetRow === -1) throw new Error('Order not found.');

    const pendingDispatch = safeReadDispatchRows_().filter(function(r) {
      return normalize_(r['UUID']) === uuid && !isYesValue_(r['Done']);
    });
    if (pendingDispatch.length) {
      throw new Error(
        'This order still has ' + pendingDispatch.length +
        ' pending dispatch request(s). Mark them done before completing the order.'
      );
    }

    const now = timestamp_();

    setIfColumnExists_(sh, header, targetRow, 'Order Closed', 'Yes');
    setIfColumnExists_(sh, header, targetRow, 'Closed Timestamp', now);
    setIfColumnExists_(
      sh, header, targetRow, 'Closed By', session.name || session.username
    );

    /* Mirror the closure on the plan header so analysis sheets agree. */
    ensurePlanSheet_();
    const planSh = getSheet_(APP.SHEETS.PLAN);
    const planHeader = getHeader_(planSh);
    const planLast = planSh.getLastRow();

    if (planLast >= 2) {
      const planUuidCol = headerIndex_(planHeader, 'UUID');

      if (planUuidCol !== -1) {
        const planValues = planSh
          .getRange(2, 1, planLast - 1, planSh.getLastColumn())
          .getDisplayValues();

        for (let j = 0; j < planValues.length; j++) {
          if (normalize_(planValues[j][planUuidCol]) === uuid) {
            setIfColumnExists_(
              planSh,
              planHeader,
              j + 2,
              'Lifecycle Status',
              APP.LIFECYCLE.CLOSED
            );
            setIfColumnExists_(
              planSh, planHeader, j + 2, 'Updated Timestamp', now
            );
            break;
          }
        }
      }
    }

    log_('ORDER_CLOSED', {
      uuid: uuid,
      orderNo: orderNo,
      by: session.username
    });

    return ok_({ uuid: uuid, orderNo: orderNo, closedAt: now });
  });
}

function apiCloseOrderJson(token, uuid) {
  return jsonString_(apiCloseOrder(token, uuid));
}

/* ============================================================================
   21. DISPATCH — EXTERNAL MASTER_DATA SHEET
   ========================================================================== */

/* Opens the external dispatch spreadsheet and ensures the Master_Data sheet
   with the required header row exists. */
function getDispatchSheet_() {
  const ss = SpreadsheetApp.openById(APP.DISPATCH_SHEET_ID);
  let sh = ss.getSheetByName(APP.DISPATCH_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(APP.DISPATCH_SHEET_NAME);
    sh.getRange(1, 1, 1, APP.DISPATCH_COLUMNS.length)
      .setValues([APP.DISPATCH_COLUMNS]);
    sh.setFrozenRows(1);
  }

  return sh;
}

function readDispatchRows_() {
  const sh = getDispatchSheet_();
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return [];

  const header = getHeader_(sh);
  const values = sh.getRange(2, 1, lastRow - 1, header.length)
    .getDisplayValues();

  const rows = [];

  values.forEach(function(row, index) {
    const obj = {};
    let hasAny = false;

    header.forEach(function(h, c) {
      if (!h) return;
      obj[h] = row[c];
      if (row[c] !== '') hasAny = true;
    });

    if (hasAny) {
      obj.__row = index + 2;
      rows.push(obj);
    }
  });

  return rows;
}

/* Reading the dispatch sheet should never break core workflow actions — if
   the external spreadsheet is missing/renamed, treat it as having no rows. */
function safeReadDispatchRows_() {
  try {
    return readDispatchRows_();
  } catch (e) {
    return [];
  }
}

/**
 * SALES — lists the user's open orders (candidates for dispatch) plus the
 * dispatch requests they have already sent.
 */
function apiGetDispatchData(token) {
  return withError_(function() {
    const session = requireSession_(token);

    const ordersData = readSheet_(APP.SHEETS.ORDERS);
    const orders = [];

    ordersData.rows.forEach(function(order) {
      if (isOrderClosed_(order)) return;
      if (
        session.role !== 'admin' &&
        normalize_(order['EMB Sales Coordinator']) !== normalize_(session.name)
      ) {
        return;
      }

      const uuid = normalize_(order['UUID']);
      if (!uuid) return;

      const obj = {};
      APP.ORDER_BASE_COLUMNS.concat(APP.ORDER_WORKFLOW_COLUMNS)
        .forEach(function(col) {
          if (order[col] !== undefined) obj[col] = order[col];
        });
      obj.__row = order.__row;
      orders.push(obj);
    });

    const requests = safeReadDispatchRows_().filter(function(r) {
      if (session.role === 'admin') return true;
      return normalize_(r['Owner']) === normalize_(session.name);
    });

    return ok_({ orders: orders, requests: requests });
  });
}

function apiGetDispatchDataJson(token) {
  return jsonString_(apiGetDispatchData(token));
}

/**
 * SALES — sends a dispatch instruction for an order (may be part qty; an
 * order can be dispatched multiple times). Appends one row to Master_Data.
 */
function apiSendDispatch(token, uuid, qty) {
  return withError_(function() {
    const session = requireSession_(token);
    uuid = normalize_(uuid);
    const q = number_(qty);

    if (!uuid) throw new Error('UUID is required.');
    if (!(q > 0)) throw new Error('Dispatch qty must be greater than 0.');

    const found = findOrderByUUID_(uuid);
    if (!found) throw new Error('Order not found.');

    const order = found.row;

    if (isOrderClosed_(order)) {
      throw new Error('This order is already completed.');
    }

    if (
      session.role !== 'admin' &&
      normalize_(order['EMB Sales Coordinator']) !== normalize_(session.name)
    ) {
      throw new Error('This order does not belong to you.');
    }

    const sh = getDispatchSheet_();
    const row = APP.DISPATCH_COLUMNS.map(function(col) {
      switch (col) {
        case 'UUID': return uuid;
        case 'Order No': return order['Order No'] || '';
        case 'Customer': return order['Customer'] || '';
        case 'Design No': return order['Design No'] || '';
        case 'Colour': return order['Colour'] || '';
        case 'Fabric': return order['Fabric'] || '';
        case 'Yarn': return order['Yarn'] || '';
        case 'QTY': return q;
        case 'Qty Unit': return order['Qty Unit'] || '';
        case 'Timestamp': return timestamp_();
        case 'Owner': return session.name || session.username || '';
        case 'Done': return 'No';
        default: return '';
      }
    });

    sh.appendRow(row);

    log_('DISPATCH_REQUESTED', {
      uuid: uuid,
      orderNo: order['Order No'],
      qty: q,
      by: session.username
    });

    return ok_({
      uuid: uuid,
      orderNo: order['Order No'],
      qty: q,
      status: 'pending'
    });
  });
}

function apiSendDispatchJson(token, uuid, qty) {
  return jsonString_(apiSendDispatch(token, uuid, qty));
}

/**
 * SALES — marks an individual dispatch request as Done (column L = Yes).
 */
function apiMarkDispatchDone(token, rowNumber) {
  return withError_(function() {
    const session = requireSession_(token);
    rowNumber = Number(rowNumber);

    const sh = getDispatchSheet_();
    const header = getHeader_(sh);
    const lastRow = sh.getLastRow();

    if (!(rowNumber >= 2 && rowNumber <= lastRow)) {
      throw new Error('Dispatch request not found.');
    }

    const ownerIdx = headerIndex_(header, 'Owner');
    if (ownerIdx !== -1) {
      const owner = normalize_(sh.getRange(rowNumber, ownerIdx + 1)
        .getDisplayValue());
      if (session.role !== 'admin' && owner !== normalize_(session.name)) {
        throw new Error('You do not own this dispatch request.');
      }
    }

    setIfColumnExists_(sh, header, rowNumber, 'Done', 'Yes');

    log_('DISPATCH_MARKED_DONE', {
      rowNumber: rowNumber,
      by: session.username
    });

    return ok_({ rowNumber: rowNumber });
  });
}

function apiMarkDispatchDoneJson(token, rowNumber) {
  return jsonString_(apiMarkDispatchDone(token, rowNumber));
}


/* ============================================================================
   22. GODOWN PENDING — SENT, WAITING FOR VERIFICATION
   ========================================================================== */

/**
 * Orders with Status "Sent To Godown" that have no G Apv row back yet.
 * Once the godown sheet returns data they move to "Pending Decisions".
 */
function apiGetGodownPending(token) {
  return withError_(function() {
    const session = requireSession_(token);

    const ordersData = readSheet_(APP.SHEETS.ORDERS);
    const gapvData = readSheet_(APP.SHEETS.GAPV);

    const gapvUuids = {};
    gapvData.rows.forEach(function(g) {
      const u = normalize_(g['UUID']);
      if (u) gapvUuids[u] = true;
    });

    const orders = [];

    ordersData.rows.forEach(function(order) {
      if (isOrderClosed_(order)) return;
      if (
        session.role !== 'admin' &&
        normalize_(order['EMB Sales Coordinator']) !== normalize_(session.name)
      ) {
        return;
      }
      if (normalize_(order['Status']) !== APP.STATUS.SENT_TO_GODOWN) return;

      const uuid = normalize_(order['UUID']);
      if (gapvUuids[uuid]) return;

      const obj = {};
      APP.ORDER_BASE_COLUMNS.concat(APP.ORDER_WORKFLOW_COLUMNS)
        .forEach(function(col) {
          if (order[col] !== undefined) obj[col] = order[col];
        });
      obj['Pending Since'] = order['Sent Timestamp'] || '';
      obj.__row = order.__row;
      orders.push(obj);
    });

    return ok_({ orders: orders });
  });
}

function apiGetGodownPendingJson(token) {
  return jsonString_(apiGetGodownPending(token));
}


/* ============================================================================
   23. DEPARTMENT DASHBOARD
   ========================================================================== */

/**
 * DEPARTMENT user — sees every order currently waiting on their step
 * (pending) plus the work their department already completed (done).
 */
function apiGetDepartmentDashboard(token) {
  return withError_(function() {
    const session = requireSession_(token);

    if (session.role !== 'department' && session.role !== 'admin') {
      throw new Error('Department access required.');
    }

    const deptKey = normalize_(session.department).toLowerCase();
    if (!deptKey) {
      throw new Error('No department is assigned to your account.');
    }

    const ordersData = readSheet_(APP.SHEETS.ORDERS);
    const orderByUuid = {};

    ordersData.rows.forEach(function(r) {
      const u = normalize_(r['UUID']);
      if (u && !isOrderClosed_(r)) orderByUuid[u] = r;
    });

    ensurePlanStepsSheet_();
    const stepsData = readSheet_(APP.SHEETS.PLAN_STEPS);

    const pending = [];
    const done = [];

    stepsData.rows.forEach(function(s) {
      if (normalize_(s['Step Key']).toLowerCase() !== deptKey) return;

      const uuid = normalize_(s['UUID']);
      const order = orderByUuid[uuid];
      if (!order) return;

      const status = normalize_(s['Step Status']).toLowerCase();

      const rec = {
        'UUID': uuid,
        'Order No': order['Order No'] || s['Order No'] || '',
        'Customer': order['Customer'] || '',
        'Design No': order['Design No'] || '',
        'Colour': order['Colour'] || '',
        'Fabric': order['Fabric'] || '',
        'Desc': s['Desc'] || '',
        'Remaining Qty': s['Remaining Qty'] || '',
        'Qty Unit': order['Qty Unit'] || '',
        'Step Name': s['Step Name'] || '',
        'Step Key': s['Step Key'] || '',
        'Planned Date': s['Planned Date'] || '',
        'Completed Date': s['Completed Date'] || '',
        'Completed By': s['Completed By'] || '',
        'Step Status': s['Step Status'] || ''
      };

      if (status === 'current') pending.push(rec);
      else if (status === 'completed') done.push(rec);
    });

    done.sort(function(a, b) {
      return String(b['Completed Date']).localeCompare(
        String(a['Completed Date'])
      );
    });

    const deptMeta = APP.DEPARTMENTS.find(function(d) {
      return d.key === deptKey;
    });

    return ok_({
      department: deptKey,
      departmentLabel: deptMeta ? deptMeta.label : deptKey,
      pending: pending,
      done: done.slice(0, 200)
    });
  });
}

function apiGetDepartmentDashboardJson(token) {
  return jsonString_(apiGetDepartmentDashboard(token));
}


/* ============================================================================
   24. COMPLETED ORDERS + LIFECYCLE
   ========================================================================== */

/**
 * Orders that finished the whole lifecycle: either the plan is Completed or
 * the order has been closed. Clicking one shows its full history.
 */
function apiGetCompletedOrders(token) {
  return withError_(function() {
    const session = requireSession_(token);
    const isAdmin = session.role === 'admin';

    const ordersData = readSheet_(APP.SHEETS.ORDERS);
    const plansData = readSheet_(APP.SHEETS.PLAN);

    const planByUuid = {};
    plansData.rows.forEach(function(p) {
      const u = normalize_(p['UUID']);
      if (u) planByUuid[u] = p;
    });

    const out = [];

    ordersData.rows.forEach(function(order) {
      const uuid = normalize_(order['UUID']);
      if (!uuid) return;

      const closed = isOrderClosed_(order);
      const plan = planByUuid[uuid] || {};

      if (
        !closed &&
        normalize_(plan['Lifecycle Status']) !== APP.LIFECYCLE.COMPLETED
      ) {
        return;
      }

      if (
        !isAdmin &&
        normalize_(order['EMB Sales Coordinator']) !== normalize_(session.name)
      ) {
        return;
      }

      out.push({
        'UUID': uuid,
        'Order No': order['Order No'] || '',
        'Date': order['Date'] || '',
        'Customer': order['Customer'] || '',
        'Design No': order['Design No'] || '',
        'Colour': order['Colour'] || '',
        'Fabric': order['Fabric'] || '',
        'Order QTY': order['Order QTY'] || '',
        'Qty Unit': order['Qty Unit'] || '',
        'Delivery Date': order['Delivery Date'] || '',
        'Godown': order['Godown'] || '',
        'Status': order['Status'] || '',
        'Production Decision': order['Production Decision'] || '',
        'Flow Sequence': plan['Flow Sequence'] || '',
        'Total Steps': plan['Total Steps'] || '',
        'Completed Steps': plan['Completed Steps'] || '',
        'Lifecycle Status': closed
          ? APP.LIFECYCLE.CLOSED
          : (plan['Lifecycle Status'] || APP.LIFECYCLE.COMPLETED),
        'Closed Timestamp': order['Closed Timestamp'] || '',
        'Closed By': order['Closed By'] || '',
        'Updated Timestamp': plan['Updated Timestamp'] || ''
      });
    });

    return ok_({ orders: out });
  });
}

function apiGetCompletedOrdersJson(token) {
  return jsonString_(apiGetCompletedOrders(token));
}

/**
 * FULL LIFECYCLE of one order — order record, plan, G Apv rows, every
 * plan step (with completion dates + who completed them), dispatch requests
 * and the closure record. Used by the progress timeline and the completed
 * order detail modal.
 */
function apiGetOrderLifecycle(token, uuid) {
  return withError_(function() {
    const session = requireSession_(token);
    uuid = normalize_(uuid);

    if (!uuid) throw new Error('UUID is required.');

    const found = findOrderByUUID_(uuid);
    if (!found) throw new Error('Order not found.');

    const order = found.row;

    if (
      session.role !== 'admin' &&
      normalize_(order['EMB Sales Coordinator']) !== normalize_(session.name)
    ) {
      throw new Error('You do not have access to this order.');
    }

    const o = {};
    APP.ORDER_BASE_COLUMNS.concat(APP.ORDER_WORKFLOW_COLUMNS)
      .forEach(function(col) {
        if (order[col] !== undefined) o[col] = order[col];
      });

    ensurePlanSheet_();
    const plansData = readSheet_(APP.SHEETS.PLAN);
    const plan = plansData.rows.find(function(p) {
      return normalize_(p['UUID']) === uuid;
    }) || {};

    const gapv = findGApvByUUID_(uuid);

    ensurePlanStepsSheet_();
    const stepsData = readSheet_(APP.SHEETS.PLAN_STEPS);
    const steps = stepsData.rows.filter(function(r) {
      return normalize_(r['UUID']) === uuid;
    });

    const dispatch = safeReadDispatchRows_().filter(function(r) {
      return normalize_(r['UUID']) === uuid;
    });

    return ok_({
      order: o,
      plan: plan,
      gapv: gapv,
      steps: steps,
      dispatch: dispatch
    });
  });
}

function apiGetOrderLifecycleJson(token, uuid) {
  return jsonString_(apiGetOrderLifecycle(token, uuid));
}


/* ============================================================================
   25. USER / CREDENTIAL MANAGEMENT
   ========================================================================== */

function apiGetUsers(token) {
  return withError_(function() {
    const session = requireSession_(token);
    if (session.role !== 'admin') {
      throw new Error('Admin access required.');
    }

    const data = readSheet_(APP.SHEETS.USERS);

    return ok_({ users: data.rows });
  });
}

function apiGetUsersJson(token) {
  return jsonString_(apiGetUsers(token));
}

/**
 * ADMIN — creates or updates an account (username, name, role, department,
 * password). Password is stored hashed. Updating an existing username keeps
 * the previous password unless a new one is supplied.
 */
function apiCreateUser(token, user) {
  return withError_(function() {
    const session = requireSession_(token);
    if (session.role !== 'admin') {
      throw new Error('Admin access required.');
    }

    user = user || {};

    const username = normalize_(user.username);
    const name = normalize_(user.name);
    const password = String(user.password || '');
    const role = normalize_(user.role).toLowerCase() || 'sales';
    const department = normalize_(user.department);
    const active = normalize_(user.active) || 'Yes';

    if (!username || !name) {
      throw new Error('Username and name are required.');
    }

    if (['admin', 'sales', 'department'].indexOf(role) === -1) {
      throw new Error('Invalid role. Choose admin, sales or department.');
    }

    if (role === 'department' && !department) {
      throw new Error('Department is required for department users.');
    }

    if (department && !APP.DEPARTMENTS.some(function(d) {
      return d.key === department;
    })) {
      throw new Error('Invalid department.');
    }

    const sh = ensureUsersSheet_();
    const data = readSheet_(APP.SHEETS.USERS);
    const header = getHeader_(sh);

    let found = -1;
    for (let i = 0; i < data.rows.length; i++) {
      if (
        normalize_(data.rows[i]['Username']).toLowerCase() ===
        username.toLowerCase()
      ) {
        found = data.rows[i].__row;
        break;
      }
    }

    if (found === -1) {
      if (!password) {
        throw new Error('Password is required for a new account.');
      }
      sh.appendRow([username, hashPassword_(password), name, role,
        department, active]);
      log_('USER_CREATED', {
        username: username, name: name, role: role,
        department: department, by: session.username
      });
      return ok_({ created: true, username: username });
    }

    /* Update an existing account. */
    setIfColumnExists_(sh, header, found, 'Sales Coordinator Name', name);
    setIfColumnExists_(sh, header, found, 'Role', role);
    setIfColumnExists_(sh, header, found, 'Department', department);
    setIfColumnExists_(sh, header, found, 'Active', active);
    if (password) {
      setIfColumnExists_(sh, header, found, 'Password',
        hashPassword_(password));
    }

    log_('USER_UPDATED', {
      username: username, name: name, role: role,
      department: department, by: session.username
    });

    return ok_({ created: false, username: username });
  });
}

function apiCreateUserJson(token, user) {
  return jsonString_(apiCreateUser(token, user));
}

/**
 * Change a password. Admins may change any account; every other user may
 * only change their own.
 */
function apiUpdatePassword(token, username, newPassword) {
  return withError_(function() {
    const session = requireSession_(token);
    username = normalize_(username);
    newPassword = String(newPassword || '');

    if (!username || !newPassword) {
      throw new Error('Username and new password are required.');
    }

    if (newPassword.length < 6) {
      throw new Error('Password must be at least 6 characters.');
    }

    if (
      session.role !== 'admin' &&
      normalize_(session.username).toLowerCase() !== username.toLowerCase()
    ) {
      throw new Error('You can only change your own password.');
    }

    const sh = ensureUsersSheet_();
    const data = readSheet_(APP.SHEETS.USERS);
    const header = getHeader_(sh);

    let found = -1;
    for (let i = 0; i < data.rows.length; i++) {
      if (
        normalize_(data.rows[i]['Username']).toLowerCase() ===
        username.toLowerCase()
      ) {
        found = data.rows[i].__row;
        break;
      }
    }

    if (found === -1) throw new Error('User not found.');

    setIfColumnExists_(sh, header, found, 'Password',
      hashPassword_(newPassword));

    log_('PASSWORD_CHANGED', {
      username: username, by: session.username
    });

    return ok_({ username: username });
  });
}

function apiUpdatePasswordJson(token, username, newPassword) {
  return jsonString_(apiUpdatePassword(token, username, newPassword));
}
