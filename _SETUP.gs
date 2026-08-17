/**
 * EMB ORDER PROCESS — ONE-TIME SHEET SETUP
 * =========================================
 *
 * This file is DELETED after the spreadsheet is configured. It contains no
 * runtime logic — it only prepares the workbook so the app works.
 *
 * HOW TO USE
 * ----------
 * 1. Attach this file + EMB_Order_Process_Code.gs to the central
 *    spreadsheet (the one containing the Orders sheet).
 * 2. Open the central spreadsheet so it is the "active" spreadsheet.
 * 3. In the Apps Script editor run:  setupApp()
 * 4. Optionally run:  addSampleUsers()  to create one account per role.
 * 5. Delete this _SETUP.gs file.
 *
 * NOTE: The first account MUST be an admin so you can log in and create the
 * rest of the accounts from the Credentials screen inside the app.
 */

/* ============================================================================
   MAIN SETUP
   ========================================================================== */

function setupApp() {
  // ---- Central workbook ----
  ensureUsersSheet_();          // Username / Password / Name / Role / Department / Active
  ensureOrderWorkflowColumns_();// Adds the workflow columns to Orders if missing

  const plan = ensurePlanSheet_();
  Logger.log('Production Plan ready: ' + plan.getName());

  const steps = ensurePlanStepsSheet_();
  Logger.log('Plan Steps ready: ' + steps.getName());

  const alloc = ensureAllocationSheet_();
  Logger.log('Production Allocation ready: ' + alloc.getName());

  ensureNotificationsSheet_();
  ensureLogSheet_();

  const gapv = getSheet_(APP.SHEETS.GAPV);
  ensureGApvColumns_(gapv);
  Logger.log('G Apv columns ready.');

  // ---- Dispatch sheet (external Master_Data spreadsheet) ----
  const dispatchSh = getDispatchSheet_();
  Logger.log('Dispatch sheet ready: ' + APP.DISPATCH_SHEET_NAME +
    ' (L = Done column)');

  // ---- Seed an admin account if the Users sheet is empty ----
  const users = readSheet_(APP.SHEETS.USERS);
  if (!users.rows.length) {
    addOrUpdateUser(
      'admin',
      'ChangeMe123!',
      'Administrator',
      'admin',
      ''
    );
    Logger.log('Seeded default admin account: admin / ChangeMe123!');
    Logger.log('>>> CHANGE THIS PASSWORD AFTER FIRST LOGIN <<<');
  } else {
    Logger.log('Users sheet already has accounts; nothing seeded.');
  }

  log_('SETUP', { message: 'EMB Order Process backend initialized.' });

  return ok_({
    message: 'Setup completed. You can now delete the _SETUP.gs file.'
  });
}


/* ============================================================================
   USER HELPERS
   ========================================================================== */

/**
 * Creates or updates one account.
 *
 *   role       : 'admin' | 'sales' | 'department'
 *   department : flow step key, only when role === 'department'
 *                (production | mending | process | qc | lace | dispatch)
 */
function addOrUpdateUser(username, plainPassword, name, role, department) {
  username = normalize_(username);
  name = normalize_(name);
  role = normalize_(role).toLowerCase() || 'sales';
  department = normalize_(department);

  if (!username || !name || !plainPassword) {
    throw new Error('Username, name and password are required.');
  }

  if (role === 'department' && !department) {
    throw new Error('Department is required for department users.');
  }

  const sh = ensureUsersSheet_();
  const data = readSheet_(APP.SHEETS.USERS);
  const header = getHeader_(sh);
  const hashed = hashPassword_(plainPassword);

  let found = -1;
  for (let i = 0; i < data.rows.length; i++) {
    if (normalize_(data.rows[i]['Username']).toLowerCase() ===
        username.toLowerCase()) {
      found = data.rows[i].__row;
      break;
    }
  }

  if (found === -1) {
    sh.appendRow([username, hashed, name, role, department, 'Yes']);
    log_('USER_CREATED', { username: username, name: name, role: role,
      department: department, by: 'setup' });
    return 'Created ' + username;
  }

  setIfColumnExists_(sh, header, found, 'Sales Coordinator Name', name);
  setIfColumnExists_(sh, header, found, 'Role', role);
  setIfColumnExists_(sh, header, found, 'Department', department);
  setIfColumnExists_(sh, header, found, 'Password', hashed);

  log_('USER_UPDATED', { username: username, name: name, role: role,
    department: department, by: 'setup' });

  return 'Updated ' + username;
}


/* ============================================================================
   OPTIONAL: SAMPLE ACCOUNTS (one per role / department)
   ========================================================================== */

function addSampleUsers() {
  const out = [];

  out.push(addOrUpdateUser('admin', 'ChangeMe123!', 'Administrator', 'admin', ''));
  out.push(addOrUpdateUser('sales.vikram', 'ChangeMe123!', 'Vikram', 'sales', ''));

  out.push(addOrUpdateUser('dept.production', 'ChangeMe123!', 'Production Incharge',
    'department', 'production'));
  out.push(addOrUpdateUser('dept.mending', 'ChangeMe123!', 'Mending Incharge',
    'department', 'mending'));
  out.push(addOrUpdateUser('dept.process', 'ChangeMe123!', 'Process Incharge',
    'department', 'process'));
  out.push(addOrUpdateUser('dept.qc', 'ChangeMe123!', 'QC Incharge',
    'department', 'qc'));
  out.push(addOrUpdateUser('dept.lace', 'ChangeMe123!', 'Lace Cutting Incharge',
    'department', 'lace'));
  out.push(addOrUpdateUser('dept.dispatch', 'ChangeMe123!', 'Dispatch Incharge',
    'department', 'dispatch'));

  Logger.log(out.join('\n'));
  return ok_({ messages: out });
}


/* ============================================================================
   OPTIONAL: GODOWN CALLBACK SHARED SECRET
   ========================================================================== */

/**
 * Run only if your godown callback posts to this web app and needs a shared
 * secret. Creates one and stores it in Script Properties.
 */
function setSharedSecret() {
  const props = PropertiesService.getScriptProperties();

  let secret = props.getProperty('SHARED_SECRET');

  if (!secret) {
    secret = Utilities.base64Encode(
      Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        Utilities.getUuid() + Date.now()
      )
    ).substring(0, 40);

    props.setProperty('SHARED_SECRET', secret);
  }

  Logger.log('SHARED_SECRET ready.');
  return ok_();
}
