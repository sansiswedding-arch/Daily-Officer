const APP = {
  TIMEZONE: 'Asia/Jakarta',
  TARGET_HOURS: 8,
  OVERTIME_BLUE_AFTER_HOURS: 9,
  DEFAULT_EMPLOYEE_ID: 'EMP001'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sansis Daily Officer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSansisDailyOfficer(ssArg) {
  const ss = ssArg || SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(APP.TIMEZONE);

  ensureSheet_(ss, 'Employees', [
    'employee_id', 'name', 'division', 'username', 'password',
    'base_salary', 'daily_rate', 'phone', 'email',
    'birth_date', 'address',
    'emergency_contact_name', 'emergency_contact_phone',
    'profile_photo_url', 'role', 'account_status'
  ]);

  ensureSheet_(ss, 'AuthSessions', [
    'token_hash', 'employee_id', 'role',
    'expires_at', 'created_at'
  ]);

  ensureSheet_(ss, 'Attendance', [
    'attendance_id', 'employee_id', 'date', 'work_mode',
    'check_in', 'check_out',
    'check_in_lat', 'check_in_lng', 'check_in_selfie_url',
    'check_out_lat', 'check_out_lng',
    'status'
  ]);

  ensureSheet_(ss, 'JobReport', [
    'report_id', 'employee_id', 'date', 'job_title',
    'description', 'documentation_url', 'created_at'
  ]);

  ensureSheet_(ss, 'VisitReports', [
    'visit_id', 'employee_id', 'date', 'visit_time',
    'purpose', 'lat', 'lng',
    'photo_1_url', 'photo_2_url', 'photo_3_url',
    'created_at'
  ]);

  ensureSheet_(ss, 'Tasks', [
    'task_id', 'employee_id', 'source', 'title',
    'task_date', 'start_time', 'end_time',
    'priority', 'status',
    'task_kind', 'deadline', 'routine_days',
    'task_assignees'
  ]);

  ensureSheet_(ss, 'TaskKPI', [
    'kpi_id', 'employee_id', 'activity', 'value',
    'status', 'cycle_start', 'created_at'
  ]);

  ensureSheet_(ss, 'KPIProofs', [
    'proof_id', 'employee_id', 'cycle_start',
    'file_name', 'proof_url', 'created_at'
  ]);

  ensureSheet_(ss, 'Reimbursements', [
    'reimburse_id', 'employee_id', 'date', 'category',
    'description', 'amount', 'proof_url', 'status', 'created_at'
  ]);

  ensureSheet_(ss, 'FinanceElements', [
    'element_id', 'employee_id', 'label', 'calculation_type',
    'amount', 'active', 'sort_order'
  ]);

  ensureSheet_(ss, 'Claims', [
    'claim_id', 'employee_id', 'claim_type', 'date',
    'claim_name', 'description', 'amount',
    'proof_url', 'status', 'created_at'
  ]);

  const emp = ss.getSheetByName('Employees');
  if (emp.getLastRow() === 1) {
    emp.appendRow([
      APP.DEFAULT_EMPLOYEE_ID,
      'Onky Soerya',
      'Marketing Relation',
      'onky',
      '',
      0,
      0,
      '',
      ''
    ]);
  }

  // Keep existing employee accounts compatible with the new login system.
  const employeeData = emp.getDataRange().getValues();
  const employeeHeaders = employeeData[0].map(String);
  const employeeIdx = {};
  employeeHeaders.forEach((h, i) => employeeIdx[h] = i);

  for (let i = 1; i < employeeData.length; i++) {
    const rowNumber = i + 1;
    if (
      employeeIdx.role != null &&
      !String(employeeData[i][employeeIdx.role] || '').trim()
    ) {
      emp.getRange(rowNumber, employeeIdx.role + 1).setValue('Employee');
    }

    if (
      employeeIdx.account_status != null &&
      !String(employeeData[i][employeeIdx.account_status] || '').trim()
    ) {
      emp.getRange(rowNumber, employeeIdx.account_status + 1).setValue('Active');
    }
  }

  const financeElements = ss.getSheetByName('FinanceElements');
  const existingFinanceElements = financeElements.getLastRow() > 1
    ? financeElements.getRange(2, 1, financeElements.getLastRow() - 1, financeElements.getLastColumn()).getValues()
    : [];

  const hasDefaultEmployeeFinance = existingFinanceElements.some(row =>
    String(row[1] || '') === APP.DEFAULT_EMPLOYEE_ID
  );

  if (!hasDefaultEmployeeFinance) {
    const currentEmployee = getEmployee_(ss, APP.DEFAULT_EMPLOYEE_ID);
    financeElements.appendRow([
      'FIN-ELEM-' + Utilities.getUuid(),
      APP.DEFAULT_EMPLOYEE_ID,
      'Base Salary Reference',
      'reference',
      Number(currentEmployee.base_salary || 0),
      true,
      1
    ]);
    financeElements.appendRow([
      'FIN-ELEM-' + Utilities.getUuid(),
      APP.DEFAULT_EMPLOYEE_ID,
      'Daily WFO Rate',
      'wfo_rate',
      Number(currentEmployee.daily_rate || 0),
      true,
      2
    ]);
  }

  return 'Setup Sansis Daily Officer selesai.';
}

function ensureAppReady_(ss) {
  const cache = CacheService.getScriptCache();
  const key = 'SANSIS_DAILY_OFFICER_SCHEMA_V8_6';

  if (cache.get(key)) return;

  setupSansisDailyOfficer(ss);
  cache.put(key, '1', 21600); // 6 hours
}


function seedDemoData() {
  setupSansisDailyOfficer();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const todayKey = formatDateKey_(now);

  const attendance = ss.getSheetByName('Attendance');
  const attendanceValues = attendance.getDataRange().getValues();
  const alreadyAttendance = attendanceValues.slice(1).some(r =>
    String(r[1]) === APP.DEFAULT_EMPLOYEE_ID &&
    dateKeyFromValue_(r[2]) === todayKey
  );

  if (!alreadyAttendance) {
    const checkIn = new Date(now.getTime() - (5 * 60 * 60 + 45 * 60) * 1000);
    attendance.appendRow([
      'ATT-DEMO-' + Date.now(),
      APP.DEFAULT_EMPLOYEE_ID,
      new Date(now.getFullYear(), now.getMonth(), now.getDate()),
      'Work From Office',
      checkIn,
      '',
      -6.5560,
      107.4460,
      '',
      '',
      '',
      'Active'
    ]);
  }

  const tasks = ss.getSheetByName('Tasks');
  const taskValues = tasks.getDataRange().getValues();
  const alreadyTasks = taskValues.slice(1).some(r =>
    String(r[1]) === APP.DEFAULT_EMPLOYEE_ID &&
    dateKeyFromValue_(r[4]) === todayKey
  );

  if (!alreadyTasks) {
    const rows = [
      ['TASK-DEMO-1', APP.DEFAULT_EMPLOYEE_ID, 'Personal', 'Review content strategy plan', new Date(now.getFullYear(), now.getMonth(), now.getDate()), '09:00', '10:30', 'Medium', 'Open'],
      ['TASK-DEMO-2', APP.DEFAULT_EMPLOYEE_ID, 'Team', 'Meeting with marketing team', new Date(now.getFullYear(), now.getMonth(), now.getDate()), '10:30', '12:00', 'High', 'Open'],
      ['TASK-DEMO-3', APP.DEFAULT_EMPLOYEE_ID, 'Personal', 'Follow up leads & client inquiries', new Date(now.getFullYear(), now.getMonth(), now.getDate()), '13:00', '14:30', 'Medium', 'Open'],
      ['TASK-DEMO-4', APP.DEFAULT_EMPLOYEE_ID, 'Team', 'Prepare weekly report', new Date(now.getFullYear(), now.getMonth(), now.getDate()), '15:00', '16:30', 'Low', 'Open'],
      ['TASK-DEMO-5', APP.DEFAULT_EMPLOYEE_ID, 'Personal', 'Update project documentation', new Date(now.getFullYear(), now.getMonth(), now.getDate()), '16:30', '17:30', 'Low', 'Open']
    ];
    tasks.getRange(tasks.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return 'Demo data berhasil dibuat.';
}


function taskAssigneeIds_(row) {
  return String(row?.task_assignees || '')
    .split(',')
    .map(x => x.trim())
    .filter(Boolean);
}

function taskVisibleToEmployee_(row, employeeId) {
  if (String(row?.employee_id || '') === String(employeeId)) return true;
  return taskAssigneeIds_(row).includes(String(employeeId));
}


function createInitialEmployeeLogin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSansisDailyOfficer(ss);

  const sheet = ss.getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let rowNumber = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.employee_id]) === APP.DEFAULT_EMPLOYEE_ID) {
      rowNumber = i + 1;
      break;
    }
  }

  if (rowNumber < 0) throw new Error('Employee default tidak ditemukan.');

  let username = String(data[rowNumber - 1][idx.username] || '').trim();
  const storedPassword = String(data[rowNumber - 1][idx.password] || '').trim();

  if (!username) {
    username = 'employee.' + String(APP.DEFAULT_EMPLOYEE_ID).toLowerCase();
    sheet.getRange(rowNumber, idx.username + 1).setValue(username);
  }

  if (idx.role != null) sheet.getRange(rowNumber, idx.role + 1).setValue('Employee');
  if (idx.account_status != null) sheet.getRange(rowNumber, idx.account_status + 1).setValue('Active');

  if (storedPassword) {
    return 'Akun sudah memiliki password. Username: ' + username;
  }

  const temporaryPassword =
    'Sns!' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);

  sheet.getRange(rowNumber, idx.password + 1).setValue(hashPassword_(temporaryPassword));

  const message =
    'Username: ' + username +
    ' | Temporary Password: ' + temporaryPassword +
    ' | Segera ganti password setelah login.';

  Logger.log(message);
  return message;
}


function createInitialAdministratorLogin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSansisDailyOfficer(ss);

  const sheet = ss.getSheetByName('Employees');
  const rows = sheetObjects_(sheet);

  const existingAdmin = rows.find(r =>
    String(r.role || '').trim().toLowerCase() === 'administrator'
  );

  if (existingAdmin) {
    return 'Administrator sudah tersedia. Username: ' +
      String(existingAdmin.username || '(belum diisi)');
  }

  const employeeId = 'ADM001';
  const username = 'admin.sansis';

  const duplicateId = rows.some(r => String(r.employee_id || '') === employeeId);
  if (duplicateId) throw new Error('ADM001 sudah digunakan.');

  const duplicateUsername = rows.some(r =>
    String(r.username || '').trim().toLowerCase() === username
  );
  if (duplicateUsername) throw new Error('Username admin.sansis sudah digunakan.');

  const temporaryPassword =
    'Adm!' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(String);

  const values = {
    employee_id: employeeId,
    name: 'Sansis Administrator',
    division: 'Administrator',
    username,
    password: hashPassword_(temporaryPassword),
    base_salary: 0,
    daily_rate: 0,
    phone: '',
    email: '',
    birth_date: '',
    address: '',
    emergency_contact_name: '',
    emergency_contact_phone: '',
    profile_photo_url: '',
    role: 'Administrator',
    account_status: 'Active'
  };

  sheet.appendRow(headers.map(h =>
    Object.prototype.hasOwnProperty.call(values, h) ? values[h] : ''
  ));

  const message =
    'Administrator Created | Username: ' + username +
    ' | Temporary Password: ' + temporaryPassword +
    ' | Segera ganti password setelah Administrator App memiliki menu Settings.';

  Logger.log(message);
  return message;
}

function loginUser(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const username = String(payload?.username || '').trim().toLowerCase();
  const password = String(payload?.password || '');

  if (!username || !password) {
    throw new Error('Username dan password wajib diisi.');
  }

  const employees = sheetObjects_(ss.getSheetByName('Employees'));
  const employee = employees.find(r =>
    String(r.username || '').trim().toLowerCase() === username
  );

  if (!employee) throw new Error('Username atau password tidak sesuai.');

  const accountStatus = String(employee.account_status || 'Active').trim();
  if (accountStatus.toLowerCase() !== 'active') {
    throw new Error('Akun sedang tidak aktif. Hubungi Administrator.');
  }

  const storedPassword = String(employee.password || '').trim();
  if (!storedPassword || !verifyPassword_(password, storedPassword)) {
    throw new Error('Username atau password tidak sesuai.');
  }

  const role = String(employee.role || 'Employee').trim() || 'Employee';
  const employeeId = String(employee.employee_id || '').trim();
  if (!employeeId) throw new Error('Employee ID akun tidak valid.');

  cleanupExpiredSessions_(ss);

  const token =
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');

  const expiresAt = new Date(Date.now() + 12 * 60 * 60 * 1000);

  ss.getSheetByName('AuthSessions').appendRow([
    hashSessionToken_(token),
    employeeId,
    role,
    expiresAt,
    new Date()
  ]);

  return {
    token,
    expiresAt: expiresAt.toISOString(),
    user: {
      employeeId,
      name: String(employee.name || 'Sansis Officer'),
      division: String(employee.division || 'Employee'),
      role,
      profilePhotoUrl: String(employee.profile_photo_url || '')
    }
  };
}

function validateSession(sessionToken) {
  try {
    const session = requireSession_(sessionToken);
    return {
      valid: true,
      user: {
        employeeId: session.employeeId,
        name: session.name,
        division: session.division,
        role: session.role,
        profilePhotoUrl: session.profilePhotoUrl
      }
    };
  } catch (err) {
    return { valid: false };
  }
}

function logoutUser(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const token = String(sessionToken || '').trim();
  if (!token) return true;

  const sheet = ss.getSheetByName('AuthSessions');
  if (!sheet || sheet.getLastRow() < 2) return true;

  const tokenHash = hashSessionToken_(token);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idxToken = headers.indexOf('token_hash');

  if (idxToken < 0) return true;

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idxToken] || '') === tokenHash) {
      sheet.deleteRow(i + 1);
    }
  }

  return true;
}

function requireSession_(sessionToken) {
  const token = String(sessionToken || '').trim();
  if (!token) throw new Error('Sesi login tidak ditemukan. Silakan login kembali.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const tokenHash = hashSessionToken_(token);
  const sessions = sheetObjects_(ss.getSheetByName('AuthSessions'));
  const row = sessions.find(r => String(r.token_hash || '') === tokenHash);

  if (!row) throw new Error('Sesi login sudah berakhir. Silakan login kembali.');

  const expiresAt = row.expires_at instanceof Date
    ? row.expires_at
    : new Date(row.expires_at);

  if (isNaN(expiresAt.getTime()) || expiresAt.getTime() <= Date.now()) {
    throw new Error('Sesi login sudah berakhir. Silakan login kembali.');
  }

  const employeeId = String(row.employee_id || '').trim();
  const employee = getEmployee_(ss, employeeId);
  if (!employee.employee_id) {
    throw new Error('Akun karyawan tidak ditemukan.');
  }

  const accountStatus = String(employee.account_status || 'Active').trim();
  if (accountStatus.toLowerCase() !== 'active') {
    throw new Error('Akun sedang tidak aktif. Hubungi Administrator.');
  }

  return {
    employeeId,
    role: String(employee.role || row.role || 'Employee').trim() || 'Employee',
    name: String(employee.name || 'Sansis Officer'),
    division: String(employee.division || 'Employee'),
    profilePhotoUrl: String(employee.profile_photo_url || '')
  };
}

function requireEmployeeSession_(sessionToken) {
  const session = requireSession_(sessionToken);
  if (session.role.toLowerCase() !== 'employee') {
    throw new Error('Akun ini bukan akun Employee.');
  }
  return session;
}

function requireAdminSession_(sessionToken) {
  const session = requireSession_(sessionToken);
  if (session.role.toLowerCase() !== 'administrator') {
    throw new Error('Akses ini hanya tersedia untuk Administrator.');
  }
  return session;
}


function cleanupExpiredSessions_(ss) {
  const sheet = ss.getSheetByName('AuthSessions');
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idxExpires = headers.indexOf('expires_at');

  for (let i = data.length - 1; i >= 1; i--) {
    const expiresAt = data[i][idxExpires] instanceof Date
      ? data[i][idxExpires]
      : new Date(data[i][idxExpires]);

    if (isNaN(expiresAt.getTime()) || expiresAt.getTime() <= Date.now()) {
      sheet.deleteRow(i + 1);
    }
  }
}

function hashSessionToken_(token) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token || ''),
    Utilities.Charset.UTF_8
  );

  return digest
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2))
    .join('');
}


function getAdminDashboardData(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  const session = requireAdminSession_(sessionToken);

  const now = new Date();
  const todayKey = formatDateKey_(now);

  const employeeRows = sheetObjects_(ss.getSheetByName('Employees'))
    .filter(r => String(r.role || 'Employee').trim().toLowerCase() === 'employee');

  const employees = employeeRows.map(r => ({
    id: String(r.employee_id || ''),
    name: String(r.name || ''),
    division: String(r.division || ''),
    username: String(r.username || ''),
    status: String(r.account_status || 'Active'),
    profilePhotoUrl: String(r.profile_photo_url || ''),
    baseSalary: Number(r.base_salary || 0),
    dailyRate: Number(r.daily_rate || 0)
  }));

  const activeEmployees = employees.filter(e => e.status.toLowerCase() === 'active');
  const inactiveEmployees = employees.filter(e => e.status.toLowerCase() !== 'active');

  const todayAttendanceRows = sheetObjects_(ss.getSheetByName('Attendance'))
    .filter(r => dateKeyFromValue_(r.date) === todayKey)
    .sort((a, b) => new Date(b.check_in || 0) - new Date(a.check_in || 0));

  const todayAttendanceByEmployee = {};
  todayAttendanceRows.forEach(r => {
    const id = String(r.employee_id || '');
    if (!todayAttendanceByEmployee[id]) todayAttendanceByEmployee[id] = r;
  });

  const activeAttendance = activeEmployees
    .map(employee => ({
      employee,
      attendance: todayAttendanceByEmployee[employee.id] || null
    }));

  const wfoCount = activeAttendance.filter(x =>
    String(x.attendance?.work_mode || '') === 'Work From Office'
  ).length;

  const wfhCount = activeAttendance.filter(x =>
    String(x.attendance?.work_mode || '') === 'Work From Home'
  ).length;

  const checkedInCount = activeAttendance.filter(x => Boolean(x.attendance?.check_in)).length;
  const checkedOutCount = activeAttendance.filter(x => Boolean(x.attendance?.check_out)).length;
  const notCheckedInCount = Math.max(0, activeEmployees.length - checkedInCount);
  const notCheckedOutCount = activeAttendance.filter(x =>
    Boolean(x.attendance?.check_in) && !x.attendance?.check_out
  ).length;

  const employeeMap = {};
  employees.forEach(e => employeeMap[e.id] = e);

  const todayJobRows = sheetObjects_(ss.getSheetByName('JobReport'))
    .filter(r => dateKeyFromValue_(r.date) === todayKey)
    .sort((a, b) => new Date(b.created_at || b.date || 0) - new Date(a.created_at || a.date || 0));

  const jobReportOverview = todayJobRows.slice(0, 10).map(r => {
    const employee = employeeMap[String(r.employee_id || '')] || {};
    const documentationUrls = String(r.documentation_url || '')
      .split('\n')
      .map(x => x.trim())
      .filter(Boolean);

    return {
      id: String(r.report_id || ''),
      employeeId: String(r.employee_id || ''),
      employeeName: String(employee.name || r.employee_id || ''),
      division: String(employee.division || ''),
      title: String(r.job_title || ''),
      description: String(r.description || ''),
      documentationUrls,
      createdAt: r.created_at instanceof Date
        ? r.created_at.toISOString()
        : String(r.created_at || '')
    };
  });

  // Current Sansis finance cycle: 5th through 4th.
  const cycle = getReportingCycle_(now);
  const cycleStartKey = dateKeyFromValue_(cycle.start);

  const cycleAttendance = sheetObjects_(ss.getSheetByName('Attendance'))
    .filter(r => r.date && isWithin_(r.date, cycle.start, cycle.end));

  const wfoDaysByEmployee = {};
  cycleAttendance.forEach(r => {
    const id = String(r.employee_id || '');
    if (String(r.work_mode || '') === 'Work From Office') {
      wfoDaysByEmployee[id] = (wfoDaysByEmployee[id] || 0) + 1;
    }
  });

  const financeElements = sheetObjects_(ss.getSheetByName('FinanceElements'))
    .filter(r => String(r.active).toLowerCase() !== 'false');

  const elementsByEmployee = {};
  financeElements.forEach(r => {
    const id = String(r.employee_id || '');
    if (!elementsByEmployee[id]) elementsByEmployee[id] = [];
    elementsByEmployee[id].push(r);
  });

  const financeRows = activeEmployees
    .map(employee => {
      const elements = elementsByEmployee[employee.id] || [];
      const wfoDays = Number(wfoDaysByEmployee[employee.id] || 0);

      let estimatedPayment = 0;

      elements.forEach(element => {
        const type = String(element.calculation_type || 'reference').trim().toLowerCase();
        const amount = Number(element.amount || 0);

        if (type === 'fixed') estimatedPayment += amount;
        if (type === 'wfo_rate') estimatedPayment += amount * wfoDays;
      });

      return {
        employeeId: employee.id,
        name: employee.name,
        division: employee.division,
        profilePhotoUrl: employee.profilePhotoUrl,
        wfoDays,
        estimatedPayment
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name));

  const totalEstimatedPayment = financeRows
    .reduce((sum, row) => sum + Number(row.estimatedPayment || 0), 0);

  return {
    administrator: {
      id: session.employeeId,
      name: session.name,
      division: session.division,
      role: session.role,
      profilePhotoUrl: session.profilePhotoUrl
    },
    employeeOverview: {
      total: employees.length,
      active: activeEmployees.length,
      inactive: inactiveEmployees.length
    },
    attendanceOverview: {
      wfo: wfoCount,
      wfh: wfhCount,
      checkedIn: checkedInCount,
      notCheckedIn: notCheckedInCount,
      checkedOut: checkedOutCount,
      notCheckedOut: notCheckedOutCount
    },
    jobReportOverview: {
      date: todayKey,
      totalActivities: todayJobRows.length,
      rows: jobReportOverview
    },
    financeOverview: {
      cycleStart: cycleStartKey,
      cycleLabel: cycle.label,
      totalEstimatedPayment,
      rows: financeRows
    },
    summary: {
      totalEmployees: employees.length,
      activeEmployees: activeEmployees.length,
      inactiveEmployees: inactiveEmployees.length,
      divisions: Array.from(new Set(
        employees.map(e => e.division).filter(Boolean)
      )).length
    },
    attendanceSummary: {
      totalActiveEmployees: activeEmployees.length,
      working: notCheckedOutCount,
      checkedOut: checkedOutCount,
      notCheckedIn: notCheckedInCount,
      totalWfo: wfoCount,
      totalWfh: wfhCount,
      totalVisits: 0
    },
    recentEmployees: employees.slice().reverse().slice(0, 5)
  };
}

function getAdminEmployeeManagementData(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const employees = sheetObjects_(ss.getSheetByName('Employees'))
    .filter(r => String(r.role || 'Employee').trim().toLowerCase() === 'employee')
    .map(r => ({
      id: String(r.employee_id || ''),
      name: String(r.name || ''),
      division: String(r.division || ''),
      username: String(r.username || ''),
      phone: String(r.phone || ''),
      email: String(r.email || ''),
      status: String(r.account_status || 'Active'),
      role: 'Employee',
      profilePhotoUrl: String(r.profile_photo_url || ''),
      hasPassword: Boolean(String(r.password || '').trim())
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const divisions = Array.from(new Set(
    employees.map(r => r.division).filter(Boolean)
  )).sort();

  return {
    employees,
    divisions,
    summary: {
      total: employees.length,
      active: employees.filter(e => e.status.toLowerCase() === 'active').length,
      inactive: employees.filter(e => e.status.toLowerCase() !== 'active').length
    }
  };
}

function createAdminEmployee(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  if (!payload) throw new Error('Data employee tidak ditemukan.');

  const sheet = ss.getSheetByName('Employees');
  const rows = sheetObjects_(sheet);

  const name = String(payload.name || '').trim();
  const division = String(payload.division || '').trim();
  const username = String(payload.username || '').trim();
  const phone = String(payload.phone || '').trim();
  const email = String(payload.email || '').trim();
  const status = String(payload.status || 'Active').trim();

  if (!name) throw new Error('Nama employee wajib diisi.');
  if (!division) throw new Error('Division wajib diisi.');
  if (!username) throw new Error('Username wajib diisi.');
  if (!/^[A-Za-z0-9._-]{4,30}$/.test(username)) {
    throw new Error('Username minimal 4 karakter dan hanya boleh berisi huruf, angka, titik, underscore, atau minus.');
  }
  if (!['Active', 'Inactive'].includes(status)) {
    throw new Error('Status akun tidak valid.');
  }

  const duplicateUsername = rows.some(r =>
    String(r.username || '').trim().toLowerCase() === username.toLowerCase()
  );
  if (duplicateUsername) throw new Error('Username tersebut sudah digunakan.');

  const employeeId = nextEmployeeId_(rows);
  const temporaryPassword =
    'Sns!' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(String);

  const values = {
    employee_id: employeeId,
    name,
    division,
    username,
    password: hashPassword_(temporaryPassword),
    base_salary: 0,
    daily_rate: 0,
    phone,
    email,
    birth_date: '',
    address: '',
    emergency_contact_name: '',
    emergency_contact_phone: '',
    profile_photo_url: '',
    role: 'Employee',
    account_status: status
  };

  sheet.appendRow(headers.map(h =>
    Object.prototype.hasOwnProperty.call(values, h) ? values[h] : ''
  ));

  return {
    employeeId,
    name,
    username,
    temporaryPassword,
    data: getAdminEmployeeManagementData(sessionToken)
  };
}

function updateAdminEmployee(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  if (!payload) throw new Error('Data employee tidak ditemukan.');

  const employeeId = String(payload.id || '').trim();
  const name = String(payload.name || '').trim();
  const division = String(payload.division || '').trim();
  const username = String(payload.username || '').trim();
  const phone = String(payload.phone || '').trim();
  const email = String(payload.email || '').trim();
  const status = String(payload.status || 'Active').trim();

  if (!employeeId) throw new Error('Employee ID tidak valid.');
  if (!name) throw new Error('Nama employee wajib diisi.');
  if (!division) throw new Error('Division wajib diisi.');
  if (!username) throw new Error('Username wajib diisi.');
  if (!/^[A-Za-z0-9._-]{4,30}$/.test(username)) {
    throw new Error('Format username tidak valid.');
  }
  if (!['Active', 'Inactive'].includes(status)) {
    throw new Error('Status akun tidak valid.');
  }

  const sheet = ss.getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Employee tidak ditemukan.');

  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const duplicateUsername = data.slice(1).some(row =>
    String(row[idx.employee_id] || '') !== employeeId &&
    String(row[idx.username] || '').trim().toLowerCase() === username.toLowerCase()
  );
  if (duplicateUsername) throw new Error('Username tersebut sudah digunakan.');

  let rowNumber = -1;
  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][idx.employee_id] || '') === employeeId &&
      String(data[i][idx.role] || 'Employee').toLowerCase() === 'employee'
    ) {
      rowNumber = i + 1;
      break;
    }
  }

  if (rowNumber < 0) throw new Error('Employee tidak ditemukan.');

  const updates = {
    name,
    division,
    username,
    phone,
    email,
    account_status: status
  };

  Object.keys(updates).forEach(key => {
    if (idx[key] != null) {
      sheet.getRange(rowNumber, idx[key] + 1).setValue(updates[key]);
    }
  });

  if (status === 'Inactive') {
    revokeEmployeeSessions_(ss, employeeId);
  }

  return getAdminEmployeeManagementData(sessionToken);
}

function setAdminEmployeeStatus(sessionToken, employeeId, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const id = String(employeeId || '').trim();
  const nextStatus = String(status || '').trim();

  if (!id) throw new Error('Employee ID tidak valid.');
  if (!['Active', 'Inactive'].includes(nextStatus)) {
    throw new Error('Status akun tidak valid.');
  }

  const sheet = ss.getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idxId = headers.indexOf('employee_id');
  const idxRole = headers.indexOf('role');
  const idxStatus = headers.indexOf('account_status');

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][idxId] || '') === id &&
      String(data[i][idxRole] || 'Employee').toLowerCase() === 'employee'
    ) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(nextStatus);

      if (nextStatus === 'Inactive') {
        revokeEmployeeSessions_(ss, id);
      }

      return getAdminEmployeeManagementData(sessionToken);
    }
  }

  throw new Error('Employee tidak ditemukan.');
}

function resetAdminEmployeePassword(sessionToken, employeeId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const id = String(employeeId || '').trim();
  if (!id) throw new Error('Employee ID tidak valid.');

  const sheet = ss.getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idxId = headers.indexOf('employee_id');
  const idxRole = headers.indexOf('role');
  const idxPassword = headers.indexOf('password');
  const idxUsername = headers.indexOf('username');

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][idxId] || '') === id &&
      String(data[i][idxRole] || 'Employee').toLowerCase() === 'employee'
    ) {
      const temporaryPassword =
        'Sns!' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);

      sheet.getRange(i + 1, idxPassword + 1)
        .setValue(hashPassword_(temporaryPassword));

      revokeEmployeeSessions_(ss, id);

      const idxName = headers.indexOf('name');

      return {
        employeeId: id,
        name: idxName >= 0 ? String(data[i][idxName] || '') : '',
        username: String(data[i][idxUsername] || ''),
        temporaryPassword
      };
    }
  }

  throw new Error('Employee tidak ditemukan.');
}

function nextEmployeeId_(rows) {
  const numbers = rows
    .map(r => String(r.employee_id || '').trim())
    .map(id => {
      const m = id.match(/^EMP(\d+)$/i);
      return m ? Number(m[1]) : 0;
    })
    .filter(n => n > 0);

  const next = (numbers.length ? Math.max.apply(null, numbers) : 0) + 1;
  return 'EMP' + String(next).padStart(3, '0');
}

function revokeEmployeeSessions_(ss, employeeId) {
  const sheet = ss.getSheetByName('AuthSessions');
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idxEmployee = headers.indexOf('employee_id');

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idxEmployee] || '') === String(employeeId || '')) {
      sheet.deleteRow(i + 1);
    }
  }
}


function getAdminAttendanceMonitoringData(sessionToken, dateKey, periodMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const mode = String(periodMode || 'Today') === 'Month' ? 'Month' : 'Today';
  const selectedDate = parseDateKeyLocal_(dateKey || formatDateKey_(new Date()));
  const selectedKey = formatDateKey_(selectedDate);

  let periodStart;
  let periodEndExclusive;
  let periodLabel;

  if (mode === 'Month') {
    periodStart = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), 1, 0, 0, 0, 0);
    periodEndExclusive = new Date(selectedDate.getFullYear(), selectedDate.getMonth() + 1, 1, 0, 0, 0, 0);
    periodLabel = Utilities.formatDate(periodStart, APP.TIMEZONE, 'MMMM yyyy');
  } else {
    periodStart = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), selectedDate.getDate(), 0, 0, 0, 0);
    periodEndExclusive = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), selectedDate.getDate() + 1, 0, 0, 0, 0);
    periodLabel = Utilities.formatDate(periodStart, APP.TIMEZONE, 'dd MMM yyyy');
  }

  const now = new Date();

  const employees = sheetObjects_(ss.getSheetByName('Employees'))
    .filter(r => String(r.role || 'Employee').trim().toLowerCase() === 'employee')
    .map(r => ({
      id: String(r.employee_id || ''),
      name: String(r.name || ''),
      division: String(r.division || ''),
      username: String(r.username || ''),
      accountStatus: String(r.account_status || 'Active'),
      profilePhotoUrl: String(r.profile_photo_url || '')
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const employeeMap = {};
  employees.forEach(e => employeeMap[e.id] = e);

  const attendanceRows = sheetObjects_(ss.getSheetByName('Attendance'))
    .filter(r => r.date && isWithin_(r.date, periodStart, periodEndExclusive))
    .sort((a, b) => new Date(b.check_in || b.date || 0) - new Date(a.check_in || a.date || 0));

  const visitRows = sheetObjects_(ss.getSheetByName('VisitReports'))
    .filter(r => r.date && isWithin_(r.date, periodStart, periodEndExclusive))
    .sort((a, b) => new Date(b.created_at || b.visit_time || b.date || 0) - new Date(a.created_at || a.visit_time || a.date || 0));

  const jobRows = sheetObjects_(ss.getSheetByName('JobReport'))
    .filter(r => r.date && isWithin_(r.date, periodStart, periodEndExclusive));

  const attendanceByEmployee = {};
  attendanceRows.forEach(r => {
    const id = String(r.employee_id || '');
    if (!attendanceByEmployee[id]) attendanceByEmployee[id] = [];
    attendanceByEmployee[id].push(r);
  });

  const visitCountByEmployee = {};
  visitRows.forEach(r => {
    const id = String(r.employee_id || '');
    visitCountByEmployee[id] = (visitCountByEmployee[id] || 0) + 1;
  });

  const jobCountByEmployee = {};
  jobRows.forEach(r => {
    const id = String(r.employee_id || '');
    jobCountByEmployee[id] = (jobCountByEmployee[id] || 0) + 1;
  });

  const rows = employees.map(employee => {
    const employeeAttendance = attendanceByEmployee[employee.id] || [];
    const accountActive = employee.accountStatus.toLowerCase() === 'active';

    if (mode === 'Month') {
      let wfoDays = 0;
      let wfhDays = 0;
      let checkedInDays = 0;
      let checkedOutDays = 0;
      let missingCheckoutDays = 0;
      let durationSeconds = 0;

      employeeAttendance.forEach(attendance => {
        if (attendance.check_in) checkedInDays++;
        if (attendance.check_out) checkedOutDays++;
        if (attendance.check_in && !attendance.check_out) missingCheckoutDays++;

        if (String(attendance.work_mode || '') === 'Work From Office') wfoDays++;
        if (String(attendance.work_mode || '') === 'Work From Home') wfhDays++;

        if (attendance.check_in) {
          const checkIn = attendance.check_in instanceof Date
            ? attendance.check_in
            : new Date(attendance.check_in);
          const checkOut = attendance.check_out
            ? (attendance.check_out instanceof Date ? attendance.check_out : new Date(attendance.check_out))
            : null;

          if (checkOut && !isNaN(checkIn.getTime()) && !isNaN(checkOut.getTime())) {
            durationSeconds += Math.max(0, Math.floor((checkOut.getTime() - checkIn.getTime()) / 1000));
          }
        }
      });

      return {
        employeeId: employee.id,
        name: employee.name,
        division: employee.division,
        username: employee.username,
        accountStatus: employee.accountStatus,
        profilePhotoUrl: employee.profilePhotoUrl,
        wfoDays,
        wfhDays,
        checkedInDays,
        checkedOutDays,
        missingCheckoutDays,
        durationSeconds,
        visitCount: Number(visitCountByEmployee[employee.id] || 0),
        jobReportCount: Number(jobCountByEmployee[employee.id] || 0),
        attendanceStatus: accountActive
          ? (checkedInDays > 0 ? 'Has Attendance' : 'No Attendance')
          : 'Inactive'
      };
    }

    const attendance = employeeAttendance[0] || null;
    let checkInIso = null;
    let checkOutIso = null;
    let durationSeconds = 0;
    let status = accountActive ? 'Not Checked In' : 'Inactive';

    if (attendance && attendance.check_in) {
      const checkIn = attendance.check_in instanceof Date
        ? attendance.check_in
        : new Date(attendance.check_in);
      const checkOut = attendance.check_out
        ? (attendance.check_out instanceof Date ? attendance.check_out : new Date(attendance.check_out))
        : null;

      checkInIso = !isNaN(checkIn.getTime()) ? checkIn.toISOString() : null;
      checkOutIso = checkOut && !isNaN(checkOut.getTime()) ? checkOut.toISOString() : null;

      const isToday = selectedKey === formatDateKey_(now);
      const end = checkOut || (isToday ? now : checkIn);

      durationSeconds = Math.max(
        0,
        Math.floor((end.getTime() - checkIn.getTime()) / 1000)
      );

      status = checkOut ? 'Checked Out' : 'Working';
    }

    return {
      employeeId: employee.id,
      name: employee.name,
      division: employee.division,
      username: employee.username,
      accountStatus: employee.accountStatus,
      profilePhotoUrl: employee.profilePhotoUrl,
      attendanceId: attendance ? String(attendance.attendance_id || '') : '',
      workMode: attendance ? String(attendance.work_mode || '') : '',
      checkInIso,
      checkOutIso,
      durationSeconds,
      attendanceStatus: status,
      visitCount: Number(visitCountByEmployee[employee.id] || 0),
      jobReportCount: Number(jobCountByEmployee[employee.id] || 0)
    };
  });

  const activeRows = rows.filter(r => r.accountStatus.toLowerCase() === 'active');

  const visits = visitRows.map(r => {
    const employee = employeeMap[String(r.employee_id || '')] || {};
    return {
      id: String(r.visit_id || ''),
      employeeId: String(r.employee_id || ''),
      employeeName: String(employee.name || r.employee_id || ''),
      division: String(employee.division || ''),
      date: dateKeyFromValue_(r.date),
      visitTime: r.visit_time instanceof Date
        ? r.visit_time.toISOString()
        : String(r.visit_time || ''),
      purpose: String(r.purpose || ''),
      lat: r.lat === '' || r.lat == null ? null : Number(r.lat),
      lng: r.lng === '' || r.lng == null ? null : Number(r.lng),
      photos: [
        String(r.photo_1_url || ''),
        String(r.photo_2_url || ''),
        String(r.photo_3_url || '')
      ].filter(Boolean)
    };
  });

  let summary;

  if (mode === 'Month') {
    summary = {
      totalActiveEmployees: activeRows.length,
      totalWfo: activeRows.reduce((sum, r) => sum + Number(r.wfoDays || 0), 0),
      totalWfh: activeRows.reduce((sum, r) => sum + Number(r.wfhDays || 0), 0),
      checkedIn: activeRows.reduce((sum, r) => sum + Number(r.checkedInDays || 0), 0),
      checkedOut: activeRows.reduce((sum, r) => sum + Number(r.checkedOutDays || 0), 0),
      notCheckedOut: activeRows.reduce((sum, r) => sum + Number(r.missingCheckoutDays || 0), 0),
      totalVisits: visits.length
    };
  } else {
    summary = {
      totalActiveEmployees: activeRows.length,
      working: activeRows.filter(r => r.attendanceStatus === 'Working').length,
      checkedOut: activeRows.filter(r => r.attendanceStatus === 'Checked Out').length,
      notCheckedIn: activeRows.filter(r => r.attendanceStatus === 'Not Checked In').length,
      notCheckedOut: activeRows.filter(r => r.checkInIso && !r.checkOutIso).length,
      totalWfo: activeRows.filter(r => r.workMode === 'Work From Office').length,
      totalWfh: activeRows.filter(r => r.workMode === 'Work From Home').length,
      checkedIn: activeRows.filter(r => Boolean(r.checkInIso)).length,
      totalVisits: visits.length
    };
  }

  return {
    date: selectedKey,
    mode,
    period: {
      start: periodStart.toISOString(),
      endExclusive: periodEndExclusive.toISOString(),
      label: periodLabel
    },
    rows,
    visits,
    divisions: Array.from(new Set(
      employees.map(e => e.division).filter(Boolean)
    )).sort(),
    employees: employees.map(e => ({
      id: e.id,
      name: e.name,
      division: e.division,
      accountStatus: e.accountStatus
    })),
    summary
  };
}

function getAdminAttendanceDetail(sessionToken, employeeId, dateKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const id = String(employeeId || '').trim();
  if (!id) throw new Error('Employee ID tidak valid.');

  const employee = getEmployee_(ss, id);
  if (!employee.employee_id || String(employee.role || 'Employee').toLowerCase() !== 'employee') {
    throw new Error('Employee tidak ditemukan.');
  }

  const selectedDate = parseDateKeyLocal_(dateKey || formatDateKey_(new Date()));
  const selectedKey = formatDateKey_(selectedDate);
  const now = new Date();

  const attendanceRows = sheetObjects_(ss.getSheetByName('Attendance'))
    .filter(r => String(r.employee_id || '') === id)
    .sort((a, b) => {
      const da = new Date(a.check_in || a.date || 0).getTime();
      const db = new Date(b.check_in || b.date || 0).getTime();
      return db - da;
    });

  const attendance = attendanceRows.find(r =>
    dateKeyFromValue_(r.date) === selectedKey
  ) || null;

  let attendanceDetail = null;

  if (attendance) {
    const checkIn = attendance.check_in
      ? (attendance.check_in instanceof Date ? attendance.check_in : new Date(attendance.check_in))
      : null;
    const checkOut = attendance.check_out
      ? (attendance.check_out instanceof Date ? attendance.check_out : new Date(attendance.check_out))
      : null;

    const isToday = selectedKey === formatDateKey_(now);
    const end = checkOut || (isToday ? now : checkIn);
    const durationSeconds = checkIn && end
      ? Math.max(0, Math.floor((end.getTime() - checkIn.getTime()) / 1000))
      : 0;

    attendanceDetail = {
      id: String(attendance.attendance_id || ''),
      date: selectedKey,
      workMode: String(attendance.work_mode || ''),
      checkInIso: checkIn && !isNaN(checkIn.getTime()) ? checkIn.toISOString() : null,
      checkOutIso: checkOut && !isNaN(checkOut.getTime()) ? checkOut.toISOString() : null,
      durationSeconds,
      status: checkOut ? 'Checked Out' : 'Working',
      checkInLat: attendance.check_in_lat === '' || attendance.check_in_lat == null
        ? null : Number(attendance.check_in_lat),
      checkInLng: attendance.check_in_lng === '' || attendance.check_in_lng == null
        ? null : Number(attendance.check_in_lng),
      checkInSelfieUrl: String(attendance.check_in_selfie_url || ''),
      checkOutLat: attendance.check_out_lat === '' || attendance.check_out_lat == null
        ? null : Number(attendance.check_out_lat),
      checkOutLng: attendance.check_out_lng === '' || attendance.check_out_lng == null
        ? null : Number(attendance.check_out_lng)
    };
  }

  const jobs = sheetObjects_(ss.getSheetByName('JobReport'))
    .filter(r =>
      String(r.employee_id || '') === id &&
      dateKeyFromValue_(r.date) === selectedKey
    )
    .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0))
    .map(r => ({
      id: String(r.report_id || ''),
      title: String(r.job_title || ''),
      description: String(r.description || ''),
      documentationUrls: String(r.documentation_url || '')
        .split('\n')
        .map(x => x.trim())
        .filter(Boolean),
      createdAt: r.created_at instanceof Date
        ? r.created_at.toISOString()
        : String(r.created_at || '')
    }));

  const visits = sheetObjects_(ss.getSheetByName('VisitReports'))
    .filter(r =>
      String(r.employee_id || '') === id &&
      dateKeyFromValue_(r.date) === selectedKey
    )
    .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0))
    .map(r => ({
      id: String(r.visit_id || ''),
      purpose: String(r.purpose || ''),
      visitTime: r.visit_time instanceof Date
        ? r.visit_time.toISOString()
        : String(r.visit_time || ''),
      lat: r.lat === '' || r.lat == null ? null : Number(r.lat),
      lng: r.lng === '' || r.lng == null ? null : Number(r.lng),
      photos: [
        String(r.photo_1_url || ''),
        String(r.photo_2_url || ''),
        String(r.photo_3_url || '')
      ].filter(Boolean)
    }));

  const history = attendanceRows
    .filter(r => dateKeyFromValue_(r.date))
    .slice(0, 10)
    .map(r => {
      const checkIn = r.check_in
        ? (r.check_in instanceof Date ? r.check_in : new Date(r.check_in))
        : null;
      const checkOut = r.check_out
        ? (r.check_out instanceof Date ? r.check_out : new Date(r.check_out))
        : null;

      const durationSeconds = checkIn && checkOut
        ? Math.max(0, Math.floor((checkOut.getTime() - checkIn.getTime()) / 1000))
        : 0;

      return {
        date: dateKeyFromValue_(r.date),
        workMode: String(r.work_mode || ''),
        checkInIso: checkIn && !isNaN(checkIn.getTime()) ? checkIn.toISOString() : null,
        checkOutIso: checkOut && !isNaN(checkOut.getTime()) ? checkOut.toISOString() : null,
        durationSeconds,
        status: checkOut ? 'Checked Out' : 'Working'
      };
    });

  return {
    employee: {
      id,
      name: String(employee.name || ''),
      division: String(employee.division || ''),
      username: String(employee.username || ''),
      accountStatus: String(employee.account_status || 'Active'),
      profilePhotoUrl: String(employee.profile_photo_url || '')
    },
    date: selectedKey,
    attendance: attendanceDetail,
    jobReport: {
      total: jobs.length,
      documented: jobs.filter(r => r.documentationUrls.length > 0).length,
      rows: jobs
    },
    visits,
    history
  };
}

function parseDateKeyLocal_(dateKey) {
  const value = String(dateKey || '').trim();
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error('Tanggal tidak valid.');

  const date = new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    12, 0, 0, 0
  );

  if (isNaN(date.getTime())) throw new Error('Tanggal tidak valid.');
  return date;
}


function getAdminFinanceData(sessionToken, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const now = new Date();
  let selectedMonth = String(monthKey || '').trim();

  if (!/^\d{4}-\d{2}$/.test(selectedMonth)) {
    const currentCycle = getReportingCycle_(now);
    selectedMonth = Utilities.formatDate(currentCycle.start, APP.TIMEZONE, 'yyyy-MM');
  }

  const parts = selectedMonth.split('-').map(Number);
  const periodStart = new Date(parts[0], parts[1] - 1, 5, 0, 0, 0, 0);
  const periodEnd = new Date(parts[0], parts[1], 5, 0, 0, 0, 0);
  const periodLast = new Date(periodEnd.getTime() - 1);
  const periodLabel =
    Utilities.formatDate(periodStart, APP.TIMEZONE, 'dd MMM') +
    ' - ' +
    Utilities.formatDate(periodLast, APP.TIMEZONE, 'dd MMM yyyy');

  const employees = sheetObjects_(ss.getSheetByName('Employees'))
    .filter(r => String(r.role || 'Employee').trim().toLowerCase() === 'employee')
    .map(r => ({
      id: String(r.employee_id || ''),
      name: String(r.name || ''),
      division: String(r.division || ''),
      status: String(r.account_status || 'Active'),
      profilePhotoUrl: String(r.profile_photo_url || '')
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const employeeMap = {};
  employees.forEach(e => employeeMap[e.id] = e);

  const attendanceRows = sheetObjects_(ss.getSheetByName('Attendance'))
    .filter(r => r.date && isWithin_(r.date, periodStart, periodEnd));

  const wfoDaysByEmployee = {};
  attendanceRows.forEach(r => {
    const id = String(r.employee_id || '');
    if (String(r.work_mode || '') === 'Work From Office') {
      wfoDaysByEmployee[id] = (wfoDaysByEmployee[id] || 0) + 1;
    }
  });

  const elementRows = sheetObjects_(ss.getSheetByName('FinanceElements'))
    .sort((a, b) =>
      String(a.employee_id || '').localeCompare(String(b.employee_id || '')) ||
      Number(a.sort_order || 0) - Number(b.sort_order || 0)
    );

  const elementsByEmployee = {};
  elementRows.forEach(r => {
    const id = String(r.employee_id || '');
    if (!elementsByEmployee[id]) elementsByEmployee[id] = [];
    elementsByEmployee[id].push(r);
  });

  const reimburseRows = sheetObjects_(ss.getSheetByName('Reimbursements'))
    .filter(r => r.date && isWithin_(r.date, periodStart, periodEnd))
    .sort((a, b) => new Date(b.created_at || b.date || 0) - new Date(a.created_at || a.date || 0));

  const claimRows = sheetObjects_(ss.getSheetByName('Claims'))
    .filter(r => r.date && isWithin_(r.date, periodStart, periodEnd))
    .sort((a, b) => new Date(b.created_at || b.date || 0) - new Date(a.created_at || a.date || 0));

  const approvedReimByEmployee = {};
  reimburseRows.forEach(r => {
    if (String(r.status || 'Pending').toLowerCase() !== 'approved') return;
    const id = String(r.employee_id || '');
    approvedReimByEmployee[id] = (approvedReimByEmployee[id] || 0) + Number(r.amount || 0);
  });

  const approvedClaimByEmployee = {};
  claimRows.forEach(r => {
    if (String(r.status || 'Pending').toLowerCase() !== 'approved') return;
    const id = String(r.employee_id || '');
    approvedClaimByEmployee[id] = (approvedClaimByEmployee[id] || 0) + Number(r.amount || 0);
  });

  const payrollRows = employees.map(employee => {
    const elements = elementsByEmployee[employee.id] || [];
    const wfoDays = Number(wfoDaysByEmployee[employee.id] || 0);

    const paymentElements = elements.map(r => {
      const type = String(r.calculation_type || 'reference').trim().toLowerCase();
      const amount = Number(r.amount || 0);
      const active = String(r.active).toLowerCase() !== 'false';

      let calculatedAmount = amount;
      let includedInEstimate = false;

      if (active && type === 'fixed') {
        calculatedAmount = amount;
        includedInEstimate = true;
      } else if (active && type === 'wfo_rate') {
        calculatedAmount = amount * wfoDays;
        includedInEstimate = true;
      }

      return {
        id: String(r.element_id || ''),
        employeeId: employee.id,
        label: String(r.label || 'Payment Element'),
        type,
        amount,
        active,
        sortOrder: Number(r.sort_order || 0),
        calculatedAmount,
        includedInEstimate
      };
    });

    const baseEstimate = paymentElements
      .filter(x => x.includedInEstimate)
      .reduce((sum, x) => sum + Number(x.calculatedAmount || 0), 0);

    const approvedReimbursement = Number(approvedReimByEmployee[employee.id] || 0);
    const approvedClaim = Number(approvedClaimByEmployee[employee.id] || 0);
    const totalPayable = baseEstimate + approvedReimbursement + approvedClaim;

    return {
      employeeId: employee.id,
      name: employee.name,
      division: employee.division,
      status: employee.status,
      profilePhotoUrl: employee.profilePhotoUrl,
      wfoDays,
      baseEstimate,
      approvedReimbursement,
      approvedClaim,
      totalPayable,
      paymentElements
    };
  });

  const reimbursements = reimburseRows.map(r => {
    const employee = employeeMap[String(r.employee_id || '')] || {};
    return {
      id: String(r.reimburse_id || ''),
      employeeId: String(r.employee_id || ''),
      employeeName: String(employee.name || r.employee_id || ''),
      division: String(employee.division || ''),
      date: dateKeyFromValue_(r.date),
      category: String(r.category || ''),
      description: String(r.description || ''),
      amount: Number(r.amount || 0),
      proofUrl: String(r.proof_url || ''),
      status: String(r.status || 'Pending'),
      createdAt: r.created_at instanceof Date
        ? r.created_at.toISOString()
        : String(r.created_at || '')
    };
  });

  const claims = claimRows.map(r => {
    const employee = employeeMap[String(r.employee_id || '')] || {};
    return {
      id: String(r.claim_id || ''),
      employeeId: String(r.employee_id || ''),
      employeeName: String(employee.name || r.employee_id || ''),
      division: String(employee.division || ''),
      claimType: String(r.claim_type || 'Achievement'),
      date: dateKeyFromValue_(r.date),
      name: String(r.claim_name || ''),
      description: String(r.description || ''),
      amount: Number(r.amount || 0),
      proofUrl: String(r.proof_url || ''),
      status: String(r.status || 'Pending'),
      createdAt: r.created_at instanceof Date
        ? r.created_at.toISOString()
        : String(r.created_at || '')
    };
  });

  const totalBaseEstimate = payrollRows
    .reduce((sum, row) => sum + Number(row.baseEstimate || 0), 0);
  const totalApprovedReimbursement = reimbursements
    .filter(r => r.status.toLowerCase() === 'approved')
    .reduce((sum, row) => sum + Number(row.amount || 0), 0);
  const totalApprovedClaim = claims
    .filter(r => r.status.toLowerCase() === 'approved')
    .reduce((sum, row) => sum + Number(row.amount || 0), 0);
  const totalPayable = payrollRows
    .reduce((sum, row) => sum + Number(row.totalPayable || 0), 0);

  return {
    month: selectedMonth,
    period: {
      start: periodStart.toISOString(),
      endExclusive: periodEnd.toISOString(),
      label: periodLabel
    },
    summary: {
      totalEmployees: employees.length,
      totalBaseEstimate,
      totalApprovedReimbursement,
      totalApprovedClaim,
      totalPayable,
      pendingReimbursements: reimbursements.filter(r => r.status.toLowerCase() === 'pending').length,
      pendingClaims: claims.filter(r => r.status.toLowerCase() === 'pending').length
    },
    employees,
    payrollRows,
    reimbursements,
    claims
  };
}

function saveAdminFinanceElement(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  if (!payload) throw new Error('Data Payment Element tidak ditemukan.');

  const employeeId = String(payload.employeeId || '').trim();
  const elementId = String(payload.id || '').trim();
  const label = String(payload.label || '').trim();
  const type = String(payload.type || '').trim().toLowerCase();
  const amount = Number(payload.amount || 0);
  const active = payload.active !== false && String(payload.active).toLowerCase() !== 'false';

  if (!employeeId) throw new Error('Pilih Employee.');
  if (!label) throw new Error('Nama Payment Element wajib diisi.');
  if (!['reference', 'fixed', 'wfo_rate'].includes(type)) {
    throw new Error('Calculation Type tidak valid.');
  }
  if (!isFinite(amount) || amount < 0) {
    throw new Error('Nominal Payment Element tidak valid.');
  }

  const employee = getEmployee_(ss, employeeId);
  if (!employee.employee_id || String(employee.role || 'Employee').toLowerCase() !== 'employee') {
    throw new Error('Employee tidak ditemukan.');
  }

  const sheet = ss.getSheetByName('FinanceElements');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  if (elementId) {
    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][idx.element_id] || '') === elementId &&
        String(data[i][idx.employee_id] || '') === employeeId
      ) {
        sheet.getRange(i + 1, idx.label + 1).setValue(label);
        sheet.getRange(i + 1, idx.calculation_type + 1).setValue(type);
        sheet.getRange(i + 1, idx.amount + 1).setValue(amount);
        sheet.getRange(i + 1, idx.active + 1).setValue(active);
        return getAdminFinanceData(sessionToken, payload.month);
      }
    }
    throw new Error('Payment Element tidak ditemukan.');
  }

  const employeeElements = sheetObjects_(sheet)
    .filter(r => String(r.employee_id || '') === employeeId);
  const nextSort = employeeElements.length
    ? Math.max.apply(null, employeeElements.map(r => Number(r.sort_order || 0))) + 1
    : 1;

  const row = headers.map(header => {
    if (header === 'element_id') return 'FIN-ELEM-' + Utilities.getUuid();
    if (header === 'employee_id') return employeeId;
    if (header === 'label') return label;
    if (header === 'calculation_type') return type;
    if (header === 'amount') return amount;
    if (header === 'active') return active;
    if (header === 'sort_order') return nextSort;
    return '';
  });

  sheet.appendRow(row);
  return getAdminFinanceData(sessionToken, payload.month);
}

function deleteAdminFinanceElement(sessionToken, elementId, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const id = String(elementId || '').trim();
  if (!id) throw new Error('Payment Element ID tidak valid.');

  const sheet = ss.getSheetByName('FinanceElements');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Payment Element tidak ditemukan.');

  const headers = data[0].map(String);
  const idxId = headers.indexOf('element_id');

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idxId] || '') === id) {
      sheet.deleteRow(i + 1);
      return getAdminFinanceData(sessionToken, monthKey);
    }
  }

  throw new Error('Payment Element tidak ditemukan.');
}

function updateAdminReimbursementStatus(sessionToken, reimburseId, status, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const id = String(reimburseId || '').trim();
  const nextStatus = String(status || '').trim();

  if (!id) throw new Error('Reimbursement ID tidak valid.');
  if (!['Pending', 'Approved', 'Rejected'].includes(nextStatus)) {
    throw new Error('Status Reimbursement tidak valid.');
  }

  const sheet = ss.getSheetByName('Reimbursements');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Reimbursement tidak ditemukan.');

  const headers = data[0].map(String);
  const idxId = headers.indexOf('reimburse_id');
  const idxStatus = headers.indexOf('status');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId] || '') === id) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(nextStatus);
      return getAdminFinanceData(sessionToken, monthKey);
    }
  }

  throw new Error('Reimbursement tidak ditemukan.');
}

function updateAdminClaimStatus(sessionToken, claimId, status, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);
  requireAdminSession_(sessionToken);

  const id = String(claimId || '').trim();
  const nextStatus = String(status || '').trim();

  if (!id) throw new Error('Claim ID tidak valid.');
  if (!['Pending', 'Approved', 'Rejected'].includes(nextStatus)) {
    throw new Error('Status Claim tidak valid.');
  }

  const sheet = ss.getSheetByName('Claims');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Claim tidak ditemukan.');

  const headers = data[0].map(String);
  const idxId = headers.indexOf('claim_id');
  const idxStatus = headers.indexOf('status');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId] || '') === id) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(nextStatus);
      return getAdminFinanceData(sessionToken, monthKey);
    }
  }

  throw new Error('Claim tidak ditemukan.');
}

function getDashboardData(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const employee = getEmployee_(ss, employeeId);
  const today = new Date();
  const todayKey = formatDateKey_(today);

  // One read per sheet. These same arrays are reused by the KPI calculation.
  const attendanceAll = sheetObjects_(ss.getSheetByName('Attendance'));
  const reportAll = sheetObjects_(ss.getSheetByName('JobReport'));
  const taskAll = sheetObjects_(ss.getSheetByName('Tasks'));

  const todayAttendance = attendanceAll
    .filter(r => String(r.employee_id) === employeeId && dateKeyFromValue_(r.date) === todayKey)
    .sort((a, b) => new Date(b.check_in || 0) - new Date(a.check_in || 0))[0] || null;

  let checkInIso = null;
  let checkOutIso = null;
  let estCheckoutIso = null;
  let durationSeconds = 0;

  if (todayAttendance && todayAttendance.check_in) {
    const checkIn = new Date(todayAttendance.check_in);
    const checkOut = todayAttendance.check_out ? new Date(todayAttendance.check_out) : null;
    const endTime = checkOut || new Date();

    checkInIso = checkIn.toISOString();
    checkOutIso = checkOut ? checkOut.toISOString() : null;
    estCheckoutIso = new Date(checkIn.getTime() + APP.TARGET_HOURS * 3600 * 1000).toISOString();
    durationSeconds = Math.max(0, Math.floor((endTime.getTime() - checkIn.getTime()) / 1000));
  }

  const weekdayMap = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const todayWeekday = weekdayMap[today.getDay()];

  const taskRows = taskAll
    .filter(r => {
      if (!taskVisibleToEmployee_(r, employeeId)) return false;
      if (String(r.status).toLowerCase() === 'done') return false;

      const kind = String(r.task_kind || '').trim() || 'Project';

      if (kind === 'Routine') {
        const days = String(r.routine_days || '')
          .split(',')
          .map(x => x.trim())
          .filter(Boolean);
        return days.includes(todayWeekday);
      }

      let deadline = null;
      if (r.deadline) {
        deadline = r.deadline instanceof Date ? r.deadline : new Date(r.deadline);
      } else if (r.task_date) {
        deadline = r.task_date instanceof Date ? r.task_date : new Date(r.task_date);
      }

      if (!deadline || isNaN(deadline.getTime())) return false;

      // Project due today or overdue remains visible as reminder.
      return dateKeyFromValue_(deadline) <= todayKey;
    })
    .map(r => {
      const kind = String(r.task_kind || '').trim() || 'Project';
      let deadlineIso = '';

      if (kind === 'Project') {
        let d = null;
        if (r.deadline) d = r.deadline instanceof Date ? r.deadline : new Date(r.deadline);
        else if (r.task_date) d = r.task_date instanceof Date ? r.task_date : new Date(r.task_date);
        if (d && !isNaN(d.getTime())) deadlineIso = d.toISOString();
      }

      return {
        id: String(r.task_id || ''),
        taskKind: kind,
        source: String(r.source || 'Personal'),
        title: String(r.title || ''),
        deadlineIso,
        routineDays: String(r.routine_days || '').split(',').map(x => x.trim()).filter(Boolean),
        priority: kind === 'Project' ? String(r.priority || 'Medium') : '',
        status: String(r.status || 'To Do')
      };
    });

  const kpi = getWeeklyKpiFromRows_(
    attendanceAll,
    reportAll,
    taskAll,
    employeeId
  );

  return {
    employee: {
      id: employee.employee_id || employeeId,
      name: employee.name || 'Employee',
      division: employee.division || '',
      profilePhotoUrl: String(employee.profile_photo_url || '')
    },
    timezone: APP.TIMEZONE,
    work: {
      hasAttendance: !!todayAttendance,
      workMode: todayAttendance ? String(todayAttendance.work_mode || '') : '',
      checkInIso,
      checkOutIso,
      estCheckoutIso,
      durationSeconds,
      targetSeconds: APP.TARGET_HOURS * 3600,
      blueAfterSeconds: APP.OVERTIME_BLUE_AFTER_HOURS * 3600
    },
    tasks: taskRows,
    kpi
  };
}

function getJobReportPageData(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const now = new Date();
  const todayKey = formatDateKey_(now);
  const cycle = getReportingCycle_(now);

  const rows = sheetObjects_(ss.getSheetByName('JobReport'))
    .filter(r => String(r.employee_id) === employeeId)
    .sort((a, b) => new Date(b.created_at || b.date || 0) - new Date(a.created_at || a.date || 0));

  const normalizeDocs = value => {
    if (!value) return [];
    return String(value).split('\n').map(x => x.trim()).filter(Boolean);
  };

  const mapReport = r => ({
    id: String(r.report_id || ''),
    date: dateKeyFromValue_(r.date),
    title: String(r.job_title || ''),
    description: String(r.description || ''),
    documentation: normalizeDocs(r.documentation_url),
    createdAt: r.created_at instanceof Date ? r.created_at.toISOString() : String(r.created_at || '')
  });

  const todayReports = rows
    .filter(r => dateKeyFromValue_(r.date) === todayKey)
    .map(mapReport);

  // History hanya menampilkan periode aktif: tanggal 5 s.d. tanggal 4 berikutnya.
  // Saat masuk tanggal 5, history periode sebelumnya otomatis tidak ditampilkan.
  const history = rows
    .filter(r => isWithin_(r.date, cycle.start, cycle.end))
    .slice(0, 100)
    .map(mapReport);

  const documentedJobs = todayReports.filter(r => r.documentation.length > 0).length;

  return {
    period: {
      start: cycle.start.toISOString(),
      endExclusive: cycle.end.toISOString(),
      label: cycle.label
    },
    today: {
      reports: todayReports,
      totalJobs: todayReports.length,
      documentedJobs,
      requiredJobs: 3,
      readyForCheckout: todayReports.length >= 3 && documentedJobs >= 3
    },
    history
  };
}

function saveJobReport(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload) throw new Error('Data Daily Job Report tidak ditemukan.');

  const title = String(payload.title || '').trim();
  const description = String(payload.description || '').trim();

  if (!title) throw new Error('Nama pekerjaan wajib diisi.');
  if (!description) throw new Error('Deskripsi pekerjaan wajib diisi.');

  const docs = Array.isArray(payload.documentation)
    ? payload.documentation.filter(Boolean).slice(0, 3)
    : [];

  if (docs.length < 1) {
    throw new Error('Tambahkan minimal 1 dokumentasi untuk Daily Job Report ini.');
  }

  const now = new Date();
  const urls = docs.map((dataUrl, index) =>
    saveDataUrlToDrive_(
      dataUrl,
      employeeId + '_job_' + formatFileStamp_(now) + '_' + (index + 1) + '.jpg'
    )
  );

  ss.getSheetByName('JobReport').appendRow([
    'JOB-' + Utilities.getUuid(),
    employeeId,
    new Date(now.getFullYear(), now.getMonth(), now.getDate()),
    title,
    description,
    urls.join('\n'),
    now
  ]);

  return getJobReportPageData(sessionToken);
}

function deleteJobReport(sessionToken, reportId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const id = String(reportId || '').trim();
  if (!id) throw new Error('Report ID tidak valid.');

  const sheet = ss.getSheetByName('JobReport');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Daily Job Report tidak ditemukan.');

  const headers = data[0].map(String);
  const idxReportId = headers.indexOf('report_id');
  const idxEmployeeId = headers.indexOf('employee_id');
  if (idxReportId < 0 || idxEmployeeId < 0) throw new Error('Struktur sheet JobReport tidak sesuai.');

  for (let i = data.length - 1; i >= 1; i--) {
    if (
      String(data[i][idxReportId]) === id &&
      String(data[i][idxEmployeeId]) === employeeId
    ) {
      sheet.deleteRow(i + 1);
      return getJobReportPageData(sessionToken);
    }
  }

  throw new Error('Daily Job Report tidak ditemukan.');
}

function getAttendancePageData(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const todayKey = formatDateKey_(new Date());

  const attendanceRows = sheetObjects_(ss.getSheetByName('Attendance'))
    .filter(r => String(r.employee_id) === employeeId && dateKeyFromValue_(r.date) === todayKey)
    .sort((a, b) => new Date(b.check_in || 0) - new Date(a.check_in || 0));

  const attendance = attendanceRows[0] || null;

  const reportRows = sheetObjects_(ss.getSheetByName('JobReport'))
    .filter(r => String(r.employee_id) === employeeId && dateKeyFromValue_(r.date) === todayKey);

  const documentedJobs = reportRows.filter(r => String(r.documentation_url || '').trim() !== '').length;

  const visitRows = sheetObjects_(ss.getSheetByName('VisitReports'))
    .filter(r => String(r.employee_id) === employeeId && dateKeyFromValue_(r.date) === todayKey)
    .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));

  let durationSeconds = 0;
  let checkInIso = null;
  let checkOutIso = null;
  let estimatedCheckOutIso = null;

  if (attendance && attendance.check_in) {
    const checkIn = new Date(attendance.check_in);
    const checkOut = attendance.check_out ? new Date(attendance.check_out) : null;
    const end = checkOut || new Date();

    checkInIso = checkIn.toISOString();
    checkOutIso = checkOut ? checkOut.toISOString() : null;
    estimatedCheckOutIso = new Date(checkIn.getTime() + APP.TARGET_HOURS * 3600 * 1000).toISOString();
    durationSeconds = Math.max(0, Math.floor((end.getTime() - checkIn.getTime()) / 1000));
  }

  return {
    attendance: {
      exists: !!attendance,
      id: attendance ? String(attendance.attendance_id || '') : '',
      workMode: attendance ? String(attendance.work_mode || '') : '',
      checkInIso,
      checkOutIso,
      estimatedCheckOutIso,
      durationSeconds,
      targetSeconds: APP.TARGET_HOURS * 3600,
      blueAfterSeconds: APP.OVERTIME_BLUE_AFTER_HOURS * 3600,
      status: attendance ? String(attendance.status || '') : ''
    },
    jobReport: {
      totalJobs: reportRows.length,
      documentedJobs,
      requiredJobs: 3,
      readyForCheckout: reportRows.length >= 3 && documentedJobs >= 3
    },
    visitsToday: visitRows.map(r => ({
      id: String(r.visit_id || ''),
      purpose: String(r.purpose || ''),
      visitTime: r.visit_time instanceof Date ? r.visit_time.toISOString() : String(r.visit_time || ''),
      lat: Number(r.lat || 0),
      lng: Number(r.lng || 0)
    }))
  };
}

function checkInAttendance(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload || !payload.workMode) throw new Error('Pilih Work From Office atau Work From Home.');
  if (!payload.selfieDataUrl) throw new Error('Selfie check-in wajib diambil.');
  validateCoordinates_(payload.lat, payload.lng);

  const todayKey = formatDateKey_(new Date());
  const attendanceSheet = ss.getSheetByName('Attendance');
  const rows = sheetObjects_(attendanceSheet);

  const existing = rows.find(r =>
    String(r.employee_id) === employeeId &&
    dateKeyFromValue_(r.date) === todayKey &&
    !r.check_out
  );

  if (existing) throw new Error('Kamu sudah melakukan check-in hari ini dan belum check-out.');

  const now = new Date();
  const id = 'ATT-' + Utilities.getUuid();
  const selfieUrl = saveDataUrlToDrive_(payload.selfieDataUrl, employeeId + '_checkin_' + formatFileStamp_(now) + '.jpg');

  attendanceSheet.appendRow([
    id,
    employeeId,
    new Date(now.getFullYear(), now.getMonth(), now.getDate()),
    String(payload.workMode),
    now,
    '',
    Number(payload.lat),
    Number(payload.lng),
    selfieUrl,
    '',
    '',
    'Active'
  ]);

  return getAttendancePageData(sessionToken);
}

function checkOutAttendance(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const todayKey = formatDateKey_(new Date());
  const attendanceSheet = ss.getSheetByName('Attendance');

  const data = attendanceSheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Belum ada check-in aktif.');

  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let rowNumber = -1;
  let attendanceRow = null;

  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i];
    if (
      String(r[idx.employee_id]) === employeeId &&
      dateKeyFromValue_(r[idx.date]) === todayKey &&
      !r[idx.check_out]
    ) {
      rowNumber = i + 1;
      attendanceRow = r;
      break;
    }
  }

  if (rowNumber < 0) throw new Error('Tidak ada check-in aktif untuk di-check-out.');

  const reportRows = sheetObjects_(ss.getSheetByName('JobReport'))
    .filter(r => String(r.employee_id) === employeeId && dateKeyFromValue_(r.date) === todayKey);

  const documentedJobs = reportRows.filter(r => String(r.documentation_url || '').trim() !== '').length;

  if (reportRows.length < 3 || documentedJobs < 3) {
    throw new Error(
      'Check-out belum dapat dilakukan. Lengkapi minimal 3 Daily Job Report dan pastikan minimal 3 pekerjaan memiliki dokumentasi.'
    );
  }

  if (payload && payload.lat != null && payload.lng != null) {
    validateCoordinates_(payload.lat, payload.lng);
    attendanceSheet.getRange(rowNumber, idx.check_out_lat + 1).setValue(Number(payload.lat));
    attendanceSheet.getRange(rowNumber, idx.check_out_lng + 1).setValue(Number(payload.lng));
  }

  attendanceSheet.getRange(rowNumber, idx.check_out + 1).setValue(new Date());
  attendanceSheet.getRange(rowNumber, idx.status + 1).setValue('Completed');

  return getAttendancePageData(sessionToken);
}

function submitVisitReport(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload || !String(payload.purpose || '').trim()) {
    throw new Error('Isi keperluan atau tujuan visit.');
  }

  validateCoordinates_(payload.lat, payload.lng);

  const photos = Array.isArray(payload.photos) ? payload.photos.filter(Boolean) : [];
  if (photos.length < 3) {
    throw new Error('Report Visit membutuhkan minimal 3 foto.');
  }

  const now = new Date();

  const urls = photos.slice(0, 3).map((dataUrl, index) =>
    saveDataUrlToDrive_(
      dataUrl,
      employeeId + '_visit_' + formatFileStamp_(now) + '_' + (index + 1) + '.jpg'
    )
  );

  ss.getSheetByName('VisitReports').appendRow([
    'VIS-' + Utilities.getUuid(),
    employeeId,
    new Date(now.getFullYear(), now.getMonth(), now.getDate()),
    now,
    String(payload.purpose).trim(),
    Number(payload.lat),
    Number(payload.lng),
    urls[0] || '',
    urls[1] || '',
    urls[2] || '',
    now
  ]);

  return getAttendancePageData(sessionToken);
}

function saveDataUrlToDrive_(dataUrl, fileName) {
  const match = String(dataUrl || '').match(/^data:(image\/(?:jpeg|jpg|png|webp));base64,(.+)$/);
  if (!match) throw new Error('Format gambar tidak valid.');

  const mimeType = match[1] === 'image/jpg' ? 'image/jpeg' : match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);

  const folder = getOrCreateUploadFolder_();
  const file = folder.createFile(blob);
  return file.getUrl();
}

function saveGenericDataUrlToDrive_(dataUrl, fileName) {
  const match = String(dataUrl || '').match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Format file tidak valid.');

  const mimeType = String(match[1] || '').toLowerCase();
  const allowed = [
    'image/jpeg', 'image/jpg', 'image/png', 'image/webp',
    'application/pdf'
  ];

  if (!allowed.includes(mimeType)) {
    throw new Error('Bukti KPI harus berupa JPG, PNG, WEBP, atau PDF.');
  }

  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5 * 1024 * 1024) {
    throw new Error('Ukuran bukti KPI maksimal 5 MB.');
  }

  const safeName = String(fileName || 'Bukti_KPI')
    .replace(/[^\w.\-() ]+/g, '_')
    .slice(0, 120);

  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  const folder = getOrCreateUploadFolder_();
  const file = folder.createFile(blob);
  return file.getUrl();
}


function getOrCreateUploadFolder_() {
  const folderName = 'Sansis Daily Officer Uploads';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function validateCoordinates_(lat, lng) {
  const latitude = Number(lat);
  const longitude = Number(lng);

  if (!isFinite(latitude) || !isFinite(longitude)) {
    throw new Error('GPS wajib aktif dan lokasi harus berhasil dibaca.');
  }
  if (latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
    throw new Error('Koordinat GPS tidak valid.');
  }
}

function formatFileStamp_(date) {
  return Utilities.formatDate(date, APP.TIMEZONE, 'yyyyMMdd_HHmmss');
}


function getTaskManagerPageData(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;


  const employeeRows = sheetObjects_(ss.getSheetByName('Employees'))
    .filter(r => String(r.employee_id || '').trim() !== '');

  const employeeMap = {};
  employeeRows.forEach(r => {
    const id = String(r.employee_id || '');
    employeeMap[id] = {
      id,
      name: String(r.name || r.employee_id || ''),
      division: String(r.division || '')
    };
  });

  const teamMembers = employeeRows
    .filter(r =>
      String(r.role || 'Employee').trim().toLowerCase() === 'employee' &&
      String(r.account_status || 'Active').trim().toLowerCase() === 'active'
    )
    .map(r => employeeMap[String(r.employee_id || '')])
    .sort((a, b) => a.name.localeCompare(b.name));

  // A task is visible when:
  // 1. it was created by this employee, OR
  // 2. this employee is listed in task_assignees.
  const rows = sheetObjects_(ss.getSheetByName('Tasks'))
    .filter(r => taskVisibleToEmployee_(r, employeeId));

  const tasks = rows.map(r => {
    const taskKind = String(r.task_kind || '').trim() || 'Project';
    const legacyDate = dateKeyFromValue_(r.task_date);
    const legacyEnd = normalizeTime_(r.end_time);
    const ownerId = String(r.employee_id || '');
    const isOwner = ownerId === employeeId;
    let deadlineIso = '';

    if (r.deadline) {
      const d = r.deadline instanceof Date ? r.deadline : new Date(r.deadline);
      if (!isNaN(d.getTime())) deadlineIso = d.toISOString();
    } else if (taskKind === 'Project' && legacyDate) {
      const time = legacyEnd || '23:59';
      const d = new Date(legacyDate + 'T' + time + ':00');
      if (!isNaN(d.getTime())) deadlineIso = d.toISOString();
    }

    const routineDays = String(r.routine_days || '')
      .split(',')
      .map(x => x.trim())
      .filter(Boolean);

    const assigneeIds = taskAssigneeIds_(r);

    const assigneeNames = assigneeIds
      .map(id => employeeMap[id]?.name || id)
      .filter(Boolean);

    return {
      id: String(r.task_id || ''),
      taskKind,
      source: String(r.source || 'Personal'),
      title: String(r.title || ''),
      taskDate: legacyDate,
      priority: taskKind === 'Project' ? String(r.priority || 'Medium') : '',
      status: String(r.status || 'To Do'),
      deadlineIso,
      routineDays,
      assigneeIds,
      assigneeNames,
      ownerId,
      ownerName: employeeMap[ownerId]?.name || ownerId,
      isOwner,
      assignedToMe: !isOwner && assigneeIds.includes(employeeId)
    };
  }).sort((a, b) => {
    if (a.taskKind !== b.taskKind) return a.taskKind === 'Project' ? -1 : 1;
    if (a.taskKind === 'Project') {
      const da = a.deadlineIso ? new Date(a.deadlineIso).getTime() : Number.MAX_SAFE_INTEGER;
      const db = b.deadlineIso ? new Date(b.deadlineIso).getTime() : Number.MAX_SAFE_INTEGER;
      return da - db;
    }
    return a.title.localeCompare(b.title);
  });

  const kpiCycle = getReportingCycle_(new Date());
  const currentCycleStartKey = dateKeyFromValue_(kpiCycle.start);

  const kpi = sheetObjects_(ss.getSheetByName('TaskKPI'))
    .filter(r => {
      if (String(r.employee_id) !== employeeId) return false;

      const explicitCycle = dateKeyFromValue_(r.cycle_start);
      if (explicitCycle) return explicitCycle === currentCycleStartKey;

      // Compatibility for rows created before cycle_start existed.
      if (r.created_at) {
        const legacyCycle = getReportingCycle_(
          r.created_at instanceof Date ? r.created_at : new Date(r.created_at)
        );
        return dateKeyFromValue_(legacyCycle.start) === currentCycleStartKey;
      }

      return false;
    })
    .map(r => ({
      id: String(r.kpi_id || ''),
      activity: String(r.activity || ''),
      value: r.value === '' || r.value == null ? 0 : Number(r.value),
      status: String(r.status || 'Uncomplete') === 'Complete' ? 'Complete' : 'Uncomplete',
      cycleStart: currentCycleStartKey,
      createdAt: r.created_at instanceof Date
        ? r.created_at.toISOString()
        : String(r.created_at || '')
    }))
    .sort((a, b) => new Date(a.createdAt || 0) - new Date(b.createdAt || 0));

  const totalKpiWeight = kpi.reduce((sum, item) => sum + Number(item.value || 0), 0);
  const completeKpiWeight = kpi
    .filter(item => item.status === 'Complete')
    .reduce((sum, item) => sum + Number(item.value || 0), 0);
  const kpiPercent = totalKpiWeight > 0
    ? Math.round((completeKpiWeight / totalKpiWeight) * 100)
    : 0;

  const allKpiComplete = kpi.length > 0 && kpi.every(item => item.status === 'Complete');

  const kpiProofRows = sheetObjects_(ss.getSheetByName('KPIProofs'))
    .filter(r =>
      String(r.employee_id) === employeeId &&
      dateKeyFromValue_(r.cycle_start) === currentCycleStartKey
    )
    .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));

  const latestKpiProof = kpiProofRows[0]
    ? {
        id: String(kpiProofRows[0].proof_id || ''),
        fileName: String(kpiProofRows[0].file_name || ''),
        proofUrl: String(kpiProofRows[0].proof_url || ''),
        createdAt: kpiProofRows[0].created_at instanceof Date
          ? kpiProofRows[0].created_at.toISOString()
          : String(kpiProofRows[0].created_at || '')
      }
    : null;

  return {
    tasks,
    teamMembers,
    kpi,
    kpiSummary: {
      totalWeight: totalKpiWeight,
      completeWeight: completeKpiWeight,
      percent: kpiPercent,
      completeCount: kpi.filter(item => item.status === 'Complete').length,
      totalCount: kpi.length,
      allComplete: allKpiComplete,
      cycleStart: currentCycleStartKey,
      cycleLabel: kpiCycle.label,
      proof: latestKpiProof
    },
    summary: {
      total: tasks.length,
      todo: tasks.filter(t => String(t.status).toLowerCase() === 'to do').length,
      progress: tasks.filter(t => String(t.status).toLowerCase() === 'in progress').length,
      review: tasks.filter(t => String(t.status).toLowerCase() === 'review').length,
      done: tasks.filter(t => String(t.status).toLowerCase() === 'done').length,
      projects: tasks.filter(t => t.taskKind === 'Project').length,
      routines: tasks.filter(t => t.taskKind === 'Routine').length,
      assignedToMe: tasks.filter(t => t.assignedToMe).length
    }
  };
}

function saveTaskManagerTask(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload) throw new Error('Data task tidak ditemukan.');

  const title = String(payload.title || '').trim();
  const taskKind = String(payload.taskKind || 'Project').trim();
  const source = String(payload.source || 'Personal').trim();
  const priority = String(payload.priority || 'Medium').trim();
  const status = String(payload.status || 'To Do').trim();
  const deadline = String(payload.deadline || '').trim();
  const routineDays = Array.isArray(payload.routineDays)
    ? payload.routineDays.map(x => String(x).trim()).filter(Boolean)
    : [];
  const assigneeIds = Array.isArray(payload.assigneeIds)
    ? payload.assigneeIds.map(x => String(x).trim()).filter(Boolean)
    : [];

  if (!title) throw new Error('Nama task wajib diisi.');
  if (!['Project', 'Routine'].includes(taskKind)) throw new Error('Jenis task tidak valid.');
  if (!['Personal', 'Team'].includes(source)) throw new Error('Task Type tidak valid.');
  if (!['To Do', 'In Progress', 'Review', 'Done'].includes(status)) throw new Error('Status task tidak valid.');

  if (source === 'Team') {
    if (!assigneeIds.length) throw new Error('Pilih minimal 1 anggota tim untuk menerima task.');

    const validEmployeeIds = new Set(
      sheetObjects_(ss.getSheetByName('Employees'))
        .map(r => String(r.employee_id || '').trim())
        .filter(Boolean)
    );

    const invalidAssignee = assigneeIds.some(id => !validEmployeeIds.has(id));
    if (invalidAssignee) throw new Error('Terdapat anggota tim yang tidak valid.');
  }

  let deadlineValue = '';
  let taskDateValue = '';

  if (taskKind === 'Project') {
    if (!['Low', 'Medium', 'High'].includes(priority)) throw new Error('Priority task tidak valid.');
    if (!deadline) throw new Error('Deadline project wajib diisi.');

    deadlineValue = new Date(deadline);
    if (isNaN(deadlineValue.getTime())) throw new Error('Deadline project tidak valid.');

    taskDateValue = new Date(
      deadlineValue.getFullYear(),
      deadlineValue.getMonth(),
      deadlineValue.getDate()
    );
  } else {
    const allowedDays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    const invalid = routineDays.some(day => !allowedDays.includes(day));
    if (invalid || !routineDays.length) throw new Error('Pilih minimal 1 hari untuk Routine Task.');
  }

  const sheet = ss.getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const rowValues = {};
  rowValues.employee_id = employeeId;
  rowValues.source = source;
  rowValues.title = title;
  rowValues.task_date = taskDateValue;
  rowValues.start_time = '';
  rowValues.end_time = '';
  rowValues.priority = taskKind === 'Project' ? priority : '';
  rowValues.status = status;
  rowValues.task_kind = taskKind;
  rowValues.deadline = taskKind === 'Project' ? deadlineValue : '';
  rowValues.routine_days = taskKind === 'Routine' ? routineDays.join(',') : '';
  rowValues.task_assignees = source === 'Team' ? assigneeIds.join(',') : '';

  if (payload.id) {
    let rowNumber = -1;

    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][idx.task_id]) === String(payload.id) &&
        String(data[i][idx.employee_id]) === employeeId
      ) {
        rowNumber = i + 1;
        break;
      }
    }

    if (rowNumber < 0) throw new Error('Task tidak ditemukan.');

    Object.keys(rowValues).forEach(key => {
      if (idx[key] != null) sheet.getRange(rowNumber, idx[key] + 1).setValue(rowValues[key]);
    });

    return getTaskManagerPageData(sessionToken);
  }

  const newRow = headers.map(header => {
    if (header === 'task_id') return 'TASK-' + Utilities.getUuid();
    return Object.prototype.hasOwnProperty.call(rowValues, header) ? rowValues[header] : '';
  });

  sheet.appendRow(newRow);
  return getTaskManagerPageData(sessionToken);
}


function saveTaskKpi(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload) throw new Error('Data KPI tidak ditemukan.');

  const activity = String(payload.activity || '').trim();
  const valueRaw = payload.value;
  const status = String(payload.status || 'Uncomplete').trim();

  if (!activity) throw new Error('Activity wajib diisi.');
  if (valueRaw === '' || valueRaw == null || isNaN(Number(valueRaw))) {
    throw new Error('Value KPI wajib berupa angka.');
  }

  const value = Number(valueRaw);
  if (!isFinite(value) || value <= 0) {
    throw new Error('Value KPI harus lebih dari 0.');
  }

  if (!['Complete', 'Uncomplete'].includes(status)) {
    throw new Error('Status KPI tidak valid.');
  }

  const cycle = getReportingCycle_(new Date());
  const cycleStartValue = new Date(cycle.start);

  const sheet = ss.getSheetByName('TaskKPI');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  if (payload.id) {
    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][idx.kpi_id]) === String(payload.id) &&
        String(data[i][idx.employee_id]) === employeeId
      ) {
        sheet.getRange(i + 1, idx.activity + 1).setValue(activity);
        sheet.getRange(i + 1, idx.value + 1).setValue(value);
        sheet.getRange(i + 1, idx.status + 1).setValue(status);
        if (idx.cycle_start != null) {
          sheet.getRange(i + 1, idx.cycle_start + 1).setValue(cycleStartValue);
        }
        return getTaskManagerPageData(sessionToken);
      }
    }

    throw new Error('KPI tidak ditemukan.');
  }

  const row = headers.map(header => {
    if (header === 'kpi_id') return 'KPI-' + Utilities.getUuid();
    if (header === 'employee_id') return employeeId;
    if (header === 'activity') return activity;
    if (header === 'value') return value;
    if (header === 'status') return status;
    if (header === 'cycle_start') return cycleStartValue;
    if (header === 'created_at') return new Date();
    return '';
  });

  sheet.appendRow(row);
  return getTaskManagerPageData(sessionToken);
}

function uploadKpiProof(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload || !payload.fileDataUrl) {
    throw new Error('Bukti KPI belum dipilih.');
  }

  const pageData = getTaskManagerPageData(sessionToken);
  if (!pageData.kpiSummary?.allComplete) {
    throw new Error('Upload bukti KPI hanya dapat dilakukan setelah seluruh status KPI Complete.');
  }

  const fileName = String(payload.fileName || 'Bukti KPI').trim() || 'Bukti KPI';
  const cycleStart = pageData.kpiSummary.cycleStart;
  const cycleParts = cycleStart.split('-').map(Number);
  const cycleStartValue = new Date(cycleParts[0], cycleParts[1] - 1, cycleParts[2]);

  const proofUrl = saveGenericDataUrlToDrive_(
    payload.fileDataUrl,
    employeeId + '_KPI_' + fileName
  );

  ss.getSheetByName('KPIProofs').appendRow([
    'KPI-PROOF-' + Utilities.getUuid(),
    employeeId,
    cycleStartValue,
    fileName,
    proofUrl,
    new Date()
  ]);

  return getTaskManagerPageData(sessionToken);
}

function updateTaskKpiStatus(sessionToken, kpiId, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const id = String(kpiId || '').trim();
  const newStatus = String(status || '').trim();

  if (!id) throw new Error('KPI ID tidak valid.');
  if (!['Complete', 'Uncomplete'].includes(newStatus)) {
    throw new Error('Status KPI tidak valid.');
  }

  const sheet = ss.getSheetByName('TaskKPI');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('KPI tidak ditemukan.');

  const headers = data[0].map(String);
  const idxId = headers.indexOf('kpi_id');
  const idxEmployee = headers.indexOf('employee_id');
  const idxStatus = headers.indexOf('status');

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][idxId]) === id &&
      String(data[i][idxEmployee]) === employeeId
    ) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(newStatus);
      return getTaskManagerPageData(sessionToken);
    }
  }

  throw new Error('KPI tidak ditemukan.');
}

function deleteTaskKpi(sessionToken, kpiId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const id = String(kpiId || '').trim();
  if (!id) throw new Error('KPI ID tidak valid.');

  const sheet = ss.getSheetByName('TaskKPI');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('KPI tidak ditemukan.');

  const headers = data[0].map(String);
  const idxKpiId = headers.indexOf('kpi_id');
  const idxEmployeeId = headers.indexOf('employee_id');

  for (let i = data.length - 1; i >= 1; i--) {
    if (
      String(data[i][idxKpiId]) === id &&
      String(data[i][idxEmployeeId]) === employeeId
    ) {
      sheet.deleteRow(i + 1);
      return getTaskManagerPageData(sessionToken);
    }
  }

  throw new Error('KPI tidak ditemukan.');
}

function updateTaskManagerStatus(sessionToken, taskId, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const id = String(taskId || '').trim();
  const newStatus = String(status || '').trim();

  if (!id) throw new Error('Task ID tidak valid.');
  if (!['To Do', 'In Progress', 'Review', 'Done'].includes(newStatus)) {
    throw new Error('Status task tidak valid.');
  }

  const sheet = ss.getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Task tidak ditemukan.');

  const headers = data[0].map(String);
  const idxTaskId = headers.indexOf('task_id');
  const idxEmployeeId = headers.indexOf('employee_id');
  const idxAssignees = headers.indexOf('task_assignees');
  const idxStatus = headers.indexOf('status');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxTaskId]) !== id) continue;

    const ownerId = String(data[i][idxEmployeeId] || '');
    const assigneeIds = idxAssignees >= 0
      ? String(data[i][idxAssignees] || '').split(',').map(x => x.trim()).filter(Boolean)
      : [];

    const canUpdate = ownerId === employeeId ||
      assigneeIds.includes(employeeId);

    if (!canUpdate) throw new Error('Task ini tidak dibagikan kepada akun kamu.');

    sheet.getRange(i + 1, idxStatus + 1).setValue(newStatus);
    return getTaskManagerPageData(sessionToken);
  }

  throw new Error('Task tidak ditemukan.');
}

function deleteTaskManagerTask(sessionToken, taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const id = String(taskId || '').trim();
  if (!id) throw new Error('Task ID tidak valid.');

  const sheet = ss.getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Task tidak ditemukan.');

  const headers = data[0].map(String);
  const idxTaskId = headers.indexOf('task_id');
  const idxEmployeeId = headers.indexOf('employee_id');

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idxTaskId]) !== id) continue;

    if (String(data[i][idxEmployeeId]) !== employeeId) {
      throw new Error('Task Team yang dibagikan kepada kamu hanya dapat dihapus oleh pembuat task.');
    }

    sheet.deleteRow(i + 1);
    return getTaskManagerPageData(sessionToken);
  }

  throw new Error('Task tidak ditemukan.');
}

function rTaskDate_(dateValue, timeValue) {
  const key = dateKeyFromValue_(dateValue);
  const time = normalizeTime_(timeValue) || '00:00';
  return key + 'T' + time + ':00';
}


function getFinancePageData(sessionToken, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const now = new Date();

  let selectedMonth = String(monthKey || '').trim();
  if (!/^\d{4}-\d{2}$/.test(selectedMonth)) {
    const currentCycle = getReportingCycle_(now);
    selectedMonth = Utilities.formatDate(currentCycle.start, APP.TIMEZONE, 'yyyy-MM');
  }

  const monthParts = selectedMonth.split('-').map(Number);
  const year = monthParts[0];
  const month = monthParts[1] - 1;

  // Finance period runs from the 5th through the 4th of the following month.
  const periodStart = new Date(year, month, 5, 0, 0, 0, 0);
  const periodEnd = new Date(year, month + 1, 5, 0, 0, 0, 0);
  const periodLastDay = new Date(periodEnd.getTime() - 1);
  const periodLabel =
    Utilities.formatDate(periodStart, APP.TIMEZONE, 'dd MMM') +
    ' - ' +
    Utilities.formatDate(periodLastDay, APP.TIMEZONE, 'dd MMM yyyy');

  const attendanceRows = sheetObjects_(ss.getSheetByName('Attendance'))
    .filter(r =>
      String(r.employee_id) === employeeId &&
      r.date &&
      isWithin_(r.date, periodStart, periodEnd)
    );

  const wfoRows = attendanceRows.filter(r =>
    String(r.work_mode || '').toLowerCase() === 'work from office'
  );

  // Payment elements are read-only in Employee App.
  // Administrator will manage label, amount, number of elements and calculation type.
  let elementRows = sheetObjects_(ss.getSheetByName('FinanceElements'))
    .filter(r =>
      String(r.employee_id) === employeeId &&
      String(r.active).toLowerCase() !== 'false'
    )
    .sort((a, b) => Number(a.sort_order || 0) - Number(b.sort_order || 0));

  const paymentElements = elementRows.map(r => {
    const type = String(r.calculation_type || 'reference').trim().toLowerCase();
    const amount = Number(r.amount || 0);
    let calculatedAmount = amount;
    let calculationNote = 'Reference';

    if (type === 'wfo_rate') {
      calculatedAmount = wfoRows.length * amount;
      calculationNote = `${wfoRows.length} WFO × ${amount}`;
    } else if (type === 'fixed') {
      calculatedAmount = amount;
      calculationNote = 'Fixed payment';
    } else {
      calculatedAmount = amount;
      calculationNote = 'Reference only';
    }

    return {
      id: String(r.element_id || ''),
      label: String(r.label || 'Payment Element'),
      type,
      amount,
      calculatedAmount,
      calculationNote,
      includedInEstimate: type === 'fixed' || type === 'wfo_rate'
    };
  });

  const dailyWfoElement = paymentElements.find(x => x.type === 'wfo_rate');
  const dailyRate = Number(dailyWfoElement?.amount || 0);

  const estimatedSalary = paymentElements
    .filter(x => x.includedInEstimate)
    .reduce((sum, x) => sum + Number(x.calculatedAmount || 0), 0);

  const reimburseRows = sheetObjects_(ss.getSheetByName('Reimbursements'))
    .filter(r =>
      String(r.employee_id) === employeeId &&
      r.date &&
      isWithin_(r.date, periodStart, periodEnd)
    )
    .sort((a, b) => new Date(b.created_at || b.date || 0) - new Date(a.created_at || a.date || 0));

  const reimbursements = reimburseRows.map(r => ({
    id: String(r.reimburse_id || ''),
    date: dateKeyFromValue_(r.date),
    category: String(r.category || ''),
    description: String(r.description || ''),
    amount: Number(r.amount || 0),
    proofUrl: String(r.proof_url || ''),
    status: String(r.status || 'Pending'),
    createdAt: r.created_at instanceof Date ? r.created_at.toISOString() : String(r.created_at || '')
  }));

  const sumReimbursementByStatus = status => reimbursements
    .filter(r => String(r.status).toLowerCase() === status.toLowerCase())
    .reduce((sum, r) => sum + Number(r.amount || 0), 0);

  const claimRows = sheetObjects_(ss.getSheetByName('Claims'))
    .filter(r =>
      String(r.employee_id) === employeeId &&
      r.date &&
      isWithin_(r.date, periodStart, periodEnd)
    )
    .sort((a, b) => new Date(b.created_at || b.date || 0) - new Date(a.created_at || a.date || 0));

  const claims = claimRows.map(r => ({
    id: String(r.claim_id || ''),
    claimType: String(r.claim_type || 'Achievement'),
    date: dateKeyFromValue_(r.date),
    name: String(r.claim_name || ''),
    description: String(r.description || ''),
    amount: Number(r.amount || 0),
    proofUrl: String(r.proof_url || ''),
    status: String(r.status || 'Pending'),
    createdAt: r.created_at instanceof Date ? r.created_at.toISOString() : String(r.created_at || '')
  }));

  const sumClaimByStatus = status => claims
    .filter(r => String(r.status).toLowerCase() === status.toLowerCase())
    .reduce((sum, r) => sum + Number(r.amount || 0), 0);

  const attendanceLog = wfoRows
    .sort((a, b) => new Date(b.date || 0) - new Date(a.date || 0))
    .map(r => ({
      date: dateKeyFromValue_(r.date),
      checkIn: r.check_in instanceof Date ? r.check_in.toISOString() : String(r.check_in || ''),
      checkOut: r.check_out instanceof Date ? r.check_out.toISOString() : String(r.check_out || ''),
      workMode: String(r.work_mode || ''),
      earning: dailyRate
    }));

  return {
    month: selectedMonth,
    period: {
      start: periodStart.toISOString(),
      endExclusive: periodEnd.toISOString(),
      label: periodLabel
    },
    payroll: {
      wfoDays: wfoRows.length,
      totalAttendanceDays: attendanceRows.length,
      estimatedSalary,
      dailyRate,
      paymentElements
    },
    reimbursement: {
      pendingAmount: sumReimbursementByStatus('Pending'),
      approvedAmount: sumReimbursementByStatus('Approved'),
      rejectedAmount: sumReimbursementByStatus('Rejected'),
      totalRequests: reimbursements.length,
      rows: reimbursements
    },
    claims: {
      pendingAmount: sumClaimByStatus('Pending'),
      approvedAmount: sumClaimByStatus('Approved'),
      rejectedAmount: sumClaimByStatus('Rejected'),
      totalRequests: claims.length,
      rows: claims
    },
    attendanceLog
  };
}

function saveReimbursement(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload) throw new Error('Data reimbursement tidak ditemukan.');

  const date = String(payload.date || '').trim();
  const category = String(payload.category || '').trim();
  const description = String(payload.description || '').trim();
  const amount = Number(payload.amount || 0);

  if (!date) throw new Error('Tanggal reimbursement wajib diisi.');
  if (!category) throw new Error('Kategori reimbursement wajib dipilih.');
  if (!description) throw new Error('Deskripsi reimbursement wajib diisi.');
  if (!isFinite(amount) || amount <= 0) throw new Error('Nominal reimbursement harus lebih dari 0.');
  if (!payload.proofDataUrl) throw new Error('Bukti transaksi wajib dilampirkan.');

  const dateParts = date.split('-').map(Number);
  const dateValue = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
  const now = new Date();

  const proofUrl = saveDataUrlToDrive_(
    payload.proofDataUrl,
    employeeId + '_reimburse_' + formatFileStamp_(now) + '.jpg'
  );

  ss.getSheetByName('Reimbursements').appendRow([
    'REIM-' + Utilities.getUuid(),
    employeeId,
    dateValue,
    category,
    description,
    amount,
    proofUrl,
    'Pending',
    now
  ]);

  const claimCycle = getReportingCycle_(dateValue);
  return getFinancePageData(sessionToken, Utilities.formatDate(claimCycle.start, APP.TIMEZONE, 'yyyy-MM'));
}


function saveAchievementClaim(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload) throw new Error('Data klaim tidak ditemukan.');

  const date = String(payload.date || '').trim();
  const name = String(payload.name || '').trim();
  const description = String(payload.description || '').trim();
  const amount = Number(payload.amount || 0);

  if (!date) throw new Error('Tanggal klaim wajib diisi.');
  if (!name) throw new Error('Nama klaim wajib diisi.');
  if (!description) throw new Error('Keterangan klaim wajib diisi.');
  if (!isFinite(amount) || amount <= 0) throw new Error('Nominal klaim harus lebih dari 0.');
  if (!payload.proofDataUrl) throw new Error('Bukti klaim wajib dilampirkan.');

  const dateParts = date.split('-').map(Number);
  const dateValue = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
  const now = new Date();

  const proofUrl = saveDataUrlToDrive_(
    payload.proofDataUrl,
    employeeId + '_achievement_claim_' + formatFileStamp_(now) + '.jpg'
  );

  ss.getSheetByName('Claims').appendRow([
    'CLAIM-' + Utilities.getUuid(),
    employeeId,
    'Achievement',
    dateValue,
    name,
    description,
    amount,
    proofUrl,
    'Pending',
    now
  ]);

  const cycle = getReportingCycle_(dateValue);
  return getFinancePageData(sessionToken, Utilities.formatDate(cycle.start, APP.TIMEZONE, 'yyyy-MM'));
}

function deleteAchievementClaim(sessionToken, claimId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const id = String(claimId || '').trim();
  if (!id) throw new Error('Claim ID tidak valid.');

  const sheet = ss.getSheetByName('Claims');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Klaim tidak ditemukan.');

  const headers = data[0].map(String);
  const idxId = headers.indexOf('claim_id');
  const idxEmployee = headers.indexOf('employee_id');
  const idxStatus = headers.indexOf('status');
  const idxDate = headers.indexOf('date');

  for (let i = data.length - 1; i >= 1; i--) {
    if (
      String(data[i][idxId]) === id &&
      String(data[i][idxEmployee]) === employeeId
    ) {
      if (String(data[i][idxStatus] || 'Pending').toLowerCase() !== 'pending') {
        throw new Error('Hanya klaim dengan status Pending yang dapat dihapus.');
      }

      const dateValue = data[i][idxDate] instanceof Date
        ? data[i][idxDate]
        : new Date(data[i][idxDate]);

      sheet.deleteRow(i + 1);

      const cycle = getReportingCycle_(dateValue);
      return getFinancePageData(sessionToken, Utilities.formatDate(cycle.start, APP.TIMEZONE, 'yyyy-MM'));
    }
  }

  throw new Error('Klaim tidak ditemukan.');
}

function deleteReimbursement(sessionToken, reimburseId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const id = String(reimburseId || '').trim();
  if (!id) throw new Error('Reimbursement ID tidak valid.');

  const sheet = ss.getSheetByName('Reimbursements');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Reimbursement tidak ditemukan.');

  const headers = data[0].map(String);
  const idxId = headers.indexOf('reimburse_id');
  const idxEmployee = headers.indexOf('employee_id');
  const idxStatus = headers.indexOf('status');
  const idxDate = headers.indexOf('date');

  for (let i = data.length - 1; i >= 1; i--) {
    if (
      String(data[i][idxId]) === id &&
      String(data[i][idxEmployee]) === employeeId
    ) {
      if (String(data[i][idxStatus]).toLowerCase() !== 'pending') {
        throw new Error('Hanya reimbursement berstatus Pending yang dapat dihapus.');
      }

      const dateValue = data[i][idxDate] instanceof Date
        ? data[i][idxDate]
        : new Date(data[i][idxDate]);

      const cycle = getReportingCycle_(dateValue);
      sheet.deleteRow(i + 1);
      return getFinancePageData(sessionToken, Utilities.formatDate(cycle.start, APP.TIMEZONE, 'yyyy-MM'));
    }
  }

  throw new Error('Reimbursement tidak ditemukan.');
}


function getSettingsPageData(sessionToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  const employee = getEmployee_(ss, employeeId);
  const birthDate = employee.birth_date
    ? dateKeyFromValue_(employee.birth_date)
    : '';

  return {
    employee: {
      id: String(employee.employee_id || employeeId),
      name: String(employee.name || ''),
      division: String(employee.division || ''),
      username: String(employee.username || ''),
      phone: String(employee.phone || ''),
      email: String(employee.email || ''),
      birthDate,
      address: String(employee.address || ''),
      emergencyContactName: String(employee.emergency_contact_name || ''),
      emergencyContactPhone: String(employee.emergency_contact_phone || ''),
      profilePhotoUrl: String(employee.profile_photo_url || '')
    },
    security: {
      hasPassword: !!String(employee.password || '').trim()
    }
  };
}

function savePersonalBiodata(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload) throw new Error('Data biodata tidak ditemukan.');

  const name = String(payload.name || '').trim();
  const phone = String(payload.phone || '').trim();
  const email = String(payload.email || '').trim();
  const birthDate = String(payload.birthDate || '').trim();
  const address = String(payload.address || '').trim();
  const emergencyContactName = String(payload.emergencyContactName || '').trim();
  const emergencyContactPhone = String(payload.emergencyContactPhone || '').trim();

  if (!name) throw new Error('Nama lengkap wajib diisi.');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Format email belum valid.');
  }

  const sheet = ss.getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let rowNumber = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.employee_id]) === employeeId) {
      rowNumber = i + 1;
      break;
    }
  }

  if (rowNumber < 0) throw new Error('Data karyawan tidak ditemukan.');

  let profilePhotoUrl = String(data[rowNumber - 1][idx.profile_photo_url] || '');
  if (payload.profilePhotoDataUrl) {
    profilePhotoUrl = saveDataUrlToDrive_(
      payload.profilePhotoDataUrl,
      employeeId + '_profile_' + formatFileStamp_(new Date()) + '.jpg'
    );
  }

  sheet.getRange(rowNumber, idx.name + 1).setValue(name);
  sheet.getRange(rowNumber, idx.phone + 1).setValue(phone);
  sheet.getRange(rowNumber, idx.email + 1).setValue(email);

  if (birthDate) {
    const parts = birthDate.split('-').map(Number);
    sheet.getRange(rowNumber, idx.birth_date + 1)
      .setValue(new Date(parts[0], parts[1] - 1, parts[2]));
  } else {
    sheet.getRange(rowNumber, idx.birth_date + 1).clearContent();
  }

  sheet.getRange(rowNumber, idx.address + 1).setValue(address);
  sheet.getRange(rowNumber, idx.emergency_contact_name + 1).setValue(emergencyContactName);
  sheet.getRange(rowNumber, idx.emergency_contact_phone + 1).setValue(emergencyContactPhone);
  sheet.getRange(rowNumber, idx.profile_photo_url + 1).setValue(profilePhotoUrl);

  return getSettingsPageData(sessionToken);
}

function updateAccountCredentials(sessionToken, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAppReady_(ss);

  const session = requireEmployeeSession_(sessionToken);
  const employeeId = session.employeeId;

  if (!payload) throw new Error('Data akun tidak ditemukan.');

  const username = String(payload.username || '').trim();
  const currentPassword = String(payload.currentPassword || '');
  const newPassword = String(payload.newPassword || '');
  const confirmPassword = String(payload.confirmPassword || '');

  if (!username) throw new Error('Username wajib diisi.');
  if (!/^[A-Za-z0-9._-]{4,30}$/.test(username)) {
    throw new Error('Username minimal 4 karakter dan hanya boleh berisi huruf, angka, titik, underscore, atau tanda minus.');
  }

  const sheet = ss.getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let rowNumber = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx.employee_id]) === employeeId) {
      rowNumber = i + 1;
      break;
    }
  }

  if (rowNumber < 0) throw new Error('Data karyawan tidak ditemukan.');

  const duplicate = data.slice(1).some((row, index) =>
    index + 2 !== rowNumber &&
    String(row[idx.username] || '').trim().toLowerCase() === username.toLowerCase()
  );
  if (duplicate) throw new Error('Username tersebut sudah digunakan.');

  const storedPassword = String(data[rowNumber - 1][idx.password] || '');
  const hasPassword = !!storedPassword.trim();

  if (hasPassword && newPassword && !verifyPassword_(currentPassword, storedPassword)) {
    throw new Error('Current password tidak sesuai.');
  }

  if (newPassword) {
    if (newPassword.length < 8) throw new Error('Password baru minimal 8 karakter.');
    if (newPassword !== confirmPassword) throw new Error('Konfirmasi password belum sama.');
  } else if (!hasPassword) {
    throw new Error('Buat password pertama untuk mengaktifkan kredensial akun.');
  }

  sheet.getRange(rowNumber, idx.username + 1).setValue(username);

  if (newPassword) {
    sheet.getRange(rowNumber, idx.password + 1).setValue(hashPassword_(newPassword));
  }

  return getSettingsPageData(sessionToken);
}

function hashPassword_(password) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + String(password),
    Utilities.Charset.UTF_8
  );
  const hash = digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
  return 'sha256$' + salt + '$' + hash;
}

function verifyPassword_(password, storedValue) {
  const stored = String(storedValue || '');

  // Backward compatibility jika versi lama pernah menyimpan plain text.
  if (!stored.startsWith('sha256$')) {
    return String(password) === stored;
  }

  const parts = stored.split('$');
  if (parts.length !== 3) return false;

  const salt = parts[1];
  const expected = parts[2];
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + String(password),
    Utilities.Charset.UTF_8
  );
  const actual = digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
  return actual === expected;
}


function getReportingCycle_(referenceDate) {
  const ref = referenceDate ? new Date(referenceDate) : new Date();
  const year = ref.getFullYear();
  const month = ref.getMonth();
  const day = ref.getDate();

  let start;
  let end;

  if (day >= 5) {
    start = new Date(year, month, 5, 0, 0, 0, 0);
    end = new Date(year, month + 1, 5, 0, 0, 0, 0);
  } else {
    start = new Date(year, month - 1, 5, 0, 0, 0, 0);
    end = new Date(year, month, 5, 0, 0, 0, 0);
  }

  const lastDay = new Date(end.getTime() - 1);

  return {
    start,
    end,
    label: Utilities.formatDate(start, APP.TIMEZONE, 'dd MMM') +
      ' - ' +
      Utilities.formatDate(lastDay, APP.TIMEZONE, 'dd MMM yyyy')
  };
}

function getWeeklyKpi_(ss, employeeId) {
  // Kept as a compatibility wrapper for any older calls.
  return getWeeklyKpiFromRows_(
    sheetObjects_(ss.getSheetByName('Attendance')),
    sheetObjects_(ss.getSheetByName('JobReport')),
    sheetObjects_(ss.getSheetByName('Tasks')),
    employeeId
  );
}

function getWeeklyKpiFromRows_(attendanceAll, reportAll, taskAll, employeeId) {
  const now = new Date();
  const cycle = getReportingCycle_(now);

  const attendanceRows = attendanceAll
    .filter(r =>
      String(r.employee_id) === employeeId &&
      isWithin_(r.date, cycle.start, cycle.end)
    );

  let workSeconds = 0;
  const attendanceDays = new Set();

  attendanceRows.forEach(r => {
    const key = dateKeyFromValue_(r.date);
    if (key) attendanceDays.add(key);

    if (r.check_in) {
      const ci = new Date(r.check_in);
      const co = r.check_out ? new Date(r.check_out) : new Date();
      workSeconds += Math.max(0, Math.floor((co.getTime() - ci.getTime()) / 1000));
    }
  });

  const reportRows = reportAll
    .filter(r =>
      String(r.employee_id) === employeeId &&
      isWithin_(r.date, cycle.start, cycle.end)
    );

  const submittedJobs = reportRows.length;
  const documentedJobs = reportRows.filter(r =>
    String(r.documentation_url || '').trim() !== ''
  ).length;

  // Task Completion remains based on the current task database.
  const taskRows = taskAll
    .filter(r => taskVisibleToEmployee_(r, employeeId));

  const completedTasks = taskRows
    .filter(r => String(r.status).toLowerCase() === 'done')
    .length;

  return {
    period: {
      start: cycle.start.toISOString(),
      endExclusive: cycle.end.toISOString(),
      label: cycle.label
    },
    work: {
      achievedSeconds: workSeconds
    },
    reports: {
      completedJobs: submittedJobs,
      documentedJobs
    },
    tasks: {
      completed: completedTasks,
      total: taskRows.length,
      percent: taskRows.length ? percent_(completedTasks, taskRows.length) : 0
    },
    attendance: {
      days: attendanceDays.size
    }
  };
}

function getEmployee_(ss, employeeId) {
  return sheetObjects_(ss.getSheetByName('Employees'))
    .find(r => String(r.employee_id) === employeeId) || {};
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#022136')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return sheet;
  }

  // Tambahkan kolom baru dengan aman jika versi aplikasi membutuhkan field tambahan.
  const existingHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn()))
    .getValues()[0]
    .map(h => String(h).trim());

  const missingHeaders = headers.filter(h => !existingHeaders.includes(h));
  if (missingHeaders.length) {
    const startCol = existingHeaders.length + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
    sheet.getRange(1, startCol, 1, missingHeaders.length)
      .setFontWeight('bold')
      .setBackground('#022136')
      .setFontColor('#ffffff');
  }

  sheet.setFrozenRows(1);
  return sheet;
}

function sheetObjects_(sheet) {
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());

  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell != null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function dateKeyFromValue_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value).slice(0, 10);
  return formatDateKey_(d);
}

function formatDateKey_(date) {
  return Utilities.formatDate(date, APP.TIMEZONE, 'yyyy-MM-dd');
}

function isWithin_(value, start, end) {
  if (!value) return false;
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return false;
  return d >= start && d < end;
}

function normalizeTime_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, APP.TIMEZONE, 'HH:mm');
  }
  return String(value);
}

function percent_(value, target) {
  if (!target) return 0;
  return Math.max(0, Math.min(100, Math.round((value / target) * 100)));
}
