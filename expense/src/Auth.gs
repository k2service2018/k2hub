var HASH_ITERATIONS = 5e3;

function hashPassword_(password, salt) {
  var bytes = Utilities.newBlob(salt + '|' + password).getBytes();
  for (var i = 0; i < HASH_ITERATIONS; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return bytesToHex_(bytes);
}

function bytesToHex_(bytes) {
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] < 0 ? bytes[i] + 256 : bytes[i]).toString(16);
    hex += (b.length === 1 ? '0' : '') + b;
  }
  return hex;
}

function constantTimeEquals_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function passwordProblem_(pw) {
  if (!pw || String(pw).length < CFG.MIN_PASSWORD_LEN) {
    return 'รหัสผ่านต้องยาวอย่างน้อย ' + CFG.MIN_PASSWORD_LEN + ' ตัวอักษร';
  }
  return null;
}

function validEmail_(email) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(email || '').trim());
}

function usersSheet_() {
  return sheet_(CFG.USERS_SHEET, USERS_SCHEMA);
}

function readUsers_() {
  var sh = usersSheet_();
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function(h) {
    return String(h).trim();
  });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {
      _row: r + 1
    };
    var blank = true;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = values[r][c];
      if (values[r][c] !== '' && values[r][c] !== null) blank = false;
    }
    if (!blank) rows.push(obj);
  }
  return rows;
}

function userHeaders_() {
  var sh = usersSheet_();
  if (sh.getLastColumn() < 1 || sh.getLastRow() < 1) return USERS_SCHEMA.slice();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) {
    return String(h).trim();
  });
  return headers.filter(String).length ? headers : USERS_SCHEMA.slice();
}

function appendUser_(obj) {
  var sh = usersSheet_();
  var headers = userHeaders_();
  sh.appendRow(headers.map(function(h) {
    return h in obj ? obj[h] : '';
  }));
  return sh.getLastRow();
}

function updateUserRow_(rowNumber, obj) {
  var sh = usersSheet_();
  var headers = userHeaders_();
  var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    if (headers[c] && headers[c] in obj) current[c] = obj[headers[c]];
  }
  sh.getRange(rowNumber, 1, 1, headers.length).setValues([ current ]);
}

function findUserByEmail_(email) {
  if (!email) return null;
  var target = normEmail_(email);
  var rows = readUsers_();
  for (var i = 0; i < rows.length; i++) {
    if (normEmail_(rows[i].Email) === target) return rows[i];
  }
  return null;
}

function normEmail_(v) {
  return String(v == null ? '' : v).trim().toLowerCase();
}

function isActive_(user) {
  if (!user) return false;
  var v = user.Active;
  return !(v === false || String(v).toLowerCase() === 'false' || v === '');
}

function nowIso_() {
  return Utilities.formatDate(new Date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function cellStr_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return String(v);
}

function createSession_(user) {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var payload = {
    email: normEmail_(user.Email),
    name: user.FullName || user.Email,
    role: user.Role || ROLES.SALE,
    exp: Date.now() + CFG.SESSION.IDLE_SEC * 1e3,
    touched: Date.now()
  };
  PropertiesService.getScriptProperties().setProperty(CFG.SESSION.PROP_PREFIX + token, JSON.stringify(payload));
  return token;
}

function getSession_(token) {
  if (!token) return null;
  var props = PropertiesService.getScriptProperties();
  var key = CFG.SESSION.PROP_PREFIX + token;
  var raw = props.getProperty(key);
  if (!raw) return null;
  var payload;
  try {
    payload = JSON.parse(raw);
  } catch (e) {
    return null;
  }
  var now = Date.now();
  if (!payload.exp || payload.exp < now) {
    props.deleteProperty(key);
    return null;
  }
  if (now - Number(payload.touched || 0) > CFG.SESSION.REFRESH_AFTER_SEC * 1e3) {
    payload.exp = now + CFG.SESSION.IDLE_SEC * 1e3;
    payload.touched = now;
    props.setProperty(key, JSON.stringify(payload));
  }
  return payload;
}

function logoutAllSessionsFor_(email) {
  var target = normEmail_(email);
  if (!target) return 0;
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var killed = 0;
  Object.keys(all).forEach(function(k) {
    if (k.indexOf(CFG.SESSION.PROP_PREFIX) !== 0) return;
    var p;
    try {
      p = JSON.parse(all[k]);
    } catch (e) {
      return;
    }
    if (p && normEmail_(p.email) === target) {
      props.deleteProperty(k);
      killed++;
    }
  });
  return killed;
}

var CURRENT_USER = null;

function auth_(token) {
  var s = getSession_(token);
  if (!s) throw new Error('AUTH_REQUIRED');
  CURRENT_USER = identityFor_(s);
  return CURRENT_USER;
}

function me_() {
  if (!CURRENT_USER) throw new Error('AUTH_REQUIRED');
  return CURRENT_USER;
}

function identityFor_(session) {
  var user = findUserByEmail_(session.email);
  var role = user ? user.Role || ROLES.SALE : session.role || ROLES.SALE;
  var superAdmin = role === ROLES.SUPER_ADMIN;
  var active = isActive_(user);
  var checkerDuty = active && role === ROLES.ADMIN || isAssignedIn_('CheckerEmail', session.email);
  var approverDuty = active && role === ROLES.MANAGER || isAssignedIn_('ApproverEmail', session.email);
  var financeDuty = role === ROLES.FINANCE || CFG.FINANCE_EMAILS.map(normEmail_).indexOf(normEmail_(session.email)) >= 0;
  return {
    email: normEmail_(session.email),
    name: user && user.FullName || session.name || session.email,
    role: role,
    active: active,
    registered: !!user,
    profile: user ? profileOf_(user) : null,
    isSuperAdmin: superAdmin,
    isFinance: superAdmin || financeDuty,
    isChecker: superAdmin || checkerDuty,
    isApprover: superAdmin || approverDuty,
    checkerDuty: checkerDuty,
    approverDuty: approverDuty,
    financeDuty: financeDuty,
    canClaim: superAdmin || CFG.ROLES_NO_CLAIM.indexOf(role) < 0
  };
}

function profileOf_(user) {
  var p = {
    email: normEmail_(user.Email),
    role: user.Role || ROLES.SALE,
    active: isActive_(user)
  };
  Object.keys(PROFILE_COLUMN_MAP).forEach(function(key) {
    p[key] = cellStr_(user[PROFILE_COLUMN_MAP[key]]);
  });
  p.createdAt = cellStr_(user.CreatedAt);
  p.lastLogin = cellStr_(user.LastLogin);
  return p;
}

function isAssignedIn_(column, email) {
  var target = normEmail_(email);
  if (!target) return false;
  var rows = readUsers_();
  for (var i = 0; i < rows.length; i++) {
    if (normEmail_(rows[i][column]) === target) return true;
  }
  return false;
}

function roleEmails_(role) {
  return readUsers_().filter(function(u) {
    return (u.Role || ROLES.SALE) === role && isActive_(u);
  }).map(function(u) {
    return normEmail_(u.Email);
  }).filter(function(v, i, a) {
    return v && a.indexOf(v) === i;
  });
}

function superAdminEmails_() {
  return roleEmails_(ROLES.SUPER_ADMIN).concat([ normEmail_(secret_('SUPER_ADMIN_EMAIL')) ]).filter(function(v, i, a) {
    return v && a.indexOf(v) === i;
  });
}

function login(email, password) {
  email = String(email || '').trim();
  var user = findUserByEmail_(email);
  var fail = {
    ok: false,
    error: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง / Incorrect email or password'
  };
  if (!user) {
    hashPassword_(password || '', 'dummy');
    return fail;
  }
  if (!isActive_(user)) {
    return {
      ok: false,
      error: 'บัญชีนี้ยังไม่ถูกเปิดใช้งาน — กรุณารอผู้ดูแลระบบอนุมัติ'
    };
  }
  var computed = hashPassword_(password || '', user.Salt);
  if (!constantTimeEquals_(computed, String(user.PasswordHash))) return fail;
  updateUserRow_(user._row, {
    LastLogin: nowIso_()
  });
  var token = createSession_(user);
  return {
    ok: true,
    token: token,
    user: {
      email: normEmail_(user.Email),
      name: user.FullName || user.Email,
      role: user.Role || ROLES.SALE
    }
  };
}

function logout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty(CFG.SESSION.PROP_PREFIX + token);
  return {
    ok: true
  };
}

function selfRegister(payload) {
  payload = payload || {};
  var email = normEmail_(payload.email);
  if (!validEmail_(email)) return {
    ok: false,
    errors: [ 'รูปแบบอีเมลไม่ถูกต้อง' ]
  };
  var errors = [];
  var pwProblem = passwordProblem_(payload.password);
  if (pwProblem) errors.push(pwProblem);
  if (payload.password !== payload.confirmPassword) errors.push('รหัสผ่านทั้งสองช่องไม่ตรงกัน');
  REGISTER_FIELDS.forEach(function(f) {
    if (f.required && String(payload[f.key] || '').trim() === '') {
      errors.push('กรุณากรอก "' + f.th + ' / ' + f.en + '"');
    }
  });
  var approver = normEmail_(payload.approverEmail);
  var checker = normEmail_(payload.checkerEmail);
  if (approver && !validEmail_(approver)) errors.push('รูปแบบอีเมล Manager ผู้อนุมัติไม่ถูกต้อง');
  if (checker && !validEmail_(checker)) errors.push('รูปแบบอีเมล Admin ผู้ตรวจสอบไม่ถูกต้อง');
  if (approver && approver === email) errors.push('อีเมล Manager ผู้อนุมัติต้องไม่ใช่อีเมลของคุณเอง');
  if (checker && checker === email) errors.push('อีเมล Admin ผู้ตรวจสอบต้องไม่ใช่อีเมลของคุณเอง');
  if (errors.length) return {
    ok: false,
    errors: errors
  };
  var lock = LockService.getScriptLock();
  lock.waitLock(15e3);
  try {
    if (findUserByEmail_(email)) return {
      ok: false,
      errors: [ 'อีเมลนี้มีบัญชีอยู่แล้ว กรุณาเข้าสู่ระบบ' ]
    };
    var salt = Utilities.getUuid();
    var rec = {
      Email: email,
      Role: ROLES.SALE,
      PasswordHash: hashPassword_(payload.password, salt),
      Salt: salt,
      Active: !!CFG.AUTO_APPROVE_REGISTRATION,
      CreatedAt: nowIso_(),
      LastLogin: '',
      UpdatedAt: nowIso_(),
      ApprovedBy: CFG.AUTO_APPROVE_REGISTRATION ? 'auto' : ''
    };
    Object.keys(PROFILE_COLUMN_MAP).forEach(function(key) {
      rec[PROFILE_COLUMN_MAP[key]] = String(payload[key] || '').trim();
    });
    rec.CheckerEmail = checker;
    rec.ApproverEmail = approver;
    appendUser_(rec);
    syncMasterFromUsers_();
    if (!CFG.AUTO_APPROVE_REGISTRATION) {
      notifyAdminsOfRegistration_(rec);
      return {
        ok: true,
        active: false,
        message: 'สมัครเรียบร้อย — รอผู้ดูแลระบบเปิดใช้งานบัญชี'
      };
    }
    var signedIn = login(email, payload.password);
    signedIn.active = true;
    signedIn.message = 'สมัครเรียบร้อย ยินดีต้อนรับ';
    return signedIn;
  } finally {
    lock.releaseLock();
  }
}

function getCurrentUser(token) {
  var s = getSession_(token);
  if (!s) return {
    ok: false
  };
  return {
    ok: true,
    identity: identityFor_(s)
  };
}

function changeMyPassword(token, oldPassword, newPassword) {
  var me = auth_(token);
  var problem = passwordProblem_(newPassword);
  if (problem) return {
    ok: false,
    errors: [ problem ]
  };
  var user = findUserByEmail_(me.email);
  if (!user) return {
    ok: false,
    errors: [ 'ไม่พบผู้ใช้' ]
  };
  var computed = hashPassword_(oldPassword || '', user.Salt);
  if (!constantTimeEquals_(computed, String(user.PasswordHash))) {
    return {
      ok: false,
      errors: [ 'รหัสผ่านเดิมไม่ถูกต้อง' ]
    };
  }
  setUserPassword_(user._row, newPassword);
  return {
    ok: true,
    signedOut: true,
    message: 'เปลี่ยนรหัสผ่านแล้ว — กรุณาเข้าสู่ระบบใหม่'
  };
}

function setUserPassword_(row, newPassword) {
  var salt = Utilities.getUuid();
  updateUserRow_(row, {
    PasswordHash: hashPassword_(newPassword, salt),
    Salt: salt,
    UpdatedAt: nowIso_()
  });
  var rows = readUsers_();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]._row === row) {
      logoutAllSessionsFor_(rows[i].Email);
      break;
    }
  }
}

function requestPasswordReset(email) {
  var user = findUserByEmail_(email);
  if (user && isActive_(user)) {
    var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    PropertiesService.getScriptProperties().setProperty(CFG.SESSION.RESET_PREFIX + token, JSON.stringify({
      email: normEmail_(user.Email),
      exp: Date.now() + CFG.SESSION.RESET_TTL_SEC * 1e3
    }));
    var url = webAppUrl_() + '?reset=' + token;
    sendMail_(user.Email, 'ตั้งรหัสผ่านใหม่ — ' + CFG.APP_TITLE, '<p>สวัสดีครับ</p>' + '<p>มีการขอตั้งรหัสผ่านใหม่สำหรับระบบ <b>' + escapeHtml_(CFG.APP_TITLE) + '</b></p>' + '<p>คลิกลิงก์ด้านล่างเพื่อตั้งรหัสผ่านใหม่ (ลิงก์หมดอายุใน 30 นาที)</p>' + actionButtonHtml_(url, 'ตั้งรหัสผ่านใหม่') + '<p style="color:#666;font-size:12px">หากคุณไม่ได้เป็นผู้ร้องขอ กรุณาเพิกเฉยต่ออีเมลฉบับนี้</p>');
  }
  return {
    ok: true,
    message: 'ถ้าอีเมลนี้มีบัญชีอยู่ ระบบได้ส่งลิงก์ตั้งรหัสผ่านใหม่ไปให้แล้ว'
  };
}

function getReset_(token) {
  if (!token) return null;
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(CFG.SESSION.RESET_PREFIX + token);
  if (!raw) return null;
  var payload;
  try {
    payload = JSON.parse(raw);
  } catch (e) {
    return null;
  }
  if (!payload.exp || payload.exp < Date.now()) {
    props.deleteProperty(CFG.SESSION.RESET_PREFIX + token);
    return null;
  }
  return payload;
}

function validateResetToken(token) {
  var payload = getReset_(token);
  return payload ? {
    ok: true,
    email: payload.email
  } : {
    ok: false,
    errors: [ 'ลิงก์ไม่ถูกต้องหรือหมดอายุแล้ว' ]
  };
}

function resetPasswordWithToken(token, newPassword) {
  var payload = getReset_(token);
  if (!payload) return {
    ok: false,
    errors: [ 'ลิงก์ไม่ถูกต้องหรือหมดอายุแล้ว' ]
  };
  var problem = passwordProblem_(newPassword);
  if (problem) return {
    ok: false,
    errors: [ problem ]
  };
  var user = findUserByEmail_(payload.email);
  if (!user) return {
    ok: false,
    errors: [ 'ไม่พบผู้ใช้' ]
  };
  setUserPassword_(user._row, newPassword);
  PropertiesService.getScriptProperties().deleteProperty(CFG.SESSION.RESET_PREFIX + token);
  return {
    ok: true,
    message: 'ตั้งรหัสผ่านใหม่เรียบร้อย กรุณาเข้าสู่ระบบ'
  };
}

function seedSuperAdmin_() {
  var email = normEmail_(secret_('SUPER_ADMIN_EMAIL'));
  if (!email || findUserByEmail_(email)) return false;
  var salt = Utilities.getUuid();
  appendUser_({
    Email: email,
    FullName: 'Super Admin',
    Role: ROLES.SUPER_ADMIN,
    PasswordHash: hashPassword_(secret_('SEED_PASSWORD'), salt),
    Salt: salt,
    Active: true,
    CreatedAt: nowIso_(),
    LastLogin: '',
    UpdatedAt: nowIso_(),
    ApprovedBy: 'seed'
  });
  return true;
}