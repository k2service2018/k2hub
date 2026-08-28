function updateMyProfile(token, payload) {
  var me = auth_(token);
  payload = payload || {};
  var errors = [];
  REGISTER_FIELDS.forEach(function(f) {
    if (f.required && String(payload[f.key] || '').trim() === '') {
      errors.push('กรุณากรอก "' + f.th + ' / ' + f.en + '"');
    }
  });
  var approver = normEmail_(payload.approverEmail);
  var checker = normEmail_(payload.checkerEmail);
  if (approver && !validEmail_(approver)) errors.push('รูปแบบอีเมล Manager ผู้อนุมัติไม่ถูกต้อง');
  if (checker && !validEmail_(checker)) errors.push('รูปแบบอีเมล Admin ผู้ตรวจสอบไม่ถูกต้อง');
  if (approver && approver === me.email) errors.push('อีเมล Manager ผู้อนุมัติต้องไม่ใช่อีเมลของคุณเอง');
  if (checker && checker === me.email) errors.push('อีเมล Admin ผู้ตรวจสอบต้องไม่ใช่อีเมลของคุณเอง');
  if (errors.length) return {
    ok: false,
    errors: errors
  };
  var lock = LockService.getScriptLock();
  lock.waitLock(15e3);
  try {
    var user = findUserByEmail_(me.email);
    if (!user) return {
      ok: false,
      errors: [ 'ไม่พบบัญชีของคุณ' ]
    };
    var upd = {
      UpdatedAt: nowIso_()
    };
    Object.keys(PROFILE_COLUMN_MAP).forEach(function(key) {
      upd[PROFILE_COLUMN_MAP[key]] = String(payload[key] || '').trim();
    });
    upd.CheckerEmail = checker;
    upd.ApproverEmail = approver;
    updateUserRow_(user._row, upd);
    syncMasterFromUsers_();
    return {
      ok: true,
      message: 'อัปเดตข้อมูลเรียบร้อย',
      identity: identityFor_({
        email: me.email,
        name: payload.name,
        role: me.role
      })
    };
  } finally {
    lock.releaseLock();
  }
}

function syncMasterFromUsers_() {
  var rows = readUsers_().filter(isActive_);
  if (!rows.length) return;
  var master = sheet_(CFG.MASTER_SHEET, MASTER_COLUMNS);
  var headers = master.getRange(1, 1, 1, master.getLastColumn()).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  function mergeColumn(name, values) {
    var c = headers.indexOf(name);
    if (c < 0) return;
    var existing = master.getLastRow() > 1 ? master.getRange(2, c + 1, master.getLastRow() - 1, 1).getValues().map(function(r) {
      return String(r[0] || '').trim();
    }) : [];
    var merged = (DEFAULT_MASTER[name] || []).concat(existing).concat(values).filter(function(v, i, arr) {
      return v !== '' && arr.indexOf(v) === i;
    });
    if (master.getLastRow() > 1) master.getRange(2, c + 1, master.getLastRow() - 1, 1).clearContent();
    if (merged.length) {
      master.getRange(2, c + 1, merged.length, 1).setValues(merged.map(function(v) {
        return [ v ];
      }));
    }
  }
  mergeColumn('employees', rows.map(function(u) {
    return cellStr_(u.FullName);
  }));
  mergeColumn('divisions', rows.map(function(u) {
    return cellStr_(u.Division);
  }));
  mergeColumn('jobTitles', rows.map(function(u) {
    return cellStr_(u.JobTitle);
  }));
}

function listUsers(token, filter) {
  var me = auth_(token);
  if (!me.isSuperAdmin) return {
    ok: false,
    errors: [ 'เฉพาะผู้ดูแลระบบเท่านั้น' ]
  };
  filter = filter || {};
  var rows = readUsers_().map(function(u) {
    return {
      row: u._row,
      email: normEmail_(u.Email),
      name: cellStr_(u.FullName),
      role: cellStr_(u.Role) || ROLES.SALE,
      active: isActive_(u),
      division: cellStr_(u.Division),
      jobTitle: cellStr_(u.JobTitle),
      checkerEmail: cellStr_(u.CheckerEmail),
      approverEmail: cellStr_(u.ApproverEmail),
      phone: cellStr_(u.Phone),
      createdAt: cellStr_(u.CreatedAt),
      lastLogin: cellStr_(u.LastLogin)
    };
  }).filter(function(u) {
    if (!u.email) return false;
    if (filter.status === 'Pending' && u.active) return false;
    if (filter.status === 'Active' && !u.active) return false;
    if (filter.role && u.role !== filter.role) return false;
    if (filter.q) {
      var hay = [ u.email, u.name, u.division, u.jobTitle ].join(' ').toLowerCase();
      if (hay.indexOf(String(filter.q).toLowerCase()) < 0) return false;
    }
    return true;
  });
  rows.sort(function(a, b) {
    if (a.active !== b.active) return a.active ? 1 : -1;
    return String(a.name).localeCompare(String(b.name));
  });
  return safeOut_({
    ok: true,
    rows: rows,
    roles: ROLE_LIST
  });
}

function saveUser(token, email, changes) {
  var me = auth_(token);
  if (!me.isSuperAdmin) return {
    ok: false,
    errors: [ 'เฉพาะผู้ดูแลระบบเท่านั้น' ]
  };
  changes = changes || {};
  var target = normEmail_(email);
  if (changes.role && ROLE_LIST.indexOf(changes.role) < 0) {
    return {
      ok: false,
      errors: [ 'บทบาทไม่ถูกต้อง' ]
    };
  }
  if (target === me.email && changes.role && changes.role !== ROLES.SUPER_ADMIN) {
    return {
      ok: false,
      errors: [ 'ถอดสิทธิ์ผู้ดูแลของตัวเองไม่ได้ — ให้ผู้ดูแลคนอื่นเป็นผู้ดำเนินการ' ]
    };
  }
  if (target === me.email && changes.active === false) {
    return {
      ok: false,
      errors: [ 'ปิดใช้งานบัญชีของตัวเองไม่ได้' ]
    };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(15e3);
  try {
    var user = findUserByEmail_(target);
    if (!user) return {
      ok: false,
      errors: [ 'ไม่พบผู้ใช้ ' + email ]
    };
    var wasInactive = !isActive_(user);
    var upd = {
      UpdatedAt: nowIso_(),
      ApprovedBy: me.email
    };
    [ 'name', 'division', 'jobTitle', 'checkerEmail', 'approverEmail', 'phone', 'employeeCode' ].forEach(function(k) {
      if (changes[k] !== undefined && changes[k] !== null) upd[PROFILE_COLUMN_MAP[k]] = changes[k];
    });
    if (changes.role) upd.Role = changes.role;
    if (changes.active !== undefined) upd.Active = !!changes.active;
    updateUserRow_(user._row, upd);
    syncMasterFromUsers_();
    if (wasInactive && changes.active === true) notifyAccountActivated_(user);
    return {
      ok: true,
      message: 'อัปเดต ' + target + ' แล้ว'
    };
  } finally {
    lock.releaseLock();
  }
}

function adminResetPassword(token, email, newPassword) {
  var me = auth_(token);
  if (!me.isSuperAdmin) return {
    ok: false,
    errors: [ 'เฉพาะผู้ดูแลระบบเท่านั้น' ]
  };
  var problem = passwordProblem_(newPassword);
  if (problem) return {
    ok: false,
    errors: [ problem ]
  };
  var user = findUserByEmail_(email);
  if (!user) return {
    ok: false,
    errors: [ 'ไม่พบผู้ใช้' ]
  };
  setUserPassword_(user._row, newPassword);
  return {
    ok: true,
    message: 'ตั้งรหัสผ่านใหม่ให้ ' + normEmail_(email) + ' แล้ว — ผู้ใช้ถูกออกจากระบบทุกเครื่อง'
  };
}

function webAppUrl_() {
  if (secret_('WEBAPP_URL')) return secret_('WEBAPP_URL');
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function sendMail_(to, subject, htmlBody, attachments) {
  if (!CFG.NOTIFY_BY_EMAIL) return;
  var recipients = (Array.isArray(to) ? to : [ to ]).filter(Boolean).join(',');
  if (!recipients) return;
  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: htmlBody,
      attachments: attachments || [],
      name: CFG.COMPANY || CFG.APP_TITLE
    });
  } catch (err) {
    console.error('ส่งอีเมลไม่สำเร็จ: ' + err);
  }
}

function notifyAdminsOfRegistration_(user) {
  var url = webAppUrl_();
  sendMail_(superAdminEmails_(), '[' + CFG.APP_TITLE + '] มีผู้สมัครใหม่: ' + user.FullName, '<p>มีผู้สมัครใช้งานระบบเบิกค่าใช้จ่าย รอเปิดใช้งานบัญชี</p>' + '<ul>' + '<li><b>ชื่อ:</b> ' + escapeHtml_(user.FullName) + '</li>' + '<li><b>อีเมล:</b> ' + escapeHtml_(user.Email) + '</li>' + '<li><b>แผนก:</b> ' + escapeHtml_(user.Division) + '</li>' + '<li><b>ตำแหน่ง:</b> ' + escapeHtml_(user.JobTitle) + '</li>' + '<li><b>Manager ผู้อนุมัติ:</b> ' + escapeHtml_(user.ApproverEmail) + '</li>' + '</ul>' + actionButtonHtml_(url ? url + '?page=admin' : '', 'เปิดหน้าจัดการผู้ใช้'));
}

function notifyAccountActivated_(user) {
  sendMail_(user.Email, '[' + CFG.APP_TITLE + '] เปิดใช้งานบัญชีแล้ว', '<p>สวัสดีคุณ ' + escapeHtml_(user.FullName) + '</p>' + '<p>บัญชีของคุณถูกเปิดใช้งานแล้ว เข้าสู่ระบบเพื่อเริ่มกรอกใบเบิกได้ทันที</p>' + actionButtonHtml_(webAppUrl_(), 'เข้าสู่ระบบ'));
}

function actionButtonHtml_(link, text) {
  if (!link) return '';
  return '<p style="margin-top:18px"><a href="' + link + '" ' + 'style="background:#1f3864;color:#fff;padding:10px 22px;border-radius:6px;' + 'text-decoration:none;font-family:sans-serif;font-size:14px">' + text + '</a></p>';
}

function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}