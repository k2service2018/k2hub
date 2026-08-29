function doGet(e) {
  var p = e && e.parameter || {};
  var t = HtmlService.createTemplateFromFile('Index');
  t.initialPage = safeParam_(p.page);
  t.initialClaimId = safeParam_(p.id);
  t.resetToken = safeParam_(p.reset);
  return t.evaluate().setTitle(CFG.APP_TITLE + ' — ' + CFG.DOC_NO).addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function safeParam_(v) {
  return String(v == null ? '' : v).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 128);
}

function ss_() {
  if (secret_('SPREADSHEET_ID')) return SpreadsheetApp.openById(secret_('SPREADSHEET_ID'));
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('ไม่พบ Spreadsheet — กรุณาใส่ SPREADSHEET_ID ใน Secrets.gs');
  }
  return active;
}

function sheet_(name, headers) {
  var book = ss_();
  var sh = book.getSheetByName(name);
  if (!sh) {
    sh = book.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([ headers ]).setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function setup() {
  sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
  sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS);
  sheet_(CFG.APPROVAL_SHEET, APPROVAL_COLUMNS);
  sheet_(CFG.USERS_SHEET, USERS_SCHEMA);
  ensureUserColumns_();
  var master = ss_().getSheetByName(CFG.MASTER_SHEET);
  if (!master) {
    master = sheet_(CFG.MASTER_SHEET, MASTER_COLUMNS);
    master.getRange(2, 4, 3, 1).setValues([ [ 'BB' ], [ 'BP' ], [ 'TRC' ] ]);
    master.getRange('A2').setNote('เติมรายชื่อพนักงานในคอลัมน์นี้ (แถวละ 1 ชื่อ) เพื่อให้ฟอร์มแนะนำอัตโนมัติ');
  }
  var seeded = seedSuperAdmin_();
  var lines = [ 'Setup เรียบร้อย: ' + ss_().getUrl() ];
  if (seeded) {
    lines.push('');
    lines.push('สร้างบัญชีผู้ดูแลระบบแล้ว:');
    lines.push('  อีเมล: ' + secret_('SUPER_ADMIN_EMAIL'));
    lines.push('  รหัสผ่าน: ' + secret_('SEED_PASSWORD'));
    lines.push('  ⚠ เข้าสู่ระบบแล้วเปลี่ยนรหัสผ่านทันที');
  }
  var missing = missingSecrets_();
  if (missing.length) {
    lines.push('');
    lines.push('⚠ ยังไม่ได้ตั้งค่าใน Secrets.gs: ' + missing.join(', '));
  }
  return lines.join('\n');
}

function ensureUserColumns_() {
  var sh = usersSheet_();
  var lastCol = sh.getLastColumn();
  var current = lastCol >= 1 && sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h).trim();
  }) : [];
  var missing = USERS_SCHEMA.filter(function(c) {
    return current.indexOf(c) === -1;
  });
  if (!missing.length) return 0;
  sh.getRange(1, current.length + 1, 1, missing.length).setValues([ missing ]).setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
  return missing.length;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Expense Claim').addItem('Setup ชีตข้อมูล', 'setup').addItem('เปิดฟอร์ม (Web App)', 'showWebAppUrl_').addToUi();
}

function showWebAppUrl_() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Web App URL', url || 'ยังไม่ได้ deploy — ไปที่ Deploy > New deployment > Web app', SpreadsheetApp.getUi().ButtonSet.OK);
}

function getBootstrap(token) {
  var session = getSession_(token);
  return safeOut_({
    cfg: {
      appTitle: CFG.APP_TITLE,
      docNo: CFG.DOC_NO,
      docTitle: CFG.DOC_TITLE,
      company: CFG.COMPANY,
      maxRows: CFG.MAX_ROWS,
      currency: CFG.CURRENCY,
      lockAfterSubmit: CFG.LOCK_AFTER_SUBMIT,
      autoApproveRegistration: CFG.AUTO_APPROVE_REGISTRATION,
      requirePaymentStep: CFG.REQUIRE_PAYMENT_STEP,
      minPasswordLen: CFG.MIN_PASSWORD_LEN,
      idleSec: CFG.SESSION.IDLE_SEC
    },
    headerFields: HEADER_FIELDS,
    itemFields: ITEM_FIELDS,
    registerFields: REGISTER_FIELDS,
    groups: GROUPS,
    signBlocks: SIGN_BLOCKS,
    workflow: WORKFLOW,
    statuses: STATUSES,
    statusLabels: STATUS_LABELS,
    stages: STAGES,
    reviewFlags: REVIEW_FLAGS,
    roles: ROLE_LIST,
    master: readMaster_(),
    identity: session ? identityFor_(session) : null
  });
}

function involvedInClaim_(claim, me) {
  var email = normEmail_(me.email);
  if (normEmail_(claim.createdBy) === email) return true;
  if (normEmail_(claim.checkedBy) === email) return true;
  if (normEmail_(claim.approvedBy) === email) return true;
  if (normEmail_(claim.paidBy) === email) return true;
  if (normEmail_(claim.checkerEmail) === email) return true;
  if (normEmail_(claim.approverEmail) === email) return true;
  return queueCheck_(claim, me) || queueApprove_(claim, me) || queuePay_(claim, me);
}

function requireClaimRight_() {
  if (me_().canClaim) return null;
  return {
    ok: false,
    errors: [ 'บทบาท ' + me_().role + ' ไม่ได้ทำใบเบิกในระบบนี้' ]
  };
}

function requireActiveAccount_(token) {
  var session = getSession_(token);
  if (!session) return {
    ok: false,
    authRequired: true,
    errors: [ 'กรุณาเข้าสู่ระบบก่อนใช้งาน' ]
  };
  CURRENT_USER = identityFor_(session);
  if (!CURRENT_USER.registered) {
    return {
      ok: false,
      authRequired: true,
      errors: [ 'ไม่พบบัญชีของคุณในระบบ' ]
    };
  }
  if (!CURRENT_USER.active) {
    return {
      ok: false,
      inactive: true,
      errors: [ 'บัญชีของคุณยังไม่ถูกเปิดใช้งาน — กรุณารอผู้ดูแลระบบ' ]
    };
  }
  return null;
}

function profileOfEmail_(email) {
  var user = findUserByEmail_(email);
  return user ? profileOf_(user) : null;
}

function readMaster_() {
  var out = {};
  MASTER_COLUMNS.forEach(function(k) {
    out[k] = (DEFAULT_MASTER[k] || []).slice();
  });
  var sh = ss_().getSheetByName(CFG.MASTER_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  headers.forEach(function(h, c) {
    var key = String(h || '').trim();
    if (!out.hasOwnProperty(key)) return;
    values.forEach(function(r) {
      var v = String(r[c] == null ? '' : r[c]).trim();
      if (v !== '' && out[key].indexOf(v) < 0) out[key].push(v);
    });
  });
  return out;
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function round2_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function rowTotal_(item) {
  return round2_(claimableFields_().reduce(function(sum, f) {
    return sum + num_(item[f.key]);
  }, 0));
}

function fleetTotal_(item) {
  return round2_(fleetFields_().reduce(function(sum, f) {
    return sum + num_(item[f.key]);
  }, 0));
}

function isFleetItem_(item) {
  return fleetTotal_(item) !== 0 || num_(item.liter) !== 0;
}

function isEmptyItem_(item) {
  if (rowTotal_(item) !== 0) return false;
  if (fleetTotal_(item) !== 0) return false;
  if (num_(item.liter) !== 0) return false;
  return ![ 'date', 'receiptNo', 'description' ].some(function(k) {
    return String(item[k] || '').trim() !== '';
  });
}

function validate_(payload) {
  var errors = [];
  var supportDivs = readMaster_().supportDivs;
  HEADER_FIELDS.forEach(function(f) {
    if (f.required && String(payload.header[f.key] || '').trim() === '') {
      errors.push('กรุณากรอก "' + f.th + ' / ' + f.en + '"');
    }
  });
  var items = (payload.items || []).filter(function(it) {
    return !isEmptyItem_(it);
  });
  if (!items.length) errors.push('ต้องมีรายการค่าใช้จ่ายอย่างน้อย 1 บรรทัด');
  if (items.length > CFG.MAX_ROWS) {
    errors.push('แบบฟอร์มรองรับสูงสุด ' + CFG.MAX_ROWS + ' บรรทัด (กรอกมา ' + items.length + ' บรรทัด) — กรุณาแยกเป็นใบใหม่');
  }
  items.forEach(function(it, i) {
    var n = i + 1;
    if (!String(it.date || '').trim()) errors.push('บรรทัดที่ ' + n + ': ยังไม่ได้ระบุวันที่');
    if (!String(it.description || '').trim()) errors.push('บรรทัดที่ ' + n + ': ยังไม่ได้ระบุรายละเอียด');
    if (isFleetItem_(it)) {
      if (fleetTotal_(it) <= 0) errors.push('บรรทัดที่ ' + n + ' (ค่าน้ำมัน): จำนวนเงินต้องมากกว่า 0');
      if (num_(it.liter) <= 0) errors.push('บรรทัดที่ ' + n + ' (ค่าน้ำมัน): ยังไม่ได้ระบุจำนวนลิตร');
    } else if (rowTotal_(it) <= 0) errors.push('บรรทัดที่ ' + n + ': ยอดรวมต้องมากกว่า 0');
    if (it.supportDiv && supportDivs.indexOf(String(it.supportDiv)) < 0) {
      errors.push('บรรทัดที่ ' + n + ': Support Div. ต้องเป็น ' + supportDivs.join(', '));
    }
    amountFields_().forEach(function(f) {
      if (num_(it[f.key]) < 0) errors.push('บรรทัดที่ ' + n + ': "' + f.en + '" ติดลบไม่ได้');
    });
  });
  return {
    errors: errors,
    items: items
  };
}

function nextClaimId_(yearMonth) {
  var ym = String(yearMonth || '').replace(/[^0-9]/g, '').slice(0, 6) || Utilities.formatDate(new Date, Session.getScriptTimeZone(), 'yyyyMM');
  var sh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
  var prefix = 'EXP-' + ym + '-';
  var max = 0;
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var id = String(r[0] || '');
      if (id.indexOf(prefix) === 0) {
        var seq = parseInt(id.slice(prefix.length), 10);
        if (!isNaN(seq) && seq > max) max = seq;
      }
    });
  }
  return prefix + ('000' + (max + 1)).slice(-3);
}

function findClaimRow_(sh, claimId) {
  if (sh.getLastRow() < 2) return -1;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(claimId)) return i + 2;
  }
  return -1;
}

function saveClaim(token, payload) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var claimGate = requireClaimRight_();
  if (claimGate) return claimGate;
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    payload = payload || {};
    payload.header = payload.header || {};
    var status = 'Draft';
    var strict = payload.strict === true;
    var check = validate_(payload);
    if (strict && check.errors.length) {
      return {
        ok: false,
        errors: check.errors
      };
    }
    if (!strict && !String(payload.header.employee || '').trim()) {
      return {
        ok: false,
        errors: [ 'กรุณากรอกชื่อพนักงานก่อนบันทึกฉบับร่าง' ]
      };
    }
    var items = check.items;
    var me = me_();
    var user = {
      email: me.email,
      name: me.name
    };
    var now = new Date;
    var claimsSh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
    var claimId = String(payload.claimId || '').trim();
    var row = claimId ? findClaimRow_(claimsSh, claimId) : -1;
    if (claimId && row < 0) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    var existing = null;
    if (row > 0) {
      existing = rowToObject_(CLAIM_COLUMNS, claimsSh.getRange(row, 1, 1, CLAIM_COLUMNS.length).getValues()[0]);
      if (!canEditClaim_(existing, me)) {
        return {
          ok: false,
          errors: [ 'ใบเบิกนี้สถานะ "' + statusLabel_(existing.status) + '" — คุณไม่มีสิทธิ์แก้ไขในขั้นนี้' ]
        };
      }
      status = existing.status;
    }
    if (!claimId) {
      claimId = nextClaimId_(payload.header.yearMonth);
    }
    var grandTotal = round2_(items.reduce(function(s, it) {
      return s + rowTotal_(it);
    }, 0));
    var totalLiter = round2_(items.reduce(function(s, it) {
      return s + num_(it.liter);
    }, 0));
    var fuelTotal = round2_(items.reduce(function(s, it) {
      return s + fleetTotal_(it);
    }, 0));
    var record = {
      claimId: claimId,
      status: status,
      yearMonth: payload.header.yearMonth || '',
      employee: payload.header.employee || '',
      division: payload.header.division || '',
      jobTitle: payload.header.jobTitle || '',
      claimDate: payload.header.claimDate || '',
      itemCount: items.length,
      totalLiter: totalLiter,
      grandTotal: grandTotal,
      note: payload.note || '',
      createdBy: user.email,
      createdAt: now,
      updatedBy: user.email,
      updatedAt: now,
      submittedAt: '',
      pdfUrl: '',
      approverEmail: normEmail_(payload.approverEmail) || defaultApprover_(),
      approvedBy: '',
      approvedAt: '',
      rejectReason: '',
      paidBy: '',
      paidAt: ''
    };
    if (existing) {
      [ 'createdBy', 'createdAt', 'revision', 'pdfUrl', 'submittedAt', 'checkerEmail', 'checkedBy', 'checkedAt', 'checkComment', 'approvedBy', 'approvedAt', 'approveComment', 'rejectedBy', 'rejectedAt', 'rejectStage', 'rejectReason', 'paidBy', 'paidAt' ].forEach(function(k) {
        record[k] = existing[k] || record[k] || '';
      });
      record.approverEmail = record.approverEmail || existing.approverEmail || '';
      claimsSh.getRange(row, 1, 1, CLAIM_COLUMNS.length).setValues([ objectToRow_(CLAIM_COLUMNS, record) ]);
      if (normEmail_(existing.createdBy) !== me.email) {
        logApproval_(claimId, 'Edit', existing.status, existing.status, 'แก้ไขรายการระหว่างขั้นตอน');
      }
    } else {
      claimsSh.appendRow(objectToRow_(CLAIM_COLUMNS, record));
    }
    writeItems_(claimId, items);
    return {
      ok: true,
      claimId: claimId,
      status: status,
      grandTotal: grandTotal,
      totalLiter: totalLiter,
      fuelTotal: fuelTotal,
      itemCount: items.length,
      approverEmail: record.approverEmail,
      statusLabel: statusLabel_(status),
      message: existing && existing.status !== 'Draft' ? 'บันทึกการแก้ไขใบ ' + claimId + ' แล้ว (สถานะคงเดิม)' : 'บันทึกฉบับร่างแล้ว'
    };
  } finally {
    lock.releaseLock();
  }
}

function defaultApprover_() {
  var me = me_();
  return me.profile ? normEmail_(me.profile.approverEmail) : '';
}

function saveAndSubmit(token, payload) {
  payload = payload || {};
  payload.strict = true;
  var saved = saveClaim(token, payload);
  if (!saved.ok) return saved;
  var sent = submitForApproval(token, saved.claimId, {
    checkerEmail: payload.checkerEmail,
    approverEmail: payload.approverEmail,
    comment: payload.comment
  });
  if (!sent.ok) {
    sent.claimId = saved.claimId;
    sent.savedAsDraft = true;
    return sent;
  }
  return {
    ok: true,
    claimId: saved.claimId,
    status: sent.status,
    approverEmail: sent.approverEmail,
    grandTotal: saved.grandTotal,
    itemCount: saved.itemCount,
    message: sent.message
  };
}

function writeItems_(claimId, items) {
  var sh = sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS);
  deleteItems_(sh, claimId);
  if (!items.length) return;
  var rows = items.map(function(it, i) {
    var rec = {
      claimId: claimId,
      seq: i + 1,
      rowTotal: rowTotal_(it),
      reviewFlag: it.reviewFlag || '',
      reviewNote: it.reviewNote || ''
    };
    ITEM_FIELDS.forEach(function(f) {
      rec[f.key] = f.type === 'money' || f.type === 'number' ? num_(it[f.key]) : it[f.key] || '';
    });
    return objectToRow_(ITEM_COLUMNS, rec);
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, ITEM_COLUMNS.length).setValues(rows);
}

function deleteItems_(sh, claimId) {
  if (sh.getLastRow() < 2) return;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(claimId)) sh.deleteRow(i + 2);
  }
}

function rowToObject_(columns, row) {
  var o = {};
  columns.forEach(function(c, i) {
    o[c] = row[i];
  });
  return o;
}

function objectToRow_(columns, obj) {
  return columns.map(function(c) {
    var v = obj[c];
    return v === undefined || v === null ? '' : v;
  });
}

function getClaim(token, claimId) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var claimsSh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
  var row = findClaimRow_(claimsSh, claimId);
  if (row < 0) return {
    ok: false,
    errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
  };
  var claim = rowToObject_(CLAIM_COLUMNS, claimsSh.getRange(row, 1, 1, CLAIM_COLUMNS.length).getValues()[0]);
  [ 'claimDate', 'createdAt', 'updatedAt', 'submittedAt', 'checkedAt', 'approvedAt', 'rejectedAt', 'paidAt' ].forEach(function(k) {
    claim[k] = fmtDate_(claim[k]);
  });
  claim.yearMonth = fmtMonth_(claim.yearMonth);
  var me = me_();
  var perms = claimPermissions_(claim, me);
  var visible = perms.isOwner || perms.canEdit || perms.canCheck || perms.canApprove || perms.canReject || perms.canPay || me.isFinance || me.isSuperAdmin || me.isChecker || normEmail_(claim.checkerEmail) === me.email || normEmail_(claim.approverEmail) === me.email || normEmail_(claim.approvedBy) === me.email;
  if (!visible) return {
    ok: false,
    errors: [ 'คุณไม่มีสิทธิ์เข้าถึงใบเบิกนี้' ]
  };
  claim.statusLabel = statusLabel_(claim.status);
  var itemsSh = sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS);
  var items = [];
  if (itemsSh.getLastRow() > 1) {
    itemsSh.getRange(2, 1, itemsSh.getLastRow() - 1, ITEM_COLUMNS.length).getValues().filter(function(r) {
      return String(r[0]) === String(claimId);
    }).sort(function(a, b) {
      return Number(a[1]) - Number(b[1]);
    }).forEach(function(r) {
      var o = rowToObject_(ITEM_COLUMNS, r);
      o.date = fmtDate_(o.date);
      items.push(o);
    });
  }
  return safeOut_({
    ok: true,
    claim: claim,
    items: items,
    permissions: perms,
    history: getApprovalHistory(claimId).rows
  });
}

function fmtMonth_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return v === null || v === undefined ? '' : String(v);
}

function safeOut_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  if (Array.isArray(v)) return v.map(safeOut_);
  if (v && typeof v === 'object') {
    var o = {};
    Object.keys(v).forEach(function(k) {
      o[k] = safeOut_(v[k]);
    });
    return o;
  }
  return v === undefined ? '' : v;
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function listClaims(token, filter) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  filter = filter || {};
  var sh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
  if (sh.getLastRow() < 2) return {
    ok: true,
    rows: []
  };
  var identity = me_();
  var me = identity.email;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, CLAIM_COLUMNS.length).getValues().map(function(r) {
    var o = rowToObject_(CLAIM_COLUMNS, r);
    [ 'claimDate', 'createdAt', 'updatedAt', 'submittedAt', 'checkedAt', 'approvedAt', 'rejectedAt', 'paidAt' ].forEach(function(k) {
      o[k] = fmtDate_(o[k]);
    });
    o.yearMonth = fmtMonth_(o.yearMonth);
    o.statusLabel = statusLabel_(o.status);
    return o;
  }).filter(function(o) {
    if (!o.claimId) return false;
    if (!identity.isSuperAdmin && !identity.isFinance && !identity.isChecker) {
      if (!involvedInClaim_(o, identity)) return false;
    }
    if (filter.mine && me && normEmail_(o.createdBy) !== me) return false;
    if (filter.status && o.status !== filter.status) return false;
    if (filter.yearMonth && String(o.yearMonth) !== String(filter.yearMonth)) return false;
    if (filter.q) {
      var hay = [ o.claimId, o.employee, o.division, o.jobTitle, o.note ].join(' ').toLowerCase();
      if (hay.indexOf(String(filter.q).toLowerCase()) < 0) return false;
    }
    return true;
  });
  rows.sort(function(a, b) {
    return String(b.claimId).localeCompare(String(a.claimId));
  });
  return safeOut_({
    ok: true,
    rows: rows
  });
}

function deleteClaim(token, claimId) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    var me = me_();
    var sh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
    var row = findClaimRow_(sh, claimId);
    if (row < 0) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    var claim = rowToObject_(CLAIM_COLUMNS, sh.getRange(row, 1, 1, CLAIM_COLUMNS.length).getValues()[0]);
    if (normEmail_(claim.createdBy) !== me.email && !me.isSuperAdmin) {
      return {
        ok: false,
        errors: [ 'ลบได้เฉพาะใบเบิกของตัวเอง' ]
      };
    }
    if (claim.status !== 'Draft') {
      return {
        ok: false,
        errors: [ 'ลบได้เฉพาะใบที่เป็นฉบับร่าง — ใบที่ส่งแล้วให้ใช้การตีกลับแทน' ]
      };
    }
    sh.deleteRow(row);
    deleteItems_(sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS), claimId);
    logApproval_(claimId, 'Delete', 'Draft', '-', '');
    return {
      ok: true,
      message: 'ลบใบเบิก ' + claimId + ' แล้ว'
    };
  } finally {
    lock.releaseLock();
  }
}