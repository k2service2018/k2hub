function approvalSheet_() {
  return sheet_(CFG.APPROVAL_SHEET, APPROVAL_COLUMNS);
}

function logApproval_(claimId, action, fromStatus, toStatus, comment) {
  var me = me_();
  approvalSheet_().appendRow(objectToRow_(APPROVAL_COLUMNS, {
    at: new Date,
    claimId: claimId,
    action: action,
    actor: me.email,
    role: me.role,
    fromStatus: fromStatus,
    toStatus: toStatus,
    comment: comment || ''
  }));
}

function stageByStatus_(status) {
  for (var i = 0; i < STAGES.length; i++) if (STAGES[i].status === status) return STAGES[i];
  return null;
}

function loadClaimRow_(claimId) {
  var sh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
  var row = findClaimRow_(sh, claimId);
  if (row < 0) return null;
  return {
    sheet: sh,
    row: row,
    claim: rowToObject_(CLAIM_COLUMNS, sh.getRange(row, 1, 1, CLAIM_COLUMNS.length).getValues()[0])
  };
}

function writeClaimFields_(ctx, changes) {
  Object.keys(changes).forEach(function(k) {
    var c = CLAIM_COLUMNS.indexOf(k);
    if (c >= 0) ctx.sheet.getRange(ctx.row, c + 1).setValue(changes[k]);
  });
}

function canCheckClaim_(claim, me) {
  if (me.isSuperAdmin) return true;
  if (normEmail_(claim.createdBy) === me.email) return false;
  var assigned = normEmail_(claim.checkerEmail);
  if (assigned) return assigned === me.email;
  return me.isChecker;
}

function checkerPool_(checker) {
  if (checker) return [ checker ];
  var pool = roleEmails_(ROLES.ADMIN);
  return pool.length ? pool : superAdminEmails_();
}

function approverPool_(approver, me) {
  if (approver) return [ approver ];
  var pool = roleEmails_(ROLES.MANAGER).filter(function(e) {
    return !me || e !== me.email;
  });
  return pool.length ? pool : superAdminEmails_();
}

function printerPool_(claim, me) {
  var checked = normEmail_(claim.checkedBy);
  var pool = checked ? [ checked ] : roleEmails_(ROLES.ADMIN);
  return pool.filter(function(e) {
    return e && e !== normEmail_(me && me.email);
  });
}

function effectiveApprover_(claim) {
  var assigned = normEmail_(claim.approverEmail);
  if (!assigned) return '';
  if (assigned === normEmail_(claim.checkedBy)) return '';
  if (assigned === normEmail_(claim.createdBy)) return '';
  return assigned;
}

function canApproveClaim_(claim, me) {
  if (me.isSuperAdmin) return true;
  if (normEmail_(claim.createdBy) === me.email) return false;
  if (normEmail_(claim.checkedBy) === me.email) return false;
  var assigned = effectiveApprover_(claim);
  if (assigned) return assigned === me.email;
  return me.isApprover;
}

function canEditClaim_(claim, me) {
  if (me.isSuperAdmin) return true;
  if (claim.status === 'Paid') return false;
  var isOwner = normEmail_(claim.createdBy) === me.email;
  if (isOwner && EDITABLE_STATUSES.indexOf(claim.status) >= 0) return true;
  if (claim.status === 'Submitted' && canCheckClaim_(claim, me)) return true;
  if (claim.status === 'Checked' && canApproveClaim_(claim, me)) return true;
  return false;
}

function orphanStage_(role) {
  return !roleEmails_(role).length;
}

function queueCheck_(claim, me) {
  if (claim.status !== 'Submitted') return false;
  if (normEmail_(claim.createdBy) === me.email) return false;
  var assigned = normEmail_(claim.checkerEmail);
  if (assigned) return assigned === me.email;
  if (me.checkerDuty) return true;
  return me.isSuperAdmin && orphanStage_(ROLES.ADMIN);
}

function queueApprove_(claim, me) {
  if (claim.status !== 'Checked') return false;
  if (normEmail_(claim.createdBy) === me.email) return false;
  if (normEmail_(claim.checkedBy) === me.email) return false;
  var assigned = effectiveApprover_(claim);
  if (assigned) return assigned === me.email;
  if (me.approverDuty) return true;
  return me.isSuperAdmin && orphanStage_(ROLES.MANAGER);
}

function queuePay_(claim, me) {
  if (!CFG.REQUIRE_PAYMENT_STEP) return false;
  if (claim.status !== 'Approved') return false;
  if (me.financeDuty) return true;
  return me.isSuperAdmin && orphanStage_(ROLES.FINANCE) && !CFG.FINANCE_EMAILS.length;
}

function claimPermissions_(claim, me) {
  me = me || me_();
  var isOwner = normEmail_(claim.createdBy) === me.email;
  return {
    isOwner: isOwner,
    queueCheck: queueCheck_(claim, me),
    queueApprove: queueApprove_(claim, me),
    queuePay: queuePay_(claim, me),
    canEdit: canEditClaim_(claim, me),
    canSubmit: isOwner && EDITABLE_STATUSES.indexOf(claim.status) >= 0,
    canCheck: claim.status === 'Submitted' && canCheckClaim_(claim, me),
    canApprove: claim.status === 'Checked' && canApproveClaim_(claim, me),
    canReject: claim.status === 'Submitted' && canCheckClaim_(claim, me) || claim.status === 'Checked' && canApproveClaim_(claim, me),
    canPay: claim.status === 'Approved' && me.isFinance,
    canDelete: isOwner && claim.status === 'Draft'
  };
}

function submitForApproval(token, claimId, options) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var claimGate = requireClaimRight_();
  if (claimGate) return claimGate;
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    options = options || {};
    var me = me_();
    var ctx = loadClaimRow_(claimId);
    if (!ctx) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    var claim = ctx.claim;
    if (normEmail_(claim.createdBy) !== me.email && !me.isSuperAdmin) {
      return {
        ok: false,
        errors: [ 'ส่งขออนุมัติได้เฉพาะใบเบิกของตัวเอง' ]
      };
    }
    if (EDITABLE_STATUSES.indexOf(claim.status) < 0) {
      return {
        ok: false,
        errors: [ 'ใบเบิกสถานะ "' + statusLabel_(claim.status) + '" ส่งซ้ำไม่ได้' ]
      };
    }
    var checker = normEmail_(options.checkerEmail) || normEmail_(claim.checkerEmail) || normEmail_(me.profile && me.profile.checkerEmail);
    var approver = normEmail_(options.approverEmail) || normEmail_(claim.approverEmail) || normEmail_(me.profile && me.profile.approverEmail);
    var checkerPool = checkerPool_(checker);
    if (!checkerPool.length) {
      return {
        ok: false,
        errors: [ 'ยังไม่มีผู้ใช้บทบาท Admin ในระบบ — ให้ผู้ดูแลตั้งผู้ตรวจสอบก่อน' ]
      };
    }
    if (approver && approver === me.email && !me.isSuperAdmin) {
      return {
        ok: false,
        errors: [ 'อนุมัติใบเบิกของตัวเองไม่ได้ — ให้ผู้ดูแลตั้ง Manager คนอื่นให้บัญชีนี้' ]
      };
    }
    var approverPool = approverPool_(approver, me);
    if (!approverPool.length) {
      return {
        ok: false,
        errors: [ 'ยังไม่มีผู้ใช้บทบาท Manager ในระบบ — ให้ผู้ดูแลตั้งผู้อนุมัติก่อน' ]
      };
    }
    var resubmit = claim.status === 'Rejected';
    var now = new Date;
    writeClaimFields_(ctx, {
      status: 'Submitted',
      revision: (Number(claim.revision) || 0) + (resubmit ? 1 : 0),
      checkerEmail: checker,
      approverEmail: approver,
      submittedAt: now,
      updatedBy: me.email,
      updatedAt: now,
      checkedBy: '',
      checkedAt: '',
      checkComment: '',
      approvedBy: '',
      approvedAt: '',
      approveComment: '',
      rejectedBy: '',
      rejectedAt: '',
      rejectStage: '',
      rejectReason: ''
    });
    if (resubmit) clearItemReviews_(claimId);
    logApproval_(claimId, resubmit ? 'Resubmit' : 'Submit', claim.status, 'Submitted', options.comment || '');
    var fresh = loadClaimRow_(claimId).claim;
    notifyStage_(fresh, checkerPool, 'check', me, '');
    return {
      ok: true,
      status: 'Submitted',
      claimId: claimId,
      checkerEmail: checker,
      approverEmail: approver,
      message: resubmit ? 'ส่งใบเบิก ' + claimId + ' ให้ตรวจสอบใหม่แล้ว (แก้ไขครั้งที่ ' + ((Number(claim.revision) || 0) + 1) + ')' : 'ส่งใบเบิก ' + claimId + ' ให้ Admin ตรวจสอบแล้ว'
    };
  } finally {
    lock.releaseLock();
  }
}

function getClaimForReview(token, claimId) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var res = getClaim(token, claimId);
  if (!res.ok) return res;
  var me = me_();
  if (!res.permissions.canCheck && !res.permissions.canApprove && !me.isSuperAdmin && !me.isFinance) {
    return {
      ok: false,
      errors: [ 'คุณไม่มีสิทธิ์ตรวจสอบใบเบิกนี้' ]
    };
  }
  var reviewed = res.items.filter(function(it) {
    return it.reviewFlag;
  }).length;
  var flagged = res.items.filter(function(it) {
    return it.reviewFlag === 'flag';
  }).length;
  res.review = {
    total: res.items.length,
    reviewed: reviewed,
    flagged: flagged,
    passed: reviewed - flagged,
    complete: res.items.length > 0 && reviewed === res.items.length,
    flags: REVIEW_FLAGS
  };
  res.requester = profileOfEmail_(res.claim.createdBy);
  return safeOut_(res);
}

function saveItemReview(token, claimId, reviews, comment) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    var me = me_();
    var ctx = loadClaimRow_(claimId);
    if (!ctx) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    var perms = claimPermissions_(ctx.claim, me);
    if (!perms.canCheck && !perms.canApprove && !me.isSuperAdmin) {
      return {
        ok: false,
        errors: [ 'คุณไม่มีสิทธิ์ตรวจสอบใบเบิกนี้' ]
      };
    }
    var validFlags = REVIEW_FLAGS.map(function(f) {
      return f.key;
    });
    var bySeq = {};
    (reviews || []).forEach(function(r) {
      var flag = String(r.flag || '');
      if (validFlags.indexOf(flag) < 0) flag = '';
      bySeq[String(r.seq)] = {
        flag: flag,
        note: String(r.note || '')
      };
    });
    var sh = sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS);
    if (sh.getLastRow() < 2) return {
      ok: false,
      errors: [ 'ใบเบิกนี้ไม่มีรายการ' ]
    };
    var flagCol = ITEM_COLUMNS.indexOf('reviewFlag') + 1;
    var noteCol = ITEM_COLUMNS.indexOf('reviewNote') + 1;
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, ITEM_COLUMNS.length).getValues();
    var saved = 0;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) !== String(claimId)) continue;
      var hit = bySeq[String(values[i][1])];
      if (!hit) continue;
      sh.getRange(i + 2, flagCol).setValue(hit.flag);
      sh.getRange(i + 2, noteCol).setValue(hit.note);
      saved++;
    }
    if (comment !== undefined && comment !== null) {
      writeClaimFields_(ctx, {
        checkComment: comment,
        updatedBy: me.email,
        updatedAt: new Date
      });
    }
    var flagged = Object.keys(bySeq).filter(function(k) {
      return bySeq[k].flag === 'flag';
    }).length;
    return {
      ok: true,
      saved: saved,
      flagged: flagged,
      message: 'บันทึกผลการตรวจ ' + saved + ' รายการแล้ว' + (flagged ? ' (ติดปัญหา ' + flagged + ' รายการ)' : '')
    };
  } finally {
    lock.releaseLock();
  }
}

function checkClaim(token, claimId, comment) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    var me = me_();
    var ctx = loadClaimRow_(claimId);
    if (!ctx) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    var claim = ctx.claim;
    if (claim.status !== 'Submitted') {
      return {
        ok: false,
        errors: [ 'ใบเบิกนี้สถานะ "' + statusLabel_(claim.status) + '" ไม่อยู่ระหว่างรอตรวจสอบ' ]
      };
    }
    if (!canCheckClaim_(claim, me)) {
      return {
        ok: false,
        errors: [ 'คุณไม่มีสิทธิ์ตรวจสอบใบเบิกนี้' ]
      };
    }
    var approver = effectiveApprover_({
      approverEmail: claim.approverEmail,
      createdBy: claim.createdBy,
      checkedBy: me.email
    });
    var approverPool = approverPool_(approver, {
      email: normEmail_(claim.createdBy),
      isSuperAdmin: false
    }).filter(function(e) {
      return e !== me.email;
    });
    if (!approverPool.length) return {
      ok: false,
      errors: [ 'ยังไม่มีผู้ใช้บทบาท Manager ในระบบ — ให้ผู้ดูแลตั้งผู้อนุมัติก่อน' ]
    };
    var flagged = claimItems_(claimId).filter(function(it) {
      return it.reviewFlag === 'flag';
    });
    if (flagged.length) {
      return {
        ok: false,
        errors: [ 'ยังมี ' + flagged.length + ' รายการที่ติดธง "ติดปัญหา" (บรรทัดที่ ' + flagged.map(function(it) {
          return it.seq;
        }).join(', ') + ') — กรุณาแก้ไขรายการหรือตีกลับให้ผู้จัดทำก่อน' ]
      };
    }
    var now = new Date;
    writeClaimFields_(ctx, {
      status: 'Checked',
      checkedBy: me.email,
      checkedAt: now,
      checkComment: comment || '',
      approverEmail: approver,
      updatedBy: me.email,
      updatedAt: now
    });
    logApproval_(claimId, 'Check', 'Submitted', 'Checked', comment || '');
    var fresh = loadClaimRow_(claimId).claim;
    notifyStage_(fresh, approverPool, 'approve', me, comment);
    notifyOwner_(fresh, 'Checked', me, comment, null);
    return {
      ok: true,
      status: 'Checked',
      message: 'ตรวจสอบผ่าน — ส่งให้ ' + approverPool.join(', ') + ' อนุมัติแล้ว'
    };
  } finally {
    lock.releaseLock();
  }
}

function approveClaim(token, claimId, comment) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    var me = me_();
    var ctx = loadClaimRow_(claimId);
    if (!ctx) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    var claim = ctx.claim;
    if (claim.status !== 'Checked') {
      return {
        ok: false,
        errors: [ claim.status === 'Submitted' ? 'ใบเบิกนี้ต้องผ่าน Admin ตรวจสอบก่อนจึงอนุมัติได้' : 'ใบเบิกนี้สถานะ "' + statusLabel_(claim.status) + '" ไม่อยู่ระหว่างรออนุมัติ' ]
      };
    }
    if (!canApproveClaim_(claim, me)) {
      return {
        ok: false,
        errors: [ 'คุณไม่ใช่ผู้อนุมัติของใบเบิกนี้' ]
      };
    }
    var now = new Date;
    writeClaimFields_(ctx, {
      status: 'Approved',
      approvedBy: me.email,
      approvedAt: now,
      approveComment: comment || '',
      updatedBy: me.email,
      updatedAt: now
    });
    logApproval_(claimId, 'Approve', 'Checked', 'Approved', comment || '');
    var fresh = loadClaimRow_(claimId).claim;
    var pdf = attachPdfIfEnabled_(claimId);
    notifyOwner_(fresh, 'Approved', me, comment, pdf);
    notifyPrint_(fresh, printerPool_(fresh, me), me, pdf);
    if (CFG.REQUIRE_PAYMENT_STEP) {
      var financePool = roleEmails_(ROLES.FINANCE).concat(CFG.FINANCE_EMAILS.map(normEmail_));
      notifyStage_(fresh, financePool, 'pay', me, comment);
    }
    return {
      ok: true,
      status: 'Approved',
      message: 'อนุมัติใบเบิก ' + claimId + ' แล้ว'
    };
  } finally {
    lock.releaseLock();
  }
}

function rejectClaim(token, claimId, reason) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  if (!String(reason || '').trim()) {
    return {
      ok: false,
      errors: [ 'กรุณาระบุเหตุผลที่ตีกลับ เพื่อให้ผู้จัดทำแก้ไขได้ถูกจุด' ]
    };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    var me = me_();
    var ctx = loadClaimRow_(claimId);
    if (!ctx) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    var claim = ctx.claim;
    var perms = claimPermissions_(claim, me);
    if (!perms.canReject) {
      return {
        ok: false,
        errors: [ 'ใบเบิกนี้ไม่อยู่ในขั้นที่คุณตีกลับได้ (สถานะ: ' + statusLabel_(claim.status) + ')' ]
      };
    }
    var stage = claim.status === 'Submitted' ? 'ตรวจสอบ (Admin)' : 'อนุมัติ (Manager)';
    var now = new Date;
    writeClaimFields_(ctx, {
      status: 'Rejected',
      rejectedBy: me.email,
      rejectedAt: now,
      rejectStage: stage,
      rejectReason: reason,
      updatedBy: me.email,
      updatedAt: now
    });
    logApproval_(claimId, 'Reject', claim.status, 'Rejected', reason);
    var fresh = loadClaimRow_(claimId).claim;
    notifyOwner_(fresh, 'Rejected', me, reason, null);
    return {
      ok: true,
      status: 'Rejected',
      message: 'ตีกลับใบเบิก ' + claimId + ' แล้ว — ผู้จัดทำแก้ไขในเลขที่เดิมได้'
    };
  } finally {
    lock.releaseLock();
  }
}

function markPaid(token, claimId, comment) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var lock = LockService.getScriptLock();
  lock.waitLock(3e4);
  try {
    var me = me_();
    if (!me.isFinance) return {
      ok: false,
      errors: [ 'เฉพาะฝ่ายบัญชี/การเงินเท่านั้น' ]
    };
    var ctx = loadClaimRow_(claimId);
    if (!ctx) return {
      ok: false,
      errors: [ 'ไม่พบใบเบิกเลขที่ ' + claimId ]
    };
    if (ctx.claim.status !== 'Approved') {
      return {
        ok: false,
        errors: [ 'ต้องผ่านการอนุมัติจาก Manager ก่อนจึงบันทึกการจ่ายได้' ]
      };
    }
    var now = new Date;
    writeClaimFields_(ctx, {
      status: 'Paid',
      paidBy: me.email,
      paidAt: now,
      updatedBy: me.email,
      updatedAt: now
    });
    logApproval_(claimId, 'Pay', 'Approved', 'Paid', comment || '');
    notifyOwner_(loadClaimRow_(claimId).claim, 'Paid', me, comment, null);
    return {
      ok: true,
      status: 'Paid',
      message: 'บันทึกการจ่ายเงินใบ ' + claimId + ' แล้ว'
    };
  } finally {
    lock.releaseLock();
  }
}

function listPendingApprovals(token) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var me = me_();
  var res = listClaims(token, {
    all: true
  });
  if (!res.ok) return res;
  var rows = res.rows.map(function(c) {
    c.permissions = claimPermissions_(c, me);
    c.statusLabel = statusLabel_(c.status);
    return c;
  }).filter(function(c) {
    return c.permissions.queueCheck || c.permissions.queueApprove || c.permissions.queuePay;
  });
  return safeOut_({
    ok: true,
    rows: rows,
    counts: {
      toCheck: rows.filter(function(c) {
        return c.permissions.queueCheck;
      }).length,
      toApprove: rows.filter(function(c) {
        return c.permissions.queueApprove;
      }).length,
      toPay: rows.filter(function(c) {
        return c.permissions.queuePay;
      }).length
    }
  });
}

function getApprovalHistory(claimId) {
  var sh = approvalSheet_();
  if (sh.getLastRow() < 2) return {
    ok: true,
    rows: []
  };
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, APPROVAL_COLUMNS.length).getValues().map(function(r) {
    var o = rowToObject_(APPROVAL_COLUMNS, r);
    o.at = o.at instanceof Date ? Utilities.formatDate(o.at, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(o.at || '');
    return o;
  }).filter(function(o) {
    return String(o.claimId) === String(claimId);
  });
  return safeOut_({
    ok: true,
    rows: rows
  });
}

function statusLabel_(status, lang) {
  if (lang === 'en') return STATUS_LABELS_EN[status] || status || '';
  return STATUS_LABELS[status] || status || '';
}

function claimItems_(claimId) {
  var sh = sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ITEM_COLUMNS.length).getValues().filter(function(r) {
    return String(r[0]) === String(claimId);
  }).map(function(r) {
    return rowToObject_(ITEM_COLUMNS, r);
  }).sort(function(a, b) {
    return Number(a.seq) - Number(b.seq);
  });
}

function clearItemReviews_(claimId) {
  var sh = sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS);
  if (sh.getLastRow() < 2) return;
  var flagCol = ITEM_COLUMNS.indexOf('reviewFlag') + 1;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(claimId)) {
      sh.getRange(i + 2, flagCol, 1, 2).clearContent();
    }
  }
}

function attachPdfIfEnabled_(claimId) {
  if (!CFG.ATTACH_PDF_TO_EMAIL) return null;
  try {
    var res = exportPdfInternal_(claimId);
    if (!res.ok) return null;
    var id = String(res.url).match(/[-\w]{25,}/);
    return id ? DriveApp.getFileById(id[0]).getBlob().setName(claimId + '.pdf') : null;
  } catch (err) {
    console.error('แนบ PDF ไม่สำเร็จ: ' + err);
    return null;
  }
}

function claimSummaryHtml_(claim, lang) {
  var rows = [ [ mailText_('sum_claimno', lang), escapeHtml_(claim.claimId) + (Number(claim.revision) ? ' (' + mailText_('sum_revision', lang) + claim.revision + ')' : '') ], [ mailText_('sum_employee', lang), escapeHtml_(claim.employee) + ' &lt;' + escapeHtml_(claim.createdBy) + '&gt;' ], [ mailText_('sum_division', lang), escapeHtml_(claim.division) ], [ mailText_('sum_month', lang), escapeHtml_(fmtMonth_(claim.yearMonth)) ], [ mailText_('sum_items', lang), escapeHtml_(claim.itemCount) ], [ mailText_('sum_total', lang), '<b>' + Number(claim.grandTotal || 0).toLocaleString('en-US', {
    minimumFractionDigits: 2
  }) + ' ' + CFG.CURRENCY + '</b>' ], [ mailText_('sum_status', lang), escapeHtml_(statusLabel_(claim.status, lang)) ] ];
  return '<table cellpadding="6" style="border-collapse:collapse;font-family:sans-serif;font-size:13px">' + rows.map(function(r) {
    return '<tr><td style="color:#5a6577">' + r[0] + '</td><td>' + r[1] + '</td></tr>';
  }).join('') + '</table>';
}

function notifyStage_(claim, recipients, stageKey, actor, comment) {
  var stage = null;
  for (var i = 0; i < STAGES.length; i++) if (STAGES[i].key === stageKey) stage = STAGES[i];
  if (!stage) return;
  var url = webAppUrl_();
  var link = url ? url + '?page=inbox&id=' + encodeURIComponent(claim.claimId) : '';
  var headKey = {
    check: 'head_check',
    approve: 'head_approve',
    pay: 'head_pay'
  }[stageKey] || 'head_action';
  var pdf = stageKey === 'check' || stageKey === 'approve' ? attachPdfIfEnabled_(claim.claimId) : null;
  sendMailLocalized_(recipients, function(lang) {
    return {
      subject: '[' + (lang === 'en' ? stage.actionEn : stage.action) + '] ' + mailText_('subj_claim', lang) + ' ' + claim.claimId + ' — ' + claim.employee,
      html: '<p style="font-size:15px"><b>' + mailText_(headKey, lang) + '</b></p>' + claimSummaryHtml_(claim, lang) + '<p style="font-size:13px;color:#5a6577">' + mailText_('lbl_sent_by', lang) + escapeHtml_(actor.name || actor.email) + '</p>' + (comment ? '<p><b>' + mailText_('lbl_note', lang) + ':</b> ' + escapeHtml_(comment) + '</p>' : '') + actionButtonHtml_(link, mailText_('btn_inbox', lang)) + (pdf ? '<p style="color:#666;font-size:12px">' + mailText_('pdf_attached', lang) + '</p>' : '')
    };
  }, pdf ? [ pdf ] : []);
}

function notifyPrint_(claim, recipients, actor, pdfBlob) {
  if (!recipients || !recipients.length) return;
  var url = webAppUrl_();
  var link = url ? url + '?id=' + encodeURIComponent(claim.claimId) : '';
  sendMailLocalized_(recipients, function(lang) {
    return {
      subject: '[' + mailText_('subj_print', lang) + '] ' + mailText_('subj_claim', lang) + ' ' + claim.claimId + ' — ' + claim.employee,
      html: '<p style="font-size:15px"><b>' + mailText_('head_print', lang) + '</b></p>' + claimSummaryHtml_(claim, lang) + '<p style="font-size:13px;color:#5a6577">' + mailText_('lbl_approved_by', lang) + escapeHtml_(actor.name || actor.email) + '</p>' + actionButtonHtml_(link, mailText_('btn_print', lang)) + (pdfBlob ? '<p style="color:#666;font-size:12px">' + mailText_('pdf_attached', lang) + '</p>' : '')
    };
  }, pdfBlob ? [ pdfBlob ] : []);
}

function notifyOwner_(claim, toStatus, actor, comment, pdfBlob) {
  var url = webAppUrl_();
  var link = url ? url + '?id=' + encodeURIComponent(claim.claimId) : '';
  var known = [ 'Checked', 'Approved', 'Rejected', 'Paid' ].indexOf(toStatus) >= 0;
  sendMailLocalized_(claim.createdBy, function(lang) {
    var heading = known ? mailText_('own_head_' + toStatus, lang) : mailText_('own_head_other', lang) + statusLabel_(toStatus, lang);
    var subjectWord = known ? mailText_('own_subj_' + toStatus, lang) : statusLabel_(toStatus, lang);
    return {
      subject: '[' + mailText_('subj_yours', lang) + '] ' + subjectWord + ' — ' + claim.claimId,
      html: '<p style="font-size:15px"><b>' + escapeHtml_(heading) + '</b></p>' + claimSummaryHtml_(claim, lang) + '<p><b>' + mailText_('lbl_action_by', lang) + '</b> ' + escapeHtml_(actor.name || actor.email) + '</p>' + (comment ? '<p style="background:#fff5f4;border-left:4px solid #b42318;padding:10px">' + '<b>' + mailText_(toStatus === 'Rejected' ? 'lbl_reject_reason' : 'lbl_note', lang) + ':</b><br>' + escapeHtml_(comment) + '</p>' : '') + actionButtonHtml_(link, mailText_(toStatus === 'Rejected' ? 'btn_revise' : 'btn_open', lang))
    };
  }, pdfBlob ? [ pdfBlob ] : []);
}