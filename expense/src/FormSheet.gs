var FIRST_ITEM_ROW = 9;

var TOTAL_ROW = 34;

var LAST_COL = 20;

function colLetterToIndex_(letter) {
  var n = 0;
  for (var i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}

function indexToColLetter_(index) {
  var s = '';
  while (index > 0) {
    var r = (index - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    index = Math.floor((index - 1) / 26);
  }
  return s;
}

function signerName_(email) {
  var p = profileOfEmail_(email);
  return p && p.name ? p.name : String(email || '');
}

function claimForRender_(claimId) {
  var ctx = loadClaimRow_(claimId);
  if (!ctx) throw new Error('ไม่พบใบเบิกเลขที่ ' + claimId);
  var claim = ctx.claim;
  [ 'claimDate', 'submittedAt', 'checkedAt', 'approvedAt', 'rejectedAt', 'paidAt' ].forEach(function(k) {
    claim[k] = fmtDate_(claim[k]);
  });
  claim.yearMonth = fmtMonth_(claim.yearMonth);
  var items = claimItems_(claimId).map(function(o) {
    o.date = fmtDate_(o.date);
    return o;
  });
  return {
    claim: claim,
    items: items
  };
}

function insertLogo_(sh) {
  var id = secret_('LOGO_FILE_ID');
  if (!id) return;
  try {
    var img = sh.insertImage(DriveApp.getFileById(id).getBlob(), LAST_COL - 2, 3, 8, 14);
    img.setWidth(100).setHeight(85);
  } catch (e) {}
}

function renderFormSheet_(claimId) {
  var data = claimForRender_(claimId);
  var book = ss_();
  var sh = book.getSheetByName(CFG.FORM_SHEET);
  if (sh) book.deleteSheet(sh);
  sh = book.insertSheet(CFG.FORM_SHEET);
  var claim = data.claim;
  var items = data.items;
  sh.getRange(1, 1, 45, LAST_COL).clearFormat();
  sh.setHiddenGridlines(true);
  HEADER_FIELDS.forEach(function(f, i) {
    var r = i + 1;
    sh.getRange(r, 1, 1, 2).merge().setValue(f.en).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(r, 3, 1, 3).merge().setValue(claim[f.key] || '').setFontSize(9).setNumberFormat('@').setHorizontalAlignment('left').setVerticalAlignment('middle').setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange(2, 6, 4, 7).merge().setValue((CFG.COMPANY ? CFG.COMPANY + '\n' : '') + CFG.DOC_TITLE).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(1, LAST_COL - 5, 1, 6).merge().setValue(CFG.DOC_NO).setFontSize(8).setHorizontalAlignment('right');
  sh.getRange(2, LAST_COL - 5, 1, 6).merge().setValue(claim.claimId).setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');
  insertLogo_(sh);
  var groupSpans = {};
  ITEM_FIELDS.forEach(function(f) {
    if (!f.groupKey) return;
    var c = colLetterToIndex_(f.col);
    var g = groupSpans[f.groupKey];
    if (!g) groupSpans[f.groupKey] = {
      from: c,
      to: c
    }; else {
      g.from = Math.min(g.from, c);
      g.to = Math.max(g.to, c);
    }
  });
  GROUPS.forEach(function(g) {
    var span = groupSpans[g.key];
    if (!span) return;
    var label = g.en + (g.account ? ' / ' + g.account : '');
    var rng = span.from === span.to ? sh.getRange(7, span.from) : sh.getRange(7, span.from, 1, span.to - span.from + 1).merge();
    rng.setValue(label).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true).setBackground(g.color).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange(8, 1).setValue('No.');
  ITEM_FIELDS.forEach(function(f) {
    sh.getRange(8, colLetterToIndex_(f.col)).setValue(f.en);
  });
  sh.getRange(8, LAST_COL).setValue('Total');
  sh.getRange(8, 1, 1, LAST_COL).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setBackground('#dce6f1').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  var body = [];
  for (var i = 0; i < CFG.MAX_ROWS; i++) {
    var it = items[i];
    var row = new Array(LAST_COL).fill('');
    row[0] = i + 1;
    if (it) {
      ITEM_FIELDS.forEach(function(f) {
        var v = it[f.key];
        var c = colLetterToIndex_(f.col) - 1;
        if (f.type === 'money' || f.type === 'number') row[c] = num_(v) || ''; else row[c] = v || '';
      });
      row[LAST_COL - 1] = num_(it.rowTotal) || '';
    }
    body.push(row);
  }
  sh.getRange(FIRST_ITEM_ROW, 1, CFG.MAX_ROWS, LAST_COL).setValues(body).setFontSize(8).setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#808080', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(FIRST_ITEM_ROW, 1, CFG.MAX_ROWS, 1).setHorizontalAlignment('center');
  sh.getRange(FIRST_ITEM_ROW, 2, CFG.MAX_ROWS, 1).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');
  var supportDivs = readMaster_().supportDivs;
  sh.getRange(FIRST_ITEM_ROW, 5, CFG.MAX_ROWS, 1).setHorizontalAlignment('center').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(supportDivs, true).setAllowInvalid(false).setHelpText('เลือกแผนกที่รับผิดชอบค่าใช้จ่าย: ' + supportDivs.join(' / ')).build());
  sh.getRange(FIRST_ITEM_ROW, 6, CFG.MAX_ROWS, 1).setNumberFormat('#,##0.00');
  sh.getRange(FIRST_ITEM_ROW, 7, CFG.MAX_ROWS, LAST_COL - 6).setNumberFormat('#,##0.00');
  GROUPS.forEach(function(g) {
    var span = groupSpans[g.key];
    if (!span || g.key === 'support') return;
    sh.getRange(FIRST_ITEM_ROW, span.from, CFG.MAX_ROWS, span.to - span.from + 1).setBackground(g.color);
  });
  sh.getRange(TOTAL_ROW, 1, 1, 4).merge().setValue('Total').setFontWeight('bold').setHorizontalAlignment('center');
  var last = FIRST_ITEM_ROW + CFG.MAX_ROWS - 1;
  for (var c = 6; c <= LAST_COL; c++) {
    var L = indexToColLetter_(c);
    sh.getRange(TOTAL_ROW, c).setFormula('=SUM(' + L + FIRST_ITEM_ROW + ':' + L + last + ')');
  }
  sh.getRange(TOTAL_ROW, 1, 1, LAST_COL).setFontSize(9).setFontWeight('bold').setBackground('#dce6f1').setNumberFormat('#,##0.00').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(TOTAL_ROW, 1, 1, 4).setNumberFormat('@');
  sh.getRange(36, 1).setValue('Work Flow :').setFontWeight('bold').setFontSize(8);
  WORKFLOW.forEach(function(w, i) {
    sh.getRange(36 + i, 2, 1, 3).merge().setValue(i + 1 + '. ' + w).setFontSize(8);
  });
  var signCols = {
    prepared: 6,
    received: 11,
    approved: 16
  };
  SIGN_BLOCKS.forEach(function(b) {
    var c = signCols[b.key];
    var by = claim[b.byField] || '';
    var at = claim[b.atField] || '';
    sh.getRange(37, c, 1, 4).merge().setBorder(null, null, true, null, null, null, '#444444', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.getRange(38, c, 1, 4).merge().setValue(b.en).setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(39, c, 1, 4).merge().setValue(by ? signerName_(by) : '……......./……........../…...........').setFontSize(9).setHorizontalAlignment('center');
    sh.getRange(40, c, 1, 4).merge().setValue(at ? 'วันที่ ' + at : '').setFontSize(8).setFontColor('#5a6577').setHorizontalAlignment('center');
  });
  var statusText = 'สถานะ: ' + statusLabel_(claim.status) + (Number(claim.revision) ? '  •  แก้ไขครั้งที่ ' + claim.revision : '') + (claim.status === 'Rejected' && claim.rejectReason ? '  •  เหตุผลที่ตีกลับ: ' + claim.rejectReason : '');
  sh.getRange(42, 1, 1, LAST_COL).merge().setValue(statusText).setFontSize(8).setFontColor('#5a6577');
  sh.setColumnWidth(1, 28);
  sh.setColumnWidth(2, 66);
  sh.setColumnWidth(3, 156);
  sh.setColumnWidth(4, 202);
  sh.setColumnWidth(5, 54);
  sh.setColumnWidth(6, 42);
  for (var cc = 7; cc <= LAST_COL - 1; cc++) sh.setColumnWidth(cc, 58);
  sh.setColumnWidth(LAST_COL, 70);
  sh.setRowHeight(1, 21);
  sh.setRowHeight(2, 21);
  for (var hr = 3; hr <= 6; hr++) sh.setRowHeight(hr, 28);
  sh.setRowHeight(7, 30);
  sh.setRowHeight(8, 34);
  for (var sr = 35; sr <= 37; sr++) sh.setRowHeight(sr, 26);
  SpreadsheetApp.flush();
  return sh;
}

function exportPdf(token, claimId) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var visible = getClaim(token, claimId);
  if (!visible.ok) return visible;
  return exportPdfInternal_(claimId);
}

function exportPdfInternal_(claimId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(6e4);
  var sh = null;
  try {
    sh = renderFormSheet_(claimId);
    var book = ss_();
    var url = 'https://docs.google.com/spreadsheets/d/' + book.getId() + '/export?' + [ 'format=pdf', 'gid=' + sh.getSheetId(), 'portrait=false', 'size=A4', 'fitw=true', 'scale=4', 'gridlines=false', 'printtitle=false', 'sheetnames=false', 'pagenumbers=false', 'top_margin=0.35', 'bottom_margin=0.35', 'left_margin=0.18', 'right_margin=0.18' ].join('&');
    var res = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      throw new Error('สร้าง PDF ไม่สำเร็จ (HTTP ' + res.getResponseCode() + ')');
    }
    var name = claimId + '.pdf';
    var blob = res.getBlob().setName(name);
    var folder = secret_('PDF_FOLDER_ID') ? DriveApp.getFolderById(secret_('PDF_FOLDER_ID')) : DriveApp.getRootFolder();
    var dup = folder.getFilesByName(name);
    while (dup.hasNext()) dup.next().setTrashed(true);
    var file = folder.createFile(blob);
    var fileUrl = file.getUrl();
    var claimsSh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
    var row = findClaimRow_(claimsSh, claimId);
    if (row > 0) claimsSh.getRange(row, CLAIM_COLUMNS.indexOf('pdfUrl') + 1).setValue(fileUrl);
    return {
      ok: true,
      url: fileUrl,
      name: name,
      bytes: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    return {
      ok: false,
      errors: [ String(err.message || err) ]
    };
  } finally {
    if (sh) {
      try {
        ss_().deleteSheet(sh);
      } catch (e) {}
    }
    lock.releaseLock();
  }
}