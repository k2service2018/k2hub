var MAIL_EN_ROLES = [ 'Manager' ];

var STATUS_LABELS_EN = {
  Draft: 'Draft',
  Submitted: 'Awaiting Admin check',
  Checked: 'Awaiting Manager approval',
  Approved: 'Approved',
  Rejected: 'Returned for revision',
  Paid: 'Paid'
};

var MAIL_TEXT = {
  sum_claimno: { th: 'เลขที่ใบเบิก', en: 'Claim no.' },
  sum_revision: { th: 'แก้ไขครั้งที่ ', en: 'revision ' },
  sum_employee: { th: 'ผู้จัดทำ', en: 'Prepared by' },
  sum_division: { th: 'ฝ่าย / แผนก', en: 'Division' },
  sum_month: { th: 'ปี / เดือน', en: 'Year / month' },
  sum_items: { th: 'จำนวนรายการ', en: 'Items' },
  sum_total: { th: 'ยอดรวม', en: 'Grand total' },
  sum_status: { th: 'สถานะปัจจุบัน', en: 'Current status' },
  head_check: { th: 'มีใบเบิกรอคุณตรวจสอบ', en: 'An expense claim is waiting for your check' },
  head_approve: { th: 'มีใบเบิกผ่านการตรวจสอบแล้ว รอคุณอนุมัติ', en: 'An expense claim has been checked and is waiting for your approval' },
  head_pay: { th: 'มีใบเบิกได้รับอนุมัติแล้ว รอบันทึกการจ่ายเงิน', en: 'An approved expense claim is waiting for payment to be recorded' },
  head_action: { th: 'มีใบเบิกรอคุณดำเนินการ', en: 'An expense claim is waiting for your action' },
  head_print: { th: 'ใบเบิกได้รับอนุมัติแล้ว รอพิมพ์เอกสาร', en: 'This expense claim is approved and ready to print' },
  subj_claim: { th: 'ใบเบิก', en: 'Expense claim' },
  subj_print: { th: 'พิมพ์เอกสาร', en: 'Print' },
  subj_yours: { th: 'ใบเบิกของคุณ', en: 'Your claim' },
  lbl_sent_by: { th: 'ส่งโดย ', en: 'Sent by ' },
  lbl_approved_by: { th: 'อนุมัติโดย ', en: 'Approved by ' },
  lbl_action_by: { th: 'ดำเนินการโดย:', en: 'Action by:' },
  lbl_note: { th: 'หมายเหตุ', en: 'Note' },
  lbl_reject_reason: { th: 'เหตุผลที่ตีกลับ', en: 'Reason for return' },
  btn_inbox: { th: 'เปิดรายการที่รอดำเนินการ', en: 'Open my action list' },
  btn_print: { th: 'เปิดใบเบิกเพื่อพิมพ์', en: 'Open the claim to print' },
  btn_open: { th: 'เปิดใบเบิก', en: 'Open the claim' },
  btn_revise: { th: 'เปิดแก้ไขใบเบิก', en: 'Open and revise' },
  btn_login: { th: 'เข้าสู่ระบบ', en: 'Sign in' },
  btn_reset: { th: 'ตั้งรหัสผ่านใหม่', en: 'Set a new password' },
  pdf_attached: { th: 'เอกสารฉบับเต็มแนบมาในรูปแบบ PDF', en: 'The full document is attached as a PDF.' },
  own_head_Checked: { th: 'ใบเบิกของคุณผ่านการตรวจสอบแล้ว — ส่งต่อให้ Manager อนุมัติ', en: 'Your expense claim has been checked and sent to the Manager for approval' },
  own_head_Approved: { th: 'ใบเบิกของคุณได้รับการอนุมัติแล้ว', en: 'Your expense claim has been approved' },
  own_head_Rejected: { th: 'ใบเบิกของคุณถูกตีกลับ กรุณาแก้ไขแล้วส่งใหม่ในเลขที่เดิม', en: 'Your expense claim was returned. Please revise it and resubmit under the same claim number.' },
  own_head_Paid: { th: 'ใบเบิกของคุณได้รับการจ่ายเงินแล้ว', en: 'Your expense claim has been paid' },
  own_head_other: { th: 'สถานะใบเบิกเปลี่ยนเป็น ', en: 'Your expense claim status changed to ' },
  own_subj_Checked: { th: 'ผ่านการตรวจสอบแล้ว', en: 'Checked' },
  own_subj_Approved: { th: 'ได้รับการอนุมัติแล้ว', en: 'Approved' },
  own_subj_Rejected: { th: 'ถูกตีกลับให้แก้ไข', en: 'Returned for revision' },
  own_subj_Paid: { th: 'จ่ายเงินแล้ว', en: 'Paid' },
  act_subject: { th: 'เปิดใช้งานบัญชีแล้ว', en: 'Your account is now active' },
  act_greet: { th: 'สวัสดีคุณ ', en: 'Hello ' },
  act_body: { th: 'บัญชีของคุณถูกเปิดใช้งานแล้ว เข้าสู่ระบบเพื่อเริ่มกรอกใบเบิกได้ทันที', en: 'Your account has been activated. Sign in to start submitting expense claims.' },
  rst_subject: { th: 'ตั้งรหัสผ่านใหม่ — ', en: 'Password reset — ' },
  rst_greet: { th: 'สวัสดีครับ', en: 'Hello,' },
  rst_line1: { th: 'มีการขอตั้งรหัสผ่านใหม่สำหรับระบบ ', en: 'A password reset was requested for ' },
  rst_line2: { th: 'คลิกลิงก์ด้านล่างเพื่อตั้งรหัสผ่านใหม่ (ลิงก์หมดอายุใน 30 นาที)', en: 'Click the link below to set a new password. The link expires in 30 minutes.' },
  rst_ignore: { th: 'หากคุณไม่ได้เป็นผู้ร้องขอ กรุณาเพิกเฉยต่ออีเมลฉบับนี้', en: 'If you did not request this, please ignore this email.' }
};

function mailText_(key, lang) {
  var entry = MAIL_TEXT[key];
  if (!entry) return '';
  return lang === 'en' ? entry.en : entry.th;
}

function mailLangForRole_(role) {
  return MAIL_EN_ROLES.indexOf(String(role || '')) >= 0 ? 'en' : 'th';
}

function groupByMailLang_(to) {
  var out = {
    th: [],
    en: []
  };
  var list = (Array.isArray(to) ? to : [ to ]).filter(Boolean);
  if (!list.length) return out;
  var roleByEmail = {};
  try {
    readUsers_().forEach(function(u) {
      roleByEmail[normEmail_(u.Email)] = u.Role || ROLES.SALE;
    });
  } catch (e) {}
  list.forEach(function(addr) {
    out[mailLangForRole_(roleByEmail[normEmail_(addr)])].push(addr);
  });
  return out;
}

function sendMailLocalized_(to, build, attachments) {
  var groups = groupByMailLang_(to);
  [ 'th', 'en' ].forEach(function(lang) {
    if (!groups[lang].length) return;
    var msg = build(lang);
    if (!msg) return;
    sendMail_(groups[lang], msg.subject, msg.html, attachments);
  });
}
