var CFG = {
  APP_TITLE: 'Expense Claim',
  CLAIMS_SHEET: 'Claims',
  ITEMS_SHEET: 'ClaimItems',
  MASTER_SHEET: 'Master',
  USERS_SHEET: 'Users',
  APPROVAL_SHEET: 'ApprovalLog',
  FORM_SHEET: '_FORM_RENDER',
  MAX_ROWS: 25,
  DOC_NO: 'FIN-2026-EXP-001.00',
  DOC_TITLE: 'Expense Claim',
  COMPANY: '',
  CURRENCY: 'THB',
  LOCK_AFTER_SUBMIT: true,
  FINANCE_EMAILS: [],
  AUTO_APPROVE_REGISTRATION: true,
  MIN_PASSWORD_LEN: 8,
  SESSION: {
    IDLE_SEC: 60 * 60 * 2,
    REFRESH_AFTER_SEC: 60 * 5,
    RESET_TTL_SEC: 60 * 30,
    PROP_PREFIX: 'sess_',
    RESET_PREFIX: 'reset_'
  },
  NOTIFY_BY_EMAIL: true,
  ATTACH_PDF_TO_EMAIL: true,
  REQUIRE_PAYMENT_STEP: false
};

var SECRET_KEYS = [ 'SPREADSHEET_ID', 'SUPER_ADMIN_EMAIL', 'SEED_PASSWORD', 'PDF_FOLDER_ID', 'WEBAPP_URL' ];

function secret_(key) {
  var v = '';
  try {
    if (typeof SECRETS !== 'undefined' && SECRETS && SECRETS[key]) v = SECRETS[key];
  } catch (e) {}
  if (!v) {
    try {
      v = PropertiesService.getScriptProperties().getProperty(key) || '';
    } catch (e2) {}
  }
  return String(v || '');
}

function missingSecrets_() {
  return [ 'SPREADSHEET_ID', 'SUPER_ADMIN_EMAIL', 'SEED_PASSWORD' ].filter(function(k) {
    return !secret_(k);
  });
}

var ROLES = {
  SALE: 'Sale',
  ADMIN: 'Admin',
  MANAGER: 'Manager',
  FINANCE: 'Finance',
  SUPER_ADMIN: 'Super Admin'
};

var ROLE_LIST = [ ROLES.SALE, ROLES.ADMIN, ROLES.MANAGER, ROLES.FINANCE, ROLES.SUPER_ADMIN ];

var USERS_SCHEMA = [ 'Email', 'FullName', 'Role', 'PasswordHash', 'Salt', 'Active', 'CreatedAt', 'LastLogin', 'EmployeeCode', 'Division', 'JobTitle', 'DefaultSupportDiv', 'CheckerEmail', 'ApproverEmail', 'Phone', 'UpdatedAt', 'ApprovedBy' ];

var MASTER_COLUMNS = [ 'employees', 'divisions', 'jobTitles', 'supportDivs' ];

var DEFAULT_MASTER = {
  employees: [],
  divisions: [ 'Sales Representative', 'Marketing' ],
  jobTitles: [ 'Sales Representative', 'Marketing' ],
  supportDivs: [ 'BB', 'BP', 'TRC' ]
};

var REGISTER_FIELDS = [ {
  key: 'name',
  en: 'Full Name',
  th: 'ชื่อ-นามสกุล',
  type: 'text',
  required: true
}, {
  key: 'employeeCode',
  en: 'Employee ID',
  th: 'รหัสพนักงาน',
  type: 'text',
  required: false
}, {
  key: 'division',
  en: 'Division / Dept.',
  th: 'ฝ่าย / แผนก',
  type: 'text',
  required: true,
  master: 'divisions'
}, {
  key: 'jobTitle',
  en: 'Job Title',
  th: 'ตำแหน่ง',
  type: 'text',
  required: true,
  master: 'jobTitles'
}, {
  key: 'defaultSupportDiv',
  en: 'Default Support Div.',
  th: 'หน่วยงานที่สังกัด',
  type: 'select',
  required: false,
  master: 'supportDivs'
}, {
  key: 'phone',
  en: 'Phone',
  th: 'เบอร์ติดต่อ',
  type: 'text',
  required: false
} ];

var PROFILE_COLUMN_MAP = {
  name: 'FullName',
  employeeCode: 'EmployeeCode',
  division: 'Division',
  jobTitle: 'JobTitle',
  defaultSupportDiv: 'DefaultSupportDiv',
  checkerEmail: 'CheckerEmail',
  approverEmail: 'ApproverEmail',
  phone: 'Phone'
};

var STAGES = [ {
  key: 'prepare',
  status: 'Draft',
  next: 'Submitted',
  actorRole: ROLES.SALE,
  th: 'จัดทำ',
  en: 'Prepare',
  action: 'ส่งให้ตรวจสอบ',
  actionEn: 'Send for checking',
  signKey: 'prepared'
}, {
  key: 'check',
  status: 'Submitted',
  next: 'Checked',
  actorRole: ROLES.ADMIN,
  th: 'รอตรวจสอบ',
  en: 'Awaiting check',
  action: 'ตรวจสอบผ่าน',
  actionEn: 'Mark as checked',
  signKey: 'received'
}, {
  key: 'approve',
  status: 'Checked',
  next: 'Approved',
  actorRole: ROLES.MANAGER,
  th: 'รออนุมัติ',
  en: 'Awaiting approval',
  action: 'อนุมัติ',
  actionEn: 'Approve',
  signKey: 'approved'
}, {
  key: 'pay',
  status: 'Approved',
  next: 'Paid',
  actorRole: ROLES.FINANCE,
  th: 'รอจ่ายเงิน',
  en: 'Awaiting payment',
  action: 'บันทึกจ่ายเงิน',
  actionEn: 'Record payment',
  signKey: ''
} ];

var EDITABLE_STATUSES = [ 'Draft', 'Rejected' ];

var APPROVAL_COLUMNS = [ 'at', 'claimId', 'action', 'actor', 'role', 'fromStatus', 'toStatus', 'comment' ];

var HEADER_FIELDS = [ {
  key: 'yearMonth',
  en: 'Year / Month',
  th: 'ปี / เดือน',
  type: 'month',
  cell: 'B1',
  required: true
}, {
  key: 'employee',
  en: 'Employee Name',
  th: 'ชื่อพนักงาน',
  type: 'text',
  cell: 'B2',
  required: true,
  master: 'employees'
}, {
  key: 'division',
  en: 'Division / Dept.',
  th: 'ฝ่าย / แผนก',
  type: 'text',
  cell: 'B3',
  required: true,
  master: 'divisions'
}, {
  key: 'jobTitle',
  en: 'Job Title',
  th: 'ตำแหน่ง',
  type: 'text',
  cell: 'B4',
  required: false,
  master: 'jobTitles'
}, {
  key: 'claimDate',
  en: 'Date',
  th: 'วันที่ยื่นเบิก',
  type: 'date',
  cell: 'B5',
  required: true
} ];

var ITEM_FIELDS = [ {
  key: 'date',
  col: 'B',
  en: 'Date',
  th: 'วันที่',
  type: 'date'
}, {
  key: 'receiptNo',
  col: 'C',
  en: 'Receipt No.',
  th: 'เลขที่ใบเสร็จ',
  type: 'text'
}, {
  key: 'description',
  col: 'D',
  en: 'Description',
  th: 'รายละเอียด',
  type: 'text'
}, {
  key: 'supportDiv',
  col: 'E',
  en: 'Support Div.',
  th: 'หน่วยงานที่รับผิดชอบ',
  type: 'select',
  options: [ 'BB', 'BP', 'TRC' ],
  groupKey: 'support'
}, {
  key: 'liter',
  col: 'F',
  en: 'Liter',
  th: 'ลิตร',
  type: 'number',
  decimals: 2
}, {
  key: 'gasoline',
  col: 'G',
  en: 'Gasoline (Fleet Card)',
  th: 'ค่าน้ำมัน (Fleet Card)',
  type: 'money',
  amount: true,
  fleetCard: true,
  groupKey: 'traveling'
}, {
  key: 'carRental',
  col: 'H',
  en: 'Car Rental',
  th: 'ค่าเช่ารถ',
  type: 'money',
  amount: true,
  groupKey: 'traveling'
}, {
  key: 'accommodation',
  col: 'I',
  en: 'Accommodation',
  th: 'ค่าที่พัก',
  type: 'money',
  amount: true,
  groupKey: 'traveling'
}, {
  key: 'airTicket',
  col: 'J',
  en: 'Air Ticket',
  th: 'ค่าตั๋วเครื่องบิน',
  type: 'money',
  amount: true,
  groupKey: 'traveling'
}, {
  key: 'gift',
  col: 'K',
  en: 'Gift',
  th: 'ของขวัญ',
  type: 'money',
  amount: true,
  groupKey: 'entertainment'
}, {
  key: 'foodBeverage',
  col: 'L',
  en: 'Food and Beverage',
  th: 'ค่าอาหารและเครื่องดื่ม',
  type: 'money',
  amount: true,
  groupKey: 'entertainment'
}, {
  key: 'golfCaddy',
  col: 'M',
  en: 'Golf + Caddy',
  th: 'ค่ากอล์ฟ + แคดดี้',
  type: 'money',
  amount: true,
  groupKey: 'entertainment'
}, {
  key: 'expressWay',
  col: 'N',
  en: 'Express Way',
  th: 'ค่าทางด่วน',
  type: 'money',
  amount: true,
  groupKey: 'traffic'
}, {
  key: 'parking',
  col: 'O',
  en: 'Parking',
  th: 'ค่าจอดรถ',
  type: 'money',
  amount: true,
  groupKey: 'traffic'
}, {
  key: 'taxi',
  col: 'P',
  en: 'Taxi',
  th: 'ค่าแท็กซี่',
  type: 'money',
  amount: true,
  groupKey: 'traffic'
}, {
  key: 'btsBus',
  col: 'Q',
  en: 'BTS / Bus',
  th: 'ค่า BTS / รถเมล์',
  type: 'money',
  amount: true,
  groupKey: 'traffic'
}, {
  key: 'postage',
  col: 'R',
  en: 'Postage',
  th: 'ค่าไปรษณีย์',
  type: 'money',
  amount: true,
  groupKey: 'postage'
}, {
  key: 'other',
  col: 'S',
  en: 'Other',
  th: 'อื่น ๆ',
  type: 'money',
  amount: true,
  groupKey: 'other'
} ];

var GROUPS = [ {
  key: 'support',
  en: 'BB / BP / TRC',
  th: 'หน่วยงาน',
  account: '',
  color: '#eef2f7'
}, {
  key: 'traveling',
  en: 'Traveling',
  th: 'ค่าเดินทาง',
  account: '71461000',
  color: '#e8f1fb'
}, {
  key: 'entertainment',
  en: 'Entertainment',
  th: 'ค่ารับรอง',
  account: '71411000',
  color: '#fdeee6'
}, {
  key: 'traffic',
  en: 'Traffic Fare',
  th: 'ค่าพาหนะ',
  account: '71431000',
  color: '#e9f6ec'
}, {
  key: 'postage',
  en: 'Postage',
  th: 'ค่าไปรษณีย์',
  account: '71421000',
  color: '#f3ecfb'
}, {
  key: 'other',
  en: 'Other',
  th: 'อื่น ๆ',
  account: '',
  color: '#f2f2f2'
} ];

var SIGN_BLOCKS = [ {
  key: 'prepared',
  en: 'Prepared',
  th: 'ผู้จัดทำ (Sale)',
  cell: 'F39',
  byField: 'createdBy',
  atField: 'submittedAt'
}, {
  key: 'received',
  en: 'Received',
  th: 'ผู้ตรวจสอบ (Admin)',
  cell: 'K39',
  byField: 'checkedBy',
  atField: 'checkedAt'
}, {
  key: 'approved',
  en: 'Managing Director',
  th: 'ผู้อนุมัติ (Manager)',
  cell: 'P39',
  byField: 'approvedBy',
  atField: 'approvedAt'
} ];

var WORKFLOW = [ 'Summarize monthly expenses, categorized by expense type.', 'Attach all relevant documents, such as original receipts.', 'Submit to your supervisor for review and approval.', 'Submit to the finance department for reimbursement.' ];

var STATUSES = [ 'Draft', 'Submitted', 'Checked', 'Approved', 'Rejected', 'Paid' ];

var STATUS_LABELS = {
  Draft: 'ฉบับร่าง',
  Submitted: 'รอ Admin ตรวจสอบ',
  Checked: 'รอ Manager อนุมัติ',
  Approved: 'อนุมัติแล้ว',
  Rejected: 'ตีกลับให้แก้ไข',
  Paid: 'จ่ายเงินแล้ว'
};

var CLAIM_COLUMNS = [ 'claimId', 'revision', 'status', 'yearMonth', 'employee', 'division', 'jobTitle', 'claimDate', 'itemCount', 'totalLiter', 'grandTotal', 'note', 'createdBy', 'createdAt', 'updatedBy', 'updatedAt', 'pdfUrl', 'submittedAt', 'checkerEmail', 'checkedBy', 'checkedAt', 'checkComment', 'approverEmail', 'approvedBy', 'approvedAt', 'approveComment', 'rejectedBy', 'rejectedAt', 'rejectStage', 'rejectReason', 'paidBy', 'paidAt' ];

var ITEM_COLUMNS = [ 'claimId', 'seq' ].concat(ITEM_FIELDS.map(function(f) {
  return f.key;
})).concat([ 'rowTotal', 'reviewFlag', 'reviewNote' ]);

var REVIEW_FLAGS = [ {
  key: '',
  th: 'ยังไม่ตรวจ',
  en: 'Not checked',
  icon: '○',
  color: '#c3ccda'
}, {
  key: 'ok',
  th: 'ผ่าน',
  en: 'Passed',
  icon: '✓',
  color: '#10794a'
}, {
  key: 'flag',
  th: 'ติดปัญหา',
  en: 'Flagged',
  icon: '⚑',
  color: '#b45309'
} ];

function amountFields_() {
  return ITEM_FIELDS.filter(function(f) {
    return f.amount;
  });
}

function claimableFields_() {
  return ITEM_FIELDS.filter(function(f) {
    return f.amount && !f.fleetCard;
  });
}

function fleetFields_() {
  return ITEM_FIELDS.filter(function(f) {
    return f.amount && f.fleetCard;
  });
}

function fieldByKey_(key) {
  for (var i = 0; i < ITEM_FIELDS.length; i++) {
    if (ITEM_FIELDS[i].key === key) return ITEM_FIELDS[i];
  }
  return null;
}