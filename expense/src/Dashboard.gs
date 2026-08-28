var DASH_STATUSES = [ 'Submitted', 'Checked', 'Approved', 'Paid' ];

function canSeeAllSpending_(me) {
  return !!(me.isSuperAdmin || me.isFinance || me.isChecker || me.isApprover);
}

function dashPerson_(claim) {
  return {
    email: normEmail_(claim.createdBy),
    name: String(claim.employee || claim.createdBy || '')
  };
}

function getDashboard(token, filter) {
  var gate = requireActiveAccount_(token);
  if (gate) return gate;
  var me = me_();
  var seesAll = canSeeAllSpending_(me);
  filter = filter || {};

  var wantPerson = seesAll ? normEmail_(filter.person) : me.email;
  var wantMonth = String(filter.yearMonth || '');
  var wantStatus = String(filter.status || '');

  var empty = {
    ok: true,
    seesAll: seesAll,
    people: [],
    months: [],
    totals: { claim: 0, fuel: 0, spend: 0, liter: 0, claims: 0, items: 0 },
    byType: [],
    byGroup: [],
    byPerson: []
  };

  var claimsSh = sheet_(CFG.CLAIMS_SHEET, CLAIM_COLUMNS);
  if (claimsSh.getLastRow() < 2) return safeOut_(empty);

  var people = {};
  var months = {};
  var picked = {};
  var claimCount = 0;

  claimsSh.getRange(2, 1, claimsSh.getLastRow() - 1, CLAIM_COLUMNS.length).getValues().forEach(function(r) {
    var c = rowToObject_(CLAIM_COLUMNS, r);
    if (!c.claimId) return;
    if (DASH_STATUSES.indexOf(c.status) < 0) return;
    if (!seesAll && normEmail_(c.createdBy) !== me.email) return;

    var who = dashPerson_(c);
    var ym = fmtMonth_(c.yearMonth);
    people[who.email] = who.name;
    if (ym) months[ym] = true;

    if (wantPerson && who.email !== wantPerson) return;
    if (wantMonth && ym !== wantMonth) return;
    if (wantStatus && c.status !== wantStatus) return;

    picked[String(c.claimId)] = who;
    claimCount++;
  });

  var claimable = claimableFields_();
  var fleet = fleetFields_();
  var byType = {};
  claimable.forEach(function(f) { byType[f.key] = 0; });
  var byPerson = {};
  var totalClaim = 0, totalFuel = 0, totalLiter = 0, itemCount = 0;

  var itemsSh = sheet_(CFG.ITEMS_SHEET, ITEM_COLUMNS);
  if (itemsSh.getLastRow() > 1) {
    itemsSh.getRange(2, 1, itemsSh.getLastRow() - 1, ITEM_COLUMNS.length).getValues().forEach(function(r) {
      var it = rowToObject_(ITEM_COLUMNS, r);
      var who = picked[String(it.claimId)];
      if (!who) return;

      var rowClaim = 0;
      claimable.forEach(function(f) {
        var v = num_(it[f.key]);
        if (!v) return;
        byType[f.key] += v;
        rowClaim += v;
      });
      var rowFuel = fleet.reduce(function(s, f) { return s + num_(it[f.key]); }, 0);
      var rowLiter = num_(it.liter);

      totalClaim += rowClaim;
      totalFuel += rowFuel;
      totalLiter += rowLiter;
      itemCount++;

      var p = byPerson[who.email];
      if (!p) p = byPerson[who.email] = { email: who.email, name: who.name, claim: 0, fuel: 0, liter: 0, claims: 0 };
      p.claim += rowClaim;
      p.fuel += rowFuel;
      p.liter += rowLiter;
    });
  }

  Object.keys(picked).forEach(function(cid) {
    var p = byPerson[picked[cid].email];
    if (p) p.claims++;
  });

  var groupTotals = {};
  var typeRows = claimable.map(function(f) {
    var amount = round2_(byType[f.key]);
    groupTotals[f.groupKey] = round2_((groupTotals[f.groupKey] || 0) + amount);
    return { key: f.key, th: f.th, en: f.en, groupKey: f.groupKey, amount: amount };
  }).filter(function(t) { return t.amount !== 0; }).sort(function(a, b) { return b.amount - a.amount; });

  var groupRows = GROUPS.map(function(g) {
    return { key: g.key, th: g.th, en: g.en, account: g.account, color: g.color, amount: round2_(groupTotals[g.key] || 0) };
  }).filter(function(g) { return g.amount !== 0; }).sort(function(a, b) { return b.amount - a.amount; });

  var personRows = Object.keys(byPerson).map(function(k) {
    var p = byPerson[k];
    p.claim = round2_(p.claim);
    p.fuel = round2_(p.fuel);
    p.liter = round2_(p.liter);
    p.spend = round2_(p.claim + p.fuel);
    return p;
  }).sort(function(a, b) { return b.spend - a.spend; });

  return safeOut_({
    ok: true,
    seesAll: seesAll,
    people: Object.keys(people).map(function(e) { return { email: e, name: people[e] }; }).sort(function(a, b) { return a.name.localeCompare(b.name); }),
    months: Object.keys(months).sort().reverse(),
    totals: {
      claim: round2_(totalClaim),
      fuel: round2_(totalFuel),
      spend: round2_(totalClaim + totalFuel),
      liter: round2_(totalLiter),
      claims: claimCount,
      items: itemCount
    },
    byType: typeRows,
    byGroup: groupRows,
    byPerson: personRows
  });
}
