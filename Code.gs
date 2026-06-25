// ═══════════════════════════════════════════════════════════════
// ENTERPRISE BUSINESS CONTROL TOWER — Apps Script Backend
// Integrates Marketplace + Open Marketplace Sheets
// ═══════════════════════════════════════════════════════════════

var CONFIG = {
  MARKETPLACE_SHEET_ID: '1AvZZtujyTbV_fUQ-azIeOnm5xxrujWum',
  OMP_SHEET_ID:         '10v3bhgKY1C2dadT2nchXD0uZQ3GImxvK',
  CACHE_TTL:      240,   // seconds (4 min — just inside 5-min cache window)
  SLA_OK_MAX:     2,     // days — within SLA
  SLA_WARN_MAX:   5,     // days — warning zone
  LONG_PENDING_DAYS: 30, // days pending before flagged as exception
  MAX_TABLE_ROWS: 300,   // per source
};

// ─────────────────────────────────────────────────────────────
// ENTRY POINTS
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Enterprise Business Control Tower')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Called from the frontend via google.script.run
function getDashboardData(filtersJson) {
  try {
    var filters = filtersJson ? JSON.parse(filtersJson) : {};
    var cacheKey = 'dash_' + JSON.stringify(filters);
    var cache    = CacheService.getScriptCache();
    var cached   = cache.get(cacheKey);
    if (cached) return cached;

    var mpRaw   = readSheet(CONFIG.MARKETPLACE_SHEET_ID);
    var ompRaw  = readSheet(CONFIG.OMP_SHEET_ID);

    var mpAll   = normalizeMarketplace(mpRaw);
    var ompAll  = normalizeOMP(ompRaw);

    var mpData  = filterMarketplace(mpAll, filters);
    var ompData = filterOMP(ompAll, filters);

    var result = JSON.stringify({
      success:       true,
      lastUpdated:   new Date().toISOString(),
      filters:       filters,
      enterprise:    calcEnterpriseKPIs(mpData, ompData),
      marketplace:   calcMarketplaceKPIs(mpData),
      omp:           calcOMPKPIs(ompData),
      trends:        calcTrends(mpData),
      tat:           calcTAT(mpData),
      sla:           calcSLA(mpData),
      exceptions:    detectExceptions(mpData, ompData),
      table:         buildTable(mpData, ompData),
      filterOptions: buildFilterOptions(mpAll, ompAll),
    });

    try { cache.put(cacheKey, result, CONFIG.CACHE_TTL); } catch(e) {}
    return result;

  } catch (err) {
    return JSON.stringify({ success: false, error: err.message, stack: err.stack });
  }
}

// ─────────────────────────────────────────────────────────────
// SHEET READING
// ─────────────────────────────────────────────────────────────

function readSheet(id) {
  var ss    = SpreadsheetApp.openById(id);
  var sheet = ss.getSheets()[0];
  var vals  = sheet.getDataRange().getValues();
  if (!vals || vals.length < 2) return { headers: [], rows: [] };

  var headers = vals[0].map(function(h) {
    return String(h).trim().toLowerCase()
      .replace(/\s+/g, '_')
      .replace(/[^a-z0-9_]/g, '');
  });

  var rows = vals.slice(1).filter(function(row) {
    return row.some(function(c) { return c !== '' && c !== null && c !== undefined; });
  });

  return { headers: headers, rows: rows };
}

// ─────────────────────────────────────────────────────────────
// NORMALIZATION — MARKETPLACE
// Columns: id, business_name, state, vendor_type, business_category,
//          gstin_number, status, current_status, status_changed_on,
//          created_date, review_submission_date, party_id,
//          level1_rejected1..level4, onboarded_date,
//          Firstshipmentdate, LifetimeValue
// ─────────────────────────────────────────────────────────────

function normalizeMarketplace(raw) {
  var idx = buildIndex(raw.headers);
  var now = new Date();

  return raw.rows.map(function(row) {
    var created   = parseDate(gv(row, idx, 'created_date'));
    var onboarded = parseDate(gv(row, idx, 'onboarded_date'));
    var shipment  = parseDate(gv(row, idx, 'firstshipmentdate'));
    var status    = normStatus(gv(row, idx, 'status'));
    var curStatus = normCurrentStatus(gv(row, idx, 'current_status'));
    var gstin     = String(gv(row, idx, 'gstin_number') || '').trim();
    var ltv       = parseNumber(gv(row, idx, 'lifetimevalue'));
    var onbTAT    = dateDiffDays(created, onboarded);
    var shipTAT   = dateDiffDays(onboarded, shipment);
    var age       = dateDiffDays(created, now);

    return {
      id:           String(gv(row, idx, 'id') || '').replace(/,/g, '').trim(),
      name:         String(gv(row, idx, 'business_name') || '').trim(),
      state:        String(gv(row, idx, 'state') || '').trim(),
      vendorType:   String(gv(row, idx, 'vendor_type') || '').trim(),
      category:     normCategory(gv(row, idx, 'business_category')),
      gstin:        gstin,
      status:       status,
      currentStatus: curStatus,
      statusChangedOn: parseDate(gv(row, idx, 'status_changed_on')),
      createdDate:  created,
      reviewDate:   parseDate(gv(row, idx, 'review_submission_date')),
      onboardedDate: onboarded,
      shipmentDate: shipment,
      ltv:          ltv,
      onbTAT:       onbTAT,
      shipTAT:      shipTAT,
      age:          age,
      hasGST:       gstin.length > 5,
      hasShipment:  !!shipment,
      source:       'Marketplace',
    };
  }).filter(function(r) { return r.id || r.name; });
}

// ─────────────────────────────────────────────────────────────
// NORMALIZATION — OPEN MARKETPLACE
// Columns: id, gstin, business_name, status, vendor_type,
//          business_category, additional_documents,
//          proof_of_premises, electricity_bill, kyc_document,
//          gst_portal_screenshot_bank_details, cancelled_cheque,
//          msme_certificate, aadhaar, owner_pan, entity_pan,
//          gst_certificate
// ─────────────────────────────────────────────────────────────

var OMP_DOC_FIELDS = [
  { key: 'additional_documents',                  label: 'Additional Docs' },
  { key: 'proof_of_premises',                     label: 'Proof of Premises' },
  { key: 'electricity_bill',                      label: 'Electricity Bill' },
  { key: 'kyc_document',                          label: 'KYC Document' },
  { key: 'gst_portal_screenshot_bank_details',    label: 'GST Portal / Bank' },
  { key: 'cancelled_cheque',                      label: 'Cancelled Cheque' },
  { key: 'msme_certificate',                      label: 'MSME Certificate' },
  { key: 'aadhaar',                               label: 'Aadhaar' },
  { key: 'owner_pan',                             label: 'Owner PAN' },
  { key: 'entity_pan',                            label: 'Entity PAN' },
  { key: 'gst_certificate',                       label: 'GST Certificate' },
];

function normalizeOMP(raw) {
  var idx = buildIndex(raw.headers);

  return raw.rows.map(function(row) {
    var gstin  = String(gv(row, idx, 'gstin') || '').trim();
    var status = normStatus(gv(row, idx, 'status'));
    var docs   = {};
    var filled = 0;

    OMP_DOC_FIELDS.forEach(function(f) {
      var v = parseInt(gv(row, idx, f.key) || 0, 10) || 0;
      docs[f.key] = v;
      if (v > 0) filled++;
    });

    var docPct = Math.round((filled / OMP_DOC_FIELDS.length) * 100);
    var missingDocs = OMP_DOC_FIELDS.filter(function(f) { return (docs[f.key] || 0) === 0; })
                                     .map(function(f) { return f.label; });

    return {
      id:           String(gv(row, idx, 'id') || '').replace(/,/g, '').trim(),
      gstin:        gstin,
      name:         String(gv(row, idx, 'business_name') || '').trim(),
      status:       status,
      vendorType:   String(gv(row, idx, 'vendor_type') || '').trim(),
      category:     normCategory(gv(row, idx, 'business_category')),
      docs:         docs,
      docsFilled:   filled,
      docsTotal:    OMP_DOC_FIELDS.length,
      docPct:       docPct,
      missingDocs:  missingDocs,
      hasGST:       gstin.length > 5,
      source:       'Open Marketplace',
    };
  }).filter(function(r) { return r.id || r.name; });
}

// ─────────────────────────────────────────────────────────────
// FILTERS
// ─────────────────────────────────────────────────────────────

function filterMarketplace(data, f) {
  return data.filter(function(r) {
    if (f.vertical === 'OMP') return false;
    if (f.category && f.category !== 'All' && r.category !== f.category) return false;
    if (f.vendorType && f.vendorType !== 'All' && r.vendorType !== f.vendorType) return false;
    if (f.status && f.status !== 'All' && r.status !== f.status) return false;
    if (f.state && f.state !== 'All' && r.state !== f.state) return false;
    if (f.search) {
      var q = f.search.toLowerCase();
      if (!(r.name || '').toLowerCase().includes(q) &&
          !(r.gstin || '').toLowerCase().includes(q) &&
          !(r.id || '').includes(q)) return false;
    }
    if (f.period && f.period !== 'All') {
      var d = r.createdDate;
      if (!d) return false;
      var now = new Date();
      if (f.period === 'Today') {
        var s = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        if (d < s) return false;
      } else if (f.period === 'MTD') {
        var s = new Date(now.getFullYear(), now.getMonth(), 1);
        if (d < s) return false;
      } else if (f.period === 'YTD') {
        var s = new Date(now.getFullYear(), 0, 1);
        if (d < s) return false;
      } else if (f.period === 'Custom' && f.startDate && f.endDate) {
        var s = new Date(f.startDate), e = new Date(f.endDate);
        if (d < s || d > e) return false;
      }
    }
    return true;
  });
}

function filterOMP(data, f) {
  return data.filter(function(r) {
    if (f.vertical === 'Marketplace') return false;
    if (f.category && f.category !== 'All' && r.category !== f.category) return false;
    if (f.vendorType && f.vendorType !== 'All' && r.vendorType !== f.vendorType) return false;
    if (f.status && f.status !== 'All' && r.status !== f.status) return false;
    if (f.search) {
      var q = f.search.toLowerCase();
      if (!(r.name || '').toLowerCase().includes(q) &&
          !(r.gstin || '').toLowerCase().includes(q) &&
          !(r.id || '').includes(q)) return false;
    }
    return true;
  });
}

// ─────────────────────────────────────────────────────────────
// KPI CALCULATIONS
// ─────────────────────────────────────────────────────────────

function calcEnterpriseKPIs(mp, omp) {
  var totalCases     = mp.length + omp.length;
  var totalCompleted = count(mp, 'status', 'COMPLETED') + count(omp, 'status', 'COMPLETED');
  var totalPending   = countFn(mp, function(r) { return r.status === 'DRAFT' || r.status === 'IN_REVIEW'; })
                     + countFn(omp, function(r) { return r.status === 'DRAFT' || r.status === 'IN_REVIEW'; });
  var totalRejected  = count(mp, 'status', 'REJECTED') + count(omp, 'status', 'REJECTED');
  var totalLTV       = mp.reduce(function(s, r) { return s + (r.ltv || 0); }, 0);
  var sla            = calcSLA(mp);
  var tats           = mp.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; }).map(function(r) { return r.onbTAT; });
  var avgTAT         = tats.length ? Math.round(avg(tats)) : 0;
  var exceptions     = countExceptions(mp, omp);

  return {
    totalCases:     totalCases,
    totalCompleted: totalCompleted,
    totalPending:   totalPending,
    totalRejected:  totalRejected,
    totalLTV:       totalLTV,
    avgTAT:         avgTAT,
    completionPct:  pct(totalCompleted, totalCases),
    slaOKPct:       sla.okPct,
    slaBreachPct:   sla.breachPct,
    exceptions:     exceptions,
    mpTotal:        mp.length,
    ompTotal:       omp.length,
    mpPct:          pct(mp.length, totalCases),
    ompPct:         pct(omp.length, totalCases),
    mpCompleted:    count(mp, 'status', 'COMPLETED'),
    ompCompleted:   count(omp, 'status', 'COMPLETED'),
  };
}

function calcMarketplaceKPIs(data) {
  var total      = data.length;
  var completed  = count(data, 'status', 'COMPLETED');
  var draft      = count(data, 'status', 'DRAFT');
  var inReview   = count(data, 'status', 'IN_REVIEW');
  var rejected   = count(data, 'status', 'REJECTED');
  var active     = countFn(data, function(r) { return r.currentStatus === 'ACTIVE' || r.currentStatus === 'REACTIVATED'; });
  var churned    = count(data, 'currentStatus', 'CHURNED');
  var deact      = count(data, 'currentStatus', 'DEACTIVATED');
  var react      = count(data, 'currentStatus', 'REACTIVATED');
  var withShip   = countFn(data, function(r) { return r.hasShipment; });
  var noShip     = completed - withShip;
  var missingGST = countFn(data, function(r) { return !r.hasGST; });

  var ltv        = data.reduce(function(s, r) { return s + (r.ltv || 0); }, 0);
  var avgLTV     = total ? ltv / total : 0;

  var tats       = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; }).map(function(r) { return r.onbTAT; });
  var avgTAT     = tats.length ? Math.round(avg(tats)) : 0;

  var now       = new Date();
  var today0    = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var mtd0      = new Date(now.getFullYear(), now.getMonth(), 1);
  var ytd0      = new Date(now.getFullYear(), 0, 1);

  return {
    total:        total,
    completed:    completed,
    draft:        draft,
    inReview:     inReview,
    rejected:     rejected,
    pending:      draft + inReview,
    active:       active,
    churned:      churned,
    deactivated:  deact,
    reactivated:  react,
    withShipment: withShip,
    noShipment:   noShip < 0 ? 0 : noShip,
    missingGST:   missingGST,
    totalLTV:     ltv,
    avgLTV:       avgLTV,
    avgTAT:       avgTAT,
    completionPct: pct(completed, total),
    shipmentRate:  pct(withShip, completed),
    activeRate:    pct(active, completed),
    // Date KPIs
    createdToday: countFn(data, function(r) { return r.createdDate && r.createdDate >= today0; }),
    createdMTD:   countFn(data, function(r) { return r.createdDate && r.createdDate >= mtd0; }),
    createdYTD:   countFn(data, function(r) { return r.createdDate && r.createdDate >= ytd0; }),
    onbToday:     countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= today0; }),
    onbMTD:       countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= mtd0; }),
    onbYTD:       countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= ytd0; }),
    shipToday:    countFn(data, function(r) { return r.shipmentDate && r.shipmentDate >= today0; }),
    shipMTD:      countFn(data, function(r) { return r.shipmentDate && r.shipmentDate >= mtd0; }),
    shipYTD:      countFn(data, function(r) { return r.shipmentDate && r.shipmentDate >= ytd0; }),
    // Splits
    statusSplit:     objToArr(groupBy(data, 'status')),
    curStatusSplit:  objToArr(groupBy(data, 'currentStatus')),
    categorySplit:   objToArr(groupBy(data, 'category')),
    vendorSplit:     objToArr(groupBy(data, 'vendorType')),
    stateSplit:      objToArr(groupBy(data, 'state')).slice(0, 12),
  };
}

function calcOMPKPIs(data) {
  var total     = data.length;
  var completed = count(data, 'status', 'COMPLETED');
  var draft     = count(data, 'status', 'DRAFT');
  var inReview  = count(data, 'status', 'IN_REVIEW');
  var rejected  = count(data, 'status', 'REJECTED');
  var withGST   = countFn(data, function(r) { return r.hasGST; });

  var avgDoc = total ? Math.round(data.reduce(function(s, r) { return s + r.docPct; }, 0) / total) : 0;

  var docStats = OMP_DOC_FIELDS.map(function(f) {
    var submitted = countFn(data, function(r) { return (r.docs[f.key] || 0) > 0; });
    return {
      key:       f.key,
      label:     f.label,
      submitted: submitted,
      missing:   total - submitted,
      pct:       pct(submitted, total),
    };
  });

  return {
    total:         total,
    completed:     completed,
    draft:         draft,
    inReview:      inReview,
    rejected:      rejected,
    pending:       draft + inReview,
    withGST:       withGST,
    missingGST:    total - withGST,
    avgDocPct:     avgDoc,
    pipelinePct:   pct(completed, total),
    completionPct: pct(completed, total),
    statusSplit:   objToArr(groupBy(data, 'status')),
    categorySplit: objToArr(groupBy(data, 'category')),
    vendorSplit:   objToArr(groupBy(data, 'vendorType')),
    docStats:      docStats,
  };
}

// ─────────────────────────────────────────────────────────────
// TRENDS
// ─────────────────────────────────────────────────────────────

function calcTrends(data) {
  var now    = new Date();
  var months = [];

  for (var i = 11; i >= 0; i--) {
    var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push({
      label: d.toLocaleString('default', { month: 'short' }) + ' ' + String(d.getFullYear()).slice(2),
      year:  d.getFullYear(),
      month: d.getMonth(),
    });
  }

  var monthly = months.map(function(m) {
    return {
      label:     m.label,
      created:   countFn(data, function(r) { return r.createdDate && r.createdDate.getFullYear() === m.year && r.createdDate.getMonth() === m.month; }),
      onboarded: countFn(data, function(r) { return r.onboardedDate && r.onboardedDate.getFullYear() === m.year && r.onboardedDate.getMonth() === m.month; }),
      shipped:   countFn(data, function(r) { return r.shipmentDate && r.shipmentDate.getFullYear() === m.year && r.shipmentDate.getMonth() === m.month; }),
    };
  });

  var years  = [now.getFullYear() - 2, now.getFullYear() - 1, now.getFullYear()];
  var yearly = years.map(function(y) {
    return {
      label:     String(y),
      created:   countFn(data, function(r) { return r.createdDate && r.createdDate.getFullYear() === y; }),
      onboarded: countFn(data, function(r) { return r.onboardedDate && r.onboardedDate.getFullYear() === y; }),
    };
  });

  return { monthly: monthly, yearly: yearly };
}

// ─────────────────────────────────────────────────────────────
// TAT
// ─────────────────────────────────────────────────────────────

function calcTAT(data) {
  var withTAT = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; });
  if (!withTAT.length) return { avg: 0, min: 0, max: 0, median: 0, distribution: [], categoryTAT: [], vendorTAT: [] };

  var vals = withTAT.map(function(r) { return r.onbTAT; }).sort(function(a, b) { return a - b; });
  var mean = avg(vals);
  var med  = vals[Math.floor(vals.length / 2)];

  var buckets = [
    { label: '0–2 days',  min: 0,  max: 2 },
    { label: '3–5 days',  min: 3,  max: 5 },
    { label: '6–10 days', min: 6,  max: 10 },
    { label: '11–20 days',min: 11, max: 20 },
    { label: '21–30 days',min: 21, max: 30 },
    { label: '31+ days',  min: 31, max: Infinity },
  ].map(function(b) {
    b.count = countFn(withTAT, function(r) { return r.onbTAT >= b.min && r.onbTAT <= b.max; });
    return b;
  });

  var catMap = {};
  withTAT.forEach(function(r) {
    var k = r.category || 'Unknown';
    if (!catMap[k]) catMap[k] = [];
    catMap[k].push(r.onbTAT);
  });
  var categoryTAT = Object.keys(catMap).map(function(k) {
    return { category: k, avg: Math.round(avg(catMap[k])), count: catMap[k].length };
  }).sort(function(a, b) { return b.avg - a.avg; });

  var vMap = {};
  withTAT.forEach(function(r) {
    var k = r.vendorType || 'Unknown';
    if (!vMap[k]) vMap[k] = [];
    vMap[k].push(r.onbTAT);
  });
  var vendorTAT = Object.keys(vMap).map(function(k) {
    return { vendorType: k, avg: Math.round(avg(vMap[k])), count: vMap[k].length };
  }).sort(function(a, b) { return b.avg - a.avg; }).slice(0, 10);

  return {
    avg:          Math.round(mean),
    min:          vals[0],
    max:          vals[vals.length - 1],
    median:       Math.round(med),
    totalWithTAT: withTAT.length,
    distribution: buckets,
    categoryTAT:  categoryTAT,
    vendorTAT:    vendorTAT,
  };
}

// ─────────────────────────────────────────────────────────────
// SLA
// ─────────────────────────────────────────────────────────────

function calcSLA(data) {
  var withTAT = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; });
  var total   = withTAT.length;
  if (!total) return { ok: 0, warn: 0, breach: 0, okPct: 0, warnPct: 0, breachPct: 0, avgDelay: 0 };

  var ok     = countFn(withTAT, function(r) { return r.onbTAT <= CONFIG.SLA_OK_MAX; });
  var warn   = countFn(withTAT, function(r) { return r.onbTAT > CONFIG.SLA_OK_MAX && r.onbTAT <= CONFIG.SLA_WARN_MAX; });
  var breach = countFn(withTAT, function(r) { return r.onbTAT > CONFIG.SLA_WARN_MAX; });

  var breachVals  = withTAT.filter(function(r) { return r.onbTAT > CONFIG.SLA_WARN_MAX; }).map(function(r) { return r.onbTAT; });
  var avgDelay    = breachVals.length ? Math.round(avg(breachVals) - CONFIG.SLA_WARN_MAX) : 0;

  return {
    ok:       ok,
    warn:     warn,
    breach:   breach,
    okPct:    pct(ok, total),
    warnPct:  pct(warn, total),
    breachPct:pct(breach, total),
    avgDelay: avgDelay,
    total:    total,
  };
}

// ─────────────────────────────────────────────────────────────
// EXCEPTION DETECTION
// ─────────────────────────────────────────────────────────────

function countExceptions(mp, omp) {
  var n = 0;
  n += mp.filter(function(r) { return !r.hasGST; }).length;
  n += mp.filter(function(r) { return r.status === 'COMPLETED' && !r.hasShipment; }).length;
  n += mp.filter(function(r) { return (r.status === 'DRAFT' || r.status === 'IN_REVIEW') && r.age > CONFIG.LONG_PENDING_DAYS; }).length;
  n += omp.filter(function(r) { return !r.hasGST; }).length;
  n += omp.filter(function(r) { return r.status === 'DRAFT' && r.docPct < 30; }).length;
  return n;
}

function detectExceptions(mp, omp) {
  var mpMissingGST = mp.filter(function(r) { return !r.hasGST; }).slice(0, 50).map(exRow);
  var ompMissingGST = omp.filter(function(r) { return !r.hasGST; }).slice(0, 50).map(exRowOMP);

  var missingShipment = mp.filter(function(r) { return r.status === 'COMPLETED' && !r.hasShipment; })
    .slice(0, 50).map(function(r) {
      return { id: r.id, name: r.name, category: r.category, vendorType: r.vendorType, onboardedDate: fmtDate(r.onboardedDate), source: 'Marketplace' };
    });

  var longPendingMP = mp.filter(function(r) {
    return (r.status === 'DRAFT' || r.status === 'IN_REVIEW') && r.age > CONFIG.LONG_PENDING_DAYS;
  }).slice(0, 50).map(function(r) {
    return { id: r.id, name: r.name, category: r.category, status: r.status, age: r.age, source: 'Marketplace' };
  });

  var longPendingOMP = omp.filter(function(r) { return r.status === 'DRAFT' && r.docPct < 30; })
    .slice(0, 50).map(function(r) {
      return { id: r.id, name: r.name, category: r.category, docPct: r.docPct, source: 'Open Marketplace' };
    });

  var highTAT = mp.filter(function(r) { return r.onbTAT > CONFIG.SLA_WARN_MAX; })
    .slice(0, 50).map(function(r) {
      return { id: r.id, name: r.name, tat: r.onbTAT, category: r.category, source: 'Marketplace' };
    }).sort(function(a, b) { return b.tat - a.tat; });

  var rejectedMP  = mp.filter(function(r) { return r.status === 'REJECTED'; }).slice(0, 30).map(exRow);
  var rejectedOMP = omp.filter(function(r) { return r.status === 'REJECTED'; }).slice(0, 30).map(exRowOMP);

  // Duplicate GST in MP
  var gstMap = {};
  mp.forEach(function(r) { if (r.gstin) gstMap[r.gstin] = (gstMap[r.gstin] || 0) + 1; });
  var dupGST = mp.filter(function(r) { return r.gstin && gstMap[r.gstin] > 1; })
    .slice(0, 30).map(function(r) {
      return { id: r.id, name: r.name, gstin: r.gstin, count: gstMap[r.gstin] };
    });

  // Missing key docs in OMP
  var missingKeyDocs = omp.filter(function(r) {
    return ['kyc_document', 'gst_certificate', 'aadhaar', 'owner_pan'].some(function(d) { return (r.docs[d] || 0) === 0; });
  }).slice(0, 50).map(function(r) {
    return { id: r.id, name: r.name, missingDocs: r.missingDocs.slice(0, 5), docPct: r.docPct };
  });

  return {
    missingGST:     { mp: mpMissingGST, omp: ompMissingGST, total: mpMissingGST.length + ompMissingGST.length },
    missingShipment:{ items: missingShipment, total: missingShipment.length },
    longPending:    { mp: longPendingMP, omp: longPendingOMP, total: longPendingMP.length + longPendingOMP.length },
    highTAT:        { items: highTAT, total: highTAT.length },
    rejected:       { mp: rejectedMP, omp: rejectedOMP, total: rejectedMP.length + rejectedOMP.length },
    duplicateGST:   { items: dupGST, total: dupGST.length },
    missingDocs:    { items: missingKeyDocs, total: missingKeyDocs.length },
  };
}

function exRow(r) {
  return { id: r.id, name: r.name, category: r.category, vendorType: r.vendorType, status: r.status, source: 'Marketplace' };
}
function exRowOMP(r) {
  return { id: r.id, name: r.name, category: r.category, vendorType: r.vendorType, status: r.status, source: 'Open Marketplace' };
}

// ─────────────────────────────────────────────────────────────
// TABLE DATA
// ─────────────────────────────────────────────────────────────

function buildTable(mp, omp) {
  var mpRows = mp.slice(0, CONFIG.MAX_TABLE_ROWS).map(function(r) {
    return {
      id:          r.id,
      name:        r.name,
      source:      'Marketplace',
      category:    r.category,
      vendorType:  r.vendorType,
      status:      r.status,
      curStatus:   r.currentStatus,
      gstin:       r.gstin || '—',
      state:       r.state || '—',
      createdDate: fmtDate(r.createdDate),
      onbDate:     fmtDate(r.onboardedDate),
      shipDate:    fmtDate(r.shipmentDate),
      tat:         r.onbTAT !== null ? r.onbTAT : '—',
      ltv:         r.ltv > 0 ? fmtCurrency(r.ltv) : '—',
      hasGST:      r.hasGST,
      hasShip:     r.hasShipment,
      docPct:      '—',
    };
  });

  var ompRows = omp.slice(0, CONFIG.MAX_TABLE_ROWS).map(function(r) {
    return {
      id:          r.id,
      name:        r.name,
      source:      'Open Marketplace',
      category:    r.category,
      vendorType:  r.vendorType,
      status:      r.status,
      curStatus:   '—',
      gstin:       r.gstin || '—',
      state:       '—',
      createdDate: '—',
      onbDate:     '—',
      shipDate:    '—',
      tat:         '—',
      ltv:         '—',
      hasGST:      r.hasGST,
      hasShip:     false,
      docPct:      r.docPct + '%',
    };
  });

  return {
    rows:     mpRows.concat(ompRows),
    mpCount:  mpRows.length,
    ompCount: ompRows.length,
    total:    mpRows.length + ompRows.length,
  };
}

// ─────────────────────────────────────────────────────────────
// FILTER OPTIONS
// ─────────────────────────────────────────────────────────────

function buildFilterOptions(mp, omp) {
  var all = mp.concat(omp);
  return {
    categories: unique(all.map(function(r) { return r.category; })).filter(Boolean).sort(),
    vendorTypes: unique(all.map(function(r) { return r.vendorType; })).filter(Boolean).sort(),
    states:     unique(mp.map(function(r) { return r.state; })).filter(Boolean).sort(),
    statuses:   ['COMPLETED', 'IN_REVIEW', 'DRAFT', 'REJECTED'],
  };
}

// ─────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ─────────────────────────────────────────────────────────────

function buildIndex(headers) {
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  return idx;
}

function gv(row, idx, key) {
  return (idx[key] !== undefined) ? row[idx[key]] : '';
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  var s = String(val).trim();
  if (!s) return null;
  // Google Sheets serial numbers come as numbers
  if (!isNaN(Number(s)) && Number(s) > 1000) {
    var n = Number(s);
    var d = new Date((n - 25569) * 86400000); // Excel epoch
    if (!isNaN(d.getTime())) return d;
  }
  // "Month Day, Year, H:MM am" format
  var d = new Date(s.replace(/,\s*\d+:\d+\s*(am|pm)/i, '').trim());
  if (!isNaN(d.getTime())) return d;
  var d2 = new Date(s);
  if (!isNaN(d2.getTime())) return d2;
  return null;
}

function parseNumber(val) {
  if (!val && val !== 0) return 0;
  if (typeof val === 'number') return val;
  return parseFloat(String(val).replace(/,/g, '').replace(/[^0-9.\-]/g, '')) || 0;
}

function dateDiffDays(d1, d2) {
  if (!d1 || !d2) return null;
  return Math.round((d2.getTime() - d1.getTime()) / 86400000);
}

function normStatus(v) {
  var s = String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  var map = { 'COMPLETE': 'COMPLETED', 'INREVIEW': 'IN_REVIEW', 'REVIEW': 'IN_REVIEW', 'IN_REVIEW': 'IN_REVIEW', 'ACTIVE': 'ACTIVE' };
  return map[s] || s || 'UNKNOWN';
}

function normCurrentStatus(v) {
  var s = String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  var map = { 'CHURN': 'CHURNED', 'INACTIVE': 'DEACTIVATED', 'DEACTIVATE': 'DEACTIVATED' };
  return map[s] || s || 'UNKNOWN';
}

function normCategory(v) {
  var s = String(v || '').trim();
  if (!s) return 'Others';
  return s;
}

function count(arr, field, val) {
  return arr.filter(function(r) { return r[field] === val; }).length;
}
function countFn(arr, fn) {
  return arr.filter(fn).length;
}
function avg(arr) {
  return arr.length ? arr.reduce(function(a, b) { return a + b; }, 0) / arr.length : 0;
}
function pct(n, d) {
  return d ? Math.round((n / d) * 100) : 0;
}
function groupBy(arr, field) {
  var res = {};
  arr.forEach(function(r) { var k = r[field] || 'Unknown'; res[k] = (res[k] || 0) + 1; });
  return res;
}
function objToArr(obj) {
  return Object.keys(obj).map(function(k) { return { label: k, count: obj[k] }; })
               .sort(function(a, b) { return b.count - a.count; });
}
function unique(arr) {
  var seen = {}; return arr.filter(function(v) { if (seen[v]) return false; seen[v] = true; return true; });
}
function fmtDate(d) {
  if (!d) return '—';
  return d.getDate() + '/' + (d.getMonth() + 1) + '/' + d.getFullYear();
}
function fmtCurrency(v) {
  if (!v) return '₹0';
  if (v >= 10000000) return '₹' + (v / 10000000).toFixed(2) + ' Cr';
  if (v >= 100000)   return '₹' + (v / 100000).toFixed(2) + ' L';
  return '₹' + Math.round(v).toLocaleString('en-IN');
}
