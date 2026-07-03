// ═══════════════════════════════════════════════════════════════
// ENTERPRISE BUSINESS CONTROL TOWER — Apps Script Backend
// Integrates Marketplace + Open Marketplace + EPR Sheets
// ═══════════════════════════════════════════════════════════════

var CONFIG = {
  MARKETPLACE_SHEET_ID: '1AvZZtujyTbV_fUQ-azIeOnm5xxrujWum',
  OMP_SHEET_ID:         '1z_4LmDjK1aMgcCR0MOVQ595wr_A4--zN5gWoJVWR6kw',
  EPR_SHEET_ID:         '1Mvkshz6Es3V37GYuXVMIq4L8ql3MCCuIxYjc77YEq5c',
  CACHE_TTL:      55,         // 55 s — must be < 60 s frontend poll so each auto-refresh gets fresh data
};

// Business-vertical mapping for COP Seller Details (Marketplace) categories.
// Metal + Institutional Business roll up under Infra Business per business rule.
var VERTICAL_MAP = {
  'Metal':                  'Infra Business',
  'Institutional Business': 'Infra Business',
  'IB':                     'Infra Business',
  'Re-Commerce':            'Recommerce',
  'Recommerce':             'Recommerce',
  'AFR':                    'AFR',
  'GOA DRS':                'DRS',
  'DRS':                    'DRS',
};

// Open Marketplace is tracked from FY 26-27 (Apr 1, 2026) onward —
// older OMP records are excluded from the whole dashboard.
var OMP_DATA_START = new Date(2026, 3, 1);

// ─────────────────────────────────────────────────────────────
// ENTRY POINTS
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Enterprise Business Control Tower')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData(filtersJson) {
  try {
    var filters  = filtersJson ? JSON.parse(filtersJson) : {};
    var cacheKey = 'dash_v5_' + JSON.stringify(filters);
    var cache    = CacheService.getScriptCache();
    var cached   = cache.get(cacheKey);
    if (cached) return cached;

    var mpRaw  = readSheet(CONFIG.MARKETPLACE_SHEET_ID);
    var ompRaw = readSheet(CONFIG.OMP_SHEET_ID);
    var eprRaw = CONFIG.EPR_SHEET_ID ? readSheet(CONFIG.EPR_SHEET_ID) : { headers: [], rows: [] };

    var mpAll  = normalizeMarketplace(mpRaw);
    var ompAll = normalizeOMP(ompRaw).filter(function(r) {
      return !r.createdDate || r.createdDate >= OMP_DATA_START;
    });
    var eprAll = normalizeEPR(eprRaw);

    var mpData  = filterByPeriod(mpAll, filters);
    var ompData = filterByPeriod(ompAll, filters);
    var eprData = filterByPeriod(eprAll, filters);

    var result = JSON.stringify({
      success:     true,
      lastUpdated: new Date().toISOString(),
      filters:     filters,
      enterprise:  calcEnterpriseKPIs(mpData, ompData, eprData),
      tat:         calcTAT(mpData),
      verticals:   calcVerticalKPIs(mpAll, ompAll, eprAll, filters),
      fiscalYears: buildFiscalYears(mpAll.concat(ompAll).concat(eprAll)),
    });

    try { cache.put(cacheKey, result, CONFIG.CACHE_TTL); } catch(e) {}
    return result;

  } catch(err) {
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
// Columns (sheet): id, business_name, state, vendor_type,
//   business_category, gstin_number, status, current_status,
//   status_changed_on, created_date, review_submission_date,
//   party_id, onboarded_date, Firstshipmentdate, LifetimeValue
// ─────────────────────────────────────────────────────────────

function normalizeMarketplace(raw) {
  var idx = buildIndex(raw.headers);
  var now = new Date();

  return raw.rows.map(function(row) {
    var created   = parseDate(gv(row, idx, 'created_date'));
    var onboarded = parseDate(gv(row, idx, 'onboarded_date'));
    var shipment  = parseDate(gv(row, idx, 'firstshipmentdate'));
    var gstin     = String(gv(row, idx, 'gstin_number') || '').trim();
    var ltv       = parseNumber(gv(row, idx, 'lifetimevalue'));
    var cat       = normCategory(gv(row, idx, 'business_category'));

    var l1r1 = parseDate(gv(row, idx, 'level1_rejected1'));
    var l1r2 = parseDate(gv(row, idx, 'level1_rejected2'));
    var l1   = parseDate(gv(row, idx, 'level1'));
    var l2r1 = parseDate(gv(row, idx, 'level2_rejected1'));
    var l2r2 = parseDate(gv(row, idx, 'level2_rejected2'));
    var l2   = parseDate(gv(row, idx, 'level2'));
    var l3r1 = parseDate(gv(row, idx, 'level3_rejected1'));
    var l3r2 = parseDate(gv(row, idx, 'level3_rejected2'));
    var l3   = parseDate(gv(row, idx, 'level3'));
    var l4r1 = parseDate(gv(row, idx, 'level4_rejected1'));
    var l4r2 = parseDate(gv(row, idx, 'level4_rejected2'));
    var l4   = parseDate(gv(row, idx, 'level4'));
    var rejCount    = [l1r1,l1r2,l2r1,l2r2,l3r1,l3r2,l4r1,l4r2].filter(Boolean).length;
    var reviewStage = l4 ? 4 : l3 ? 3 : l2 ? 2 : l1 ? 1 : 0;

    return {
      id:             String(gv(row, idx, 'id') || '').replace(/,/g, '').trim(),
      name:           String(gv(row, idx, 'business_name') || '').trim(),
      state:          String(gv(row, idx, 'state') || '').trim(),
      vendorType:     String(gv(row, idx, 'vendor_type') || '').trim(),
      category:       cat,
      vertical:       VERTICAL_MAP[cat] || 'Marketplace',
      gstin:          gstin,
      partyId:        String(gv(row, idx, 'party_id') || '').trim(),
      status:         normStatus(gv(row, idx, 'status')),
      currentStatus:  normCurrentStatus(gv(row, idx, 'current_status')),
      statusChangedOn: parseDate(gv(row, idx, 'status_changed_on')),
      createdDate:    created,
      reviewDate:     parseDate(gv(row, idx, 'review_submission_date')),
      onboardedDate:  onboarded,
      shipmentDate:   shipment,
      ltv:            ltv,
      level1: l1, level1Rej1: l1r1, level1Rej2: l1r2,
      level2: l2, level2Rej1: l2r1, level2Rej2: l2r2,
      level3: l3, level3Rej1: l3r1, level3Rej2: l3r2,
      level4: l4, level4Rej1: l4r1, level4Rej2: l4r2,
      reviewStage:    reviewStage,
      rejectionCount: rejCount,
      onbTAT:         dateDiffDays(created, onboarded),
      shipTAT:        dateDiffDays(onboarded, shipment),
      age:            dateDiffDays(created, now),
      hasGST:         isValidGSTIN(gstin),
      hasShipment:    !!shipment,
      source:         'Marketplace',
    };
  }).filter(function(r) { return r.id || r.name; });
}

// ─────────────────────────────────────────────────────────────
// NORMALIZATION — OPEN MARKETPLACE
// Columns (sheet): seller_id, seller_name, business_category,
//   vendor_type, onboarding_status, onboarding_created_date,
//   onboarding_updated_date, onboarded_date, onboarding_age_days,
//   mobile, email, gstin, gstin_status, state, city,
//   total_listings, total_orders, first_listing_date,
//   first_order_date, days_to_first_listing, days_to_first_order,
//   last_listing_date, last_order_date, listing_activation_status,
//   transaction_activation_status, total_quantity, total_gmv,
//   cancelled_orders, completed_orders, seller_rating, osv_consent,
//   + 11 document columns
// ─────────────────────────────────────────────────────────────

var OMP_DOC_FIELDS = [
  { key: 'additional_documents',               label: 'Additional Docs'   },
  { key: 'proof_of_premises',                  label: 'Proof of Premises' },
  { key: 'electricity_bill',                   label: 'Electricity Bill'  },
  { key: 'kyc_document',                       label: 'KYC Document'      },
  { key: 'gst_portal_screenshot_bank_details', label: 'GST Portal / Bank' },
  { key: 'cancelled_cheque',                   label: 'Cancelled Cheque'  },
  { key: 'msme_certificate',                   label: 'MSME Certificate'  },
  { key: 'aadhaar',                            label: 'Aadhaar'           },
  { key: 'owner_pan',                          label: 'Owner PAN'         },
  { key: 'entity_pan',                         label: 'Entity PAN'        },
  { key: 'gst_certificate',                    label: 'GST Certificate'   },
];

function normalizeOMP(raw) {
  var idx = buildIndex(raw.headers);
  var now = new Date();

  return raw.rows.map(function(row) {
    var gstin     = String(gv(row, idx, 'gstin') || '').trim();
    var created   = parseDate(gv(row, idx, 'onboarding_created_date'));
    var onboarded = parseDate(gv(row, idx, 'onboarded_date'));
    var docs      = {};
    var filled    = 0;

    OMP_DOC_FIELDS.forEach(function(f) {
      var n = parseInt(gv(row, idx, f.key) || 0, 10) || 0;
      docs[f.key] = n;
      if (n > 0) filled++;
    });

    var docPct      = Math.round((filled / OMP_DOC_FIELDS.length) * 100);
    var missingDocs = OMP_DOC_FIELDS
      .filter(function(f) { return (docs[f.key] || 0) === 0; })
      .map(function(f) { return f.label; });

    var listingStatus = String(gv(row, idx, 'listing_activation_status') || '').trim().toUpperCase();
    var txnStatus     = String(gv(row, idx, 'transaction_activation_status') || '').trim().toUpperCase();
    var rating        = parseNumber(gv(row, idx, 'seller_rating'));
    var osvRaw        = String(gv(row, idx, 'osv_consent') || '').trim().toUpperCase();

    return {
      id:              String(gv(row, idx, 'seller_id') || '').replace(/,/g, '').trim(),
      name:            String(gv(row, idx, 'seller_name') || '').trim(),
      category:        normCategory(gv(row, idx, 'business_category')),
      vendorType:      String(gv(row, idx, 'vendor_type') || '').trim(),
      status:          normStatus(gv(row, idx, 'onboarding_status')),
      createdDate:     created,
      onboardedDate:   onboarded,
      updatedDate:     parseDate(gv(row, idx, 'onboarding_updated_date')),
      onbAgeDays:      parseNumber(gv(row, idx, 'onboarding_age_days')),
      mobile:          String(gv(row, idx, 'mobile') || '').trim(),
      email:           String(gv(row, idx, 'email') || '').trim(),
      gstin:           gstin,
      gstinStatus:     String(gv(row, idx, 'gstin_status') || '').trim(),
      state:           String(gv(row, idx, 'state') || '').trim(),
      city:            String(gv(row, idx, 'city') || '').trim(),
      totalListings:   parseNumber(gv(row, idx, 'total_listings')),
      totalOrders:     parseNumber(gv(row, idx, 'total_orders')),
      totalQuantity:   parseNumber(gv(row, idx, 'total_quantity')),
      totalGMV:        parseNumber(gv(row, idx, 'total_gmv')),
      cancelledOrders: parseNumber(gv(row, idx, 'cancelled_orders')),
      completedOrders: parseNumber(gv(row, idx, 'completed_orders')),
      firstListingDate: parseDate(gv(row, idx, 'first_listing_date')),
      firstOrderDate:   parseDate(gv(row, idx, 'first_order_date')),
      lastListingDate:  parseDate(gv(row, idx, 'last_listing_date')),
      lastOrderDate:    parseDate(gv(row, idx, 'last_order_date')),
      daysToFirstListing: parseNumber(gv(row, idx, 'days_to_first_listing')),
      daysToFirstOrder:   parseNumber(gv(row, idx, 'days_to_first_order')),
      listingStatus:   listingStatus,
      txnStatus:       txnStatus,
      // 'INACTIVE' / 'NOT_ACTIVE' / 'NEVER_TRANSACTED' must NOT count as positive
      hasListing:      listingStatus.indexOf('ACTIVE') > -1 && listingStatus.indexOf('INACTIVE') === -1 && listingStatus.indexOf('NOT') === -1,
      hasTransacted:   txnStatus.indexOf('TRANSACTED') > -1 && txnStatus.indexOf('NOT') === -1 && txnStatus.indexOf('NEVER') === -1,
      sellerRating:    rating,
      osvConsent:      osvRaw === 'CONSENT_ACCEPTED',
      docs:            docs,
      docsFilled:      filled,
      docsTotal:       OMP_DOC_FIELDS.length,
      docPct:          docPct,
      missingDocs:     missingDocs,
      hasGST:          isValidGSTIN(gstin),
      onbTAT:          dateDiffDays(created, onboarded),
      age:             dateDiffDays(created, now),
      source:          'Open Marketplace',
    };
  }).filter(function(r) { return r.id || r.name; });
}

// ─────────────────────────────────────────────────────────────
// NORMALIZATION — EPR
// Mirrors Marketplace layout; reads from EPR_SHEET_ID when set
// ─────────────────────────────────────────────────────────────

function normalizeEPR(raw) {
  if (!raw || !raw.rows || raw.rows.length === 0) return [];
  var idx = buildIndex(raw.headers);
  var now = new Date();

  return raw.rows.map(function(row) {
    var fieldsJson = {};
    try {
      var fj = String(gv(row, idx, 'fields_json') || '');
      if (fj) fieldsJson = JSON.parse(fj);
    } catch(e) {}

    var co    = ((fieldsJson.service_provider_details || {}).company_details) || {};
    var gstin = String(co.gstin_number || gv(row, idx, 'gstin_number') || gv(row, idx, 'gstin') || '').trim();
    var state = String(co.state || '').trim();

    var created   = parseDate(gv(row, idx, 'created_date'));
    var onboarded = parseDate(gv(row, idx, 'onboarded_date'));
    var reviewd   = parseDate(gv(row, idx, 'review_submission_date'));

    // EPR created_date can be a bulk-import date later than the true onboarding
    // date, which makes created→onboarded negative. Fall back to review→onboarded.
    var tat = dateDiffDays(created, onboarded);
    if (tat === null || tat < 0) tat = dateDiffDays(reviewd, onboarded);
    if (tat !== null && tat < 0) tat = null;

    return {
      id:            String(gv(row, idx, 'id') || '').replace(/,/g, '').trim(),
      name:          String(gv(row, idx, 'business_name') || '').trim(),
      gstin:         gstin,
      status:        normStatus(gv(row, idx, 'status')),
      vendorType:    String(gv(row, idx, 'vendor_type') || '').trim(),
      category:      normCategory(gv(row, idx, 'business_category')),
      state:         state,
      email:         String(gv(row, idx, 'email') || '').trim(),
      mobile:        String(gv(row, idx, 'mobile') || '').trim(),
      segmentName:   String(gv(row, idx, 'segment_name') || '').trim(),
      rejectReason:  String(gv(row, idx, 'reject_reason') || '').trim(),
      globalPartyId: String(gv(row, idx, 'global_party_id') || '').trim(),
      eInvoiceTag:   String(gv(row, idx, 'e_invoice_tag') || '').trim(),
      approvalLevels: String(gv(row, idx, 'approval_levels') || '').trim(),
      reviewDate:    reviewd,
      createdDate:   created,
      onboardedDate: onboarded,
      onbTAT:        tat,
      age:           dateDiffDays(created, now),
      hasGST:        isValidGSTIN(gstin),
      source:        'EPR',
    };
  }).filter(function(r) { return r.id || r.name; });
}

// ─────────────────────────────────────────────────────────────
// PERIOD FILTER  (the only server-side filter; the rest are
// applied client-side inside the vertical detail view)
// ─────────────────────────────────────────────────────────────

function parseYMD(str, endOfDay) {
  var p = String(str || '').split('-');
  if (p.length !== 3) return null;
  var d = endOfDay
    ? new Date(+p[0], +p[1] - 1, +p[2], 23, 59, 59, 999)
    : new Date(+p[0], +p[1] - 1, +p[2]);
  return isNaN(d.getTime()) ? null : d;
}

function applyDateFilter(r, f) {
  if (!f || !f.period || f.period === 'All') return true;
  var d = r.createdDate;
  if (!d) return false;
  var now = new Date();
  var ps  = null;
  var pe  = null;
  if (f.period === 'Today') {
    ps = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  } else if (f.period === 'Custom' && f.startDate && f.endDate) {
    ps = parseYMD(f.startDate, false);
    pe = parseYMD(f.endDate, true);
  } else if (String(f.period).indexOf('FY') === 0) {
    // "FY25-26" → Apr 1 2025 – Mar 31 2026
    var m = String(f.period).match(/FY(\d{2})-(\d{2})/);
    if (m) {
      var fyStart = 2000 + parseInt(m[1], 10);
      ps = new Date(fyStart, 3, 1);
      pe = new Date(fyStart + 1, 2, 31, 23, 59, 59, 999);
    }
  }
  if (ps && d < ps) return false;
  if (pe && d > pe) return false;
  return true;
}

function filterByPeriod(data, f) {
  return data.filter(function(r) { return applyDateFilter(r, f); });
}

// ─────────────────────────────────────────────────────────────
// KPI CALCULATIONS
// ─────────────────────────────────────────────────────────────

function calcEnterpriseKPIs(mp, omp, epr) {
  var all            = mp.concat(omp).concat(epr);
  var totalCases     = all.length;
  var totalCompleted = count(all, 'status', 'COMPLETED');
  var totalDraft     = count(all, 'status', 'DRAFT');
  var totalInReview  = count(all, 'status', 'IN_REVIEW');
  var totalRejected  = count(all, 'status', 'REJECTED');
  var tats           = all.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; }).map(function(r) { return r.onbTAT; });

  return {
    totalCases:     totalCases,
    totalCompleted: totalCompleted,
    totalPending:   totalDraft + totalInReview,
    totalDraft:     totalDraft,
    totalInReview:  totalInReview,
    totalRejected:  totalRejected,
    avgTAT:         tats.length ? Math.round(avg(tats)) : null,
    completionPct:  pct(totalCompleted, totalCases),
    mpTotal:        mp.length,
    ompTotal:       omp.length,
    eprTotal:       epr.length,
    categorySplit:  objToArr(groupBy(all, 'category')).slice(0, 12),
  };
}

// ─────────────────────────────────────────────────────────────
// VERTICAL KPIs — honours the Period filter (FY / MTD / …) only.
// Category / status / search filters are NOT applied here so the
// vertical definitions stay intact.
// ─────────────────────────────────────────────────────────────

function calcVerticalKPIs(mpAll, ompAll, eprAll, filters) {
  var f = filters || {};
  function inPeriod(arr) { return arr.filter(function(r) { return applyDateFilter(r, f); }); }
  mpAll  = inPeriod(mpAll);
  ompAll = inPeriod(ompAll);
  eprAll = inPeriod(eprAll);

  var now     = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 86400000);

  function vertRow(r) {
    return {
      id:          r.id,
      name:        r.name,
      category:    r.category,
      vendorType:  r.vendorType,
      status:      r.status,
      gstin:       r.gstin || '—',
      hasGST:      r.hasGST,
      state:       r.state || '—',
      createdDate: fmtDate(r.createdDate),
      onbDate:     fmtDate(r.onboardedDate),
      tat:         (r.onbTAT !== null && r.onbTAT >= 0) ? r.onbTAT : '—',
    };
  }

  function vStats(data, mode) {
    var total     = data.length;
    var completed = count(data, 'status', 'COMPLETED');
    var draft     = count(data, 'status', 'DRAFT');
    var inReview  = count(data, 'status', 'IN_REVIEW');
    var rejected  = count(data, 'status', 'REJECTED');
    var withGST   = countFn(data, function(r) { return r.hasGST; });
    var tats      = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; }).map(function(r) { return r.onbTAT; });

    var completedThisWeek = countFn(data, function(r) {
      return r.onboardedDate && r.onboardedDate >= weekAgo;
    });

    // Transaction signal per source: OMP = TRANSACTED activation status,
    // Marketplace verticals = first shipment recorded, EPR = not applicable.
    var transacted = mode === 'omp' ? countFn(data, function(r) { return r.hasTransacted; })
                   : mode === 'mp'  ? countFn(data, function(r) { return r.hasShipment; })
                   : null;

    var latestFirst = data.slice().sort(function(a, b) {
      return (b.createdDate ? b.createdDate.getTime() : 0) - (a.createdDate ? a.createdDate.getTime() : 0);
    });

    return {
      total:            total,
      completed:        completed,
      draft:            draft,
      inReview:         inReview,
      rejected:         rejected,
      withGST:          withGST,
      missingGST:       total - withGST,
      completionPct:    pct(completed, total),
      completedThisWeek: completedThisWeek,
      avgTAT:           tats.length ? Math.round(avg(tats)) : null,
      transacted:       transacted,
      pctTransacted:    transacted === null ? null : pct(transacted, completed),
      rows:             latestFirst.map(vertRow),
      rowsTotal:        total,
    };
  }

  function byVertical(name) {
    return mpAll.filter(function(r) { return r.vertical === name; });
  }

  return {
    AFR:           vStats(byVertical('AFR'),            'mp'),
    DRS:           vStats(byVertical('DRS'),            'mp'),
    EPR:           vStats(eprAll,                        'epr'),
    InfraBusiness: vStats(byVertical('Infra Business'), 'mp'),
    OMP:           vStats(ompAll,                        'omp'),
    Recommerce:    vStats(byVertical('Recommerce'),     'mp'),
  };
}

// ─────────────────────────────────────────────────────────────
// TAT  (Marketplace onboarding TAT only)
// ─────────────────────────────────────────────────────────────

function calcTAT(data) {
  var withTAT = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; });
  if (!withTAT.length) return { avg: null, min: null, max: null, median: null, totalWithTAT: 0, distribution: [] };

  var vals = withTAT.map(function(r) { return r.onbTAT; }).sort(function(a, b) { return a - b; });

  var buckets = [
    { label: '0–2 days',   min: 0,  max: 2        },
    { label: '3–5 days',   min: 3,  max: 5        },
    { label: '6–10 days',  min: 6,  max: 10       },
    { label: '11–20 days', min: 11, max: 20       },
    { label: '21–30 days', min: 21, max: 30       },
    { label: '31+ days',   min: 31, max: Infinity },
  ].map(function(b) {
    return { label: b.label, count: countFn(withTAT, function(r) { return r.onbTAT >= b.min && r.onbTAT <= b.max; }) };
  });

  return {
    avg:          Math.round(avg(vals)),
    min:          vals[0],
    max:          vals[vals.length - 1],
    median:       Math.round(vals[Math.floor(vals.length / 2)]),
    totalWithTAT: withTAT.length,
    distribution: buckets,
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
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  // GSheets serial number (Apps Script may not yet format as Date if cell is unformatted)
  if (typeof val === 'number') {
    if (val < 1) return null;
    var d = new Date((val - 25569) * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  var s = String(val).trim();
  if (!s || s === '—' || s === '-' || s === 'N/A') return null;
  // Strip appended time: "Jun 30, 2025, 4:35 pm" → "Jun 30, 2025"
  var noTime = s.replace(/,?\s+\d{1,2}:\d{2}(:\d{2})?(\s*(am|pm))?/i, '').trim();
  var d1 = new Date(noTime);
  if (!isNaN(d1.getTime())) return d1;
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

function parseNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  return parseFloat(String(val).replace(/,/g, '').replace(/[^0-9.\-]/g, '')) || 0;
}

function dateDiffDays(d1, d2) {
  if (!d1 || !d2) return null;
  var diff = Math.round((d2.getTime() - d1.getTime()) / 86400000);
  return diff;
}

function isValidGSTIN(g) {
  // Strict 15-char alphanumeric GSTIN (e.g. 27AAACR1234A1Z5); anything
  // shorter/junk ("NA", "URP", "0") counts as GST-inactive.
  return /^[0-9A-Z]{15}$/.test(String(g || '').trim().toUpperCase());
}

function normStatus(v) {
  var s = String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  if (!s) return 'UNKNOWN';
  var map = {
    'COMPLETED':    'COMPLETED',
    'COMPLETE':     'COMPLETED',
    'ONBOARDED':    'COMPLETED',
    'DRAFT':        'DRAFT',
    'PENDING':      'DRAFT',
    'NEW':          'DRAFT',
    'IN_REVIEW':    'IN_REVIEW',
    'INREVIEW':     'IN_REVIEW',
    'REVIEW':       'IN_REVIEW',
    'UNDER_REVIEW': 'IN_REVIEW',
    'IN_PROGRESS':  'IN_REVIEW',
    'REJECTED':     'REJECTED',
    'REJECT':       'REJECTED',
    'DECLINED':     'REJECTED',
  };
  return map[s] || s;
}

function normCurrentStatus(v) {
  var s = String(v || '').toUpperCase().trim().replace(/\s+/g, '_');
  if (!s) return 'UNKNOWN';
  var map = {
    'ACTIVE':       'ACTIVE',
    'CHURNED':      'CHURNED',
    'CHURN':        'CHURNED',
    'DEACTIVATED':  'DEACTIVATED',
    'DEACTIVATE':   'DEACTIVATED',
    'INACTIVE':     'DEACTIVATED',
    'REACTIVATED':  'REACTIVATED',
    'REACTIVATE':   'REACTIVATED',
  };
  return map[s] || s;
}

function normCategory(v) {
  var s = String(v || '').trim();
  return s || 'Others';
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
  arr.forEach(function(r) {
    var k = r[field] || 'Unknown';
    res[k] = (res[k] || 0) + 1;
  });
  return res;
}
function objToArr(obj) {
  return Object.keys(obj)
    .map(function(k) { return { label: k, count: obj[k] }; })
    .sort(function(a, b) { return b.count - a.count; });
}
function fyStartYear(d) {
  // Indian financial year: Apr 1 – Mar 31
  return d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;
}
function buildFiscalYears(all) {
  // All financial years present in the data, newest first (e.g. 'FY26-27')
  var nowFY = fyStartYear(new Date());
  var minFY = nowFY;
  all.forEach(function(r) {
    var d = r.createdDate || r.onboardedDate;
    if (d) {
      var y = fyStartYear(d);
      if (y >= 2015 && y < minFY) minFY = y;
    }
  });
  var list = [];
  for (var y = nowFY; y >= minFY; y--) {
    list.push('FY' + String(y).slice(2) + '-' + String(y + 1).slice(2));
  }
  return list;
}
function fmtDate(d) {
  if (!d) return '—';
  var dd = d.getDate();
  var mm = d.getMonth() + 1;
  var yy = d.getFullYear();
  return (dd < 10 ? '0' : '') + dd + '/' + (mm < 10 ? '0' : '') + mm + '/' + yy;
}
