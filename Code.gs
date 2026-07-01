// ═══════════════════════════════════════════════════════════════
// ENTERPRISE BUSINESS CONTROL TOWER — Apps Script Backend
// Integrates Marketplace + Open Marketplace + EPR Sheets
// ═══════════════════════════════════════════════════════════════

var CONFIG = {
  MARKETPLACE_SHEET_ID: '1AvZZtujyTbV_fUQ-azIeOnm5xxrujWum',
  OMP_SHEET_ID:         '1z_4LmDjK1aMgcCR0MOVQ595wr_A4--zN5gWoJVWR6kw',
  EPR_SHEET_ID:         '',   // set this when the EPR Google Sheet is ready
  CACHE_TTL:      55,         // 55 s — must be < 60 s frontend poll so each auto-refresh gets fresh data
  MAX_TABLE_ROWS: 300,        // per source
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

function getDashboardData(filtersJson) {
  try {
    var filters  = filtersJson ? JSON.parse(filtersJson) : {};
    var cacheKey = 'dash_v3_' + JSON.stringify(filters);
    var cache    = CacheService.getScriptCache();
    var cached   = cache.get(cacheKey);
    if (cached) return cached;

    var mpRaw  = readSheet(CONFIG.MARKETPLACE_SHEET_ID);
    var ompRaw = readSheet(CONFIG.OMP_SHEET_ID);
    var eprRaw = CONFIG.EPR_SHEET_ID ? readSheet(CONFIG.EPR_SHEET_ID) : { headers: [], rows: [] };

    var mpAll  = normalizeMarketplace(mpRaw);
    var ompAll = normalizeOMP(ompRaw);
    var eprAll = normalizeEPR(eprRaw);

    var mpData  = filterMarketplace(mpAll, filters);
    var ompData = filterOMP(ompAll, filters);
    var eprData = filterEPR(eprAll, filters);

    var result = JSON.stringify({
      success:       true,
      lastUpdated:   new Date().toISOString(),
      filters:       filters,
      enterprise:    calcEnterpriseKPIs(mpData, ompData, eprData),
      marketplace:   calcMarketplaceKPIs(mpData),
      omp:           calcOMPKPIs(ompData),
      epr:           calcEPRKPIs(eprData),
      trends:        calcTrends(mpData, ompData, eprData),
      tat:           calcTAT(mpData),
      table:         buildTable(mpData, ompData, eprData),
      filterOptions: buildFilterOptions(
        (!filters.vertical || filters.vertical === 'All' || filters.vertical === 'Marketplace') ? mpAll  : [],
        (!filters.vertical || filters.vertical === 'All' || filters.vertical === 'OMP')         ? ompAll : [],
        (!filters.vertical || filters.vertical === 'All' || filters.vertical === 'EPR')         ? eprAll : []
      ),
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

    return {
      id:             String(gv(row, idx, 'id') || '').replace(/,/g, '').trim(),
      name:           String(gv(row, idx, 'business_name') || '').trim(),
      state:          String(gv(row, idx, 'state') || '').trim(),
      vendorType:     String(gv(row, idx, 'vendor_type') || '').trim(),
      category:       normCategory(gv(row, idx, 'business_category')),
      gstin:          gstin,
      status:         normStatus(gv(row, idx, 'status')),
      currentStatus:  normCurrentStatus(gv(row, idx, 'current_status')),
      statusChangedOn: parseDate(gv(row, idx, 'status_changed_on')),
      createdDate:    created,
      reviewDate:     parseDate(gv(row, idx, 'review_submission_date')),
      onboardedDate:  onboarded,
      shipmentDate:   shipment,
      ltv:            ltv,
      onbTAT:         dateDiffDays(created, onboarded),
      shipTAT:        dateDiffDays(onboarded, shipment),
      age:            dateDiffDays(created, now),
      hasGST:         gstin.length > 5,
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

    var listingStatus = String(gv(row, idx, 'listing_activation_status') || '').trim();
    var txnStatus     = String(gv(row, idx, 'transaction_activation_status') || '').trim();
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
      hasListing:      listingStatus.indexOf('ACTIVE') > -1,
      hasTransacted:   txnStatus.indexOf('TRANSACTED') > -1 && txnStatus.indexOf('NOT') === -1,
      sellerRating:    rating,
      osvConsent:      osvRaw === 'CONSENT_ACCEPTED',
      docs:            docs,
      docsFilled:      filled,
      docsTotal:       OMP_DOC_FIELDS.length,
      docPct:          docPct,
      missingDocs:     missingDocs,
      hasGST:          gstin.length > 5,
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
    // accept both 'gstin' and 'gstin_number' column names
    var gstin     = String(gv(row, idx, 'gstin') || gv(row, idx, 'gstin_number') || '').trim();
    var created   = parseDate(gv(row, idx, 'created_date'));
    var onboarded = parseDate(gv(row, idx, 'onboarded_date'));

    return {
      id:            String(gv(row, idx, 'id') || '').replace(/,/g, '').trim(),
      name:          String(gv(row, idx, 'business_name') || '').trim(),
      gstin:         gstin,
      status:        normStatus(gv(row, idx, 'status')),
      vendorType:    String(gv(row, idx, 'vendor_type') || '').trim(),
      category:      normCategory(gv(row, idx, 'business_category')),
      createdDate:   created,
      onboardedDate: onboarded,
      onbTAT:        dateDiffDays(created, onboarded),
      age:           dateDiffDays(created, now),
      hasGST:        gstin.length > 5,
      source:        'EPR',
    };
  }).filter(function(r) { return r.id || r.name; });
}

// ─────────────────────────────────────────────────────────────
// FILTERS
// ─────────────────────────────────────────────────────────────

function applyDateFilter(r, f) {
  if (!f.period || f.period === 'All') return true;
  var d = r.createdDate;
  if (!d) return false;
  var now = new Date();
  var ps  = null;
  var pe  = null;
  if (f.period === 'Today') {
    ps = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  } else if (f.period === 'MTD') {
    ps = new Date(now.getFullYear(), now.getMonth(), 1);
  } else if (f.period === 'YTD') {
    ps = new Date(now.getFullYear(), 0, 1);
  } else if (f.period === 'Custom' && f.startDate && f.endDate) {
    ps = new Date(f.startDate);
    pe = new Date(f.endDate);
    pe.setHours(23, 59, 59, 999);
  }
  if (ps && d < ps) return false;
  if (pe && d > pe) return false;
  return true;
}

function applyCommonFilters(r, f) {
  if (f.category   && f.category   !== 'All' && r.category   !== f.category)   return false;
  if (f.vendorType && f.vendorType !== 'All' && r.vendorType !== f.vendorType)  return false;
  if (f.status     && f.status     !== 'All' && r.status     !== f.status)      return false;
  if (f.search) {
    var q = f.search.toLowerCase();
    if (!(r.name  || '').toLowerCase().includes(q) &&
        !(r.gstin || '').toLowerCase().includes(q) &&
        !(r.id    || '').includes(q)) return false;
  }
  return true;
}

function filterMarketplace(data, f) {
  return data.filter(function(r) {
    if (f.vertical && f.vertical !== 'All' && f.vertical !== 'Marketplace') return false;
    if (f.state && f.state !== 'All' && r.state !== f.state) return false;
    if (!applyCommonFilters(r, f)) return false;
    if (!applyDateFilter(r, f))    return false;
    return true;
  });
}

function filterOMP(data, f) {
  return data.filter(function(r) {
    if (f.vertical && f.vertical !== 'All' && f.vertical !== 'OMP') return false;
    if (f.state && f.state !== 'All' && r.state !== f.state) return false;
    if (!applyCommonFilters(r, f)) return false;
    if (!applyDateFilter(r, f))    return false;
    return true;
  });
}

function filterEPR(data, f) {
  return data.filter(function(r) {
    if (f.vertical && f.vertical !== 'All' && f.vertical !== 'EPR') return false;
    if (!applyCommonFilters(r, f)) return false;
    if (!applyDateFilter(r, f))    return false;
    return true;
  });
}

// ─────────────────────────────────────────────────────────────
// KPI CALCULATIONS
// ─────────────────────────────────────────────────────────────

function calcEnterpriseKPIs(mp, omp, epr) {
  var all           = mp.concat(omp).concat(epr);
  var totalCases    = all.length;
  var totalCompleted = count(all, 'status', 'COMPLETED');
  var totalPending  = countFn(all, function(r) { return r.status === 'DRAFT' || r.status === 'IN_REVIEW'; });
  var totalRejected = count(all, 'status', 'REJECTED');
  var totalLTV      = mp.reduce(function(s, r) { return s + (r.ltv || 0); }, 0);
  var ompGMV        = omp.reduce(function(s, r) { return s + (r.totalGMV || 0); }, 0);
  var tatSrc        = mp.concat(omp).concat(epr);
  var tats          = tatSrc.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; }).map(function(r) { return r.onbTAT; });

  return {
    totalCases:     totalCases,
    totalCompleted: totalCompleted,
    totalPending:   totalPending,
    totalRejected:  totalRejected,
    totalLTV:       totalLTV,
    ompGMV:         ompGMV,
    avgTAT:         tats.length ? Math.round(avg(tats)) : 0,
    completionPct:  pct(totalCompleted, totalCases),
    mpTotal:        mp.length,
    ompTotal:       omp.length,
    eprTotal:       epr.length,
    mpPct:          pct(mp.length,  totalCases),
    ompPct:         pct(omp.length, totalCases),
    eprPct:         pct(epr.length, totalCases),
    mpCompleted:    count(mp,  'status', 'COMPLETED'),
    ompCompleted:   count(omp, 'status', 'COMPLETED'),
    eprCompleted:   count(epr, 'status', 'COMPLETED'),
  };
}

function calcMarketplaceKPIs(data) {
  var total     = data.length;
  var completed = count(data, 'status', 'COMPLETED');
  var draft     = count(data, 'status', 'DRAFT');
  var inReview  = count(data, 'status', 'IN_REVIEW');
  var rejected  = count(data, 'status', 'REJECTED');
  var active    = countFn(data, function(r) { return r.currentStatus === 'ACTIVE' || r.currentStatus === 'REACTIVATED'; });
  var churned   = count(data, 'currentStatus', 'CHURNED');
  var deact     = count(data, 'currentStatus', 'DEACTIVATED');
  var react     = count(data, 'currentStatus', 'REACTIVATED');
  var withShip  = countFn(data, function(r) { return r.hasShipment; });
  var noShip    = Math.max(0, completed - withShip);
  var misGST    = countFn(data, function(r) { return !r.hasGST; });
  var ltv       = data.reduce(function(s, r) { return s + (r.ltv || 0); }, 0);
  var tats      = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; }).map(function(r) { return r.onbTAT; });

  var now    = new Date();
  var today0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var mtd0   = new Date(now.getFullYear(), now.getMonth(), 1);
  var ytd0   = new Date(now.getFullYear(), 0, 1);

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
    noShipment:   noShip,
    missingGST:   misGST,
    totalLTV:     ltv,
    avgLTV:       total ? ltv / total : 0,
    avgTAT:       tats.length ? Math.round(avg(tats)) : 0,
    completionPct: pct(completed, total),
    shipmentRate:  pct(withShip, completed),
    activeRate:    pct(active, completed),
    createdToday:  countFn(data, function(r) { return r.createdDate   && r.createdDate   >= today0; }),
    createdMTD:    countFn(data, function(r) { return r.createdDate   && r.createdDate   >= mtd0;   }),
    createdYTD:    countFn(data, function(r) { return r.createdDate   && r.createdDate   >= ytd0;   }),
    onbToday:      countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= today0; }),
    onbMTD:        countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= mtd0;   }),
    onbYTD:        countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= ytd0;   }),
    shipToday:     countFn(data, function(r) { return r.shipmentDate  && r.shipmentDate  >= today0; }),
    shipMTD:       countFn(data, function(r) { return r.shipmentDate  && r.shipmentDate  >= mtd0;   }),
    shipYTD:       countFn(data, function(r) { return r.shipmentDate  && r.shipmentDate  >= ytd0;   }),
    statusSplit:    objToArr(groupBy(data, 'status')),
    curStatusSplit: objToArr(groupBy(data, 'currentStatus')),
    categorySplit:  objToArr(groupBy(data, 'category')),
    vendorSplit:    objToArr(groupBy(data, 'vendorType')),
    stateSplit:     objToArr(groupBy(data, 'state')).slice(0, 12),
  };
}

function calcOMPKPIs(data) {
  var total      = data.length;
  var completed  = count(data, 'status', 'COMPLETED');
  var draft      = count(data, 'status', 'DRAFT');
  var inReview   = count(data, 'status', 'IN_REVIEW');
  var rejected   = count(data, 'status', 'REJECTED');
  var withGST    = countFn(data, function(r) { return r.hasGST; });
  var avgDoc     = total ? Math.round(data.reduce(function(s, r) { return s + r.docPct; }, 0) / total) : 0;
  var withListing   = countFn(data, function(r) { return r.hasListing; });
  var withTransacted= countFn(data, function(r) { return r.hasTransacted; });
  var withOSV    = countFn(data, function(r) { return r.osvConsent; });
  var totalGMV   = data.reduce(function(s, r) { return s + (r.totalGMV || 0); }, 0);
  var totalQty   = data.reduce(function(s, r) { return s + (r.totalQuantity || 0); }, 0);
  var totalOrders= data.reduce(function(s, r) { return s + (r.totalOrders || 0); }, 0);
  var rated      = data.filter(function(r) { return r.sellerRating > 0; });
  var avgRating  = rated.length ? Math.round((rated.reduce(function(s, r) { return s + r.sellerRating; }, 0) / rated.length) * 10) / 10 : 0;
  var tats       = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; }).map(function(r) { return r.onbTAT; });

  var now    = new Date();
  var today0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var mtd0   = new Date(now.getFullYear(), now.getMonth(), 1);
  var ytd0   = new Date(now.getFullYear(), 0, 1);

  var docStats = OMP_DOC_FIELDS.map(function(f) {
    var submitted = countFn(data, function(r) { return (r.docs[f.key] || 0) > 0; });
    return { key: f.key, label: f.label, submitted: submitted, missing: total - submitted, pct: pct(submitted, total) };
  });

  return {
    total:            total,
    completed:        completed,
    draft:            draft,
    inReview:         inReview,
    rejected:         rejected,
    pending:          draft + inReview,
    withGST:          withGST,
    missingGST:       total - withGST,
    avgDocPct:        avgDoc,
    pipelinePct:      pct(draft + inReview, total),
    completionPct:    pct(completed, total),
    withListing:      withListing,
    withTransacted:   withTransacted,
    withOSV:          withOSV,
    totalGMV:         totalGMV,
    totalQuantity:    totalQty,
    totalOrders:      totalOrders,
    avgRating:        avgRating,
    avgTAT:           tats.length ? Math.round(avg(tats)) : 0,
    listingRate:      pct(withListing, completed),
    transactionRate:  pct(withTransacted, completed),
    createdToday:     countFn(data, function(r) { return r.createdDate   && r.createdDate   >= today0; }),
    createdMTD:       countFn(data, function(r) { return r.createdDate   && r.createdDate   >= mtd0;   }),
    createdYTD:       countFn(data, function(r) { return r.createdDate   && r.createdDate   >= ytd0;   }),
    onbToday:         countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= today0; }),
    onbMTD:           countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= mtd0;   }),
    onbYTD:           countFn(data, function(r) { return r.onboardedDate && r.onboardedDate >= ytd0;   }),
    statusSplit:      objToArr(groupBy(data, 'status')),
    categorySplit:    objToArr(groupBy(data, 'category')),
    vendorSplit:      objToArr(groupBy(data, 'vendorType')),
    stateSplit:       objToArr(groupBy(data, 'state')).slice(0, 12),
    docStats:         docStats,
  };
}

function calcEPRKPIs(data) {
  var total     = data.length;
  var completed = count(data, 'status', 'COMPLETED');
  var draft     = count(data, 'status', 'DRAFT');
  var inReview  = count(data, 'status', 'IN_REVIEW');
  var rejected  = count(data, 'status', 'REJECTED');
  var withGST   = countFn(data, function(r) { return r.hasGST; });

  return {
    total:         total,
    completed:     completed,
    draft:         draft,
    inReview:      inReview,
    rejected:      rejected,
    pending:       draft + inReview,
    withGST:       withGST,
    missingGST:    total - withGST,
    pipelinePct:   pct(draft + inReview, total),
    completionPct: pct(completed, total),
    statusSplit:   objToArr(groupBy(data, 'status')),
    categorySplit: objToArr(groupBy(data, 'category')),
    vendorSplit:   objToArr(groupBy(data, 'vendorType')),
  };
}

// ─────────────────────────────────────────────────────────────
// TRENDS  (MP + OMP + EPR — all now carry date fields)
// ─────────────────────────────────────────────────────────────

function calcTrends(mp, omp, epr) {
  var data = mp.concat(omp || []).concat(epr || []);
  var now  = new Date();
  var months = [];

  for (var i = 11; i >= 0; i--) {
    var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push({
      label: d.toLocaleString('default', { month: 'short' }) + ' \'' + String(d.getFullYear()).slice(2),
      year:  d.getFullYear(),
      month: d.getMonth(),
    });
  }

  var monthly = months.map(function(m) {
    return {
      label:     m.label,
      created:   countFn(data, function(r) { return r.createdDate   && r.createdDate.getFullYear()   === m.year && r.createdDate.getMonth()   === m.month; }),
      onboarded: countFn(data, function(r) { return r.onboardedDate && r.onboardedDate.getFullYear() === m.year && r.onboardedDate.getMonth() === m.month; }),
      shipped:   countFn(mp,   function(r) { return r.shipmentDate  && r.shipmentDate.getFullYear()  === m.year && r.shipmentDate.getMonth()  === m.month; }),
    };
  });

  var years  = [now.getFullYear() - 2, now.getFullYear() - 1, now.getFullYear()];
  var yearly = years.map(function(y) {
    return {
      label:     String(y),
      created:   countFn(data, function(r) { return r.createdDate   && r.createdDate.getFullYear()   === y; }),
      onboarded: countFn(data, function(r) { return r.onboardedDate && r.onboardedDate.getFullYear() === y; }),
    };
  });

  return { monthly: monthly, yearly: yearly };
}

// ─────────────────────────────────────────────────────────────
// TAT  (Marketplace onboarding TAT only)
// ─────────────────────────────────────────────────────────────

function calcTAT(data) {
  var withTAT = data.filter(function(r) { return r.onbTAT !== null && r.onbTAT >= 0; });
  if (!withTAT.length) return { avg: 0, min: 0, max: 0, median: 0, totalWithTAT: 0, distribution: [], categoryTAT: [], vendorTAT: [] };

  var vals = withTAT.map(function(r) { return r.onbTAT; }).sort(function(a, b) { return a - b; });

  var buckets = [
    { label: '0–2 days',   min: 0,  max: 2          },
    { label: '3–5 days',   min: 3,  max: 5          },
    { label: '6–10 days',  min: 6,  max: 10         },
    { label: '11–20 days', min: 11, max: 20         },
    { label: '21–30 days', min: 21, max: 30         },
    { label: '31+ days',   min: 31, max: Infinity   },
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
  var categoryTAT = Object.keys(catMap)
    .map(function(k) { return { category: k, avg: Math.round(avg(catMap[k])), count: catMap[k].length }; })
    .sort(function(a, b) { return b.avg - a.avg; });

  var vMap = {};
  withTAT.forEach(function(r) {
    var k = r.vendorType || 'Unknown';
    if (!vMap[k]) vMap[k] = [];
    vMap[k].push(r.onbTAT);
  });
  var vendorTAT = Object.keys(vMap)
    .map(function(k) { return { vendorType: k, avg: Math.round(avg(vMap[k])), count: vMap[k].length }; })
    .sort(function(a, b) { return b.avg - a.avg; })
    .slice(0, 10);

  return {
    avg:          Math.round(avg(vals)),
    min:          vals[0],
    max:          vals[vals.length - 1],
    median:       Math.round(vals[Math.floor(vals.length / 2)]),
    totalWithTAT: withTAT.length,
    distribution: buckets,
    categoryTAT:  categoryTAT,
    vendorTAT:    vendorTAT,
  };
}

// ─────────────────────────────────────────────────────────────
// TABLE DATA
// ─────────────────────────────────────────────────────────────

function buildTable(mp, omp, epr) {
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
      curStatus:   r.hasTransacted ? 'TRANSACTED' : (r.hasListing ? 'ACTIVE' : 'NOT ACTIVE'),
      gstin:       r.gstin || '—',
      state:       r.state || '—',
      city:        r.city  || '—',
      createdDate: fmtDate(r.createdDate),
      onbDate:     fmtDate(r.onboardedDate),
      shipDate:    '—',
      tat:         r.onbTAT !== null ? r.onbTAT : '—',
      ltv:         r.totalGMV > 0 ? fmtCurrency(r.totalGMV) : '—',
      hasGST:      r.hasGST,
      hasShip:     r.hasTransacted,
      docPct:      r.docPct + '%',
      totalListings: r.totalListings,
      totalOrders:   r.totalOrders,
      sellerRating:  r.sellerRating > 0 ? r.sellerRating : '—',
      osvConsent:    r.osvConsent,
    };
  });

  var eprRows = epr.slice(0, CONFIG.MAX_TABLE_ROWS).map(function(r) {
    return {
      id:          r.id,
      name:        r.name,
      source:      'EPR',
      category:    r.category,
      vendorType:  r.vendorType,
      status:      r.status,
      curStatus:   '—',
      gstin:       r.gstin || '—',
      state:       '—',
      createdDate: fmtDate(r.createdDate),
      onbDate:     fmtDate(r.onboardedDate),
      shipDate:    '—',
      tat:         r.onbTAT !== null ? r.onbTAT : '—',
      ltv:         '—',
      hasGST:      r.hasGST,
      hasShip:     false,
      docPct:      '—',
    };
  });

  var rows = mpRows.concat(ompRows).concat(eprRows);
  return {
    rows:     rows,
    mpCount:  mpRows.length,
    ompCount: ompRows.length,
    eprCount: eprRows.length,
    total:    rows.length,
  };
}

// ─────────────────────────────────────────────────────────────
// FILTER OPTIONS
// ─────────────────────────────────────────────────────────────

function buildFilterOptions(mp, omp, epr) {
  var all = mp.concat(omp).concat(epr);

  // Statuses: known order first, then any extras found in data
  var known  = ['COMPLETED', 'IN_REVIEW', 'DRAFT', 'REJECTED'];
  var inData = unique(all.map(function(r) { return r.status; }).filter(Boolean));
  var statuses = known.filter(function(s) { return inData.indexOf(s) > -1; })
                      .concat(inData.filter(function(s) { return known.indexOf(s) === -1; }));

  return {
    categories:  unique(all.map(function(r) { return r.category;   })).filter(Boolean).sort(),
    vendorTypes: unique(all.map(function(r) { return r.vendorType; })).filter(Boolean).sort(),
    states:      unique(mp.concat(omp).map(function(r) { return r.state; })).filter(Boolean).sort(),
    statuses:    statuses,
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
function unique(arr) {
  var seen = Object.create(null);
  return arr.filter(function(v) {
    if (seen[v]) return false;
    seen[v] = true;
    return true;
  });
}
function fmtDate(d) {
  if (!d) return '—';
  var dd = d.getDate();
  var mm = d.getMonth() + 1;
  var yy = d.getFullYear();
  return (dd < 10 ? '0' : '') + dd + '/' + (mm < 10 ? '0' : '') + mm + '/' + yy;
}
function fmtCurrency(v) {
  if (!v) return '₹0';
  if (v >= 10000000) return '₹' + (v / 10000000).toFixed(2) + ' Cr';
  if (v >= 100000)   return '₹' + (v / 100000).toFixed(2) + ' L';
  return '₹' + Math.round(v).toLocaleString('en-IN');
}
