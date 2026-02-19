// @ts-nocheck
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
function refreshSummary() {
    /***** ======================= CONFIG ======================= *****/
    var LTP_SHEET = (typeof LTP_SHEET_NAME !== 'undefined') ? LTP_SHEET_NAME : 'LTP';
    var HDATA_SHEET = (typeof SRC_SHEET_NAME !== 'undefined') ? SRC_SHEET_NAME : 'HData';
    var DSEX_SHEET = 'DSEX';
    var TZ = 'Asia/Dhaka';
    var LOOKBACK_DAYS = (typeof this.LOOKBACK_DAYS !== 'undefined') ? this.LOOKBACK_DAYS : 180;
    var MAX_HOLD_DAYS = 7;
    // Stops and bounds
    var STOP_FLOOR_PCT = (typeof this.STOP_FLOOR_PCT !== 'undefined') ? this.STOP_FLOOR_PCT : 0.02;
    var ATR_LEN = 14;
    var ATR_STOP_MULT = 1.8;
    var MIN_TARGET_PCT = 0.10;
    var MAX_TARGET_PCT = 0.20;
    var MIN_SAMPLES = 30;
    // Achievability settings
    var DESIRED_HIT_PROB = 0.60;
    var TARGET_FALLBACK_ATR_MULT = 1.6;
    // Universe & gating
    var ACTION_LIST = ['Purchased', 'Watch', 'Listed'];
    var MAX_CANDIDATES = 40;
    var MIN_VALID_DAYS = 60;
    // Filters
    var REQUIRE_PRICE_INCREASE = true;
    var MIN_PRICE_CHG_FILTER_PCT = 3;
    var REQUIRE_VALUE_INCREASE = true;
    var MIN_VALUE_INC_FILTER_PCT = 10;
    // Optional guards (OFF)
    var DEF = {
        USE_LIQUIDITY_GUARD: false,
        LIQ_MIN_VAL_MN_MED20: 10,
        ATR_BAND_MIN: 1.2,
        ATR_BAND_MAX: 8.0
    };
    // Bayesian prior for Win%
    var GLOBAL_WIN_MEAN = 0.45;
    var PRIOR_WEIGHT = 50;
    // ----- Investment settings (score-weighted) -----
    var TOTAL_CAPITAL_BDT = 100000;
    var NUM_SLOTS = 10;
    var BASE_PER_TRADE = TOTAL_CAPITAL_BDT / NUM_SLOTS;
    var WIN_BASE_PCT = 60;
    var WEIGHT_MIN = 0.5;
    var WEIGHT_MAX = 2.0;
    function investAmountFromWin_(winPct) {
        if (!Number.isFinite(winPct) || winPct <= 0)
            return 0;
        var w = Math.min(WEIGHT_MAX, Math.max(WEIGHT_MIN, (winPct / WIN_BASE_PCT)));
        return Math.round((BASE_PER_TRADE * w) / 100) * 100;
    }
    var DBG = {
        ltpRows: 0,
        hRows: 0,
        builtItems: 0,
        symbolsInSeries: 0,
        afterGate: 0,
        metricBlankRemoved: 0,
        positivityRemoved: 0,
        dedupedKept: 0,
        pinnedPlusKept: 0
    };
    /***** ======================= UTILS ======================= *****/
    var toNum = function (v) { return (typeof v === 'number') ? v : (v == null || v === '') ? NaN : parseFloat(String(v).replace(/[, ]/g, '')); };
    var fmtDate = function (d) { return Utilities.formatDate(d, TZ, 'dd-MMM-yy'); };
    var norm = function (s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '').trim(); };
    var median = function (a) {
        var b = a.filter(Number.isFinite).slice().sort(function (x, y) { return x - y; });
        if (!b.length)
            return NaN;
        var m = Math.floor(b.length / 2);
        return (b.length % 2) ? b[m] : (b[m - 1] + b[m]) / 2;
    };
    var hhmm = Utilities.formatDate(new Date(), TZ, 'HH:mm');
    var minutesNow = (+hhmm.slice(0, 2)) * 60 + (+hhmm.slice(3, 5));
    // ⛔ Skip updating between 09:00 and 10:00 (DSE shows pre-data)
    var BLOCK_START = 9 * 60;
    var BLOCK_END = 10 * 60;
    if (minutesNow >= BLOCK_START && minutesNow < BLOCK_END) {
        Logger.log('refreshSummary skipped between 09:00–10:00 (pre-data time).');
        return;
    }
    var AFTER_1430 = minutesNow >= (14 * 60 + 30);
    var IN_SESSION = minutesNow >= (10 * 60) && minutesNow < (14 * 60 + 30);
    function isTradingDayBD_(d) {
        if (d === void 0) { d = new Date(); }
        var ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
        var x = new Date(ds + 'T00:00:00');
        var wd = x.getDay(); // Fri=5, Sat=6
        return wd !== 5 && wd !== 6;
    }
    function addTradingDaysBD_(date, n) {
        var d = new Date(date);
        d.setHours(0, 0, 0, 0);
        var add = 0;
        while (add < n) {
            d.setDate(d.getDate() + 1);
            var wd = d.getDay();
            if (wd !== 5 && wd !== 6)
                add++;
        }
        return d;
    }
    function chooseBuyDate_() {
        var todayStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
        var today = new Date(todayStr + 'T00:00:00');
        var wd = today.getDay();
        if (!AFTER_1430 && wd !== 5 && wd !== 6)
            return today;
        return addTradingDaysBD_(today, 1);
    }
    function ema_(arr, len) {
        if (!arr || !arr.length)
            return [];
        var k = 2 / (len + 1);
        var out = new Array(arr.length);
        var e = arr[0];
        for (var i = 0; i < arr.length; i++) {
            var v = Number.isFinite(arr[i]) ? arr[i] : e;
            e = (i ? (v - e) * k + e : v);
            out[i] = e;
        }
        return out;
    }
    function macdHist_(cl) {
        if (!cl || cl.length < 26)
            return { hist: NaN, hist3: [] };
        var e12 = ema_(cl, 12);
        var e26 = ema_(cl, 26);
        var macd = cl.map(function (_, i) { return e12[i] - e26[i]; });
        var sig = ema_(macd, 9);
        var hist = macd[cl.length - 1] - sig[sig.length - 1];
        var hist3 = macd.slice(-3).map(function (_, k) { return macd[cl.length - 3 + k] - sig[sig.length - 3 + k]; });
        return { hist: hist, hist3: hist3 };
    }
    function linregSlopePct_(arr, days) {
        if (!arr || arr.length < days)
            return 0;
        var s = arr.slice(-days);
        var n = s.length;
        var sx = 0, sy = 0, sxx = 0, sxy = 0;
        for (var i = 0; i < n; i++) {
            var x = i + 1;
            var y = s[i];
            sx += x;
            sy += y;
            sxx += x * x;
            sxy += x * y;
        }
        var m = (n * sxy - sx * sy) / (n * sxx - sx * sx);
        var last = s[n - 1];
        return (last ? (m / last) * 100 : 0);
    }
    function atr_(hi, lo, cl, len) {
        var n = cl.length;
        var tr = new Array(n).fill(NaN);
        for (var i = 1; i < n; i++) {
            if (!Number.isFinite(hi[i]) || !Number.isFinite(lo[i]) || !Number.isFinite(cl[i - 1]))
                continue;
            tr[i] = Math.max(hi[i] - lo[i], Math.abs(hi[i] - cl[i - 1]), Math.abs(lo[i] - cl[i - 1]));
        }
        var out = new Array(n).fill(NaN);
        var sum = 0;
        for (var i = 0; i < n; i++) {
            var v = tr[i];
            if (!Number.isFinite(v)) {
                out[i] = (i ? out[i - 1] : NaN);
                continue;
            }
            if (i <= len) {
                sum += v;
                out[i] = (i === len) ? (sum / len) : NaN;
            }
            else {
                out[i] = (out[i - 1] * (len - 1) + v) / len;
            }
        }
        return out;
    }
    function rsi_(cl, len) {
        if (len === void 0) { len = 14; }
        if (!cl || cl.length < len + 1)
            return NaN;
        var g = 0, l = 0;
        for (var i = 1; i <= len; i++) {
            var d = cl[i] - cl[i - 1];
            if (d > 0)
                g += d;
            else
                l -= Math.min(d, 0);
        }
        g /= len;
        l /= len;
        var rs = (l === 0) ? 100 : g / l;
        var r = 100 - (100 / (1 + rs));
        for (var i = len + 1; i < cl.length; i++) {
            var d = cl[i] - cl[i - 1];
            var G = Math.max(d, 0);
            var L = Math.max(-d, 0);
            g = (g * (len - 1) + G) / len;
            l = (l * (len - 1) + L) / len;
            rs = (l === 0) ? 100 : g / l;
            r = 100 - (100 / (1 + rs));
        }
        return r;
    }
    /***** ======================= TREND (3-pillars base) ======================= *****/
    function trendPillars_(series, buyStartMs) {
        var hist = series.filter(function (b) { return b.ms < buyStartMs; });
        var N = 10;
        if (hist.length < Math.max(N, 20))
            return { sym: '=', score: 0, flags: { insufficient: true } };
        var lastN = hist.slice(-N);
        var clAll = hist.map(function (x) { return x.close; }).filter(Number.isFinite);
        // Pillar 1: Location
        var e20 = ema_(clAll, 20);
        var lastClose = clAll[clAll.length - 1];
        var lastE20 = e20[e20.length - 1];
        var loc = (Number.isFinite(lastClose) && Number.isFinite(lastE20) && lastClose > lastE20) &&
            (linregSlopePct_(e20.slice(-10), Math.min(10, e20.length)) > 0);
        // Pillar 2: Momentum
        var _a = macdHist_(clAll), mh = _a.hist, hist3 = _a.hist3;
        var macdUp = Number.isFinite(mh) && mh > 0;
        var macdRise = (hist3 && hist3.length === 3) ? (hist3[0] <= hist3[1] && hist3[1] <= hist3[2]) : false;
        var mom = (macdUp && macdRise) || (rsi_(clAll, 14) >= 55);
        // Pillar 3: Structure
        var HH = 0, HL = 0;
        for (var i = 1; i < lastN.length; i++) {
            var p = lastN[i - 1], c = lastN[i];
            if (Number.isFinite(c.high) && Number.isFinite(p.high) && c.high > p.high)
                HH++;
            if (Number.isFinite(c.low) && Number.isFinite(p.low) && c.low > p.low)
                HL++;
        }
        var struct = (HH >= 4 && HL >= 4) || (linregSlopePct_(lastN.map(function (x) { return x.close; }), Math.min(7, lastN.length)) > 0);
        var pass = (loc ? 1 : 0) + (mom ? 1 : 0) + (struct ? 1 : 0);
        return { sym: pass >= 2 ? '⬆️' : (pass === 0 ? '⬇️' : '='), score: pass, flags: { Location: loc, Momentum: mom, Structure: struct } };
    }
    /***** ======================= DSEX HELPERS ======================= *****/
    function readDSEX_CloseSeries_(ss, cutoffMs, buyStartMs) {
        var sh = ss.getSheetByName(DSEX_SHEET);
        if (!sh)
            return null;
        var lc = sh.getLastColumn();
        var lr = sh.getLastRow();
        if (lr < 2)
            return null;
        var H = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (h) { return String(h).trim().toUpperCase(); });
        var iDate = H.findIndex(function (h) { return h === 'DATE'; });
        var iDSEX = H.findIndex(function (h) { return h.includes('DSEX INDEX'); });
        if (iDate < 0 || iDSEX < 0)
            return null;
        var vals = sh.getRange(2, 1, lr - 1, lc).getValues();
        var ser = [];
        for (var _i = 0, vals_1 = vals; _i < vals_1.length; _i++) {
            var r = vals_1[_i];
            var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
            if (isNaN(d))
                continue;
            var ms = d.getTime();
            if (ms < cutoffMs || ms >= buyStartMs)
                continue;
            var c = Number(r[iDSEX]);
            if (Number.isFinite(c))
                ser.push({ ms: ms, close: c });
        }
        ser.sort(function (a, b) { return a.ms - b.ms; });
        return ser.length ? ser : null;
    }
    function regimeFromCloses_(cl) {
        if (!cl || cl.length < 40) {
            return { ok: true, sym: '=' };
        }
        var slope20 = linregSlopePct_(cl, 20);
        var hist = macdHist_(cl).hist;
        var macd = Number.isFinite(hist) ? hist : 0;
        var strongUp = (slope20 > 0.05) || (macd > 0.05);
        var neutralOk = (slope20 > -0.10) || (macd > -0.05);
        if (strongUp)
            return { ok: true, sym: '⬆️' };
        if (neutralOk)
            return { ok: true, sym: '=' };
        return { ok: false, sym: '⬇️' };
    }
    function displayTrendSym_(sym) {
        if (sym === '⬆️')
            return '⬆️';
        if (sym === '⬇️')
            return '⬇️';
        return '〰️';
    }
    /***** ======================= SHEETS / HEADERS ======================= *****/
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ltp = ss.getSheetByName(LTP_SHEET);
    var hdata = ss.getSheetByName(HDATA_SHEET);
    if (!ltp || !hdata)
        return safeWrite_([], ACTION_LIST, 'Missing LTP/HData');
    var lr = ltp.getLastRow();
    var lc = ltp.getLastColumn();
    if (lr < 2)
        return safeWrite_([], ACTION_LIST, 'No rows in LTP');
    var HRowDisp = ltp.getRange(1, 1, 1, lc).getDisplayValues()[0];
    var NU = HRowDisp.map(function (h) { return norm(String(h).replace(/\*/g, '')); });
    var findHeader = function (keys) {
        var arr = Array.isArray(keys) ? keys : [keys];
        for (var i = 0; i < NU.length; i++) {
            for (var _i = 0, arr_1 = arr; _i < arr_1.length; _i++) {
                var k = arr_1[_i];
                if (NU[i].includes(norm(k)))
                    return i;
            }
        }
        return null;
    };
    var colCode = findHeader(['TRADINGCODE']);
    var colSector = findHeader(['SECTOR']);
    var colValMn = findHeader(['VALUEMN', 'VALUE']);
    var colLtpPx = findHeader(['LTP']);
    var colYcpPx = findHeader(['YCP']);
    var colCloseP = findHeader(['CLOSEP']);
    var colValSpike = findHeader(['VALUESPIKE']);
    if (colCode == null)
        return safeWrite_([], ACTION_LIST, 'Missing TRADING CODE in LTP');
    var ltpVals = ltp.getRange(2, 1, lr - 1, lc).getValues();
    DBG.ltpRows = ltpVals.length;
    var LTP_CODES = new Set(ltpVals.map(function (r) { return norm(r[colCode]); }).filter(Boolean));
    var valueSpikeByCode = new Map();
    for (var _i = 0, ltpVals_1 = ltpVals; _i < ltpVals_1.length; _i++) {
        var r = ltpVals_1[_i];
        var code = norm(r[colCode]);
        if (!code)
            continue;
        if (colValSpike != null) {
            var vs = toNum(r[colValSpike]);
            if (Number.isFinite(vs))
                valueSpikeByCode.set(code, vs);
        }
    }
    var ltpPriceByCode = new Map();
    for (var _a = 0, ltpVals_2 = ltpVals; _a < ltpVals_2.length; _a++) {
        var r = ltpVals_2[_a];
        var code = norm(r[colCode]);
        if (!code)
            continue;
        if (colLtpPx != null) {
            var px = toNum(r[colLtpPx]);
            if (Number.isFinite(px))
                ltpPriceByCode.set(code, px);
        }
    }
    // HData mapping
    var h_lc = hdata.getLastColumn();
    var Hh = hdata.getRange(1, 1, 1, h_lc).getDisplayValues()[0].map(function (h) { return norm(String(h).replace(/\*/g, '')); });
    var iDate = Hh.indexOf('DATE') >= 0 ? Hh.indexOf('DATE') : 1;
    var iCode = Hh.findIndex(function (x) { return x.includes('TRADINGCODE'); }) >= 0 ? Hh.findIndex(function (x) { return x.includes('TRADINGCODE'); }) : 2;
    var iHigh = Hh.indexOf('HIGH');
    var iLow = Hh.indexOf('LOW');
    var iOpen = Hh.findIndex(function (x) { return x.startsWith('OPEN'); });
    var iClose = Hh.findIndex(function (x) { return x.startsWith('CLOSE'); });
    var iYcp = Hh.findIndex(function (x) { return x.includes('YCP'); });
    var iVal = Hh.findIndex(function (x) { return x === 'VALUEMN' || x === 'VALUE' || x.includes('VALUE'); });
    var iVol = Hh.findIndex(function (x) { return x === 'VOLUME' || x.includes('VOLUME'); });
    if ([iDate, iCode, iClose].some(function (i) { return i < 0; }))
        return safeWrite_([], ACTION_LIST, 'Missing Date/Code/Close in HData');
    var h_lr = hdata.getLastRow();
    if (h_lr < 2)
        return safeWrite_([], ACTION_LIST, 'HData empty');
    var todayISO = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var todayStart = new Date(todayISO + 'T00:00:00');
    var todayStartMs = todayStart.getTime();
    var cutoff = new Date(todayStart);
    cutoff.setDate(cutoff.getDate() - LOOKBACK_DAYS);
    var cutoffMs = cutoff.getTime();
    var hVals = hdata.getRange(2, 1, h_lr - 1, h_lc).getValues();
    DBG.hRows = hVals.length;
    var seriesByCode = Object.create(null);
    var todayHMap = new Map();
    var hasHDataToday = false;
    for (var r = 0; r < hVals.length; r++) {
        var row = hVals[r];
        var code = norm(row[iCode]);
        if (!code || !LTP_CODES.has(code))
            continue;
        var d = row[iDate] instanceof Date ? row[iDate] : new Date(row[iDate]);
        if (isNaN(d))
            continue;
        var ms = d.getTime();
        if (ms < cutoffMs)
            continue;
        if (!AFTER_1430 && ms >= todayStartMs)
            continue;
        var item = {
            ms: ms,
            close: toNum(row[iClose]),
            open: (iOpen >= 0) ? toNum(row[iOpen]) : NaN,
            high: (iHigh >= 0) ? toNum(row[iHigh]) : NaN,
            low: (iLow >= 0) ? toNum(row[iLow]) : NaN,
            ycp: (iYcp >= 0) ? toNum(row[iYcp]) : NaN,
            val: (iVal >= 0) ? toNum(row[iVal]) : NaN,
            vol: (iVol >= 0) ? toNum(row[iVol]) : NaN
        };
        (seriesByCode[code] || (seriesByCode[code] = [])).push(item);
        if (ms >= todayStartMs && ms <= (todayStartMs + 24 * 3600 * 1000 - 1)) {
            todayHMap.set(code, { close: item.close, val: item.val, vol: item.vol });
            hasHDataToday = true;
        }
    }
    for (var k in seriesByCode)
        seriesByCode[k].sort(function (a, b) { return a.ms - b.ms; });
    var isHolidayToday = (!isTradingDayBD_() || !hasHDataToday);
    var todaySource = ((AFTER_1430 && isTradingDayBD_()) || hasHDataToday) ? 'HData' : 'LTP';
    /***** -------- Regime (prefer DSEX sheet) -------- *****/
    var buyDateDefault = chooseBuyDate_();
    var buyStartMs = new Date(Utilities.formatDate(buyDateDefault, TZ, 'yyyy-MM-dd') + 'T00:00:00').getTime();
    var dsexSer = readDSEX_CloseSeries_(ss, cutoffMs, buyStartMs);
    var regimeOK, regimeSym, regimeSource;
    if (dsexSer) {
        var cl = dsexSer.map(function (x) { return x.close; });
        var rg = regimeFromCloses_(cl);
        regimeOK = rg.ok;
        regimeSym = rg.sym;
        regimeSource = 'DSEX sheet';
    }
    else {
        var idxSeries = seriesByCode.DSEX || [];
        var cl = idxSeries.map(function (x) { return x.close; }).filter(Number.isFinite);
        var rg = regimeFromCloses_(cl);
        regimeOK = rg.ok;
        regimeSym = rg.sym;
        regimeSource = 'HData';
    }
    /***** -------- Light stats & maps -------- *****/
    var prevYcp = new Map(), prevVal = new Map(), prevClose = new Map(), prevHigh = new Map();
    var lightStats = new Map();
    function medianOfLastN_(arr, n) {
        var a = arr.slice(-n).filter(Number.isFinite);
        if (!a.length)
            return NaN;
        return median(a);
    }
    for (var _b = 0, _c = Object.entries(seriesByCode); _b < _c.length; _b++) {
        var _d = _c[_b], code = _d[0], arr = _d[1];
        if (!arr || arr.length < MIN_VALID_DAYS)
            continue;
        var pi = -1;
        for (var k = arr.length - 1; k >= 0; k--) {
            if (arr[k].ms < buyStartMs) {
                pi = k;
                break;
            }
        }
        if (pi >= 0) {
            var p = arr[pi];
            if (Number.isFinite(p.ycp))
                prevYcp.set(code, p.ycp);
            if (Number.isFinite(p.val))
                prevVal.set(code, p.val);
            if (Number.isFinite(p.close))
                prevClose.set(code, p.close);
            if (Number.isFinite(p.high))
                prevHigh.set(code, p.high);
        }
        var usable = arr.filter(function (x) { return x.ms < buyStartMs && Number.isFinite(x.close); });
        if (usable.length < MIN_VALID_DAYS)
            continue;
        var prevC = prevClose.get(code);
        if (!Number.isFinite(prevC) || prevC <= 0)
            continue;
        var cl = arr.map(function (x) { return x.close; });
        var e20 = ema_(cl, 20);
        var last = cl[cl.length - 1];
        var priceGtEma = Number.isFinite(last) && Number.isFinite(e20[e20.length - 1]) && last > e20[e20.length - 1];
        var emaSlopeUp = linregSlopePct_(e20.slice(-10), Math.min(10, e20.length)) > 0;
        var liqMed20 = medianOfLastN_(arr.map(function (x) { return Number.isFinite(x.val) ? x.val : NaN; }), 20);
        var tN = trendPillars_(arr, buyStartMs);
        var vSpike = valueSpikeByCode.get(code) || 0;
        var addDemand = (Number.isFinite(vSpike) && vSpike > 0) ? 1 : 0;
        var baseScore = tN.score;
        var score = baseScore + addDemand;
        var sym = void 0;
        if (score >= 3)
            sym = '⬆️';
        else if (score === 0)
            sym = '⬇️';
        else
            sym = '=';
        lightStats.set(code, {
            prevClose: prevC,
            analyzedDays: usable.length,
            priceGtEma: priceGtEma,
            emaSlopeUp: emaSlopeUp,
            trendSym: sym,
            trendScore: score,
            trendFlags: Object.assign({}, tN.flags, { Demand: addDemand > 0 }),
            liqMed20: liqMed20
        });
    }
    // Existing Summary (capture Action + frozen fields + ordering)
    var existing = (function () {
        var map = new Map();
        var orderPurchased = [];
        var orderWatch = [];
        var sh = ss.getSheetByName('Summary');
        if (!sh || sh.getLastRow() < 2)
            return { map: map, orderPurchased: orderPurchased, orderWatch: orderWatch };
        var lr = sh.getLastRow(), lc = sh.getLastColumn();
        var H = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(String);
        var idxBuy = H.findIndex(function (h) { return /BUY DATE/i.test(h); });
        var idxCode = H.findIndex(function (h) { return /^CODE$/i.test(h); });
        var idxEntry = H.findIndex(function (h) { return /^ENTRY$/i.test(h); });
        var idxAction = H.findIndex(function (h) { return /^ACTION$/i.test(h); });
        var rows = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();
        rows.forEach(function (r) {
            var code = norm(r[idxCode] || '');
            if (!code)
                return;
            var action = (idxAction >= 0 ? String(r[idxAction] || '').trim() : '') || 'Listed';
            var buyTxt = (idxBuy >= 0 ? String(r[idxBuy] || '').trim() : '');
            var entry = (idxEntry >= 0 ? toNum(r[idxEntry]) : NaN);
            map.set(code, { action: action, buyDateText: buyTxt, frozenEntry: entry });
            if (action === 'Purchased')
                orderPurchased.push(code);
            else if (action === 'Watch')
                orderWatch.push(code);
        });
        return { map: map, orderPurchased: orderPurchased, orderWatch: orderWatch };
    })();
    function todaysCloseFromLTPRow_(row) {
        var ltpv = (colLtpPx != null) ? toNum(row[colLtpPx]) : NaN;
        var ycp = (colYcpPx != null) ? toNum(row[colYcpPx]) : NaN;
        if (Number.isFinite(ltpv) && ltpv > 0)
            return ltpv;
        if (Number.isFinite(ycp) && ycp > 0)
            return ycp;
        return NaN;
    }
    function lastNBefore_(series, buyStartMs2, n, picker) {
        var vals = [];
        for (var k = series.length - 1; k >= 0 && vals.length < n; k--) {
            var row = series[k];
            if (row && row.ms < buyStartMs2) {
                var v = picker(row);
                if (Number.isFinite(v))
                    vals.push(v);
            }
        }
        return vals;
    }
    function valueInc_Last1vs3Avg_(series, buyStartMs2, d0ValMn) {
        if (!Number.isFinite(d0ValMn))
            return { txt: '', num: NaN };
        var prev3 = lastNBefore_(series, buyStartMs2, 3, function (r) { return r.val; });
        if (prev3.length < 3)
            return { txt: '', num: NaN };
        var avgPrev3 = (prev3[0] + prev3[1] + prev3[2]) / 3;
        if (!Number.isFinite(avgPrev3) || avgPrev3 <= 0)
            return { txt: '', num: NaN };
        var changePct = (d0ValMn / avgPrev3 - 1) * 100;
        var num = Math.round(changePct);
        return { txt: "".concat(num, "%"), num: num };
    }
    function priceChg_TodayPlus2_(series, buyStartMs2, d0Px, isHoliday) {
        if (isHoliday) {
            var prev_1 = lastNBefore_(series, buyStartMs2, 3, function (r) { return r.close; });
            if (prev_1.length < 3)
                return { txt: '', num: NaN };
            var p1_1 = prev_1[0], p2_1 = prev_1[1], p3 = prev_1[2];
            if (!(p1_1 > 0 && p2_1 > 0 && p3 > 0))
                return { txt: '', num: NaN };
            var c1_1 = (p1_1 / p2_1 - 1) * 100;
            var c2_1 = (p2_1 / p3 - 1) * 100;
            var num_1 = Math.round((c1_1 + c2_1) / 2);
            return { txt: "".concat(num_1, "%"), num: num_1 };
        }
        if (!Number.isFinite(d0Px))
            return { txt: '', num: NaN };
        var prev = lastNBefore_(series, buyStartMs2, 2, function (r) { return r.close; });
        if (prev.length < 2)
            return { txt: '', num: NaN };
        var p1 = prev[0], p2 = prev[1];
        if (!(p1 > 0 && p2 > 0))
            return { txt: '', num: NaN };
        var c1 = (d0Px / p1 - 1) * 100;
        var c2 = (p1 / p2 - 1) * 100;
        var num = Math.round((c1 + c2) / 2);
        return { txt: "".concat(num, "%"), num: num };
    }
    /***** ===== Option B Trigger & Predicted Entry ===== *****/
    function computeTriggerB_(series, buyStartMs2) {
        var hist = series.filter(function (b) { return b.ms < buyStartMs2; });
        if (hist.length < Math.max(MIN_VALID_DAYS, ATR_LEN + 5))
            return { ok: false };
        var cl = hist.map(function (x) { return x.close; }).filter(Number.isFinite);
        var hi = hist.map(function (x) { return x.high; });
        var lo = hist.map(function (x) { return x.low; });
        if (cl.length < ATR_LEN + 5)
            return { ok: false };
        var e20 = ema_(cl, 20);
        var e20Last = e20[e20.length - 1];
        var e20SlopeUp = linregSlopePct_(e20.slice(-10), Math.min(10, e20.length)) > 0;
        var closeLast = cl[cl.length - 1];
        if (!(Number.isFinite(closeLast) && Number.isFinite(e20Last) && closeLast > e20Last && e20SlopeUp))
            return { ok: false };
        var atrArr = atr_(hi, lo, cl, ATR_LEN);
        var atrLast = atrArr[atrArr.length - 1];
        if (!Number.isFinite(atrLast))
            return { ok: false };
        var trigger = e20Last + 0.3 * atrLast;
        return { ok: true, trigger: trigger, e20Last: e20Last, atrLast: atrLast };
    }
    function predictEntryB_(series, buyStartMs2, ltpPxNow, isHolidayToday2) {
        var t = computeTriggerB_(series, buyStartMs2);
        if (!t.ok)
            return { plannedDate: null, plannedText: '', entryPx: NaN, reason: 'setup-not-ok' };
        if (!isHolidayToday2 && IN_SESSION && Number.isFinite(ltpPxNow) && ltpPxNow > 0 && ltpPxNow <= t.trigger) {
            var todayStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
            var today = new Date(todayStr + 'T00:00:00');
            return { plannedDate: today, plannedText: fmtDate(today), entryPx: +t.trigger.toFixed(2), reason: 'live-hit' };
        }
        var baseBuy = chooseBuyDate_();
        return { plannedDate: baseBuy, plannedText: fmtDate(baseBuy), entryPx: +t.trigger.toFixed(2), reason: 'planned-next' };
    }
    /***** ===== Build items (for all tickers, so pinned rows stay even if they fail) ===== *****/
    var itemsByCode = new Map();
    var itemsList = [];
    for (var _e = 0, ltpVals_3 = ltpVals; _e < ltpVals_3.length; _e++) {
        var r = ltpVals_3[_e];
        var code = norm(r[colCode]);
        if (!code)
            continue;
        var series = seriesByCode[code];
        if (!series || !series.length)
            continue;
        var s = lightStats.get(code);
        if (!s)
            continue;
        var sector = (colSector != null) ? r[colSector] : '';
        // 🚫 Exclude Mutual Fund
        if (sector && String(sector).toUpperCase().includes('MUTUAL'))
            continue;
        var d0ValMn = NaN, d0Px = NaN;
        if (((AFTER_1430 && isTradingDayBD_()) || hasHDataToday)) {
            var t = todayHMap.get(code);
            if (t) {
                d0ValMn = t.val;
                d0Px = t.close;
            }
        }
        else {
            d0ValMn = (colValMn != null) ? toNum(r[colValMn]) : NaN;
            d0Px = todaysCloseFromLTPRow_(r);
            if (colCloseP != null) {
                var cp = toNum(r[colCloseP]);
                if (cp === 0) {
                    // already using LTP/YCP
                }
            }
        }
        var _f = valueInc_Last1vs3Avg_(series, buyStartMs, d0ValMn), valueIncTxt = _f.txt, valueIncNum = _f.num;
        var _g = priceChg_TodayPlus2_(series, buyStartMs, d0Px, isHolidayToday), priceChgTxt = _g.txt, priceChgNum = _g.num;
        var useLive = (!hasHDataToday && isTradingDayBD_() && IN_SESSION);
        var predEnt = predictEntryB_(series, buyStartMs, useLive ? d0Px : NaN, isHolidayToday);
        var entry = predEnt.entryPx;
        var buyDate = predEnt.plannedDate || chooseBuyDate_();
        var buyDateText = predEnt.plannedText || fmtDate(buyDate);
        var stop_1 = Number.isFinite(entry) ? +(entry * (1 - STOP_FLOOR_PCT)).toFixed(2) : NaN;
        var item = {
            code: code,
            sector: sector,
            series: series,
            s: s,
            valueIncNum: valueIncNum,
            priceChgNum: priceChgNum,
            valueIncTxt: valueIncTxt,
            priceChgTxt: priceChgTxt,
            entry: entry,
            stop: stop_1,
            buyDateText: buyDateText,
            entryReason: predEnt.reason,
            action: 'Listed',
            targetPct: NaN,
            target: NaN,
            winPct: NaN,
            medHold: '',
            invest: 0
        };
        itemsByCode.set(code, item);
        itemsList.push(item);
    }
    DBG.builtItems = itemsList.length;
    function recalcEntryFromLTP_(predictedEntry, ltpNow, trendScore, hasDemand) {
        if (!Number.isFinite(predictedEntry) || !Number.isFinite(ltpNow))
            return predictedEntry;
        var diffPct = ((predictedEntry - ltpNow) / ltpNow) * 100;
        if (Math.abs(diffPct) <= 1.5)
            return +predictedEntry.toFixed(2);
        var bufferPct;
        if (trendScore >= 3 || hasDemand)
            bufferPct = 1.0;
        else if (trendScore >= 2)
            bufferPct = 2.0;
        else
            bufferPct = 4.0;
        return +(ltpNow * (1 - bufferPct / 100)).toFixed(2);
    }
    /***** =================== Gates (for Listed shortlist) =================== *****/
    function passGates_(it) {
        var sc = it.s.trendScore;
        var vSpike = valueSpikeByCode.get(it.code) || 0;
        var hasDemand = Number.isFinite(vSpike) && vSpike > 0;
        if (!Number.isFinite(sc) || (sc < 1 && !hasDemand))
            return false;
        if (DEF.USE_LIQUIDITY_GUARD && Number.isFinite(it.s.liqMed20) && it.s.liqMed20 < DEF.LIQ_MIN_VAL_MN_MED20)
            return false;
        if (REQUIRE_VALUE_INCREASE && Number.isFinite(it.valueIncNum) && it.valueIncNum < MIN_VALUE_INC_FILTER_PCT)
            return false;
        if (REQUIRE_PRICE_INCREASE && Number.isFinite(it.priceChgNum)) {
            if (!hasDemand && it.priceChgNum < MIN_PRICE_CHG_FILTER_PCT)
                return false;
            if (hasDemand && it.priceChgNum < (MIN_PRICE_CHG_FILTER_PCT - 2))
                return false;
        }
        if (!Number.isFinite(it.entry))
            return false;
        var currentLtp = ltpPriceByCode.get(it.code);
        if (Number.isFinite(currentLtp)) {
            it.entry = recalcEntryFromLTP_(it.entry, currentLtp, sc, hasDemand);
        }
        if (!Number.isFinite(it.entry))
            return false;
        var ENTRY_TOLERANCE_PCT = 3;
        if (Number.isFinite(currentLtp)) {
            var maxAllowed = it.entry * (1 + ENTRY_TOLERANCE_PCT / 100);
            if (currentLtp > maxAllowed)
                return false;
        }
        return true;
    }
    var pool = itemsList.filter(passGates_);
    DBG.afterGate = pool.length;
    /***** =================== PIN / SECTION logic (Purchased & Watch) =================== *****/
    var frozenPurchased = [];
    var frozenWatch = [];
    for (var _h = 0, _j = existing.map.entries(); _h < _j.length; _h++) {
        var _k = _j[_h], code = _k[0], exInfo = _k[1];
        var base = itemsByCode.get(code);
        if (!base)
            continue;
        var it = base;
        it.action = exInfo.action;
        if (exInfo.action === 'Purchased') {
            var freezeBuy = exInfo.buyDateText && exInfo.buyDateText.trim() ? exInfo.buyDateText : base.buyDateText;
            var freezeEntry = Number.isFinite(exInfo.frozenEntry) ? exInfo.frozenEntry : base.entry;
            it.buyDateText = freezeBuy;
            it.entry = freezeEntry;
            frozenPurchased.push(it);
        }
        else if (exInfo.action === 'Watch') {
            it.buyDateText = base.buyDateText;
            frozenWatch.push(it);
        }
    }
    var orderedPurchased = existing.orderPurchased.map(function (c) { return frozenPurchased.find(function (x) { return x.code === c; }); }).filter(Boolean);
    var orderedWatch = existing.orderWatch.map(function (c) { return frozenWatch.find(function (x) { return x.code === c; }); }).filter(Boolean);
    var pinnedCodes = new Set(__spreadArray(__spreadArray([], orderedPurchased.map(function (x) { return x.code; }), true), orderedWatch.map(function (x) { return x.code; }), true));
    var listed = pool.filter(function (x) { return !pinnedCodes.has(x.code); }).map(function (x) {
        x.action = 'Listed';
        return x;
    });
    listed.sort(function (a, b) {
        var sa = Number.isFinite(a.s && a.s.trendScore) ? a.s.trendScore : -Infinity;
        var sb = Number.isFinite(b.s && b.s.trendScore) ? b.s.trendScore : -Infinity;
        if (sb !== sa)
            return sb - sa;
        var wa = Number.isFinite(a.winPct) ? a.winPct : -Infinity;
        var wb = Number.isFinite(b.winPct) ? b.winPct : -Infinity;
        if (wb !== wa)
            return wb - wa;
        return String(a.code).localeCompare(String(b.code));
    });
    var merged = __spreadArray(__spreadArray(__spreadArray([], orderedPurchased, true), orderedWatch, true), listed.slice(0, MAX_CANDIDATES), true);
    /***** =================== Targets / Stops / Win% / Investment =================== *****/
    function nextPeakReturns_(arr, H) {
        var out = [];
        for (var i = 0; i < arr.length - 1; i++) {
            var e = arr[i].close;
            if (!Number.isFinite(e) || e <= 0)
                continue;
            var end = Math.min(arr.length - 1, i + H);
            var peak = -Infinity;
            for (var j = i + 1; j <= end; j++) {
                var c = arr[j].close;
                if (Number.isFinite(c) && c > peak)
                    peak = c;
            }
            if (peak > -Infinity) {
                var rt = (peak / e) - 1;
                if (Number.isFinite(rt))
                    out.push(rt);
            }
        }
        return out;
    }
    function quantile_(a, q) {
        var b = a.filter(Number.isFinite).slice().sort(function (x, y) { return x - y; });
        if (!b.length)
            return NaN;
        var pos = (b.length - 1) * q;
        var lo = Math.floor(pos);
        var hi = Math.ceil(pos);
        if (lo === hi)
            return b[lo];
        return b[lo] + (b[hi] - b[lo]) * (pos - lo);
    }
    function clamp(x, lo, hi) {
        return Math.max(lo, Math.min(hi, x));
    }
    function predictTargetPct_Achievable_(arr, entryPx) {
        var ret = nextPeakReturns_(arr, MAX_HOLD_DAYS);
        if (ret.length >= MIN_SAMPLES) {
            var tPct = quantile_(ret, 1 - DESIRED_HIT_PROB);
            if (!Number.isFinite(tPct))
                tPct = MIN_TARGET_PCT;
            return clamp(tPct, MIN_TARGET_PCT, MAX_TARGET_PCT);
        }
        var hi = arr.map(function (x) { return x.high; }), lo = arr.map(function (x) { return x.low; }), cl = arr.map(function (x) { return x.close; });
        var a = atr_(hi, lo, cl, ATR_LEN);
        var atrLast = a[a.length - 1];
        var atrPct = (Number.isFinite(atrLast) && Number.isFinite(cl[cl.length - 1]) && cl[cl.length - 1] > 0) ? (atrLast / cl[cl.length - 1]) : NaN;
        if (Number.isFinite(atrPct) && atrPct > 0)
            return clamp(atrPct * TARGET_FALLBACK_ATR_MULT, MIN_TARGET_PCT, MAX_TARGET_PCT);
        var last10 = arr.slice(-10).map(function (x) { return x.close; }).filter(Number.isFinite);
        if (last10.length >= 5) {
            var minC = Math.min.apply(null, last10);
            var maxC = Math.max.apply(null, last10);
            var swingPct = (maxC - minC) / (entryPx || maxC);
            return clamp(swingPct * 0.7, MIN_TARGET_PCT, MAX_TARGET_PCT);
        }
        return MIN_TARGET_PCT;
    }
    function computeWinAndHold_(arr, tgtPct) {
        var wins = 0, totals = 0;
        var holdWins = [];
        for (var i = 0; i < arr.length - 1; i++) {
            var e = arr[i].close;
            if (!Number.isFinite(e) || e <= 0)
                continue;
            var tgt = e * (1 + tgtPct);
            var end = Math.min(arr.length - 1, i + MAX_HOLD_DAYS);
            var hit = false;
            var hold = NaN;
            for (var j = i + 1; j <= end; j++) {
                var c = arr[j].close;
                if (Number.isFinite(c) && c >= tgt) {
                    hit = true;
                    hold = j - i;
                    break;
                }
            }
            totals++;
            if (hit) {
                wins++;
                holdWins.push(hold);
            }
        }
        var rawWin = totals ? (wins * 100 / totals) : 0;
        var winPct = (((rawWin / 100) * arr.length + GLOBAL_WIN_MEAN * PRIOR_WEIGHT) / (arr.length + PRIOR_WEIGHT) * 100);
        var medHold = holdWins.length ? Math.min(Math.round(median(holdWins)), MAX_HOLD_DAYS) : '';
        return { winPct: winPct, medHold: medHold };
    }
    for (var _l = 0, merged_1 = merged; _l < merged_1.length; _l++) {
        var it = merged_1[_l];
        var arr = it.series;
        var entryPx = it.entry;
        if (!Number.isFinite(entryPx))
            continue;
        var hi = arr.map(function (x) { return x.high; }), lo = arr.map(function (x) { return x.low; }), cl = arr.map(function (x) { return x.close; });
        var a = atr_(hi, lo, cl, ATR_LEN), atrLast = a[a.length - 1];
        var atrPct = (Number.isFinite(atrLast) && Number.isFinite(cl[cl.length - 1]) && cl[cl.length - 1] > 0) ? (atrLast / cl[cl.length - 1]) * 100 : NaN;
        if (Number.isFinite(atrPct)) {
            var atrAbs = entryPx * (atrPct / 100);
            var atrStop = entryPx - ATR_STOP_MULT * atrAbs;
            it.stop = +Math.min(+it.stop || Infinity, atrStop).toFixed(2);
        }
        var tPct = predictTargetPct_Achievable_(arr, entryPx);
        var _m = computeWinAndHold_(arr, tPct), winPct = _m.winPct, medHold = _m.medHold;
        it.targetPct = tPct;
        it.target = +(entryPx * (1 + tPct)).toFixed(2);
        it.winPct = +(+winPct).toFixed(1);
        it.medHold = medHold;
        it.invest = investAmountFromWin_(it.winPct);
    }
    /***** ======================= WRITE ======================= *****/
    var finalRows = merged.map(function (it) {
        var buyDate = it.buyDateText || fmtDate(chooseBuyDate_());
        var exitDate = (function () {
            var d = new Date(buyDate);
            var base = isNaN(d) ? chooseBuyDate_() : d;
            return addTradingDaysBD_(base, it.medHold || MAX_HOLD_DAYS);
        })();
        var dsexSym = displayTrendSym_(regimeSym);
        var stockSym = displayTrendSym_(it.s.trendSym);
        var trendText = "'" + dsexSym + ' | ' + stockSym;
        return [
            buyDate,
            fmtDate(exitDate),
            it.code,
            it.sector,
            ltpPriceByCode.get(it.code) || '',
            it.entry,
            it.stop,
            it.target,
            "".concat(Math.round((it.targetPct || 0) * 100), "%"),
            (it.medHold || ''),
            Number.isFinite(it.winPct) ? "".concat(Math.round(it.winPct), "%") : '',
            it.valueIncTxt,
            it.priceChgTxt,
            trendText,
            'OK',
            it.action
        ];
    });
    DBG.pinnedPlusKept = finalRows.length;
    writeSummary_(finalRows, ACTION_LIST);
    writeLog_(Object.assign({
        note: 'Entry=EMA20+0.3*ATR; live LTP 10:00–14:30 used when HData isn’t ready. Action: Purchased (freeze Buy Date & Entry), Watch (pinned, rolling date). Sections: Purchased → Watch → Listed. Investment=score-weighted by Win%. Trend has 4 pillars: Location, Momentum, Structure, Demand(ValueSpike>0). DSEX regime included in Trend column as DSEX | Stock.',
        tz: TZ,
        time: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
        todaySource: todaySource,
        shortlisted: finalRows.length,
        regimeOK: regimeOK ? 'OK' : 'Caution',
        regimeSym: regimeSym,
        regimeSource: regimeSource
    }, DBG));
    function applyThinBordersToUsedRange_(sh, headersLen) {
        var lastR = sh.getLastRow();
        var lastC = Math.min(sh.getMaxColumns(), headersLen);
        if (lastR > 0 && lastC > 0)
            sh.getRange(1, 1, lastR, lastC).setBorder(true, true, true, true, true, true);
    }
    function writeSummary_(rows, actionList) {
        var headers = [
            'Buy Date', 'Exit Date', 'CODE', 'SECTOR', 'Investment',
            'ENTRY', 'STOP', 'TARGET', 'Target %', 'Hold Days',
            'Win %', 'Value Increase %', 'Price Change %', 'Trend',
            'Trend Check', 'Action'
        ];
        var sh = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
        if (sh.getMaxColumns() < headers.length)
            sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        var lastRow = sh.getLastRow();
        if (lastRow > 1)
            sh.getRange(2, 1, lastRow - 1, headers.length).clearContent();
        if (rows && rows.length)
            sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
        if (rows && rows.length) {
            var rule = SpreadsheetApp.newDataValidation().requireValueInList(actionList, true).setAllowInvalid(false).build();
            sh.getRange(2, headers.length, rows.length, 1).setDataValidation(rule);
        }
        applyThinBordersToUsedRange_(sh, headers.length);
    }
    function writeLog_(obj) {
        var sh = ss.getSheetByName('SummaryLog') || ss.insertSheet('SummaryLog');
        var lr2 = sh.getLastRow(), lc2 = sh.getLastColumn();
        if (lr2 && lc2)
            sh.getRange(1, 1, lr2, lc2).clearContent();
        var rows = Object.entries(obj).map(function (_a) {
            var k = _a[0], v = _a[1];
            return [k, String(v)];
        });
        if (rows.length)
            sh.getRange(1, 1, rows.length, 2).setValues(rows);
        applyThinBordersToUsedRange_(sh, 2);
    }
    function safeWrite_(rows, actionList, note) {
        writeSummary_(rows || [], actionList);
        writeLog_({ note: note, tz: TZ, time: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), shortlisted: 0 });
    }
}
/** UI menu */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Summary Utils')
        .addItem('Refresh Summary', 'refreshSummary')
        .addToUi();
}
