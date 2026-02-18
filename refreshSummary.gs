function refreshSummary() {
  /***** ======================= CONFIG ======================= *****/
  const LTP_SHEET   = (typeof LTP_SHEET_NAME !== 'undefined') ? LTP_SHEET_NAME : 'LTP';
  const HDATA_SHEET = (typeof SRC_SHEET_NAME !== 'undefined') ? SRC_SHEET_NAME : 'HData';
  const DSEX_SHEET  = 'DSEX';

  const TZ = 'Asia/Dhaka';
  const LOOKBACK_DAYS = (typeof this.LOOKBACK_DAYS !== 'undefined') ? this.LOOKBACK_DAYS : 180;
  const MAX_HOLD_DAYS = 7;

  // Stops and bounds
  const STOP_FLOOR_PCT   = (typeof this.STOP_FLOOR_PCT !== 'undefined') ? this.STOP_FLOOR_PCT : 0.02;
  const ATR_LEN          = 14;
  const ATR_STOP_MULT    = 1.8;
  const MIN_TARGET_PCT   = 0.10;
  const MAX_TARGET_PCT   = 0.20;
  const MIN_SAMPLES      = 30;

  // Achievability settings
  const DESIRED_HIT_PROB = 0.60;
  const TARGET_FALLBACK_ATR_MULT = 1.6;

  // Universe & gating
  const ACTION_LIST      = ['Purchased','Watch','Listed'];
  const MAX_CANDIDATES   = 40;
  const MIN_VALID_DAYS   = 60;

  // Filters
  const REQUIRE_PRICE_INCREASE    = true;
  const MIN_PRICE_CHG_FILTER_PCT  = 3;

  const REQUIRE_VALUE_INCREASE    = true;
  const MIN_VALUE_INC_FILTER_PCT  = 10;

  // Optional guards (OFF)
  const DEF = {
    USE_LIQUIDITY_GUARD : false,
    LIQ_MIN_VAL_MN_MED20: 10,
    ATR_BAND_MIN        : 1.2,
    ATR_BAND_MAX        : 8.0,
  };

  // Bayesian prior for Win%
  const GLOBAL_WIN_MEAN = 0.45;
  const PRIOR_WEIGHT    = 50;

  // ----- Investment settings (score-weighted) -----
  const TOTAL_CAPITAL_BDT  = 100000;
  const NUM_SLOTS          = 10;
  const BASE_PER_TRADE     = TOTAL_CAPITAL_BDT / NUM_SLOTS;
  const WIN_BASE_PCT       = 60;
  const WEIGHT_MIN         = 0.5;
  const WEIGHT_MAX         = 2.0;
  function investAmountFromWin_(winPct){
    if (!Number.isFinite(winPct) || winPct <= 0) return 0;
    const w = Math.min(WEIGHT_MAX, Math.max(WEIGHT_MIN, (winPct / WIN_BASE_PCT)));
    return Math.round((BASE_PER_TRADE * w) / 100) * 100;
  }

  const DBG = { ltpRows:0, hRows:0, builtItems:0, symbolsInSeries:0, afterGate:0,
                metricBlankRemoved:0, positivityRemoved:0, dedupedKept:0, pinnedPlusKept:0 };

  /***** ======================= UTILS ======================= *****/
  const toNum   = v => (typeof v==='number') ? v : (v==null||v==='') ? NaN : (parseFloat(String(v).replace(/[, ]/g,'')));
  const fmtDate = d => Utilities.formatDate(d, TZ, 'dd-MMM-yy');
  const norm    = s => String(s||'').toUpperCase().replace(/[^A-Z0-9]/g,'').trim();
  const median  = a => { const b=a.filter(Number.isFinite).slice().sort((x,y)=>x-y); if(!b.length)return NaN; const m=Math.floor(b.length/2); return (b.length%2)? b[m] : (b[m-1]+b[m])/2; };

  const hhmm = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  const minutesNow = (+hhmm.slice(0,2))*60 + (+hhmm.slice(3,5));

  const BLOCK_START = 9 * 60;
  const BLOCK_END   = 10 * 60;
  if (minutesNow >= BLOCK_START && minutesNow < BLOCK_END) {
    Logger.log('refreshSummary skipped between 09:00–10:00 (pre-data time).');
    return;
  }

  const AFTER_1430 = minutesNow >= (14*60+30);
  const IN_SESSION = minutesNow >= (10*60) && minutesNow < (14*60+30);

  function isTradingDayBD_(d = new Date()){
    const ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    const x  = new Date(ds+'T00:00:00'); const wd = x.getDay();
    return wd !== 5 && wd !== 6;
  }
  function addTradingDaysBD_(date, n){
    const d=new Date(date); d.setHours(0,0,0,0);
    let add=0; while(add<n){ d.setDate(d.getDate()+1); const wd=d.getDay(); if (wd!==5 && wd!==6) add++; }
    return d;
  }
  function chooseBuyDate_(){
    const todayStr = Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd');
    const today = new Date(todayStr+'T00:00:00'); const wd=today.getDay();
    if (!AFTER_1430 && wd!==5 && wd!==6) return today;
    return addTradingDaysBD_(today,1);
  }

  function ema_(arr,len){
    if(!arr||!arr.length) return [];
    const k=2/(len+1), out=new Array(arr.length); let e=arr[0];
    for(let i=0;i<arr.length;i++){ const v=Number.isFinite(arr[i])?arr[i]:e; e=(i? (v-e)*k + e : v); out[i]=e; }
    return out;
  }
  function macdHist_(cl){
    if(!cl||cl.length<26) return {hist:NaN, hist3:[]};
    const e12=ema_(cl,12), e26=ema_(cl,26);
    const macd=cl.map((_,i)=>e12[i]-e26[i]), sig=ema_(macd,9);
    const hist = macd[cl.length-1]-sig[sig.length-1];
    const hist3 = macd.slice(-3).map((_,k)=>macd[cl.length-3+k]-sig[sig.length-3+k]);
    return {hist, hist3};
  }
  function linregSlopePct_(arr, days){
    if(!arr||arr.length<days) return 0;
    const s=arr.slice(-days), n=s.length;
    let sx=0,sy=0,sxx=0,sxy=0;
    for(let i=0;i<n;i++){ const x=i+1,y=s[i]; sx+=x; sy+=y; sxx+=x*x; sxy+=x*y; }
    const m=(n*sxy - sx*sy)/(n*sxx - sx*sx), last=s[n-1];
    return (last? (m/last)*100 : 0);
  }
  function atr_(hi,lo,cl,len){
    const n=cl.length, tr=new Array(n).fill(NaN);
    for(let i=1;i<n;i++){
      if(!Number.isFinite(hi[i])||!Number.isFinite(lo[i])||!Number.isFinite(cl[i-1])) continue;
      tr[i]=Math.max(hi[i]-lo[i], Math.abs(hi[i]-cl[i-1]), Math.abs(lo[i]-cl[i-1]));
    }
    const out=new Array(n).fill(NaN); let sum=0;
    for(let i=0;i<n;i++){
      const v=tr[i];
      if(!Number.isFinite(v)){ out[i]=(i?out[i-1]:NaN); continue; }
      if(i<=len){ sum+=v; out[i]=(i===len)? (sum/len) : NaN; }
      else out[i]=(out[i-1]*(len-1)+v)/len;
    }
    return out;
  }
  function rsi_(cl, len=14){
    if(!cl||cl.length<len+1) return NaN;
    let g=0,l=0;
    for(let i=1;i<=len;i++){ const d=cl[i]-cl[i-1]; if(d>0) g+=d; else l-=Math.min(d,0); }
    g/=len; l/=len;
    let rs=(l===0)?100:g/l, r=100-(100/(1+rs));
    for(let i=len+1;i<cl.length;i++){
      const d=cl[i]-cl[i-1], G=Math.max(d,0), L=Math.max(-d,0);
      g=(g*(len-1)+G)/len; l=(l*(len-1)+L)/len;
      rs=(l===0)?100:g/l; r=100-(100/(1+rs));
    }
    return r;
  }

  /***** ======================= TREND (3-pillars base) ======================= *****/
  function trendPillars_(series, buyStartMs){
    const hist = series.filter(b => b.ms < buyStartMs);
    const N = 10; if (hist.length < Math.max(N, 20)) return {sym:'=', score:0, flags:{insufficient:true}};
    const lastN = hist.slice(-N); const clAll = hist.map(x=>x.close).filter(Number.isFinite);

    const e20 = ema_(clAll, 20);
    const lastClose = clAll[clAll.length-1], lastE20 = e20[e20.length-1];
    const loc = (Number.isFinite(lastClose) && Number.isFinite(lastE20) && lastClose > lastE20) &&
                (linregSlopePct_(e20.slice(-10), Math.min(10, e20.length)) > 0);

    const {hist:mh, hist3} = macdHist_(clAll);
    const macdUp    = Number.isFinite(mh) && mh > 0;
    const macdRise  = (hist3 && hist3.length===3) ? (hist3[0] <= hist3[1] && hist3[1] <= hist3[2]) : false;
    const mom       = (macdUp && macdRise) || (rsi_(clAll,14) >= 55);

    let HH=0, HL=0;
    for (let i=1;i<lastN.length;i++){
      const p=lastN[i-1], c=lastN[i];
      if (Number.isFinite(c.high)&&Number.isFinite(p.high) && c.high>p.high) HH++;
      if (Number.isFinite(c.low) &&Number.isFinite(p.low)  && c.low >p.low ) HL++;
    }
    const struct = (HH>=4 && HL>=4) || (linregSlopePct_(lastN.map(x=>x.close), Math.min(7,lastN.length)) > 0);

    const pass = (loc?1:0)+(mom?1:0)+(struct?1:0);
    return {sym: pass>=2?'⬆️':(pass===0?'⬇️':'='), score:pass, flags:{Location:loc,Momentum:mom,Structure:struct}};
  }

  /***** ======================= DSEX HELPERS ======================= *****/
  function readDSEX_CloseSeries_(ss, cutoffMs, buyStartMs){
    const sh = ss.getSheetByName(DSEX_SHEET); if (!sh) return null;
    const lc = sh.getLastColumn(), lr = sh.getLastRow(); if (lr < 2) return null;
    const H = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(h => String(h).trim().toUpperCase());
    const iDate = H.findIndex(h => h === 'DATE'); const iDSEX = H.findIndex(h => h.includes('DSEX INDEX'));
    if (iDate < 0 || iDSEX < 0) return null;

    const vals = sh.getRange(2,1,lr-1,lc).getValues(); const ser = [];
    for (const r of vals){
      const d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]); if (isNaN(d)) continue;
      const ms = d.getTime(); if (ms < cutoffMs || ms >= buyStartMs) continue;
      const c = Number(r[iDSEX]); if (Number.isFinite(c)) ser.push({ ms, close: c });
    }
    ser.sort((a,b)=>a.ms-b.ms); return ser.length ? ser : null;
  }

  function regimeFromCloses_(cl){
    if (!cl || cl.length < 40) return { ok: true, sym: '=' };

    const slope20 = linregSlopePct_(cl, 20);
    const { hist } = macdHist_(cl);
    const macd = Number.isFinite(hist) ? hist : 0;

    const strongUp = (slope20 > 0.05) || (macd > 0.05);
    const neutralOk = (slope20 > -0.10) || (macd > -0.05);

    let ok, sym;
    if (strongUp) {
      ok  = true;
      sym = '⬆️';
    } else if (neutralOk) {
      ok  = true;
      sym = '=';
    } else {
      ok  = false;
      sym = '⬇️';
    }

    return { ok, sym };
  }

  /***** ======================= SHEETS / HEADERS ======================= *****/
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ltp = ss.getSheetByName(LTP_SHEET), hdata = ss.getSheetByName(HDATA_SHEET);
  if (!ltp || !hdata) return safeWrite_([], ACTION_LIST, 'Missing LTP/HData');

  const lr = ltp.getLastRow(), lc = ltp.getLastColumn();
  if (lr < 2) return safeWrite_([], ACTION_LIST, 'No rows in LTP');

  const HRowDisp = ltp.getRange(1,1,1,lc).getDisplayValues()[0];
  const NU = HRowDisp.map(h=>norm(String(h).replace(/\*/g,'')));
  const findHeader = (keys) => {
    const arr = Array.isArray(keys)? keys : [keys];
    for (let i=0;i<NU.length;i++){ for(const k of arr){ if (NU[i].includes(norm(k))) return i; } }
    return null;
  };

  const colCode     = findHeader(['TRADINGCODE']);
  const colSector   = findHeader(['SECTOR']);
  const colValMn    = findHeader(['VALUEMN','VALUE']);
  const colLtpPx    = findHeader(['LTP']);
  const colYcpPx    = findHeader(['YCP']);
  const colCloseP   = findHeader(['CLOSEP']);
  const colValSpike = findHeader(['VALUESPIKE']);

  if (colCode == null) return safeWrite_([], ACTION_LIST, 'Missing TRADING CODE in LTP');

  const MIN_LTP_E_FILTER = 5.00;
  const ltpColEIndex = 4; // Column E (1-indexed) in LTP sheet
  function ltpFromColumnE_(row){
    const v = (colLtpPx != null) ? toNum(row[colLtpPx]) : toNum(row[ltpColEIndex]);
    return Number.isFinite(v) ? v : NaN;
  }

  const ltpVals = ltp.getRange(2,1,lr-1,lc).getValues(); DBG.ltpRows = ltpVals.length;
  const LTP_CODES = new Set(ltpVals.map(r => norm(r[colCode])).filter(Boolean));

  const valueSpikeByCode = new Map();
  for (const r of ltpVals) {
    const code = norm(r[colCode]);
    if (!code) continue;
    if (colValSpike != null) {
      const vs = toNum(r[colValSpike]);
      if (Number.isFinite(vs)) valueSpikeByCode.set(code, vs);
    }
  }

  const h_lc = hdata.getLastColumn();
  const Hh = hdata.getRange(1,1,1,h_lc).getDisplayValues()[0].map(h => norm(String(h).replace(/\*/g,'')));
  const iDate = Hh.indexOf('DATE') >= 0 ? Hh.indexOf('DATE') : 1;
  const iCode = Hh.findIndex(x=>x.includes('TRADINGCODE')) >= 0 ? Hh.findIndex(x=>x.includes('TRADINGCODE')) : 2;
  const iHigh = Hh.indexOf('HIGH');
  const iLow  = Hh.indexOf('LOW');
  const iOpen = Hh.findIndex(x=>x.startsWith('OPEN'));
  const iClose= Hh.findIndex(x=>x.startsWith('CLOSE'));
  const iYcp  = Hh.findIndex(x=>x.includes('YCP'));
  const iVal  = Hh.findIndex(x=>x==='VALUEMN' || x==='VALUE' || x.includes('VALUE'));
  const iVol  = Hh.findIndex(x=>x==='VOLUME' || x.includes('VOLUME'));
  if ([iDate, iCode, iClose].some(i=>i<0)) return safeWrite_([], ACTION_LIST, 'Missing Date/Code/Close in HData');

  const h_lr = hdata.getLastRow(); if (h_lr < 2) return safeWrite_([], ACTION_LIST, 'HData empty');

  const todayISO = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const todayStart = new Date(todayISO+'T00:00:00'); const todayStartMs = todayStart.getTime();
  const cutoff = new Date(todayStart); cutoff.setDate(cutoff.getDate() - LOOKBACK_DAYS); const cutoffMs = cutoff.getTime();
  const hVals = hdata.getRange(2,1,h_lr-1,h_lc).getValues(); DBG.hRows = hVals.length;

  const seriesByCode = Object.create(null);
  const todayHMap = new Map(); let hasHDataToday = false;

  for (let r=0; r<hVals.length; r++){
    const row = hVals[r];
    const code = norm(row[iCode]); if (!code || !LTP_CODES.has(code)) continue;
    const d    = row[iDate] instanceof Date ? row[iDate] : new Date(row[iDate]); if (isNaN(d)) continue;
    const ms   = d.getTime(); if (ms < cutoffMs) continue; if (!AFTER_1430 && ms >= todayStartMs) continue;
    const item = {
      ms,
      close: toNum(row[iClose]),
      open : (iOpen>=0)? toNum(row[iOpen]) : NaN,
      high : (iHigh>=0)? toNum(row[iHigh]) : NaN,
      low  : (iLow >=0)? toNum(row[iLow])  : NaN,
      ycp  : (iYcp >=0)? toNum(row[iYcp])  : NaN,
      val  : (iVal >=0)? toNum(row[iVal])  : NaN,
      vol  : (iVol >=0)? toNum(row[iVol])  : NaN,
    };
    (seriesByCode[code] || (seriesByCode[code]=[])).push(item);
    if (ms >= todayStartMs && ms <= (todayStartMs + 24*3600*1000 - 1)) {
      todayHMap.set(code, {close: item.close, val: item.val, vol: item.vol});
      hasHDataToday = true;
    }
  }
  for (const k in seriesByCode) seriesByCode[k].sort((a,b)=>a.ms-b.ms);

  // Missing same-day HData during market hours should fall back to LTP,
  // not be treated as a market holiday.
  const isHolidayToday = !isTradingDayBD_();
  const todaySource    = ((AFTER_1430 && isTradingDayBD_()) || hasHDataToday) ? 'HData' : 'LTP';

  const buyDateDefault  = chooseBuyDate_();
  const buyStartMs      = new Date(Utilities.formatDate(buyDateDefault, TZ, 'yyyy-MM-dd') + 'T00:00:00').getTime();
  const dsexSer = readDSEX_CloseSeries_(ss, cutoffMs, buyStartMs);
  let regimeOK, regimeSym, regimeSource;
  if (dsexSer) {
    const cl = dsexSer.map(x=>x.close);
    const rg = regimeFromCloses_(cl);
    regimeOK=rg.ok; regimeSym=rg.sym; regimeSource='DSEX sheet';
  } else {
    const idxSeries = seriesByCode['DSEX'] || [];
    const cl = idxSeries.map(x=>x.close).filter(Number.isFinite);
    const rg = regimeFromCloses_(cl);
    regimeOK=rg.ok; regimeSym=rg.sym; regimeSource='HData';
  }

  const prevYcp = new Map(), prevVal = new Map(), prevClose = new Map(), prevHigh = new Map();
  const lightStats = new Map();
  function medianOfLastN_(arr, n){ const a=arr.slice(-n).filter(Number.isFinite); if (!a.length) return NaN; return median(a); }

  for (const [code, arr] of Object.entries(seriesByCode)) {
    if (!arr || arr.length < MIN_VALID_DAYS) continue;
    let pi = -1; for (let k=arr.length-1;k>=0;k--){ if (arr[k].ms < buyStartMs){ pi=k; break; } }
    if (pi >= 0) {
      const p=arr[pi];
      if (Number.isFinite(p.ycp))   prevYcp.set(code,p.ycp);
      if (Number.isFinite(p.val))   prevVal.set(code,p.val);
      if (Number.isFinite(p.close)) prevClose.set(code,p.close);
      if (Number.isFinite(p.high))  prevHigh.set(code,p.high);
    }
    const usable = arr.filter(x => x.ms < buyStartMs && Number.isFinite(x.close));
    if (usable.length < MIN_VALID_DAYS) continue;
    const prevC = prevClose.get(code); if (!Number.isFinite(prevC) || prevC <= 0) continue;
    const cl=arr.map(x=>x.close); const e20=ema_(cl,20); const last=cl[cl.length-1];
    const priceGtEma = Number.isFinite(last)&&Number.isFinite(e20[e20.length-1])&&last>e20[e20.length-1];
    const emaSlopeUp = linregSlopePct_(e20.slice(-10), Math.min(10,e20.length)) > 0;
    const liqMed20 = medianOfLastN_(arr.map(x=>Number.isFinite(x.val)?x.val:NaN), 20);
    const tN = trendPillars_(arr, buyStartMs);

    const vSpike = valueSpikeByCode.get(code) || 0;
    const addDemand = (Number.isFinite(vSpike) && vSpike > 0) ? 1 : 0;

    const baseScore = tN.score;
    const score = baseScore + addDemand;

    let sym;
    if (score >= 3) sym = '⬆️';
    else if (score === 0) sym = '⬇️';
    else sym = '=';

    lightStats.set(code, {
      prevClose: prevC,
      analyzedDays: usable.length,
      priceGtEma,
      emaSlopeUp,
      trendSym: sym,
      trendScore: score,
      trendFlags: Object.assign({}, tN.flags, { Demand: addDemand > 0 }),
      liqMed20
    });
  }

  const existing = (function(){
    const map = new Map(); const orderPurchased=[], orderWatch=[];
    const sh = ss.getSheetByName('Summary'); if (!sh || sh.getLastRow()<2) return {map, orderPurchased, orderWatch};
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    const H = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(String);
    const idxBuy    = H.findIndex(h=>/BUY DATE/i.test(h));
    const idxCode   = H.findIndex(h=>/^CODE$/i.test(h));
    const idxEntry  = H.findIndex(h=>/^ENTRY$/i.test(h));
    const idxAction = H.findIndex(h=>/^ACTION$/i.test(h));
    const rows = sh.getRange(2,1,lr-1,lc).getDisplayValues();
    rows.forEach((r)=>{
      const code = norm(r[idxCode]||''); if (!code) return;
      const action = (idxAction>=0? String(r[idxAction]||'').trim() : '') || 'Listed';
      const buyTxt = (idxBuy>=0? String(r[idxBuy]||'').trim() : '');
      const entry  = (idxEntry>=0? toNum(r[idxEntry]) : NaN);
      map.set(code, {action, buyDateText: buyTxt, frozenEntry: entry});
      if (action==='Purchased') orderPurchased.push(code);
      else if (action==='Watch') orderWatch.push(code);
    });
    return {map, orderPurchased, orderWatch};
  })();

  function todaysCloseFromLTPRow_(row) {
    const ltpv = (colLtpPx!=null) ? toNum(row[colLtpPx]) : NaN;
    const ycp  = (colYcpPx!=null) ? toNum(row[colYcpPx]) : NaN;
    if (Number.isFinite(ltpv) && ltpv > 0) return ltpv;
    if (Number.isFinite(ycp)  && ycp  > 0) return ycp;
    return NaN;
  }
  function lastNBefore_(series, buyStartMs, n, picker){
    const vals=[]; for(let k=series.length-1;k>=0 && vals.length<n;k--){ const row=series[k]; if (row && row.ms < buyStartMs){ const v = picker(row); if (Number.isFinite(v)) vals.push(v); } } return vals;
  }
  function valueInc_Last1vs3Avg_(series, buyStartMs, d0ValMn){
    if (!Number.isFinite(d0ValMn)) return {txt:'', num:NaN};
    const prev3 = lastNBefore_(series, buyStartMs, 3, r => r.val);
    if (prev3.length < 3) return {txt:'', num:NaN};
    const avgPrev3 = (prev3[0]+prev3[1]+prev3[2]) / 3; if (!Number.isFinite(avgPrev3) || avgPrev3 <= 0) return {txt:'', num:NaN};
    const changePct = (d0ValMn / avgPrev3 - 1) * 100; const num = Math.round(changePct);
    return {txt: num + '%', num};
  }
  function priceChg_TodayPlus2_(series, buyStartMs, d0Px, isHoliday) {
    if (isHoliday) {
      const prev = lastNBefore_(series, buyStartMs, 3, r => r.close);
      if (prev.length < 3) return {txt:'', num:NaN};
      const [p1, p2, p3] = prev; if (!(p1>0 && p2>0 && p3>0)) return {txt:'', num:NaN};
      const c1 = (p1/p2 - 1) * 100; const c2 = (p2/p3 - 1) * 100; const num = Math.round((c1 + c2) / 2);
      return {txt: num + '%', num};
    } else {
      if (!Number.isFinite(d0Px)) return {txt:'', num:NaN};
      const prev = lastNBefore_(series, buyStartMs, 2, r=>r.close);
      if (prev.length < 2) return {txt:'', num:NaN};
      const [p1, p2] = prev; if (!(p1>0 && p2>0)) return {txt:'', num:NaN};
      const c1 = (d0Px/p1 - 1) * 100; const c2 = (p1/p2  - 1) * 100; const num = Math.round((c1 + c2) / 2);
      return {txt: num + '%', num};
    }
  }

  function computeTriggerB_(series, buyStartMs){
    const hist = series.filter(b => b.ms < buyStartMs);
    if (hist.length < Math.max(MIN_VALID_DAYS, ATR_LEN+5)) return { ok:false };
    const cl = hist.map(x=>x.close).filter(Number.isFinite);
    const hi = hist.map(x=>x.high); const lo = hist.map(x=>x.low);
    if (cl.length < ATR_LEN+5) return { ok:false };
    const e20 = ema_(cl, 20); const e20Last = e20[e20.length-1];
    const e20SlopeUp = linregSlopePct_(e20.slice(-10), Math.min(10,e20.length)) > 0;
    const closeLast = cl[cl.length-1];
    if (!(Number.isFinite(closeLast) && Number.isFinite(e20Last) && closeLast>e20Last && e20SlopeUp)) return { ok:false };
    const atrArr = atr_(hi, lo, cl, ATR_LEN); const atrLast = atrArr[atrArr.length-1];
    if (!Number.isFinite(atrLast)) return { ok:false };
    const trigger = e20Last + 0.3 * atrLast;
    return { ok:true, trigger, e20Last, atrLast };
  }
  function predictEntryB_(series, buyStartMs, ltpPxNow, isHolidayToday){
    const t = computeTriggerB_(series, buyStartMs);
    if (!t.ok) return { plannedDate: null, plannedText: '', entryPx: NaN, reason: 'setup-not-ok' };
    if (!isHolidayToday && IN_SESSION && Number.isFinite(ltpPxNow) && ltpPxNow>0 && ltpPxNow <= t.trigger) {
      const todayStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
      const today = new Date(todayStr+'T00:00:00');
      return { plannedDate: today, plannedText: fmtDate(today), entryPx: +t.trigger.toFixed(2), reason: 'live-hit' };
    }
    const baseBuy = chooseBuyDate_();
    return { plannedDate: baseBuy, plannedText: fmtDate(baseBuy), entryPx: +t.trigger.toFixed(2), reason: 'planned-next' };
  }

  const itemsByCode = new Map();
  const itemsList = [];
  for (const r of ltpVals) {
    const code = norm(r[colCode]); if (!code) continue;
    const ltpFromColE = ltpFromColumnE_(r);
    if (!(Number.isFinite(ltpFromColE) && ltpFromColE > MIN_LTP_E_FILTER)) continue;
    const series = seriesByCode[code]; if (!series || !series.length) continue;
    const s = lightStats.get(code); if (!s) continue;

    const sector = (colSector!=null)? r[colSector] : '';
    if (sector && String(sector).toUpperCase().includes('MUTUAL')) continue;

    let d0ValMn = NaN, d0Px = NaN;
    if (((AFTER_1430 && isTradingDayBD_()) || hasHDataToday)) {
      const t = todayHMap.get(code); if (t){ d0ValMn = t.val; d0Px = t.close; }
    } else {
      d0ValMn = (colValMn != null) ? toNum(r[colValMn]) : NaN;
      d0Px    = todaysCloseFromLTPRow_(r);
      if (colCloseP!=null) { const cp = toNum(r[colCloseP]); if (cp===0) {} }
    }

    const ltpColC = toNum(r[2]); // LTP sheet, column C

    const {txt: valueIncTxt, num: valueIncNum} = valueInc_Last1vs3Avg_(series, buyStartMs, d0ValMn);
    const {txt: priceChgTxt, num: priceChgNum} = priceChg_TodayPlus2_(series, buyStartMs, d0Px, isHolidayToday);

    const useLive = (!hasHDataToday && isTradingDayBD_() && IN_SESSION);
    const predEnt = predictEntryB_(series, buyStartMs, useLive ? d0Px : NaN, isHolidayToday);

    let entry = predEnt.entryPx;
    let buyDate = predEnt.plannedDate || chooseBuyDate_();
    let buyDateText = predEnt.plannedText || fmtDate(buyDate);

    let stop = Number.isFinite(entry) ? +(entry * (1 - STOP_FLOOR_PCT)).toFixed(2) : NaN;

    itemsByCode.set(code, {
      code, sector, series, s, valueIncNum, priceChgNum, valueIncTxt, priceChgTxt,
      entry, stop, buyDateText, entryReason: predEnt.reason, action: 'Listed',
      targetPct: NaN, target: NaN, winPct: NaN, medHold: '', invest: 0, ltpColC
    });
    itemsList.push(itemsByCode.get(code));
  }
  DBG.builtItems = itemsList.length;

  function passGates_(it){
    const sc      = it.s.trendScore;
    const vSpike  = valueSpikeByCode.get(it.code) || 0;
    const hasDemand = Number.isFinite(vSpike) && vSpike > 0;

    if (!Number.isFinite(sc) || (sc < 1 && !hasDemand)) return false;

    if (DEF.USE_LIQUIDITY_GUARD &&
        Number.isFinite(it.s.liqMed20) &&
        it.s.liqMed20 < DEF.LIQ_MIN_VAL_MN_MED20) {
      return false;
    }

    if (REQUIRE_VALUE_INCREASE &&
        Number.isFinite(it.valueIncNum) &&
        it.valueIncNum < MIN_VALUE_INC_FILTER_PCT) {
      return false;
    }

    if (REQUIRE_PRICE_INCREASE &&
        Number.isFinite(it.priceChgNum)) {
      if (!hasDemand && it.priceChgNum < MIN_PRICE_CHG_FILTER_PCT) return false;
      if (hasDemand  && it.priceChgNum < (MIN_PRICE_CHG_FILTER_PCT - 2)) return false;
    }

    return Number.isFinite(it.entry);
  }

  let pool = itemsList.filter(passGates_);
  const bestMap = new Map();
  for (const it of pool){
    const cur = bestMap.get(it.code);
    if (!cur || it.valueIncNum > cur.valueIncNum || (it.valueIncNum === cur.valueIncNum && it.priceChgNum > cur.priceChgNum)) {
      bestMap.set(it.code, it);
    }
  }
  pool = Array.from(bestMap.values());
  pool.sort((A,B)=>{
    const aV = Number.isFinite(A.valueIncNum)?A.valueIncNum:-Infinity;
    const bV = Number.isFinite(B.valueIncNum)?B.valueIncNum:-Infinity;
    if (bV !== aV) return bV - aV;
    const aP = Number.isFinite(A.priceChgNum)?A.priceChgNum:-Infinity;
    const bP = Number.isFinite(B.priceChgNum)?B.priceChgNum:-Infinity;
    return bP - aP;
  });
  pool = pool.slice(0, MAX_CANDIDATES);

  const frozenPurchased = [];
  const frozenWatch = [];

  for (const [code, exInfo] of existing.map.entries()) {
    const base = itemsByCode.get(code);
    if (!base) continue;
    let it = base;
    it.action = exInfo.action;

    if (exInfo.action === 'Purchased') {
      const freezeBuy = exInfo.buyDateText && exInfo.buyDateText.trim() ? exInfo.buyDateText : base.buyDateText;
      const freezeEntry = Number.isFinite(exInfo.frozenEntry) ? exInfo.frozenEntry : base.entry;
      it.buyDateText = freezeBuy;
      it.entry = freezeEntry;
      frozenPurchased.push(it);
    } else if (exInfo.action === 'Watch') {
      it.buyDateText = base.buyDateText;
      frozenWatch.push(it);
    }
  }

  const orderedPurchased = existing.orderPurchased.map(c => frozenPurchased.find(x=>x.code===c)).filter(Boolean);
  const orderedWatch     = existing.orderWatch.map(c => frozenWatch.find(x=>x.code===c)).filter(Boolean);

  const pinnedCodes = new Set([...orderedPurchased.map(x=>x.code), ...orderedWatch.map(x=>x.code)]);
  const listed = pool.filter(x => !pinnedCodes.has(x.code)).map(x=>{ x.action='Listed'; return x; });

  const merged = [...orderedPurchased, ...orderedWatch, ...listed];

  function nextPeakReturns_(arr, H){
    const out=[];
    for(let i=0;i<arr.length-1;i++){
      const e=arr[i].close; if(!Number.isFinite(e)||e<=0) continue;
      const end=Math.min(arr.length-1,i+H); let peak=-Infinity;
      for(let j=i+1;j<=end;j++){
        const c=arr[j].close; if(Number.isFinite(c)&&c>peak) peak=c;
      }
      if (peak>-Infinity){
        const r=(peak/e)-1;
        if (Number.isFinite(r)) out.push(r);
      }
    }
    return out;
  }
  function quantile_(a, q){
    const b=a.filter(Number.isFinite).slice().sort((x,y)=>x-y); if (!b.length) return NaN;
    const pos=(b.length-1)*q; const lo=Math.floor(pos), hi=Math.ceil(pos);
    if (lo===hi) return b[lo];
    return b[lo] + (b[hi]-b[lo])*(pos-lo);
  }
  function clamp(x, lo, hi){ return Math.max(lo, Math.min(hi, x)); }
  function predictTargetPct_Achievable_(arr, entry){
    const ret = nextPeakReturns_(arr, MAX_HOLD_DAYS);
    if (ret.length >= MIN_SAMPLES) {
      let tPct = quantile_(ret, 1 - DESIRED_HIT_PROB); if (!Number.isFinite(tPct)) tPct = MIN_TARGET_PCT;
      return clamp(tPct, MIN_TARGET_PCT, MAX_TARGET_PCT);
    }
    const hi=arr.map(x=>x.high), lo=arr.map(x=>x.low), cl=arr.map(x=>x.close);
    const a = atr_(hi,lo,cl,ATR_LEN), atrLast = a[a.length-1];
    const atrPct = (Number.isFinite(atrLast)&&Number.isFinite(cl[cl.length-1])&&cl[cl.length-1]>0)? (atrLast/cl[cl.length-1]) : NaN;
    if (Number.isFinite(atrPct) && atrPct>0) return clamp(atrPct*TARGET_FALLBACK_ATR_MULT, MIN_TARGET_PCT, MAX_TARGET_PCT);
    const last10 = arr.slice(-10).map(x=>x.close).filter(Number.isFinite);
    if (last10.length >= 5) {
      const minC = Math.min.apply(null,last10), maxC = Math.max.apply(null,last10);
      const swingPct = (maxC - minC) / (entry || maxC); return clamp(swingPct*0.7, MIN_TARGET_PCT, MAX_TARGET_PCT);
    }
    return MIN_TARGET_PCT;
  }
  function computeWinAndHold_(arr, tgtPct){
    let wins=0, totals=0, holdWins=[];
    for(let i=0;i<arr.length-1;i++){
      const e=arr[i].close; if(!Number.isFinite(e)||e<=0) continue;
      const tgt=e*(1+tgtPct); const end=Math.min(arr.length-1,i+MAX_HOLD_DAYS);
      let hit=false, hold=NaN;
      for(let j=i+1;j<=end;j++){
        const c = arr[j].close;
        if(Number.isFinite(c)&&c>=tgt){ hit=true; hold=j-i; break; }
      }
      totals++; if(hit){ wins++; holdWins.push(hold); }
    }
    const rawWin = totals? (wins*100/totals) : 0;
    const winPct = (((rawWin/100)*arr.length + GLOBAL_WIN_MEAN*PRIOR_WEIGHT) / (arr.length + PRIOR_WEIGHT) * 100);
    const medHold = holdWins.length ? Math.min(Math.round(median(holdWins)), MAX_HOLD_DAYS) : '';
    return {winPct, medHold};
  }

  for (const it of merged) {
    const arr = it.series;
    const entryPx = it.entry;
    if (!Number.isFinite(entryPx)) continue;

    const hi=arr.map(x=>x.high), lo=arr.map(x=>x.low), cl=arr.map(x=>x.close);
    const a = atr_(hi,lo,cl,ATR_LEN), atrLast = a[a.length-1];
    const atrPct = (Number.isFinite(atrLast)&&Number.isFinite(cl[cl.length-1])&&cl[cl.length-1]>0)? (atrLast/cl[cl.length-1])*100 : NaN;
    if (Number.isFinite(atrPct)) {
      const atrAbs = entryPx * (atrPct/100);
      const atrStop = entryPx - ATR_STOP_MULT*atrAbs;
      it.stop = +Math.min(+it.stop||Infinity, atrStop).toFixed(2);
    }

    const tPct = predictTargetPct_Achievable_(arr, entryPx);
    const {winPct, medHold} = computeWinAndHold_(arr, tPct);

    it.targetPct = tPct;
    it.target    = +(entryPx * (1 + tPct)).toFixed(2);
    it.winPct    = +(+winPct).toFixed(1);
    it.medHold   = medHold;

    it.invest    = investAmountFromWin_(it.winPct);
  }

  const finalRows = merged.map(it => {
    const buyDate = it.buyDateText || fmtDate(chooseBuyDate_());
    const exitDate = (function(){
      const d = new Date(buyDate); const base = isNaN(d) ? chooseBuyDate_() : d;
      return addTradingDaysBD_(base, it.medHold || MAX_HOLD_DAYS);
    })();

    function displayTrendSym_(sym) {
      if (sym === '⬆️') return '⬆️';
      if (sym === '⬇️') return '⬇️';
      return '〰️';
    }

    const dsexSym  = displayTrendSym_(regimeSym);
    const stockSym = displayTrendSym_(it.s.trendSym);
    const trendText = "'" + dsexSym + ' | ' + stockSym;

    return [
      buyDate, fmtDate(exitDate), it.code, it.sector,
      Number.isFinite(it.ltpColC) ? it.ltpColC : '',
      it.entry, it.stop, it.target, Math.round((it.targetPct||0)*100)+'%', (it.medHold||''),
      Number.isFinite(it.winPct)? Math.round(it.winPct)+'%' : '',
      it.valueIncTxt, it.priceChgTxt,
      trendText,
      'OK', it.action
    ];
  });

  DBG.pinnedPlusKept = finalRows.length;

  writeSummary_(finalRows, ACTION_LIST);
  writeLog_(Object.assign({
    note:'Entry=EMA20+0.3*ATR; live LTP 10:00–14:30 used when HData isn’t ready. Action: Purchased (freeze Buy Date & Entry), Watch (pinned, rolling date). Sections: Purchased → Watch → Listed. Column E now shows LTP from LTP sheet column C. Trend has 4 pillars: Location, Momentum, Structure, Demand(ValueSpike>0). DSEX regime included in Trend column as DSEX | Stock.',
    tz:TZ, time:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    todaySource, shortlisted: finalRows.length,
    regimeOK: regimeOK ? 'OK' : 'Caution', regimeSym, regimeSource
  }, DBG));

  function applyThinBordersToUsedRange_(sh, headersLen) {
    const lastR = sh.getLastRow(); const lastC = Math.min(sh.getMaxColumns(), headersLen);
    if (lastR > 0 && lastC > 0) sh.getRange(1,1,lastR,lastC).setBorder(true,true,true,true,true,true);
  }
  function writeSummary_(rows, ACTION_LIST) {
    const headers = [
      'Buy Date','Exit Date','CODE','SECTOR','LTP',
      'ENTRY','STOP','TARGET','Target %','Hold Days',
      'Win %','Value Increase %','Price Change %','Trend',
      'Trend Check','Action'
    ];
    const sh = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
    if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    const lr = sh.getLastRow(); if (lr > 1) sh.getRange(2,1,lr-1,headers.length).clearContent();
    if (rows && rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
    if (rows && rows.length){
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(ACTION_LIST, true).setAllowInvalid(false).build();
      sh.getRange(2, headers.length, rows.length, 1).setDataValidation(rule);
    }
    applyThinBordersToUsedRange_(sh, headers.length);
  }
  function writeLog_(obj){
    const sh = ss.getSheetByName('SummaryLog') || ss.insertSheet('SummaryLog');
    const lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr && lc) sh.getRange(1,1,lr,lc).clearContent();
    const rows = Object.entries(obj).map(([k,v])=>[k,String(v)]);
    if (rows.length) sh.getRange(1,1,rows.length,2).setValues(rows);
    applyThinBordersToUsedRange_(sh, 2);
  }
  function safeWrite_(rows, ACTION_LIST, note){
    writeSummary_([], ACTION_LIST);
    writeLog_({ note, tz:TZ, time:Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), shortlisted: 0 });
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Summary Utils')
    .addItem('Refresh Summary', 'refreshSummary')
    .addToUi();
}
