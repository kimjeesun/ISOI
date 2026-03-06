// ============================================================
// 올리브영 랭킹 트래커 + 뷰어 수 트래커 — GAS Web App
// ============================================================

const CONFIG = {
  CATEGORIES: ["전체TOP100", "스킨케어"],
  KEEP_DAYS:  30,
  SHEET_RAW:    "📊 원본데이터",
  SHEET_BRAND:  "📦 브랜드집계",
  SHEET_LIVE:   "🔴 실시간현황",
  SHEET_VIEWER: "👁 뷰어트래킹",
};

// ─────────────────────────────────────────
// Web App 수신
// ─────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // SECRET 검증 (필요시 주석 해제)
    // const SECRET = "abc123";
    // if (payload.secret !== SECRET) {
    //   return jsonResponse_({ ok: false, error: "Unauthorized" });
    // }

    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const rows       = payload.rows       || [];
    const viewerRows = payload.viewerRows || [];
    const dateStr    = payload.dateStr;
    const timeStr    = payload.timeStr;

    if (rows.length === 0 && viewerRows.length === 0) {
      return jsonResponse_({ ok: false, error: "No data" });
    }

    // 랭킹 데이터 저장
    if (rows.length > 0) {
      appendRawData_(ss, rows);
      refreshLiveSheet_(ss, rows, dateStr, timeStr);
      refreshBrandSheet_(ss, dateStr);
    }

    // 뷰어 데이터 저장
    let viewerSaved = 0;
    if (viewerRows.length > 0) {
      viewerSaved = appendViewerData_(ss, viewerRows);
      refreshViewerSheet_(ss);
    }

    pruneOldData_(ss);

    return jsonResponse_({ ok: true, saved: rows.length, viewerSaved });

  } catch (err) {
    Logger.log("doPost 오류: " + err.message);
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function doGet(e) {
  return jsonResponse_({ ok: true, status: "running" });
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────
// 초기 설정
// ─────────────────────────────────────────
function setup() {
  initSheets_();
  SpreadsheetApp.getUi().alert(
    "✅ 설정 완료!\n배포 > 새 배포 > 웹 앱으로 배포해주세요."
  );
}

// ─────────────────────────────────────────
// 뷰어 수 데이터 저장
// ─────────────────────────────────────────
const VIEWER_HEADERS = ["날짜", "시각", "상품명", "URL", "뷰어 수"];

function appendViewerData_(ss, viewerRows) {
  const sh = ss.getSheetByName(CONFIG.SHEET_VIEWER);
  if (!sh) return 0;

  if (sh.getLastRow() === 0) {
    sh.appendRow(VIEWER_HEADERS);
    sh.getRange(1, 1, 1, VIEWER_HEADERS.length)
      .setBackground("#7B1FA2").setFontColor("white").setFontWeight("bold");
    sh.setFrozenRows(1);
  }

  const data = viewerRows.map(r => [
    r.dateStr, r.timeStr, r.productName, r.url, r.viewerCount
  ]);
  sh.getRange(sh.getLastRow() + 1, 1, data.length, VIEWER_HEADERS.length).setValues(data);
  return data.length;
}

// ─────────────────────────────────────────
// 뷰어 추이 시트 (상품별 시간대 추이)
// ─────────────────────────────────────────
function refreshViewerSheet_(ss) {
  const rawSh  = ss.getSheetByName(CONFIG.SHEET_VIEWER);
  if (!rawSh || rawSh.getLastRow() < 2) return;

  // 오늘 날짜
  const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");

  const allData = rawSh.getDataRange().getValues();
  const todayRows = allData.slice(1).filter(r => {
    const d = r[0];
    const dStr = (d instanceof Date)
      ? Utilities.formatDate(d, "Asia/Seoul", "yyyy-MM-dd")
      : String(d).substring(0, 10);
    return dStr === today;
  });
  if (!todayRows.length) return;

  // 시각값 정규화 헬퍼 (Date 객체 or 문자열 모두 HH:mm으로 변환)
  const toTimeStr = (t) => (t instanceof Date)
    ? Utilities.formatDate(t, "Asia/Seoul", "HH:mm")
    : String(t).substring(0, 5);

  // 상품 목록 & 시각 목록 (모두 정규화된 문자열로)
  const products = [...new Set(todayRows.map(r => r[2]))]; // 상품명
  const times    = [...new Set(todayRows.map(r => toTimeStr(r[1])))].sort();

  // 추이 시트 생성/초기화 (별도 시트)
  const sheetName = "👁 뷰어추이";
  let trendSh = ss.getSheetByName(sheetName);
  if (!trendSh) trendSh = ss.insertSheet(sheetName);
  trendSh.clearContents();
  trendSh.clearFormats();

  let r = 1;

  // 타이틀
  trendSh.getRange(r, 1).setValue("👁 상품 뷰어 수 시간대별 추이");
  trendSh.getRange(r, 1, 1, times.length + 1).merge()
    .setBackground("#7B1FA2").setFontColor("white")
    .setFontWeight("bold").setFontSize(13).setHorizontalAlignment("left");
  r++;

  trendSh.getRange(r, 1).setValue(`${today} 기준`);
  trendSh.getRange(r, 1, 1, times.length + 1).merge()
    .setFontColor("#555").setHorizontalAlignment("center").setFontSize(10);
  r++;

  // 헤더: 상품명 | 시각1 | 시각2 | ...
  trendSh.getRange(r, 1).setValue("상품명 \\ 수집시각");
  times.forEach((t, i) => trendSh.getRange(r, i + 2).setValue(t));
  trendSh.getRange(r, 1, 1, times.length + 1)
    .setBackground("#F3E5F5").setFontWeight("bold").setHorizontalAlignment("center");
  trendSh.getRange(r, 1).setHorizontalAlignment("left"); // 상품명 헤더만 좌측
  r++;

  // 상품별 시각별 뷰어 수 (실제 데이터 있는 상품만)
  products.forEach(product => {
    // 해당 상품의 데이터가 하나도 없으면 행 건너뜀
    const hasData = times.some(t => todayRows.find(row => toTimeStr(row[1]) === t && row[2] === product));
    if (!hasData) return;
    trendSh.getRange(r, 1).setValue(product).setFontWeight("bold").setHorizontalAlignment("left");
    times.forEach((t, i) => {
      const hit = todayRows.find(row => toTimeStr(row[1]) === t && row[2] === product);
      if (hit) {
        const cell = trendSh.getRange(r, i + 2);
        cell.setValue(hit[4]); // 뷰어 수

        // 뷰어 수에 따라 색상 강조
        const v = Number(hit[4]);
        if      (v >= 500) cell.setBackground("#E53935").setFontColor("white");
        else if (v >= 200) cell.setBackground("#FB8C00").setFontColor("white");
        else if (v >= 100) cell.setBackground("#FDD835");
        else if (v >= 50)  cell.setBackground("#E8F5E9");
      } else {
        trendSh.getRange(r, i + 2).setValue("-").setFontColor("#ccc");
      }
    });
    r++;
  });

  // 범례
  r += 2;
  trendSh.getRange(r, 1).setValue("범례");
  trendSh.getRange(r, 1).setFontWeight("bold");
  r++;
  [
    ["🔴 500명 이상", "#E53935", "white"],
    ["🟠 200~499명",  "#FB8C00", "white"],
    ["🟡 100~199명",  "#FDD835", "black"],
    ["🟢 50~99명",    "#E8F5E9", "black"],
  ].forEach(([label, bg, fc]) => {
    trendSh.getRange(r, 1).setValue(label).setBackground(bg).setFontColor(fc);
    r++;
  });

  trendSh.setColumnWidth(1, 550);  // A: 상품명 넓게
  for (let i = 2; i <= times.length + 1; i++) {
    trendSh.setColumnWidth(i, 80);  // B~: 시각별 뷰어 수
  }
}

// ─────────────────────────────────────────
// 원본 데이터 저장
// ─────────────────────────────────────────
const RAW_HEADERS = [
  "날짜","시각","카테고리","순위","브랜드","상품명",
  "현재가","정가","할인율(%)","세일","쿠폰","증정","오늘드림"
];

function appendRawData_(ss, rows) {
  const sh = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (sh.getLastRow() === 0) {
    sh.appendRow(RAW_HEADERS);
    sh.getRange(1, 1, 1, RAW_HEADERS.length)
      .setBackground("#34A853").setFontColor("white").setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  const data = rows.map(r => [
    r.dateStr, r.timeStr, r.category, r.rank, r.brand, r.name,
    r.curPrice || "", r.orgPrice || "", r.discount || "",
    r.hasSale || "", r.hasCoupon || "", r.hasGift || "", r.hasDelivery || ""
  ]);
  sh.getRange(sh.getLastRow() + 1, 1, data.length, RAW_HEADERS.length).setValues(data);
}

// ─────────────────────────────────────────
// 실시간 현황 시트
// ─────────────────────────────────────────
function refreshLiveSheet_(ss, rows, dateStr, timeStr) {
  const sh = ss.getSheetByName(CONFIG.SHEET_LIVE);
  sh.clearContents(); sh.clearFormats();
  let r = 1;

  sh.getRange(r, 1).setValue("🔴 올리브영 실시간 랭킹");
  sh.getRange(r, 1, 1, 9).merge()
    .setBackground("#D93025").setFontColor("white")
    .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
  r++;
  sh.getRange(r, 1).setValue(`수집 시각: ${dateStr} ${timeStr} KST`);
  sh.getRange(r, 1, 1, 9).merge()
    .setFontColor("#555").setHorizontalAlignment("center").setFontSize(10);
  r += 2;

  const cats = [...new Set(rows.map(x => x.category))];
  for (const cat of cats) {
    const catRows = rows.filter(x => x.category === cat);
    if (!catRows.length) continue;

    sh.getRange(r, 1).setValue(`▶ ${cat}  (${catRows.length}개)`);
    sh.getRange(r, 1, 1, 9).merge()
      .setBackground("#1A73E8").setFontColor("white").setFontWeight("bold").setFontSize(12);
    r++;
    sh.getRange(r, 1, 1, 9).setValues([["순위","브랜드","상품명","현재가","할인율","세일","쿠폰","증정","오늘드림"]])
      .setBackground("#E8F0FE").setFontWeight("bold");
    r++;

    catRows.forEach(item => {
      sh.getRange(r, 1, 1, 9).setValues([[
        item.rank, item.brand, item.name,
        item.curPrice ? `₩${Number(item.curPrice).toLocaleString()}` : "-",
        item.discount ? `${item.discount}%` : "-",
        item.hasSale     === "Y" ? "✅" : "",
        item.hasCoupon   === "Y" ? "✅" : "",
        item.hasGift     === "Y" ? "✅" : "",
        item.hasDelivery === "Y" ? "✅" : "",
      ]]);
      // 가운데 정렬: A(순위), D(현재가), E~I(할인율/세일/쿠폰/증정/오늘드림)
      sh.getRange(r, 1).setHorizontalAlignment("center"); // 순위
      sh.getRange(r, 4, 1, 6).setHorizontalAlignment("center"); // 현재가~오늘드림
      if      (item.rank === 1) sh.getRange(r, 1, 1, 9).setBackground("#FFF9C4");
      else if (item.rank === 2) sh.getRange(r, 1, 1, 9).setBackground("#F5F5F5");
      else if (item.rank === 3) sh.getRange(r, 1, 1, 9).setBackground("#FFF3E0");
      r++;
    });
    r += 2;
  }

  // 열 너비 고정
  sh.setColumnWidth(1, 50);   // A: 순위
  sh.setColumnWidth(2, 120);  // B: 브랜드
  sh.setColumnWidth(3, 400);  // C: 상품명 (넓게)
  sh.setColumnWidth(4, 110);  // D: 현재가
  sh.setColumnWidth(5, 80);   // E: 할인율
  sh.setColumnWidth(6, 60);   // F: 세일
  sh.setColumnWidth(7, 60);   // G: 쿠폰
  sh.setColumnWidth(8, 60);   // H: 증정
  sh.setColumnWidth(9, 80);   // I: 오늘드림

  // 헤더도 가운데 정렬
  sh.getRange(1, 1, sh.getLastRow(), 9).setVerticalAlignment("middle");
}

// ─────────────────────────────────────────
// 브랜드 집계 시트
// ─────────────────────────────────────────
function refreshBrandSheet_(ss, targetDate) {
  const rawSh   = ss.getSheetByName(CONFIG.SHEET_RAW);
  const brandSh = ss.getSheetByName(CONFIG.SHEET_BRAND);
  brandSh.clearContents(); brandSh.clearFormats();

  const allData   = rawSh.getDataRange().getValues();
  if (allData.length < 2) return;

  const I = { date:0,time:1,cat:2,rank:3,brand:4,name:5,
              cur:6,org:7,disc:8,sale:9,coupon:10,gift:11,delivery:12 };
  // 날짜 비교: Date 객체 또는 문자열 모두 처리
  const toDateStr = (v) => (v instanceof Date)
    ? Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd")
    : String(v).substring(0, 10);
  const targetDateStr = toDateStr(targetDate);
  const todayRows = allData.slice(1).filter(r => toDateStr(r[I.date]) === targetDateStr);
  if (!todayRows.length) return;

  let r = 1;
  for (const catName of CONFIG.CATEGORIES) {
    const catRows = todayRows.filter(row => row[I.cat] === catName);
    if (!catRows.length) continue;

    const brandMap = {};
    catRows.forEach(row => {
      const b = row[I.brand] || "미상";
      if (!brandMap[b]) brandMap[b] = { count:0,rankSum:0,topRank:999,
                                         saleCount:0,couponCount:0,priceSum:0,priceCount:0 };
      const d = brandMap[b];
      d.count++;
      d.rankSum  += Number(row[I.rank]);
      d.topRank   = Math.min(d.topRank, Number(row[I.rank]));
      if (row[I.sale]   === "Y") d.saleCount++;
      if (row[I.coupon] === "Y") d.couponCount++;
      if (row[I.cur] > 0) { d.priceSum += Number(row[I.cur]); d.priceCount++; }
    });

    const timeSlots = [...new Set(catRows.map(row => row[I.time]))];
    brandSh.getRange(r, 1).setValue(`📦 ${catName} — 브랜드별 집계 (${targetDate})`);
    brandSh.getRange(r, 1, 1, 8).merge()
      .setBackground("#1A73E8").setFontColor("white")
      .setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
    r++;
    brandSh.getRange(r, 1).setValue(`오늘 수집 ${timeSlots.length}회 / 총 ${catRows.length}건`);
    brandSh.getRange(r, 1, 1, 8).merge()
      .setFontColor("#555").setFontSize(10).setHorizontalAlignment("center");
    r++;
    const hdrs = ["브랜드","등장 횟수","평균 순위","최고 순위","랭킹 점수","세일 횟수","쿠폰 횟수","평균 가격(원)"];
    brandSh.getRange(r, 1, 1, 8).setValues([hdrs])
      .setBackground("#E8F0FE").setFontWeight("bold").setHorizontalAlignment("center");
    r++;

    const sorted = Object.entries(brandMap).map(([name, v]) => ({
      name, count: v.count,
      avgRank:  Math.round(v.rankSum / v.count * 10) / 10,
      topRank:  v.topRank,
      score:    Math.round(v.count * 1000 / (v.rankSum / v.count)),
      saleCount: v.saleCount, couponCount: v.couponCount,
      avgPrice: v.priceCount > 0 ? Math.round(v.priceSum / v.priceCount) : 0,
    })).sort((a, b) => b.count !== a.count ? b.count - a.count : a.avgRank - b.avgRank);

    sorted.forEach((b, i) => {
      brandSh.getRange(r, 1, 1, 8).setValues([[
        b.name, b.count, b.avgRank, `${b.topRank}위`, b.score,
        b.saleCount || "", b.couponCount || "",
        b.avgPrice > 0 ? b.avgPrice : "-",
      ]]);
      if (i < 5) {
        const colors = ["#FFF176","#F5F5F5","#FFE0B2","#E8F5E9","#E3F2FD"];
        brandSh.getRange(r, 1, 1, 8).setBackground(colors[i]);
      }
      r++;
    });
    r += 3;
  }
}

// ─────────────────────────────────────────
// 오래된 데이터 삭제
// ─────────────────────────────────────────
function pruneOldData_(ss) {
  const sh     = ss.getSheetByName(CONFIG.SHEET_RAW);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - CONFIG.KEEP_DAYS);
  const cutStr = Utilities.formatDate(cutoff, "Asia/Seoul", "yyyy-MM-dd");
  const data   = sh.getDataRange().getValues();
  let del = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && data[i][0] < cutStr) { sh.deleteRow(i + 1); del++; }
  }
  if (del > 0) Logger.log(`🗑 ${del}행 삭제`);
}

// ─────────────────────────────────────────
// 헬퍼
// ─────────────────────────────────────────
function initSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEET_RAW, CONFIG.SHEET_LIVE, CONFIG.SHEET_BRAND,
   CONFIG.SHEET_VIEWER, "👁 뷰어추이"].forEach(n => {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });
  ["Sheet1","시트1"].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 5) try { ss.deleteSheet(sh); } catch(_) {}
  });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🫒 올리브영")
    .addItem("📦 브랜드 집계 새로고침", "refreshBrandSheetManual")
    .addItem("👁 뷰어 추이 새로고침",   "refreshViewerManual")
    .addSeparator()
    .addItem("⚙️ 초기 설정 (최초 1회)", "setup")
    .addToUi();
}

function refreshBrandSheetManual() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const rawSh = ss.getSheetByName(CONFIG.SHEET_RAW);
  const data  = rawSh.getDataRange().getValues();
  const dates = data.slice(1).map(r => r[0]).filter(Boolean);
  if (!dates.length) { SpreadsheetApp.getUi().alert("수집된 데이터가 없습니다."); return; }
  const latest = dates.reduce((a, b) => a > b ? a : b);
  refreshBrandSheet_(ss, latest);
  SpreadsheetApp.getUi().alert(`✅ ${latest} 기준 브랜드 집계 완료`);
}

function refreshViewerManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  refreshViewerSheet_(ss);
  SpreadsheetApp.getUi().alert("✅ 뷰어 추이 새로고침 완료");
}

// 원본데이터 시트 열 너비 수동 조정
function fixRawSheetColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (!sh) { SpreadsheetApp.getUi().alert("📊 원본데이터 시트를 찾을 수 없습니다."); return; }

  sh.setColumnWidth(1, 110);  // A: 날짜
  sh.setColumnWidth(2, 70);   // B: 시각
  sh.setColumnWidth(3, 100);  // C: 카테고리
  sh.setColumnWidth(4, 50);   // D: 순위
  sh.setColumnWidth(5, 110);  // E: 브랜드
  sh.setColumnWidth(6, 350);  // F: 상품명 ← 넓게
  sh.setColumnWidth(7, 90);   // G: 현재가
  sh.setColumnWidth(8, 90);   // H: 정가
  sh.setColumnWidth(9, 80);   // I: 할인율
  sh.setColumnWidth(10, 55);  // J: 세일
  sh.setColumnWidth(11, 55);  // K: 쿠폰
  sh.setColumnWidth(12, 55);  // L: 증정
  sh.setColumnWidth(13, 70);  // M: 오늘드림

  SpreadsheetApp.getUi().alert("✅ 원본데이터 열 너비 조정 완료!");
}