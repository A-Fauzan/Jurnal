// ╔══════════════════════════════════════════════════════════╗
// ║  Jurnal Kerja Harian — BIG                              ║
// ║  Google Apps Script API                                  ║
// ╚══════════════════════════════════════════════════════════╝

const SPREADSHEET_ID = "ISI_ID_GOOGLE_SHEET_KAMU";
const SECRET_KEY     = "jurnal-fauzan-big-rahasia";
const SHEET_DATA     = "jurnal";

const KATEGORI = ["Pemetaan","Analisis GIS","Administrasi","Rapat","Pengembangan","Pelatihan","Koordinasi","Lainnya"];

function doGet(e) {
  const p        = (e && e.parameter) ? e.parameter : {};
  const action   = p.action   || "ping";
  const callback = p.callback || "";

  if (action !== "ping" && p.key !== SECRET_KEY)
    return respond({ error: "Unauthorized" }, callback);

  let result;
  try {
    if      (action === "ping")        result = { status:"ok" };
    else if (action === "get_jurnal")  result = getJurnal(p.month||"", p.minggu||"");
    else if (action === "add_jurnal")  result = addJurnal(p);
    else if (action === "delete_jurnal") result = deleteJurnal(p.id);
    else if (action === "get_summary") result = getSummary(p.month||"");
    else result = { error: "Unknown action" };
  } catch(err) {
    result = { error: err.message };
  }
  return respond(result, callback);
}

function doPost(e) { return doGet(e); }

function getSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_DATA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATA);
    sheet.appendRow(["id","tanggal","kategori","kegiatan","durasi_jam","output","catatan"]);
    sheet.getRange(1,1,1,7).setFontWeight("bold");
  }
  return sheet;
}

function getJurnal(month, minggu) {
  const sheet = getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return [];
  const rows = sheet.getRange(2,1,last-1,7).getValues();
  return rows.filter(r => {
    if (!r[0]) return false;
    const tgl = tglStr(r[1]);
    if (month && !tgl.startsWith(month)) return false;
    return true;
  }).map(r => ({
    id: Number(r[0]), tanggal: tglStr(r[1]),
    kategori: r[2], kegiatan: r[3],
    durasi_jam: Number(r[4]), output: r[5], catatan: r[6]
  })).sort((a,b) => new Date(b.tanggal)-new Date(a.tanggal));
}

function addJurnal(p) {
  const sheet = getSheet();
  const newId = nextId(sheet);
  const today = p.tanggal || fmtDate(new Date());
  sheet.appendRow([newId, today, p.kategori||"Lainnya",
    p.kegiatan||"", Number(p.durasi_jam)||1,
    p.output||"", p.catatan||""]);
  return { success:true, id:newId };
}

function deleteJurnal(id) {
  const sheet = getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return { error:"Tidak ada data" };
  const ids = sheet.getRange(2,1,last-1,1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (Number(ids[i][0]) === Number(id)) { sheet.deleteRow(i+2); return {success:true}; }
  }
  return { error:"ID tidak ditemukan" };
}

function getSummary(month) {
  const rows   = getJurnal(month, "");
  const total  = rows.reduce((s,r) => s+r.durasi_jam, 0);
  const byKat  = {};
  rows.forEach(r => { byKat[r.kategori] = (byKat[r.kategori]||0) + r.durasi_jam; });
  const byDate = {};
  rows.forEach(r => { byDate[r.tanggal] = (byDate[r.tanggal]||0) + r.durasi_jam; });
  return { total_jam: total, total_kegiatan: rows.length, by_kategori: byKat, by_date: byDate, month };
}

function tglStr(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v,"Asia/Jakarta","yyyy-MM-dd");
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);
  return s.substring(0,10);
}
function fmtDate(d) { return Utilities.formatDate(d,"Asia/Jakarta","yyyy-MM-dd"); }
function nextId(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const ids = sheet.getRange(2,1,last-1,1).getValues().flat().map(Number).filter(n=>n>0);
  return ids.length ? Math.max(...ids)+1 : 1;
}
function respond(data, callback) {
  const json = JSON.stringify(data);
  const out  = callback ? callback+"("+json+")" : json;
  return ContentService.createTextOutput(out)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
