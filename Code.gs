/**
 * ANOVA + Regresyon oyunları için ortak veri toplayıcı (Google Apps Script).
 *
 * KURULUM (özet — ayrıntı için KURULUM.md):
 *   1. Bir Google E-Tablosu aç.
 *   2. Uzantılar > Apps Script menüsünden bu kodu yapıştır.
 *   3. Dağıt (Deploy) > Yeni dağıtım > tür: Web uygulaması
 *        - "Yürüt: Ben"   /  "Erişim: Herkes"
 *   4. Çıkan /exec ile biten adresi kopyala ve iki HTML dosyasındaki
 *      SHEET_URL değişkenine yapıştır.
 *
 * Sekmeler (yoksa otomatik oluşturulur):
 *   anova      -> [zaman, zorluk, skor]
 *   regresyon  -> [zaman, boyut, omur, skor]
 */

var HEADERS = {
  anova:     ['zaman', 'zorluk', 'skor'],
  regresyon: ['zaman', 'boyut', 'omur', 'skor']
};

/* ---- YAZMA: oyun bir tur bitince buraya POST eder ---- */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var name = d.sheet;
    if (!HEADERS[name]) return out_({ ok: false, error: 'bilinmeyen sheet: ' + name });

    var sh = getSheet_(name);
    var row = (name === 'anova')
      ? [new Date(), d.zorluk, Number(d.skor)]
      : [new Date(), Number(d.boyut), Number(d.omur), Number(d.skor)];

    sh.appendRow(row);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ---- OKUMA: "Sınıf verisini yükle" buraya GET (JSONP) ile gelir ---- */
function doGet(e) {
  var name = e.parameter.sheet;
  var cb   = e.parameter.callback;   // JSONP callback adı
  var resp;

  if (!HEADERS[name]) {
    resp = { ok: false, error: 'bilinmeyen sheet: ' + name };
  } else {
    var sh = getSheet_(name);
    var values = sh.getDataRange().getValues();
    if (values.length > 0) values.shift();   // başlık satırını çıkar
    resp = { ok: true, rows: values };
  }

  var json = JSON.stringify(resp);
  if (cb) {
    // JSONP: tarayıcıda CORS sorununu tamamen ortadan kaldırır
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---- yardımcılar ---- */
function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS[name]);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
