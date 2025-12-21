const EXPORT_CONFIG = {
  folderId: "1HZuq3DfhqlNVam8UANpQE2p2Vc_t7OoR",
  name: "Pontgyűjtő",
  timeFormat: "yyyy-MM-dd_HHmmss",
  filesToInsert: [
    { filename: "kód.html",   displayName: "Kód.gs" },
    { filename: "index.html", displayName: "index.html" },
    { filename: "java.html",  displayName: "java.html" },
    { filename: "css.html",   displayName: "css.html" }
  ]
};

function addLineNumbers(code) {
  // Normalizáljuk a sortöréseket, hogy a sorindex stabil legyen
  var normalized = (code || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var lines = normalized.split("\n");

  // Szélesség (pl. 1-999 sorig 3 számjegy)
  var width = String(lines.length).length;

  return lines.map(function(line, i) {
    var n = String(i + 1).padStart(width, "0");
    return n + " | " + line;
  }).join("\n");
}

function exportCodeToTxtFile() {
  var folder = DriveApp.getFolderById(EXPORT_CONFIG.folderId);

  var now = new Date();
  var dateString = Utilities.formatDate(now, Session.getScriptTimeZone(), EXPORT_CONFIG.timeFormat);

var fileName = EXPORT_CONFIG.name + "_" + dateString + ".txt";
var txtContent = EXPORT_CONFIG.name + "\n\n";

  // ===== 1) ÖSSZES FILE: SORSZÁMOZOTT NÉZET =====
  txtContent += "=== ÖSSZES FILE - SORSZÁMOZOTT NÉZET ===\n\n";

  EXPORT_CONFIG.filesToInsert.forEach(function(file) {
    var fileContent = HtmlService.createTemplateFromFile(file.filename).getRawContent();

    txtContent += file.displayName + ":\n";
    txtContent += "------------------------\n";
    txtContent += addLineNumbers(fileContent) + "\n\n";
  });

  // ===== 2) ÖSSZES FILE: TISZTA (MÁSOLHATÓ) NÉZET =====
  txtContent += "=== ÖSSZES FILE - TISZTA (MÁSOLHATÓ) NÉZET ===\n\n";

  EXPORT_CONFIG.filesToInsert.forEach(function(file) {
    var fileContent = HtmlService.createTemplateFromFile(file.filename).getRawContent();

    txtContent += file.displayName + ":\n";
    txtContent += "------------------------\n";
    txtContent += fileContent + "\n\n";
  });

  var outFile = folder.createFile(fileName, txtContent, MimeType.PLAIN_TEXT);
  Logger.log("A(z) '" + fileName + "' fájl létrehozva lett.");

  showFileLink(outFile.getUrl());
}

function exportCodeToRtfFile() {
  var folder = DriveApp.getFolderById(EXPORT_CONFIG.folderId);

  var now = new Date();
  var dateString = Utilities.formatDate(now, Session.getScriptTimeZone(), EXPORT_CONFIG.timeFormat);

var fileName = EXPORT_CONFIG.name + "_" + dateString + ".rtf";
rtfContent += "\\fs24\\b " + escapeRTF(EXPORT_CONFIG.name) + "\\b0\\par\n\\par\n";


  // Színpaletta
  var colorTable =
    "{\\colortbl ;" +
    "\\red255\\green0\\blue0;" +   // cf1 - kulcsszavak
    "\\red0\\green128\\blue0;" +   // cf2 - stringek
    "\\red0\\green0\\blue255;" +   // cf3 - kommentek
    "}\n";

  // RTF fejléc
  var rtfHeader =
    "{\\rtf1\\ansi\\deff0\n" +
    "{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}\n" +
    colorTable;

  var rtfContent = "";

  // Cím
  rtfContent += "\\fs24\\b " + escapeRTF(EXPORT_CONFIG.title) + "\\b0\\par\n\\par\n";

  // =========================================================
  // 1) ÖSSZES FILE – SORSZÁMOZOTT NÉZET
  // =========================================================
  rtfContent += "\\b " + escapeRTF("ÖSSZES FILE – SORSZÁMOZOTT NÉZET") + "\\b0\\par\n\\par\n";

  EXPORT_CONFIG.filesToInsert.forEach(function(file) {
    var fileContent = HtmlService.createTemplateFromFile(file.filename).getRawContent();
    var numberedContent = addLineNumbers(fileContent);

    rtfContent += "\\b " + escapeRTF(file.displayName) + "\\b0\\par\n";
    rtfContent += syntaxHighlightRTF(numberedContent) + "\\par\n\\par\n";
  });

  // =========================================================
  // 2) ÖSSZES FILE – TISZTA (MÁSOLHATÓ) NÉZET
  // =========================================================
  rtfContent += "\\page\n";
  rtfContent += "\\b " + escapeRTF("ÖSSZES FILE – TISZTA (MÁSOLHATÓ) NÉZET") + "\\b0\\par\n\\par\n";

  EXPORT_CONFIG.filesToInsert.forEach(function(file) {
    var fileContent = HtmlService.createTemplateFromFile(file.filename).getRawContent();

    rtfContent += "\\b " + escapeRTF(file.displayName) + "\\b0\\par\n";
    rtfContent += syntaxHighlightRTF(fileContent) + "\\par\n\\par\n";
  });

  var fullRtf = rtfHeader + rtfContent + "}";

  var outFile = folder.createFile(fileName, fullRtf, MimeType.RTF);
  Logger.log("A(z) '" + fileName + "' fájl létrehozva lett.");

  showFileLink(outFile.getUrl());
}

/**
 * Egyszerű szintaxis kiemelés RTF formátumban.
 * Ez egy alapvető példa, amely kulcsszavakat, stringeket és kommenteket emel ki.
 */
function syntaxHighlightRTF(code) {
  // Színek:
  // \cf1 - Piros (kulcsszavak)
  // \cf2 - Zöld (stringek)
  // \cf3 - Kék (kommentek)
  
  // Escape speciális karaktereket és kezeljük az Unicode karaktereket
  code = escapeRTF(code);
  
  // Kommentek kiemelése (// és /* */)
  code = code.replace(/(\/\/[^\n]*|\/\*[\s\S]*?\*\/)/g, "\\cf3 $1\\cf0 ");
  
  // Stringek kiemelése ("", '', ``)
  code = code.replace(/(".*?"|'.*?'|`.*?`)/g, "\\cf2 $1\\cf0 ");
  
  // Kulcsszavak kiemelése
  var keywords = ["function", "var", "let", "const", "if", "else", "for", "while", "return", "switch", "case", "break", "continue", "default", "new", "try", "catch", "finally", "throw", "class", "extends", "super", "import", "export", "from", "as", "async", "await", "typeof", "instanceof", "void", "delete", "in", "of"];
  var keywordPattern = new RegExp("\\b(" + keywords.join("|") + ")\\b", "g");
  code = code.replace(keywordPattern, "\\cf1 $1\\cf0 ");
  
  // Visszatérés sorok átalakítása RTF új sorra
  code = code.replace(/\n/g, "\\line ");
  
  return code;
}

/**
 * Speciális RTF karakterek és Unicode karakterek kezelése.
 * @param {string} text - A szöveg, amit RTF-re kell konvertálni.
 * @returns {string} - Az RTF-nek megfelelően escapelt szöveg.
 */

function escapeRTF(text) {
  // null/undefined → üres string, minden más → stringgé konvertálás
  if (text === null || typeof text === 'undefined') text = '';
  text = String(text);

  return text.replace(/\\/g, '\\\\')  // Backslash
             .replace(/{/g, '\\{')    // Open brace
             .replace(/}/g, '\\}')    // Close brace
             .replace(/[\u0080-\uFFFF]/g, function(chr) { // Unicode karakterek
               var code = chr.charCodeAt(0);
               return '\\u' + code + '?';
             });
}

function showFileLink(url) {
  var htmlContent = '<p>A fájl sikeresen létrejött/frissült. <a href="' + url + '" target="_blank">Kattints ide a megnyitáshoz</a>.</p>';
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
                              .setWidth(300)
                              .setHeight(100);
  
  try {
    // Próbáljuk meg megjeleníteni a linket egy HTML dialógusablakban
    // Itt feltételezzük, hogy a script Google Sheets-ből vagy Google Docs-ból fut
    var ui;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (e) {
      ui = DocumentApp.getUi();
    }
    ui.showModalDialog(htmlOutput, 'Fájl Elérhetősége');
  } catch (e) {
    // Ha nem Sheets sem Docs, akkor csak logoljuk a linket egyszer
    Logger.log("Fájl elérhetősége: " + url);
  }
}