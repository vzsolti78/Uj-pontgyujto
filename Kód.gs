// App meta / PWA konstansok
const APP_NAME       = 'Pontgyűjtő';
const APP_SHORT_NAME = 'Pontgyűjtő';
const THEME_COLOR    = '#3b82f6';
const ICON_URL       = 'https://res.cloudinary.com/dml7b81n6/image/upload/v1765280307/Pontgyujto_ikon_upsfe1.png';
const POINT_EXPIRATION_DAYS = 90;
const PROPS = PropertiesService.getScriptProperties();
const SPREADSHEET_ID_MAIN  = PROPS.getProperty('SPREADSHEET_ID_MAIN');
const SPREADSHEET_ID_ADMIN = PROPS.getProperty('SPREADSHEET_ID_ADMIN');

function getMainSpreadsheet_() {
  if (!SPREADSHEET_ID_MAIN) {
    throw new Error('SPREADSHEET_ID_MAIN nincs beállítva.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID_MAIN);
}

function getAdminSpreadsheet_() {
  if (!SPREADSHEET_ID_ADMIN) {
    throw new Error('SPREADSHEET_ID_ADMIN nincs beállítva.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID_ADMIN);
}

function doGet(e) {
  e = e || {};
  var params = e.parameter || {};

  // Manifest kérés kezelése (?manifest=1)
  if (String(params.manifest || '') === '1') {
    var manifest = {
      name: APP_NAME,
      short_name: APP_SHORT_NAME,
      display: 'standalone',
      start_url: './',
      scope: './',
      background_color: '#ffffff',
      theme_color: THEME_COLOR,
      icons: [
        { src: ICON_URL, sizes: '192x192', type: 'image/png', purpose: 'any maskable' },
        { src: ICON_URL, sizes: '512x512', type: 'image/png', purpose: 'any maskable' }
      ]
    };
    return ContentService
      .createTextOutput(JSON.stringify(manifest))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Fájl tartalom lekérése, ha fileId paraméter van
  if (params.fileId) {
    return getFileContent(params.fileId);
  }

  var page = params.page || 'index';
  var template = HtmlService.createTemplateFromFile(page);

  template.homeUrl = ScriptApp.getService().getUrl();
  template.osszesitokUrl = ScriptApp.getService().getUrl() + '?page=osszesitok';

  return template.evaluate()
    .setTitle(APP_NAME)
    .setFaviconUrl(ICON_URL)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getFileContent(fileId) {
  try {
    // Az azonosítóval megadott fájl tartalmának olvasása
    var file = DriveApp.getFileById(fileId);
    var content = file.getBlob().getDataAsString(); // A fájl tartalmát szövegként olvassuk be
    return ContentService.createTextOutput(content).setMimeType(ContentService.MimeType.TEXT);
  } catch (e) {
    return ContentService.createTextOutput("Hiba történt: " + e.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function addPointCredit(pointValue, pointReason) {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('Pontok');
  var archivSheet = ss.getSheetByName('archiv');
  var lastYearSheet = ss.getSheetByName('utolsó1év');

  // Ha nincs archiv sheet, létrehozzuk, és fejléceket adunk hozzá
  if (!archivSheet) {
    archivSheet = ss.insertSheet('archiv');
    archivSheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
  }

  // Ha nincs utolsó1év sheet, létrehozzuk, és fejléceket adunk hozzá
  if (!lastYearSheet) {
    lastYearSheet = ss.insertSheet('utolsó1év');
    lastYearSheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
  }

  // Dátum formázása
  var now = new Date();
  var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm:ss');

  // Új sor összeállítása a Pontok sheet számára:
  // A: Dátum, B: pontValue, C: "pont jóváírás", D: "Jóváírás", E: pointReason
  var newRow = [formattedDate, pointValue, "pont jóváírás", "Jóváírás", pointReason];
  sheet.appendRow(newRow);

  // Archiv és utolsó1év lapokhoz hasonló formátumban adjuk hozzá a sort
  // Itt is betartjuk a 5-oszlopos struktúrát: Dátum, Pontok, Büntetést kiszabta, Büntetési kategória, Engedetlenség típusa
  // Mivel ez jóváírás, nem egy konkrét büntető fél, tegyük "Rendszer" vagy "Jóváírás" szót a 3. és 4. mezőbe:
  archivSheet.appendRow([now, pointValue, "Rendszer", "Jóváírás", pointReason]);
  lastYearSheet.appendRow([now, pointValue, "Rendszer", "Jóváírás", pointReason]);

  // Pontok G1 cellájának frissítése, ugyanúgy, mint a addPoints függvényben
  var range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1); // B oszlop (Összesített pontok)
  var sum = range.getValues().reduce(function(acc, row) {
    return acc + (row[0] || 0);
  }, 0);
  sheet.getRange('G1').setValue(sum);

  // Balance frissítése az új pontok alapján
  updateBalanceFromLastYear();
  updateBalanceSheet();
}

function createTasksForPenalty(itemName, purchaseCategory) {
  // Ellenőrizzük, hogy a Tasks API engedélyezve van-e
  if (!Tasks) {
    throw new Error('A Google Tasks API nincs engedélyezve a projektben.');
  }

  // Feladatlisták lekérése
  var taskLists = Tasks.Tasklists.list().items;
  if (!taskLists) {
    throw new Error('Nem találhatók feladatlisták.');
  }

  // Megkeressük az 'Első értesítés' és 'Második értesítés' listákat
  var firstNotificationList = taskLists.find(function(list) {
    return list.title === 'Első értesítés';
  });
  var secondNotificationList = taskLists.find(function(list) {
    return list.title === 'Második értesítés';
  });

  if (!firstNotificationList || !secondNotificationList) {
    throw new Error('Nem találhatók a szükséges feladatlisták.');
  }

  // Határidők beállítása
  var today = new Date();

  // Első értesítés: 14 nap múlva, reggel 6 órakor
  var firstDeadline = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 14, 6, 0, 0);

  // Második értesítés: 28 nap múlva, reggel 6 órakor
  var secondDeadline = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 28, 6, 0, 0);

  // Első értesítés feladat létrehozása
  var task1 = {
    title: '1. értesítés határideje a mai napon lejár!',
    due: firstDeadline.toISOString().split('.')[0] + 'Z', // Pontos idő
    notes: 'Vásárlási kategória: ' + purchaseCategory + '\n' +
           'Megvásárolt tétel: ' + itemName
  };

  // Második értesítés feladat létrehozása
  var task2 = {
    title: '2. értesítés határideje a mai napon lejár!',
    due: secondDeadline.toISOString().split('.')[0] + 'Z', // Pontos idő
    notes: 'Vásárlási kategória: ' + purchaseCategory + '\n' +
           'Megvásárolt tétel: ' + itemName
  };

  // Feladatok hozzáadása a megfelelő listákhoz
  Tasks.Tasks.insert(task1, firstNotificationList.id);
  Tasks.Tasks.insert(task2, secondNotificationList.id);

  Logger.log('Értesítések sikeresen létrehozva a Google Tasks-ben.');
}

function sendPushbulletNotification(title, body) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('PUSHBULLET_API_KEY');

  if (!apiKey) {
    throw new Error('A Pushbullet API kulcs nincs beállítva a Script Properties-ben.');
  }

  var url = 'https://api.pushbullet.com/v2/pushes';
  var payload = {
    type: 'note',
    title: title,
    body: body
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Access-Token': apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error('Pushbullet értesítés küldése sikertelen: ' + response.getContentText());
  }
}

function getBalanceOnly() {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');
  return balanceSheet.getRange('B3').getValue(); // Csak a B3 cella értékét adjuk vissza
}

/**
 * Betölti a 'speciális' lap adatait, és elkészíti a HTML-táblázatot.
 * A 3. oszlop (C) tartalmazhat súgószöveget. Ha van benne szöveg,
 * a sor mellé megjelenik egy "?" ikon, amelyre kattintva megjelenik a súgó.
 */
function loadSpecialToolsPurchaseList() {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('speciális');
  if (!sheet) throw new Error('A "speciális" lap nem található.');

  // Teljes adattartomány lekérése
  var data = sheet.getDataRange().getValues();

  // HTML táblázat fejléce
  var html = '<table id="specialToolsPurchaseTable" class="table"><thead><tr>';
  html += '<th>Eszköz megnevezése</th>';
  html += '<th>Pont érték</th>';
  html += '<th>Darabszám</th>';
  html += '<th>Összköltség</th>';
  html += '<th>Vásárlás</th>';
  html += '</tr></thead><tbody>';

  // Az első sor (i=0) feltételezzük, hogy a táblázat fejléc, így i=1-től indulunk
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // row[0] = Eszköz megnevezése (A oszlop)
    // row[1] = Pont érték (B oszlop)
    // row[2] = Súgószöveg (C oszlop) - ha van

    // Új <tr> kezdete
    html += '<tr>';

    // 1) Eszköz megnevezése + (ha van) súgóikon
    var helpText = row[2] || ''; // C oszlop
    if (helpText) {
      // Ha van súgószöveg, megjelenítjük a "?" ikont
      var escapedHelpText = escapeForHTMLAttribute(helpText);
      html += '<td>' + row[0] +
              ' <span class="help-icon" onclick="showHelpText(event, \'' + escapedHelpText + '\')">?</span></td>';
    } else {
      // Ha nincs súgószöveg, csak a megnevezést írjuk ki
      html += '<td>' + row[0] + '</td>';
    }

    // 2) Pont érték
    html += '<td>' + row[1] + ' pont</td>';

    // 3) Darabszám input
    html += '<td>' +
            '<input type="number" class="quantity-input" ' +
            ' id="special-tools-quantity-' + i + '"' +
            ' style="width: 50px;"' +
            ' min="0" data-price="' + row[1] + '"' +
            ' oninput="calculateSpecialToolsCost(' + i + ')">' +
            '</td>';

    // 4) Összköltség cella (kezdésként 0 pont)
    html += '<td id="special-tools-cost-' + i + '">0 pont</td>';

    // 5) Vásárlás gomb
    html += '<td>' +
            '<button class="buy-button" onclick="buySpecialTool(' + i + ')">Megveszem</button>' +
            '</td>';

    // Sor lezárása
    html += '</tr>';
  }

  html += '</tbody></table>';
  return html;
}

/**
 * Segédfüggvény a speciális karakterek escape-eléséhez HTML-attribútumokba.
 * Ezt már megtalálod a kódban, de legyen itt is a teljesség kedvéért.
 */
function escapeForHTMLAttribute(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/'/g, '&#39;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r');
}

function loadCardPurchaseList() {
  var sheet = getMainSpreadsheet_().getSheetByName('kártya');
  var data = sheet.getDataRange().getValues();
  
  var html = '<table id="cardPurchaseTable" class="table"><thead><tr>';
  html += '<th>Kártya megnevezése</th>';
  html += '<th>Pont érték</th>';
  html += '<th>Darabszám</th>';
  html += '<th>Összköltség</th>';
  html += '<th>Vásárlás</th>';
  html += '</tr></thead><tbody>';

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    html += '<tr>';
    html += '<td>' + row[0] + '</td>';  // Kártya neve
    html += '<td>' + row[1] + ' pont</td>';  // Pont érték

    // Darabszám input mező
    html += '<td><input type="number" class="quantity-input" id="card-quantity-' + i + '" style="width: 50px;" min="0" data-price="' + row[1] + '" oninput="calculateCardCost(' + i + ')"></td>';

    // Összköltség cella
    html += '<td id="card-cost-' + i + '">0 pont</td>';

    // Vásárlás gomb
    html += '<td><button class="buy-button" onclick="buyCard(' + i + ')">Megveszem</button></td>';
    html += '</tr>';
  }

  html += '</tbody></table>';
  return html;
}

function deductBalanceAndProcessSale(cost, spreadsheetRowIndex, quantities) {
  var ss = getMainSpreadsheet_(); // Helyettesítsd a saját Spreadsheet ID-val
  var balanceSheet = ss.getSheetByName('Balance');
  var priceListSheet = ss.getSheetByName('vásárlás');
  var soldSheet = ss.getSheetByName('eladott'); // 'eladott' lap

  // >>> ÚJ LOGIKA: Dátum, B1, B2, B4 külön kezelése <<<
  var today = new Date();
  var dayOfMonth = today.getDate(); // Megnézzük, hanyadika van

  var prevMonthBalance = Number(balanceSheet.getRange('B1').getValue()) || 0; // Előző havi pontok (bruttó)
  var currMonthBalance = Number(balanceSheet.getRange('B2').getValue()) || 0; // Aktuális havi pontok (bruttó)
  var usedBalance      = Number(balanceSheet.getRange('B4').getValue()) || 0; // Felhasznált pontok
  var lostPoints       = Number(balanceSheet.getRange('B5').getValue()) || 0; // Elveszett pontok

  // Felhasználható egyenleg: B1 + B2 - B4 - B5
  var totalBalance = prevMonthBalance + currMonthBalance - usedBalance - lostPoints;
  if (totalBalance < 0) totalBalance = 0;

  // Ellenőrzés: van-e elég pont a vásárláshoz
  if (totalBalance < cost) {
    throw new Error('Nincs elég egyenleg a vásárláshoz.');
  }

  // NEM nyúlunk B1-hez és B2-höz, csak a felhasznált egyenleget növeljük
  var newUsedBalance = usedBalance + cost;
  balanceSheet.getRange('B4').setValue(newUsedBalance);

  // Új felhasználható egyenleg (B1 + B2 - B4 - B5)
  var newTotalBalance = prevMonthBalance + currMonthBalance - newUsedBalance - lostPoints;
  if (newTotalBalance < 0) newTotalBalance = 0;
  balanceSheet.getRange('B3').setValue(newTotalBalance);

  // >>> Az alábbi részek az eredeti logika szerint: <<<
  // Megvásárolt tétel feldolgozása
  var rowData = priceListSheet.getRange(spreadsheetRowIndex, 1, 1, priceListSheet.getLastColumn()).getValues()[0];
  var headers = priceListSheet.getRange(1, 2, 1, 4).getValues()[0]; // Fejlécek (1. sor, 2-5 oszlop)
  
  var itemName = rowData[0]; // Büntetőeszköz megnevezése

  // Sor törlése az 'vásárlás' lapról
  priceListSheet.deleteRow(spreadsheetRowIndex);

  // Összköltség és vásárlás dátuma
  var totalCost = cost;
  var purchaseDate = new Date();

  // Új sor összeállítása az 'eladott' munkafüzethez
  var newRow = [];
  newRow.push(itemName); // Büntetőeszköz megnevezése

  // Ár és mennyiség hozzáadása az egyes oszlopokhoz a kívánt formátumban
  for (var j = 1; j <= 4; j++) {
    var priceCell = rowData[j];
    var price = priceCell;

    // Ha az ár tartalmazza a 'pont' szót, eltávolítjuk
    if (typeof priceCell === 'string' && priceCell.toLowerCase().includes('pont')) {
      price = parseFloat(priceCell.toLowerCase().replace('pont', '').trim());
    }

    var quantity = quantities[j] || 0;
    var formattedData = price + ' pont x ' + quantity;

    // Ha a mennyiség nulla, hagyjuk üresen
    if (quantity == 0) {
      formattedData = '';
    }

    newRow.push(formattedData);
  }

  newRow.push(totalCost + ' pont'); // Összköltség
  newRow.push(purchaseDate); // Vásárlás dátuma

  // Ha az 'eladott' lap üres, adjuk hozzá a megfelelő fejléceket
  if (soldSheet.getLastRow() === 0) {
    soldSheet.appendRow([
      'Büntetőeszköz megnevezés',
      '1 óra időtartamra',
      'Reggeltől fürdésig (10 óra)',
      '24 órára, szigorúan levétel nélkül',
      '5 munkanapra, reggeltől fürdésig (50 óra)',
      'Összköltség',
      'Vásárlás dátuma'
    ]);
  }

  // Új sor hozzáadása az 'eladott' laphoz
  soldSheet.appendRow(newRow);

  // Részletes vásárlási információk összegyűjtése a Tasks számára
  var purchaseDetails = '';
  for (var j = 1; j <= 4; j++) {
    var quantity = quantities[j] || 0;
    if (quantity > 0) {
      var header = headers[j - 1];
      var priceCell = rowData[j];
      var price = priceCell;

      if (typeof priceCell === 'string' && priceCell.toLowerCase().includes('pont')) {
        price = parseFloat(priceCell.toLowerCase().replace('pont', '').trim());
      }

      purchaseDetails += header + ': ' + quantity + ' db (' + price + ' pont/db)\n';
    }
  }

  // Pushbullet értesítés küldése (opcionális)
  try {
    var message = 'Új büntetés vásárlás történt.\n' +
                  'Tétel: ' + itemName + '\n' +
                  'Összköltség: ' + totalCost + ' pont\n' +
                  purchaseDetails;

    sendPushbulletNotification('Új büntetés vásárlás', message);
  } catch (error) {
    Logger.log('Hiba a Pushbullet értesítés küldésekor: ' + error.message);
  }

  // Visszaadjuk az új egyenleget, a tétel nevét és a vásárlás részleteit
  return {
    newBalance: newTotalBalance,
    itemName: itemName,
    purchaseDetails: purchaseDetails
  };
}

function deductBalanceAndProcessSpecialToolPurchase(cost, rowIndex, quantity) {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');
  var specialToolsSheet = ss.getSheetByName('speciális');
  var soldSheet = ss.getSheetByName('eladott');

  // Felhasználható egyenleg számítása: B1 + B2 - B4 - B5
  var previousMonthBalance = Number(balanceSheet.getRange('B1').getValue()) || 0;
  var currentMonthBalance  = Number(balanceSheet.getRange('B2').getValue()) || 0;
  var usedBalance          = Number(balanceSheet.getRange('B4').getValue()) || 0;
  var lostPoints           = Number(balanceSheet.getRange('B5').getValue()) || 0;

  var currentBalance = previousMonthBalance + currentMonthBalance - usedBalance - lostPoints;
  if (currentBalance < 0) currentBalance = 0;

  if (currentBalance >= cost) {
    // Frissítjük a felhasznált egyenleget
    var newUsedBalance = usedBalance + cost;
    balanceSheet.getRange('B4').setValue(newUsedBalance);

    // Új egyenleg kiszámítása és frissítése (B1 + B2 - B4 - B5)
    var newTotalBalance = previousMonthBalance + currentMonthBalance - newUsedBalance - lostPoints;
    if (newTotalBalance < 0) newTotalBalance = 0;
    balanceSheet.getRange('B3').setValue(newTotalBalance);

    // Vásárlási adatok feldolgozása
    var rowData = specialToolsSheet.getRange(rowIndex + 1, 1, 1, specialToolsSheet.getLastColumn()).getValues()[0];
    var itemName = rowData[0]; // A oszlop: megnevezés

    // Eladott tételek hozzáadása
    var purchaseDate = new Date();
    var newRow = [
      itemName,
      quantity + ' db',
      '', '', '',
      cost + ' pont',
      purchaseDate
    ];

    // Fejléc ellenőrzése és új sor beszúrása
    if (soldSheet.getLastRow() === 0) {
      soldSheet.appendRow([
        'Megnevezés',
        'Mennyiség',
        '', '', '',
        'Összköltség',
        'Vásárlás dátuma'
      ]);
    }
    soldSheet.appendRow(newRow);

    // Pontok frissítése (B7 cella)
    var currentPoints = balanceSheet.getRange('B7').getValue() || 0;
    var updatedPoints = currentPoints + cost;
    balanceSheet.getRange('B7').setValue(updatedPoints);

    // Vásárlási részletek összeállítása
    var unitPrice = (quantity > 0) ? (cost / quantity) : 0;
    var purchaseDetails = 'Mennyiség: ' + quantity + ' db\n' +
                          'Egységár: ' + unitPrice + ' pont/db';

    // >>> Itt a kért új logika <<<
    // Ha a megvásárolt tétel neve 'Automatikus napi büntetés',
    // akkor adjuk hozzá a vásárolt darabszámot a B8 cella értékéhez.
if (itemName === 'Automatikus napi büntetés') {
  var napBalanceSS = getAdminSpreadsheet_();
  var napBalanceSheet = napBalanceSS.getSheetByName('Napbalance');
  var oldVal = napBalanceSheet.getRange('B1').getValue() || 0;
  var newVal = oldVal + quantity;
  napBalanceSheet.getRange('B1').setValue(newVal);
}

    // Visszatérés a frissített adatokkal
    return {
      newBalance: newTotalBalance,
      itemName: itemName,
      purchaseDetails: purchaseDetails
    };

  } else {
    throw new Error('Nincs elég egyenleg a vásárláshoz.');
  }
}

function deductBalanceAndProcessCardPurchase(cost, rowIndex, quantity) {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');
  var cardSheet = ss.getSheetByName('kártya');
  
  var soldSheet = ss.getSheetByName('eladott'); // 'eladott' lap
  if (!soldSheet) {
    // Ha az 'eladott' munkalap nem létezik, hozzunk létre egy újat és adjuk hozzá a fejlécet
    soldSheet = ss.insertSheet('eladott');
    soldSheet.appendRow(['Kártya típus', 'Darabszám', '', '', '', 'Összköltség', 'Vásárlás dátuma']);
  }

  // Felhasználható egyenleg számítása: B1 + B2 - B4 - B5
  var previousMonthBalance = Number(balanceSheet.getRange('B1').getValue()) || 0;
  var currentMonthBalance  = Number(balanceSheet.getRange('B2').getValue()) || 0;
  var usedBalance          = Number(balanceSheet.getRange('B4').getValue()) || 0;
  var lostPoints           = Number(balanceSheet.getRange('B5').getValue()) || 0;

  var currentBalance = previousMonthBalance + currentMonthBalance - usedBalance - lostPoints;
  if (currentBalance < 0) currentBalance = 0;

  if (currentBalance >= cost) {
    // Frissítjük a felhasznált egyenleget
    var newUsedBalance = usedBalance + cost;
    balanceSheet.getRange('B4').setValue(newUsedBalance);

    // Új egyenleg (B1 + B2 - B4 - B5)
    var newTotalBalance = previousMonthBalance + currentMonthBalance - newUsedBalance - lostPoints;
    if (newTotalBalance < 0) newTotalBalance = 0;
    balanceSheet.getRange('B3').setValue(newTotalBalance);

    var rowData = cardSheet.getRange(rowIndex + 1, 1, 1, cardSheet.getLastColumn()).getValues()[0];
    var itemName = rowData[0];  // Kártya neve (A oszlop)

    var purchaseDate = new Date();
    var newRow = [];
    newRow.push(itemName);  // Kártya neve (A oszlop)
    newRow.push(quantity + ' db');  // Vásárolt darabszám (B oszlop)
    newRow.push('');  // Üres C-E oszlop
    newRow.push('');  
    newRow.push('');  
    newRow.push(cost + ' pont');  // Összköltség (F oszlop)
    newRow.push(Utilities.formatDate(purchaseDate, Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm'));  // Vásárlás dátuma (G oszlop)

    // Ha az 'eladott' lap üres, adjuk hozzá a fejléceket
    if (soldSheet.getLastRow() === 0) {
      soldSheet.appendRow(['Kártya típus', 'Darabszám', '', '', '', 'Összköltség', 'Vásárlás dátuma']);
    }

    // Új sor hozzáadása az eladott laphoz
    soldSheet.appendRow(newRow);

    // Pushbullet értesítés küldése
    try {
      var message = 'Új kártya vásárlás történt.\n' +
                    'Tétel: ' + itemName + '\n' +
                    'Mennyiség: ' + quantity + ' db\n' +
                    'Összköltség: ' + cost + ' pont';

      sendPushbulletNotification('Új kártya vásárlás', message);
    } catch (error) {
      Logger.log('Hiba a Pushbullet értesítés küldésekor: ' + error.message);
    }

    // Visszaadjuk az új egyenleget és a tétel nevét
    return { newBalance: newTotalBalance, itemName: itemName };

  } else {
    throw new Error('Nincs elég egyenleg a vásárláshoz.');
  }
}

function getPreviousMonthPoints() {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('utolsó1év'); // Az 'utolsó1év' munkalap
  if (!sheet) return 0;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); // Dátum és pontok oszlopok
  var today = new Date();
  var previousMonth = today.getMonth() - 1; // Előző hónap (0-al kezdődik)
  var currentYear = today.getFullYear();

  if (previousMonth < 0) {
    previousMonth = 11; // Ha január, akkor az előző hónap december
    currentYear -= 1;
  }

  // Összesített pontok számítása az előző hónapra
  var previousMonthPoints = data.reduce(function(acc, row) {
    var rowDate = new Date(row[0]); // Dátum az első oszlopban
    if (rowDate.getMonth() === previousMonth && rowDate.getFullYear() === currentYear) {
      return acc + (row[1] || 0); // Pontok a második oszlopban
    }
    return acc;
  }, 0);

  return previousMonthPoints;
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var name = sheet.getName();
  
  // Ha a Pontok vagy vásárlás sheet-en van változás, frissítjük a Pontbeváltás sheet-et
  if (name === 'Pontok' || name === 'vásárlás') {
    updateRedemptionSheet();
  }
  
  // Balance sheet frissítése, ha a pontok vagy vásárlás változik
  if (name === 'Pontok' || name === 'vásárlás') {
    updateBalanceSheet();  // Balance sheet frissítése a B3 cellával együtt
  }
}

// Pontbeváltás sheet frissítése
function updateRedemptionSheet() {
  var ss = getMainSpreadsheet_();
  var redemptionSheet = ss.getSheetByName('Pontbeváltás');
  var balanceSheet = ss.getSheetByName('Balance');

  // Előző hónap pontok a Balance sheet B1 cellájából
  var previousMonthPoints = balanceSheet.getRange('B1').getValue();
  
  // Aktuális hónap pontok a Balance sheet B2 cellájából
  var currentMonthPoints = balanceSheet.getRange('B2').getValue();
  
  // Felhasznált egyenleg a Balance sheet B4 cellájából
  var usedBalance = balanceSheet.getRange('B4').getValue();
  
  // Friss adatok beírása a Pontbeváltás sheet megfelelő celláiba
  redemptionSheet.getRange('A1').setValue("Előző hónap pontjai: " + previousMonthPoints);
  redemptionSheet.getRange('A2').setValue("Aktuális hónap pontjai: " + currentMonthPoints);
  redemptionSheet.getRange('A3').setValue("Felhasznált egyenleg: " + usedBalance);
}

function createAllTriggers() {
  // Napi pontok lezárása és napi összesítés 23:00-kor
  ScriptApp.newTrigger('updatePoints')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  // Heti pontok frissítése (az elmúlt hét újraszámolása)
  ScriptApp.newTrigger('updateWeeklyPoints')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  // Havi pontok frissítése (aktuális hónap újraszámolása)
  ScriptApp.newTrigger('updateMonthlyPoints')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  // Balance sheet automatikus frissítése (FIFO lejárat + B1–B5 + B3 kiszámolása)
  ScriptApp.newTrigger('updateBalanceSheet')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
}

function deleteAllTriggers() {
  // Összes trigger lekérése
  const triggers = ScriptApp.getProjectTriggers();

  // Trigger-ek törlése
  triggers.forEach((trigger) => {
    ScriptApp.deleteTrigger(trigger);
  });

  Logger.log("Minden trigger sikeresen törölve.");
}

function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    Logger.log('Handler function: ' + trigger.getHandlerFunction() + ', Trigger type: ' + trigger.getEventType());
  });
}

function updateBalanceSheet() {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');

  if (!balanceSheet) {
    Logger.log('Balance sheet not found!');
    return;
  }

  // 1) pontok lejártatása (FIFO, elveszett pontok frissítése)
  expireOldPoints();

  // 2) B1 és B2 frissítése az 'utolsó1év' lap alapján
  updateBalanceFromLastYear();

  var prevMonthBalance  = Number(balanceSheet.getRange('B1').getValue()) || 0; // előző hónapban szerzett pontok
  var currMonthBalance  = Number(balanceSheet.getRange('B2').getValue()) || 0; // aktuális hónapban szerzett pontok
  var usedBalance       = Number(balanceSheet.getRange('B4').getValue()) || 0; // felhasznált pontok
  var lostPoints        = Number(balanceSheet.getRange('B5').getValue()) || 0; // elveszett pontok

  // Felhasználható egyenleg = B1 + B2 - B4 - B5
  var totalBalance = prevMonthBalance + currMonthBalance - usedBalance - lostPoints;
  if (totalBalance < 0) {
    totalBalance = 0; // soha ne legyen mínusz
  }

  balanceSheet.getRange('B3').setValue(totalBalance);
}


function updateBalanceFromLastYear() {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');
  var lastYearSheet = ss.getSheetByName('utolsó1év'); // Az 'utolsó1év' sheet, ahol az adatok vannak

  var today = new Date();
  var currentMonth = today.getMonth();  // Aktuális hónap
  var currentYear = today.getFullYear();
  var prevMonth = currentMonth - 1;  // Előző hónap
  if (prevMonth < 0) {
    prevMonth = 11;  // Ha január, akkor az előző hónap december
  }

  var prevMonthSum = 0;
  var currentMonthSum = 0;

  // 'utolsó1év' adatok lekérése
  var data = lastYearSheet.getRange(2, 1, lastYearSheet.getLastRow() - 1, 2).getValues();
  
  data.forEach(function(row) {
    var date = new Date(row[0]);  // Dátum az A oszlopban
    var points = row[1] || 0;     // Pontok a B oszlopban

    // Összesítés az előző hónapra
    if (date.getMonth() === prevMonth && date.getFullYear() === currentYear) {
      prevMonthSum += points;
    }

    // Összesítés az aktuális hónapra
    if (date.getMonth() === currentMonth && date.getFullYear() === currentYear) {
      currentMonthSum += points;
    }
  });

  // Előző hónap egyenleg beírása a B1 cellába
  balanceSheet.getRange('B1').setValue(prevMonthSum);

  // Aktuális hónap egyenleg beírása a B2 cellába
  balanceSheet.getRange('B2').setValue(currentMonthSum);

  return `Előző hónap egyenlege: ${prevMonthSum} pont, Aktuális hónap egyenlege: ${currentMonthSum} pont.`;
}

// Törli az 1 évnél régebbi adatokat az 'utolsó1év' munkalapról
function cleanOldLastYearData() {
  var ss = getMainSpreadsheet_();
  var lastYearSheet = ss.getSheetByName('utolsó1év');
  if (!lastYearSheet) return;

  var data = lastYearSheet.getDataRange().getValues();
  var today = new Date();
  var oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate()); // Egy évvel ezelőtti dátum

  for (var i = data.length - 1; i >= 1; i--) { // Az adatok visszafelé történő ellenőrzése
    var eventDate = new Date(data[i][0]); // Dátum oszlop (A oszlop)
    if (eventDate < oneYearAgo) {
      lastYearSheet.deleteRow(i + 1); // A régi sorok törlése (a fejlécet kihagyva)
    }
  }
}

function expireOldPoints() {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');
  var lastYearSheet = ss.getSheetByName('utolsó1év');
  if (!lastYearSheet || lastYearSheet.getLastRow() < 2) {
    return 0; // nincs semmi, ami lejárhatna
  }

  // Ha nincs vagy 0 a lejárati idő, ne csináljon semmit
  var thresholdDays = (typeof POINT_EXPIRATION_DAYS !== 'undefined') ? POINT_EXPIRATION_DAYS : 90;
  if (thresholdDays <= 0) {
    return 0;
  }

  var now = new Date();
  var thresholdMillis = thresholdDays * 24 * 60 * 60 * 1000;
  var cutoffDate = new Date(now.getTime() - thresholdMillis);

  // Balance adatok
  var prevMonthBalance = Number(balanceSheet.getRange('B1').getValue()) || 0; // előző hónapban szerzett
  var currMonthBalance = Number(balanceSheet.getRange('B2').getValue()) || 0; // aktuális hónapban szerzett
  var usedBalance      = Number(balanceSheet.getRange('B4').getValue()) || 0; // felhasznált
  var lostAlready      = Number(balanceSheet.getRange('B5').getValue()) || 0; // eddig elveszett

  var totalEarned = prevMonthBalance + currMonthBalance;
  var totalConsumedBefore = usedBalance + lostAlready; // vásárlás + régebben lejárt

  // Ha nincs felhasználható pont, nem járhat le semmi
  if (totalEarned <= totalConsumedBefore) {
    return 0;
  }

  // 'utolsó1év' adatok
  var lastRow = lastYearSheet.getLastRow();
  var data = lastYearSheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A–E: Dátum, Pontok, Büntetést kiszabta, stb.

  var rows = [];
  var remainingToConsume = totalConsumedBefore;

  // 1. kör: FIFO elosztás a FELHASZNÁLT + EDDIG LEJÁRT pontokra
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var pts = Number(row[1]) || 0;
    if (pts <= 0) continue;

    var consumedFromRow = 0;
    if (remainingToConsume > 0) {
      consumedFromRow = Math.min(pts, remainingToConsume);
      remainingToConsume -= consumedFromRow;
    }
    var remaining = pts - consumedFromRow; // ennyi pont él még ebből a sorból

    rows.push({
      date: new Date(row[0]),
      points: pts,
      remaining: remaining,
      rawRow: row
    });
  }

  // 2. kör: az adott cutoffDate-nél régebbi, MÉG NEM felhasznált pontok lejártatása
  var lostThisRun = 0;
  var rowsToAppend = [];

  rows.forEach(function(info) {
    if (info.remaining <= 0) return; // ebből a sorból már minden el lett költve vagy már korábban lejárt
    if (info.date <= cutoffDate) {
      // Most jár le az innen maradt pont
      lostThisRun += info.remaining;

      // Az "elveszett pontok" lapra az EREDETI dátummal kerül át:
      rowsToAppend.push([
        info.rawRow[0],      // eredeti dátum
        info.remaining,      // most lejáró mennyiség
        info.rawRow[2],      // Büntetést kiszabta
        info.rawRow[3],      // Büntetési kategória
        info.rawRow[4]       // Engedetlenség típusa
      ]);

      info.remaining = 0;
    }
  });

  if (lostThisRun <= 0) {
    return 0; // most nem járt le semmi
  }

  // 3. "elveszett pontok" sheet frissítése (NEM törlünk az 'utolsó1év'-ből!)
  var lostSheet = ss.getSheetByName('elveszett pontok');
  if (!lostSheet) {
    lostSheet = ss.insertSheet('elveszett pontok');
    lostSheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
  }

  rowsToAppend.forEach(function(r) {
    lostSheet.appendRow(r);
  });

  // 4. Balance!B5 = összes elveszett pont (összegezve az "elveszett pontok" lapon)
  var lostLastRow = lostSheet.getLastRow();
  var totalLostPoints = 0;
  if (lostLastRow > 1) {
    var lostData = lostSheet.getRange(2, 2, lostLastRow - 1, 1).getValues(); // B oszlop
    totalLostPoints = lostData.reduce(function(acc, row) {
      return acc + (Number(row[0]) || 0);
    }, 0);
  }

  balanceSheet.getRange('B5').setValue(totalLostPoints);

  return lostThisRun;
}

function loadEladottList() {
  var sheet = getMainSpreadsheet_().getSheetByName('eladott');
  if (!sheet) throw new Error('Az "eladott" lap nem található.');

  var data = sheet.getDataRange().getValues();

  // Legújabb eladások felül
  var dataWithoutHeader = data.slice(1).reverse();

  // A 4 időtartam fejléce (ugyanaz, mint amit eddig külön oszlopokba írtál)
  var durationLabels = [
    '1 óra időtartamra / darab',
    'Reggeltől fürdésig (10 óra)',
    '24 órára, levétel nélkül',
    '5 munkanapra, reggeltől fürdésig'
  ];

  var html = '<table id="eladottTable" class="table" style="table-layout: fixed; width: 100%;">';
  html += '<thead><tr>';
  html += '<th style="border: 1px solid black; padding: 8px; width: 220px; word-wrap: break-word; white-space: normal;">Büntetőeszköz megnevezése</th>';
  html += '<th style="border: 1px solid black; padding: 8px; width: 320px; word-wrap: break-word; white-space: normal;">Eladott időtartam(ok)</th>';
  html += '<th style="border: 1px solid black; padding: 8px; width: 140px; word-wrap: break-word; white-space: normal;">Összköltség</th>';
  html += '<th style="border: 1px solid black; padding: 8px; width: 170px; word-wrap: break-word; white-space: normal;">Vásárlás dátuma</th>';
  html += '</tr></thead><tbody>';

  for (var i = 0; i < dataWithoutHeader.length; i++) {
    var row = dataWithoutHeader[i];

    // row[0] név, row[1..4] időtartamok, row[5] összköltség, row[6] dátum
    var soldParts = [];
    for (var j = 1; j <= 4; j++) {
      var cellText = (row[j] || '').toString().trim();
      if (cellText) {
        soldParts.push(
          '<div style="margin-bottom: 4px;"><strong>' + durationLabels[j - 1] + ':</strong> ' + cellText + '</div>'
        );
      }
    }

    // Ha valamiért nincs egyetlen eladott időtartam sem, akkor ezt a sort kihagyjuk
    if (soldParts.length === 0) continue;

    var purchaseDate = Utilities.formatDate(new Date(row[6]), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm');

    html += '<tr>';
    html += '<td style="border: 1px solid black; padding: 8px; width: 220px; word-wrap: break-word; white-space: normal;">' + row[0] + '</td>';
    html += '<td style="border: 1px solid black; padding: 8px; width: 320px;">' + soldParts.join('') + '</td>';
    html += '<td style="border: 1px solid black; padding: 8px; width: 140px;">' + row[5] + '</td>';
    html += '<td style="border: 1px solid black; padding: 8px; width: 170px;">' + purchaseDate + '</td>';
    html += '</tr>';
  }

  html += '</tbody></table>';
  return html;
}

function getBalance() {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');

  // Előző havi (B1) már nem kell a megjelenítéshez
  var currentMonthBalance = balanceSheet.getRange('B2').getValue(); // Aktuális havi egyenleg
  var spendableBalance = balanceSheet.getRange('B3').getValue();    // Felhasználható egyenleg

  // Két sorba tördelve, nagyobb olvashatóság miatt (HTML)
  var balanceText =
    `Aktuális hónap egyenlege: <strong>${currentMonthBalance} pont</strong><br>` +
    `Felhasználható egyenleged: <strong>${spendableBalance} pont</strong>`;

  return balanceText;
}

function loadPurchaseList() {
  var sheet = getMainSpreadsheet_().getSheetByName('vásárlás');
  var data = sheet.getDataRange().getValues();
  
  var html = '<table id="purchaseTable" class="table"><thead><tr>';
  
  // Fejléc generálása
  var headers = data[0];
  headers.forEach(function(header) {
    html += '<th>' + header + '</th>';
  });
  
  // Hozzáadjuk az Összköltség és Vásárlás oszlopokat
  html += '<th>Összköltség</th>';
  html += '<th>Vásárlás</th>';
  
  html += '</tr></thead><tbody>';
  
  for (var i = 1; i < data.length; i++) { // Az első sor a fejléc
    var row = data[i];
    var spreadsheetRowIndex = i + 1; // A táblázatban a sor indexe
    html += '<tr>';
    html += '<td>' + row[0] + '</td>';  // Termék neve
        
    var hasPurchasable = false; // Jelző, ha van legalább egy megvásárolható ár
        
    // Iterálunk a különböző árkategóriákon
    for (var j = 1; j <= 4; j++) { 
      var priceCell = row[j].toString().trim().toLowerCase();
          
      if (priceCell.includes('nem megvásárolható')) {
        html += '<td>Nem megvásárolható</td>';
      } else {
        var price = row[j];
        if (price.toString().toLowerCase().includes('pont')) {
          price = price.toString().replace(/pont/i, '').trim();
        }
        html += '<td>';
        html += price + ' pont ';
        html += '<input type="number" class="quantity-input" id="quantity-' + i + '-' + j + '" style="width: 40px;" min="0" data-price="' + price + '" data-row-index="' + i + '">';
        html += '</td>';
        hasPurchasable = true;
      }
    }
        
    if (hasPurchasable) {
      html += '<td id="cost-' + i + '">0 pont</td>';
    } else {
      html += '<td>N/A</td>';
    }
        
    // Vásárlás gomb és a táblázatbeli sor indexének átadása
    html += '<td><button class="buy-button" onclick="buyItem(' + i + ', ' + spreadsheetRowIndex + ')">Megveszem</button></td>';
    html += '</tr>';
  }

  html += '</tbody></table>';
  return html;
}

function getSpendableBalance() {
  var ss = getMainSpreadsheet_(); // Balance sheet ID
  var balanceSheet = ss.getSheetByName('Balance');
  
  var prevMonthBalance = balanceSheet.getRange('B1').getValue(); // Előző hónap egyenlege
  var currentMonthBalance = balanceSheet.getRange('B2').getValue(); // Aktuális hónap egyenlege
  var usedBalance = balanceSheet.getRange('B4').getValue(); // Felhasznált egyenleg
  var totalBalance = balanceSheet.getRange('B3').getValue(); // Összesített egyenleg (B3)

  return `Előző hónap egyenlege: ${prevMonthBalance} pont, Aktuális hónap egyenlege: ${currentMonthBalance} pont. Felhasználható egyenleged: ${totalBalance} pont.`;
}

function loadPriceList() {
  var sheet = getMainSpreadsheet_().getSheetByName('Árlista');
  var data = sheet.getDataRange().getValues();
  var backgrounds = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getBackgrounds();  // Cellák háttérszíne
  var fontColors = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getFontColors();    // Cellák szöveg színe
  var fontWeights = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getFontWeights();  // Cellák szöveg vastagsága
  var wrapStrategies = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getWrapStrategies(); // Cellák tördelési beállításai
  
  var columnWidths = [150, 100, 100, 100, 120];  // Alapértelmezett oszlopszélességek
  
  // HTML táblázat létrehozása
  var html = '<table class="table">';

  // Fejléc hozzáadása
  html += '<thead><tr>';
  for (var i = 0; i < data[0].length; i++) {
    html += '<th style="width: ' + columnWidths[i] + 'px;">' + data[0][i] + '</th>';
  }
  html += '</tr></thead>';

  // Adatok hozzáadása a táblázathoz
  html += '<tbody>';
  for (var i = 1; i < data.length; i++) {
    html += '<tr>';
    for (var j = 0; j < data[i].length; j++) {
      var backgroundColor = backgrounds[i][j];
      var fontColor = fontColors[i][j];
      var fontWeight = fontWeights[i][j];
      var wrap = wrapStrategies[i][j] === SpreadsheetApp.WrapStrategy.CLIP ? 'nowrap' : 'wrap';
      html += '<td style="background-color: ' + backgroundColor + '; color: ' + fontColor + '; font-weight: ' + fontWeight + '; width: ' + columnWidths[j] + 'px; white-space: ' + wrap + ';">' + data[i][j] + '</td>';
    }
    html += '</tr>';
  }
  html += '</tbody></table>';

  return html;
}

function getRedemptionInfo() {
  // A szöveg alkalmazkodjon a beállított lejárati időhöz
  return 'A fel nem használt pontok ' + POINT_EXPIRATION_DAYS + ' nap után elvesznek!';
}

function loadRedemptionData() {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');

  // Előző hónapról fennmaradt pontok (B1 cella)
  var previousMonthPoints = balanceSheet.getRange('B1').getValue();

  // Aktuális havi pontok a Balance sheet-ről (B2 cella)
  var monthlyPoints = balanceSheet.getRange('B2').getValue();

  // Felhasznált egyenleg (B4 cella) és jelenleg felhasználható egyenleg (B3 cella)
  var usedBalance = balanceSheet.getRange('B4').getValue(); // Felhasznált egyenleg
  var availableBalance = balanceSheet.getRange('B3').getValue(); // Felhasználható egyenleg

  // Pontbeváltási információ
  var redemptionInfo = getRedemptionInfo();

var lostPoints = balanceSheet.getRange('B5').getValue();

// A visszatérő objektumba add hozzá:
return {
  previousMonthPoints: previousMonthPoints,
  monthlyPoints: monthlyPoints,
  usedBalance: usedBalance,
  availableBalance: availableBalance,
  redemptionInfo: redemptionInfo,
  lostPoints: lostPoints   // <--- ÚJ
  };
}

function getMonthlyPoints() {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('utolsó1év'); // Most már az 'utolsó1év' sheet-ről gyűjtjük az adatokat
  if (!sheet) return { monthlyPoints: 0 };

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues(); // Az adatok lekérése (dátum, pontok stb.)
  var today = new Date();
  var currentMonth = today.getMonth() + 1; // Az aktuális hónap (0-al kezdődik)
  var currentYear = today.getFullYear();

  // Összesített pontok számítása az aktuális hónapra
  var monthlyPoints = data.reduce(function(acc, row) {
    var rowDate = new Date(row[0]); // Dátum az A oszlopban
    if (rowDate.getMonth() + 1 === currentMonth && rowDate.getFullYear() === currentYear) {
      return acc + (row[1] || 0); // Összesített pontok a B oszlopban
    }
    return acc;
  }, 0);

  return { monthlyPoints: monthlyPoints };
}

function cleanOldArchivedData() {
  var ss = getMainSpreadsheet_();
  var archivSheet = ss.getSheetByName('archiv');
  if (!archivSheet) return;

  var data = archivSheet.getDataRange().getValues();
  var today = new Date();
  var sevenDaysAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);

  for (var i = data.length - 1; i >= 1; i--) { // Kezdjük az utolsó sorral
    var eventDate = new Date(data[i][0]);
    if (eventDate < sevenDaysAgo) {
      archivSheet.deleteRow(i + 1); // Mivel az 1. sor a fejléc
    }
  }
}

function loadLast7DaysEvents() {
  var ss = getMainSpreadsheet_();
  var archivSheet = ss.getSheetByName('archiv');
  if (!archivSheet) return '';

  var data = archivSheet.getDataRange().getValues();
  var today = new Date();
  var sevenDaysAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);

  var filteredData = data.filter(function(row, index) {
    if (index === 0) return false; // Fejléc kihagyása
    var eventDate = new Date(row[0]);
    return eventDate >= sevenDaysAgo;
  });

  // Fordítsd meg a szűrt adatokat, hogy a legújabb események legyenek felül
  filteredData.reverse();

  var tableHtml = filteredData.map(function(row) {
    var formattedDate = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm:ss');
    return `<tr>
              <td>${formattedDate}</td>
              <td>${row[1]}</td>
              <td>${row[2]}</td>
              <td>${row[3]}</td>
              <td>${row[4]}</td>
            </tr>`;
  }).join('');

  return tableHtml || '<tr><td colspan="5">Nincs adat az elmúlt 7 napban.</td></tr>';
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function addPoints(points, punisher, punishmentCategory, disobedienceType) {
  try {
    var ss = getMainSpreadsheet_();
    var sheet = ss.getSheetByName('Pontok');
    var archivSheet = ss.getSheetByName('archiv'); // archiv lap
    var lastYearSheet = ss.getSheetByName('utolsó1év'); // utolsó1év lap
    var balanceSheet = ss.getSheetByName('Balance'); // Balance sheet

    // Ha nincs archiv sheet, létrehozzuk, és fejlécekkel töltjük fel
    if (!archivSheet) {
      archivSheet = ss.insertSheet('archiv');
      archivSheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
    }

    // Ha nincs Pontok sheet, létrehozzuk, és fejlécekkel töltjük fel
    if (!sheet) {
      sheet = ss.insertSheet('Pontok');
      sheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
      sheet.getRange('B1').setValue('Összesített pontok');
    }

    // Ha nincs utolsó1év sheet, létrehozzuk, és fejlécekkel töltjük fel
    if (!lastYearSheet) {
      lastYearSheet = ss.insertSheet('utolsó1év');
      lastYearSheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
    }

    // Adat hozzáadása: Dátum, Pontok, Büntetést kiszabta, Büntetési kategória, Engedetlenség típusa
    var date = new Date();
    var newRow = [date, points, punisher, punishmentCategory, disobedienceType];

    // Új sor hozzáadása a Pontok, Archiv, és utolsó1év lapokhoz
    sheet.appendRow(newRow);
    archivSheet.appendRow(newRow);
    lastYearSheet.appendRow(newRow);

    // Pontok összesítése a Pontok lapon
    var range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
    var sum = range.getValues().reduce((acc, row) => acc + (row[0] || 0), 0);
    sheet.getRange('G1').setValue(sum);

    // Frissítjük a Balance sheet-et az előző és aktuális hónap pontjainak számításához
    updateBalanceFromLastYear();

    // Hozzáadjuk az updateBalanceSheet() hívást a B3 cella frissítéséhez
    updateBalanceSheet();

    // Pontok táblázat megjelenítése
    var table = loadPointsTable();
    return { totalPoints: sum, tableHtml: table };

  } catch (error) {
    Logger.log("Error in addPoints: " + error.message);
    throw new Error("Hiba történt a pontok hozzáadása közben.");
  }
}

function loadPoints() {
  try {
    var ss = getMainSpreadsheet_();
    var sheet = ss.getSheetByName('Pontok');
    
    // Ellenőrizzük, hogy létezik-e a sheet és van-e benne adat
    if (!sheet || sheet.getLastRow() < 2) {
      return { totalPoints: 0, allPoints: 'Nincsenek pontok még.' };
    }

    // Pontok lekérése a G1 cellából
    var totalPoints = sheet.getRange('G1').getValue() || 0;  // G1 cella az összesített pontok

    // Pontok listájának lekérése (B oszlop)
    var allPoints = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
      .map(row => {
        var date = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm:ss');
        var punisher = row[2] || '';
        var category = row[3] || '';
        var disobedience = row[4] || '';
        var points = row[1] || 0;
        return `${date}: ${points} pont | ${punisher} | ${category} | ${disobedience}`;
      })
      .join(', ');
    
    return { totalPoints: totalPoints, allPoints: allPoints };
  } catch (error) {
    Logger.log("Error in loadPoints: " + error.message);
    throw new Error("Hiba történt a pontok betöltése közben.");
  }
}

function loadPointsTable() {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('Pontok'); // <-- szükséges!
  if (!sheet) return '';
  var data = sheet.getDataRange().getValues();
  var table = ''; 
  for (var i = data.length - 1; i >= 1; i--) { 
    var formattedDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm:ss');
    var punisher = data[i][2] || ''; // Büntetést kiszabta
    var punishmentCategory = data[i][3] || ''; // Büntetési kategória
    var disobedienceType = data[i][4] || ''; // Engedetlenség típusa
    var points = data[i][1] || 0; // Összesített pontok
    table += `<tr>
                <td>${formattedDate}</td>
                <td>${punisher}</td>
                <td>${punishmentCategory}</td>
                <td>${disobedienceType}</td>
                <td>${points}</td>
              </tr>`;
  }
  return table;
}

function updatePoints() {
  var ss = getMainSpreadsheet_();
  var pointsSheet = ss.getSheetByName('Pontok');
  var dailySummarySheet = ss.getSheetByName('Napi összesítések');
  var weeklySummarySheet = ss.getSheetByName('Heti összesítések');
  var monthlySummarySheet = ss.getSheetByName('Havi összesítések');

  if (!pointsSheet) return;

  // Összesített pontok a G1 cellából
  var totalPoints = Number(pointsSheet.getRange('G1').getValue()) || 0;

  // Aktuális dátum a napi összesítésekhez
  var today = new Date();
  dailySummarySheet.appendRow([today, totalPoints]);  // Dátum és összesített pontok hozzáadása

  // Heti és havi összesítések frissítése
  var weekNumber = getWeekNumber(today);
  var month = today.getMonth() + 1;
  updateWeeklySummary(weeklySummarySheet, weekNumber, totalPoints);
  updateMonthlySummary(monthlySummarySheet, month, totalPoints);

  // Ürítjük a Pontok sheet adatait (fejlécet nem töröljük)
  var lastRow = pointsSheet.getLastRow();
  if (lastRow > 1) {
    pointsSheet.getRange(2, 1, lastRow - 1, pointsSheet.getLastColumn()).clearContent(); // Sorok ürítése (fejléc marad)
  }

  // G1 cella nullázása
  pointsSheet.getRange('G1').setValue(0);

  // Pontgyűjtés után frissítjük a Balance sheet B3 celláját
  updateBalanceSheet();  // Itt hívjuk meg a függvényt, hogy frissítse a B3 cellát
}

function updateWeeklyPoints() {
  var ss = getMainSpreadsheet_();
  var dailySummarySheet = ss.getSheetByName('Napi összesítések');
  var weeklySummarySheet = ss.getSheetByName('Heti összesítések');

  var today = new Date();
  
  // A hétfőt hét 1. napjának tekintjük (ISO)
  // Megnézzük, ma hanyadik napja van a hétnek (hétfő=1, kedd=2, ..., vasárnap=7)
  // Javascriptben vasárnap = 0, hétfő=1, ... => ezzel kicsit át kell rendezni:
  // getDay() => vasárnap=0, hétfő=1, ...
  // Ezért bevezetünk egy konverziót:
  var dayOfWeek = today.getDay(); 
  if (dayOfWeek === 0) {
    dayOfWeek = 7; // vasárnap
  }

  // Az "elmúlt" hét hétfője = ma - (dayOfWeek + 6) nap, 
  // ha vasárnap van (dayOfWeek=7), akkor ma - 13 => egy héttel korábbit kapunk
  // Egyszerűbb, ha fixen 7 nappal visszamegyünk, és onnantól kiszámoljuk a hétfőt:

  // Példa: ha ma vasárnap, dayOfWeek=7, induljunk 7 napot vissza
  var endOfWeek = new Date(today);
  endOfWeek.setDate(endOfWeek.getDate() - dayOfWeek);  
  // ezzel a mostani hét "hétfőjére" jutunk
  // De nekünk az elmúlt hét vasárnapja kell => lépjünk vissza 1 napot:
  endOfWeek.setDate(endOfWeek.getDate() - 1);
  // Így endOfWeek az elmúlt hét vasárnap (23:59:59)
  endOfWeek.setHours(23, 59, 59, 999);

  // A múlt hét hétfője:
  var startOfWeek = new Date(endOfWeek);
  startOfWeek.setDate(startOfWeek.getDate() - 6); // a vasárnaptól vissza 6 nap a hétfő
  startOfWeek.setHours(0, 0, 0, 0);

  // Most kiszámoljuk, melyik hétnek a száma is ez (az elmúlt hét):
  var weekNumber = getWeekNumber(startOfWeek); 
  // getWeekNumber a te kódodban az ISO heti sorszámot adja vissza, a startOfWeek alapján megy

  // Heti pontok összesítése
  var data = dailySummarySheet.getDataRange().getValues();
  var weeklySum = data.reduce(function(acc, row) {
    var date = new Date(row[0]);
    if (date >= startOfWeek && date <= endOfWeek) {
      return acc + (row[1] || 0);
    }
    return acc;
  }, 0);

  // Frissítés a Heti összesítések lapon
  updateWeeklySummary(weeklySummarySheet, weekNumber, weeklySum);
}

function updateMonthlyPoints() {
  var ss = getMainSpreadsheet_();
  var dailySummarySheet = ss.getSheetByName('Napi összesítések');
  var monthlySummarySheet = ss.getSheetByName('Havi összesítések');

  // Időszak kezdete és vége (hónap első és utolsó napja)
  var today = new Date();
  var month = today.getMonth() + 1;  // Hónap száma
  var startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);  // Hónap első napja
  var endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);  // Hónap utolsó napja
  startOfMonth.setHours(0, 0, 0, 0);
  endOfMonth.setHours(23, 59, 59, 999);

  // Havi pontok összesítése
  var data = dailySummarySheet.getDataRange().getValues();
  var monthlySum = data.reduce((acc, row) => {
    var date = new Date(row[0]);
    if (date >= startOfMonth && date <= endOfMonth) {
      return acc + (row[1] || 0);
    }
    return acc;
  }, 0);

  // Havi összesítés frissítése, akkor is, ha a havi összeg nulla
  updateMonthlySummary(monthlySummarySheet, month, monthlySum);
}

function updateWeeklySummary(sheet, weekNumber, points) {
  // Feltételezzük, hogy a header már létezik:
  // Fejlécek: [Hét, Összesített pont, Úrnő jutalomra jogosult, Kisorsolva]

  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    // Ha még nincs adat a sheet-en, létrehozzuk a fejléceket
    sheet.appendRow(['Hét', 'Összesített pont', 'Úrnő jutalomra jogosult', 'Kisorsolva']);
  }

  // Ellenőrizzük, hogy létezik-e már az adott hét a sheetben:
  var data = sheet.getDataRange().getValues();
  var weekFound = false;
  var rowIndex = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === weekNumber) {
      weekFound = true;
      rowIndex = i + 1;
      break;
    }
  }

  if (weekFound) {
    // Hét már létezik, frissítjük a pontokat
    sheet.getRange(rowIndex, 2).setValue(points);
  } else {
    // Új sor beszúrása
    sheet.appendRow([weekNumber, points, '', '']);
    rowIndex = sheet.getLastRow();
  }

  // Úrnő jutalomra jogosult-e?
  var jutalom = points >= 500 ? 'Igen' : 'Nem';
  sheet.getRange(rowIndex, 3).setValue(jutalom);

  // Kisorsolva mező csak akkor változtatunk ha új a sor.
  // Ha új sor, akkor üresen hagyjuk a D oszlopot (Kisorsolva).
  // Ha régi sor, nem piszkáljuk, mert a felhasználó már állíthatta.

  // Ha kell, megbizonyosodunk róla, hogy a 4. oszlop létezik:
  var lastCol = sheet.getLastColumn();
  if (lastCol < 4) {
    sheet.insertColumnAfter(3);
    sheet.getRange(1,4).setValue('Kisorsolva');
  }
}

function setKisorsolva(weekNumber) {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('Heti összesítések');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === weekNumber) {
      // D oszlopba (4. oszlop) írjuk a jelzést, pl. 'TRUE'
      sheet.getRange(i+1,4).setValue('TRUE');
      return 'OK';
    }
  }

  throw new Error('A megadott hét nem található a Heti összesítések között.');
}

function updateMonthlySummary(sheet, month, points) {
  // Ellenőrizzük, hogy létezik-e a fejléc, ha nem, létrehozzuk
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    // Még nincs adat, létrehozzuk a fejlécet
    sheet.appendRow(['Hónap', 'Összesített pont', 'Jutalomra jogosult', 'Büntetésre jogosult', 'Büntetés kisorsolva', 'Jutalom kisorsolva']);
    lastRow = 1;
  }

  // Megkeressük a hónapot a táblában
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === month) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    // Új sor beszúrása ha nincs ilyen hónap
    sheet.appendRow([month, points, '', '', '', '']);
    rowIndex = sheet.getLastRow();
  } else {
    // Frissítjük a pontokat, hozzáadjuk az új pontokat a meglévőhöz
    var lastMonthPoints = sheet.getRange(rowIndex, 2).getValue();
    sheet.getRange(rowIndex, 2).setValue(lastMonthPoints + points);
  }

  // Pontok frissítése után megállapítjuk a jogosultságokat a végső pontszám alapján
  var finalPoints = sheet.getRange(rowIndex, 2).getValue();

  var jutalomJogosult = '';
  if (finalPoints <= 500) {
    jutalomJogosult = 'rabszolga';
  } else if (finalPoints <= 1999) {
    jutalomJogosult = 'senki';
  } else {
    jutalomJogosult = 'Úrnő';
  }

  var buntetesJogosult = '';
  if (finalPoints <= 500) {
    buntetesJogosult = 'Úrnő';
  } else if (finalPoints <= 1999) {
    buntetesJogosult = 'senki';
  } else {
    buntetesJogosult = 'rabszolga';
  }

  sheet.getRange(rowIndex, 3).setValue(jutalomJogosult);
  sheet.getRange(rowIndex, 4).setValue(buntetesJogosult);
  // Az 5. és 6. oszlop (Büntetés kisorsolva, Jutalom kisorsolva) értékeit nem piszkáljuk, hogy ne írjuk felül a felhasználói beállítást.
}

function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function loadDailySummary() {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('Napi összesítések');
  
  var data = sheet.getDataRange().getValues();
  
  return data.slice(1).map(row => {
    var formattedDate = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy.MM.dd');
    return `<tr><td>${formattedDate}</td><td>${row[1]}</td></tr>`;
  }).reverse().join(''); // Fordított sorrend
}

function loadWeeklySummary() {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('Heti összesítések');
  if (!sheet) {
    return '<tr><td colspan="4">Nincs adat</td></tr>';
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return '<tr><td colspan="4">Nincs adat</td></tr>';
  }

  var html = '';
  var rows = data.slice(1).reverse(); // Az adatok megfordítása
  
  rows.forEach(function(row) {
    var weekNumber = row[0];
    var points = row[1];
    var jutalom = row[2] || 'NEM';
    var kisorsolvaValue = (row[3] || '').toString().trim().toUpperCase();

    // Ha a pontszám 500 vagy több, jelenjen meg a checkbox, különben üres cella
    var checkboxCell = '';
    if (points >= 500) {
      var checked = (kisorsolvaValue === 'TRUE') ? 'checked disabled' : '';
      checkboxCell = `<input type="checkbox" ${checked} onclick="handleKisorsolvaClick(${weekNumber}, this)">`;
    }

    html += '<tr>';
    html += `<td>${weekNumber}. hét</td>`;
    html += `<td>${points}</td>`;
    html += `<td>${jutalom}</td>`;
    html += `<td>${checkboxCell}</td>`;
    html += '</tr>';
  });

  return html;
}

function loadMonthlySummary() {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('Havi összesítések');
  if (!sheet) {
    return '<table class="table"><thead><tr><th>Hónap</th><th>Összesített pont</th><th>Jutalomra jogosult</th><th>Büntetésre jogosult</th><th>Büntetés kisorsolva</th><th>Jutalom kisorsolva</th></tr></thead><tbody><tr><td colspan="6">Nincs adat</td></tr></tbody></table>';
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return '<table class="table"><thead><tr><th>Hónap</th><th>Összesített pont</th><th>Jutalomra jogosult</th><th>Büntetésre jogosult</th><th>Büntetés kisorsolva</th><th>Jutalom kisorsolva</th></tr></thead><tbody><tr><td colspan="6">Nincs adat</td></tr></tbody></table>';
  }

  var html = '<table class="table">';
  html += '<thead><tr>';
  html += '<th>Hónap</th>';
  html += '<th>Összesített pont</th>';
  html += '<th>Jutalomra jogosult</th>';
  html += '<th>Büntetésre jogosult</th>';
  html += '<th>Büntetés kisorsolva</th>';
  html += '<th>Jutalom kisorsolva</th>';
  html += '</tr></thead><tbody>';

  var rows = data.slice(1).reverse(); // Megfordítjuk, hogy a legújabb legyen felül
  
  rows.forEach(function(row) {
    var month = row[0];
    var points = row[1];
    var jutalomJogosult = row[2] || '';
    var buntetesJogosult = row[3] || '';
    var buntetesKisorsolva = (row[4] || '').toString().trim().toUpperCase();
    var jutalomKisorsolva = (row[5] || '').toString().trim().toUpperCase();

    // Meghatározzuk, hogy legyen-e checkbox
    var isCheckboxVisible = (points <= 500 || points >= 2000);

    var buntetesCheckbox = '';
    var jutalomCheckbox = '';

    if (isCheckboxVisible) {
      var buntetesChecked = (buntetesKisorsolva === 'TRUE') ? 'checked disabled' : '';
      var jutalomChecked = (jutalomKisorsolva === 'TRUE') ? 'checked disabled' : '';

      buntetesCheckbox = `<input type="checkbox" ${buntetesChecked} onclick="handleMonthlyKisorsolvaClick(${month}, 'buntetes', this)">`;
      jutalomCheckbox = `<input type="checkbox" ${jutalomChecked} onclick="handleMonthlyKisorsolvaClick(${month}, 'jutalom', this)">`;
    } else {
      // 501-1999 között nincs checkbox
      buntetesCheckbox = '';
      jutalomCheckbox = '';
    }

    html += '<tr>';
    html += `<td>${month}. hónap</td>`;
    html += `<td>${points}</td>`;
    html += `<td>${jutalomJogosult}</td>`;
    html += `<td>${buntetesJogosult}</td>`;
    html += `<td>${buntetesCheckbox}</td>`;
    html += `<td>${jutalomCheckbox}</td>`;
    html += '</tr>';
  });

  html += '</tbody></table>';
  return html;
}

function getMonthName(monthNumber) {
  var months = [
    'Január', 'Február', 'Március', 'Április', 'Május', 'Június',
    'Július', 'Augusztus', 'Szeptember', 'Október', 'November', 'December'
  ];
  return months[monthNumber - 1];
}

function loadAllSummaries() {
  try {
    var daily = loadDailySummary();
    var weekly = loadWeeklySummary();
    var monthly = loadMonthlySummary();
    return {
      daily: daily,
      weekly: weekly,
      monthly: monthly
    };
  } catch (error) {
    Logger.log("Error in loadAllSummaries: " + error.message);
    throw new Error("Hiba történt az összesítők betöltése közben.");
  }
}

// Úrnő Pontok Hozzáadása (ha szükséges további funkciókhoz)
function addUrnoPoints(points) {
  try {
    var ss = getMainSpreadsheet_();
    var sheet = ss.getSheetByName('UrnoPontok');
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(new Date());
    sheet.getRange(lastRow + 1, 2).setValue(points);
    
    // Összesített pontok frissítése
    var range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
    var sum = range.getValues().reduce((acc, row) => acc + row[0], 0);
    sheet.getRange('G1').setValue(sum);
    
    return { addedPoints: points, totalPoints: sum };
  } catch (error) {
    Logger.log("Error in addUrnoPoints: " + error.message);
    throw new Error("Hiba történt az Úrnő pontjainak hozzáadása közben.");
  }
}

function setMonthlyKisorsolva(month, type) {
  var ss = getMainSpreadsheet_();
  var sheet = ss.getSheetByName('Havi összesítések');
  var data = sheet.getDataRange().getValues();

  // type: 'buntetes' vagy 'jutalom'
  // 5. oszlop = Büntetés kisorsolva
  // 6. oszlop = Jutalom kisorsolva
  var columnIndex = (type === 'buntetes') ? 5 : 6; 

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === month) {
      sheet.getRange(i+1, columnIndex).setValue('TRUE');
      return 'OK';
    }
  }

  throw new Error('A megadott hónap nem található a Havi összesítések között.');
}

function recalcAllWeeklySums() {
  var ss = getMainSpreadsheet_();
  var dailySummarySheet = ss.getSheetByName('Napi összesítések');
  var weeklySummarySheet = ss.getSheetByName('Heti összesítések');

  // Ha nincs 'Heti összesítések' lap, létrehozzuk fejléccel
  if (!weeklySummarySheet) {
    weeklySummarySheet = ss.insertSheet('Heti összesítések');
    weeklySummarySheet.appendRow(['Hét', 'Összesített pont', 'Úrnő jutalomra jogosult', 'Kisorsolva']);
  }

  // Minden sort beolvasunk (az első sor a fejléc)
  var data = dailySummarySheet.getDataRange().getValues();
  if (data.length <= 1) {
    // Nincs érdemi adat
    Logger.log("Nincs adat a Napi összesítések lapon.");
    return;
  }

  // weeklySums objektumban gyűjtjük a hét => pont összefüggést
  var weeklySums = {};

  // Végigmegyünk a bejegyzéseken
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var date = new Date(row[0]);
    var dailyPoints = row[1] || 0;
    var wNum = getWeekNumber(date); // Ez a te getWeekNumber függvényed

    if (!weeklySums[wNum]) {
      weeklySums[wNum] = 0;
    }
    weeklySums[wNum] += dailyPoints;
  }

  // Most a weeklySums alapján minden hétre meghívjuk az updateWeeklySummary-t
  for (var weekNum in weeklySums) {
    updateWeeklySummary(weeklySummarySheet, parseInt(weekNum), weeklySums[weekNum]);
  }

  Logger.log("Heti összesítések újraszámolva és frissítve.");
}

function redeemLostPoints() {
  var ss = getMainSpreadsheet_();
  var balanceSheet = ss.getSheetByName('Balance');
  var pointsSheet  = ss.getSheetByName('Pontok');
  var archivSheet  = ss.getSheetByName('archiv');
  var lastYearSheet = ss.getSheetByName('utolsó1év');
  var lostSheet    = ss.getSheetByName('elveszett pontok');

  // archiv lap biztosítása
  if (!archivSheet) {
    archivSheet = ss.insertSheet('archiv');
    archivSheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
  }

  // utolsó1év lap biztosítása
  if (!lastYearSheet) {
    lastYearSheet = ss.insertSheet('utolsó1év');
    lastYearSheet.appendRow(['Dátum', 'Összesített pontok', 'Büntetést kiszabta', 'Büntetési kategória', 'Engedetlenség típusa']);
  }

  // Összes elveszett pont az "elveszett pontok" lapról
  var lostPoints = 0;
  if (lostSheet && lostSheet.getLastRow() > 1) {
    var lostData = lostSheet.getRange(2, 2, lostSheet.getLastRow() - 1, 1).getValues(); // B oszlop
    lostPoints = lostData.reduce(function(acc, row) {
      return acc + (Number(row[0]) || 0);
    }, 0);
  }

  if (lostPoints > 0) {
    // 1) B5 nullázása
    balanceSheet.getRange('B5').setValue(0);

    // 2) "elveszett pontok" lap kiürítése (fejléc marad)
    if (lostSheet && lostSheet.getLastRow() > 1) {
      lostSheet.getRange(2, 1, lostSheet.getLastRow() - 1, lostSheet.getLastColumn()).clearContent();
    }

    // 3) Elveszett pontok jóváírása az aktuális havi pontokhoz (B2)
    var currentMonthPoints = Number(balanceSheet.getRange('B2').getValue()) || 0;
    balanceSheet.getRange('B2').setValue(currentMonthPoints + lostPoints);

    // 4) Log bejegyzés a Pontok / archiv / utolsó1év lapokra
    var dateNow = new Date();
    var newRow = [dateNow, lostPoints, "Úrnő", "Úrnői jutalom", "Elveszett pont beváltása"];

    pointsSheet.appendRow(newRow);
    archivSheet.appendRow(newRow);
    lastYearSheet.appendRow(newRow);

    // 5) Pontok!G1 újraszámolása
    var range = pointsSheet.getRange(2, 2, pointsSheet.getLastRow() - 1, 1);
    var sum = range.getValues().reduce(function(acc, row) {
      return acc + (row[0] || 0);
    }, 0);
    pointsSheet.getRange('G1').setValue(sum);

    // 6) Balance újraszámolása
    updateBalanceFromLastYear();
    updateBalanceSheet();
  }

  // Ennyi pontot váltottál be
  return lostPoints;
}
