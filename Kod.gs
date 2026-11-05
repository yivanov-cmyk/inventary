const SS = SpreadsheetApp.getActiveSpreadsheet();
const MASTER_SHEET_NAME = "Master_Stock";
const REPORT_SHEET_NAME = "Raport_Koncowy";
const TEAMS_COUNT = 15;
const SPEC_TEAMS = 5;
const TASK_LIMIT = 170;

const KODY = {
  "admin123": "ADMIN",
  "ekipa1": "Ekipa_1", "ekipa2": "Ekipa_2", "ekipa3": "Ekipa_3", "ekipa4": "Ekipa_4",
  "ekipa5": "Ekipa_5", "ekipa6": "Ekipa_6", "ekipa7": "Ekipa_7", "ekipa8": "Ekipa_8",
  "ekipa9": "Ekipa_9", "ekipa10": "Ekipa_10", "ekipa11": "Ekipa_11", "ekipa12": "Ekipa_12","ekipa13": "Ekipa_13", "ekipa14": "Ekipa_14", "ekipa15": "Ekipa_15", "ekipa16": "Ekipa_16","meza":"Ekipa_MEZA"
};
function stressTestSimultaneousSaving() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwiShWCO9BO22LOAVphZwuuG7Fd_vC-LOoP1SqzKel6SesxvgFqeRwrv_U6n1fqvK5H/exec";
  
  const testTeams = ["Ekipa_1", "Ekipa_2", "Ekipa_3","Ekipa_4", "Ekipa_5", "Ekipa_6"];
  const LIMIT = 170; 

  const requests = testTeams.map(teamName => {
    let dummyTasks = [];
    for (let i = 0; i < LIMIT; i++) {
      dummyTasks.push({
        loc: "0150.0215001", 
        sku: "TEST-SKU-" + i,
        qty: "10",
        sysQty: "10",
        team: teamName,
        masterRow: 2 + i, 
        comment: "TEST OBCIĄŻENIOWY",
        info: "1. LICZENIE"
      });
    }

    return {
      url: WEB_APP_URL,
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        action: "save_tasks", 
        data: dummyTasks
      }),
      muteHttpExceptions: true
    };
  });

  console.log("🚀 START TESTU: Wysyłanie równoległe...");
  const startTime = new Date();

  
  const responses = UrlFetchApp.fetchAll(requests);

  const endTime = new Date();
  const totalDuration = (endTime - startTime) / 1000;

  console.log("📊 WYNIKI:");
  console.log("Całkowity czas: " + totalDuration + "s");

  responses.forEach((res, index) => {
    console.log(`Team: ${testTeams[index]} | Status: ${res.getResponseCode()}`);
  });
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const contents = JSON.parse(e.postData.contents);
    const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
    const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
    if (contents.action === "fast_status_update") {
      const updates = contents.updates; 
      const lastRowMaster = masterSheet.getLastRow();
      
      const statusRange = masterSheet.getRange(1, 15, lastRowMaster, 1);
      let statusValues = statusRange.getValues();

      updates.forEach(update => {
        let rowIdx = parseInt(update[0]);
        let newStatus = update[1];
        if (rowIdx > 1 && rowIdx <= lastRowMaster) {
          statusValues[rowIdx - 1][0] = newStatus;
        }
      });

      statusRange.setValues(statusValues);
      
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify({"success": true, "msg": "Master updated"})).setMimeType(ContentService.MimeType.JSON);
    }

    const data = contents.data; 
    const masterData = masterSheet.getDataRange().getValues();
    const reportValues = reportSheet.getDataRange().getValues(); 

    data.forEach(t => {
      const teamName = t.team || "Nieznana_Ekipa";
      const now = new Date();
      const timeStr = Utilities.formatDate(now, "GMT+1", "HH:mm:ss");

      let masterIdx = parseInt(t.masterRow);
      if (!masterIdx || masterIdx < 2) {
        let mRow = masterData.find(r => String(r[2]) === String(t.loc) && String(r[7]) === String(t.sku));
        masterIdx = mRow ? masterData.indexOf(mRow) + 1 : null;
      }

      let mData = masterIdx ? masterData[masterIdx - 1] : null;
      let sysQty = mData ? mData[8] : (t.sysQty || 0);
      let batch = mData ? mData[10] : (t.batch || "");
      let sscc = mData ? mData[9] : "";
      let palletKind = mData ? mData[13] : "";

      let isDuplicate = reportValues.some(r => 
        String(r[2]) === String(t.loc) && 
        String(r[3]) === String(t.sku) && 
        String(r[1]) === teamName &&
        r[7] === "" 
      );

      if (isDuplicate && t.info !== "2. LICZENIE") return;

      if (t.info === "2. LICZENIE") {
        let found = false;
        for (let i = reportValues.length - 1; i >= 1; i--) {
          if (String(reportValues[i][2]) === String(t.loc) && String(reportValues[i][3]) === String(t.sku) && reportValues[i][7] === "") {
            const finalStatus = (String(t.qty) === String(sysQty)) ? "OK (2.L)" : "ROZBIEŻNOŚĆ FINALNA";
            reportSheet.getRange(i + 1, 8).setValue(t.qty); 
            reportSheet.getRange(i + 1, 9).setValue(finalStatus); 
            reportSheet.getRange(i + 1, 15).setValue(t.comment); 
            
            if (masterIdx) {
              masterSheet.getRange(masterIdx, 15).setValue("ZAKOŃCZONE FINALNIE (" + teamName + ")");
            }
            found = true;
            break;
          }
        }
        if (!found) {
          reportSheet.appendRow([now, teamName, t.loc, t.sku, batch, sysQty, "", t.qty, "ROZBIEŻNOŚĆ FINALNA", t.comment, timeStr, timeStr, sscc, palletKind, "", masterIdx]);
        }
      } else {
        const status = (String(t.qty) === String(sysQty)) ? "OK" : "ROZBIEŻNOŚĆ";
        reportSheet.appendRow([
          now, teamName, t.loc, t.sku, batch, sysQty, t.qty, "", status, 
          t.comment, timeStr, timeStr, sscc, palletKind, "", masterIdx
        ]);

        if (masterIdx) {
          masterSheet.getRange(masterIdx, 15).setValue(status + " (" + teamName + ")");
        }
      }
    });

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({"success": true})).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Błąd doPost: " + err.message);
    return ContentService.createTextOutput(JSON.stringify({"success": false, "error": err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
function emergencyForceSaveTeam() {
  const TEAM_TO_SAVE = "Ekipa_6";  
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("AWARYJNY ZAPIS", "Czy na pewno chcesz wymusić zapis dla: " + TEAM_TO_SAVE + "?", ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) return;

  const teamSheet = SS.getSheetByName(TEAM_TO_SAVE);
  const lastRow = teamSheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert("Błąd: Arkusz " + TEAM_TO_SAVE + " jest pusty.");
    return;
  }

  const currentTasks = teamSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  
  const allTasksMapped = currentTasks.map(row => {
    return {
      loc: String(row[0]),
      qty: String(row[2]), 
      comment: String(row[5])
    };
  });

  const startTime = "AWARYJNY";
  
  const result = saveAllTasks(TEAM_TO_SAVE, allTasksMapped, startTime);

  if (result.success) {
    ui.alert("SUKCES", "Dane z " + TEAM_TO_SAVE + " zostały pomyślnie przetransferowane do Raportu i Mastera.", ui.ButtonSet.OK);
  } else {
    ui.alert("BŁĄD", "Coś poszło nie tak: " + result.msg, ui.ButtonSet.OK);
  }
}
function doGet() {
  return HtmlService.createTemplateFromFile('Skaner')
      .evaluate()
      .setTitle('Skaner Inwentaryzacyjny')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function repairMasterHeaders() {
  const sheet = SS.getSheetByName(MASTER_SHEET_NAME);
  
  const headers = [[
    "Hala",             // A (1)
    "Alejka",           // B (2)
    "Location",         // C (3)
    "Level",            // D (4)
    "Cell",             // E (5)
    "Klient",           // F (6)
    "Status_Miejsca",   // G (7)
    "ItemCode",         // H (8)
    "LotQuantity",      // I (9)
    "SSCC",             // J (10)
    "BATCH",            // K (11)
    "NbParcelPerPallet",// L (12)
    "BestBeforeDate",   // M (13)
    "SupportTypeCaption",// N (14)
    "Status"            // O (15) 
  ]];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers[0]);
  } else {
    sheet.getRange(1, 1, 1, 15).setValues(headers);
  }

  const headerRange = sheet.getRange(1, 1, 1, 15);
  headerRange.setBackground("#1e293b")
             .setFontColor("white")
             .setFontWeight("bold")
             .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  
  sheet.autoResizeColumns(1, 15);

  SpreadsheetApp.getUi().alert("Nagłówki w Master_Stock zostały przywrócone!");
}
function getInitialData(pass) {
  try {
    const team = KODY[pass.trim()];
    if (!team) return null;
    
    const sheet = SS.getSheetByName(team);
    const lastRow = sheet.getLastRow();
    const startTime = Utilities.formatDate(new Date(), "GMT+1", "HH:mm:ss");

    if (lastRow < 2) return { team: team, tasks: [], startTime: startTime };

    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    
    return {
      team: team,
      startTime: startTime,
      tasks: data.map(r => ({
        loc: String(r[0]),
        sku: String(r[1]),
        sysQty: r[3],
        info: String(r[4]),
        batch: String(r[6]),
        isFull: (Number(r[7]) === Number(r[8]) && Number(r[8]) > 0),
        qty: "",
        comment: "",
        masterRow: r[9],
        palletKind: String(r[10])
      }))
    };
  } catch (e) { 
    console.error("Błąd getInitialData: " + e.message);
    return null; 
  }
}
function buildReportRow(date, team, loc, sku, batch, sys, fiz1, fiz2, status, comm1, start, end, mIdx, comm2, mSheet) {
  let sscc = "";
  let bbd = "";
  
  if (mIdx && mIdx > 1) {
    const extra = mSheet.getRange(mIdx, 10, 1, 3).getValues()[0];
    sscc = extra[0]; // Kolumna J (SSCC)
    bbd = extra[2];  // Kolumna L (BBD / Data ważności)
  }

  let diff1 = (fiz1 !== "" && !isNaN(fiz1)) ? (Number(fiz1) - Number(sys)) : 0;
  return [
    date,    // A: Data
    team,    // B: Ekipa
    loc,     // C: Lok.
    sku,     // D: SKU
    batch,   // E: Batch
    sys,     // F: Oczekiwana
    fiz1,    // G: 1. Liczenie
    fiz2,    // H: 2. Liczenie
    status,  // I: Status
    comm1,   // J: Komentarz
    start,   // K: Start
    end,     // L: Koniec
    sscc,    // M: SSCC (Poprawione)
    bbd,     // N: BBD (Poprawione)
    comm2,   // O: Komentarz2
    mIdx,    // P: MasterRowIdx
    diff1,   // Q: wartosc roznicy
    "",      // R: roznica 2 licz
    "",      // S: wartosc II liczenia
    ""       // T: III Liczenie
  ];
}
function fixStatusForSecondCount() {
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Stock");
  const data = masterSheet.getRange(2, 15, masterSheet.getLastRow() - 1, 1).getValues();
  const newData = data.map(row => {
    return [row[0].toString().includes("ROZBIEŻNOŚĆ") ? "DO 2. LICZENIA" : row[0]];
  });
  masterSheet.getRange(2, 15, newData.length, 1).setValues(newData);
}
function saveAllTasks(teamName, allTasks, startTime) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(50000); 
    const toNum = (val) => {
      if (val === "" || val === null || val === undefined) return 0;
      let n = parseFloat(String(val).replace(/\s+/g, '').replace(',', '.'));
      return isNaN(n) ? 0 : n;
    };

    const teamSheet = SS.getSheetByName(teamName);
    const lastRow = teamSheet.getLastRow();
    if (lastRow < 2) {
      if (lock.hasLock()) lock.releaseLock();
      return { success: true, msg: "Brak danych." };
    }

    const currentTasks = teamSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
    const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
    
    // POBIERAMY CAŁY RAPORT RAZ
    const reportRange = reportSheet.getDataRange();
    const reportData = reportRange.getValues();
    
    // --- OPTYMALIZACJA 1: MAPA RAPORTU ---
    // Lokalizacja występuje raz, więc robimy szybki indeks
    let reportMap = {};


    for (let i = reportData.length - 1; i >= 1; i--) {
      const loc = String(reportData[i][2]);
      const sku = String(reportData[i][3]);
      const key = JSON.stringify([loc, sku]);
    
    
      if (!reportMap[key]) {
        reportMap[key] = [];
      }
    
    
      reportMap[key].push(i);
    }

    const endTime = Utilities.formatDate(new Date(), "GMT+1", "HH:mm:ss");
    const masterLastRow = masterSheet.getLastRow();
    const masterStatusRange = masterSheet.getRange(1, 15, masterLastRow, 1);
    let masterStatusValues = masterStatusRange.getValues();
    
    let newReportEntries = [];
    let reportUpdates = [];
    //let reportWasModified = false;

    currentTasks.forEach(row => {
      const loc = String(row[0]);
      const sku = String(row[1]);
      const sys = toNum(row[3]);
      const type = row[4]; 
      const masterRowIdx = parseInt(row[9]); 
      const match = allTasks.find(t =>
        String(t.loc) === loc &&
        String(t.sku) === sku &&
        Number(t.masterRow) === masterRowIdx
      );
      const fiz = match ? toNum(match.qty) : 0;
      const comment = match ? match.comment : "";

      if (type === "2. LICZENIE") {


        const key = JSON.stringify([loc, sku]);
        const candidates = reportMap[key] || [];
      
      
        const rowIndex = candidates.find(i =>
          reportData[i][7] === "" ||
          reportData[i][7] === null
        );
      
      
        if (rowIndex !== undefined) {
      
      
          const finalStatus = (fiz === sys)
            ? "OK (2.L)"
            : "ROZBIEŻNOŚĆ FINALNA";
      
      
          reportData[rowIndex][7] = fiz;
          reportData[rowIndex][8] = finalStatus;
          reportData[rowIndex][14] = comment;
      
      
          reportUpdates.push({
            row: rowIndex + 1,
            qty: fiz,
            status: finalStatus,
            comment: comment
          });
      
      
          if (masterRowIdx && masterRowIdx <= masterLastRow) {
            masterStatusValues[masterRowIdx - 1][0] =
              "ZAKOŃCZONE FINALNIE (" + teamName + ")";
          }
      
      
        } else {
      
      
          newReportEntries.push(
            buildReportRow(
              new Date(),
              teamName,
              loc,
              sku,
              row[6],
              sys,
              "",
              fiz,
              "ROZBIEŻNOŚĆ FINALNA",
              "",
              startTime,
              endTime,
              masterRowIdx,
              comment,
              masterSheet
            )
          );
      
      
        }
      
      
      } else {
        const isOk = (fiz === sys);
        const status = isOk ? "OK" : "DO 2. LICZENIA"; 
        
        if (masterRowIdx && masterRowIdx <= masterLastRow) {
          masterStatusValues[masterRowIdx - 1][0] = status + (isOk ? " (" + teamName + ")" : "");
        }
        
        newReportEntries.push(buildReportRow(new Date(), teamName, loc, sku, row[6], sys, fiz, "", isOk ? "OK" : "ROZBIEŻNOŚĆ", comment, startTime, endTime, masterRowIdx, "", masterSheet));
      }
    });

    // --- ZAPIS MASOWY (TYLKO 3 STRZAŁY DO SERWERA) ---
    
    // 1. Aktualizacja Mastera (statusy)
    masterStatusRange.setValues(masterStatusValues);
    
    // 2. Aktualizacja istniejących wierszy w Raporcie (jeśli były zmiany w 2. liczeniu)
    

    reportUpdates.forEach(update => {
    
    
      reportSheet.getRange(update.row, 8, 1, 2)
        .setValues([[
          update.qty,
          update.status
        ]]);
    
    
      reportSheet.getRange(update.row, 15)
        .setValue(update.comment);
    
    
    });

    // 3. Dodanie nowych wierszy do Raportu
    if (newReportEntries.length > 0) {
      reportSheet.getRange(reportSheet.getLastRow() + 1, 1, newReportEntries.length, 20).setValues(newReportEntries);
    }

    // Czyszczenie i nowa paczka
    teamSheet.getRange("A2:K").clearContent(); 
    SpreadsheetApp.flush(); // Wymuś zapis przed sync
    
    syncAndGetNewData(teamName);
    const nextData = getInitialDataForLoop(teamName); 

    if (lock.hasLock()) lock.releaseLock();
    return { success: true, hasNewTasks: nextData.tasks.length > 0, newData: nextData };

  } catch (e) {
    if (lock && lock.hasLock()) lock.releaseLock();
    return { success: false, msg: e.message };
  }
}
function ssaveAllTasks(teamName, allTasks, startTime) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(50000); 
    const toNum = (val) => {
      if (val === "" || val === null || val === undefined) return 0;
      let n = parseFloat(String(val).replace(/\s+/g, '').replace(',', '.'));
      return isNaN(n) ? 0 : n;
    };
    const teamSheet = SS.getSheetByName(teamName);
    const lastRow = teamSheet.getLastRow();
    if (lastRow < 2) {
      if (lock.hasLock()) lock.releaseLock();
      return { success: true, msg: "Brak danych." };
    }
    const currentTasks = teamSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
    const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
    const reportRange = reportSheet.getDataRange();
    const reportData = reportRange.getValues();
    const endTime = Utilities.formatDate(new Date(), "GMT+1", "HH:mm:ss");
    const masterLastRow = masterSheet.getLastRow();
    const masterStatusRange = masterSheet.getRange(1, 15, masterLastRow, 1);
    let masterStatusValues = masterStatusRange.getValues();
    let newReportEntries = [];
    currentTasks.forEach(row => {
      const loc = String(row[0]);
      const sku = String(row[1]);
      const sys = toNum(row[3]);
      const type = row[4]; 
      const masterRowIdx = parseInt(row[9]); 
      const match = allTasks.find(t => String(t.loc) === loc);
      const fiz = match ? toNum(match.qty) : 0;
      const comment = match ? match.comment : "";
      if (type === "2. LICZENIE") {
        let foundInReport = false;
        for (let i = reportData.length - 1; i >= 1; i--) {
          if (String(reportData[i][2]) === loc && String(reportData[i][3]) === sku && String(reportData[i][7]) === "") {
            const finalStatus = (fiz === sys) ? "OK (2.L)" : "ROZBIEŻNOŚĆ FINALNA";       
            reportSheet.getRange(i + 1, 8).setValue(fiz);          // 2. Liczenie
            reportSheet.getRange(i + 1, 9).setValue(finalStatus);  // Status
            reportSheet.getRange(i + 1, 15).setValue(comment);     // Komentarz2
            if (masterRowIdx && masterRowIdx <= masterLastRow) {
              masterStatusValues[masterRowIdx - 1][0] = "ZAKOŃCZONE FINALNIE (" + teamName + ")";
            }
            foundInReport = true;
            break;
          }
        }
        if (!foundInReport) {
          newReportEntries.push(buildReportRow(now, teamName, loc, sku, row[6], sys, "", fiz, "ROZBIEŻNOŚĆ FINALNA", "", startTime, endTime, masterRowIdx, comment, masterSheet));
        }
      } else {
        const isOk = (fiz === sys);
        const status = isOk ? "OK" : "DO 2. LICZENIA"; 
        
        if (masterRowIdx && masterRowIdx <= masterLastRow) {
          masterStatusValues[masterRowIdx - 1][0] = status + (isOk ? " (" + teamName + ")" : "");
        }
        
        newReportEntries.push(buildReportRow(new Date(), teamName, loc, sku, row[6], sys, fiz, "", isOk ? "OK" : "ROZBIEŻNOŚĆ", comment, startTime, endTime, masterRowIdx, "", masterSheet));
      }
    });

    masterStatusRange.setValues(masterStatusValues);
    if (newReportEntries.length > 0) {
      reportSheet.getRange(reportSheet.getLastRow() + 1, 1, newReportEntries.length, 20).setValues(newReportEntries);
    }

    teamSheet.getRange("A2:K").clearContent(); 
    syncAndGetNewData(teamName);
    const nextData = getInitialDataForLoop(teamName); 

    lock.releaseLock();
    return { success: true, hasNewTasks: nextData.tasks.length > 0, newData: nextData };

  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, msg: e.message };
  }
}
function getInitialDataForLoop(teamName) {
  const sheet = SS.getSheetByName(teamName);
  const lastRow = sheet.getLastRow();
  const startTime = Utilities.formatDate(new Date(), "GMT+1", "HH:mm:ss");

  if (lastRow < 2) return { team: teamName, tasks: [], startTime: startTime };

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  return {
    team: teamName,
    startTime: startTime,
    tasks: data.map(r => ({
      loc: String(r[0]),
      sku: String(r[1]),
      sysQty: r[3],
      info: String(r[4]),
      batch: String(r[6]),
      isFull: (Number(r[7]) === Number(r[8]) && Number(r[8]) > 0),
      qty: "",
      comment: "",
      masterRow: r[9],
      palletKind: String(r[10])
    }))
  };
}
function activateSecondCountTrigger() {
  const ui = SpreadsheetApp.getUi();
  const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const reportData = reportSheet.getDataRange().getValues();
  
  let counter = 0;
  for (let i = 1; i < reportData.length; i++) {
    if (reportData[i][8] === "ROZBIEŻNOŚĆ" && reportData[i][7] === "") {
      let masterRow = reportData[i][15]; // Kolumna P
      if (masterRow && !isNaN(masterRow)) {
        masterSheet.getRange(masterRow, 15).setValue("DO 2. LICZENIA");
        counter++;
      }
    }
  }
  SpreadsheetApp.flush();
  for (let j = 1; j <= TEAMS_COUNT; j++) {
    syncAndGetNewData("Ekipa_" + j);
  }
  ui.alert("Aktywowano 2. liczenie dla " + counter + " pozycji.");
}
function repairReportMasterIndices() {
  const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const reportData = reportSheet.getDataRange().getValues();
  const masterData = masterSheet.getDataRange().getValues();
  let masterMap = {};
  for (let i = 1; i < masterData.length; i++) {
    let key = String(masterData[i][2]) + "|" + String(masterData[i][7]); // C | H
    masterMap[key] = i + 1;
  }
  let indices = [];
  for (let j = 1; j < reportData.length; j++) {
    let reportKey = String(reportData[j][2]) + "|" + String(reportData[j][3]); // C | D
    indices.push([masterMap[reportKey] || ""]);
  }
  reportSheet.getRange(2, 16, indices.length, 1).setValues(indices);
  reportSheet.getRange(1, 16).setValue("MasterRowIdx").setBackground("#fbbf24");
  
  SpreadsheetApp.getUi().alert("Naprawiono! Numery wierszy zostały dopisane w kolumnie P.");
}
function wyzwijRozbieznosci() {
  const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
  const data = reportSheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    let fiz1 = data[i][6]; // Kolumna G
    let sys = data[i][5];  // Kolumna F
    let status = data[i][8]; // Kolumna I
    if (fiz1 !== "" && status === "") {
      if (parseFloat(fiz1) !== parseFloat(sys)) {
        reportSheet.getRange(i + 1, 9).setValue("ROZBIEŻNOŚĆ");
      } else {
        reportSheet.getRange(i + 1, 9).setValue("OK");
      }
    }
  }
}
function syncAndGetNewData(manualTeamName) {
  const sheet = manualTeamName ? SS.getSheetByName(manualTeamName) : SS.getActiveSheet();
  const teamName = sheet.getName();
  if (!teamName.startsWith("Ekipa_")) return;

  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const teamNum = parseInt(teamName.split("_")[1]);
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(50000); 
    const lastRowMaster = masterSheet.getLastRow();
    if (lastRowMaster < 2) { if(lock.hasLock()) lock.releaseLock(); return; }

    const masterRange = masterSheet.getRange(2, 1, lastRowMaster - 1, 15);
    const masterData = masterRange.getValues();
    const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
    const reportData = reportSheet.getDataRange().getValues();
    const reportIndex = new Map();
    for (let i = 1; i < reportData.length; i++) {
      const r = reportData[i];
      const key = `${r[2]}|${r[3]}|${r[7] === ""}`;
      reportIndex.set(key, r[1]); 
    }
    const isAntresolaTeam = (teamNum <= SPEC_TEAMS);
    const antresolaLevels = [10, 11, 12, 13];

    let reservedForMe = [];
    let busyAlleysInMyZone = new Set();

    // 1. BLOKADY KORYTARZY
    for (let i = 0; i < masterData.length; i++) {
      const status = String(masterData[i][14]);
      const currentAlley = parseInt(masterData[i][1]);
      if (status === "ZAREZERWOWANE (" + teamName + ")") {
        reservedForMe.push({ data: masterData[i], rowNum: i + 2 });
      } else if (status.startsWith("ZAREZERWOWANE")) {
        let resTeamMatch = status.match(/\d+/);
        let resTeamNum = resTeamMatch ? parseInt(resTeamMatch[0]) : 0;
        let resIsAntresola = (resTeamNum <= SPEC_TEAMS);
        
        if (resIsAntresola === isAntresolaTeam) {
          let pair = (currentAlley % 2 === 0) ? [currentAlley, currentAlley + 1] : [currentAlley - 1, currentAlley];
          busyAlleysInMyZone.add(pair[0]);
          busyAlleysInMyZone.add(pair[1]);
        }
      }
    }

    let tasksToLoad = [];
    if (reservedForMe.length > 0) {
      tasksToLoad = reservedForMe;
    } else {
      
      // --- KROK 1 DLA ANTRESOLI (STARE ZASADY: NAJPIERW 2. LICZENIE) ---
      if (isAntresolaTeam) {
        let discrepancies = masterData.map((r, i) => ({ 
          data: r, rowNum: i + 2, lvl: parseInt(r[3]), alley: parseInt(r[1]), status: String(r[14]) 
        }))
        .filter(t => {
          if (antresolaLevels.includes(t.lvl) && t.status === "DO 2. LICZENIA") {
            const prevTeam = reportIndex.get(`${t.data[2]}|${t.data[7]}|true`);
            return prevTeam !== undefined && prevTeam !== teamName; 
          }
          return false;
        });
        discrepancies.slice(0, TASK_LIMIT).forEach(t => {
          masterData[t.rowNum - 2][14] = "ZAREZERWOWANE (" + teamName + ")";
          tasksToLoad.push({ data: masterData[t.rowNum - 2], rowNum: t.rowNum });
        });
      }

      // --- KROK 2 (DLA ZWYŻEK PRIORYTET / DLA ANTRESOLI DOPEŁNIENIE): 1. LICZENIE ---
      // --- KROK B: NOWE ZADANIA (POPRAWIONA LOGIKA CIĄGŁOŚCI) ---
      if (tasksToLoad.length < TASK_LIMIT) {
        let possibleTasks = masterData.map((r, i) => ({ 
            data: r, index: i, alley: parseInt(r[1]), lvl: parseInt(r[3]), status: String(r[14]) 
          }))
          .filter(t => {
            let isLevelValid = isAntresolaTeam ? antresolaLevels.includes(t.lvl) : (t.lvl >= 20 && t.lvl <= 50); 
            return t.status === "" && isLevelValid; 
          });

        if (possibleTasks.length > 0) {
          let myCurrentAlley = (reservedForMe.length > 0) ? reservedForMe[0].data[1] : null;
          
          let targetAlley;
          if (myCurrentAlley && possibleTasks.some(t => t.alley == myCurrentAlley && !busyAlleysInMyZone.has(t.alley))) {
            targetAlley = parseInt(myCurrentAlley);
          } else {
            // Jeśli Twój korytarz jest pusty, dopiero wtedy bierzemy najniższą wolną alejkę
            let trulyFreeTasks = possibleTasks.filter(t => !busyAlleysInMyZone.has(t.alley));
            if (trulyFreeTasks.length === 0) { /* Brak zadań */ }
            else {
              trulyFreeTasks.sort((a, b) => a.alley - b.alley);
              targetAlley = trulyFreeTasks[0].alley;
            }
          }

          if (targetAlley) {
            let alleyPair = (targetAlley % 2 === 0) ? [targetAlley, targetAlley + 1] : [targetAlley - 1, targetAlley];
            
            // Pobieramy zadania z tej pary, które nie są zablokowane przez innych
            let corridorTasks = possibleTasks.filter(t => alleyPair.includes(t.alley) && !busyAlleysInMyZone.has(t.alley));
            
            corridorTasks.sort((a, b) => {
              if (b.lvl !== a.lvl) return b.lvl - a.lvl; // Najpierw góra
              return a.alley - b.alley; // Potem strona
            });

            for (let t of corridorTasks) {
              if (tasksToLoad.length >= TASK_LIMIT) break;
              masterData[t.index][14] = "ZAREZERWOWANE (" + teamName + ")";
              tasksToLoad.push({ data: masterData[t.index], rowNum: t.index + 2 }); 
            }
          }
        }
      }

      // --- KROK 3 DLA ZWYŻEK (2. LICZENIE TYLKO GDY KORYTARZ WOLNY I BRAK NOWYCH ZADAŃ) ---
      if (!isAntresolaTeam && tasksToLoad.length < TASK_LIMIT) {
        let discrepanciesVNA = masterData.map((r, i) => ({ 
          data: r, 
          rowNum: i + 2, 
          lvl: parseInt(r[3]), 
          alley: parseInt(r[1]), 
          status: String(r[14]) 
        }))
        .filter(t => {
          let isLevelMatch = (t.lvl >= 20 && t.lvl <= 50); 
          if (busyAlleysInMyZone.has(t.alley)) return false; 

          if (t.status === "DO 2. LICZENIA" && isLevelMatch) {
            // BŁYSKAWICZNE SPRAWDZENIE: pobieramy ekipę z mapy zamiast przeszukiwać arkusz
            const prevTeam = reportIndex.get(`${t.data[2]}|${t.data[7]}|true`); 
            
            
            return prevTeam !== undefined && prevTeam !== teamName; 
          }
          return false;
        });

        discrepanciesVNA.slice(0, TASK_LIMIT - tasksToLoad.length).forEach(t => {
          masterData[t.rowNum - 2][14] = "ZAREZERWOWANE (" + teamName + ")"; 
          tasksToLoad.push({ data: masterData[t.rowNum - 2], rowNum: t.rowNum }); 
        });
      }

      masterRange.setValues(masterData);
      SpreadsheetApp.flush(); 
    }

    if (lock.hasLock()) lock.releaseLock();
    if (tasksToLoad.length === 0) return;

    // --- SORTOWANIE TRASY ---
    let sorted = tasksToLoad.map(obj => obj.data);
    sorted.sort((a, b) => {
      let alleyA = parseInt(a[1]) || 0;
      let alleyB = parseInt(b[1]) || 0;
      let lvlA = parseInt(a[3]) || 0;
      let lvlB = parseInt(b[3]) || 0;
      if (lvlA !== lvlB) return lvlB - lvlA;
      if (alleyA !== alleyB) return alleyA - alleyB;
      let locA = parseInt(a[2].split('.')[1]) || 0;
      let locB = parseInt(b[2].split('.')[1]) || 0;
      return (alleyA % 2 === 0) ? locA - locB : locB - locA;
    });

    const taskMap = new Map();
    tasksToLoad.forEach(t => taskMap.set(`${t.data[2]}|${t.data[7]}`, t.rowNum));

    let formatted = sorted.map(r => {
      const key = `${r[2]}|${r[7]}`;
      const isSecondCount = reportIndex.has(`${key}|true`);
      const originalRow = taskMap.get(key);

      return [
        r[2], r[7], "", r[8], 
        isSecondCount ? "2. LICZENIE" : "1. LICZENIE", 
        "", r[10], r[8], r[11], 
        originalRow || "", r[13] || "BRAK"
      ];
    }); 

    sheet.getRange("A2:K").clearContent(); 
    if (formatted.length > 0) {
      sheet.getRange(2, 1, formatted.length, 11).setValues(formatted); 
      applyOfflineColorFormatting(sheet, formatted.length + 1); 
    }

  } catch (e) {
    if (lock && lock.hasLock()) lock.releaseLock();
    console.error("Błąd sync: " + e.message); 
  }
}

function generateMissingLevelsAndDiscrepancies() {
  const ui = SpreadsheetApp.getUi();
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const masterData = masterSheet.getDataRange().getValues();
  const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
  const reportData = reportSheet.getDataRange().getValues();
  
  const GROUND_TEAM = "Ekipa_1"; 
  const HIGH_TEAM = "Ekipa_8";   
  
  let groundTasks = []; 
  let highTasks = [];   

  const reportIndex = new Map();
  for (let i = 1; i < reportData.length; i++) {
    reportIndex.set(`${reportData[i][2]}|${reportData[i][3]}|${reportData[i][7] === ""}`, reportData[i][1]);
  }

  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const lvl = parseInt(row[3]); 
    const status = String(row[14]); 
    const location = row[2];
    const sku = row[7];

    let shouldAssignToGround = false;
    let shouldAssignToHigh = false;
    let taskType = "1. LICZENIE";

    if (lvl === 14 && (status === "" || status === "DO 2. LICZENIA")) {
      shouldAssignToGround = true;
      if (status === "DO 2. LICZENIA") taskType = "2. LICZENIE";
    } 
    else if (lvl === 60 && (status === "" || status === "DO 2. LICZENIA")) {
      shouldAssignToHigh = true;
      if (status === "DO 2. LICZENIA") taskType = "2. LICZENIE";
    }
    else if (status === "DO 2. LICZENIA") {
      shouldAssignToHigh = true;
      taskType = "2. LICZENIE";
    }

    if (shouldAssignToGround || shouldAssignToHigh) {
      const formattedRow = [
        location, sku, "", row[8], taskType, "", row[10], row[8], row[11], i + 1, row[13] || "BRAK"
      ];

      if (shouldAssignToGround) {
        groundTasks.push(formattedRow);
        masterData[i][14] = "ZAREZERWOWANE (" + GROUND_TEAM + ")";
      } else {
        highTasks.push(formattedRow);
        masterData[i][14] = "ZAREZERWOWANE (" + HIGH_TEAM + ")";
      }
    }
  }

  masterSheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);

  if (groundTasks.length > 0) {
    const s1 = SS.getSheetByName(GROUND_TEAM);
    s1.getRange("A2:K").clearContent();
    s1.getRange(2, 1, groundTasks.length, 11).setValues(groundTasks);
    applyOfflineColorFormatting(s1, groundTasks.length + 1);
  }

  if (highTasks.length > 0) {
    const s6 = SS.getSheetByName(HIGH_TEAM);
    s6.getRange("A2:K").clearContent();
    s6.getRange(2, 1, highTasks.length, 11).setValues(highTasks);
    applyOfflineColorFormatting(s6, highTasks.length + 1);
  }

  SpreadsheetApp.flush();
  ui.alert(`Koniec inwentaryzacji - podsumowanie przydziału:
  - ${GROUND_TEAM} (Lvl 14): ${groundTasks.length} zadań.
  - ${HIGH_TEAM} (Lvl 60 + reszta rozbieżności): ${highTasks.length} zadań.`);
}
function distributeDualAisleData() {
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const masterData = masterSheet.getDataRange().getValues();
  let available = [];
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][14] === "" || masterData[i][14] === "DO 1. LICZENIE") {
      available.push({ data: masterData[i], rowNum: i + 1 });
    }
  }

  let alleys = [...new Set(available.map(r => String(r.data[1])))].sort((a,b) => a-b);
  let currentTeam = 5; 
  for (let i = 0; i < alleys.length; i++) {
    if (currentTeam > TEAMS_COUNT) break;
    let alleyA = alleys[i], alleyB = alleys[i+1];
    if (alleyB && (parseInt(alleyB) === parseInt(alleyA) + 1)) {
      assignSpecificAlley("Ekipa_" + currentTeam, alleyA, available, masterSheet);
      if (currentTeam + 1 <= TEAMS_COUNT) assignSpecificAlley("Ekipa_" + (currentTeam + 1), alleyB, available, masterSheet);
      currentTeam += 2; i++; 
    }
  }
  SpreadsheetApp.getUi().alert("Rozdano alejki w trybie DWUSTRONNYM.");
}

function assignSpecificAlley(teamName, alleyLabel, allRows, masterSheet) {
  let teamTasks = allRows.filter(r => String(r.data[1]) === alleyLabel);
  if (teamTasks.length === 0) return; 
  let sheet = SS.getSheetByName(teamName);
  let sortedData = sortSyncDirection(teamTasks.map(obj => obj.data));
  let formatted = sortedData.map(r => {
    let original = teamTasks.find(t => t.data[2] === r[2] && t.data[7] === r[7]);
    return [
      r[2], r[7], "", r[8], "1. LICZENIE", "", r[10], r[8], r[11], 
      original.rowNum, 
      r[13] || "BRAK" 
    ];
  });
  sheet.getRange("A2:K").clearContent(); 
  sheet.getRange(2, 1, formatted.length, 11).setValues(formatted);
  teamTasks.forEach(item => masterSheet.getRange(item.rowNum, 15).setValue("ZAREZERWOWANE (" + teamName + ")"));
  applyOfflineColorFormatting(sheet, formatted.length + 1);
}
function sortSerpentine(chunk) {
  if (chunk.length === 0) return [];

  let mapped = chunk.map(r => {
    let details = parseLocationDetails(r[2]);
    let alleyPairKey = Math.floor(details.alley / 2); 
    return { data: r, details: details, pairKey: alleyPairKey };
  });

  mapped.sort((a, b) => {
    if (a.pairKey !== b.pairKey) return a.pairKey - b.pairKey;

    if (a.details.lvl !== b.details.lvl) return b.details.lvl - a.details.lvl;

    if (a.details.alley !== b.details.alley) return a.details.alley - b.details.alley;

    if (a.details.alley % 2 === 0) {
      if (a.details.bay !== b.details.bay) return a.details.bay - b.details.bay;
      return a.details.pos - b.details.pos;
    } else {
      if (a.details.bay !== b.details.bay) return b.details.bay - a.details.bay;
      return b.details.pos - a.details.pos;
    }
  });

  return mapped.map(m => m.data);
}
function parseLocationDetails(locStr) {
  let parts = String(locStr).split('.');
  let alley = parseInt(parts[0]);
  let sub = parts[1] || "0000000";
  return {
    alley: alley,
    bay: parseInt(sub.substring(0, 3)),   // 021
    lvl: parseInt(sub.substring(3, 5)),   // 50
    pos: parseInt(sub.substring(5))       // 01
  };
}
function sortSyncDirection(chunk) {
  let levels = {};
  chunk.forEach(r => {
    let parts = String(r[2]).split('.');
    let l = (parts.length > 1 && parts[1].length >= 4) ? parts[1].substring(3, 4) : "0";
    if (!levels[l]) levels[l] = [];
    levels[l].push(r);
  });
  let sortedLvl = Object.keys(levels).sort((a,b) => b-a);
  let res = [];
  sortedLvl.forEach(l => {
    let data = levels[l];
    data.sort((a,b) => {
      let pA = String(a[2]).split('.')[1] || "0000000", pB = String(b[2]).split('.')[1] || "0000000";
      let bayA = parseInt(pA.substring(0, 3)), bayB = parseInt(pB.substring(0, 3));
      if (bayA !== bayB) return bayA - bayB;
      return parseInt(pA.substring(4)) - parseInt(pB.substring(4));
    });
    res.push(...data);
  });
  return res;
}
function updateDashboardSheet() {
  const { heatmap, stats } = getAdminHeatmapData();
  let sheet = SS.getSheetByName("DASHBOARD") || SS.insertSheet("DASHBOARD");
  sheet.clear(); 
  const teamAssignments = getTeamAssignments();
  const allAlleys = Object.keys(heatmap).sort((a, b) => parseInt(a) - parseInt(b));
  let specRows = [];
  let stdRows = [];
  allAlleys.forEach(alley => {
    const assignedTeam = teamAssignments[alley] || "---";
    const h = heatmap[alley];
    const s = stats[alley];
    const row = [
      alley, assignedTeam, 
      h.total > 0 ? (h.done / h.total) : 0,
      s.count1 > 0 ? (s.err1 / s.count1) : 0,
      s.count2 > 0 ? (s.err2 / s.count2) : 0,
      h.done + " / " + h.total
    ];
    if (h.isAntresola) {
      specRows.push(row);
    } else {
      stdRows.push(row);
    }
  });
  sheet.getRange(1, 1).setValue("EKIPY SPECJALNE (ANTRESOLA Lvl 10,11,12)").setFontWeight("bold").setBackground("#dcfce7");
  sheet.getRange(2, 1, 1, 6).setValues([["ALEJKA", "EKIPA", "POSTĘP", "BŁĄD 1.L", "BŁĄD 2.L", "STAN"]]).setBackground("#475569").setFontColor("white");
  if(specRows.length > 0) sheet.getRange(3, 1, specRows.length, 6).setValues(specRows);
  let startStd = specRows.length + 5;
  sheet.getRange(startStd, 1).setValue("EKIPY STANDARDOWE (PIĘTRA < 10, bez 20)").setFontWeight("bold").setBackground("#dbeafe");
  sheet.getRange(startStd + 1, 1, 1, 6).setValues([["ALEJKA", "EKIPA", "POSTĘP", "BŁĄD 1.L", "BŁĄD 2.L", "STAN"]]).setBackground("#475569").setFontColor("white");
  if(stdRows.length > 0) sheet.getRange(startStd + 2, 1, stdRows.length, 6).setValues(stdRows);
  sheet.getRange(1, 3, sheet.getLastRow(), 3).setNumberFormat("0%");
  sheet.autoResizeColumns(1, 6);
}
function forceAssignTasksFromDashboard() {
  const sheet = SS.getSheetByName("DASHBOARD");
  const lastRow = sheet.getLastRow();
  const data = sheet.getDataRange().getValues();
  let rIdx = -1;
  for(let i=0; i<data.length; i++){
    if(data[i][0] === "WPISZ NR EKIPY") { rIdx = i + 1; break; }
  }
  if(rIdx === -1) return;
  const teamNum = sheet.getRange(rIdx + 1, 1).getValue();
  const alleyNum = sheet.getRange(rIdx + 1, 2).getValue();
  if (!teamNum || !alleyNum) {
    SpreadsheetApp.getUi().alert("Błąd: Wpisz numer ekipy i numer alejki!");
    return;
  }
  const teamName = "Ekipa_" + teamNum;
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const masterData = masterSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < masterData.length; i++) {
    if (parseInt(masterData[i][1]) === parseInt(alleyNum)) {
      masterSheet.getRange(i + 1, 15).setValue("ZAREZERWOWANE (" + teamName + ")");
      count++;
    }
  }
  if (count > 0) {
    SpreadsheetApp.flush();
    syncAndGetNewData(teamName); 
    SpreadsheetApp.getUi().alert("Sukces! Przypisano " + count + " wierszy alejki " + alleyNum + " do " + teamName);
    updateDashboardSheet();
  } else {
    SpreadsheetApp.getUi().alert("Nie znaleziono wierszy dla alejki " + alleyNum);
  }
}
function getAdminHeatmapData() {
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
  const masterData = masterSheet.getDataRange().getValues();
  const reportData = reportSheet.getDataRange().getValues(); 
  const heatmap = {};
  const stats = {}; 
  const antresolaLevels = [10, 11, 12];
  for (let i = 1; i < masterData.length; i++) {
    const alley = String(masterData[i][1]).trim();
    const lvl = parseInt(masterData[i][3]);
    if (!alley) continue;
    const isAntresola = antresolaLevels.includes(lvl);
    const isStandard = (lvl < 10 && lvl !== 20); 
    if (!isAntresola && !isStandard) continue; 
    if (!heatmap[alley]) {
      heatmap[alley] = { total: 0, done: 0, isAntresola: isAntresola };
      stats[alley] = { err1: 0, count1: 0, err2: 0, count2: 0 };
    }
    heatmap[alley].total++;
    const statusM = String(masterData[i][14]);
    if (statusM !== "" && statusM !== "DO 1. LICZENIE" && !statusM.startsWith("ZAREZERWOWANE")) {
      heatmap[alley].done++;
    }
  }
  for (let j = 1; j < reportData.length; j++) {
    const row = reportData[j];
    const alley = String(row[2]).split('.')[0];
    if (stats[alley]) {
      if (row[6] !== "") { 
        stats[alley].count1++;
        if (row[8] === "ROZBIEŻNOŚĆ") stats[alley].err1++;
      } else if (row[7] !== "") { 
        stats[alley].count2++;
        if (row[8] === "ROZBIEŻNOŚĆ") stats[alley].err2++;
      }
    }
  }
  return { heatmap, stats };
}
function getTeamAssignments() {
  let teamAssignments = {};
  const masterData = SS.getSheetByName(MASTER_SHEET_NAME).getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    let alley = String(masterData[i][1]), status = String(masterData[i][14]);
    if (status.includes("(") && status.includes(")")) {
      let team = status.split('(')[1].split(')')[0];
      if (!teamAssignments[alley]) teamAssignments[alley] = team;
    }
  }
  return teamAssignments;
}
function setupInventorySystem() {
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  masterSheet.getRange(1, 15).setValue("Status").setBackground("#2c3e50").setFontColor("white").setFontWeight("bold");
  masterSheet.getRange(1, 14).setValue("SupportTypeCaption").setBackground("#2c3e50").setFontColor("white").setFontWeight("bold");
  for (let i = 1; i <= TEAMS_COUNT; i++) {
    let name = "Ekipa_" + i, s = SS.getSheetByName(name) || SS.insertSheet(name);
    s.clear(); 
    s.getRange(1, 1, 1, 11).setValues([["LOKALIZACJA", "SKU", "ILOŚĆ FIZ.", "ILOŚĆ SYS.", "INFO", "KOMENTARZ", "BATCH", "LOT_QTY", "PALLET_QTY", "MASTER_ROW", "PALLET_KIND"]]).setBackground("#2c3e50").setFontColor("white").setFontWeight("bold");
    s.setFrozenRows(1); 
    s.hideColumns(4); 
    s.hideColumns(8, 4); 
  }
  SpreadsheetApp.getUi().alert("System gotowy.");
}
function distributeInitialData() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock(); 
  try {
    lock.waitLock(60000);  
    for (let i = 1; i <= TEAMS_COUNT; i++) {
      const teamName = "Ekipa_" + i;
      console.log("Przetwarzanie: " + teamName);
      syncAndGetNewData(teamName);
      SpreadsheetApp.flush(); 
      Utilities.sleep(500); 
    }
    lock.releaseLock();
    ui.alert("Sukces! Rozdzielono pracę na wszystkie ekipy (1-13).");
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    ui.alert("Błąd krytyczny dystrybucji: " + e.message);
  }
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 SYSTEM')
      .addItem('1. Przygotuj System', 'setupInventorySystem')
      .addItem('🚀 Generuj zadania', 'fetchBigQueryDataAndDistribute')
      //.addItem('🚀 2. POBIERZ I ROZDAJ (Full)', 'distributeInitialData')
      .addItem('3. Rozdaj Pracę (DWIE STRONY)', 'distributeDualAisleData')
      .addItem('🛠️ WYMUŚ ALEJKĘ DLA EKIPY', 'autoAssignNextAlleysUI')
      .addItem('🛠️ 2 Liczenie', 'wyzwijRozbieznosci')
      .addSeparator()
      .addItem('ODSWIEZ DASHBOARD','updateDashboardSheet')
      .addItem('Reset Mastera', 'resetMasterStatus')
      .addItem('Wyczyść Arkusze Ekip','forceResetAllTeamSheets')
      .addToUi();
}

function autoAssignNextAlleysUI() {
  const ui = SpreadsheetApp.getUi();
  const teamResponse = ui.prompt('AUTOMATYCZNE DOŁADOWANIE', 'Dla której ekipy pobrać alejki? (np. 5):', ui.ButtonSet.OK_CANCEL);
  if (teamResponse.getSelectedButton() !== ui.Button.OK) return;
  const teamNum = teamResponse.getResponseText().trim();
  const countResponse = ui.prompt('AUTOMATYCZNE DOŁADOWANIE', 'Ile wolnych alejek przypisać? (domyślnie 2):', ui.ButtonSet.OK_CANCEL);
  if (countResponse.getSelectedButton() !== ui.Button.OK) return;
  let numToAssign = parseInt(countResponse.getResponseText().trim()) || 2;
  const teamName = "Ekipa_" + teamNum;
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const lastRow = masterSheet.getLastRow();
  const masterRange = masterSheet.getRange(2, 1, lastRow - 1, 15);
  const masterData = masterRange.getValues();
  let reservedAlleys = new Set();
  let allAvailableAlleys = new Set();
  for (let i = 0; i < masterData.length; i++) {
    let alley = String(masterData[i][1]);
    let status = String(masterData[i][14]);
    if (status !== "") {
      reservedAlleys.add(alley);
    } else {
      allAvailableAlleys.add(alley);
    }
  }
  let trulyFreeAlleys = Array.from(allAvailableAlleys)
    .filter(a => !reservedAlleys.has(a))
    .sort((a, b) => parseInt(a) - parseInt(b));
  if (trulyFreeAlleys.length === 0) {
    ui.alert("Brak wolnych alejek w Masterze!");
    return;
  }
  let selectedAlleys = trulyFreeAlleys.slice(0, numToAssign);
  let assignedCount = 0;
  for (let i = 0; i < masterData.length; i++) {
    if (selectedAlleys.includes(String(masterData[i][1]))) {
      masterData[i][14] = "ZAREZERWOWANE (" + teamName + ")";
      assignedCount++;
    }
  }
  masterRange.setValues(masterData);
  SpreadsheetApp.flush();
  syncAndGetNewData(teamName); 
  ui.alert("Sukces! Przypisano alejki: " + selectedAlleys.join(", ") + " do " + teamName);
  updateDashboardSheet();
}
function applyOfflineColorFormatting(sheet, lastRow) {
  sheet.clearConditionalFormatRules();
  let range = sheet.getRange(2, 1, lastRow, 11);
  let ruleDone = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=NOT(ISBLANK($C2))').setBackground("#000000").setFontColor("#444444").setRanges([range]).build();
  let ruleCurrent = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISBLANK($C2); OR(ROW()=2; NOT(ISBLANK($C1))))').setBackground("#ffffff").setFontColor("#000000").setBold(true).setRanges([range]).build();
  sheet.setConditionalFormatRules([ruleDone, ruleCurrent]);
}
function resetMasterStatus() {
  const s = SS.getSheetByName(MASTER_SHEET_NAME);
  s.getRange(2, 15, s.getLastRow(), 1).clearContent();
}

function forceResetAllTeamSheets() {
  for (let i = 1; i <= TEAMS_COUNT; i++) {
    let s = SS.getSheetByName("Ekipa_" + i);
    if (s) s.getRange("A2:K").clearContent().setBackground(null);
  }
}
function fetchBigQueryDataAndDistribute() {
  const projectId = 'fml-prod-rep-pl-rfx-sbe';
  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  const sql = `
    WITH masterdata AS (
      SELECT * FROM (
        SELECT VLCART AS Article, VLCVLA, VLQVLB 
        FROM \`fml-prod-rep-pl-rfx-sbe.last_rows_sbe.hlarvlp\` 
        WHERE VLCACT IN ('MTE')
      )
      PIVOT (MAX(VLQVLB) FOR VLCVLA IN ('10' as spcb, '20' as pcb, '30' as base_lv_on_pallet))
    ),
    daty AS (
      SELECT
        hl_gei.GENSUP, hl_gei.GECART, hl_gei.GENGEI,
        CASE WHEN hl_gei.GESDLU * hl_gei.GEADLU * hl_gei.GEMDLU * hl_gei.GEJDLU > 0 THEN
          CAST(FORMAT_DATE('%Y/%m/%d', PARSE_DATE('%Y%m%d', 
            LPAD(CAST(hl_gei.GESDLU AS STRING),2,'0') || LPAD(CAST(hl_gei.GEADLU AS STRING),2,'0') ||
            LPAD(CAST(hl_gei.GEMDLU AS STRING),2,'0') || LPAD(CAST(hl_gei.GEJDLU AS STRING),2,'0'))) AS STRING)
          ELSE '1900/01/01' END AS DATA_PRZYD
      FROM \`fml-prod-rep-pl-rfx-sbe.last_rows_sbe.hlgeinp\` hl_gei
      LEFT JOIN \`fml-prod-rep-pl-rfx-sbe.last_rows_sbe.hlrecpp\` hl_rec_entete 
        ON hl_gei.GECACT = hl_rec_entete.RECACT AND hl_gei.GECDPR = hl_rec_entete.RECDPO 
        AND hl_gei.GENANN = hl_rec_entete.RENANN AND hl_gei.GENREC = hl_rec_entete.RENREC
      WHERE hl_gei.GECACT IN ('MTE') AND hl_gei.GECTST = '200'
    ),
    stock_data AS (
      SELECT 
        hl_gei.GECACT,
        hl_gei.GECART AS ItemCode,
        hl_gei.GEQGEI AS LotQuantity_raw,
        hl_gei.GENSUP AS SSCC,
        hl_gei.GELOTF AS BATCH,
        hl_emplacement.EMC1EM, hl_emplacement.EMC2EM, hl_emplacement.EMC3EM, hl_emplacement.EMC4EM, hl_emplacement.EMC5EM,
        m.pcb, m.spcb, m.base_lv_on_pallet,
        d.DATA_PRZYD,
        ts.TULTSU AS SupportTypeCaption
      FROM \`fml-prod-rep-pl-rfx-sbe.last_rows_sbe.hlgeinp\` hl_gei
      JOIN \`fml-prod-rep-pl-rfx-sbe.last_rows_sbe.hlsuppp\` hl_support ON hl_support.SUCACT = hl_gei.GECACT AND hl_support.SUNSUP = hl_gei.GENSUP
      JOIN \`fml-prod-rep-pl-rfx-sbe.last_rows_sbe.hlemplp\` hl_emplacement ON hl_support.SUCDPO = hl_emplacement.EMCDPO AND hl_support.SUNEMP = hl_emplacement.EMNEMP
      JOIN \`fml-prod-rep-pl-rfx-sbe.last_rows_sbe.hltysup\` ts ON ts.TUCTSU = hl_support.SUCTSU
      LEFT JOIN masterdata m ON hl_gei.GECART = m.Article
      LEFT JOIN daty d ON hl_gei.GENSUP = d.GENSUP AND hl_gei.GECART = d.GECART AND hl_gei.GENGEI = d.GENGEI
      WHERE hl_gei.GECTST = '200'
    )
    SELECT DISTINCT
        b.Building AS Hala,
        b.Alley AS Alejka,
        b.Location AS Location,
        b.Level AS Level,
        b.Cell AS Cell,
        CASE
            WHEN REGEXP_CONTAINS(b.GeographicalZoneCode, '(?i)MTE') THEN 'MTE'
            WHEN REGEXP_CONTAINS(b.GeographicalZoneCode, '(?i)GLE') THEN 'GLE'
            WHEN REGEXP_CONTAINS(b.GeographicalZoneCode, '(?i)CDY') THEN 'CDY'
            WHEN REGEXP_CONTAINS(b.GeographicalZoneCode, '(?i)YRP') THEN 'YRP'
            WHEN REGEXP_CONTAINS(b.GeographicalZoneCode, '(?i)LNP') THEN 'LNP'
            ELSE 'MCE'
        END AS Klient,
        CASE
            WHEN s.ItemCode IS NOT NULL THEN 'Zajęte'
            WHEN b.IsBlockedEntrance = '1' THEN 'Zaholdowane'
            ELSE 'Wolne'
        END AS Status_Miejsca,
        COALESCE(s.ItemCode, 'PUSTA') AS ItemCode,
        COALESCE(s.LotQuantity_raw / NULLIF(s.pcb, 0), 0) AS LotQuantity,
        COALESCE(s.SSCC, 'BRAK') AS SSCC,
        COALESCE(s.BATCH, 'BRAK') AS BATCH,
        COALESCE(CAST(s.pcb AS INT64), 0) AS NbParcelPerPallet,
        COALESCE(s.DATA_PRZYD, '1900/01/01') AS BestBeforeDate,
        COALESCE(s.SupportTypeCaption, 'BRAK') AS SupportTypeCaption

    FROM
        \`fml-prod-rep-pl-rfx-sbe.BO.wh_depot_sbe\` AS b
    LEFT JOIN
        stock_data AS s ON b.Location = CONCAT(s.EMC1EM, s.EMC2EM, s.EMC3EM, s.EMC4EM, s.EMC5EM)

    WHERE
        (b.PhysicalDepot = "SBE"
        AND b.Building in ("B01") 
        AND NOT REGEXP_CONTAINS(LOWER(b.Location), r'usz')
        AND b.Location NOT IN ('MISSING', 'QUALITY', 'ANULM', 'ANULS', 'ANULY', 'UTYLIZAC', 'UTYL1MEB', 'REPA.LET', 'RETAINS', 'NADW', 'WYCI', 'KARTONY', 'KIERMASZ')
        AND b.Location NOT LIKE 'BUFOR%'
        AND b.Location NOT LIKE 'LABEL%'
        AND b.Location NOT LIKE 'CLOUD%'
        AND b.Location NOT LIKE 'TRAN%'
        AND b.Location NOT LIKE 'HUB%'
        AND b.Location NOT LIKE 'COP%'
        AND b.Location NOT LIKE 'DMC%'
        and Alley not in ('ZWR',"ZMIA","TEMP","SDB","YRP","WYS","REG.","PARC","OBNI","MEZA","MTE","DYRP","LINE","KONS","KONT","J04","J03","J02","J01","DSDB","DMTE","DGLE","CONV","CLOU")
        AND b.Location LIKE '%.%' 
        
        and (SAFE_CAST(b.Alley AS INT64) >= 152 or level>"10"))
        or (Building="B02" and REGEXP_CONTAINS(b.GeographicalZoneCode, '(?i)MTE')
        AND ((level="10" and b.Alley="0259") or (level="10" and b.Alley="0262") or (level="10" and b.Alley="0260") or (level="10" and b.Alley="0261") or safe_cast(b.Alley as INT64)<=258)
        
        AND NOT REGEXP_CONTAINS(LOWER(b.Location), r'usz')
        AND b.Location NOT IN ('MISSING', 'QUALITY', 'ANULM', 'ANULS', 'ANULY', 'UTYLIZAC', 'UTYL1MEB', 'REPA.LET', 'RETAINS', 'NADW', 'WYCI', 'KARTONY', 'KIERMASZ')
        AND b.Location NOT LIKE 'BUFOR%'
        AND b.Location NOT LIKE 'LABEL%'
        AND b.Location NOT LIKE 'CLOUD%'
        AND b.Location NOT LIKE 'TRAN%'
        AND b.Location NOT LIKE 'HUB%'
        AND b.Location NOT LIKE 'COP%'
        AND b.Location NOT LIKE 'DMC%'
        and Alley not in ('ZWR',"ZMIA","TEMP","SDB","YRP","WYS","REG.","PARC","OBNI","MEZA","MTE","DYRP","LINE","KONS","KONT","J04","J03","J02","J01","DSDB","DMTE","DGLE","CONV","CLOU")
        AND b.Location LIKE '%.%' )
        #and ItemCode not LIKE "FAAA%"
             

    ORDER BY 
       b.Alley ASC,
        b.Location ASC 
  `;

  try {
    let job = {
      configuration: {
        query: {
          query: sql,
          useLegacySql: false
        }
      }
    };

    let queryJob = BigQuery.Jobs.insert(job, projectId);
    const jobId = queryJob.jobReference.jobId;

    let status = BigQuery.Jobs.get(projectId, jobId);
    while (status.status.state === 'RUNNING') {
      Utilities.sleep(1000);
      status = BigQuery.Jobs.get(projectId, jobId);
    }

    if (status.status.errorResult) {
      throw new Error("Błąd SQL BigQuery: " + status.status.errorResult.message);
    }

    const results = BigQuery.Jobs.getQueryResults(projectId, jobId);
    if (!results.rows || results.rows.length === 0) {
      ui.alert("Brak danych w BigQuery.");
      return;
    }

    const data = results.rows.map(row => row.f.map(col => col.v));

    if (masterSheet.getLastRow() > 1) {
      masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 15).clearContent();
    }
    
    masterSheet.getRange(2, 1, data.length, data[0].length).setValues(data);

    distributeInitialData();
    
    ui.alert("Sukces! Pobrano " + data.length + " wierszy i zaktualizowano Master Stock.");

  } catch (e) {
    ui.alert("Błąd krytyczny: " + e.message);
  }
}
function exportToFormatka() {
  const ui = SpreadsheetApp.getUi();
  const reportSheet = SS.getSheetByName(REPORT_SHEET_NAME);
  const formatkaSheet = SS.getSheetByName("formatka") || SS.insertSheet("formatka");
  
  const data = reportSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert("Raport jest pusty!");
    return;
  }

  const headers = [
    "Lokalizacja", "Alejka", "", "Przęsło", "Piętro", "Miejsce", "SKU", "Opis", 
    "BBD", "Partia", "Sztuki w RFX", "PCB", "Zarezerwowane", 
    "I Liczenie", "LISTA", "KOMENTARZ", "Wynik", 
    "II liczenie", "LISTA", "KOMENTARZ", "Wynik 2", 
    "Komisyjne", "LISTA", "KOMENTARZ 2", "Działanie", 
    "Ilość zliczona po III liczeniu", "Różnica"
  ];

  let output = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const locRaw = String(r[2]);
    const details = parseLocationExtended(locRaw);
    const rowIdx = i + 1; 

    
    let row = [
      locRaw, details.alley, ".", details.bay, details.lvl, details.pos, r[3], "", 
      r[13], r[4], r[5], "", "", 
      r[6], "", r[9], "", // I liczenie
      r[7], "", r[14], "", // II liczenie
      "", "", "",          // Komisyjne (V, W, X)
      "", "", ""           // Działanie (Y), III Licz (Z), Różnica (AA)
    ];
    output.push(row);
  }

  formatkaSheet.clear();
  formatkaSheet.getRange(1, 1, 1, headers.length).setValues([headers])
               .setBackground("#1e293b").setFontColor("white").setFontWeight("bold");
  
  if (output.length > 0) {
    formatkaSheet.getRange(2, 1, output.length, headers.length).setValues(output);
    
    
    const lastRow = output.length + 1;

    
    formatkaSheet.getRange(2, 25, output.length, 1).setFormulas(
      Array.from({length: output.length}, (_, i) => [`=IF(OR(K${i+2}=N${i+2}; K${i+2}=R${i+2}); "OK"; "Modyfikuj")`])
    );

    
    
    formatkaSheet.getRange(2, 27, output.length, 1).setFormulas(
      Array.from({length: output.length}, (_, i) => [
        `=IF(V${i+2}<>""; V${i+2}-K${i+2}; IF(R${i+2}<>""; R${i+2}-K${i+2}; N${i+2}-K${i+2}))`
      ])
    );
  }

  formatkaSheet.autoResizeColumns(1, headers.length);
  ui.alert("Formatka wygenerowana z formułami!");
}


function parseLocationExtended(locStr) {
  let parts = String(locStr).split('.');
  let alley = parts[0] || "";
  let sub = parts[1] || "0000000";
  
  return {
    alley: alley,
    bay: sub.substring(0, 3),   // 001
    lvl: sub.substring(3, 5),   // 10
    pos: sub.substring(5)       // 01
  };
}
