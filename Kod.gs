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
    if (!alreadyLocked && lock.hasLock()) {
      lock.releaseLock();
    }
    console.error("Błąd doPost: " + err.message);
    return ContentService.createTextOutput(JSON.stringify({"success": false, "error": err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
function updateInventoryBatchStatus(sheet, batchId, status) {
  const batch = findInventoryBatch(sheet, batchId);


  if (!batch) {
    throw new Error("Nie znaleziono partii: " + batchId);
  }


  sheet.getRange(batch.row, 3, 1, 2).setValues([[
    status,
    new Date()
  ]]);


  SpreadsheetApp.flush();
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
function getInventorySyncLog() {
  const name = "Inventory_Sync_Log";
  let sheet = SS.getSheetByName(name);


  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow([
      "BatchId",
      "Team",
      "Status",
      "UpdatedAt"
    ]);
  }


  return sheet;
}
function findInventoryBatch(sheet, batchId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;


  const cell = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(batchId)
    .matchEntireCell(true)
    .findNext();


  if (!cell) return null;


  const row = cell.getRow();


  return {
    row: row,
    team: sheet.getRange(row, 2).getValue(),
    status: sheet.getRange(row, 3).getValue()
  };
}
function saveAllTasks(teamName, allTasks, startTime, batchId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(50000); 
    if (!batchId) {
      throw new Error("Brak BatchId");
    }
    
    
    const syncLog = getInventorySyncLog();
    const existingBatch = findInventoryBatch(syncLog, batchId);
    
    
    if (existingBatch) {
      if (existingBatch.team !== teamName) {
        throw new Error("BatchId należy do innej ekipy");
      }
      if (existingBatch.status === "REPORT_SAVED") {
        const teamSheet = SS.getSheetByName(teamName);
      
      
        if (teamSheet.getLastRow() >= 2) {
          throw new Error(
            "REPORT_SAVED: arkusz ekipy nie jest pusty. Wymagana weryfikacja."
          );
        }
      
      
        syncAndGetNewData(teamName, true);
        SpreadsheetApp.flush();
      
      
        updateInventoryBatchStatus(
          syncLog,
          batchId,
          "TASKS_ASSIGNED"
        );
      
      
        const nextData = getInitialDataForLoop(teamName);
      
      
        updateInventoryBatchStatus(
          syncLog,
          batchId,
          "DONE"
        );
      
      
        if (lock.hasLock()) {
          lock.releaseLock();
        }
      
      
        return {
          success: true,
          alreadySaved: true,
          hasNewTasks: nextData.tasks.length > 0,
          newData: nextData
        };
      }
      if (existingBatch.status === "TASKS_ASSIGNED") {
        const nextData = getInitialDataForLoop(teamName);
      
      
        updateInventoryBatchStatus(
          syncLog,
          batchId,
          "DONE"
        );
      
      
        if (lock.hasLock()) {
          lock.releaseLock();
        }
      
      
        return {
          success: true,
          alreadySaved: true,
          hasNewTasks: nextData.tasks.length > 0,
          newData: nextData
        };
      }
      if (existingBatch.status === "DONE") {
        const nextData = getInitialDataForLoop(teamName);
      
      
        if (lock.hasLock()) {
          lock.releaseLock();
        }
    
    
        return {
          success: true,
          alreadySaved: true,
          hasNewTasks: nextData.tasks.length > 0,
          newData: nextData
        };
      }
    
    
      throw new Error(
        "Partia wymaga weryfikacji: " + existingBatch.status
      );
    }
    
    
    
    
    
    
      
    const toNum = (val) => {
      if (val === "" || val === null || val === undefined) return 0;
      let n = parseFloat(String(val).replace(/\s+/g, '').replace(',', '.'));
      return isNaN(n) ? 0 : n;
    };

    const teamSheet = SS.getSheetByName(teamName);
    const lastRow = teamSheet.getLastRow();
    if (lastRow < 2) {
      if (lock.hasLock()) {
        lock.releaseLock();
      }
      return { success: true, msg: "Brak danych." };
    }

    const currentTasks = teamSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    if (!Array.isArray(allTasks)) {
    throw new Error("Nieprawidłowe dane skanera");
  }
  
  
  if (allTasks.length !== currentTasks.length) {
    throw new Error(
      "Niezgodna liczba zadań. Skaner: " +
      allTasks.length +
      ", arkusz: " +
      currentTasks.length
    );
  }
  
  
  const sentTasks = new Map();
  
  
  allTasks.forEach(t => {
    const key = JSON.stringify([
      String(t.loc),
      String(t.sku),
      Number(t.masterRow)
    ]);
  
  
    if (sentTasks.has(key)) {
      throw new Error("Duplikat zadania: " + t.loc);
    }
  
  
    if (
      t.qty === "" ||
      t.qty === null ||
      t.qty === undefined ||
      !Number.isFinite(Number(String(t.qty).replace(",", ".")))
    ) {
      throw new Error("Brak lub błędna ilość: " + t.loc);
    }
  
  
    sentTasks.set(key, t);
  });
  
    currentTasks.forEach(row => {
      const key = JSON.stringify([
        String(row[0]),
        String(row[1]),
        Number(row[9])
      ]);
    
    
      if (!sentTasks.has(key)) {
        throw new Error("Brak wyniku: " + row[0]);
      }
    });
  
    const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
    syncLog.appendRow([
      batchId,
      teamName,
      "STARTED",
      new Date()
    ]);
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
            comment: comment,
            batchId: batchId
          });
      
      
          if (masterRowIdx && masterRowIdx <= masterLastRow) {
            masterStatusValues[masterRowIdx - 1][0] =
              "ZAKOŃCZONE FINALNIE (" + teamName + ")";
          }
      
      
        } else {
            const newRow = buildReportRow(
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
            );
          
          
            // U: BatchId pierwszego liczenia
            // V: BatchId drugiego liczenia
            newRow.push("", batchId);
          
          
            newReportEntries.push(newRow);
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
      // H:V — обновляем результат и BatchId вместе
      const range = reportSheet.getRange(update.row, 8, 1, 15);
      const values = range.getValues()[0];
    
    
      values[0] = update.qty;      // H: 2. Liczenie
      values[1] = update.status;   // I: Status
      values[7] = update.comment;  // O: Komentarz2
      values[14] = update.batchId; // V: BatchId 2. liczenia
    
    
      range.setValues([values]);
    });

    // 3. Dodanie nowych wierszy do Raportu
    if (newReportEntries.length > 0) {
      const rowsWithBatch = newReportEntries.map(row => {
        if (row.length === 22) {
          return row;
        }
    
    
        return [
          ...row,
          batchId,
          ""
        ];
      });
    
    
      reportSheet
        .getRange(
          reportSheet.getLastRow() + 1,
          1,
          rowsWithBatch.length,
          22
        )
        .setValues(rowsWithBatch);
    }

    // Raport i Master zostały zapisane
    SpreadsheetApp.flush();
    
    
    updateInventoryBatchStatus(
      syncLog,
      batchId,
      "REPORT_SAVED"
    );
    
    
    // Czyszczenie i nowa paczka
    teamSheet.getRange("A2:K").clearContent()
    SpreadsheetApp.flush(); // Wymuś zapis przed sync
    
    syncAndGetNewData(teamName,true);


    SpreadsheetApp.flush();
    
    
    updateInventoryBatchStatus(
      syncLog,
      batchId,
      "TASKS_ASSIGNED"
    );
    
    
    const nextData = getInitialDataForLoop(teamName);


    updateInventoryBatchStatus(
      syncLog,
      batchId,
      "DONE"
    );
    if (lock.hasLock()) {
      lock.releaseLock();
    }
    return { success: true, hasNewTasks: nextData.tasks.length > 0, newData: nextData };

  } catch (e) {
    if (lock && lock.hasLock()) lock.releaseLock();
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
function syncAndGetNewData(manualTeamName, alreadyLocked = false) {
  const sheet = manualTeamName ? SS.getSheetByName(manualTeamName) : SS.getActiveSheet();
  const teamName = sheet.getName();
  if (!teamName.startsWith("Ekipa_")) return;

  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const teamNum = parseInt(teamName.split("_")[1]);
  const lock = LockService.getScriptLock();

  try {
    if (!alreadyLocked){
      lock.waitLock(50000); 
    }
    const lastRowMaster = masterSheet.getLastRow();
    


    if (lastRowMaster < 2) {
      if (!alreadyLocked && lock.hasLock()) {
        lock.releaseLock();
      }
      return;
    }

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

    if (!alreadyLocked && lock.hasLock()) {
      lock.releaseLock();
    }
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
    if (!alreadyLocked && lock.hasLock()) {
      lock.releaseLock();
    }


    console.error("Błąd sync: " + e.message);


    if (alreadyLocked) {
      throw e;
    }
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
    syncAndGetNewData(teamName,true); 
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
      syncAndGetNewData(teamName,true);
      SpreadsheetApp.flush(); 
      Utilities.sleep(500); 
    }
    lock.releaseLock();
    ui.alert("Sukces! Rozdzielono pracę na wszystkie ekipy (1-13).");
  } catch (e) {
    if (!alreadyLocked && lock.hasLock()) {
      lock.releaseLock();
    }
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
function test15Scanners170() {
  const SEND_RESULTS = true; // true — реально сохранить результаты
  const PER_TEAM = 170;
  const TEAM_COUNT = 15;
  const REQUIRED = PER_TEAM * TEAM_COUNT;


  const masterSheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const master = masterSheet.getDataRange().getValues();


  // Берём только свободные строки Master.
  const available = [];


  for (let i = 1; i < master.length; i++) {
    const r = master[i];


    if (
      String(r[14]).trim() === "" &&
      String(r[2]).trim() !== "" &&
      String(r[7]).trim() !== ""
    ) {
      available.push({
        rowNum: i + 1,
        data: r
      });
    }
  }


  if (available.length < REQUIRED) {
    throw new Error(
      "Za mało wolnych pozycji Master: " +
      available.length + " / " + REQUIRED
    );
  }


  // Проверяем наличие всех листов до изменений.
  for (let i = 1; i <= TEAM_COUNT; i++) {
    if (!SS.getSheetByName("Ekipa_" + i)) {
      throw new Error("Brak arkusza Ekipa_" + i);
    }
  }


  const testId = Utilities.getUuid().slice(0, 8);
  const results = [];
  let index = 0;


  // Подготавливаем ровно 170 заданий на команду.
  for (let team = 1; team <= TEAM_COUNT; team++) {
    const teamName = "Ekipa_" + team;
    const sheet = SS.getSheetByName(teamName);
    const rows = [];


    for (let j = 0; j < PER_TEAM; j++) {
      const item = available[index++];
      const r = item.data;


      rows.push([
        r[2],                 // A: Lokalizacja
        r[7],                 // B: SKU
        "",                   // C
        r[8],                 // D: Ilość systemowa
        "1. LICZENIE",        // E: Typ
        "",                   // F
        r[10],                // G: Batch
        r[8],                 // H
        r[11],                // I
        item.rowNum,          // J: MasterRow
        r[13] || "BRAK"       // K
      ]);


      // Резервируем позицию за командой.
      r[14] = "ZAREZERWOWANE (" + teamName + ")";
    }


    sheet.getRange("A2:K").clearContent();
    sheet.getRange(2, 1, PER_TEAM, 11).setValues(rows);


    results.push({
      team: teamName,
      prepared: rows.length,
      status: "PREPARED"
    });
  }


  // Сохраняем резервирования в Master.
  masterSheet.getRange(2, 15, master.length - 1, 1)
    .setValues(master.slice(1).map(r => [r[14]]));


  SpreadsheetApp.flush();


  if (!SEND_RESULTS) {
    console.log(
      "PREPARED: " + REQUIRED +
      " tasks. Set SEND_RESULTS = true to save."
    );
    console.log(JSON.stringify(results, null, 2));
    return;
  }


  // Отправляем результаты через реальную saveAllTasks().
  for (let team = 1; team <= TEAM_COUNT; team++) {
    const teamName = "Ekipa_" + team;
    const data = getInitialDataForLoop(teamName);


    const tasks = data.tasks.map(t => ({
      loc: t.loc,
      sku: t.sku,
      masterRow: t.masterRow,
      qty: Number(t.sysQty) || 0,
      comment: "TEST_170"
    }));


    const batchId = "LOADTEST-" + testId + "-" + teamName;
    const started = Date.now();


    const response = saveAllTasks(
      teamName,
      tasks,
      data.startTime,
      batchId
    );


    const result = results[team - 1];
    result.status = response.success ? "OK" : "ERROR";
    result.seconds = (Date.now() - started) / 1000;
    result.batchId = batchId;
    result.message = response.msg || "";


    console.log(
      teamName + ": " +
      result.status + " | " +
      result.seconds + "s | " +
      result.message
    );
  }


  console.log(JSON.stringify(results, null, 2));
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
  syncAndGetNewData(teamName,true); 
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
        WHERE VLCACT IN ('LNP')
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
      WHERE hl_gei.GECACT IN ('LNP') AND hl_gei.GECTST = '200'
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
        AND b.Building in ("B09") 
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
function start15ScannersTest() {
  const COUNT = 15;
  const RUN_ID = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();


  let totalTasks = 0;


  for (let i = 1; i <= COUNT; i++) {
      const team = "Ekipa_" + i;
      const data = getInitialDataForLoop(team);


      const count = data.tasks.length;
      totalTasks += count;


      console.log(team + ": " + count + " zadań");
  }


  // Создаём отдельный журнал теста.
  const log = SS.getSheetByName("Test_15_Scanners")
    || SS.insertSheet("Test_15_Scanners");


  log.appendRow([
    "RunId", "Team", "BatchId", "Status",
    "Start", "End", "Seconds", "Message"
  ]);


  // 15 независимых запусков примерно через минуту.
  for (let i = 1; i <= COUNT; i++) {
    const trigger = ScriptApp.newTrigger("runOneScannerTest")
      .timeBased()
      .after(60000)
      .create();


    props.setProperty(
      "TEST_TRIGGER_" + trigger.getUniqueId(),
      JSON.stringify({
        runId: RUN_ID,
        team: "Ekipa_" + i,
        batchId: RUN_ID + "-Ekipa_" + i
      })
    );
  }


  console.log(
    "Созданы " + COUNT + " триггеров. " +
    "Всего выдано заданий: " + totalTasks +
    ". RunId: " + RUN_ID
  );
}




function runOneScannerTest(e) {
  const props = PropertiesService.getScriptProperties();
  const triggerId = String(e.triggerUid);
  const key = "TEST_TRIGGER_" + triggerId;
  const raw = props.getProperty(key);


  if (!raw) {
    throw new Error("Не найдены параметры тестового запуска");
  }


  const config = JSON.parse(raw);
  const team = config.team;
  const started = new Date();
  let status = "ERROR";
  let message = "";


  try {
    const data = getInitialDataForLoop(team);


    


    const tasks = data.tasks.map(t => ({
      loc: t.loc,
      sku: t.sku,
      masterRow: t.masterRow,
      qty: Number(t.sysQty) || 0,
      comment: "LOAD_TEST_15"
    }));


    // Настоящее сохранение, как со сканера.
    const response = testSaveWithQueue(
      team,
      tasks,
      data.startTime,
      config.batchId
    );


    if (!response.success) {
      throw new Error(response.msg || "Ошибка сохранения");
    }


    status = "OK";


  } catch (err) {
    message = err.message;


  } finally {
    const finished = new Date();


    // Удаляем только свой триггер.
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(t);
      }
    });


    props.deleteProperty(key);


    const log = SS.getSheetByName("Test_15_Scanners");


    // Отдельная блокировка для записи журнала.
    const logLock = LockService.getDocumentLock()
      || LockService.getScriptLock();


    try {
      logLock.waitLock(30000);


      log.appendRow([
        config.runId,
        team,
        config.batchId,
        status,
        started,
        finished,
        (finished - started) / 1000,
        message
      ]);


      SpreadsheetApp.flush();


    } finally {
      if (logLock.hasLock()) {
        logLock.releaseLock();
      }
    }


    console.log(
      team + " | " + status +
      " | " + ((finished - started) / 1000) +
      "s | " + message
    );
  }
}
function testSaveWithQueue(teamName, tasks, startTime, batchId) {
  const maxAttempts = 10;


  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const result = saveAllTasks(
      teamName,
      tasks,
      startTime,
      batchId
    );


    if (result.success) {
      return {
        ...result,
        attempts: attempt
      };
    }


    const msg = String(result.msg || "");


    // Повторяем только ошибку блокировки.
    if (!/lock|blokad|zablokowan/i.test(msg)) {
      return result;
    }


    if (attempt < maxAttempts) {
      Utilities.sleep(5000);
    }
  }


  return {
    success: false,
    msg: "Przekroczono limit prób zapisu."
  };
}
/**
 * Kolejka inwentaryzacji — wersja testowa.
 * Dodać jako osobny plik .gs w projekcie Apps Script.
 * Nie modyfikuje saveAllTasks() ani interfejsu skanera.
 */
const IQ = {
  QUEUE: 'Inventory_Queue',
  DATA: 'Inventory_Queue_Data',
  TEST: 'Test_15_Scanners_Queue',
  CHUNK: 30000,
  STALE_MS: 20 * 60 * 1000
};

function auditInventoryQueue() {
  const queue = SS.getSheetByName('Inventory_Queue');
  const report = SS.getSheetByName(REPORT_SHEET_NAME);
  const syncLog = SS.getSheetByName('Inventory_Sync_Log');


  if (!queue || !report) {
    throw new Error('Brak arkusza kolejki lub raportu');
  }


  const queueRows = queue.getLastRow() > 1
    ? queue.getRange(2, 1, queue.getLastRow() - 1, 9).getValues()
    : [];


  const reportRows = report.getLastRow() > 1
    ? report.getRange(2, 1, report.getLastRow() - 1, 22).getValues()
    : [];


  const syncRows = syncLog && syncLog.getLastRow() > 1
    ? syncLog.getRange(2, 1, syncLog.getLastRow() - 1, 4).getValues()
    : [];


  const reportCounts = new Map();
  const syncStatuses = new Map();


  // U = pierwsze liczenie, V = drugie liczenie.
  reportRows.forEach(row => {
    const first = String(row[20] || '').trim();
    const second = String(row[21] || '').trim();


    if (first) {
      reportCounts.set(
        first,
        (reportCounts.get(first) || 0) + 1
      );
    }


    if (second) {
      reportCounts.set(
        second,
        (reportCounts.get(second) || 0) + 1
      );
    }
  });


  syncRows.forEach(row => {
    syncStatuses.set(String(row[0]), String(row[2]));
  });


  let audit = SS.getSheetByName('Inventory_Queue_Audit');


  if (!audit) {
    audit = SS.insertSheet('Inventory_Queue_Audit');
  }


  audit.clearContents();


  const output = [[
    'BatchId',
    'Team',
    'QueueStatus',
    'SyncStatus',
    'ExpectedTasks',
    'ReportMatches',
    'Difference',
    'AuditStatus'
  ]];


  queueRows.forEach(row => {
    const batchId = String(row[0]);
    const team = String(row[1]);
    const queueStatus = String(row[2]);
    const expected = Number(row[6]) || 0;


    const syncStatus = syncStatuses.get(batchId) || 'BRAK';
    const actual = reportCounts.get(batchId) || 0;
    const difference = actual - expected;


    let auditStatus = 'PENDING';


    if (queueStatus === 'DONE' && syncStatus === 'DONE') {
      auditStatus = difference === 0
        ? 'OK'
        : 'MISMATCH';
    } else if (
      queueStatus === 'NEEDS_REVIEW' ||
      queueStatus === 'DONE'
    ) {
      auditStatus = 'CHECK';
    }


    output.push([
      batchId,
      team,
      queueStatus,
      syncStatus,
      expected,
      actual,
      difference,
      auditStatus
    ]);
  });


  audit.getRange(1, 1, output.length, 8).setValues(output);
  audit.setFrozenRows(1);
  SpreadsheetApp.flush();


  const summary = {
    total: output.length - 1,
    ok: output.slice(1).filter(r => r[7] === 'OK').length,
    mismatch: output.slice(1).filter(r => r[7] === 'MISMATCH').length,
    check: output.slice(1).filter(r => r[7] === 'CHECK').length,
    pending: output.slice(1).filter(r => r[7] === 'PENDING').length
  };


  console.log(JSON.stringify(summary));
  return summary;
}
function inventoryQueueSheets_() {
  let q = SS.getSheetByName(IQ.QUEUE);
  if (!q) {
    q = SS.insertSheet(IQ.QUEUE);
    q.appendRow(['BatchId', 'Team', 'Status', 'CreatedAt', 'StartedAt',
      'FinishedAt', 'TaskCount', 'ChunkCount', 'Message']);
  }
  let d = SS.getSheetByName(IQ.DATA);
  if (!d) {
    d = SS.insertSheet(IQ.DATA);
    d.appendRow(['BatchId', 'Part', 'Payload']);
  }
  return { q, d };
}


function inventoryQueueFind_(sheet, batchId) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const cell = sheet.getRange(2, 1, last - 1, 1)
    .createTextFinder(String(batchId)).matchEntireCell(true).findNext();
  if (!cell) return null;
  const row = cell.getRow();
  const v = sheet.getRange(row, 1, 1, 9).getValues()[0];
  return { row, batchId: v[0], team: v[1], status: v[2],
    createdAt: v[3], startedAt: v[4], finishedAt: v[5],
    count: Number(v[6]), chunks: Number(v[7]), message: String(v[8] || '') };
}


/**
 * Zapisuje pełną paczkę, zanim zostanie oznaczona jako QUEUED.
 * Powtórne zgłoszenie tego samego BatchId nie tworzy duplikatu.
 */
function enqueueInventory(teamName, allTasks, startTime, batchId) {
  if (!/^Ekipa_\d+$/.test(String(teamName))) {
    return { success: false, msg: 'Nieprawidłowa nazwa ekipy' };
  }
  if (!batchId || !Array.isArray(allTasks) || !allTasks.length) {
    return { success: false, msg: 'Brak BatchId lub zadań' };
  }
  if (!allTasks.every(t => t && t.loc != null && t.sku != null &&
      Number.isFinite(Number(t.masterRow)) && Number(t.masterRow) > 0 &&
      t.qty !== '' && t.qty != null &&
      Number.isFinite(Number(String(t.qty).replace(',', '.'))))) {
    return { success: false, msg: 'Nieprawidłowe dane zadań' };
  }


  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    return { success: false, retryable: true, msg: 'LOCK_BUSY: kolejka zajęta' };
  }
  try {
    const { q, d } = inventoryQueueSheets_();
    const existing = inventoryQueueFind_(q, batchId);
    if (existing) {
      if (existing.team !== teamName) {
        return { success: false, msg: 'BatchId należy do innej ekipy' };
      }
      return { success: true, batchId, status: existing.status,
        taskCount: existing.count, alreadyQueued: true };
    }


    const payload = JSON.stringify({ teamName, allTasks, startTime, batchId });
    const chunks = [];
    for (let i = 0; i < payload.length; i += IQ.CHUNK) {
      chunks.push([String(batchId), chunks.length + 1,
        payload.substring(i, i + IQ.CHUNK)]);
    }


    // Najpierw dane; dopiero po ich zapisaniu rejestrujemy gotowe zgłoszenie.
    d.getRange(d.getLastRow() + 1, 1, chunks.length, 3).setValues(chunks);
    SpreadsheetApp.flush();
    q.appendRow([String(batchId), teamName, 'QUEUED', new Date(), '', '',
      allTasks.length, chunks.length, '']);
    SpreadsheetApp.flush();
    return { success: true, batchId, status: 'QUEUED',
      taskCount: allTasks.length };
  } catch (err) {
    return { success: false, msg: String(err.message || err) };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}


function inventoryQueueReadPayload_(dataSheet, batchId, expectedChunks) {
  const last = dataSheet.getLastRow();
  if (last < 2) throw new Error('Brak danych paczki');
  const rows = dataSheet.getRange(2, 1, last - 1, 3).getValues()
    .filter(r => String(r[0]) === String(batchId))
    .sort((a, b) => Number(a[1]) - Number(b[1]));
  if (rows.length !== expectedChunks ||
      rows.some((r, i) => Number(r[1]) !== i + 1)) {
    throw new Error('Niekompletne fragmenty danych paczki');
  }
  const payload = JSON.parse(rows.map(r => String(r[2])).join(''));
  if (payload.batchId !== batchId ||
      !Array.isArray(payload.allTasks)) {
    throw new Error('Niezgodność danych paczki');
  }
  return payload;
}


function getInventoryQueueStatus(batchId) {
  const q = SS.getSheetByName(IQ.QUEUE);
  if (!q) return { found: false };
  const item = inventoryQueueFind_(q, batchId);
  if (!item) return { found: false };
  return { found: true, batchId: item.batchId, team: item.team,
    status: item.status, taskCount: item.count, message: item.message };
}


function processInventoryQueue() {
  const lock = LockService.getScriptLock();


  if (!lock.tryLock(1000)) {
    return { status: 'BUSY' };
  }


  let item = null;


  try {
    const { q } = inventoryQueueSheets_();
    const last = q.getLastRow();


    if (last < 2) {
      return { status: 'EMPTY' };
    }


    const rows = q.getRange(2, 1, last - 1, 9).getValues();


    // Kontrola przerwanego przetwarzania.
    const active = rows.findIndex(r => r[2] === 'PROCESSING');


    if (active >= 0) {
      const row = active + 2;
      const batchId = String(rows[active][0]);
      const startedAt = new Date(rows[active][4]).getTime();
      const age = Date.now() - startedAt;


      if (!Number.isFinite(age) || age < IQ.STALE_MS) {
        return { status: 'BUSY', batchId };
      }


      const syncStatus = iqGetSyncStatus_(batchId);


      if (syncStatus === 'DONE') {
        q.getRange(row, 3).setValue('DONE');
        q.getRange(row, 6).setValue(new Date());
        q.getRange(row, 9).setValue(
          'Potwierdzono zakończenie w Inventory_Sync_Log'
        );
      } else {
        q.getRange(row, 3).setValue('NEEDS_REVIEW');
        q.getRange(row, 9).setValue(
          'Przerwane przetwarzanie. Status synchronizacji: ' +
          syncStatus
        );
      }


      rows[active][2] = syncStatus === 'DONE'
        ? 'DONE'
        : 'NEEDS_REVIEW';


      SpreadsheetApp.flush();
    }


    // Dawne zadania oczekujące na odzyskanie wymagają kontroli.
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][2] !== 'RECOVERY_PENDING') {
        continue;
      }


      const batchId = String(rows[i][0]);
      const syncStatus = iqGetSyncStatus_(batchId);
      const status = syncStatus === 'DONE'
        ? 'DONE'
        : 'NEEDS_REVIEW';


      q.getRange(i + 2, 3).setValue(status);
      q.getRange(i + 2, 9).setValue(
        'Kontrola po przerwaniu. Status synchronizacji: ' +
        syncStatus
      );


      if (status === 'DONE') {
        q.getRange(i + 2, 6).setValue(new Date());
      }


      rows[i][2] = status;
    }


    // Wybieramy tylko jedno nowe zgłoszenie.
    const index = rows.findIndex(r => r[2] === 'QUEUED');


    if (index < 0) {
      SpreadsheetApp.flush();
      return { status: 'EMPTY' };
    }


    item = {
      row: index + 2,
      batchId: String(rows[index][0]),
      team: String(rows[index][1]),
      count: Number(rows[index][6]),
      chunks: Number(rows[index][7])
    };


    q.getRange(item.row, 3).setValue('PROCESSING');
    q.getRange(item.row, 5).setValue(new Date());
    q.getRange(item.row, 9).setValue('');


    SpreadsheetApp.flush();


  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }


  let status = 'NEEDS_REVIEW';
  let message = '';


  try {
    const dataSheet = SS.getSheetByName(IQ.DATA);


    const payload = inventoryQueueReadPayload_(
      dataSheet,
      item.batchId,
      item.chunks
    );


    if (
      payload.teamName !== item.team ||
      payload.allTasks.length !== item.count
    ) {
      throw new Error('Niezgodność danych paczki');
    }


    const result = saveAllTasks(
      payload.teamName,
      payload.allTasks,
      payload.startTime,
      payload.batchId
    );


    const syncStatus = iqGetSyncStatus_(item.batchId);


    if (result && result.success && syncStatus === 'DONE') {
      status = 'DONE';
    } else {
      message = String(
        result && result.msg ||
        'Brak potwierdzenia zakończenia zapisu'
      );
      message += ' | SyncStatus: ' + syncStatus;
    }


  } catch (err) {
    message = String(err.message || err);
  }


  // Zapis końcowego statusu pod blokadą.
  const finishLock = LockService.getScriptLock();


  if (!finishLock.tryLock(25000)) {
    console.error(
      'Nie zapisano statusu końcowego: ' + item.batchId
    );


    return {
      batchId: item.batchId,
      status: 'STATUS_UPDATE_FAILED'
    };
  }


  try {
    const q = SS.getSheetByName(IQ.QUEUE);


    const currentBatchId = String(
      q.getRange(item.row, 1).getValue()
    );


    const currentStatus = String(
      q.getRange(item.row, 3).getValue()
    );


    if (
      currentBatchId !== item.batchId ||
      currentStatus !== 'PROCESSING'
    ) {
      console.error(
        'Stan kolejki zmienił się podczas zapisu: ' +
        item.batchId
      );


      return {
        batchId: item.batchId,
        status: 'STATE_CONFLICT'
      };
    }


    q.getRange(item.row, 3).setValue(status);
    q.getRange(item.row, 6).setValue(new Date());
    q.getRange(item.row, 9).setValue(message);


    SpreadsheetApp.flush();


  } finally {
    if (finishLock.hasLock()) {
      finishLock.releaseLock();
    }
  }


  console.log(
    item.batchId + ' | ' + status + ' | ' + message
  );


  return {
    batchId: item.batchId,
    status,
    message
  };
}


/** Jednorazowo uruchomić ręcznie, aby włączyć pracownika co minutę. */
function installInventoryQueueWorker() {
  const existing = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processInventoryQueue');
  if (!existing.length) {
    ScriptApp.newTrigger('processInventoryQueue').timeBased()
      .everyMinutes(1).create();
  }
  inventoryQueueSheets_();
  console.log('Pracownik kolejki aktywny.');
}


/** Wyłącza tylko wyzwalacze pracownika kolejki. */
function stopInventoryQueueWorker() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processInventoryQueue')
    .forEach(t => ScriptApp.deleteTrigger(t));
}


/** Uruchamia 15 niezależnych zgłoszeń testowych, bez zmiany saveAllTasks(). */
function start15ScannersQueueTest() {
  const runId = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  let log = SS.getSheetByName(IQ.TEST);
  if (!log) {
    log = SS.insertSheet(IQ.TEST);
    log.appendRow(['RunId', 'Team', 'BatchId', 'Status',
      'Tasks', 'StartedAt', 'FinishedAt', 'Seconds', 'Message']);
  }
  for (let i = 1; i <= 15; i++) {
    const team = 'Ekipa_' + i;
    const trigger = ScriptApp.newTrigger('runOneScannerQueueTest')
      .timeBased().after(60000).create();
    props.setProperty('IQ_TEST_' + trigger.getUniqueId(),
      JSON.stringify({ runId, team, batchId: runId + '-' + team }));
  }
  console.log('Utworzono 15 zgłoszeń testowych. RunId: ' + runId);
}


function runOneScannerQueueTest(e) {
  const props = PropertiesService.getScriptProperties();
  const uid = String(e.triggerUid);
  const key = 'IQ_TEST_' + uid;
  const raw = props.getProperty(key);


  if (!raw) {
    throw new Error('Brak konfiguracji testu');
  }


  const cfg = JSON.parse(raw);
  const started = new Date();


  let status = 'ERROR';
  let count = 0;
  let message = '';
  let attempts = 0;


  try {
    const data = getInitialDataForLoop(cfg.team);


    const tasks = data.tasks.map(t => ({
      loc: t.loc,
      sku: t.sku,
      masterRow: t.masterRow,
      qty: Number(t.sysQty) || 0,
      comment: 'LOAD_TEST_QUEUE'
    }));


    count = tasks.length;


    if (!count) {
      status = 'EMPTY';
    } else {
      const maxAttempts = 4;


      for (let i = 1; i <= maxAttempts; i++) {
        attempts = i;


        // Ten sam BatchId przy każdej próbie.
        const response = enqueueInventory(
          cfg.team,
          tasks,
          data.startTime,
          cfg.batchId
        );


        if (response.success) {
          status = response.status;
          message = response.alreadyQueued
            ? 'Zgłoszenie już istnieje'
            : '';
          break;
        }


        message = String(response.msg || 'Nieznany błąd');


        // Ponawiamy wyłącznie przejściowy problem blokady.
        const retryable = response.retryable === true &&
          message.indexOf('LOCK_BUSY') !== -1;


        if (!retryable || i === maxAttempts) {
          status = 'ERROR';
          break;
        }


        Utilities.sleep(3000 + i * 2000);
      }
    }
  } catch (err) {
    status = 'ERROR';
    message = String(err.message || err);
  } finally {
    const finished = new Date();


    // Usuwamy wyłącznie wyzwalacz bieżącego testu.
    ScriptApp.getProjectTriggers()
      .filter(t => t.getUniqueId() === uid)
      .forEach(t => ScriptApp.deleteTrigger(t));


    props.deleteProperty(key);


    // Osobna blokada dziennika, niezależna od ScriptLock.
    const logLock = LockService.getDocumentLock();


    let logLocked = false;


    try {
      if (logLock) {
        logLocked = logLock.tryLock(30000);
      }


      if (!logLocked) {
        console.error(
          'Nie udało się uzyskać blokady dziennika: ' + cfg.team
        );
      } else {
        const log = SS.getSheetByName(IQ.TEST);


        log.appendRow([
          cfg.runId,
          cfg.team,
          cfg.batchId,
          status,
          count,
          started,
          finished,
          (finished - started) / 1000,
          'Próby: ' + attempts +
            (message ? ' | ' + message : '')
        ]);


        SpreadsheetApp.flush();
      }
    } finally {
      if (logLocked && logLock.hasLock()) {
        logLock.releaseLock();
      }
    }


    console.log(
      cfg.team +
      ' | ' + status +
      ' | zadań: ' + count +
      ' | prób: ' + attempts +
      ' | ' + message
    );
  }
}
/**
 * Odczytuje aktualny status partii z dziennika synchronizacji.
 */
function iqGetSyncStatus_(batchId) {
  const sheet = SS.getSheetByName('Inventory_Sync_Log');


  if (!sheet || sheet.getLastRow() < 2) {
    return 'NOT_FOUND';
  }


  const item = findInventoryBatch(sheet, batchId);


  return item ? String(item.status) : 'NOT_FOUND';
}


/**
 * Odzyskuje przerwane zadania kolejki.
 * Nie ponawia partii o niepewnym stanie STARTED.
 */
function recoverInventoryQueue() {
  const lock = LockService.getScriptLock();


  if (!lock.tryLock(1000)) {
    console.log('Odzyskiwanie pominięte: blokada zajęta');
    return;
  }


  try {
    const queue = SS.getSheetByName(IQ.QUEUE);


    if (!queue || queue.getLastRow() < 2) {
      return;
    }


    const rows = queue.getRange(
      2, 1, queue.getLastRow() - 1, 9
    ).getValues();


    const now = Date.now();


    for (let i = 0; i < rows.length; i++) {
      const row = i + 2;
      const batchId = String(rows[i][0]);
      const queueStatus = String(rows[i][2]);


      if (queueStatus !== 'PROCESSING') {
        continue;
      }


      const startedAt = new Date(rows[i][4]).getTime();


      // Nie ingerujemy w zadania, które mogą nadal działać.
      if (
        Number.isFinite(startedAt) &&
        now - startedAt < IQ.STALE_MS
      ) {
        continue;
      }


      const syncStatus = iqGetSyncStatus_(batchId);


      let newStatus = 'NEEDS_REVIEW';
      let message = '';


      if (syncStatus === 'DONE') {
        newStatus = 'DONE';
        message = 'Odzyskano: zapis wcześniej zakończony';


      } else if (
        syncStatus === 'REPORT_SAVED' ||
        syncStatus === 'TASKS_ASSIGNED'
      ) {
        // Istniejąca saveAllTasks() obsługuje te stany.
        // Ponowienie jest dozwolone tylko po potwierdzeniu
        // zakończenia poprzedniego wykonania.
        newStatus = 'RECOVERY_PENDING';
        message = 'Oczekuje na bezpieczne wznowienie: ' +
          syncStatus;


      } else if (syncStatus === 'STARTED') {
        message = 'Możliwy częściowy zapis. Wymagana kontrola';


      } else {
        message = 'Niepotwierdzony stan synchronizacji: ' +
          syncStatus;
      }


      queue.getRange(row, 3).setValue(newStatus);
      queue.getRange(row, 9).setValue(message);


      if (newStatus === 'DONE') {
        queue.getRange(row, 6).setValue(new Date());
      }


      console.log(
        batchId + ' | ' + syncStatus + ' -> ' + newStatus
      );
    }


    SpreadsheetApp.flush();


  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


/**
 * Instaluje cykliczną kontrolę przerwanych zadań.
 */
function installInventoryQueueRecovery() {
  const triggers = ScriptApp.getProjectTriggers();


  const exists = triggers.some(t =>
    t.getHandlerFunction() === 'recoverInventoryQueue'
  );


  if (!exists) {
    ScriptApp.newTrigger('recoverInventoryQueue')
      .timeBased()
      .everyMinutes(5)
      .create();
  }


  console.log('Kontrola odzyskiwania aktywna');
}
/**
 * Sprawdza, czy poprzedni zapis nadal może być aktywny.
 * Wymaga czasu rozpoczęcia starszego niż okres bezpieczeństwa.
 */
function iqCanRecover_(startedAt) {
  const timestamp = new Date(startedAt).getTime();


  if (!Number.isFinite(timestamp)) {
    return false;
  }


  return Date.now() - timestamp >= IQ.STALE_MS;
}


/**
 * Przywraca paczkę wyłącznie ze stanów obsługiwanych
 * przez istniejącą funkcję saveAllTasks().
 */
function iqResumeBatch_(item) {
  const syncStatus = iqGetSyncStatus_(item.batchId);


  if (syncStatus === 'DONE') {
    return { success: true, alreadySaved: true };
  }


  if (
    syncStatus !== 'REPORT_SAVED' &&
    syncStatus !== 'TASKS_ASSIGNED'
  ) {
    return {
      success: false,
      msg: 'Niebezpieczny stan synchronizacji: ' + syncStatus
    };
  }


  const dataSheet = SS.getSheetByName(IQ.DATA);


  const payload = inventoryQueueReadPayload_(
    dataSheet,
    item.batchId,
    item.chunks
  );


  return saveAllTasks(
    payload.teamName,
    payload.allTasks,
    payload.startTime,
    payload.batchId
  );
}
function processInventoryRecovery() {
  const queue = SS.getSheetByName(IQ.QUEUE);


  if (!queue || queue.getLastRow() < 2) {
    return;
  }


  const lock = LockService.getScriptLock();


  if (!lock.tryLock(1000)) {
    return;
  }


  let item = null;


  try {
    const rows = queue.getRange(
      2, 1, queue.getLastRow() - 1, 9
    ).getValues();


    // Nie uruchamiamy odzyskiwania, gdy trwa zwykłe przetwarzanie.
    if (rows.some(r => r[2] === 'PROCESSING')) {
      return;
    }


    const index = rows.findIndex(r =>
      r[2] === 'RECOVERY_PENDING' &&
      iqCanRecover_(r[4])
    );


    if (index < 0) {
      return;
    }


    item = {
      row: index + 2,
      batchId: String(rows[index][0]),
      team: String(rows[index][1]),
      chunks: Number(rows[index][7])
    };


    queue.getRange(item.row, 3).setValue('PROCESSING');
    queue.getRange(item.row, 5).setValue(new Date());
    queue.getRange(item.row, 9)
      .setValue('Automatyczne wznowienie zapisu');


    SpreadsheetApp.flush();


  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }


  if (!item) {
    return;
  }


  let result;
  let status = 'NEEDS_REVIEW';
  let message = '';


  try {
    result = iqResumeBatch_(item);


    if (result && result.success) {
      status = 'DONE';
      message = 'Odzyskano automatycznie';
    } else {
      message = String(
        result && result.msg || 'Nie udało się wznowić'
      );
    }
  } catch (err) {
    message = String(err.message || err);
  }


  queue.getRange(item.row, 3).setValue(status);
  queue.getRange(item.row, 6).setValue(new Date());
  queue.getRange(item.row, 9).setValue(message);


  SpreadsheetApp.flush();


  console.log(
    item.batchId + ' | ' + status + ' | ' + message
  );
}
function installInventoryRecoveryWorker() {
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'processInventoryRecovery'
  );


  if (!exists) {
    ScriptApp.newTrigger('processInventoryRecovery')
      .timeBased()
      .everyMinutes(1)
      .create();
  }


  console.log('Automatyczne odzyskiwanie aktywne');
}
function stopOldInventoryWorkers() {
  const handlers = [
    'processInventoryQueue',
    'recoverInventoryQueue',
    'processInventoryRecovery'
  ];


  ScriptApp.getProjectTriggers()
    .filter(t => handlers.includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));


  console.log('Zatrzymano stare wyzwalacze kolejki.');
}
