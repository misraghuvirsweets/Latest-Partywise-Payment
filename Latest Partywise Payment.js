function doGet(e) {
  var role = (e && e.parameter && e.parameter.role) ? e.parameter.role.toLowerCase() : 'counter';
  var branch = (e && e.parameter && e.parameter.branch) ? e.parameter.branch : 'Ram Mandir';
  if(role === 'savali') branch = 'Savali';
  
  var tmp = HtmlService.createTemplateFromFile('Index');
  tmp.role = role;
  tmp.urlBranch = branch; 
  
  return tmp.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Raghuvir Portal")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function markAsVerified(branch, rowIndex) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(branch);
    sheet.getRange(rowIndex, 6).setValue("Verified");
    return "SUCCESS";
  } catch(e) { return "Error: " + e.toString(); }
  finally { lock.releaseLock(); }
}

function markAsPaid(branch, rowIndex, roleName, p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(branch);
    var ts = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm:ss");
    var bankPart = p.bank ? " | Bank: " + p.bank : "";
    var details = "PAID: " + p.mode + bankPart + " | TxDate: " + p.txDate + " | Details: " + p.remark;
    sheet.getRange(rowIndex, 5).setValue("Paid");
    sheet.getRange(rowIndex, 9).setValue(ts);
    sheet.getRange(rowIndex, 10).setValue(roleName.toUpperCase());
    sheet.getRange(rowIndex, 12).setValue(details);
    return "SUCCESS";
  } catch(e) { return e.toString(); }
  finally { lock.releaseLock(); }
}

function markChequeIssuedStatus(branch, rowIndex) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(branch);
    if (!rowIndex) return;
    sheet.getRange(rowIndex, 5).setValue("Cheque Issued");
    return "SUCCESS";
  } catch(e) { return e.toString(); }
  finally { lock.releaseLock(); }
}

function markMultipleAsPaid(items, roleName, p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ts = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm:ss");
    var bankPart = p.bank ? " | Bank: " + p.bank : "";
    var details = "BULK PAID: " + p.mode + bankPart + " | TxDate: " + p.txDate + " | Details: " + p.remark;
    items.forEach(function(item) {
      var sheet = ss.getSheetByName(item.branch);
      sheet.getRange(item.rowIndex, 5).setValue("Paid");
      sheet.getRange(item.rowIndex, 9).setValue(ts);
      sheet.getRange(item.rowIndex, 10).setValue(roleName.toUpperCase());
      sheet.getRange(item.rowIndex, 12).setValue(details);
    });
    return "SUCCESS";
  } catch(e) { return e.toString(); }
  finally { lock.releaseLock(); }
}

function getParties() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
    var data = sheet ? sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat().filter(n => n && n !== "Party Name" && n !== "Other Expenses") : [];
    data.sort(); 
    data.push("Other Expenses"); 
    return data;
  } catch(e) { return ["Other Expenses"]; }
}

function addNewParty(name) {
  var lock = LockService.getScriptLock();
  try { 
    lock.waitLock(10000);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").appendRow([name]); 
    return "SUCCESS"; 
  } catch(e) { return e.toString(); }
  finally { lock.releaseLock(); }
}

function getOrCreateInvoiceFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('INVOICE_FOLDER_ID');
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {}
  }
  var folderName = "Invoice_Uploads";
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty('INVOICE_FOLDER_ID', folder.getId());
  return folder;
}

function submitForm(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(data.branch);

    if (sheet.getLastRow() > 1) {
      var lastR = sheet.getLastRow();
      var existingEntries = sheet.getRange(1, 2, lastR, 2).getValues(); 
      var newParty = String(data.party).trim().toLowerCase();
      var newBill = String(data.bill).trim().toLowerCase();
      
      for (var i = 0; i < existingEntries.length; i++) {
        if (String(existingEntries[i][0]).trim().toLowerCase() === newParty && 
            String(existingEntries[i][1]).trim().toLowerCase() === newBill) {
          return {status: "DUPLICATE", msg: "An entry with Party: " + data.party + " and Bill No: " + data.bill + " already exists in " + data.branch + "!"};
        }
      }
    }

    var methodToSave = data.method;
    var amountToSave = Number(data.amount);
    if (data.method === "Credit Note") { 
      methodToSave = "Credit"; 
      amountToSave = -Math.abs(amountToSave); 
    }
    
    var folder = getOrCreateInvoiceFolder();
    var urlList = [];
    
    if (data.files && data.files.length > 0) {
      data.files.forEach(function(f, idx) {
        var base64Data = f.base64.split(',')[1];
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", "B_" + data.bill + "_" + idx + ".jpg");
        var file = folder.createFile(blob);
        urlList.push(file.getUrl());
      });
    }

    var formTimestamp = Utilities.formatDate(new Date(), "GMT+5:30", "dd/MM/yyyy HH:mm:ss");
    
    sheet.appendRow([
      new Date(data.date), 
      data.party, 
      data.bill, 
      amountToSave, 
      methodToSave, 
      "Unverified", 
      urlList.join(", "), 
      data.chequeNo || "", 
      "", 
      "", 
      data.role.toUpperCase(), 
      data.noteRemark || "",
      formTimestamp 
    ]);
    return {status: "SUCCESS"};
  } catch(e) { 
    return {status: "ERROR", msg: e.toString()}; 
  } finally {
    lock.releaseLock();
  }
}

function getVendorHistory(branch, party, role) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var branchesToSearch = (role && (role.includes('accountant') || role === 'admin' || role === 'ea')) ? ["Ram Mandir", "Vishrambagh", "Miraj", "Savali"] : [branch];
  var history = [];
  var searchParty = String(party).trim().toLowerCase();
  
  branchesToSearch.forEach(function(bName) {
    var sh = ss.getSheetByName(bName);
    if (sh && sh.getLastRow() > 1) {
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      data.forEach(function(r, idx) {
        if (!r[1]) return;
        if (String(r[1]).trim().toLowerCase() === searchParty) {
          var m = String(r[4]), a = r[3];
          if (m === "Credit Note") { m = "Credit"; a = -Math.abs(a); }
          history.push({ 
            rowIndex: idx + 2, 
            date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+5:30", "dd/MM/yyyy") : String(r[0]), 
            rawDate: r[0] instanceof Date ? r[0].getTime() : 0, 
            party: String(r[1]), 
            bill: String(r[2]), 
            amt: a, 
            method: m, 
            link: String(r[6] || ""), 
            branch: bName, 
            verifyStatus: r[5] 
          });
        }
      });
    }
  });
  return history.sort((a, b) => b.rawDate - a.rawDate).slice(0, 15);
}

function getRecentBranchTransactions(branch) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(branch);
  if (!sh || sh.getLastRow() < 2) return [];
  
  var lastRow = sh.getLastRow();
  var data = sh.getRange(2, 1, lastRow - 1, Math.min(sh.getLastColumn(), 7)).getValues();
  
  var results = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var r = data[i];
    if (!r[0] || !r[1]) continue; // Skip truly empty rows
    
    var m = String(r[4]), a = r[3];
    if (m === "Credit Note") { m = "Credit"; a = -Math.abs(a); }
    
    results.push({ 
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+5:30", "dd/MM/yyyy") : String(r[0]), 
      party: String(r[1]), 
      bill: String(r[2]), 
      amt: a, 
      method: m, 
      branch: branch, 
      link: String(r[6] || "") 
    });
    if (results.length >= 15) break;
  }
  return results;
}

function getDashboardData(start, end, branchFilter, statusFilter, verifyFilter, partyFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var branches = ["Ram Mandir", "Vishrambagh", "Miraj", "Savali"];
  var allData = [];
  var sD = (start && start !== "") ? new Date(start).setHours(0,0,0,0) : null;
  var eD = (end   && end   !== "") ? new Date(end).setHours(23,59,59,999) : null;
  var pF = (partyFilter && partyFilter !== "") ? String(partyFilter).trim().toLowerCase() : null;
  var sf = (statusFilter && statusFilter !== "") ? String(statusFilter).trim() : "All";
  var vf = (verifyFilter && verifyFilter !== "") ? String(verifyFilter).trim() : "All";

  branches.forEach(function(bName) {
    if (branchFilter !== "All" && branchFilter !== bName) return;
    var sh = ss.getSheetByName(bName);
    if (!sh || sh.getLastRow() < 2) return;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(sh.getLastColumn(), 13)).getValues();
    rows.forEach(function(r, idx) {
      if (!r[0] || !r[1]) return;
      if (pF) {
        var partyStr = String(r[1]).trim().toLowerCase();
        var billStr  = String(r[2]).trim().toLowerCase();
        if (partyStr.indexOf(pF) === -1 && billStr.indexOf(pF) === -1) return;
      }
      var rDate = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
      if (isNaN(rDate)) return;
      if (sD && rDate < sD) return;
      if (eD && rDate > eD) return;
      var m = String(r[4]).trim();
      var a = r[3];
      if (m === "Credit Note") { m = "Credit"; a = -Math.abs(Number(a)); }
      if (sf !== "All" && m !== sf) return;
      var vStat = String(r[5] || "Unverified").trim();
      if (vf !== "All" && vStat !== vf) return;
      allData.push({ 
        rowIndex: idx + 2, 
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], "GMT+5:30", "dd/MM/yyyy") : String(r[0]), 
        party: String(r[1]), 
        bill: String(r[2]), 
        amt: a, 
        method: m, 
        verifyStatus: vStat, 
        link: String(r[6] || ""), 
        chequeNo: String(r[7] || ""), 
        rawDate: rDate, 
        branch: bName 
      });
    });
  });
  return allData.sort(function(a, b) { return b.rawDate - a.rawDate; });
}
