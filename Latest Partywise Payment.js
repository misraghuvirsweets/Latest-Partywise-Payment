function doGet(e) {
  var role = (e && e.parameter && e.parameter.role) ? e.parameter.role.toLowerCase() : 'counter';
  var branch = (e && e.parameter && e.parameter.branch) ? e.parameter.branch : 'Ram Mandir';
  if(role === 'savali') branch = 'Savali';
  var tmp = HtmlService.createTemplateFromFile('Index');
  tmp.role = role;
  tmp.urlBranch = branch; 
  return tmp.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Raghuvir Portal").addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function markAsVerified(branch, rowIndex) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(branch);
    sheet.getRange(rowIndex, 6).setValue("Verified");
    return "SUCCESS";
  } catch(e) { return e.toString(); }
}

function markAsPaid(branch, rowIndex, roleName, p) {
  try {
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
}

function markMultipleAsPaid(items, roleName, p) {
  try {
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
}

function getParties() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  var data = sheet ? sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat().filter(n => n && n !== "Party Name" && n !== "Other Expenses") : [];
  data.sort(); data.push("Other Expenses"); return data;
}

function addNewParty(name) {
  try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master").appendRow([name]); return "SUCCESS"; } catch(e) { return e.toString(); }
}

function submitForm(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(data.branch);
    var methodToSave = data.method;
    var amountToSave = Number(data.amount);
    if (data.method === "Credit Note") { methodToSave = "Credit"; amountToSave = -Math.abs(amountToSave); }
    var folderName = "Invoice_Uploads";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    var urlList = [];
    if (data.files && data.files.length > 0) {
      data.files.forEach(function(f, idx) {
        var base64Data = f.base64.split(',')[1];
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", "B_" + data.bill + "_" + idx + ".jpg");
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        urlList.push(file.getUrl());
      });
    }
    sheet.appendRow([new Date(data.date), data.party, data.bill, amountToSave, methodToSave, "Unverified", urlList.join(", "), data.chequeNo || "", "", "", data.role.toUpperCase(), data.noteRemark || ""]);
    return {status: "SUCCESS"};
  } catch(e) { return {status: "ERROR", msg: e.toString()}; }
}

function getVendorHistory(branch, party, role) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var branchesToSearch = (role && (role.includes('accountant') || role === 'admin' || role === 'ea')) ? ["Ram Mandir", "Vishrambagh", "Miraj", "Savali"] : [branch];
  var history = [];
  var searchParty = String(party).trim().toLowerCase();
  branchesToSearch.forEach(function(bName) {
    var sh = ss.getSheetByName(bName);
    if (sh && sh.getLastRow() > 1) {
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
      data.forEach(function(r) {
        if (String(r[1]).trim().toLowerCase() === searchParty) {
          var m = String(r[4]), a = r[3];
          if (m === "Credit Note") { m = "Credit"; a = -Math.abs(a); }
          history.push({ date: Utilities.formatDate(new Date(r[0]), "GMT+5:30", "dd/MM/yyyy"), rawDate: new Date(r[0]).getTime(), party: String(r[1]), bill: String(r[2]), amt: a, method: m, link: String(r[6]), branch: bName, verifyStatus: r[5] });
        }
      });
    }
  });
  return history.sort((a, b) => b.rawDate - a.rawDate).slice(0, 15);
}

function getDashboardData(start, end, branchFilter, statusFilter, verifyFilter, partyFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var branches = ["Ram Mandir", "Vishrambagh", "Miraj", "Savali"];
  var allData = [];
  var sD = start ? new Date(start).setHours(0,0,0,0) : null;
  var eD = end ? new Date(end).setHours(23,59,59,999) : null;
  var pF = (partyFilter && partyFilter !== "") ? partyFilter.trim().toLowerCase() : null;

  branches.forEach(function(bName) {
    if (branchFilter !== "All" && branchFilter !== bName) return;
    var sh = ss.getSheetByName(bName);
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
      rows.forEach(function(r, idx) {
        if (pF && String(r[1]).trim().toLowerCase().indexOf(pF) === -1 && String(r[2]).trim().toLowerCase().indexOf(pF) === -1) return;
        var rDate = new Date(r[0]).getTime();
        if (sD && rDate < sD) return;
        if (eD && rDate > eD) return;
        var m = String(r[4]).trim(), a = r[3];
        if (m === "Credit Note") { m = "Credit"; a = -Math.abs(a); }
        if (statusFilter !== "All" && m !== statusFilter) return;
        var vStat = String(r[5] || "Unverified").trim();
        if (verifyFilter !== "All" && vStat !== verifyFilter) return;
        allData.push({ rowIndex: idx + 2, date: Utilities.formatDate(new Date(r[0]), "GMT+5:30", "dd/MM/yyyy"), party: r[1], bill: r[2], amt: a, method: m, verifyStatus: vStat, link: r[6], chequeNo: r[7], rawDate: rDate, branch: bName });
      });
    }
  });
  return allData.sort((a,b) => b.rawDate - a.rawDate);
}