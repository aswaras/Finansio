// ===============================================
// GOOGLE APPS SCRIPT - Finanzo Backend
// ===============================================

// Get atau Create spreadsheet
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss;
}

function getOrCreateSheetByName(name) {
  const ss = getOrCreateSheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ===============================================
// MAIN GET REQUEST - Return all data
// ===============================================
function doGet(e) {
  try {
    const ss = getOrCreateSheet();
    
    // Get all data from sheets
    const userSheet = getOrCreateSheetByName("USER");
    const recordsSheet = getOrCreateSheetByName("RECORDS");
    const notasSheet = getOrCreateSheetByName("NOTAS");
    const notaItemsSheet = getOrCreateSheetByName("NOTA_ITEMS");

    // Get user data
    const userData = getUserData(userSheet);
    
    // Get records (catat transaksi)
    const records = getRecords(recordsSheet);
    
    // Get notas with items
    const notas = getNotasWithItems(notasSheet, notaItemsSheet);

    const response = {
      status: "success",
      data: {
        user: userData,
        theme: "light",
        activeView: "home",
        records: records,
        notas: notas
      }
    };

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error in doGet: " + error);
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// MAIN POST REQUEST - Save all data
// ===============================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    const ss = getOrCreateSheet();
    const userSheet = getOrCreateSheetByName("USER");
    const recordsSheet = getOrCreateSheetByName("RECORDS");
    const notasSheet = getOrCreateSheetByName("NOTAS");
    const notaItemsSheet = getOrCreateSheetByName("NOTA_ITEM");

    // Save user data
    saveUserData(userSheet, data.user);
    
    // Save records (catat)
    saveRecords(recordsSheet, data.records);
    
    // Save notas and items
    saveNotasWithItems(notasSheet, notaItemsSheet, data.notas);

    const response = {
      status: "success",
      message: "Data berhasil disimpan"
    };

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error in doPost: " + error);
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===============================================
// USER DATA FUNCTIONS
// ===============================================
function getUserData(sheet) {
  const headers = ["ID", "Name", "Shop", "Address", "Footer", "Color"];
  initializeSheet(sheet, headers);

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  
  if (data.length === 0 || !data[0][0]) {
    return {
      name: "Pemilik",
      shop: "Toko Finanzo",
      address: "Jl. Sukses No. 1",
      footer: "Terima kasih.\nTransfer: BCA 1234567",
      color: "#2563eb"
    };
  }

  return {
    name: data[0][1] || "Pemilik",
    shop: data[0][2] || "Toko Finanzo",
    address: data[0][3] || "Jl. Sukses No. 1",
    footer: data[0][4] || "Terima kasih.",
    color: data[0][5] || "#2563eb"
  };
}

function saveUserData(sheet, userData) {
  const headers = ["ID", "Name", "Shop", "Address", "Footer", "Color"];
  initializeSheet(sheet, headers);

  // Clear all data rows
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  // Add user data
  sheet.appendRow([
    "USER_001",
    userData.name || "Pemilik",
    userData.shop || "Toko Finanzo",
    userData.address || "Jl. Sukses No. 1",
    userData.footer || "Terima kasih.",
    userData.color || "#2563eb"
  ]);
}

// ===============================================
// RECORDS (CATAT) FUNCTIONS
// ===============================================
function getRecords(sheet) {
  const headers = ["ID", "Date", "Type", "Amount", "Note"];
  initializeSheet(sheet, headers);

  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), headers.length).getValues();
  
  const records = [];
  data.forEach(row => {
    if (row[0]) { // if ID exists
      records.push({
        id: row[0],
        date: row[1] || "",
        type: row[2] || "",
        amount: parseFloat(row[3]) || 0,
        note: row[4] || ""
      });
    }
  });

  return records;
}

function saveRecords(sheet, records) {
  const headers = ["ID", "Date", "Type", "Amount", "Note"];
  initializeSheet(sheet, headers);

  // Clear all data
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  // Add records
  records.forEach(record => {
    sheet.appendRow([
      record.id,
      record.date,
      record.type,
      record.amount,
      record.note
    ]);
  });
}

// ===============================================
// NOTAS FUNCTIONS
// ===============================================
function getNotasWithItems(notasSheet, itemsSheet) {
  const notaHeaders = ["ID", "Customer", "Phone", "Date", "Deadline", "DP", "Total"];
  const itemHeaders = ["NotaID", "ItemName", "Description", "Qty", "Price"];
  
  initializeSheet(notasSheet, notaHeaders);
  initializeSheet(itemsSheet, itemHeaders);

  const notasData = notasSheet.getRange(2, 1, Math.max(notasSheet.getLastRow() - 1, 0), notaHeaders.length).getValues();
  const itemsData = itemsSheet.getRange(2, 1, Math.max(itemsSheet.getLastRow() - 1, 0), itemHeaders.length).getValues();

  const notas = [];
  const itemsMap = {};

  // Build items map
  itemsData.forEach(row => {
    if (row[0]) { // if NotaID exists
      if (!itemsMap[row[0]]) {
        itemsMap[row[0]] = [];
      }
      itemsMap[row[0]].push({
        name: row[1] || "",
        description: row[2] || "",
        qty: parseFloat(row[3]) || 0,
        price: parseFloat(row[4]) || 0
      });
    }
  });

  // Build notas with items
  notasData.forEach(row => {
    if (row[0]) { // if ID exists
      notas.push({
        id: row[0],
        customer: row[1] || "",
        phone: row[2] || "",
        date: row[3] || "",
        deadline: row[4] || "",
        dp: parseFloat(row[5]) || 0,
        total: parseFloat(row[6]) || 0,
        items: itemsMap[row[0]] || []
      });
    }
  });

  return notas;
}

function saveNotasWithItems(notasSheet, itemsSheet, notas) {
  const notaHeaders = ["ID", "Customer", "Phone", "Date", "Deadline", "DP", "Total"];
  const itemHeaders = ["NotaID", "ItemName", "Description", "Qty", "Price"];
  
  initializeSheet(notasSheet, notaHeaders);
  initializeSheet(itemsSheet, itemHeaders);

  // Clear notas sheet
  if (notasSheet.getLastRow() > 1) {
    notasSheet.deleteRows(2, notasSheet.getLastRow() - 1);
  }

  // Clear items sheet
  if (itemsSheet.getLastRow() > 1) {
    itemsSheet.deleteRows(2, itemsSheet.getLastRow() - 1);
  }

  // Add notas and items
  notas.forEach(nota => {
    notasSheet.appendRow([
      nota.id,
      nota.customer,
      nota.phone,
      nota.date,
      nota.deadline,
      nota.dp,
      nota.total
    ]);

    // Add items for this nota
    nota.items.forEach(item => {
      itemsSheet.appendRow([
        nota.id,
        item.name,
        item.description || "",
        item.qty,
        item.price
      ]);
    });
  });
}

// ===============================================
// HELPER FUNCTIONS
// ===============================================
function initializeSheet(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#2563eb");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
  }
}
