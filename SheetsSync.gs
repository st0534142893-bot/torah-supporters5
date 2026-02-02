/**
 * קוד Google Apps Script לסנכרון דו-כיווני עם תוכנת קמפיין פורים
 * כולל חיבור מאובטח לנדרים פלוס
 * 
 * ========== הוראות התקנה ==========
 * 
 * 1. פתח את הגליון האלקטרוני שלך בגוגל
 * 2. לחץ על "הרחבות" (Extensions) -> "Apps Script"
 * 3. מחק את כל הקוד שיש שם והדבק את הקוד הזה במקום
 * 4. שנה את SPREADSHEET_ID למזהה של הגליון שלך (מופיע ב-URL)
 * 5. שנה את הגדרות NEDARIM_CONFIG למספרים שלך
 * 6. לחץ על "פריסה" (Deploy) -> "פריסה חדשה" (New deployment)
 * 7. בחר סוג: "אפליקציית אינטרנט"
 * 8. הגדר "הפעל כ": אני (Me) | "למי יש גישה": כל אחד (Anyone)
 * 9. לחץ "פרוס" והעתק את ה-URL שמתקבל
 * 
 * =====================================
 */

// ⚠️ שנה את המזהה הזה למזהה של הגליון שלך!
const SPREADSHEET_ID = '1YI6XQZObSP1vfhIVh9wXYtIufM20PFjIGGM9rB-_rc8';

// *****************************************************************************
// *** הגדרות נדרים פלוס - סודי! ***
// *****************************************************************************
const NEDARIM_CONFIG = {
  MOSAD_ID: '1000642',
  MATCHING_ID: '715',
  API_PASSWORD: 'ep348',
  API_URL: 'https://matara.pro/nedarimplus/Reports/Manage3.aspx',
  ONLINE_API_URL: 'https://www.matara.pro/nedarimplus/online/Files/Manage.aspx',
  MATCH_PLUS_API_URL: 'https://www.matara.pro/nedarimplus/V6/MatchPlus.aspx'
};

// שמות הגיליונות
const SHEET_NAMES = {
  DONORS: 'מתרימים',
  GROUPS: 'קבוצות',
  SETTINGS: 'הגדרות',
  HISTORY: 'היסטוריה',
  TOOLKIT: 'ארגז_כלים',
  ADDRESSES: 'כתובות',
  SCOUTS: 'סיירות',
  GROOM_GRANTS: 'מענקי_חתנים',
  FINANCE: 'כספים'
};

// עמודות למתרימים
const DONOR_COLUMNS = ['id', 'name', 'displayName', 'groupId', 'amount', 'personalGoal', 'history', 'nedarimMatrimId', 'createdAt', 'updatedAt'];

// עמודות לקבוצות
const GROUP_COLUMNS = ['id', 'name', 'goal', 'orderNumber', 'createdAt', 'updatedAt'];

// עמודות להיסטוריה
const HISTORY_COLUMNS = ['id', 'timestamp', 'date', 'time', 'actionType', 'entityType', 'entityId', 'entityName', 'details', 'amount', 'source', 'computerName', 'computerId', 'userId'];

/**
 * פונקציה ראשית לטיפול בבקשות GET
 */
function doGet(e) {
  var output;
  
  try {
    var action = e && e.parameter ? e.parameter.action : '';
    
    switch(action) {
      // ===== פעולות נדרים פלוס =====
      case 'getNedarimTotal':
        output = getNedarimTotalDonations();
        break;
      case 'searchNedarimRecruiters':
        var searchTerm = e && e.parameter ? e.parameter.search : '';
        output = searchNedarimRecruiters(searchTerm);
        break;
      case 'getNedarimRecruiterDetails':
        var recruiterId = e && e.parameter ? e.parameter.recruiterId : '';
        output = getNedarimRecruiterDetails(recruiterId);
        break;
      case 'getPublicConfig':
        output = getPublicConfig();
        break;
        
      // ===== פעולות סנכרון רגילות =====
      case 'getDonors':
        output = getAllDonors();
        break;
      case 'getGroups':
        output = getAllGroups();
        break;
      case 'getSettings':
        output = getAllSettings();
        break;
      case 'getAll':
        output = getAllData();
        break;
      case 'getAllComplete':
        output = getAllDataComplete();
        break;
      case 'getExtended':
        output = getAllExtendedData();
        break;
      case 'getHistory':
        var filters = {};
        if (e && e.parameter) {
          if (e.parameter.startDate) filters.startDate = e.parameter.startDate;
          if (e.parameter.endDate) filters.endDate = e.parameter.endDate;
          if (e.parameter.source) filters.source = e.parameter.source;
          if (e.parameter.actionType) filters.actionType = e.parameter.actionType;
          if (e.parameter.entityType) filters.entityType = e.parameter.entityType;
          if (e.parameter.computerName) filters.computerName = e.parameter.computerName;
        }
        output = getHistory(Object.keys(filters).length > 0 ? filters : null);
        break;
      case 'ping':
        output = { success: true, message: 'החיבור תקין!', timestamp: new Date().toISOString() };
        break;
      case 'importSimple':
        var sheetName = e && e.parameter ? e.parameter.sheet : 'גיליון1';
        output = importFromSimpleSheet(sheetName);
        break;
      case 'getSimpleData':
        output = readSimpleSheet();
        break;
      default:
        output = getAllData();
    }
  } catch (error) {
    Logger.log('שגיאה ב-doGet: ' + error.toString());
    output = { success: false, error: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * פונקציה ראשית לטיפול בבקשות POST
 */
function doPost(e) {
  var output;
  
  try {
    var postData;
    if (e && e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      postData = e.parameter;
    } else {
      throw new Error('לא התקבלו נתונים');
    }
    
    var action = postData.action || '';
    
    switch(action) {
      // ===== פעולות נדרים פלוס =====
      case 'getNedarimReport':
        output = getNedarimReport(postData);
        break;
        
      // ===== פעולות סנכרון רגילות =====
      case 'saveDonors':
        output = saveDonors(postData.donors || []);
        break;
      case 'saveGroups':
        output = saveGroups(postData.groups || []);
        break;
      case 'saveSettings':
        output = saveSettings(postData.settings || {});
        break;
      case 'saveAll':
        output = saveAllData(postData);
        break;
      case 'saveAllComplete':
        output = saveAllDataComplete(postData);
        break;
      case 'saveExtended':
        output = saveAllExtendedData(postData);
        break;
      case 'syncAll':
        output = syncAllData(postData);
        break;
      case 'addHistory':
        output = addHistoryEntry(postData.entry || postData);
        break;
      case 'addHistoryBatch':
        output = addHistoryEntries(postData.entries || []);
        break;
      default:
        output = saveAllData(postData);
    }
  } catch (error) {
    Logger.log('שגיאה ב-doPost: ' + error.toString());
    output = { success: false, error: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// *****************************************************************************
// *** פונקציות נדרים פלוס ***
// *****************************************************************************

/**
 * קבלת ההגדרות הציבוריות (ללא סיסמאות)
 */
function getPublicConfig() {
  return {
    success: true,
    mosadId: NEDARIM_CONFIG.MOSAD_ID,
    matchingId: NEDARIM_CONFIG.MATCHING_ID
  };
}

/**
 * קבלת סכום תרומות כולל מנדרים פלוס
 */
function getNedarimTotalDonations() {
  try {
    // ניסיון 1: MatchPlus API עם ShowGoal (הכי מדויק)
    var url1 = NEDARIM_CONFIG.MATCH_PLUS_API_URL + 
      '?Action=ShowGoal&MosadId=' + NEDARIM_CONFIG.MOSAD_ID + 
      '&GoalId=' + NEDARIM_CONFIG.MATCHING_ID;
    
    Logger.log('מנסה URL: ' + url1);
    
    var response1 = UrlFetchApp.fetch(url1, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response1.getResponseCode() === 200) {
      var text = response1.getContentText();
      Logger.log('תגובה מ-ShowGoal: ' + text);
      try {
        var data = JSON.parse(text);
        // בודק שדות שונים שיכולים להכיל את הסכום
        var donated = parseFloat(data.Donated) || parseFloat(data.DSum) || parseFloat(data.TotalDonated) || 0;
        var goal = parseFloat(data.Goal) || parseFloat(data.TargetSum) || 0;
        
        if (donated > 0 || goal > 0) {
          return {
            success: true,
            totalDonated: donated,
            goal: goal,
            goalId: data.GoalId || NEDARIM_CONFIG.MATCHING_ID,
            source: 'MatchPlus-ShowGoal'
          };
        }
      } catch (e) {
        Logger.log('שגיאת פרסור ShowGoal: ' + e.message);
      }
    }
    
    // ניסיון 2: Reports API
    var url2 = NEDARIM_CONFIG.API_URL + 
      '?Action=GetMatchingDetails&MosadId=' + NEDARIM_CONFIG.MOSAD_ID + 
      '&MatchingId=' + NEDARIM_CONFIG.MATCHING_ID +
      '&Password=' + NEDARIM_CONFIG.API_PASSWORD;
    
    Logger.log('מנסה URL: ' + url2);
    
    var response2 = UrlFetchApp.fetch(url2, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response2.getResponseCode() === 200) {
      var text2 = response2.getContentText();
      Logger.log('תגובה מ-GetMatchingDetails: ' + text2);
      try {
        var data2 = JSON.parse(text2);
        var donated2 = parseFloat(data2.TotalDonated) || parseFloat(data2.Donated) || parseFloat(data2.DSum) || 0;
        var goal2 = parseFloat(data2.Goal) || parseFloat(data2.TargetSum) || 0;
        
        if (donated2 > 0 || goal2 > 0) {
          return {
            success: true,
            totalDonated: donated2,
            goal: goal2,
            source: 'Reports-GetMatchingDetails'
          };
        }
      } catch (e) {
        Logger.log('שגיאת פרסור GetMatchingDetails: ' + e.message);
      }
    }
    
    // ניסיון 3: חישוב מהמתרימים
    Logger.log('מנסה לחשב סכום מהמתרימים...');
    var recruitersResult = searchNedarimRecruiters('');
    if (recruitersResult && recruitersResult.success && recruitersResult.recruiters) {
      var total = 0;
      recruitersResult.recruiters.forEach(function(r) {
        total += parseFloat(r.Amount) || parseFloat(r.Collected) || 0;
      });
      if (total > 0) {
        return {
          success: true,
          totalDonated: total,
          goal: 0,
          source: 'Calculated-FromRecruiters',
          recruitersCount: recruitersResult.recruiters.length
        };
      }
    }
    
    return {
      success: false,
      error: 'לא ניתן לקבל נתונים מנדרים פלוס',
      totalDonated: 0,
      goal: 0
    };
    
  } catch (e) {
    Logger.log('שגיאה ב-getNedarimTotalDonations: ' + e.message);
    return {
      success: false,
      error: e.message,
      totalDonated: 0,
      goal: 0
    };
  }
}

/**
 * חיפוש מתרימים בנדרים פלוס
 */
function searchNedarimRecruiters(searchTerm) {
  try {
    var url = NEDARIM_CONFIG.MATCH_PLUS_API_URL + 
      '?Action=SearchMatrim&Name=' + encodeURIComponent(searchTerm || '') + 
      '&MosadId=' + NEDARIM_CONFIG.MOSAD_ID;
    
    var response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      return {
        success: true,
        recruiters: data || []
      };
    } else {
      return {
        success: false,
        error: 'שגיאת שרת: ' + response.getResponseCode(),
        recruiters: []
      };
    }
  } catch (e) {
    Logger.log('שגיאה ב-searchNedarimRecruiters: ' + e.message);
    return {
      success: false,
      error: e.message,
      recruiters: []
    };
  }
}

/**
 * קבלת פרטי מתרים מנדרים פלוס
 */
function getNedarimRecruiterDetails(recruiterId) {
  try {
    var url = NEDARIM_CONFIG.MATCH_PLUS_API_URL + 
      '?Action=GetMatrimData&MatrimId=' + encodeURIComponent(recruiterId) + 
      '&MosadId=' + NEDARIM_CONFIG.MOSAD_ID;
    
    var response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      return {
        success: true,
        data: data
      };
    } else {
      return {
        success: false,
        error: 'שגיאת שרת: ' + response.getResponseCode()
      };
    }
  } catch (e) {
    Logger.log('שגיאה ב-getNedarimRecruiterDetails: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * קבלת דוח מנדרים פלוס
 */
function getNedarimReport(params) {
  try {
    var formData = {
      'Action': params.reportAction || 'GetReport',
      'MosadNumber': NEDARIM_CONFIG.MOSAD_ID,
      'ApiPassword': NEDARIM_CONFIG.API_PASSWORD
    };
    
    if (params.fromDate) formData['FromDate'] = params.fromDate;
    if (params.toDate) formData['ToDate'] = params.toDate;
    
    var response = UrlFetchApp.fetch(NEDARIM_CONFIG.API_URL, {
      method: 'POST',
      payload: formData,
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      try {
        var data = JSON.parse(response.getContentText());
        return {
          success: true,
          data: data
        };
      } catch (e) {
        return {
          success: true,
          data: response.getContentText()
        };
      }
    } else {
      return {
        success: false,
        error: 'שגיאת שרת: ' + response.getResponseCode()
      };
    }
  } catch (e) {
    Logger.log('שגיאה ב-getNedarimReport: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

// *****************************************************************************
// *** פונקציות סנכרון Google Sheets ***
// *****************************************************************************

function ensureSheetsExist() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  var donorsSheet = ss.getSheetByName(SHEET_NAMES.DONORS);
  if (!donorsSheet) {
    donorsSheet = ss.insertSheet(SHEET_NAMES.DONORS);
    donorsSheet.getRange(1, 1, 1, DONOR_COLUMNS.length).setValues([DONOR_COLUMNS]);
    formatHeaderRow(donorsSheet);
  }
  
  var groupsSheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
  if (!groupsSheet) {
    groupsSheet = ss.insertSheet(SHEET_NAMES.GROUPS);
    groupsSheet.getRange(1, 1, 1, GROUP_COLUMNS.length).setValues([GROUP_COLUMNS]);
    formatHeaderRow(groupsSheet);
  }
  
  var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settingsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    formatHeaderRow(settingsSheet);
  }
  
  var historySheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_NAMES.HISTORY);
    historySheet.getRange(1, 1, 1, HISTORY_COLUMNS.length).setValues([HISTORY_COLUMNS]);
    formatHeaderRow(historySheet);
  }
  
  var toolkitSheet = ss.getSheetByName(SHEET_NAMES.TOOLKIT);
  if (!toolkitSheet) {
    toolkitSheet = ss.insertSheet(SHEET_NAMES.TOOLKIT);
    toolkitSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    formatHeaderRow(toolkitSheet);
  }
  
  var addressesSheet = ss.getSheetByName(SHEET_NAMES.ADDRESSES);
  if (!addressesSheet) {
    addressesSheet = ss.insertSheet(SHEET_NAMES.ADDRESSES);
    addressesSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    formatHeaderRow(addressesSheet);
  }
  
  var scoutsSheet = ss.getSheetByName(SHEET_NAMES.SCOUTS);
  if (!scoutsSheet) {
    scoutsSheet = ss.insertSheet(SHEET_NAMES.SCOUTS);
    scoutsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    formatHeaderRow(scoutsSheet);
  }
  
  var groomSheet = ss.getSheetByName(SHEET_NAMES.GROOM_GRANTS);
  if (!groomSheet) {
    groomSheet = ss.insertSheet(SHEET_NAMES.GROOM_GRANTS);
    groomSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    formatHeaderRow(groomSheet);
  }
  
  var financeSheet = ss.getSheetByName(SHEET_NAMES.FINANCE);
  if (!financeSheet) {
    financeSheet = ss.insertSheet(SHEET_NAMES.FINANCE);
    financeSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    formatHeaderRow(financeSheet);
  }
  
  return { 
    donors: donorsSheet, 
    groups: groupsSheet, 
    settings: settingsSheet,
    history: historySheet,
    toolkit: toolkitSheet,
    addresses: addressesSheet,
    scouts: scoutsSheet,
    groomGrants: groomSheet,
    finance: financeSheet
  };
}

function formatHeaderRow(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#D4AF37');
  headerRange.setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

function getAllDonors() {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.DONORS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, donors: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var donors = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var donor = {};
      
      for (var j = 0; j < headers.length; j++) {
        var value = row[j];
        var key = headers[j];
        
        if (key === 'amount' || key === 'personalGoal') {
          donor[key] = value ? Number(value) : 0;
        }
        else if (key === 'id' || key === 'groupId') {
          donor[key] = value || '';
        }
        else if (key === 'history') {
          try {
            donor[key] = value ? JSON.parse(value) : [];
          } catch (e) {
            donor[key] = [];
          }
        }
        else {
          donor[key] = value || '';
        }
      }
      
      if (donor.id) {
        donors.push(donor);
      }
    }
    
    return { success: true, donors: donors };
  } catch (error) {
    Logger.log('שגיאה בקריאת מתרימים: ' + error.toString());
    return { success: false, error: error.toString(), donors: [] };
  }
}

function getAllGroups() {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, groups: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var groups = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var group = {};
      
      for (var j = 0; j < headers.length; j++) {
        var value = row[j];
        var key = headers[j];
        
        if (key === 'goal' || key === 'orderNumber') {
          group[key] = value ? Number(value) : 0;
        }
        else if (key === 'id') {
          group[key] = value || '';
        }
        else {
          group[key] = value || '';
        }
      }
      
      if (group.id && group.name) {
        groups.push(group);
      }
    }
    
    return { success: true, groups: groups };
  } catch (error) {
    Logger.log('שגיאה בקריאת קבוצות: ' + error.toString());
    return { success: false, error: error.toString(), groups: [] };
  }
}

function getAllSettings() {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, settings: {} };
    }
    
    var data = sheet.getDataRange().getValues();
    var settings = {};
    
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      var value = data[i][1];
      
      if (key) {
        try {
          settings[key] = JSON.parse(value);
        } catch (e) {
          settings[key] = value;
        }
      }
    }
    
    return { success: true, settings: settings };
  } catch (error) {
    Logger.log('שגיאה בקריאת הגדרות: ' + error.toString());
    return { success: false, error: error.toString(), settings: {} };
  }
}

function getAllData() {
  try {
    var donorsResult = getAllDonors();
    var groupsResult = getAllGroups();
    var settingsResult = getAllSettings();
    
    return {
      success: true,
      donors: donorsResult.donors || [],
      groups: groupsResult.groups || [],
      settings: settingsResult.settings || {},
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('שגיאה בקריאת כל הנתונים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveDonors(donors) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.DONORS);
    
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    
    if (!donors || donors.length === 0) {
      return { success: true, message: 'אין מתרימים לשמירה' };
    }
    
    var rows = [];
    var now = new Date().toISOString();
    
    for (var i = 0; i < donors.length; i++) {
      var donor = donors[i];
      var row = [];
      
      for (var j = 0; j < DONOR_COLUMNS.length; j++) {
        var col = DONOR_COLUMNS[j];
        var value = donor[col];
        
        if (col === 'history') {
          row.push(JSON.stringify(value || []));
        }
        else if (col === 'updatedAt') {
          row.push(now);
        }
        else if (col === 'createdAt') {
          row.push(donor.createdAt || now);
        }
        else {
          row.push(value !== undefined ? value : '');
        }
      }
      
      rows.push(row);
    }
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, DONOR_COLUMNS.length).setValues(rows);
    }
    
    return { success: true, message: 'נשמרו ' + donors.length + ' מתרימים', count: donors.length };
  } catch (error) {
    Logger.log('שגיאה בשמירת מתרימים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveGroups(groups) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
    
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    
    if (!groups || groups.length === 0) {
      return { success: true, message: 'אין קבוצות לשמירה' };
    }
    
    var rows = [];
    var now = new Date().toISOString();
    
    for (var i = 0; i < groups.length; i++) {
      var group = groups[i];
      var row = [];
      
      for (var j = 0; j < GROUP_COLUMNS.length; j++) {
        var col = GROUP_COLUMNS[j];
        var value = group[col];
        
        if (col === 'updatedAt') {
          row.push(now);
        }
        else if (col === 'createdAt') {
          row.push(group.createdAt || now);
        }
        else {
          row.push(value !== undefined ? value : '');
        }
      }
      
      rows.push(row);
    }
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, GROUP_COLUMNS.length).setValues(rows);
    }
    
    return { success: true, message: 'נשמרו ' + groups.length + ' קבוצות', count: groups.length };
  } catch (error) {
    Logger.log('שגיאה בשמירת קבוצות: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveSettings(settings) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    
    if (!settings || Object.keys(settings).length === 0) {
      return { success: true, message: 'אין הגדרות לשמירה' };
    }
    
    var rows = [];
    var keys = Object.keys(settings);
    
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      var value = settings[key];
      
      if (typeof value === 'object') {
        value = JSON.stringify(value);
      }
      
      rows.push([key, value]);
    }
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }
    
    return { success: true, message: 'נשמרו ' + keys.length + ' הגדרות', count: keys.length };
  } catch (error) {
    Logger.log('שגיאה בשמירת הגדרות: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveAllData(data) {
  try {
    var results = {
      donors: { success: false },
      groups: { success: false },
      settings: { success: false }
    };
    
    if (data.donors) {
      results.donors = saveDonors(data.donors);
    }
    
    if (data.groups) {
      results.groups = saveGroups(data.groups);
    }
    
    if (data.settings) {
      results.settings = saveSettings(data.settings);
    }
    
    return {
      success: true,
      results: results,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('שגיאה בשמירת כל הנתונים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function syncAllData(clientData) {
  try {
    var serverData = getAllData();
    
    if (!serverData.success) {
      return { success: false, error: 'שגיאה בקריאת נתונים מהגליון' };
    }
    
    if ((!serverData.donors || serverData.donors.length === 0) && 
        (!serverData.groups || serverData.groups.length === 0)) {
      return saveAllData(clientData);
    }
    
    var mergedDonors = mergeDonors(clientData.donors || [], serverData.donors || []);
    var mergedGroups = mergeGroups(clientData.groups || [], serverData.groups || []);
    
    var saveResult = saveAllData({
      donors: mergedDonors,
      groups: mergedGroups,
      settings: clientData.settings || {}
    });
    
    return {
      success: true,
      donors: mergedDonors,
      groups: mergedGroups,
      settings: clientData.settings || {},
      merged: true,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('שגיאה בסנכרון: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function mergeDonors(clientDonors, serverDonors) {
  var merged = {};
  
  for (var i = 0; i < serverDonors.length; i++) {
    var donor = serverDonors[i];
    if (donor.id) {
      merged[donor.id] = donor;
    }
  }
  
  for (var j = 0; j < clientDonors.length; j++) {
    var clientDonor = clientDonors[j];
    if (clientDonor.id) {
      var existing = merged[clientDonor.id];
      
      if (existing) {
        var clientDate = clientDonor.updatedAt ? new Date(clientDonor.updatedAt) : new Date(0);
        var serverDate = existing.updatedAt ? new Date(existing.updatedAt) : new Date(0);
        
        if (clientDate >= serverDate) {
          merged[clientDonor.id] = clientDonor;
        }
      } else {
        merged[clientDonor.id] = clientDonor;
      }
    }
  }
  
  return Object.values(merged);
}

function mergeGroups(clientGroups, serverGroups) {
  var merged = {};
  
  for (var i = 0; i < serverGroups.length; i++) {
    var group = serverGroups[i];
    if (group.id) {
      merged[group.id] = group;
    }
  }
  
  for (var j = 0; j < clientGroups.length; j++) {
    var clientGroup = clientGroups[j];
    if (clientGroup.id) {
      var existing = merged[clientGroup.id];
      
      if (existing) {
        var clientDate = clientGroup.updatedAt ? new Date(clientGroup.updatedAt) : new Date(0);
        var serverDate = existing.updatedAt ? new Date(existing.updatedAt) : new Date(0);
        
        if (clientDate >= serverDate) {
          merged[clientGroup.id] = clientGroup;
        }
      } else {
        merged[clientGroup.id] = clientGroup;
      }
    }
  }
  
  return Object.values(merged);
}

function testConnection() {
  try {
    ensureSheetsExist();
    var data = getAllData();
    Logger.log('חיבור תקין! מתרימים: ' + (data.donors ? data.donors.length : 0) + ', קבוצות: ' + (data.groups ? data.groups.length : 0));
    return { success: true, message: 'החיבור תקין!', data: data };
  } catch (error) {
    Logger.log('שגיאה בבדיקת חיבור: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function readSimpleSheet(sourceSheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sourceSheet = ss.getSheetByName(sourceSheetName || 'גיליון1');
    
    if (!sourceSheet) {
      var sheets = ss.getSheets();
      if (sheets.length > 0) {
        sourceSheet = sheets[0];
      } else {
        return { success: false, error: 'לא נמצא גיליון' };
      }
    }
    
    var data = sourceSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, donors: [], groups: [], message: 'הגיליון ריק' };
    }
    
    var donors = [];
    var groupsMap = {};
    var groupId = 1;
    var donorId = 1;
    
    for (var i = 1; i < data.length; i++) {
      var name = data[i][0];
      var groupName = data[i][1];
      
      if (!name) continue;
      
      if (groupName && !groupsMap[groupName]) {
        groupsMap[groupName] = {
          id: groupId++,
          name: groupName,
          goal: 0,
          orderNumber: Object.keys(groupsMap).length + 1,
          createdAt: new Date().toISOString()
        };
      }
      
      donors.push({
        id: donorId++,
        name: name,
        displayName: name,
        groupId: groupName ? groupsMap[groupName].id : 1,
        amount: 0,
        personalGoal: 2100,
        history: [],
        createdAt: new Date().toISOString()
      });
    }
    
    var groups = Object.values(groupsMap);
    
    return {
      success: true,
      donors: donors,
      groups: groups,
      message: 'נקראו ' + donors.length + ' מתרימים ו-' + groups.length + ' קבוצות',
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('שגיאה בקריאת גיליון פשוט: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function importFromSimpleSheet(sourceSheetName) {
  try {
    var result = readSimpleSheet(sourceSheetName);
    if (!result.success) return result;
    
    saveDonors(result.donors);
    saveGroups(result.groups);
    
    return {
      success: true,
      message: 'יובאו ' + result.donors.length + ' מתרימים ו-' + result.groups.length + ' קבוצות',
      donors: result.donors.length,
      groups: result.groups.length
    };
  } catch (error) {
    Logger.log('שגיאה בייבוא: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function addHistoryEntry(entry) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
    
    var now = new Date();
    var row = [
      entry.id || 'hist_' + now.getTime(),
      now.toISOString(),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
      entry.actionType || '',
      entry.entityType || '',
      entry.entityId || '',
      entry.entityName || '',
      entry.details || '',
      entry.amount || 0,
      entry.source || 'manual',
      entry.computerName || '',
      entry.computerId || '',
      entry.userId || ''
    ];
    
    sheet.appendRow(row);
    
    return { success: true, id: row[0] };
  } catch (error) {
    Logger.log('שגיאה בהוספת היסטוריה: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function addHistoryEntries(entries) {
  try {
    if (!entries || entries.length === 0) {
      return { success: true, message: 'אין רשומות להוספה' };
    }
    
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
    
    var rows = [];
    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      var now = entry.timestamp ? new Date(entry.timestamp) : new Date();
      
      rows.push([
        entry.id || 'hist_' + now.getTime() + '_' + i,
        now.toISOString(),
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
        entry.actionType || '',
        entry.entityType || '',
        entry.entityId || '',
        entry.entityName || '',
        entry.details || '',
        entry.amount || 0,
        entry.source || 'manual',
        entry.computerName || '',
        entry.computerId || '',
        entry.userId || ''
      ]);
    }
    
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HISTORY_COLUMNS.length).setValues(rows);
    }
    
    return { success: true, count: rows.length };
  } catch (error) {
    Logger.log('שגיאה בהוספת רשומות היסטוריה: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function getHistory(filters) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, history: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var history = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var entry = {};
      
      for (var j = 0; j < headers.length; j++) {
        entry[headers[j]] = row[j] || '';
      }
      
      if (filters) {
        var entryDate = new Date(entry.timestamp);
        
        if (filters.startDate) {
          var startDate = new Date(filters.startDate);
          startDate.setHours(0, 0, 0, 0);
          if (entryDate < startDate) continue;
        }
        
        if (filters.endDate) {
          var endDate = new Date(filters.endDate);
          endDate.setHours(23, 59, 59, 999);
          if (entryDate > endDate) continue;
        }
        
        if (filters.source && entry.source !== filters.source) continue;
        if (filters.actionType && entry.actionType !== filters.actionType) continue;
        if (filters.entityType && entry.entityType !== filters.entityType) continue;
        if (filters.computerName && entry.computerName !== filters.computerName) continue;
      }
      
      history.push(entry);
    }
    
    history.sort(function(a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    
    return { success: true, history: history, count: history.length };
  } catch (error) {
    Logger.log('שגיאה בקריאת היסטוריה: ' + error.toString());
    return { success: false, error: error.toString(), history: [] };
  }
}

function saveGenericData(sheetName, key, value) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, error: 'גיליון לא נמצא: ' + sheetName };
    }
    
    var valueStr = typeof value === 'object' ? JSON.stringify(value) : String(value);
    
    var data = sheet.getDataRange().getValues();
    var foundRow = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow > 0) {
      sheet.getRange(foundRow, 2).setValue(valueStr);
    } else {
      sheet.appendRow([key, valueStr]);
    }
    
    return { success: true };
  } catch (error) {
    Logger.log('שגיאה בשמירת נתונים ל-' + sheetName + ': ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function loadGenericData(sheetName, key) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: null };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        var value = data[i][1];
        try {
          return { success: true, data: JSON.parse(value) };
        } catch (e) {
          return { success: true, data: value };
        }
      }
    }
    
    return { success: true, data: null };
  } catch (error) {
    Logger.log('שגיאה בקריאת נתונים מ-' + sheetName + ': ' + error.toString());
    return { success: false, error: error.toString(), data: null };
  }
}

function saveAllExtendedData(data) {
  try {
    var results = {};
    
    if (data.toolkitTips !== undefined) {
      results.toolkit = saveGenericData(SHEET_NAMES.TOOLKIT, 'toolkitTips', data.toolkitTips);
    }
    
    if (data.scoutsSchedule !== undefined) {
      results.scouts = saveGenericData(SHEET_NAMES.SCOUTS, 'scoutsSchedule', data.scoutsSchedule);
    }
    
    if (data.groomGrants !== undefined) {
      results.groomGrants = saveGenericData(SHEET_NAMES.GROOM_GRANTS, 'groomGrants', data.groomGrants);
    }
    
    if (data.addresses !== undefined) {
      results.addresses = saveGenericData(SHEET_NAMES.ADDRESSES, 'addresses', data.addresses);
    }
    
    if (data.financeState !== undefined) {
      results.finance = saveGenericData(SHEET_NAMES.FINANCE, 'financeState', data.financeState);
    }
    
    if (data.liveViewSettings !== undefined) {
      results.liveView = saveGenericData(SHEET_NAMES.SETTINGS, 'liveViewSettings', data.liveViewSettings);
    }
    
    return { success: true, results: results };
  } catch (error) {
    Logger.log('שגיאה בשמירת נתונים נוספים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function getAllExtendedData() {
  try {
    return {
      success: true,
      toolkitTips: loadGenericData(SHEET_NAMES.TOOLKIT, 'toolkitTips').data,
      scoutsSchedule: loadGenericData(SHEET_NAMES.SCOUTS, 'scoutsSchedule').data,
      groomGrants: loadGenericData(SHEET_NAMES.GROOM_GRANTS, 'groomGrants').data,
      addresses: loadGenericData(SHEET_NAMES.ADDRESSES, 'addresses').data,
      financeState: loadGenericData(SHEET_NAMES.FINANCE, 'financeState').data,
      liveViewSettings: loadGenericData(SHEET_NAMES.SETTINGS, 'liveViewSettings').data
    };
  } catch (error) {
    Logger.log('שגיאה בקריאת נתונים נוספים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function getAllDataComplete() {
  try {
    var basicData = getAllData();
    var extendedData = getAllExtendedData();
    
    return {
      success: true,
      donors: basicData.donors || [],
      groups: basicData.groups || [],
      settings: basicData.settings || {},
      toolkitTips: extendedData.toolkitTips || [],
      scoutsSchedule: extendedData.scoutsSchedule || null,
      groomGrants: extendedData.groomGrants || [],
      addresses: extendedData.addresses || [],
      financeState: extendedData.financeState || null,
      liveViewSettings: extendedData.liveViewSettings || null,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('שגיאה בקריאת כל הנתונים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveAllDataComplete(data) {
  try {
    var basicResults = saveAllData(data);
    var extendedResults = saveAllExtendedData(data);
    
    return {
      success: true,
      basicResults: basicResults,
      extendedResults: extendedResults,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    Logger.log('שגיאה בשמירת כל הנתונים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
