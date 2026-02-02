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
  SETTINGS: 'הגדרות'
};

// ========== מבנה עמודות פשוט וברור ==========
// עמודות למתרימים
const DONOR_COLUMNS = [
  'id',              // מזהה ייחודי
  'name',            // שם מלא (מקורי)
  'displayName',     // שם תצוגה (אפשר לערוך)
  'groupId',         // מזהה קבוצה
  'groupName',       // שם הקבוצה (לנוחות)
  'amount',          // סכום שגויס
  'personalGoal',    // יעד אישי
  'source',          // מקור: nedarim_plus / manual
  'nedarimMatrimId', // מספר מתרים בנדרים פלוס
  'createdAt',       // תאריך יצירה
  'updatedAt'        // תאריך עדכון
];

// עמודות לקבוצות
const GROUP_COLUMNS = [
  'id',              // מזהה ייחודי
  'name',            // שם הקבוצה
  'goal',            // יעד הקבוצה
  'orderNumber',     // סדר בתצוגה
  'showInLiveView',  // האם להציג בלייב
  'createdAt',       // תאריך יצירה
  'updatedAt'        // תאריך עדכון
];

// *****************************************************************************
// *** פונקציות HTTP ראשיות ***
// *****************************************************************************

function doGet(e) {
  var output;
  
  try {
    var action = e && e.parameter ? e.parameter.action : '';
    
    switch(action) {
      // פעולות נדרים פלוס
      case 'getNedarimTotal':
        output = getNedarimTotalDonations();
        break;
      case 'searchNedarimRecruiters':
        var searchTerm = e && e.parameter ? e.parameter.search : '';
        output = searchNedarimRecruiters(searchTerm);
        break;
      case 'getPublicConfig':
        output = { success: true, mosadId: NEDARIM_CONFIG.MOSAD_ID, matchingId: NEDARIM_CONFIG.MATCHING_ID };
        break;
        
      // פעולות סנכרון
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
        
      default:
        output = { success: false, error: 'פעולה לא מוכרת: ' + action };
    }
    
  } catch (error) {
    Logger.log('שגיאה ב-doGet: ' + error.toString());
    output = { success: false, error: error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var output;
  
  try {
    var requestData = {};
    
    if (e && e.postData && e.postData.contents) {
      try {
        requestData = JSON.parse(e.postData.contents);
      } catch (parseError) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: 'שגיאה בפרסור JSON'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    var action = requestData.action || '';
    
    switch(action) {
      case 'saveDonors':
        output = saveDonors(requestData.donors, requestData.groups);
        break;
      case 'saveGroups':
        output = saveGroups(requestData.groups);
        break;
      case 'saveSettings':
        output = saveSettings(requestData.settings);
        break;
      case 'saveAll':
        output = saveAllData(requestData);
        break;
        
      default:
        output = { success: false, error: 'פעולה לא מוכרת: ' + action };
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
 * קבלת סכום תרומות כולל מנדרים פלוס
 * מחשב את הסכום האמיתי מכל המתרימים
 */
function getNedarimTotalDonations() {
  try {
    var totalFromRecruiters = 0;
    var goalFromAPI = 0;
    
    // ניסיון 1: קבלת היעד מ-ShowGoal API
    Logger.log('מנסה לקבל יעד מ-ShowGoal API...');
    try {
      var url1 = NEDARIM_CONFIG.MATCH_PLUS_API_URL + 
        '?Action=ShowGoal&MosadId=' + NEDARIM_CONFIG.MOSAD_ID + 
        '&GoalId=' + NEDARIM_CONFIG.MATCHING_ID;
      
      var response1 = UrlFetchApp.fetch(url1, { method: 'GET', muteHttpExceptions: true });
      
      if (response1.getResponseCode() === 200) {
        var data = JSON.parse(response1.getContentText());
        goalFromAPI = parseFloat(data.Goal) || parseFloat(data.TargetSum) || 0;
        
        // אם יש גם סכום תרומות - נשתמש בו
        var donatedFromAPI = parseFloat(data.Donated) || parseFloat(data.DSum) || parseFloat(data.TotalDonated) || 0;
        if (donatedFromAPI > 0) {
          Logger.log('סכום מ-API: ' + donatedFromAPI + ', יעד: ' + goalFromAPI);
          return { 
            success: true, 
            totalDonated: donatedFromAPI, 
            goal: goalFromAPI, 
            source: 'ShowGoal' 
          };
        }
      }
    } catch (e) {
      Logger.log('שגיאה ב-ShowGoal: ' + e.message);
    }
    
    // ניסיון 2: חישוב הסכום האמיתי מכל המתרימים
    Logger.log('מחשב סכום כולל מכל המתרימים...');
    
    var hebrewLetters = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ', 'ק', 'ר', 'ש', 'ת'];
    var allRecruiters = {};
    
    // חיפוש ריק קודם
    try {
      var emptyResult = searchNedarimRecruiters('');
      if (emptyResult && emptyResult.success && emptyResult.recruiters) {
        emptyResult.recruiters.forEach(function(r) {
          var id = r.Id || r.MatrimId || r.Name;
          if (id) allRecruiters[id] = r;
        });
        Logger.log('חיפוש ריק: ' + emptyResult.recruiters.length + ' מתרימים');
      }
    } catch (e) {
      Logger.log('שגיאה בחיפוש ריק: ' + e.message);
    }
    
    // חיפוש לפי אותיות עבריות
    for (var i = 0; i < hebrewLetters.length; i++) {
      var letter = hebrewLetters[i];
      try {
        var result = searchNedarimRecruiters(letter);
        if (result && result.success && result.recruiters) {
          result.recruiters.forEach(function(r) {
            var id = r.Id || r.MatrimId || r.Name;
            if (id) allRecruiters[id] = r;
          });
        }
      } catch (e) {
        Logger.log('שגיאה בחיפוש ' + letter + ': ' + e.message);
      }
    }
    
    // חישוב הסכום הכולל
    var recruitersArray = [];
    for (var key in allRecruiters) {
      if (allRecruiters.hasOwnProperty(key)) {
        recruitersArray.push(allRecruiters[key]);
      }
    }
    
    for (var j = 0; j < recruitersArray.length; j++) {
      var r = recruitersArray[j];
      var amount = parseFloat(r.Cumule) || parseFloat(r.Amount) || parseFloat(r.Collected) || 0;
      totalFromRecruiters += amount;
    }
    
    Logger.log('סכום כולל מ-' + recruitersArray.length + ' מתרימים: ' + totalFromRecruiters);
    
    if (totalFromRecruiters > 0) {
      return {
        success: true,
        totalDonated: totalFromRecruiters,
        goal: goalFromAPI,
        source: 'Calculated-FromAllRecruiters',
        recruitersCount: recruitersArray.length
      };
    }
    
    return { success: false, error: 'לא הצלחתי לקבל נתונים', totalDonated: 0, goal: goalFromAPI };
    
  } catch (e) {
    Logger.log('שגיאה כללית: ' + e.message);
    return { success: false, error: e.message, totalDonated: 0, goal: 0 };
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
    
    var response = UrlFetchApp.fetch(url, { method: 'GET', muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      
      if (Array.isArray(data)) {
        return { success: true, recruiters: data };
      } else if (data && data.Matrim) {
        return { success: true, recruiters: data.Matrim };
      } else if (data && data.recruiters) {
        return { success: true, recruiters: data.recruiters };
      }
      
      return { success: true, recruiters: [] };
    }
    
    return { success: false, error: 'שגיאת שרת', recruiters: [] };
    
  } catch (e) {
    Logger.log('שגיאה בחיפוש: ' + e.message);
    return { success: false, error: e.message, recruiters: [] };
  }
}

// *****************************************************************************
// *** פונקציות גיליון ***
// *****************************************************************************

/**
 * יצירת גיליונות אם לא קיימים
 */
function ensureSheetsExist() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // גיליון מתרימים
  var donorsSheet = ss.getSheetByName(SHEET_NAMES.DONORS);
  if (!donorsSheet) {
    donorsSheet = ss.insertSheet(SHEET_NAMES.DONORS);
    donorsSheet.getRange(1, 1, 1, DONOR_COLUMNS.length).setValues([DONOR_COLUMNS]);
    formatHeader(donorsSheet);
  } else {
    // בדיקה ועדכון כותרות
    var headers = donorsSheet.getRange(1, 1, 1, donorsSheet.getLastColumn()).getValues()[0];
    if (headers.length !== DONOR_COLUMNS.length || headers[0] !== DONOR_COLUMNS[0]) {
      donorsSheet.getRange(1, 1, 1, DONOR_COLUMNS.length).setValues([DONOR_COLUMNS]);
      formatHeader(donorsSheet);
    }
  }
  
  // גיליון קבוצות
  var groupsSheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
  if (!groupsSheet) {
    groupsSheet = ss.insertSheet(SHEET_NAMES.GROUPS);
    groupsSheet.getRange(1, 1, 1, GROUP_COLUMNS.length).setValues([GROUP_COLUMNS]);
    formatHeader(groupsSheet);
  }
  
  // גיליון הגדרות
  var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settingsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    formatHeader(settingsSheet);
  }
  
  return { donors: donorsSheet, groups: groupsSheet, settings: settingsSheet };
}

function formatHeader(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#D4AF37');
  headerRange.setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

// *****************************************************************************
// *** קריאת נתונים ***
// *****************************************************************************

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
        var key = headers[j];
        var value = row[j];
        
        if (!key) continue;
        
        if (key === 'amount' || key === 'personalGoal') {
          donor[key] = value ? Number(value) : 0;
        } else if (key === 'groupId') {
          donor[key] = value !== '' && value !== null ? value : '';
        } else {
          donor[key] = value !== undefined && value !== null ? value : '';
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
        var key = headers[j];
        var value = row[j];
        
        if (!key) continue;
        
        if (key === 'goal' || key === 'orderNumber') {
          group[key] = value ? Number(value) : 0;
        } else if (key === 'showInLiveView') {
          group[key] = value !== false && value !== 'false';
        } else {
          group[key] = value !== undefined && value !== null ? value : '';
        }
      }
      
      if (group.id) {
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
    var donors = getAllDonors();
    var groups = getAllGroups();
    var settings = getAllSettings();
    
    return {
      success: true,
      donors: donors.donors || [],
      groups: groups.groups || [],
      settings: settings.settings || {},
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('שגיאה בקריאת כל הנתונים: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// *****************************************************************************
// *** שמירת נתונים ***
// *****************************************************************************

function saveDonors(donors, groups) {
  try {
    ensureSheetsExist();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.DONORS);
    
    // מחיקת נתונים ישנים
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    
    if (!donors || donors.length === 0) {
      return { success: true, message: 'אין מתרימים לשמירה', count: 0 };
    }
    
    // בניית מפת קבוצות (id -> name)
    var groupsMap = {};
    if (groups && groups.length > 0) {
      for (var g = 0; g < groups.length; g++) {
        groupsMap[groups[g].id] = groups[g].name || '';
      }
    } else {
      // אם אין קבוצות, ננסה לקרוא מהגיליון
      var groupsResult = getAllGroups();
      if (groupsResult.groups) {
        for (var g = 0; g < groupsResult.groups.length; g++) {
          groupsMap[groupsResult.groups[g].id] = groupsResult.groups[g].name || '';
        }
      }
    }
    
    // בניית שורות
    var rows = [];
    var now = new Date().toISOString();
    
    for (var i = 0; i < donors.length; i++) {
      var donor = donors[i];
      var row = [];
      
      for (var j = 0; j < DONOR_COLUMNS.length; j++) {
        var col = DONOR_COLUMNS[j];
        
        switch(col) {
          case 'groupName':
            row.push(groupsMap[donor.groupId] || '');
            break;
          case 'updatedAt':
            row.push(now);
            break;
          case 'createdAt':
            row.push(donor.createdAt || now);
            break;
          case 'source':
            row.push(donor.source || donor.fromNedarimPlus ? 'nedarim_plus' : 'manual');
            break;
          default:
            var value = donor[col];
            row.push(value !== undefined && value !== null ? value : '');
        }
      }
      
      rows.push(row);
    }
    
    // כתיבה לגיליון
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, DONOR_COLUMNS.length).setValues(rows);
    }
    
    Logger.log('נשמרו ' + donors.length + ' מתרימים');
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
    
    // מחיקת נתונים ישנים
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    
    if (!groups || groups.length === 0) {
      return { success: true, message: 'אין קבוצות לשמירה', count: 0 };
    }
    
    // בניית שורות
    var rows = [];
    var now = new Date().toISOString();
    
    for (var i = 0; i < groups.length; i++) {
      var group = groups[i];
      var row = [];
      
      for (var j = 0; j < GROUP_COLUMNS.length; j++) {
        var col = GROUP_COLUMNS[j];
        
        switch(col) {
          case 'updatedAt':
            row.push(now);
            break;
          case 'createdAt':
            row.push(group.createdAt || now);
            break;
          case 'showInLiveView':
            row.push(group.showInLiveView !== false);
            break;
          default:
            var value = group[col];
            row.push(value !== undefined && value !== null ? value : '');
        }
      }
      
      rows.push(row);
    }
    
    // כתיבה לגיליון
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, GROUP_COLUMNS.length).setValues(rows);
    }
    
    Logger.log('נשמרו ' + groups.length + ' קבוצות');
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
    
    // מחיקת נתונים ישנים
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    
    if (!settings || Object.keys(settings).length === 0) {
      return { success: true, message: 'אין הגדרות לשמירה', count: 0 };
    }
    
    // בניית שורות
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
    
    // כתיבה לגיליון
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }
    
    Logger.log('נשמרו ' + keys.length + ' הגדרות');
    return { success: true, message: 'נשמרו ' + keys.length + ' הגדרות', count: keys.length };
    
  } catch (error) {
    Logger.log('שגיאה בשמירת הגדרות: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveAllData(data) {
  try {
    var results = {};
    
    // שומרים קבוצות קודם
    if (data.groups) {
      results.groups = saveGroups(data.groups);
    }
    
    // שומרים מתרימים עם הקבוצות
    if (data.donors) {
      results.donors = saveDonors(data.donors, data.groups);
    }
    
    // שומרים הגדרות
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

// *****************************************************************************
// *** פונקציית בדיקה ***
// *****************************************************************************

function testConnection() {
  Logger.log('=== בדיקת חיבור ===');
  
  // בדיקת גיליון
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('✅ גיליון נפתח בהצלחה: ' + ss.getName());
  } catch (e) {
    Logger.log('❌ שגיאה בפתיחת גיליון: ' + e.message);
    return { success: false, error: 'לא ניתן לפתוח את הגיליון' };
  }
  
  // בדיקת נדרים פלוס
  var nedarimResult = getNedarimTotalDonations();
  Logger.log('נדרים פלוס: ' + JSON.stringify(nedarimResult));
  
  // בדיקת נתונים
  var data = getAllData();
  Logger.log('מתרימים: ' + (data.donors ? data.donors.length : 0));
  Logger.log('קבוצות: ' + (data.groups ? data.groups.length : 0));
  
  return {
    success: true,
    spreadsheet: ss.getName(),
    nedarimPlus: nedarimResult,
    donorsCount: data.donors ? data.donors.length : 0,
    groupsCount: data.groups ? data.groups.length : 0
  };
}
