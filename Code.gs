// =============================================================================
// Code.gs - קוד צד השרת המאובטח למערכת קמפיין פורים
// =============================================================================
// ⚠️ קובץ זה רץ בשרתי Google בלבד - הלקוח לא יכול לראות את התוכן שלו!
// =============================================================================

// *****************************************************************************
// *** הגדרות סודיות - אלו לא נחשפות ללקוח! ***
// *****************************************************************************

const CONFIG = {
  // Google Sheets
  SPREADSHEET_ID: '1YI6XQZObSP1vfhIVh9wXYtIufM20PFjIGGM9rB-_rc8',
  
  // נדרים פלוס API - סודי!
  NEDARIM: {
    MOSAD_ID: '1000642',
    MATCHING_ID: '715',
    API_PASSWORD: 'ep348',
    API_URL: 'https://matara.pro/nedarimplus/Reports/Manage3.aspx',
    ONLINE_API_URL: 'https://www.matara.pro/nedarimplus/online/Files/Manage.aspx',
    MATCH_PLUS_API_URL: 'https://www.matara.pro/nedarimplus/V6/MatchPlus.aspx'
  },
  
  // אבטחה
  SECURITY: {
    ALLOWED_EMAILS: [],
    ENABLE_AUDIT_LOG: true,
    MAX_REQUESTS_PER_MINUTE: 60
  }
};

// *****************************************************************************
// *** פונקציות ציבוריות - זמינות ללקוח ***
// *****************************************************************************

/**
 * קבלת ההגדרות הציבוריות בלבד (ללא סודות)
 * @returns {Object} הגדרות ציבוריות
 */
function getPublicConfig() {
  return {
    mosadId: CONFIG.NEDARIM.MOSAD_ID,
    matchingId: CONFIG.NEDARIM.MATCHING_ID
    // ללא API_PASSWORD!
  };
}

/**
 * קבלת קישור לגליון (ללא חשיפת המזהה בקוד הלקוח)
 * @returns {string} URL לגליון
 */
function getSpreadsheetUrl() {
  return 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID + '/edit';
}

/**
 * קבלת סכום תרומות כולל מנדרים פלוס
 * @returns {Object} תוצאה עם totalDonated, goal וכו'
 */
function getNedarimTotalDonations() {
  try {
    // ניסיון 1: ShowGoal
    const url1 = CONFIG.NEDARIM.ONLINE_API_URL + 
      '?Action=ShowGoal&Ession=&S=1&Mosession=' + CONFIG.NEDARIM.MATCHING_ID + 
      '&MosadId=' + CONFIG.NEDARIM.MOSAD_ID;
    
    const response1 = UrlFetchApp.fetch(url1, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response1.getResponseCode() === 200) {
      const text = response1.getContentText();
      try {
        const data = JSON.parse(text);
        if (data && data.DSum !== undefined) {
          return {
            success: true,
            totalDonated: parseFloat(data.DSum) || 0,
            goal: parseFloat(data.Goal) || 0,
            goalId: data.GoalId || '',
            source: 'ShowGoal'
          };
        }
      } catch (e) {
        // לא JSON - ממשיך לניסיון הבא
      }
    }
    
    // ניסיון 2: GetMosad
    const url2 = CONFIG.NEDARIM.ONLINE_API_URL + 
      '?Action=GetMosad&MosadId=' + CONFIG.NEDARIM.MOSAD_ID;
    
    const response2 = UrlFetchApp.fetch(url2, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response2.getResponseCode() === 200) {
      const data2 = JSON.parse(response2.getContentText());
      if (data2) {
        return {
          success: true,
          totalDonated: parseFloat(data2.TotalDonated) || parseFloat(data2.DSum) || 0,
          goal: parseFloat(data2.Goal) || 0,
          mosadName: data2.Name || '',
          source: 'GetMosad'
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
 * חיפוש מתרימים במערכת MatchPlus
 * @param {string} searchTerm - מונח חיפוש
 * @returns {Object} תוצאות החיפוש
 */
function searchNedarimRecruiters(searchTerm) {
  try {
    const url = CONFIG.NEDARIM.MATCH_PLUS_API_URL + 
      '?Action=SearchMatrim&Name=' + encodeURIComponent(searchTerm || '') + 
      '&MosadId=' + CONFIG.NEDARIM.MOSAD_ID;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
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
 * @param {string} recruiterId - מזהה המתרים
 * @returns {Object} פרטי המתרים
 */
function getNedarimRecruiterDetails(recruiterId) {
  try {
    const url = CONFIG.NEDARIM.MATCH_PLUS_API_URL + 
      '?Action=GetMatrimData&MatrimId=' + encodeURIComponent(recruiterId) + 
      '&MosadId=' + CONFIG.NEDARIM.MOSAD_ID;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
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
 * קבלת דוח מפורט מנדרים פלוס
 * @param {Object} params - פרמטרים לדוח
 * @returns {Object} נתוני הדוח
 */
function getNedarimReport(params) {
  try {
    const formData = {
      'Action': params.action || 'GetReport',
      'MosadNumber': CONFIG.NEDARIM.MOSAD_ID,
      'ApiPassword': CONFIG.NEDARIM.API_PASSWORD
    };
    
    if (params.fromDate) formData['FromDate'] = params.fromDate;
    if (params.toDate) formData['ToDate'] = params.toDate;
    
    const response = UrlFetchApp.fetch(CONFIG.NEDARIM.API_URL, {
      method: 'POST',
      payload: formData,
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      try {
        const data = JSON.parse(response.getContentText());
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
