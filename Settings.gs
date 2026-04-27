/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Settings.gs
 * VERSION: 1.6 (Utility Integration - Generic Upload)
 * SYNC STATUS: Fully Synchronized with SettingsData.html & Utils.gs
 */

/**
 * Settings Module - Backend
 * Standardized under SparkHub Architecture Blueprint.
 */

/**
 * Retrieves global environment, branding, storage, and registry settings.
 */
function getSystemSettings() {
  try {
    var props = PropertiesService.getScriptProperties();
    var logoId = props.getProperty('SYSTEM_LOGO_ID');
    var logoUrl = logoId ? ("https://drive.google.com/thumbnail?id=" + logoId + "&sz=w500") : "";
    var appSvg = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Cpath d='M54 20 L30 55 L52 55 L42 80 L70 42 L48 42 Z' fill='%236366F1'/%3E%3Ccircle cx='68' cy='28' r='7' fill='%2306B6D4'/%3E%3C/svg%3E";
    
    function clean(val, fallback) {
      if (!val) return fallback;
      var str = String(val).trim();
      return (str.indexOf('#') === 0) ? str : '#' + str;
    }
    
    function getSafeProp(key, fallback) {
      var val = props.getProperty(key) || props.getProperty(key.toLowerCase()) || props.getProperty(key.toUpperCase());
      return clean(val, fallback);
    }

    return {
      environment: props.getProperty('ENVIRONMENT') || 'Production',
      adminEmail: props.getProperty('ADMIN_EMAIL') || '',
      systemName: props.getProperty('SYSTEM_NAME') || 'SparkHub',
      systemLogoUrl: logoUrl,
      systemLogoId: logoId || '',
      appFallbackLogo: appSvg,
      emailFallbackLogo: props.getProperty('EMAIL_FALLBACK_LOGO') || '',
      rootFolderId: props.getProperty('ROOT_FOLDER_ID') || '',
      mainDbId: props.getProperty('DATABASE_ID') || '',
      notifDbId: props.getProperty('NOTIF_DATABASE_ID') || '',
      installedModules: props.getProperty('INSTALLED_MODULES') || '',
      installedPlugins: props.getProperty('INSTALLED_PLUGINS') || '',
      themePrimary: getSafeProp('THEME_PRIMARY', '#C40004'), 
      themeAccent: getSafeProp('THEME_ACCENT', '#FDDD64'),   
      themeDark: getSafeProp('THEME_DARK', '#323232'),
      themeBg: getSafeProp('THEME_BG', '#FDDD64'),
      themeHover: getSafeProp('THEME_HOVER', '#A30003')
    };
  } catch(e) {
    return { environment: 'Production', systemName: 'SparkHub', themePrimary: '#C40004' };
  }
}

/**
 * Validates a Google Sheet ID structure.
 */
function validateDatabase(id) {
  try {
    var ss = SpreadsheetApp.openById(id);
    if (!ss.getSheetByName("Users") || !ss.getSheetByName("Templates")) throw new Error("Missing Core Sheets.");
    return true;
  } catch (e) { throw new Error("Database Validation Failed: " + e.message); }
}

/**
 * Uploads the custom system logo to Drive utilizing the Generic Upload Engine.
 */
function uploadSystemLogo(base64, filename) {
  try {
    // 1. Navigate/Create the specific Assets path
    var assetsFolder = getSystemSubfolder("System Assets");
    var sFolder = getOrCreateFolder(assetsFolder, "Settings");
    var iFolder = getOrCreateFolder(sFolder, "Images");
    
    // 2. Utilize the Generic Utility
    return uploadBase64File(base64, filename, iFolder);
  } catch (e) { 
    return { error: "Logo Sync Error: " + e.message }; 
  }
}

/**
 * Saves environment overrides globally.
 */
function saveSystemSettings(settings) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (settings.mainDbId) validateDatabase(settings.mainDbId);
    
    if (settings.environment) props.setProperty('ENVIRONMENT', settings.environment);
    if (settings.adminEmail) props.setProperty('ADMIN_EMAIL', settings.adminEmail);
    if (settings.systemName) props.setProperty('SYSTEM_NAME', settings.systemName);
    if (settings.rootFolderId) props.setProperty('ROOT_FOLDER_ID', settings.rootFolderId);
    if (settings.mainDbId) props.setProperty('DATABASE_ID', settings.mainDbId);
    
    if (settings.themePrimary) props.setProperty('THEME_PRIMARY', settings.themePrimary);
    if (settings.systemLogoId) props.setProperty('SYSTEM_LOGO_ID', settings.systemLogoId);

    return "Success! Settings updated.";
  } catch (e) { return "Error: " + e.message; }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */