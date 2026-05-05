/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Settings.gs
 * VERSION: 1.6 (Utility Integration - Generic Upload)
 * SYNC STATUS: Fully Synchronized with SettingsData.html & Utils.gs
 */

// ========================================================================
// 1. REGISTRY EXPORTS
// ========================================================================
function System_getTriggers() {
  return ["System:INSTALL"];
}

function Settings_getTriggers() {
  return ["Settings:UPDATE"];
}

function Settings_getPlaceholders() {
  return ["systemName", "systemLogoUrl", "environment", "adminEmail"];
}

// ========================================================================
// 2. CORE PROCESSORS
// ========================================================================
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

// ========================================================================
// 3. READ / GET FUNCTIONS
// ========================================================================
/**
 * Retrieves global environment, branding, storage, and registry settings.
 */
function getSystemSettings() {
  try {
    var props = PropertiesService.getScriptProperties();
    var logoId = props.getProperty('SYSTEM_LOGO_ID');
    var logoUrl = logoId ? ("https://drive.google.com/thumbnail?id=" + logoId + "&sz=w500") : ("https://drive.google.com/thumbnail?id=1HLIJ-GJSHOmh7CPGR56x5fKH8sptM44e&sz=w500");
    var appSvg = "https://i.imgur.com/0iPmgVk.png";
    
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
      environment: props.getProperty('ENVIRONMENT') || 'Sandbox',
      authMode: props.getProperty('AUTH_MODE') || 'SSO', // <-- NEW LINE
      adminEmail: props.getProperty('ADMIN_EMAIL') || '',
      systemName: props.getProperty('SYSTEM_NAME') || 'SparkHub',
      systemLogoUrl: logoUrl,
      systemLogoId: logoId || '',
      appFallbackLogo: appSvg,
      emailFallbackLogo: props.getProperty('EMAIL_FALLBACK_LOGO') || 'https://i.imgur.com/Nlcwog7.png',
      rootFolderId: props.getProperty('ROOT_FOLDER_ID') || '',
      mainDbId: props.getProperty('DATABASE_ID') || '',
      logsDbId: props.getProperty('LOGS_DATABASE_ID') || '',
      hasWebhookSecret: !!props.getProperty('WEBHOOK_SECRET'), // <-- NEW: Boolean check
      installedModules: props.getProperty('INSTALLED_MODULES') || '',
      installedPlugins: props.getProperty('INSTALLED_PLUGINS') || '',
      themePrimary: getSafeProp('THEME_PRIMARY', '#666DF2'), 
      themeAccent: getSafeProp('THEME_ACCENT', '#0BC4D9'),   
      themeDark: getSafeProp('THEME_DARK', '#00283A'),
      themeBg: getSafeProp('THEME_BG', '#F1F5F9'),
      themeHover: getSafeProp('THEME_HOVER', '#7E84F2')
    };
  } catch(e) {
    return { environment: 'Production', systemName: 'SparkHub', themePrimary: '#666DF2' };
  }
}

// ========================================================================
// 4. WRITE / SAVE FUNCTIONS
// ========================================================================
/**
 * Saves environment overrides globally.
 */
function saveSystemSettings(settings) {
  try {
    var props = PropertiesService.getScriptProperties();
    
    // Validate databases if IDs are being changed
    if (settings.mainDbId) validateDatabase(settings.mainDbId);
    if (settings.logsDbId) validateLogsDatabase(settings.logsDbId);
    
    // Core Infrastructure & Identity
    if (settings.environment) props.setProperty('ENVIRONMENT', settings.environment);
    if (settings.authMode) props.setProperty('AUTH_MODE', settings.authMode);
    if (settings.adminEmail) props.setProperty('ADMIN_EMAIL', settings.adminEmail);
    if (settings.systemName) props.setProperty('SYSTEM_NAME', settings.systemName);
    if (settings.rootFolderId) props.setProperty('ROOT_FOLDER_ID', settings.rootFolderId);
    if (settings.mainDbId) props.setProperty('DATABASE_ID', settings.mainDbId);
    if (settings.logsDbId) props.setProperty('LOGS_DATABASE_ID', settings.logsDbId);
    if (settings.fallbackLogoUrl) props.setProperty('EMAIL_FALLBACK_LOGO', settings.fallbackLogoUrl);
    
    // THEME ENGINE SYNC: Explicitly save all 5 color tokens
    if (settings.themePrimary) props.setProperty('THEME_PRIMARY', settings.themePrimary);
    if (settings.themeAccent) props.setProperty('THEME_ACCENT', settings.themeAccent);
    if (settings.themeDark) props.setProperty('THEME_DARK', settings.themeDark);
    if (settings.themeHover) props.setProperty('THEME_HOVER', settings.themeHover);
    if (settings.themeBg) props.setProperty('THEME_BG', settings.themeBg);
    
    // Assets
    if (settings.systemLogoId) props.setProperty('SYSTEM_LOGO_ID', settings.systemLogoId);

    SystemEvent.emit("Settings", "UPDATE", "System Configuration", "WARN", "Global Settings", "Core system architecture, identity, or theme settings were modified.");
    return "Success! Settings updated.";
  } catch (e) { 
    return "Error: " + e.message;
  }
}

// ========================================================================
// 5. INTERNAL HELPERS
// ========================================================================
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

function validateLogsDatabase(id) {
  try {
    var ss = SpreadsheetApp.openById(id);
    if (!ss.getSheetByName("System Logs")) throw new Error("Missing System Logs Sheet.");
    return true;
  } catch (e) { throw new Error("Logs Database Validation Failed: " + e.message); }
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */