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
  return ["System:INSTALL", "System:MODULE_INSTALLED"];
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
    var assetsFolder = getSystemSubfolder("System Assets");
    var sFolder = getOrCreateFolder(assetsFolder, "Settings");
    var iFolder = getOrCreateFolder(sFolder, "Images");
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
      authMode: props.getProperty('AUTH_MODE') || 'SSO', 
      adminEmail: props.getProperty('ADMIN_EMAIL') || '',
      systemName: props.getProperty('SYSTEM_NAME') || 'SparkHub',
      systemLogoUrl: logoUrl,
      systemLogoId: logoId || '',
      appFallbackLogo: appSvg,
      emailFallbackLogo: props.getProperty('EMAIL_FALLBACK_LOGO') || 'https://i.imgur.com/Nlcwog7.png',
      rootFolderId: props.getProperty('ROOT_FOLDER_ID') || '',
      mainDbId: props.getProperty('DATABASE_ID') || '',
      logsDbId: props.getProperty('LOGS_DATABASE_ID') || '',
      hasWebhookSecret: !!props.getProperty('WEBHOOK_SECRET'), 
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
/**
 * Saves environment overrides globally.
 */
function saveSystemSettings(settings) {
  try {
    var props = PropertiesService.getScriptProperties();
    
    // IMMUTABLE ANCHOR: Block the core repository name from being used as a white-label identity
    if (settings.systemName && settings.systemName.trim().toLowerCase() === 'sparkhub') {
      throw new Error("The name 'SparkHub' is restricted. Please provide a custom white-label system name.");
    }
    
    if (settings.mainDbId) validateDatabase(settings.mainDbId);
    if (settings.logsDbId) validateLogsDatabase(settings.logsDbId);

    if (settings.environment) props.setProperty('ENVIRONMENT', settings.environment);
    if (settings.authMode) props.setProperty('AUTH_MODE', settings.authMode);
    if (settings.adminEmail) props.setProperty('ADMIN_EMAIL', settings.adminEmail);
    if (settings.systemName) props.setProperty('SYSTEM_NAME', settings.systemName);
    if (settings.rootFolderId) props.setProperty('ROOT_FOLDER_ID', settings.rootFolderId);
    if (settings.mainDbId) props.setProperty('DATABASE_ID', settings.mainDbId);
    if (settings.logsDbId) props.setProperty('LOGS_DATABASE_ID', settings.logsDbId);
    if (settings.fallbackLogoUrl) props.setProperty('EMAIL_FALLBACK_LOGO', settings.fallbackLogoUrl);

    // Theme Engine Sync
    if (settings.themePrimary) props.setProperty('THEME_PRIMARY', settings.themePrimary);
    if (settings.themeAccent) props.setProperty('THEME_ACCENT', settings.themeAccent);
    if (settings.themeDark) props.setProperty('THEME_DARK', settings.themeDark);
    if (settings.themeHover) props.setProperty('THEME_HOVER', settings.themeHover);
    if (settings.themeBg) props.setProperty('THEME_BG', settings.themeBg);
    
    if (settings.systemLogoId) props.setProperty('SYSTEM_LOGO_ID', settings.systemLogoId);

    SystemEvent.emit("Settings", "UPDATE", "System Configuration", "WARN", "Global Settings", "Core system architecture, identity, or theme settings were modified.");
    return "Success! Settings updated.";

  } catch (e) { 
    return "Error: " + e.message;
  }
}

/**
 * Executes the dynamic installation of external modules, auto-assigns wrappers,
 * seeds module-specific templates, and dispatches global BCC announcements.
 */
function performModuleInstallation(moduleName) {
  try {
    if (!moduleName) return "Error: Module name is required.";
    var mName = moduleName.trim();
    
    // 1. Verify module registry & run DB setup
    var getTriggers = this[mName + "_getTriggers"];
    if (typeof getTriggers !== 'function') return "Error: Could not find backend registry for '" + mName + "'.";
    var setupFunc = this["setup" + mName + "Database"];
    if (typeof setupFunc === 'function') {
      var setupRes = setupFunc();
      if (setupRes && setupRes.error) return "Installation Failed during DB Setup: " + setupRes.error;
    }

    // 2. Register Module
    var props = PropertiesService.getScriptProperties();
    var installed = props.getProperty('INSTALLED_MODULES') || "";
    var modules = installed.split(',').map(function(m){ return m.trim(); }).filter(function(m){ return m !== ""; });
    for (var i = 0; i < modules.length; i++) {
      if (modules[i].toLowerCase() === mName.toLowerCase()) return "Notice: Module '" + mName + "' is already installed.";
    }
    modules.push(mName);
    props.setProperty('INSTALLED_MODULES', modules.join(','));

    // 3. Module-Specific Template Seeding & Auto-Assignment
    var extWrapperName = "External Communication";
    var intWrapperName = "Internal Communication";
    try {
      if (typeof getWrappersList === 'function') {
        var wrappers = getWrappersList().filter(function(w) { return w.status === "Active"; });
        var extWrapper = wrappers.find(function(w) { return w.name.toLowerCase().includes("external") || w.name.toLowerCase().includes("client"); });
        if (extWrapper) extWrapperName = extWrapper.name;
        
        var intWrapper = wrappers.find(function(w) { return w.name.toLowerCase().includes("internal") || w.name.toLowerCase().includes("user"); });
        if (intWrapper) intWrapperName = intWrapper.name;
      }

      // Delegate to the specific external module's template seeder
      var seedFuncModule = this["seed" + mName + "Templates"];
      if (typeof seedFuncModule === 'function') seedFuncModule(extWrapperName, intWrapperName);
    } catch(e) { console.error("Template Seeding Error: " + e.message); }

    // 4. Dispatch the Event and the BCC Announcement
    SystemEvent.emit("Settings", "UPDATE", "Module Installed", "INFO", mName, "The " + mName + " module was successfully provisioned.", "System", {});
    
    try {
      var bccEmails = [];
      if (typeof getUsersList === 'function') {
        var usersRes = getUsersList();
        if (usersRes && usersRes.data) bccEmails = usersRes.data.filter(function(u) { return u.status === 'Active' && u.email; }).map(function(u) { return u.email; });
      }
      if (bccEmails.length > 0) {
        var bccString = bccEmails.join(',');
        var adminEmail = getSystemSettings().adminEmail || bccEmails[0];
        SystemEvent.emit("System", "MODULE_INSTALLED", "Module Installed", "INFO", mName, "The " + mName + " module was successfully installed.", adminEmail, { bcc: bccString, moduleName: mName });
      }
    } catch(e) { console.error("Announcement Error: " + e.message); }

    return "Success: Module '" + mName + "' installed and registered!";
  } catch (e) {
    return "Installation Error: " + e.message;
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
 * Retrieves the published Web App URL to allow safe, cross-origin auto-refreshing.
 */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */