/**
 * Main Initialization Module (Router)
 * Standardized under SparkHub Architecture Blueprint.
 * * CORE RESPONSIBILITIES:
 * - Entry point (doGet) for the Web Application.
 * - Bridges the Module/Plugin Registry from ScriptProperties to the UI.
 * - Routes to the Installation Wizard or the Modular Dashboard.
 */

/**
 * Entry point for the SparkHub Web Application.
 * Handles environment detection, dynamic branding, and registry-based rendering.
 * @return {HtmlService.HtmlOutput} The evaluated HTML template.
 */
function doGet() {
  var props = PropertiesService.getScriptProperties();
  var env = props.getProperty('ENVIRONMENT');
  var userEmail = Session.getActiveUser().getEmail();
  
  // 1. INSTALLATION CHECK
  var isInstalled = (env !== null && env !== "");
  
  // 2. FETCH SYSTEM SETTINGS (Registry + Theme)
  var settings = getSystemSettings();
  
  // 3. APP INITIALIZATION
  var template = HtmlService.createTemplateFromFile('Index');
  template.isInstalled = isInstalled;
  template.userEmail = userEmail;
  template.systemName = settings.systemName;
  
  // Pass Registry data to allow Index.html to perform conditional rendering
  template.installedModules = settings.installedModules;
  template.installedPlugins = settings.installedPlugins;
  
  // Hybrid Logo Logic
  template.systemLogoUrl = settings.systemLogoId ? settings.systemLogoUrl : settings.appFallbackLogo;
  template.appFallbackLogo = settings.appFallbackLogo;
  
  // Pass Theme Engine variables
  template.themePrimary = settings.themePrimary;
  template.themeAccent = settings.themeAccent;
  template.themeDark = settings.themeDark;
  template.themeBg = settings.themeBg;
  template.themeHover = settings.themeHover;

  // 4. SECURITY & PERMISSIONS
  if (isInstalled) {
    var role = 'Guest';
    
    // EXPLICIT AUTHENTICATION OVERRIDE
    if (settings.authMode === 'Local') {
      role = 'Guest'; // Defer entirely to frontend localStorage
    } else {
      // SSO Mode
      if (userEmail !== '') {
        role = getUserRole();
      }
    }

    if (role === 'Inactive') {
      return serveAccessDeniedScreen(settings);
    }
    
    template.userRole = role;
    template.username = (role === 'Guest') ? '' : getLoggedInUsername();
    template.userPermissions = (role === 'Guest') ? '{}' : getUserPermissions(role);
  } else {
    template.userRole = "Administrator";
    template.username = userEmail.split('@')[0];
    template.userPermissions = '{"ALL":["ALL"]}';
  }
  
  var htmlOutput = template.evaluate()
      .setTitle(settings.systemName)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  // Set Dynamic Favicon
  if (settings.systemLogoId) {
    htmlOutput.setFaviconUrl(settings.systemLogoUrl + "&ext=.png");
  }
  else {
    // Ensure this points to a standard URL (e.g., .png), not a Data URI/SVG
    htmlOutput.setFaviconUrl(settings.appFallbackLogo);
  }
  
  return htmlOutput;
}

/**
 * Serves a branded HTML error page when a user's account is deactivated.
 */
function serveAccessDeniedScreen(s) {
  var displayLogo = s.systemLogoId ? s.systemLogoUrl : s.appFallbackLogo;
  
  var errorHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${s.systemName} - Access Denied</title>
      <style>
        body { background-color: ${s.themeDark}; font-family: 'Inter', sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; color: white; }
        .error-card { background: #FFFFFF; padding: 48px; border-radius: 24px; box-shadow: 0 20px 50px rgba(0,0,0,0.3); text-align: center; max-width: 420px; border-top: 8px solid ${s.themePrimary}; color: #0F172A; }
        .error-card h1 { margin-top: 0; font-size: 24px; font-weight: 800; }
        .error-card p { opacity: 0.7; line-height: 1.6; font-size: 15px; }
      </style>
    </head>
    <body>
      <div class="error-card">
        <img src="${displayLogo}" alt="${s.systemName} Logo" style="height: 64px; margin-bottom: 24px; object-fit: contain;">
        <h1>Access Denied</h1>
        <p>Your account has been deactivated. You do not have permission to access ${s.systemName}. Please contact your administrator.</p>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(errorHtml)
      .setTitle(s.systemName + ' - Access Denied')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl(s.systemLogoId ? s.systemLogoUrl + "&ext=.png" : s.appFallbackLogo);
}

/**
 * [SPARKHUB INTEGRITY ANCHOR: END]
 */