/**
 * [SPARKHUB INTEGRITY HEADER: START]
 * FILE: Main.gs
 * VERSION: 1.9 (True Modular Injection + Core Logic Restoration)
 * SYNC STATUS: Fully Synchronized with Index.html & Core Modules
 */

/**
 * Main Initialization Module (Router)
 * Standardized under SparkHub Architecture Blueprint.
 * * CORE RESPONSIBILITIES:
 * - Entry point (doGet) for the Web Application.
 * - Bridges the Module/Plugin Registry from ScriptProperties to the UI.
 * - Routes to the Installation Wizard or the Modular Dashboard.
 */
function doGet(e) {
  // --- COMMUNICATION INTERCEPTOR: LINK TRACKING ---
  if (e && e.parameter && e.parameter.action === 'trackEvent' && e.parameter.dest) {
    var destUrl = decodeURIComponent(e.parameter.dest);
    var userEmail = e.parameter.u || "Unknown User";
    
    // Log the click event to the global audit trail
    SystemEvent.emit("Templates", "TRACKING", "Email Link Clicked", "INFO", userEmail, "Navigated to: " + destUrl);
    
    // 1. Fetch the primary brand color
    var props = PropertiesService.getScriptProperties();
    var themeStr = props.getProperty('SYSTEM_THEME');
    var primaryColor = '#2563eb'; // Default SparkHub Blue fallback
    if (themeStr) {
      try {
        var theme = JSON.parse(themeStr);
        if (theme.primary) primaryColor = theme.primary;
      } catch(err) { /* ignore parse error */ }
    }
    
    // 2. Inject the dynamic brand color into the redirect button CSS
    var redirectHtml = '<!DOCTYPE html><html><head><style>' +
      'body { font-family: sans-serif; text-align: center; padding-top: 60px; color: #64748b; }' +
      '.btn { display: inline-block; margin-top: 25px; padding: 12px 24px; background-color: ' + primaryColor + '; color: #ffffff; text-decoration: none; border-radius: 6px; font-weight: 600; transition: opacity 0.2s; }' +
      '.btn:hover { opacity: 0.9; }' +
      '</style></head><body>' +
      '<h2>Redirecting...</h2>' +
      '<p>You are being securely routed to your destination.</p>' +
      '<a href="' + destUrl + '" target="_top" class="btn">Continue Now</a>' +
      '<script>' +
      '  setTimeout(function() { ' +
      '    try { window.top.location.replace("' + destUrl + '"); } catch(err) {} ' +
      '  }, 800);' +
      '</script></body></html>';
      
    return HtmlService.createHtmlOutput(redirectHtml);
  }
  // ------------------------------------------------
  
  var props = PropertiesService.getScriptProperties();
  var env = props.getProperty('ENVIRONMENT');
  var userEmail = Session.getActiveUser().getEmail();
  var isInstalled = (env !== null && env !== "");
  var settings = getSystemSettings();
  
  var template = HtmlService.createTemplateFromFile('Index');
  template.isInstalled = isInstalled;
  template.userEmail = userEmail;
  template.systemName = settings.systemName;
  template.resetToken = (e && e.parameter && e.parameter.token) ? e.parameter.token : "";
  
  // DYNAMIC UI REGISTRY
  var includes = ['SettingsData', 'UsersData', 'TemplatesData', 'LogsData'];
  if (settings.installedModules) {
    settings.installedModules.split(',').forEach(function(m) {
      var fileName = m.trim() + "Data";
      if (includes.indexOf(fileName) === -1) includes.push(fileName);
    });
  }
  template.includeList = includes; 

  template.installedModules = settings.installedModules;
  template.installedPlugins = settings.installedPlugins;
  template.systemLogoUrl = settings.systemLogoId ? settings.systemLogoUrl : settings.appFallbackLogo;
  template.appFallbackLogo = settings.appFallbackLogo;
  template.themePrimary = settings.themePrimary;
  template.themeAccent = settings.themeAccent;
  template.themeDark = settings.themeDark;
  template.themeBg = settings.themeBg;
  template.themeHover = settings.themeHover;
  template.authMode = settings.authMode;

  var role = 'Guest';
  var layout = 'DEFAULT'; 
  if (isInstalled) {
    if (settings.authMode === 'Local') {
      role = 'Guest';
    } else {
      role = (userEmail === '' ? 'Guest' : getUserRole());
      if (role !== 'Guest' && role !== 'Inactive') layout = getResolvedDashboardLayout(userEmail, role);
    }
  } else {
    role = 'Administrator';
  }
  
  if (role === 'Inactive') return serveAccessDeniedScreen(settings);

  template.userRole = role;
  template.dashboardLayout = layout;
  template.username = (role === 'Guest') ? '' : getLoggedInUsername();
  template.userFirstName = (role === 'Guest' || !isInstalled) ? '' : getLoggedInUserFirstName();
  template.userPermissions = (role === 'Guest' || !isInstalled) ? '{"ALL":["ALL"]}' : getUserPermissions(role);

  return template.evaluate()
      .setTitle(settings.systemName)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl(settings.systemLogoId ? settings.systemLogoUrl + "&ext=.png" : settings.appFallbackLogo);
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