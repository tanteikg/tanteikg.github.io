// ================================
// LEGACY WEBVIEW GRACEFUL FAILOVER
// ================================
if (navigator.userAgent.includes("Trident")) {
  /*
    Trident is the webview in use. Do one of the following:
    1. Provide an alternate add-in experience that doesn't use any of the
      HTML5 features that aren't supported in Trident (IE 11).
    2. Enable the add-in to gracefully fail by adding a message to the UI
      that says something similar to:
      "This add-in won't run in your version of Office. Please upgrade
      either to perpetual Office 2021 (or later) or to a Microsoft 365
      account."
  */
  let legacyMessage = document.getElementById("legacyMessage");
  let mainUI = document.getElementById("main");
  legacyMessage.style.display = "block";
  mainUI.style.display = "none";
} else if (navigator.userAgent.includes("Edge")) {
  /*
    EdgeHTML is the browser in use. Do one of the following:
    1. Provide an alternate add-in experience that's supported in EdgeHTML
      (Microsoft Edge Legacy).
    2. Enable the add-in to gracefully fail by adding a message to the UI
      that says something similar to:
      "This add-in won't run in your version of Office. Please upgrade
      either to perpetual Office 2021 (or later) or to a Microsoft 365
      account."
  */
  let legacyMessage = document.getElementById("legacyMessage");
  let mainUI = document.getElementById("main");
  legacyMessage.style.display = "block";
  mainUI.style.display = "none";
} else {
  /* 
    A webview other than Trident or EdgeHTML is in use.
    Provide a full-featured version of the add-in here.
  */
}
