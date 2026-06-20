// Background Service Worker for JP-VI Translator Extension

// Set side panel behavior to open when clicking the extension icon
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch((error) => console.error(error));
