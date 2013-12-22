// Opera.Wang+autoArchive@gmail.com GPL/MPL
"use strict";

const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
// if use custom resouce, refer here
// http://mdn.beonex.com/en/JavaScript_code_modules/Using.html

const sss = Cc["@mozilla.org/content/style-sheet-service;1"].getService(Ci.nsIStyleSheetService);
const userCSS = Services.io.newURI("chrome://awsomeAutoArchive/content/autoArchive.css", null, null);
const targetWindows = [ "mail:3pane" ];

function loadIntoWindow(window) {
  if ( !window ) return; // windows is the global host context
  let document = window.document; // XULDocument
  let type = document.documentElement.getAttribute('windowtype'); // documentElement maybe 'messengerWindow' / 'addressbookWindow'
  if ( targetWindows.indexOf(type) < 0 ) return;
  autoArchive.Load(window);
}

var windowListener = {
  onOpenWindow: function(aWindow) {
    let onLoadWindow = function() {
      aWindow.removeEventListener("load", onLoadWindow, false);
      let msgComposeWindow = aWindow.document.getElementById("msgcomposeWindow");
      if ( msgComposeWindow ) msgComposeWindow.removeEventListener("compose-window-reopen", onLoadWindow, false);
      loadIntoWindow(aWindow);
    };
    aWindow.addEventListener("load", onLoadWindow, false);
    let msgComposeWindow = aWindow.document.getElementById("msgcomposeWindow");
    if ( msgComposeWindow ) {
      if ( aWindow.ComposeFieldsReady ) loadIntoWindow(aWindow); // compose-window-reopen won't work for TB24 2013.06.10?
      msgComposeWindow.addEventListener("compose-window-reopen", onLoadWindow, false);
    }
  },
  //onCloseWindow: function(aWindow) {}, onWindowTitleChange: function(aWindow) {},
  observe: function(subject, topic, data) {
    if ( topic == "xul-window-registered") {
      windowListener.onOpenWindow( subject.QueryInterface(Ci.nsIInterfaceRequestor).getInterface(Ci.nsIDOMWindow) );
    }
  },
};

// A toplevel window in a XUL app is an nsXULWindow.  Inside that there is an nsGlobalWindow (aka nsIDOMWindow).
function startup(aData, aReason) {
  Services.console.logStringMessage("Awesome Auto Archive startup...");
  Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
  Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
  autoArchivePref.initPerf(__SCRIPT_URI_SPEC__);
  Cu.import("chrome://awsomeAutoArchive/content/autoArchive.jsm");
  Cu.import("chrome://awsomeAutoArchive/content/autoArchivePrefDialog.jsm");
  //autoArchiveUtil.setChangeCallback( function(clean) { autoArchive.clearCache(clean); } );
  // Load into any existing windows, but not hidden/cached compose window, until compose window recycling is disabled by bug https://bugzilla.mozilla.org/show_bug.cgi?id=777732
  let windows = Services.wm.getEnumerator(null);
  while (windows.hasMoreElements()) {
    let domWindow = windows.getNext().QueryInterface(Ci.nsIDOMWindow);
    if ( domWindow.document.readyState == "complete" && targetWindows.indexOf(domWindow.document.documentElement.getAttribute('windowtype')) >= 0 ) {
      loadIntoWindow(domWindow);
    } else {
      windowListener.onOpenWindow(domWindow);
    }
  }
  // Wait for new windows
  Services.obs.addObserver(windowListener, "xul-window-registered", false);
  // validator warnings on the below line, ignore it
  if ( !sss.sheetRegistered(userCSS, sss.USER_SHEET) ) sss.loadAndRegisterSheet(userCSS, sss.USER_SHEET); // will be unregister when shutdown
}
 
function shutdown(aData, aReason) {
  // When the application is shutting down we normally don't have to clean up any UI changes made
  // but we have to abort running jobs
  if ( sss.sheetRegistered(userCSS, sss.USER_SHEET) ) sss.unregisterSheet(userCSS, sss.USER_SHEET);
  try {
    Services.obs.removeObserver(windowListener, "xul-window-registered");
    // Unload from any existing windows
    let windows = Services.wm.getEnumerator(null);
    while (windows.hasMoreElements()) {
      let domWindow = windows.getNext().QueryInterface(Ci.nsIDOMWindow);
      autoArchive.unLoad(domWindow); // won't check windowtype as unload will check
      // Do CC & GC, comment out allTraces when release
      domWindow.QueryInterface(Ci.nsIInterfaceRequestor).getInterface(Ci.nsIDOMWindowUtils).garbageCollect(
        // Cc["@mozilla.org/cycle-collector-logger;1"].createInstance(Ci.nsICycleCollectorListener).allTraces()
      );
    }
    autoArchive.cleanup();
  } catch (err) {Cu.reportError(err);}
  if (aReason == APP_SHUTDOWN) return;
  Services.strings.flushBundles(); // clear string bundles
  Cu.unload("chrome://awsomeAutoArchive/content/autoArchivePrefDialog.jsm");
  Cu.unload("chrome://awsomeAutoArchive/content/autoArchive.jsm");
  Cu.unload("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
  Cu.unload("chrome://awsomeAutoArchive/content/log.jsm");
  try {
    autoArchive = autoArchivePref = autoArchiveLog = null;
  } catch (err) {}
  // flushStartupCache
  // Init this, so it will get the notification.
  //Cc["@mozilla.org/xul/xul-prototype-cache;1"].getService(Ci.nsISupports);
  Services.obs.notifyObservers(null, "startupcache-invalidate", null);
  Cu.schedulePreciseGC( Cu.forceGC );
  Services.console.logStringMessage("Awesome Auto Archive shutdown");
}

function install(aData, aReason) {}
function uninstall(aData, aReason) {}
