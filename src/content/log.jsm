// Opera Wang, 2010/1/15
// GPL V3 / MPL
// debug utils
"use strict";
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm, stack: Cs } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource://gre/modules/XPCOMUtils.jsm");
Cu.import("resource:///modules/iteratorUtils.jsm"); // import toXPCOMArray
const popupImage = "chrome://awsomeAutoArchive/content/icon_popup.png";
var EXPORTED_SYMBOLS = ["autoArchiveLog"];
let autoArchiveLog = {
  popupDelay: 4,
  setPopupDelay: function(delay) {
    this.popupDelay = delay;
  },
  popupListener: {
    QueryInterface: XPCOMUtils.generateQI([Ci.nsISupports, Ci.nsIObserver]), // not needed, just be safe
    observe: function(subject, topic, cookie) {
      if ( topic == 'alertclickcallback' ) { // or alertfinished / alertshow(Gecko22)
        let type = 'global:console';
        let logWindow = Services.wm.getMostRecentWindow(type);
        if ( logWindow ) return logWindow.focus();
        Services.ww.openWindow(null, 'chrome://global/content/console.xul', type, 'chrome,titlebar,toolbar,centerscreen,resizable,dialog=yes', null);
      } else if ( topic == 'alertfinished' ) {
        delete popupWins[cookie];
      }
    }
  },
  popup: function(title, msg) {
    let delay = this.popupDelay;
    if ( delay <= 0 ) return;
    /*
    http://mdn.beonex.com/en/Working_with_windows_in_chrome_code.html 
    https://developer.mozilla.org/en-US/docs/XPCOM_Interface_Reference/nsIAlertsService
    https://developer.mozilla.org/en-US/Add-ons/Code_snippets/Alerts_and_Notifications
    Before Gecko 22, alert-service won't work with bb4win, use xul instead
    https://bugzilla.mozilla.org/show_bug.cgi?id=782211
    From Gecko 22, nsIAlertsService also use XUL on all platforms and easy to pass args, but difficult to get windows, so hard to change display time
    let alertsService = Cc["@mozilla.org/alerts-service;1"].getService(Ci.nsIAlertsService);
    alertsService.showAlertNotification(popupImage, title, msg, true, cookie, this.popupListener, name);
    */
    let cookie = Date.now();
    let args = [popupImage, title, msg, true, cookie, 0, '', '', null, this.popupListener];
    // win is nsIDOMJSWindow, nsIDOMWindow
    let win = Services.ww.openWindow(null, 'chrome://global/content/alerts/alert.xul', "_blank", 'chrome,titlebar=no,popup=yes',
      // https://alexvincent.us/blog/?p=451
      // https://groups.google.com/forum/#!topic/mozilla.dev.tech.js-engine/NLDZFQJV1dU
      toXPCOMArray(args.map( function(arg) {
        let variant = Cc["@mozilla.org/variant;1"].createInstance(Ci.nsIWritableVariant);
        if ( arg && typeof(arg) == 'object' ) variant.setAsInterface(Ci.nsIObserver, arg); // to pass the listener interface
        else variant.setFromVariant(arg);
        return variant;
      } ), Ci.nsIMutableArray)); // nsIMutableArray can't pass JavaScript Object
    popupWins[cookie] = Cu.getWeakReference(win);
    // sometimes it's too slow to set win.arguments here when the xul window is reused.
    // win.arguments = args;
    let popupLoad = function() {
      win.removeEventListener('load', popupLoad, false);
      if ( win.document ) {
        let alertBox = win.document.getElementById('alertBox');
        if ( alertBox ) alertBox.style.animationDuration = delay + "s";
        let text = win.document.getElementById('alertTextLabel');
        if ( text && win.arguments[3] ) text.classList.add('awsome_auto_archive-popup-clickable');
      }
      win.moveWindowToEnd = function() { // work around https://bugzilla.mozilla.org/show_bug.cgi?id=324570,  Make simultaneous notifications from alerts service work
        let x = win.screen.availLeft + win.screen.availWidth - win.outerWidth;
        let y = win.screen.availTop + win.screen.availHeight - win.outerHeight;
        let windows = Services.wm.getEnumerator('alert:alert');
        while (windows.hasMoreElements()) {
          let alertWindow = windows.getNext();
          if (alertWindow != win && alertWindow.screenY > win.outerHeight) y = Math.min(y, alertWindow.screenY - win.outerHeight);
        }
        let WINDOW_MARGIN = 10; y += -WINDOW_MARGIN; x += -WINDOW_MARGIN;
        win.moveTo(x, y);
      }
    };
    if ( win.document.readyState == "complete" ) popupLoad();
    else win.addEventListener('load', popupLoad, false);
  },
  cleanup: function() {
    try {
      this.info("Log cleanup");
      for ( let cookie in popupWins ) {
        let newwin = popupWins[cookie].get();
        this.info("close window:" + cookie);
        if ( newwin && newwin.document && !newwin.closed ) newwin.close();
      };
      popupWins = {};
      this.info("Log cleanup done");
    } catch(err){}
  },
  
  now: function() { //author: meizz
    let format = "yyyy-MM-dd hh:mm:ss.SSS ";
    let time = new Date();
    let o = {
      "M+" : time.getMonth()+1, //month
      "d+" : time.getDate(),    //day
      "h+" : time.getHours(),   //hour
      "m+" : time.getMinutes(), //minute
      "s+" : time.getSeconds(), //second
      "q+" : Math.floor((time.getMonth()+3)/3),  //quarter
      "S+" : time.getMilliseconds() //millisecond
    }
    
    if(/(y+)/.test(format)) format=format.replace(RegExp.$1,
      (time.getFullYear()+"").substr(4 - RegExp.$1.length));
    for(let k in o)if(new RegExp("("+ k +")").test(format))
      format = format.replace(RegExp.$1,
        RegExp.$1.length==1 ? o[k] :
          ("000"+ o[k]).substr((""+ o[k]).length+3-RegExp.$1.length));
    return format;
  },
  
  verbose: false,
  setVerbose: function(verbose) {
    this.verbose = verbose;
  },

  info: function(msg,popup,force) {
    if (!force && !this.verbose) return;
    this.log(this.now() + msg,popup,true);
  },

  log: function(msg,popup,info) {
    if ( ( typeof(info) != 'undefined' && info ) || !Components || !Cs || !Cs.caller ) {
      Services.console.logStringMessage(msg);
    } else {
      let scriptError = Cc["@mozilla.org/scripterror;1"].createInstance(Ci.nsIScriptError);
      scriptError.init(msg, Cs.caller.filename, Cs.caller.sourceLine, Cs.caller.lineNumber, 0, scriptError.warningFlag, "chrome javascript");
      Services.console.logMessage(scriptError);
    }
    if (popup) {
      if ( typeof(popup) == 'number' ) popup = 'Warning!';
      this.popup(popup,msg);
    }
  },
  
  // from errorUtils.js
  objectTreeAsString: function(o, recurse, compress, level) {
    let s = "";
    let pfx = "";
    let tee = "";
    try {
      if (recurse === undefined)
        recurse = 0;
      if (level === undefined)
        level = 0;
      if (compress === undefined)
        compress = true;
      
      for (let junk = 0; junk < level; junk++)
        pfx += (compress) ? "| " : "|  ";
      
      tee = (compress) ? "+ " : "+- ";
      
      if (typeof(o) != "object") {
        s += pfx + tee + " (" + typeof(o) + ") " + o + "\n";
      }
      else {
        for (let i in o) {
          let t = "";
          try {
            t = typeof(o[i]);
          } catch (err) {
            s += pfx + tee + " (exception) " + err + "\n";
          }
          switch (t) {
            case "function":
              let sfunc = String(o[i]).split("\n");
              if ( typeof(sfunc[2]) != 'undefined' && sfunc[2] == "    [native code]" )
                sfunc = "[native code]";
              else
                sfunc = sfunc.length + " lines";
              s += pfx + tee + i + " (function) " + sfunc + "\n";
              break;
            case "object":
              s += pfx + tee + i + " (object) " + o[i] + "\n";
              if (!compress)
                s += pfx + "|\n";
              if ((i != "parent") && (recurse))
                s += this.objectTreeAsString(o[i], recurse - 1,
                                             compress, level + 1);
              break;
            case "string":
              if (o[i].length > 200)
                s += pfx + tee + i + " (" + t + ") " + o[i].length + " chars\n";
              else
                s += pfx + tee + i + " (" + t + ") '" + o[i] + "'\n";
              break;
            case "":
              break;
            default:
              s += pfx + tee + i + " (" + t + ") " + o[i] + "\n";
          }
          if (!compress)
            s += pfx + "|\n";
        }
      }
    } catch (ex) {
      s += pfx + tee + " (exception) " + ex + "\n";
    }
    s += pfx + "*\n";
    return s;
  },
  
  logObject: function(obj, name, maxDepth, curDepth) {
    this.info(name + ":\n" + this.objectTreeAsString(obj,maxDepth,true));
  },
  
  logException: function(e, popup) {
    let msg = "";
    if ( typeof(e.name) != 'undefined' && typeof(e.message) != 'undefined' ) {
      msg += e.name + ": " + e.message + "\n";
    }
    if ( e.stack ) {
      msg += e.stack;
    }
    if ( e.location ) {
      msg += e.location + "\n";
    }
    if ( msg == '' ){
      msg += " " + e + "\n";
    }
    msg = 'Caught Exception ' + msg;
    let fileName= e.fileName || e.filename || Cs.caller.filename;
    let lineNumber= e.lineNumber || Cs.caller.lineNumber;
    let sourceLine= e.sourceLine || Cs.caller.sourceLine;
    let scriptError = Cc["@mozilla.org/scripterror;1"].createInstance(Ci.nsIScriptError);
    scriptError.init(msg, fileName, sourceLine, lineNumber, e.columnNumber, scriptError.errorFlag, "chrome javascript");
    Services.console.logMessage(scriptError);
    if ( typeof(popup) == 'undefined' || popup ) this.popup("Exception", msg);
  },
  
};
let popupWins = {};
