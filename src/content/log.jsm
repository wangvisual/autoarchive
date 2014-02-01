// Opera Wang, 2010/1/15
// GPL V3 / MPL
// debug utils
"use strict";
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm, stack: Cs } = Components;
Cu.import("resource://gre/modules/Services.jsm");
const popupImage = "chrome://awsomeAutoArchive/content/icon_popup.png";
var EXPORTED_SYMBOLS = ["autoArchiveLog"];
let autoArchiveLog = {
  popupDelay: 4,
  popupWins: [],
  setPopupDelay: function(delay) {
    this.popupDelay = delay;
  },
  popupListener: {
    observe: function(subject, topic, cookie) {
      if ( topic == 'alertclickcallback' ) { // or alertfinished / alertshow
        let type = 'global:console';
        let logWindow = Services.wm.getMostRecentWindow(type);
        if ( logWindow ) return logWindow.focus();
        Services.ww.openWindow(null, 'chrome://global/content/console.xul', type, 'chrome,titlebar,toolbar,centerscreen,resizable,dialog=yes', null);
      } else if ( topic == 'alertfinished' ) {
        let index = cookie.log.popupWins.indexOf(cookie.winRef);
        if ( index >= 0 ) cookie.log.popupWins.splice(index, 1);
      }
    }
  },
  popup: function(title, msg) {
    let delay = this.popupDelay;
    if ( delay <= 0 ) return;
    // alert-service won't work with bb4win, use xul instead
    // http://mdn.beonex.com/en/Working_with_windows_in_chrome_code.html 
    let win = Services.ww.openWindow(null, 'chrome://global/content/alerts/alert.xul', '_alert', 'chrome,titlebar=no,popup=yes', null ); // nsIDOMJSWindow, nsIDOMWindow
    let winRef = Cu.getWeakReference(win);
    // sometimes it's too slow to set here, but mostly should be OK and it's the way suggested in MDN
    win.arguments = [popupImage, title, msg, true, {winRef: winRef, log: this}/*cookie*/, 0, '', '', null, this.popupListener];
    let popupLoad = function() {
      win.removeEventListener('load', popupLoad, false);
      if ( win.document ) {
        let alertBox = win.document.getElementById('alertBox');
        if ( alertBox ) alertBox.style.animationDuration = delay + "s";
        let text = win.document.getElementById('alertTextLabel');
        if ( text && win.arguments[3] ) text.classList.add('awsome_auto_archive-popup-clickable');
      }
    };
    if ( win.document.readyState == "complete" ) popupLoad();
    else win.addEventListener('load', popupLoad, false);
    this.popupWins.push(winRef);
  },
  cleanup: function() {
    try {
      this.info("Log cleanup");
      this.popupWins.forEach( function(winRef) {
        let newwin = winRef.get();
        if ( newwin && newwin.document && !newwin.closed ) newwin.close();
      } );
      delete this.popupWins;
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
