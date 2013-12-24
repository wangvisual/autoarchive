// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchive"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource://app/modules/gloda/utils.js");
Cu.import("resource://gre/modules/FileUtils.jsm");
//Cu.import("resource://gre/modules/Dict.jsm");
//Cu.import("chrome://awsomeAutoArchive/content/autoArchiveUtil.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
Cu.import("chrome://awsomeAutoArchive/content/aop.jsm");
//Cu.import("chrome://awsomeAutoArchive/content/sprintf.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveService.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");

const XULNS = "http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul";
const statusbarIconID = "autoArchive-statusbar-icon";

let autoArchive = {
  //strBundle: Services.strings.createBundle('chrome://awsomeAutoArchive/locale/awsome_auto_archive.properties'),
  createPopup: function(aWindow) {
  },
  Load: function(aWindow) {
    return autoArchive.realLoad(aWindow);
  },

  realLoad: function(aWindow) {
    try {
      autoArchiveLog.info("Load for " + aWindow.location.href);
      let doc = aWindow.document;
      //let winref = Cu.getWeakReference(aWindow);
      //let docref = Cu.getWeakReference(doc);
      if ( typeof(aWindow._autoarchive) != 'undefined' ) autoArchiveLog.info("Already loaded, return");
      aWindow._autoarchive = { createdElements:[], hookedFunctions:[] };
      if ( typeof(aWindow.MessageDisplayWidget) != 'undefined' || 1 ) { // messeage display window
        let status_bar = doc.getElementById('status-bar');
        if ( status_bar ) { // add status bar icon
          let statusbarIcon = doc.createElementNS(XULNS, "statusbarpanel");
          statusbarIcon.id = statusbarIconID;
          statusbarIcon.setAttribute('class', 'statusbarpanel-iconic');
          statusbarIcon.setAttribute('src', 'chrome://awsomeAutoArchive/content/icon.png');
          statusbarIcon.setAttribute('tooltiptext', 'statusbarTooltipID');
          //statusbarIcon.setAttribute('tooltip', statusbarTooltipID);
          //statusbarIcon.setAttribute('popup', contextMenuID);
          //statusbarIcon.setAttribute('context', contextMenuID);
          status_bar.insertBefore(statusbarIcon, null);
          aWindow._autoarchive.createdElements.push(statusbarIconID);
          autoArchiveLog.info("statusbarIcon");
        }
      }
      if ( aWindow._autoarchive.hookedFunctions.length ) {
        autoArchiveLog.info('create popup');
        //this.createPopup(aWindow);
      }
      aWindow.addEventListener("unload", autoArchive.onUnLoad, false);
    }catch(err) {
      autoArchiveLog.logException(err);
    }
  },
 
  onUnLoad: function(event) {
    autoArchiveLog.info('onUnLoad');
    let aWindow = event.currentTarget;
    if ( aWindow ) autoArchive.unLoad(aWindow);
  },

  unLoad: function(aWindow) {
    try {
      autoArchiveLog.info('unload');
      if ( typeof(aWindow._autoarchive) != 'undefined' ) {
        autoArchiveLog.info('unhook');
        aWindow.removeEventListener("unload", autoArchive.onUnLoad, false);
        aWindow._autoarchive.hookedFunctions.forEach( function(hooked) {
          hooked.unweave();
        } );
        let doc = aWindow.document;
        for ( let node of aWindow._autoarchive.createdElements ) {
          if ( typeof(node) == 'string' ) node = doc.getElementById(node);
          if ( node && node.parentNode ) {
            autoArchiveLog.info("removed node " + node);
            node.parentNode.removeChild(node);
          }
        }
        delete aWindow._autoarchive;
      }
    } catch (err) {
      autoArchiveLog.logException(err);  
    }
    autoArchiveLog.info('unload done');
  },

  cleanup: function() {
    try {
      autoArchiveLog.info('autoArchive cleanup');
      if ( this.timer ) this.timer.cancel();
      this.timer = null;
      autoArchiveService.cleanup();
      autoArchivePref.cleanup();
    } catch (err) {
      autoArchiveLog.logException(err);  
    }
    Cu.unload("chrome://awsomeAutoArchive/content/aop.jsm");
    Cu.unload("chrome://awsomeAutoArchive/content/autoArchiveService.jsm");
    autoArchiveLog.info('autoArchive cleanup done');
    //autoArchiveaop = null;
  },
};
