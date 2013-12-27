// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchive"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource://app/modules/gloda/utils.js");
Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
Cu.import("chrome://awsomeAutoArchive/content/aop.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveService.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveUtil.jsm");

const XULNS = "http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul";
const statusbarIconID = "autoArchive-statusbar-icon";
const popupsetID = "autoArchive-statusbar-popup";
const contextMenuID = "autoArchive-statusbar-contextmenu";
const statusbarIconSrc = 'chrome://awsomeAutoArchive/content/icon.png';
const statusbarIconSrcWait = 'chrome://awsomeAutoArchive/content/icon_wait.png';
const statusbarIconSrcRun = 'chrome://awsomeAutoArchive/content/icon_run.png';

let autoArchive = {
  //strBundle: Services.strings.createBundle('chrome://awsomeAutoArchive/locale/awsome_auto_archive.properties'),
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
          statusbarIcon.setAttribute('src', statusbarIconSrc);
          statusbarIcon.setAttribute('tooltiptext', autoArchiveUtil.Name + " " + autoArchiveUtil.Version);
          statusbarIcon.setAttribute('popup', contextMenuID);
          statusbarIcon.setAttribute('context', contextMenuID);
          status_bar.insertBefore(statusbarIcon, null);
          aWindow._autoarchive.createdElements.push(statusbarIconID);
          aWindow._autoarchive.statusCallback = function(status, detail) {
            if ( status == autoArchiveService.STATUS_SLEEP ) {
              statusbarIcon.setAttribute('src', statusbarIconSrc);
            } else if ( status == autoArchiveService.STATUS_WAITIDLE ) {
              statusbarIcon.setAttribute('src', statusbarIconSrcWait);
            } else if ( status == autoArchiveService.STATUS_RUN ) {
              statusbarIcon.setAttribute('src', statusbarIconSrcRun);
            }
            statusbarIcon.setAttribute('tooltiptext', autoArchiveUtil.Name + " " + autoArchiveUtil.Version + "\n" + detail);
          };
          autoArchiveService.addStatusListener(aWindow._autoarchive.statusCallback);
        }
        
      }
      this.createPopup(aWindow);
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
        if ( aWindow._autoarchive.statusCallback ) autoArchiveService.removeStatusListener(aWindow._autoarchive.statusCallback);
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
    autoArchiveLog.info('autoArchive cleanup done');
  },
  
  createPopup: function(aWindow) {
    let doc = aWindow.document;
    let popupset = doc.createElementNS(XULNS, "popupset");
    popupset.id = popupsetID;
    let menupopup = doc.createElementNS(XULNS, "menupopup");
    menupopup.id = contextMenuID;
    [ ["Option", "chrome://messenger/skin/accountcentral/account-settings.png", function() { aWindow.openDialog("chrome://awsomeAutoArchive/content/autoArchivePrefDialog.xul", "Opt", "chrome,dialog,modal,resizable"); }],
      ["Addon Homepage", "chrome://mozapps/skin/extensions/category-extensions.png", function(){ autoArchiveUtil.loadUseProtocol("https://addons.mozilla.org/en-US/thunderbird/addon/awsomewutoarchive/"); }],
      ["Help", "chrome://global/skin/icons/question-64.png", function(){ autoArchiveUtil.loadUseProtocol("https://github.com/wangvisual/autoarchive/wiki/Help"); }],
      ["Report Bug", "chrome://global/skin/icons/warning-64.png", function(){ autoArchiveUtil.loadUseProtocol("https://github.com/wangvisual/autoarchive/issues"); }],
      //["Donate", "chrome://awsomeAutoArchive/skin/donate.png", function(){ autoArchiveUtil.loadUseProtocol("https://addons.mozilla.org/en-US/thunderbird/addon/awsomewutoarchive/developers"); }],
    ].forEach( function(menu) {
      let item = doc.createElementNS(XULNS, "menuitem");
      item.setAttribute('label', menu[0]);
      item.setAttribute('image', menu[1]);
      item.addEventListener('command', menu[2], false);
      item.setAttribute('class', "menuitem-iconic");
      menupopup.insertBefore(item, null);
    } );
    popupset.insertBefore(menupopup, null);
    doc.documentElement.insertBefore(popupset, null);
    aWindow._autoarchive.createdElements.push(popupsetID);
  },
};
