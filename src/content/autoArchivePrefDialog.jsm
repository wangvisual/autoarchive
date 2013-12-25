// MPL/GPL
// Opera.Wang 2013/06/04
"use strict";
var EXPORTED_SYMBOLS = ["autoArchivePrefDialog"];

const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
//Cu.import("resource:///modules/mailServices.js");
//Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("resource:///modules/folderUtils.jsm");
Cu.import("resource:///modules/MailUtils.js");
Cu.import("resource://gre/modules/AddonManager.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
const SEAMONKEY_ID = "{92650c4d-4b8e-4d2a-b7eb-24ecf4f6b63a}";
const XUL = "http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul";
const ruleClass = 'awsome_auto_archive-rule';

let autoArchivePrefDialog = {
  Name: "Awesome Auto Archive xxx", // might get changed by getAddonByID function call
  Version: 'unknown',
  isSeaMonkey: Services.appinfo.ID == SEAMONKEY_ID,
  Applicaton: ( Services.appinfo.ID == SEAMONKEY_ID ) ? 'Seamonky' : 'Thunderbird',
  initName: function() {
    autoArchiveLog.log("autoArchivePrefDialog initName");
    if ( this.Version != 'unknown' ) return;
    AddonManager.getAddonByID('awsomeautoarchive@opera.wang', function(addon) {
      if ( !self ) return;
      self.Version = addon.version;
      self.Name = addon.name;
    });
  },
  loadInTopWindow: function(win, url) {
    win.openDialog("chrome://messenger/content/", "_blank", "chrome,dialog=no,all", null,
      { tabType: "contentTab", tabParams: {contentPage: Services.io.newURI(url, null, null) } });
  },
  loadUseProtocol: function(url) {
    try {
      Cc["@mozilla.org/uriloader/external-protocol-service;1"].getService(Ci.nsIExternalProtocolService).loadURI(Services.io.newURI(url, null, null), null);
    } catch (err) {
      autoArchiveLog.logException(err);
    }
  },
  loadDonate: function(pay) {
    let url = "https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=893LVBYFXCUP4&lc=US&item_name=Expression%20Search&no_note=0&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_LG%2egif%3aNonHostedGuest";
    if ( typeof(pay) != 'undefined' ) {
      if ( pay == 'alipay' ) url = "https://me.alipay.com/operawang";
      if ( pay == 'mozilla' ) url = "https://addons.mozilla.org/en-US/thunderbird/addon/ldapinfoshow/developers?src=api"; // Meet the developer page
    }
    this.loadUseProtocol(url);
  },
  sendEmailWithTB: function(url) {
    MailServices.compose.OpenComposeWindowWithURI(null, Services.io.newURI(url, null, null));
  },
  loadTab: function(args) {
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    if (mail3PaneWindow) {
      let tabmail = mail3PaneWindow.document.getElementById("tabmail");
      if ( !tabmail ) return;
      mail3PaneWindow.focus();
      tabmail.openTab(args.type, args);
    }
  },
  onSyncFromPreference: function(doc,self) {
    let textbox = self;
    let preference = doc.getElementById('facebook_token_expire');
    let actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
    let date = new Date((+actualValue)*1000);
    return date.toLocaleFormat("%Y/%m/%d %H:%M:%S");
  },
  cleanup: function() {
  },
  
  showPrettyTooltip: function(URI,pretty) {
    return decodeURIComponent(URI).replace(/(.*\/)[^/]*/, '$1') + pretty;
  },
  updateFolderStyle: function(folderPicker, folderPopup, win) {
    let msgFolder = {value: '', prettyName: 'N/A'};
    let updateStyle = function() {
      folderPopup.selectFolder(msgFolder); // false alarm by addon validator
      folderPopup._setCssSelectors(msgFolder, folderPicker);
    };
    try {
      msgFolder = MailUtils.getFolderForURI(folderPicker.value);
      if ( win ) win.setTimeout( updateStyle, 1 );// use timer to wait for the XBL bindings add SelectFolder / _setCssSelectors to popup
      else updateStyle();
    } catch(err) {}
    folderPicker.setAttribute("label", msgFolder.prettyName);
    folderPicker.setAttribute('tooltiptext', self.showPrettyTooltip(msgFolder.ValueUTF8||msgFolder.value, msgFolder.prettyName));
  },
  onFolderPick: function(folderPicker, aEvent, folderPopup) {
    let folder = aEvent.target._folder;
    if ( !folder ) return;
    let value = folder.URI || folder.folderURL;
    folderPicker.value = value; // must set value before set label, or next line may fail when previous value is empty
    self.updateFolderStyle(folderPicker, folderPopup, null);
  },
  initFolderPick: function(doc, win, folderPicker, folderPopup) {
    folderPicker.addEventListener('command', function(aEvent) { return self.onFolderPick(folderPicker, aEvent, folderPopup); }, false);
    folderPicker.classList.add("folderMenuItem");

    folderPopup.setAttribute("type", "folder");
    folderPopup.setAttribute("mode", "newFolder");
    folderPopup.setAttribute("showFileHereLabel", "true");
    folderPopup.setAttribute("fileHereLabel", "here");
    folderPopup.classList.add("menulist-menupopup");
    folderPopup.classList.add("searchPopup");

    let menuitem = doc.createElementNS(XUL, "menuitem");
    menuitem.setAttribute("label", "N/A");
    menuitem.setAttribute("value", "");
    menuitem.setAttribute("class", "folderMenuItem");
    menuitem.setAttribute("SpecialFolder", "Virtual");
    folderPopup.insertBefore(menuitem, null);
    let menuseparator = doc.createElementNS(XUL, "menuseparator");
    folderPopup.insertBefore(menuseparator, null);
    self.updateFolderStyle(folderPicker, folderPopup, win);
  },
  
  creatNewRule: function(win) {
    return self.creatOneRule(win.document, win, {action: 'archive', enable: true, sub: 0, age: autoArchivePref.options.default_days}, null);
  },

  creatOneRule: function(doc, win, rule, ref) {
    try {
      let group = doc.getElementById('awsome_auto_archive-rules');
      if ( !group ) return;
      let enable = doc.createElementNS(XUL, "checkbox");
      enable.setAttribute("checked", rule.enable);
      enable.setAttribute("rule", 'enable');
      
      let menulistAction = doc.createElementNS(XUL, "menulist");
      let menupopupAction = doc.createElementNS(XUL, "menupopup");
      ["archive", "copy", "delete", "move"].forEach( function(action) {
        let menuitem = doc.createElementNS(XUL, "menuitem");
        menuitem.setAttribute("label", action);
        menuitem.setAttribute("value", action);
        menupopupAction.insertBefore(menuitem, null);
      } );
      menulistAction.insertBefore(menupopupAction, null);
      menulistAction.setAttribute("value", rule.action || 'archive');
      menulistAction.setAttribute("rule", 'action');
      
      let menulistSrc = doc.createElementNS(XUL, "menulist");
      let menupopupSrc = doc.createElementNS(XUL, "menupopup");
      menulistSrc.insertBefore(menupopupSrc, null);
      menulistSrc.value = rule.src || '';
      menulistSrc.setAttribute("rule", 'src');
      
      let menulistSub = doc.createElementNS(XUL, "menulist");
      let menupopupSub = doc.createElementNS(XUL, "menupopup");
      let types = [ {key: "only", value: 0}, { key: "sub", value: 1}, {key: "sub_keep", value: 2} ];
      types.forEach( function(type) {
        let menuitem = doc.createElementNS(XUL, "menuitem");
        menuitem.setAttribute("label", type.key);
        menuitem.setAttribute("value", type.value);
        menupopupSub.insertBefore(menuitem, null);
      } );
      menulistSub.insertBefore(menupopupSub, null);
      menulistSub.setAttribute("value", rule.sub || 0);
      menulistSub.setAttribute("rule", 'sub');
      
      let to = doc.createElementNS(XUL, "label");
      to.setAttribute("value", "To");
      let menulistDest = doc.createElementNS(XUL, "menulist");
      let menupopupDest = doc.createElementNS(XUL, "menupopup");
      menulistDest.insertBefore(menupopupDest, null);
      menulistDest.value = rule.dest || '';
      menulistDest.setAttribute("rule", 'dest');
      
      let matches = doc.createElementNS(XUL, "label");
      matches.setAttribute("value", "matches");
      let subject = doc.createElementNS(XUL, "textbox");
      subject.setAttribute("value", rule.subject || '');
      subject.setAttribute("rule", 'subject');
      
      let after = doc.createElementNS(XUL, "label");
      after.setAttribute("value", "After");
      let age = doc.createElementNS(XUL, "textbox");
      age.setAttribute("type", "number");
      age.setAttribute("min", "0");
      age.setAttribute("value", typeof(rule.age)!='undefined' ? rule.age : autoArchivePref.options.default_days);
      age.setAttribute("rule", 'age');
      age.setAttribute("size", "4");
      let days = doc.createElementNS(XUL, "label");
      days.setAttribute("value", "days");
      
      let up = doc.createElementNS(XUL, "toolbarbutton");
      up.setAttribute("label", '\u2191');
      up.addEventListener("command", function(aEvent) { self.upDownRule(hbox, true, doc, win); }, false );
      
      let down = doc.createElementNS(XUL, "toolbarbutton");
      down.setAttribute("label", '\u2193');
      down.addEventListener("command", function(aEvent) { self.upDownRule(hbox, false, doc, win); }, false );
      
      let remove = doc.createElementNS(XUL, "toolbarbutton");
      remove.setAttribute("label", "x");
      remove.setAttribute("icon", "remove");
      remove.classList.add("awsome_auto_archive-delete-rule");
      remove.addEventListener("command", function(aEvent) { self.removeRule(hbox); }, false );
      
      let hbox = doc.createElementNS(XUL, "hbox");
      hbox.classList.add(ruleClass);
      [enable, menulistAction, menulistSrc, menulistSub, to, menulistDest, matches, subject, after, age, days, up, down, remove].forEach( function(item) {
        hbox.insertBefore(item, null);
      } );
      group.insertBefore(hbox, ref);
      self.initFolderPick(doc, win, menulistSrc, menupopupSrc);
      self.initFolderPick(doc, win, menulistDest, menupopupDest);
      self.checkAction(menulistAction, to, menulistDest);
      self.checkEnable(enable, hbox);
      menulistAction.addEventListener('command', function(aEvent) { self.checkAction(menulistAction, to, menulistDest); } );
      enable.addEventListener('command', function(aEvent) { self.checkEnable(enable, hbox); } );
    } catch(err) {
      autoArchiveLog.logException(err);
    }
    return true;
  },
  
  upDownRule: function(hbox, isUp, doc, win) {
    try {
      let ref = isUp ? hbox.previousSibling : hbox;
      let remove = isUp ? hbox : hbox.nextSibling;
      if ( ref.classList.contains(ruleClass) &&  remove.classList.contains(ruleClass) ) {
        let rule = this.getOneRule(remove);
        autoArchiveLog.logObject(rule, 'temp rule', 1);
        remove.parentNode.removeChild(remove);
        // remove.parentNode.insertBefore(remove, ref); // lost all unsaved values
        this.creatOneRule(doc, win, rule, ref);
      }
    } catch(err) {
      autoArchiveLog.logException(err);
    }
  },
  
  removeRule: function(hbox) {
    hbox.parentNode.removeChild(hbox);
  },

  checkEnable: function(enable, hbox) {
    if ( enable.checked ) {
      hbox.classList.remove("awsome_auto_archive-disable");
    } else {
      hbox.classList.add("awsome_auto_archive-disable");
    }
  },
  
  checkAction: function(menulistAction, to, menulistDest) {
    return to.style.visibility = menulistDest.style.visibility = ["archive", "delete"].indexOf(menulistAction.value) >= 0 ? 'hidden': 'visible';
  },

  loadPerfWindow: function(win) {
    try {
      let doc = win.document;
      if ( autoArchivePref.rules.length ) {
        autoArchivePref.rules.forEach( function(rule) {
          self.creatOneRule(doc, win, rule, null);
        } );
      } else {
        self.creatNewRule(win);
      }
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },
  getOneRule: function(hbox) {
    let rule = {};
    for ( let item of hbox.childNodes ) {
      let key = item.getAttribute('rule');
      if ( key ) {
        let value = item.value || item.checked;
        if ( item.getAttribute("type") == 'number' ) value = item.valueNumber;
        rule[key] = value;
      }
    }
    return rule;
  },
  acceptPerfWindow: function(win) {
    try {
      let doc = win.document;
      let group = doc.getElementById('awsome_auto_archive-rules');
      if ( !group ) return;
      let rules = [];
      for ( let hbox of group.childNodes ) {
        if ( hbox.classList.contains(ruleClass) ) {
          let rule = this.getOneRule(hbox);
          if ( Object.keys(rule).length > 0 ) rules.push(rule);
        }
      }
      autoArchiveLog.logObject(rules,'new rules',1);
      autoArchivePref.setPerf('rules',JSON.stringify(rules));
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },

}
let self = autoArchivePrefDialog;
self.initName();
