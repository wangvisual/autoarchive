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

let autoArchivePrefDialog = {
  Name: "Awesome Auto Archive xxx", // might get changed by getAddonByID function call
  Version: 'unknown',
  isSeaMonkey: Services.appinfo.ID == SEAMONKEY_ID,
  Applicaton: ( Services.appinfo.ID == SEAMONKEY_ID ) ? 'Seamonky' : 'Thunderbird',
  initName: function() {
    autoArchiveLog.log("autoArchivePrefDialog initName");
    if ( this.Version != 'unknown' ) return;
    AddonManager.getAddonByID('awsomeautoarchive@opera.wang', function(addon) {
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
  onFolderPick: function(folderPicker, aEvent) {
    let folder = aEvent.target._folder;
    if ( !folder ) return;
    let label = folder.prettyName || folder.name;
    let value = folder.URI || folder.folderURL;
    folderPicker.value = value; // must set value before set label, or next line may fail when previous value is empty
    folderPicker.setAttribute("label", label); 
    folderPicker.setAttribute('tooltiptext', self.showPrettyTooltip(value, label));
  },
  initFolderPick: function(doc, win, folderPicker, folderPopup) {
    folderPicker.addEventListener('command', function(aEvent) { return self.onFolderPick(folderPicker, aEvent); }, false);
    folderPicker.classList.add("folderMenuItem");

    folderPopup.setAttribute("type", "folder");
    folderPopup.setAttribute("maxwidth", "300");
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
    
    let msgFolder = {value: '', prettyName: 'N/A'};
    try {
      msgFolder = MailUtils.getFolderForURI(folderPicker.value);
      win.setTimeout( function() { // use timer to wait for the XBL bindings add selectFolder / _setCssSelectors to popup
        try {
          folderPopup.selectFolder(msgFolder);
          folderPopup._setCssSelectors(msgFolder, folderPicker);
        } catch(err) {}
      }, 1 );
    } catch(err) {}
    folderPicker.setAttribute("label", msgFolder.prettyName);
    folderPicker.setAttribute('tooltiptext', self.showPrettyTooltip(msgFolder.ValueUTF8||msgFolder.value, msgFolder.prettyName));
  },
  
  creatOneRule: function(doc, win, rule, parent) {
    let checkbox = doc.createElementNS(XUL, "checkbox");
    checkbox.setAttribute("checked", rule.enable);

    let menulistAction = doc.createElementNS(XUL, "menulist");
    let menupopupAction = doc.createElementNS(XUL, "menupopup");
    ["archive", "copy", "delete", "move"].forEach( function(action) {
      let menuitem = doc.createElementNS(XUL, "menuitem");
      menuitem.setAttribute("label", action);
      menuitem.setAttribute("value", action);
      menupopupAction.insertBefore(menuitem, null);
    } );
    menulistAction.insertBefore(menupopupAction, null);
    menulistAction.setAttribute("value", rule.action);

    let menulistSrc = doc.createElementNS(XUL, "menulist");
    let menupopupSrc = doc.createElementNS(XUL, "menupopup");
    menulistSrc.insertBefore(menupopupSrc, null);
    menulistSrc.value = rule.src;

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
    menulistSub.setAttribute("value", rule.sub);
    
    let to = doc.createElementNS(XUL, "label");
    to.setAttribute("value", "To");
    let menulistDest = doc.createElementNS(XUL, "menulist");
    let menupopupDest = doc.createElementNS(XUL, "menupopup");
    menulistDest.insertBefore(menupopupDest, null);
    menulistDest.value = rule.dest || '';

    let after = doc.createElementNS(XUL, "label");
    after.setAttribute("value", "After");
    let age = doc.createElementNS(XUL, "textbox");
    age.setAttribute("type", "number");
    age.setAttribute("min", "0");
    age.setAttribute("value", rule.age);
    age.setAttribute("maxwidth", "60");
    let days = doc.createElementNS(XUL, "label");
    days.setAttribute("value", "days");
    
    let remove = doc.createElementNS(XUL, "button");
    remove.setAttribute("label", "X");
    //remove.setAttribute("oncommand", self.);
    
    let hbox = doc.createElementNS(XUL, "hbox");
    [checkbox, menulistAction, menulistSrc, menulistSub, to, menulistDest, after, age, days, remove].forEach( function(item) {
      hbox.insertBefore(item, null);
    } );
    parent.insertBefore(hbox, null);
    self.initFolderPick(doc, win, menulistSrc, menupopupSrc);
    self.initFolderPick(doc, win, menulistDest, menupopupDest);
  },

  loadPerfWindow: function(win) {
    try {
      let doc = win.document;
      let group = doc.getElementById('awsome_auto_archive-rules');
      autoArchiveLog.info('group:' + group);
      autoArchivePref.rules.forEach( function(rule) {
        self.creatOneRule(doc, win, rule, group);
      } );
    } catch (err) { /*autoArchiveLog.logException(err);*/ }
    return true;
  },
  acceptPerfWindow: function(win) {
    try {
      let disabled = [];
      for ( let checkbox of win.document.getElementById('ldapinfoshow-enable-servers').childNodes ) {
        if ( checkbox.key && !checkbox.checked ) disabled.push(checkbox.key);
      }
      this.prefs.setCharPref("disabled_servers", disabled.join(','));
      if ( !this.options.warned_about_fbli && ( this.options.load_from_facebook || this.options.load_from_linkedin ) ) {
        this.prefs.setBoolPref("warned_about_fbli", true);
        let strBundle = Services.strings.createBundle('chrome://awsomeAutoArchive/locale/ldapinfoshow.properties');
        this.loadUseProtocol("http://code.google.com/p/ldapinfo/wiki/Help");
        let result = Services.prompt.confirm(win, strBundle.GetStringFromName("prompt.warning"), strBundle.GetStringFromName("prompt.confirm.fbli"));
        if ( !result ) {
          this.prefs.setBoolPref("load_from_facebook", false);
          this.prefs.setBoolPref("load_from_linkedin", false);
          return false;
        }
      }
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },

}
let self = autoArchivePrefDialog;
self.initName();
