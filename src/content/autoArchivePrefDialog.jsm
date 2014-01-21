// MPL/GPL
// Opera.Wang 2013/06/04
"use strict";
var EXPORTED_SYMBOLS = ["autoArchivePrefDialog"];

const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
//Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("resource:///modules/iteratorUtils.jsm");
Cu.import("resource:///modules/folderUtils.jsm");
Cu.import("resource:///modules/MailUtils.js");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveService.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
const SEAMONKEY_ID = "{92650c4d-4b8e-4d2a-b7eb-24ecf4f6b63a}";
const XUL = "http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul";
const ruleClass = 'awsome_auto_archive-rule';

let autoArchivePrefDialog = {
  strBundle: Services.strings.createBundle('chrome://awsomeAutoArchive/locale/awsome_auto_archive.properties'),
  _doc: null,
  _win: null,
  cleanup: function() {
  },
  
  showPrettyTooltip: function(URI,pretty) {
    return decodeURIComponent(URI).replace(/(.*\/)[^/]*/, '$1') + pretty;
  },
  updateFolderStyle: function(folderPicker, folderPopup, init) {
    let msgFolder = {value: '', prettyName: 'N/A'};
    let updateStyle = function() {
      let hasError = false;
      try {
        if ( typeof(folderPopup.selectFolder) != 'undefined' ) folderPopup.selectFolder(msgFolder); // false alarm by addon validator
        else return;
        folderPopup._setCssSelectors(msgFolder, folderPicker);
      } catch(err) {
        hasError = true;
        //autoArchiveLog.logException(err);
      }
      if ( hasError ) {
        autoArchiveLog.info("Error: folder '" + msgFolder.prettyName + "' can't be selected");
        folderPicker.classList.add("awsome_auto_archive-folderError");
        folderPicker.classList.remove("folderMenuItem");
      } else {
        folderPicker.classList.remove("awsome_auto_archive-folderError");
        folderPicker.classList.add("folderMenuItem");
      }
    };
    try {
      msgFolder = MailUtils.getFolderForURI(folderPicker.value);
      if ( init ) this._win.setTimeout( updateStyle, 0 );// use timer to wait for the XBL bindings add SelectFolder / _setCssSelectors to popup
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
    self.updateFolderStyle(folderPicker, folderPopup, false);
  },
  initFolderPick: function(folderPicker, folderPopup, isSrc) {
    folderPicker.addEventListener('command', function(aEvent) { return self.onFolderPick(folderPicker, aEvent, folderPopup); }, false);
    folderPicker.classList.add("folderMenuItem");
    folderPicker.setAttribute("sizetopopup", "none");

    folderPopup.setAttribute("type", "folder");
    if ( !isSrc ) {
      folderPopup.setAttribute("mode", "filing");
      folderPopup.setAttribute("showFileHereLabel", "true");
    }
    folderPopup.classList.add("menulist-menupopup");
    folderPopup.classList.add("searchPopup");
    self.updateFolderStyle(folderPicker, folderPopup, true);
  },
  createRuleHeader: function() {
    try {
      let doc = this._doc;
      let container = doc.getElementById('awsome_auto_archive-rules');
      if ( !container ) return;
      while (container.firstChild) container.removeChild(container.firstChild);
      let row = doc.createElementNS(XUL, "row");
      ["", "action", "source", "scope", "dest", "from", "recipient", "subject", "age", "", "", ""].forEach( function(label) {
        let item = doc.createElementNS(XUL, "label");
        item.setAttribute('value', label ? self.strBundle.GetStringFromName("perfdialog." + label) : "");
        row.insertBefore(item, null);
      } );
      row.id = "awsome_auto_archive-rules-header";
      container.insertBefore(row, null);
    } catch (err) {
      autoArchiveLog.logException(err);
    }
  },
  creatOneRule: function(rule, ref) {
    try {
      let doc = this._doc;
      let container = doc.getElementById('awsome_auto_archive-rules');
      if ( !container ) return;
      let row = doc.createElementNS(XUL, "row");

      let enable = doc.createElementNS(XUL, "checkbox");
      enable.setAttribute("checked", rule.enable);
      enable.setAttribute("rule", 'enable');
      
      let menulistAction = doc.createElementNS(XUL, "menulist");
      let menupopupAction = doc.createElementNS(XUL, "menupopup");
      ["archive", "copy", "delete", "move"].forEach( function(action) {
        let menuitem = doc.createElementNS(XUL, "menuitem");
        menuitem.setAttribute("label", self.strBundle.GetStringFromName("perfdialog.action."+action));
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
      let types = [ {key: self.strBundle.GetStringFromName("perfdialog.type.only"), value: 0}, { key: self.strBundle.GetStringFromName("perfdialog.type.sub"), value: 1}, {key: self.strBundle.GetStringFromName("perfdialog.type.sub_keep"), value: 2} ];
      types.forEach( function(type) {
        let menuitem = doc.createElementNS(XUL, "menuitem");
        menuitem.setAttribute("label", type.key);
        menuitem.setAttribute("value", type.value);
        menupopupSub.insertBefore(menuitem, null);
      } );
      menulistSub.insertBefore(menupopupSub, null);
      menulistSub.setAttribute("value", rule.sub || 0);
      menulistSub.setAttribute("rule", 'sub');
      
      let menulistDest = doc.createElementNS(XUL, "menulist");
      let menupopupDest = doc.createElementNS(XUL, "menupopup");
      menulistDest.insertBefore(menupopupDest, null);
      menulistDest.value = rule.dest || '';
      menulistDest.setAttribute("rule", 'dest');

      let sender = doc.createElementNS(XUL, "textbox");
      sender.setAttribute("value", rule.from || '');
      sender.setAttribute("rule", 'from');
      sender.setAttribute("size", "10");
      
      let recipient = doc.createElementNS(XUL, "textbox");
      recipient.setAttribute("value", rule.recipient || '');
      recipient.setAttribute("rule", 'recipient');
      recipient.setAttribute("size", "10");
      
      let subject = doc.createElementNS(XUL, "textbox");
      subject.setAttribute("value", rule.subject || '');
      subject.setAttribute("rule", 'subject');
      
      let age = doc.createElementNS(XUL, "textbox");
      age.setAttribute("type", "number");
      age.setAttribute("min", "0");
      age.setAttribute("value", typeof(rule.age)!='undefined' ? rule.age : autoArchivePref.options.default_days);
      age.setAttribute("rule", 'age');
      age.setAttribute("size", "4");
      
      let up = doc.createElementNS(XUL, "toolbarbutton");
      up.setAttribute("label", '\u2191');
      up.addEventListener("command", function(aEvent) { self.upDownRule(row, true); }, false );
      
      let down = doc.createElementNS(XUL, "toolbarbutton");
      down.setAttribute("label", '\u2193');
      down.addEventListener("command", function(aEvent) { self.upDownRule(row, false); }, false );
      
      let remove = doc.createElementNS(XUL, "toolbarbutton");
      remove.setAttribute("label", "x");
      remove.classList.add("awsome_auto_archive-delete-rule");
      remove.addEventListener("command", function(aEvent) { self.removeRule(row); }, false );
      
      row.classList.add(ruleClass);
      [enable, menulistAction, menulistSrc, menulistSub, menulistDest, sender, recipient, subject, age, up, down, remove].forEach( function(item) {
        row.insertBefore(item, null);
      } );
      container.insertBefore(row, ref);
      self.initFolderPick(menulistSrc, menupopupSrc, true);
      self.initFolderPick(menulistDest, menupopupDest, false);
      self.checkAction(menulistAction, menulistDest, menulistSub);
      self.checkEnable(enable, row);
      menulistAction.addEventListener('command', function(aEvent) { self.checkAction(menulistAction, menulistDest, menulistSub); }, false );
      enable.addEventListener('command', function(aEvent) { self.checkEnable(enable, row); }, false );
      row.addEventListener('focus', function(aEvent) { self.checkFocus(row); }, true );
      row.addEventListener('click', function(aEvent) { self.checkFocus(row); }, true );
      return row;
    } catch(err) {
      autoArchiveLog.logException(err);
    }
  },
  
  focusRow: null,
  checkFocus: function(row) {
    if ( this.focusRow && this.focusRow != row )  this.focusRow.removeAttribute('awsome_auto_archive-focused');
    row.setAttribute('awsome_auto_archive-focused', true);
    this.focusRow = row;
  },
  
  upDownRule: function(row, isUp) {
    try {
      let ref = isUp ? row.previousSibling : row;
      let remove = isUp ? row : row.nextSibling;
      if ( ref && remove && ref.classList.contains(ruleClass) && remove.classList.contains(ruleClass) ) {
        let rule = this.getOneRule(remove);
        autoArchiveLog.logObject(rule, 'temp rule', 1);
        remove.parentNode.removeChild(remove);
        // remove.parentNode.insertBefore(remove, ref); // lost all unsaved values
        let newBox = this.creatOneRule(rule, ref)
        this.checkFocus( isUp ? newBox : row );
        this.syncToPerf();
      }
    } catch(err) {
      autoArchiveLog.logException(err);
    }
  },
  
  removeRule: function(row) {
    row.parentNode.removeChild(row);
    this.syncToPerf();
  },

  checkEnable: function(enable, row) {
    if ( enable.checked ) {
      row.classList.remove("awsome_auto_archive-disable");
    } else {
      row.classList.add("awsome_auto_archive-disable");
    }
  },
  
  checkAction: function(menulistAction, menulistDest, menulistSub) {
    let limit = ["archive", "delete"].indexOf(menulistAction.value) >= 0;
    if ( limit && menulistSub.value == 2 ) menulistSub.value = 1;
    menulistDest.style.visibility = limit ? 'hidden': 'visible';
    menulistSub.firstChild.lastChild.style.display = limit ? 'none': '-moz-box';
  },
  
  starStopNow: function(dry_run) {
    autoArchiveService.starStopNow(this.getRules(), dry_run);
  },
  
  statusCallback: function(status, detail) {
    let run_button = self._doc.getElementById('awsome_auto_archive-action');
    let dry_button = self._doc.getElementById('awsome_auto_archive-dry-run');
    if ( !run_button || !dry_button ) return;
    if ( [autoArchiveService.STATUS_SLEEP, autoArchiveService.STATUS_WAITIDLE].indexOf(status) >= 0 ) {
      // change run_button to "Run"
      run_button.setAttribute("label", self.strBundle.GetStringFromName("perfdialog.action.button.run"));
      dry_button.setAttribute("label", self.strBundle.GetStringFromName("perfdialog.action.button.dryrun"));
    } else if ( status == autoArchiveService.STATUS_RUN ) {
      // change run_button to "Stop"
      run_button.setAttribute("label", self.strBundle.GetStringFromName("perfdialog.action.button.stop"));
      dry_button.setAttribute("label", self.strBundle.GetStringFromName("perfdialog.action.button.stop"));
    }
    run_button.setAttribute("tooltiptext", detail);
    dry_button.setAttribute("tooltiptext", detail);
  },
  
  creatNewRule: function() {
    this.checkFocus(
      this.creatOneRule({action: 'archive', enable: true, sub: 0, age: autoArchivePref.options.default_days}, null)
    );
  },
  changeRule: function(how) {
    if ( !this.focusRow ) return;
    if ( how == 'up' ) this.upDownRule(this.focusRow, true);
    else if ( how == 'down' ) this.upDownRule(this.focusRow, false);
    else if ( how == 'remove' ) this.removeRule(this.focusRow);
  },
  
  syncFromPerf: function(win) { // this need 0.5s for 8 rules
    autoArchiveLog.info('syncFromPerf');
     // if not modal, user can open 2nd pref window, we will close the old one, and close/unLoadPerfWindow seems a sync call, so we are fine
    if ( this._win && this._win != win && !this._win.closed ) this._win.close();
    autoArchiveLog.info('syncFromPerf set win');
    this._win = win;
    this._doc = win.document;
    let preference = this._doc.getElementById("pref_rules");
    let actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
    //autoArchiveLog.info('act:' + actualValue);
    if ( actualValue === this.oldvalue ) return;
    this.createRuleHeader();
    let rules = JSON.parse(actualValue);
    if ( rules.length ) {
      rules.forEach( function(rule) {
        self.creatOneRule(rule, null);
      } );
      this.oldvalue = actualValue;
    } else {
      this.creatNewRule();
    }
    autoArchiveLog.info('syncFromPerf done');
  },
  
  syncToPerf: function() { // this need 0.005s for 8 rules
    autoArchiveLog.info('syncToPerf');
    let value = JSON.stringify(this.getRules());
    //autoArchiveLog.info('value:' + value);
    this.oldvalue = value;
    autoArchiveLog.info('syncToPerf done');
    return value;
  },

  loadPerfWindow: function(win) {
    try {
      autoArchiveLog.info('loadPerfWindow');
      //this._win = win;
      //this._doc = win.document;
      this.instantApply = this._doc.getElementById('awsome_auto_archive-preferences').instantApply || false;
      autoArchivePref.setInstantApply(this.instantApply);
      autoArchiveService.addStatusListener(this.statusCallback);
      this.fillIdentities(false);
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },
  getOneRule: function(row) {
    let rule = {};
    for ( let item of row.childNodes ) {
      let key = item.getAttribute('rule');
      if ( key ) {
        let value = item.value || item.checked;
        if ( item.getAttribute("type") == 'number' ) value = item.valueNumber;
        if ( key == 'sub' ) value = Number(value); // menulist.value is always 'string'
        rule[key] = value;
      }
    }
    return rule;
  },
  
  getRules: function() {
    let rules = [];
    try {
      let container = this._doc.getElementById('awsome_auto_archive-rules');
      if ( !container ) return rules;
      for ( let row of container.childNodes ) {
        if ( row.classList.contains(ruleClass) ) {
          let rule = this.getOneRule(row);
          if ( Object.keys(rule).length > 0 ) rules.push(rule);
        }
      }
      // autoArchiveLog.logObject(rules,'got rules',1);
    } catch (err) { autoArchiveLog.logException(err); }
    return rules;
  },
  acceptPerfWindow: function() {
    autoArchiveLog.info("acceptPerfWindow");
    //autoArchivePref.setPerf('rules',JSON.stringify(this.getRules()));
  },
  unLoadPerfWindow: function() {
    autoArchiveService.removeStatusListener(this.statusCallback);
    if ( this.instantApply ) autoArchivePref.validateRules();
    delete this._doc;
    delete this._win;
    delete this.oldvalue;
    delete this.instantApply;
    autoArchiveLog.info("prefwindow unload");
    return true;
  },
  
  //https://github.com/protz/thunderbird-stdlib/blob/master/misc.js
  fillIdentities: function(aSkipNntp) {
    let doc = self._doc;
    let group = doc.getElementById('awsome_auto_archive-IDs');
    let pane = doc.getElementById('awsome_auto_archive-perfpane');
    if ( !group || !pane ) return;
    let firstNonNull = null, gIdentities = {}, gAccounts = {};
    for (let account in fixIterator(MailServices.accounts.accounts, Ci.nsIMsgAccount)) {
      let server = account.incomingServer;
      if (aSkipNntp && (!server || server.type != "pop3" && server.type != "imap")) {
        continue;
      }
      for (let id in fixIterator(account.identities, Ci.nsIMsgIdentity)) {
        // We're only interested in identities that have a real email.
        if (id.email) {
          gIdentities[id.email.toLowerCase()] = id;
          gAccounts[id.email.toLowerCase()] = account;
          if (!firstNonNull) firstNonNull = id;
        }
      }
    }
    gIdentities["default"] = MailServices.accounts.defaultAccount.defaultIdentity || firstNonNull;
    gAccounts["default"] = MailServices.accounts.defaultAccount;
    Object.keys(gIdentities).sort().forEach( function(id) {
      let button = doc.createElementNS(XUL, "button");
      button.setAttribute("label", id);
      button.addEventListener("command", function(aEvent) { self._win.openDialog("chrome://messenger/content/am-identity-edit.xul", "dlg", "", {identity: gIdentities[id], account: gAccounts[id], result:false }); }, false );
      group.insertBefore(button, null);
    } );
    pane.style.minHeight = pane.contentHeight + 10 + "px"; // reset the pane height after fill Identities, to prevent vertical scrollbar
  },

}

let self = autoArchivePrefDialog;
