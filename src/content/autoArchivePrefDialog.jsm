// MPL/GPL
// Opera.Wang 2013/06/04
"use strict";
var EXPORTED_SYMBOLS = ["autoArchivePrefDialog"];

const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource://app/modules/gloda/utils.js");
//Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("resource:///modules/iteratorUtils.jsm");
Cu.import("resource:///modules/folderUtils.jsm");
Cu.import("resource:///modules/MailUtils.js");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveService.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveUtil.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
const perfDialogTooltipID = "awsome_auto_archive-perfDialogTooltip";
const XUL = "http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul";
const ruleClass = 'awsome_auto_archive-rule';
const ruleHeaderContextMenuID = 'awsome_auto_archive-rule-header-context';

let autoArchivePrefDialog = {
  strBundle: Services.strings.createBundle('chrome://awsomeAutoArchive/locale/awsome_auto_archive.properties'),
  _doc: null,
  _win: null,
  cleanup: function() {
    autoArchiveLog.info("autoArchivePrefDialog cleanup");
    if ( this._win && !this._win.closed ) this._win.close();
    autoArchiveLog.info("autoArchivePrefDialog cleanup done");
  },
  
  showPrettyTooltip: function(URI,pretty) {
    return decodeURIComponent(URI).replace(/(.*\/)[^/]*/, '$1') + pretty;
  },
  getPrettyName: function(msgFolder) {
    // access msgFolder.prettyName for non-existing local folder may cause creating wrong folder
    if ( !( msgFolder instanceof Ci.nsIMsgFolder ) || msgFolder.server.type != 'none' || autoArchiveUtil.folderExists(msgFolder) ) return msgFolder.prettyName;
    return msgFolder.URI.replace(/^.*\/([^\/]+)/,'$1');
  },
  getFolderAndSetLabel: function(folderPicker, setLabel) {
    let msgFolder = {value: '', prettyName: 'N/A', server: {}};
    try {
      msgFolder = MailUtils.getFolderForURI(folderPicker.value);
    } catch(err) {}
    if ( !this._doc || !setLabel ) return msgFolder;
    let showFolderAs = this._doc.getElementById('pref.show_folder_as');
    let label = "";
    switch ( showFolderAs.value ) {
      case 0:
        label = self.getPrettyName(msgFolder);
        break;
      case 1:
        label = "[" + msgFolder.server.prettyName + "] " + ( msgFolder == msgFolder.rootFolder ? "/" : self.getPrettyName(msgFolder) );
        break;
      case 2:
      default:
        label = self.showPrettyTooltip(msgFolder.ValueUTF8||msgFolder.value, self.getPrettyName(msgFolder));
        break;
    }
    folderPicker.setAttribute("label", label);
    folderPicker.setAttribute("folderStyle", showFolderAs.value); // for css to set correct length
    return msgFolder;
  },
  changeShowFolderAs: function() {
    if ( !this._doc ) return;
    let container = this._doc.getElementById('awsome_auto_archive-rules');
    if ( !container ) return;
    for ( let row of container.childNodes ) {
      if ( row.classList.contains(ruleClass) ) {
        for ( let item of row.childNodes ) {
          let key = item.getAttribute('rule');
          if ( ["src", "dest"].indexOf(key) >= 0 /*&& item.style.visibility != 'hidden'*/ )
            this.getFolderAndSetLabel(item, true);
        }
      }
    }
  },
  updateFolderStyle: function(folderPicker, folderPopup, init) {
    let msgFolder = this.getFolderAndSetLabel(folderPicker, false);
    let updateStyle = function() {
      let hasError = !autoArchiveUtil.folderExists(msgFolder);
      try {
        if ( typeof(folderPopup.selectFolder) != 'undefined' ) folderPopup.selectFolder(msgFolder); // false alarm by addon validator
        else return;
        if ( !hasError ) folderPopup._setCssSelectors(msgFolder, folderPicker); // _setCssSelectors may also create wrong local folders
      } catch(err) {
        hasError = true;
        //autoArchiveLog.logException(err);
      }
      if ( hasError ) {
        autoArchiveLog.info("Error: folder '" + self.getPrettyName(msgFolder) + "' can't be selected");
        folderPicker.classList.add("awsome_auto_archive-folderError");
        folderPicker.classList.remove("folderMenuItem");
      } else {
        folderPicker.classList.remove("awsome_auto_archive-folderError");
        folderPicker.classList.add("folderMenuItem");
      }
      if ( msgFolder.noSelect ) folderPicker.setAttribute("NoSelect", "true");
      else folderPicker.removeAttribute("NoSelect");
      self.getFolderAndSetLabel(folderPicker, true);
    };
    if ( msgFolder.rootFolder ) {
      if ( init ) this._win.setTimeout( updateStyle, 0 ); // use timer to wait for the XBL bindings add SelectFolder / _setCssSelectors to popup
      else updateStyle();
    }
    folderPicker.setAttribute('tooltiptext', self.showPrettyTooltip(msgFolder.ValueUTF8||msgFolder.value, self.getPrettyName(msgFolder)));
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
    /* Folders like [Gmail] are disabled by default, enable them, also true for IMAP servers like mercury/32, http://kb.mozillazine.org/Grey_italic_folders */
    let nsResolver = function(prefix) {
      let ns = { 'xul': "http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul" };
      return ns[prefix] || null;
    }
    folderPicker.addEventListener('popupshown', function(aEvent) {
      try {
        let menuitems = self._doc.evaluate(".//xul:menuitem[@disabled='true' and @generated='true']", folderPicker, nsResolver, Ci.nsIDOMXPathResult.UNORDERED_NODE_SNAPSHOT_TYPE, null);
        for ( let i=0 ; i < menuitems.snapshotLength; i++ ) {
          let menuitem = menuitems.snapshotItem(i);
          if ( menuitem._folder && menuitem._folder.noSelect ) {
            menuitem.removeAttribute('disabled');
            menuitem.setAttribute("NoSelect", "true"); // so it will show as in folder pane
          }
        }
      } catch (err) { autoArchiveLog.logException(err); }
    }, false);
    folderPicker.classList.add("folderMenuItem");
    folderPicker.setAttribute("sizetopopup", "none");
    folderPicker.setAttribute("crop", "center");

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
      ["", "action", "source", "scope", "dest", "from", "recipient", "subject", "size", "tags", "age", "", "", "picker"].forEach( function(label) {
        let item;
        if ( label == 'picker' ) {
          item = doc.createElementNS(XUL, "image");
          item.classList.add("tree-columnpicker-icon");
          item.addEventListener('click', function (event) { return doc.getElementById(ruleHeaderContextMenuID).openPopup(item, 'after_start', 0, 0, true, false, event); }, false );
          item.setAttribute("tooltiptext", self.strBundle.GetStringFromName("perfdialog.tooltip.picker"));
        } else {
          item = doc.createElementNS(XUL, "label");
          item.setAttribute('value', label ? self.strBundle.GetStringFromName("perfdialog." + label) : "");
          item.setAttribute('rule', label); // header does not have class ruleClass
        }
        let preference = doc.getElementById('pref.show_' + label);
        if ( preference ) {
          let actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
          item.style.display = actualValue ? '-moz-box': 'none';
        }
        row.insertBefore(item, null);
      } );
      row.id = "awsome_auto_archive-rules-header";
      row.setAttribute('context', ruleHeaderContextMenuID);
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
      menulistSub.setAttribute("tooltiptext", self.strBundle.GetStringFromName('perfdialog.tooltip.scope'));
      
      let menulistDest = doc.createElementNS(XUL, "menulist");
      let menupopupDest = doc.createElementNS(XUL, "menupopup");
      menulistDest.insertBefore(menupopupDest, null);
      menulistDest.value = rule.dest || '';
      menulistDest.setAttribute("rule", 'dest');

      let [from, recipient, subject, size, tags, age] = [
        // filter, size, default, tooltip, type, min          
        ["from", "10", '', perfDialogTooltipID],
        ["recipient", "10", '', perfDialogTooltipID],
        ["subject", '',  '', perfDialogTooltipID],
        ["size", '5', '', perfDialogTooltipID],
        ["tags", '10', '', perfDialogTooltipID],
        ["age", "4", autoArchivePref.options.default_days, '', 'number', "0"] ].map( function(attributes) {
          let element = doc.createElementNS(XUL, "textbox");
          let [filter, size, defaultValue, tooltip, type, min] = attributes;
          element.setAttribute("rule", filter);
          if ( size ) element.setAttribute("size", size);
          element.setAttribute("value", typeof(rule[filter]) != 'undefined' ? rule[filter] : defaultValue);
          if ( tooltip ) element.tooltip = tooltip;
          if ( type ) element.setAttribute("type", type);
          if ( typeof(min) != 'undefined' ) element.setAttribute("min", "0");
          let preference = doc.getElementById('pref.show_' + filter);
          let actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
          element.style.display = actualValue ? '-moz-box': 'none';
          return element;
        } );
      
      let [up, down, remove ] = [
        ['\u2191', function(aEvent) { self.upDownRule(row, true); }, ''],
        ['\u2193', function(aEvent) { self.upDownRule(row, false); }, ''],
        ['x', function(aEvent) { self.removeRule(row); }, 'awsome_auto_archive-delete-rule'] ].map( function(attributes) {
          let element = doc.createElementNS(XUL, "toolbarbutton");
          element.setAttribute("label", attributes[0]);
          element.addEventListener("command", attributes[1], false );
          if (attributes[2]) element.classList.add(attributes[2]);
          return element;
        } );
      
      row.classList.add(ruleClass);
      [enable, menulistAction, menulistSrc, menulistSub, menulistDest, from, recipient, subject, size, tags, age, up, down, remove].forEach( function(item) {
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
        remove.parentNode.removeChild(remove);
        // remove.parentNode.insertBefore(remove, ref); // lost all unsaved values
        let newBox = this.creatOneRule(rule, ref)
        this.checkFocus( isUp ? newBox : row );
        this.syncToPerf(true);
      }
    } catch(err) {
      autoArchiveLog.logException(err);
    }
  },
  
  removeRule: function(row) {
    row.parentNode.removeChild(row); // will cause 'Error: TypeError: temp is null Source file: chrome://global/content/bindings/preferences.xml Line: 1172'
    this.syncToPerf(true);
  },
  
  revertRules: function() {
    if ( !this._doc ) return;
    this.syncToPerf(true);
    let preference = this._doc.getElementById("pref.rules");
    autoArchiveLog.info("Revert rules from\n" + preference.value + "\nto\n" + this._savedRules);
    preference.value = this._savedRules; // perfpane.userChangedValue is the same
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
    if ( [autoArchiveService.STATUS_SLEEP, autoArchiveService.STATUS_WAITIDLE, autoArchiveService.STATUS_FINISH, autoArchiveService.STATUS_HIBERNATE].indexOf(status) >= 0 ) {
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
  
  creatNewRule: function(rule) {
    if ( !rule ) rule = {action: 'archive', enable: true, sub: 0, age: autoArchivePref.options.default_days};
    this.checkFocus( this.creatOneRule(rule, null) );
    this.syncToPerf(true);
  },
  changeRule: function(how) {
    if ( !this.focusRow ) return;
    if ( how == 'up' ) this.upDownRule(this.focusRow, true);
    else if ( how == 'down' ) this.upDownRule(this.focusRow, false);
    else if ( how == 'remove' ) this.removeRule(this.focusRow);
  },
  
  syncFromPerf: function(win) { // this need 0.5s for 8 rules
    //autoArchiveLog.info('syncFromPerf');
    // if not modal, user can open 2nd pref window, we will close the old one, and close/unLoadPerfWindow seems a sync call, so we are fine
    if ( this._win && this._win != win && !this._win.closed ) this._win.close();
    this._win = win;
    this._doc = win.document;
    let preference = this._doc.getElementById("pref.rules");
    let actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
    if ( actualValue === this.oldvalue ) return;
    this.createRuleHeader();
    let rules = JSON.parse(actualValue);
    if ( rules.length ) {
      rules.forEach( function(rule) {
        self.creatOneRule(rule, null);
      } );
      this.oldvalue = actualValue;
    } else if ( !win.arguments || !win.arguments[0] ) { // don't create empty rule if loadPerfWindow will create new rule based on selected email
      this.creatNewRule();
    }
    //autoArchiveLog.info('syncFromPerf done');
  },
  
  syncToPerf: function(store2pref) { // this need 0.005s for 8 rules
    //autoArchiveLog.info('syncToPerf');
    let value = JSON.stringify(this.getRules());
    this.oldvalue = value; // need before set preference.value, which will cause syncFromPref
    if ( store2pref ) {
      let preference = this._doc.getElementById("pref.rules");
      preference.value = value;
    }
    //autoArchiveLog.info('syncToPerf done');
    return value;
  },
  
  PopupShowing: function(event) {
    try {
      let doc = event.view.document;
      let tooltip = doc.getElementById(perfDialogTooltipID);
      let line1 = tooltip.firstChild.firstChild;
      let line2 = line1.nextSibling;
      let line3 = line2.nextSibling;
      let line4 = line3.nextSibling;
      let triggerNode = event.target.triggerNode;
      let rule = triggerNode.getAttribute('rule');
      let supportRE = autoArchiveService.advancedTerms[rule].some( function(term) {
        return MailServices.filters.getCustomTerm(term);
      } );
      let str = function(label) { return self.strBundle.GetStringFromName("perfdialog.tooltip." + label); };
      line1.value = (triggerNode.value == "") ? str("emptyFilter") : triggerNode.value;
      line2.value = supportRE ? str("hasRE") : ( ["size", "tags"].indexOf(rule) >= 0 ? str("line2." + rule) : str("noRE") );
      line3.value = str("line3." + rule);
      line4.value = str("negativeSearch") + str("negative." + rule + "Example");
    } catch (err) {
      autoArchiveLog.logException(err);
    }
    return true;
  },
  
  syncFromPerf4Filter: function(obj) {
    //autoArchiveLog.info('syncFromPerf4Filter:' + obj.getAttribute("preference"));
    let doc = obj.ownerDocument;
    let perfID = obj.getAttribute("preference");
    let preference = doc.getElementById(perfID);
    let actualValue = preference.value !== undefined ? preference.value : preference.defaultValue;
    let oldValue = obj.oldValue;
    if ( oldValue != actualValue ) {
      obj.setAttribute("checked", actualValue);
      let container = doc.getElementById('awsome_auto_archive-rules');
      if ( !container ) return;
      for ( let row of container.childNodes ) {
        for ( let item of row.childNodes ) {
          let key = item.getAttribute('rule');
          if ( "pref.show_" + key == perfID )
            item.style.display = actualValue ? '-moz-box': 'none';
        }
      }
    }
    obj.oldValue = actualValue;
    return actualValue;
  },
  
  syncToPerf4Filter: function(obj) {
    //autoArchiveLog.info('syncToPerf4Filter:' + obj.getAttribute("preference"));
    let preference = obj.ownerDocument.getElementById(obj.getAttribute("preference"));
    preference.value = obj.getAttribute("checked") ? true : false;
    return preference.value;
  },

  loadPerfWindow: function(win) {
    try {
      autoArchiveLog.info('loadPerfWindow');
      if ( !this._win ) this.syncFromPerf(win); // SeaMonkey may have one dialog open from addon manager, and then open another one from icon or context menu
      this.instantApply = this._doc.getElementById('awsome_auto_archive-preferences').instantApply || false;
      autoArchivePref.setInstantApply(this.instantApply);
      if ( this.instantApply ) { // only use synctopreference for instantApply, else use acceptPerfWindow
        // must be a onsynctopreference attribute, not a event handler, ref preferences.xml
        this._doc.getElementById('awsome_auto_archive-rules').setAttribute("onsynctopreference", 'return autoArchivePrefDialog.syncToPerf();');
        // no need to show 'Apply' button for instantApply
        let extra1 = this._doc.documentElement.getButton("extra1");
        if ( extra1 && extra1.parentNode ) extra1.parentNode.removeChild(extra1);
      }
      autoArchiveService.addStatusListener(this.statusCallback);
      this.fillIdentities(false);
      let tooltip = this._doc.getElementById(perfDialogTooltipID);
      if ( tooltip ) tooltip.addEventListener("popupshowing", this.PopupShowing, true);
      this._savedRules = autoArchivePref.options.rules;
      if ( win.arguments && win.arguments[0] ) { // new rule based on message selected, not including in the revert all
        let msgHdr = win.arguments[0];
        this.creatNewRule( {action: 'archive', enable: true, src: msgHdr.folder.URI, sub: 0, from: this.getSearchStringFromAddress(msgHdr.mime2DecodedAuthor),
          recipient: this.getSearchStringFromAddress(msgHdr.mime2DecodedTo || msgHdr.mime2DecodedRecipients), subject: msgHdr.mime2DecodedSubject, age: autoArchivePref.options.default_days} );
      }
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },
  getSearchStringFromAddress: function(mails) {
    // GetDisplayNameInAddressBook() in http://mxr.mozilla.org/comm-central/source/mailnews/base/src/nsMsgDBView.cpp
    try {
      let parsedMails = GlodaUtils.parseMailAddresses(mails);
      let returnMails = [];
      for ( let i = 0; i < parsedMails.count; i++ ) {
        let email = parsedMails.addresses[i], card, displayName;
        if ( !autoArchivePref.options.generate_rule_use ) {
          if ( Services.prefs.getBoolPref("mail.showCondensedAddresses") ) { // the usage of getSearchStringFromAddress might be few, so won't add Observer 
            let allAddressBooks = MailServices.ab.directories;
            while ( !card && allAddressBooks.hasMoreElements()) {
              let addressBook = allAddressBooks.getNext().QueryInterface(Ci.nsIAbDirectory);
              if ( addressBook instanceof Ci.nsIAbDirectory /*&& !addressBook.isRemote*/ ) {
                try {
                  card = addressBook.cardForEmailAddress(email); // case-insensitive && sync, only return 1st one if multiple match, but it search on all email addresses
                } catch (err) {}
                if ( card ) {
                  let PreferDisplayName = Number(card.getProperty('PreferDisplayName', 1));
                  if (PreferDisplayName) displayName = card.displayName;
                }
              }
            }
          }
          if ( !displayName ) displayName = parsedMails.names[i] || parsedMails.fullAddresses[i];
          displayName = displayName.replace(/['"<>]/g,'');
          if ( parsedMails.fullAddresses[i].indexOf(displayName) != -1 ) email = displayName;
        }
        let search = (autoArchivePref.options.generate_rule_use == 2) ? email : email.replace(/(.*@).*/, '$1');
        if ( returnMails.indexOf(search) < 0 ) returnMails.push(search);
      }
      return returnMails.join(", ");
    } catch (err) { autoArchiveLog.logException(err); }
    return mails;
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
    } catch (err) { autoArchiveLog.logException(err); throw err; } // throw the error out so syncToPerf won't get an empty rules
    return rules;
  },
  acceptPerfWindow: function() {
    try {
      autoArchiveLog.info("acceptPerfWindow");
      if ( !this.instantApply ) autoArchivePref.setPerf('rules', this.syncToPerf());
    } catch (err) { autoArchiveLog.logException(err); }
    return true;
  },
  unLoadPerfWindow: function() {
    if ( !autoArchiveService || !autoArchivePref || !autoArchiveLog || !autoArchiveUtil ) return true;
    if ( this._savedRules != autoArchivePref.options.rules ) autoArchiveUtil.backupRules(autoArchivePref.options.rules);
    autoArchiveService.removeStatusListener(this.statusCallback);
    let tooltip = this._doc.getElementById(perfDialogTooltipID);
    if ( tooltip ) tooltip.removeEventListener("popupshowing", this.PopupShowing, true);
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
    let tabbox = doc.getElementById('awsome_auto_archive-tabbox');
    if ( !group || !pane || !tabbox ) return;
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
    
    try {
      let perfDialog = self._doc.getElementById('awsome_auto_archive-prefs');
      let buttonBox = self._doc.getAnonymousElementByAttribute(perfDialog, "anonid", "dlg-buttons");
      let targetWinHeight = buttonBox.scrollHeight + pane.contentHeight;
      let currentWinHeight = perfDialog.height;
      if ( currentWinHeight < targetWinHeight+62 ) perfDialog.setAttribute('height', targetWinHeight+62);
      perfDialog.style.minHeight = targetWinHeight + "px";
      let width = Number(perfDialog.width || perfDialog.getAttribute("width"));
      let targetWidth = Number(tabbox.clientWidth  || tabbox.scrollWidth) + 36;
      if ( width < targetWidth ) perfDialog.setAttribute("width", targetWidth);
    } catch (err) {autoArchiveLog.logException(err);}
  },
  
  applyChanges: function() {
    Array.prototype.forEach.call( self._doc.getElementById("awsome_auto_archive-prefs").preferencePanes, function(pane) {
      pane.writePreferences(true);
    } );
  },

}

let self = autoArchivePrefDialog;
