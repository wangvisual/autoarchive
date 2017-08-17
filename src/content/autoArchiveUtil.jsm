// MPL/GPL
// Opera.Wang 2013/06/04
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveUtil"];

const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource:///modules/iteratorUtils.jsm");
Cu.import("resource://gre/modules/AddonManager.jsm");
Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("resource://gre/modules/osfile.jsm");
Cu.import("resource://gre/modules/Promise.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
const SEAMONKEY_ID = "{92650c4d-4b8e-4d2a-b7eb-24ecf4f6b63a}";
const XUL = "http://www.mozilla.org/keymaster/gatekeeper/there.is.only.xul";

let autoArchiveUtil = {
  Name: "Awesome Auto Archive", // might get changed by getAddonByID function call
  Version: 'unknown',
  //isSeaMonkey: Services.appinfo.ID == SEAMONKEY_ID,
  //Application: Services.appinfo.name,
  NS_MSG_ERROR_IMAP_COMMAND_FAILED: 0x80550021,
  initName: function() {
    autoArchiveLog.info("autoArchiveUtil initName");
    if ( this.Version != 'unknown' ) return;
    AddonManager.getAddonByID('awsomeautoarchive@opera.wang', function(addon) {
      if ( !self || !autoArchiveLog ) return;
      self.Version = addon.version;
      self.Name = addon.name;
      autoArchiveLog.info(self.Name + " " + self.Version);
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
    let url = "https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=893LVBYFXCUP4&lc=US&item_name=" + this.Name + "&no_note=0&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_LG%2egif%3aNonHostedGuest";
    if ( typeof(pay) != 'undefined' ) {
      if ( pay == 'alipay' ) return this.loadTab( {type: 'contentTab', contentPage: "chrome://awsomeAutoArchive/content/qr-alipay.png"});
      if ( pay == 'mozilla' ) url = "https://addons.mozilla.org/en-US/thunderbird/addon/awesome-auto-archive/developers?src=api"; // Meet the developer page
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
  getErrorMsg: function(pStatus) { // https://developer.mozilla.org/en/docs/Table_Of_Errors
    if ( ( pStatus & 0x80590000 ) == 0x80590000 ) { // http://dxr.mozilla.org/mozilla-central/source/xpcom/base/nsError.h
      let ldapBundle = Services.strings.createBundle('chrome://mozldap/locale/ldap.properties');
      try { return ldapBundle.GetStringFromID(pStatus & 0xff); } catch(err) {};
    } else {
      for ( let p in Cr ) {
          if ( Cr[p] == pStatus ) return p;
      }
    }
    if ( pStatus == this.NS_MSG_ERROR_IMAP_COMMAND_FAILED ) return 'NS_MSG_ERROR_IMAP_COMMAND_FAILED';
    return 'Unknown Error';
  },
  // http://stackoverflow.com/questions/4498866/actual-numbers-to-the-human-readable-values
  readablizeBytes: function(bytes) {
    if ( bytes <= 0 ) return 0;
    let s = ' KMGTPE';
    let e = Math.floor(Math.log(bytes) / Math.log(1024));
    return (bytes / Math.pow(1024, e)).toFixed(2) + s[e];
  },
  getKeyFromTag: function(myTag) {
    myTag = myTag.toLowerCase();
    if ( myTag == 'na' ) return myTag;
    var tagArray = MailServices.tags.getAllTags({});
    // consider two tags, one is "ABC", the other is "ABCD", when searching for "AB", perfect is return both.
    // but here I just return the best fit "ABC".
    let uniqueKey = '';
    let myTagLen = myTag.length;
    let lenDiff = 10000000; // big enough?
    for (var i = 0; i < tagArray.length; ++i) {
      let tag = tagArray[i].tag.toLowerCase();
      if (tag.indexOf(myTag) >= 0 && ( tag.length-myTagLen < lenDiff ) ) {
        uniqueKey = tagArray[i].key;
        lenDiff = tag.length-myTagLen;
        if ( lenDiff == 0 ) break;
      }
    }
    if (uniqueKey != '') return uniqueKey;
    autoArchiveLog.log( "Can't find tag '" + myTag + "' in '" + tagArray.map((element) => { return element.tag; }).join(', ') + "'", 'Error');
    return "..unknown.."; // to prevent tags not in use and search hit more messages than expected
  },
  addSearchTerm: function(searchSession, attr, str, op) { // simple version of the one in expression search
    try {
      let aCustomId;
      if ( typeof(attr) == 'object' && attr.type == Ci.nsMsgSearchAttrib.Custom ) {
        aCustomId = attr.customId;
        attr = Ci.nsMsgSearchAttrib.Custom;
      }
      let term = searchSession.createTerm();
      term.attrib = attr;
      let value = term.value;
      // This is tricky - value.attrib must be set before actual values, from searchTestUtils.js 
      value.attrib = attr;
      if (attr == Ci.nsMsgSearchAttrib.JunkPercent)
        value.junkPercent = str;
      else if (attr == Ci.nsMsgSearchAttrib.Priority)
        value.priority = str;
      else if (attr == Ci.nsMsgSearchAttrib.Date)
        value.date = str;
      else if (attr == Ci.nsMsgSearchAttrib.MsgStatus || attr == Ci.nsMsgSearchAttrib.FolderFlag || attr == Ci.nsMsgSearchAttrib.Uint32HdrProperty)
        value.status = str;
      else if (attr == Ci.nsMsgSearchAttrib.Size)
        value.size = str;
      else if (attr == Ci.nsMsgSearchAttrib.AgeInDays)
        value.age = str;
      else if (attr == Ci.nsMsgSearchAttrib.Label)
        value.label = str;
      else if (attr == Ci.nsMsgSearchAttrib.JunkStatus)
        value.junkStatus = str;
      else if (attr == Ci.nsMsgSearchAttrib.HasAttachmentStatus)
        value.status = nsMsgMessageFlags.Attachment;
      else
        value.str = str;
      if (attr == Ci.nsMsgSearchAttrib.Custom)
        term.customId = aCustomId;
      term.value = value;
      term.op = op;
      term.booleanAnd = true;
      searchSession.appendTerm(term);
    } catch (err) { autoArchiveLog.logException(err); }
  },
  sizeToKB: function(expression) {
    try {
      let match = expression.match(/^([-.\d]*)(\w*)/i); // default KB, can be M,G
      if ( match.length == 3 ) {
        let [, size, scale] = match;
        if ( scale == '' ) scale = 1;
        if ( /^m/i.test(scale) ) {
          scale = 1024;
        } else if ( /^G/i.test(scale) ) {
          scale = 1024 * 1024;
        } else if ( /^K/i.test(scale) ) {
          scale = 1;
        } else if ( scale != 1 ) {
          autoArchiveLog.log("Unknown size scale:'" + scale + "', can be K(default),M,G", 1);
          scale = 1;
        }
        return size * scale; 
      }
    } catch (err) { autoArchiveLog.logException(err); }
    return -1;
  },
  folderExists: function(folder) {
    return folder && ( folder.parent || folder.isServer );
  },
  createFolderAsync: function(folder) {
    return folder.server && folder.server.type != 'none' && folder.server.protocolInfo && folder.server.protocolInfo.foldersCreatedAsync;
  },
  getSearchSessionString: function(session) {
    let folders = [], condition = "", scopes = session.countSearchScopes();
    for ( let count = 0; count < scopes; count ++ ) {
      let folder = {};
      session.getNthSearchScope(count, {}, folder);
      if ( folder && folder.value && folder.value.QueryInterface ) {
        folder.value.QueryInterface(Ci.nsIMsgFolder);
        folders.push(folder.value.URI);
      }
    }
    
    for ( let searchTerm in fixIterator(session.searchTerms, Ci.nsIMsgSearchTerm) ) {
      if (condition.length > 0) condition += " ";
      if (searchTerm.matchAll)
        condition += "ALL";
      else {
        condition += searchTerm.booleanAnd ? "AND" : "OR";
        condition += searchTerm.beginsGrouping && !searchTerm.endsGrouping ? " {" : "";
      }
      let converter = Cc["@mozilla.org/intl/scriptableunicodeconverter"].createInstance(Ci.nsIScriptableUnicodeConverter);
      converter.charset = 'UTF-8';
      let termString = converter.ConvertToUnicode(searchTerm.termAsString); // termAsString is ACString
      condition += " (" + termString + ")"; 
      // "}" may not balanced with "{", but who cares
      condition += searchTerm.endsGrouping && !searchTerm.beginsGrouping ? " }" : "";
    }
    return "Searching\n" + folders.join('\n') + "\nwith\n" + condition;
  },
  rulesBackupPath: OS.Path.join(OS.Constants.Path.profileDir, "awsomeautoarchive"),
  backupRules: function(rules, rules_to_keep) {
    try {
      let encoder = new TextEncoder();
      let fileName = OS.Path.join(self.rulesBackupPath, "rules." + new Date().toISOString().replace(/\W/g, '_'));
      OS.File.makeDir(self.rulesBackupPath, {ignoreExisting: true}).then( () => {
        return OS.File.writeAtomic(fileName, encoder.encode(rules), {tmpPath: fileName + ".tmp"}).then( () => {
          Services.console.logStringMessage("success backup rule to " + fileName);
          let entries = [];
          let iterator = new OS.File.DirectoryIterator(self.rulesBackupPath);
          iterator.forEach( entry => {
            if ("winLastWriteDate" in entry)
              entries.push({entry: entry, creationDate: entry.winCreationDate});
            else return OS.File.stat(entry.path).then( info => {
              entries.push({entry:entry, creationDate: info.creationDate});
            });
          }).then( () => {
            iterator.close();
            return Promise.all(
              entries.sort( (a, b) => {
                return b.creationDate - a.creationDate;
              }).filter( (element, index) => {
                return index >= rules_to_keep;
              }).map( entry => {
                Services.console.logStringMessage("remove old backup rule file " + entry.entry.path);
                return OS.File.remove(entry.entry.path);
              })
            );
          }, reason => {
            iterator.close();
            throw reason;
          });
        });
      }).then(null, Cu.reportError);
    } catch (err) { autoArchiveLog.logException(err); }
  },
  loadSavedRules: function(win, callback) {
    try {
      const nsIFilePicker = Ci.nsIFilePicker;
      let fp = Cc["@mozilla.org/filepicker;1"].createInstance(nsIFilePicker);
      fp.init(win, "Saved Rules", nsIFilePicker.modeOpen);
      try {
        fp.displayDirectory = new FileUtils.File(self.rulesBackupPath);
      } catch (err) {autoArchiveLog.logException(err);}
      fp.open( function fpCallback_done(aResult) {
        if (aResult == nsIFilePicker.returnOK) {
          if ( fp.file && fp.file.path && fp.file.isFile() && fp.file.isReadable() ) {
            OS.File.read(fp.file.path).then( (array) => {
              let decoder = new TextDecoder();
              let text = decoder.decode(array);
              return callback(text);
            }, Cu.reportError);
          }
        }
      } );
    } catch (err) {
      autoArchiveLog.logException(err);
    }
  },
  cleanup: function() {
  },
}

let self = autoArchiveUtil;
self.initName();
