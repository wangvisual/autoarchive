// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveService"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm, stack: Cs } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource:///modules/MailUtils.js");
Cu.import("resource:///modules/virtualFolderWrapper.js");
Cu.import("resource://gre/modules/XPCOMUtils.jsm");
Cu.import("resource:///modules/iteratorUtils.jsm"); // import toXPCOMArray
Cu.import("chrome://awsomeAutoArchive/content/aop.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveUtil.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");

let autoArchiveService = {
  timer: Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer), // used to schedule the action, preStart => waitIdle => kicksOff => watchDog, can't use this timer during watchDog(running rule)
  statusListeners: [],
  STATUS_INIT: 0,
  STATUS_HIBERNATE: 1,
  STATUS_SLEEP: 2,
  STATUS_WAITIDLE: 3,
  STATUS_RUN: 4,
  STATUS_FINISH: 5,
  _status: [],
  isExceed: false,
  numOfMessages: 0,
  totalSize: 0,
  summary: {}, // { archive: 10, delete: 10, copy: 1, move 20}
  dry_run: false, // set by 'Dry Run' button, there's another one autoArchivePref.options.dry_run is set by perf
  showStatusText: 0, // only show once for STATUS_HIBERNATE on status bar after change, as we use timer
  preStart: function(time) {
    this.timer.cancel();
    this.cleanupIdleObserver();
    let hibernate = autoArchivePref.options.hibernate;
    if ( hibernate == -1 ) return this.updateStatus(this.STATUS_HIBERNATE, "Schedule Disabled"); // never start schdule
    else if ( hibernate == 0 ) return this.start(time);
    else if ( hibernate < -1 ) {
      if ( hibernate + Math.round(Services.startup.getStartupInfo().main/1000) == 0   ) return this.updateStatus(this.STATUS_HIBERNATE, "Schedule Disabled till Thunderbird restart"); // stop schedule for this session
      else {
        autoArchivePref.setPerf('hibernate', 0);
        return this.start(time);
      }
    } else { // disable for some time
      if ( Date.now() > hibernate*1000 ) {
        autoArchivePref.setPerf('hibernate', 0);
        return this.start(time);
      }
      let date = new Date(hibernate*1000);
      if ( this.showStatusText != hibernate ) this.updateStatus(this.STATUS_HIBERNATE, "Schedule Disabled till " + date.toLocaleDateString() + " " + date.toLocaleTimeString());
      this.showStatusText = hibernate;
      this.timer.initWithCallback( function() {
        if ( autoArchiveLog && self ) self.preStart(time);
      }, 60*1000, Ci.nsITimer.TYPE_ONE_SHOT ); // check every mintues
    }
  },
  start: function(time) {
    this.showStatusText = 0;
    let date = new Date(Date.now() + time*1000);
    this.updateStatus(this.STATUS_SLEEP, "Will wakeup @ " + date.toLocaleDateString() + " " + date.toLocaleTimeString());
    this.timer.initWithCallback( function() {
      if ( autoArchiveLog && self ) self.waitTillIdle();
    }, time*1000, Ci.nsITimer.TYPE_ONE_SHOT );
  },
  idleService: Cc["@mozilla.org/widget/idleservice;1"].getService(Ci.nsIIdleService),
  idleObserver: {
    delay: null,
    observe: function(_idleService, topic, data) {
      // topic: idle, active
      if ( topic == 'idle' && self ) self.doArchive();
    }
  },
  waitTillIdle: function() {
    let idleTime = Math.round(this.idleService.idleTime/1000);
    let needDelay = autoArchivePref.options.idle_delay - idleTime;
    autoArchiveLog.info("Computer already idle for " + idleTime + " seconds");
    if ( needDelay <= 0 ) return self.doArchive();
    this.updateStatus(this.STATUS_WAITIDLE, "Wait for " + needDelay + " more seconds idle time");
    this.idleObserver.delay = needDelay;
    this.idleService.addIdleObserver(this.idleObserver, this.idleObserver.delay); // the notification may delay, as Gecko pool OS every 5 seconds
  },
  cleanupIdleObserver: function() {
    if ( this.idleObserver.delay ) {
      this.idleService.removeIdleObserver(this.idleObserver, this.idleObserver.delay);
      this.idleObserver.delay = null;
    }
  },
  cleanup: function() {
    autoArchiveLog.info("autoArchiveService cleanup");
    this.clear();
    this.statusListeners = [];
    this.timer = null;
    autoArchiveLog.info("autoArchiveService cleanup done");
  },
  removeFolderListener: function(listener) {
    MailServices.mailSession.RemoveFolderListener(listener);
    let index = self.folderListeners.indexOf(listener);
    if ( index >= 0 ) self.folderListeners.splice(index, 1);
    else autoArchiveLog.info("Can't remove FolderListener", 1);
  },
  clear: function() {
    try {
      this.cleanupIdleObserver();
      this.timer.cancel();
      this.folderListeners.forEach( function(listener) {
        self.removeFolderListener(listener);
      } );
      this.hookedFunctions.forEach( function(hooked) {
        hooked.unweave();
      } );
      this.closeAllFoldersDB();
      this.hookedFunctions = [];
      this.rules = [];
      this.ruleIndex = 0;
      this.copyGroups = [];
      this.status = [];
      this.wait4Folders = {};
      this._searchSession = null;
      this.folderListeners = [];
      this.isExceed = this.dry_run = false;
      this.showStatusText = true;
      this.numOfMessages = this.totalSize = this.showStatusText = 0;
      this.dryRunLogItems = [];
      this._status = [];
      this.summary = {};
      if ( this.serverStatus['_listeners_'] ) {
        this.serverStatus['_listeners_'].forEach( function(listener) {
          let url = listener.URI.QueryInterface(Ci.nsIMsgMailNewsUrl);
          url.UnRegisterListener(listener);
          autoArchiveLog.info("UnRegister server verify listener for server " + listener.URI.prePath);
        } );
        delete this.serverStatus['_listeners_'];
      }
    } catch(err) { autoArchiveLog.logException(err); }
  },
  rules: [],
  wait4Folders: {},
  accessedFolders: {}, // in onSearchHit, we check if the message exists in dest folder and open it's DB, need to null them later, also we need to null other folders
  closeAllFoldersDB: function() {
    // null all msgDatabase to prevent memory leak, TB might close it later too (https://bugzilla.mozilla.org/show_bug.cgi?id=723248), but just in case user set a very long timeout value
    if ( Object.keys(this.accessedFolders).length ) autoArchiveLog.info("autoArchiveService closeAllFoldersDB");
    Object.keys(this.accessedFolders).forEach( function(uri) {
      try {
        let folder = MailUtils.getFolderForURI(uri);
        if ( !MailServices.mailSession.IsFolderOpenInWindow(folder) && !(folder.flags & (Ci.nsMsgFolderFlags.Trash | Ci.nsMsgFolderFlags.Inbox | Ci.nsMsgFolderFlags.SentMail )) ) {
          autoArchiveLog.info("close msgDatabase for " + uri);
          folder.msgDatabase = null;
        } else {
          autoArchiveLog.info("not close msgDatabase for " + uri);
        }
      } catch(err) { autoArchiveLog.logException(err); }
    } );
    this.accessedFolders = {};
  },
  starStopNow: function(rules, dry_run) {
    if (this._status && this._status[0] == this.STATUS_RUN) autoArchiveService.reportSummaryAndStartNext();
    else autoArchiveService.doArchive(rules, dry_run);
  },
  doArchive: function(rules, dry_run) {
    try {
      autoArchiveLog.info("autoArchiveService doArchive");
      this.clear();
      this.serverStatus = {}; // won't cache between runs
      this.rules = autoArchivePref.validateRules(rules).filter( function(rule) {
        return rule.enable;
      } );
      if ( dry_run ) this.dry_run = true;
      this.updateStatus(this.STATUS_RUN, "Total " + this.rules.length + " rule(s)", this.ruleIndex, this.rules.length);
      autoArchiveLog.logObject(this.rules, 'this.rules',1);
      this.doMoveOrArchiveOne();
    } catch(err) { autoArchiveLog.logException(err); }
  },
  folderListeners: [], // may contain dynamic ones
  folderListener: {
    called: false, // to prevent call doMoveOrArchiveOne etc more than once
    OnItemEvent: function(folder, event) {
      if ( event.toString() != "FolderLoaded" || !folder || !('URI' in folder) ) return;
      let updateNext = false;
      if ( folder.URI ) autoArchiveLog.info("FolderLoaded " + folder.URI);
      else {
        updateNext = true; // the kick off one
        this.called = false;
      }
      if ( self.wait4Folders[folder.URI] ) { // might not be the one we request to update
        delete self.wait4Folders[folder.URI];
        updateNext = true;
      }
      if ( updateNext ) {
        for ( let uri in self.wait4Folders ) {
          let success = false;
          try {
            autoArchiveLog.info("updateFolder " + uri);
            let folder = MailUtils.getFolderForURI(uri);
            if ( folder && folder.parent && folder != folder.rootFolder ) {
              folder.updateFolder(null); // this might call folderLIstener.OnItemEvent sync!
              success = true;
            }
          } catch(err) { autoArchiveLog.logException(err); }
          if ( !success ) delete self.wait4Folders[uri]; // and try update next folder
          else break; // wait for OnItemEvent called again
        }
      }

      if ( Object.keys(self.wait4Folders).length == 0 && !this.called ) {
        this.called = true;
        self.removeFolderListener(self.folderListener);
        autoArchiveLog.info("All FolderLoaded");
        self.doMoveOrArchiveOne();
      }
    },
  },
  copyListener: function(group) { // this listener is for Copy/Delete/Move actions
    this.QueryInterface = XPCOMUtils.generateQI([Ci.nsISupports, Ci.nsIMsgCopyServiceListener, Ci.nsIMsgFolderListener]);
    this.OnStartCopy = function() {
      autoArchiveLog.info("OnStart " + group.action);
      let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
      try {
        if ( mail3PaneWindow && mail3PaneWindow.gFolderDisplay && mail3PaneWindow.gFolderDisplay.view && mail3PaneWindow.gFolderDisplay.view.dbView ) {
          mail3PaneWindow.gFolderDisplay.hintMassMoveStarting();
          mail3PaneWindow.gFolderDisplay._nextViewIndexAfterDelete = null; // then when call hintMassMoveCompleted, no selection change
        }
      } catch(err) { autoArchiveLog.logException(err); }
    };
    this.OnProgress = function(aProgress, aProgressMax) {
      //autoArchiveLog.info("OnProgress " + aProgress + "/"+ aProgressMax);
    };
    this.OnStopCopy = function(aStatus) {
      autoArchiveLog.info("OnStop " + group.action + " 0x" + aStatus.toString(16));
      if ( aStatus ) autoArchiveLog.log(autoArchiveUtil.Name + ": " + group.action + " failed with " + autoArchiveUtil.getErrorMsg(aStatus), "Error!");
      else self.summary[group.action] = ( self.summary[group.action] || 0 ) + group.messages.length;
      if ( group.action == 'delete' || group.action == 'move' ) self.wait4Folders[group.src] = true;
      if ( group.action == 'copy' || group.action == 'move' ) self.wait4Folders[group.dest] = true;
      let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
      try {
        if ( mail3PaneWindow && mail3PaneWindow.gFolderDisplay && mail3PaneWindow.gFolderDisplay.view && mail3PaneWindow.gFolderDisplay.view.dbView ) mail3PaneWindow.gFolderDisplay.hintMassMoveCompleted();
      } catch(err) { autoArchiveLog.logException(err); }
      if ( self.copyGroups.length ) self.doCopyDeleteMoveOne(self.copyGroups.shift());
      else self.updateFolders();
    };
    this.SetMessageKey = function(aKey) {};
    this.GetMessageId = function() {};
    let numberOfMessages = group.messages.length;
    this.msgsDeleted = function(aMsgList) { // Ci.nsIMsgFolderListener, for realDelete message, thus can't get onStopCopy/msgsMoveCopyCompleted
      autoArchiveLog.info("msgsDeleted");
      self.wait4Folders[group.src] = true;
      for (let iMsgHdr = 0; iMsgHdr < aMsgList.length; iMsgHdr++) {
        let msgHdr = aMsgList.queryElementAt(iMsgHdr, Ci.nsIMsgDBHdr);
        let index = group.messages.indexOf(msgHdr);
        if ( index >= 0 ) group.messages.splice(index, 1);
      }
      if ( group.messages.length == 0 ) {
        autoArchiveLog.info("All msgsDeleted");
        MailServices.mfn.removeListener(this);
        // if server connection fail, still report msgs deleted, but delete will fail, this is a BUG
        self.summary[group.action] = ( self.summary[group.action] || 0 ) + numberOfMessages;
        if ( self.copyGroups.length ) self.doCopyDeleteMoveOne(self.copyGroups.shift());
        else self.updateFolders();
      }
    };
  },
  
  getFoldersFromWait4Folders: function() {
    let folders = [];
    Object.keys(self.wait4Folders).forEach( function(uri) {
      let folder;
      try {
        folder = MailUtils.getFolderForURI(uri);
      } catch(err) { autoArchiveLog.logException(err); }
      if ( folder ) {
        if ( self.wait4Folders[uri] != 2 ) folders.push(folder); // if action is copy, do NOT check if server is online or not
      } else delete self.wait4Folders[uri];
    } );
    return folders;
  },

  // Check if all servers in current rule is online before updateFolders
  serverStatus: {
    // key: { OK: true, time: Date.now() },
    // _bad_: false
    // _listeners_: []
  },
  serverListener: function(key) {
    this.key = key;
    this.URI = null; // will be replaced by real URI
    this.OnStartRunningUrl = function(uri) {};
    this.OnStopRunningUrl = function(uri, aExitCode) {
      if ( !self || !autoArchiveLog ) return;
      let index = self.serverStatus['_listeners_'].indexOf(this);
      if ( index >= 0 ) self.serverStatus['_listeners_'].splice(index, 1);
      autoArchiveLog.info("OnStopRunningUrl: server " + this.key);
      if ( Components.isSuccessCode(aExitCode) ) autoArchiveLog.info(uri.prePath + " OK");
      else {
        self.serverStatus['_bad_'] = true;
        autoArchiveLog.log("Check Mail Server, got " + autoArchiveUtil.getErrorMsg(aExitCode) + "(" + aExitCode.toString(16) + ") for " + uri.prePath, 1);
      }
      self.serverStatus[this.key] = { OK: Components.isSuccessCode(aExitCode), time: Date.now() };
      if ( self.serverStatus['_listeners_'].length == 0 ) {
        autoArchiveLog.info("All servers checking done, has bad server? : " + self.serverStatus['_bad_']);
        if ( self.serverStatus['_bad_'] ) {
          self._searchSession = null;
          return self.doMoveOrArchiveOne(); // next rule
        } else return self.updateFolders();
      }
    };
  },
  checkServers: function() {
    let servers = {}, hasBad = false;
    self.serverStatus['_bad_'] = false;
    self.serverStatus['_listeners_'] = [];
    this.getFoldersFromWait4Folders().some( function(folder) {
      let server = folder.server, needCheck = false;
      if ( ['none', 'nntp', 'rss', 'exquilla'].indexOf(server.type) < 0 && !servers[server.key] ) {
        if ( Services.io.offline ) {
          autoArchiveLog.log("Skip rule due to offline now");
          return ( hasBad = true ); // break 'some'
        }
        if ( self.serverStatus[server.key] ) {
          if ( Date.now() - self.serverStatus[server.key]['time'] > ( self.serverStatus[server.key]['OK'] ? 180000 : 60000 ) ) { // positive cache 3 minutes, neg cache 1 minute
            autoArchiveLog.info("Need re-check server:" + server.prettyName);
            delete self.serverStatus[server.key];
            needCheck = true;
          } else if ( !self.serverStatus[server.key]['OK'] ) {
            autoArchiveLog.log("Skip bad server " + server.prettyName);
            return ( hasBad = true );
          }
        } else needCheck = true;
      }
      if ( needCheck ) {
        if ( !Services.logins.isLoggedIn || server.passwordPromptRequired ) return ( hasBad = true );
        autoArchiveLog.info("needCheck mail server: " + server.prettyName);
        // serverBusy means we already getting new Messages
        if ( server.serverBusy || ( server instanceof Ci.nsIPop3IncomingServer && server.runningProtocol ) )
          self.serverStatus[server.key] = { OK: true, time: Date.now() };
        else servers[server.key] = server;
      }
      return false;
    } );
    if ( hasBad ) {
      self._searchSession = null;
      return self.doMoveOrArchiveOne(); // next rule
    }
    let count = 0;
    for ( let key in servers ) {
      try {
        let server = servers[key], listener = new self.serverListener(key), URI;
        // verifyLogon will zero popstate.dat and cause duplicate mails for POP3
        if ( server instanceof Ci.nsIPop3IncomingServer ) {
          let inbox = server.rootMsgFolder.getFolderWithFlags(Ci.nsMsgFolderFlags.Inbox);
          // server.loginAtStartUp: check for new mail @ startup
          // server.doBiff: check for new mail every n minutes
          // server.downloadOnBiff: automatically download new messages
          let headers_only = false;
          try {
            headers_only = Services.prefs.getBoolPref("mail.server." + key + ".headers_only");
          } catch(err) {}
          let download = ( server.loginAtStartUp || server.doBiff ) && server.downloadOnBiff && !headers_only;
          autoArchiveLog.info( ( download ? "downloading" : "checking" ) + " emails for server " + server.prettyName);
          URI = download ? MailServices.pop3.GetNewMail(null, listener, inbox, server) : MailServices.pop3.CheckForNewMail(null, listener, inbox, server);
        } else URI = server.verifyLogon(listener, null);
        self.serverStatus['_listeners_'].push(listener); // the listener can be unregistered if clear / stop
        listener.URI = URI;
        autoArchiveLog.info("Checking if server " + key + " on line using " + URI.spec);
        count ++;
      } catch(err) { autoArchiveLog.logException(err); }
    }
    if ( count == 0 ) return self.updateFolders(); // continue to update folder and then search
  },
  
  // updateFolders may get called before when we run search ( when _searchSession was set )
  // or get called after we doing one group of Move/Delete/Copy, or one Archive ( when _searchSession was null )
  // any case we will chain doMoveOrArchiveOne here or in folderListener, and let it to decide either start process a new rule, or continue to search
  updateFolders: function() {
    MailServices.mailSession.AddFolderListener(self.folderListener, Ci.nsIFolderListener.event);
    self.folderListeners.push(self.folderListener);
    self.folderListener.OnItemEvent({URI:''}, 'FolderLoaded'); // we fake one to kicks off the update sequence
  },

  copyGroups: [], // [ {src: src, dest: dest, action: move, messages[]}, ...]
  hookedFunctions: [],
  searchListener: function(rule, srcFolder, destFolder) {
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    if (!mail3PaneWindow) return self.doMoveOrArchiveOne();
    this.QueryInterface = XPCOMUtils.generateQI([Ci.nsISupports, Ci.nsIFolderListener, Ci.nsIMsgSearchNotify]);
    this.messages = [];
    this.missingFolders = {};
    this.messagesDest = {};
    let allTags = {};
    let searchHit = 0;
    let duplicateHit = [];
    let skipReason = { duplicate: 0, exceed: 0, cantDelete: 0, deleted: 0, srcLocked: 0, destLocked: 0, flaged: 0, unread: 0, tags: 0, cantArchive: 0, offline: 0 };
    let actionSize = 0;
    let listener = this;
    for ( let tag of MailServices.tags.getAllTags({}) ) {
      allTags[tag.key.toLowerCase()] = true;
    };
    this.hasTag = function(msgHdr) {
      return msgHdr.getStringProperty('keywords').toLowerCase().split(' ').some( function(key) { // may contains X-Keywords like NonJunk etc
        return ( key in allTags );
      } );
    };
    this.onSearchHit = function(msgHdr, folder) {
      //autoArchiveLog.info("search hit message:" + msgHdr.mime2DecodedSubject);
      //let str = ''; let e = msgHdr.propertyEnumerator; let str = "property:\n"; while ( e.hasMore() ) { let k = e.getNext(); str += k + ":" + msgHdr.getStringProperty(k) + "\n"; }; autoArchiveLog.info(str);
      searchHit ++;
      if ( self.isExceed ) return skipReason.exceed++;
      if ( self.numOfMessages > 0 // if only one big message exceed the size limit, we still accept it
        && ( ( autoArchivePref.options.messages_number_limit > 0 && self.numOfMessages >= autoArchivePref.options.messages_number_limit )
          || ( autoArchivePref.options.messages_size_limit > 0 && self.totalSize + msgHdr.messageSize > autoArchivePref.options.messages_size_limit * 1024 * 1024 ) ) ) {
        self.isExceed = true;
        return skipReason.exceed++;
      }
      if ( !msgHdr.messageId || !msgHdr.folder || !msgHdr.folder.URI || msgHdr.folder.URI == rule.dest ) return;
      if ( ['delete', 'move'].indexOf(rule.action) >= 0 && !msgHdr.folder.canDeleteMessages ) return skipReason.cantDelete++;
      if ( msgHdr.flags & (Ci.nsMsgMessageFlags.Expunged|Ci.nsMsgMessageFlags.IMAPDeleted) ) return skipReason.deleted++;
      let age = ( Date.now() / 1000 - msgHdr.dateInSeconds ) / 3600 / 24;
      if ( ["move", "delete", "archive"].indexOf(rule.action) >= 0 ) {
        if ( msgHdr.folder.locked ) return skipReason.srcLocked++;
        if ( msgHdr.isFlagged && ( !autoArchivePref.options.enable_flag || age < autoArchivePref.options.age_flag ) ) return skipReason.flaged++;
        if ( !msgHdr.isRead && ( !autoArchivePref.options.enable_unread || age < autoArchivePref.options.age_unread ) ) return skipReason.unread++;
        if ( typeof(rule.tags) == 'undefined' && this.hasTag(msgHdr) && ( !autoArchivePref.options.enable_tag || age < autoArchivePref.options.age_tag ) ) return skipReason.tags++;
      }
      if ( rule.action == 'archive' ) {
        if ( self.folderIsOf(msgHdr.folder, Ci.nsMsgFolderFlags.Archive) ) return skipReason.cantArchive++;
        let getIdentityForHeader = mail3PaneWindow.getIdentityForHeader || mail3PaneWindow.GetIdentityForHeader; // TB & SeaMonkey use different name
        if ( !getIdentityForHeader || !getIdentityForHeader(msgHdr).archiveEnabled ) return skipReason.cantArchive++;
      }
      
      if ( Services.io.offline && msgHdr.folder.server && msgHdr.folder.server.type != 'none' ) return skipReason.offline++; // https://bugzilla.mozilla.org/show_bug.cgi?id=956598
      if ( ["copy", "move"].indexOf(rule.action) >= 0 ) {
        // check if dest folder has already has the message
        let realDest = rule.dest, additonal = '', additonalNames = [];
        let supportHierarchy = ( rule.sub == 2 ) && !srcFolder.getFlag(Ci.nsMsgFolderFlags.Virtual) && !destFolder.getFlag(Ci.nsMsgFolderFlags.Virtual) && destFolder.canCreateSubfolders;
        if ( supportHierarchy && (destFolder.server instanceof Ci.nsIImapIncomingServer)) supportHierarchy = !destFolder.server.isGMailServer;
        if ( supportHierarchy ) {
          // for local folder like URI (string) 'mailbox://nobody@Local%20Folders/test/hello%20world', the name of leaf level will be 'hello world'
          // but IMAP folders don't use %20,
          // when we create folders, we can only use 'hello world', never use 'hello%20world'
          let tmpFolder = msgHdr.folder, rootFolder = tmpFolder.server.rootFolder;
          while ( tmpFolder && tmpFolder != srcFolder && tmpFolder != rootFolder ) { // this loop might run each time for every messages hit
            additonalNames.unshift(tmpFolder.name); // [ 'test', 'hello world' ]
            tmpFolder = tmpFolder.parent;
          }
          if ( additonalNames.length ) additonal = '/' + additonalNames.join('/');
          realDest = rule.dest + additonal;
        }
        //autoArchiveLog.info(msgHdr.mime2DecodedSubject + " : " + msgHdr.folder.URI + " => " + realDest);
        let realDestFolder = MailUtils.getFolderForURI(realDest);
        if ( Services.io.offline && realDestFolder.server && realDestFolder.server.type != 'none' ) return skipReason.offline++;
        if ( realDestFolder.locked ) return skipReason.destLocked++;
        // BatchMessageMover using createStorageIfMissing/createSubfolder
        // CopyFolders using createSubfolder
        // https://github.com/gark87/SmartFilters/blob/master/src/chrome/content/backend/imapfolders.jsm using createSubfolder
        // https://github.com/mozilla/releases-comm-central/blob/master/mailnews/imap/test/unit/test_localToImapFilter.js using CopyFolders, but it's empty folders
        // http://thunderbirddocs.blogspot.com/2005/12/mozilla-thunderbird-creating-folders.html
        // http://mxr.mozilla.org/comm-central/source/mailnews/imap/src/nsImapMailFolder.cpp
        // http://mxr.mozilla.org/comm-central/source/mailnews/local/src/nsLocalMailFolder.cpp
        // If target folder already exists but not subscribed, sometimes createStorageIfMissing will not trigger OnStopRunningUrl
        
        // msgDatabase is a getter that will always try and load the message database! so null it if not use if any more
        let destHdr, msgDatabase, offlineStream;
        try {
          //autoArchiveLog.logObject(realDestFolder,'realDestFolder',0); // Don't do this, access some property may automatically create wrong type of local mail folder, eg
          // "expungedBytes","flags","sortOrder","name","prettyName","prettiestName","abbreviatedName","subFolders","descendants"
          msgDatabase = realDestFolder.msgDatabase; // exception when folder not exists
          self.accessedFolders[realDest] = 1;
          destHdr = msgDatabase.getMsgHdrForMessageID(msgHdr.messageId);
          offlineStream = realDestFolder.offlineStoreInputStream;
        } catch(err) {
          // 0x80004005 (NS_ERROR_FAILURE)
          // 0x80520012 (NS_ERROR_FILE_NOT_FOUND) [nsIMsgFolder.offlineStoreInputStream]
          // 0x80550006 [nsIMsgFolder.msgDatabase]
          // 0X80520015 (NS_ERROR_FILE_ACCESS_DENIED) [nsIMsgFolder.offlineStoreInputStream]
          if ( [0x80004005, 0x80520012, 0x80550006, 0X80520015].indexOf(err.result) < 0  ) autoArchiveLog.logException(err, 0);
        }
        if ( offlineStream && msgDatabase && !autoArchiveUtil.folderExists(realDestFolder) && destFolder.msgStore ) {
          // may false alarm, but just keep here
          autoArchiveLog.info("Found hidden folder '" + realDestFolder.URI + "', update folder tree");
          destFolder.msgStore.discoverSubFolders(destFolder, true);
          if ( mail3PaneWindow.gFolderTreeView && mail3PaneWindow.gFolderTreeView._rebuild ) mail3PaneWindow.gFolderTreeView._rebuild();
        }
        if ( destHdr ) {
          //autoArchiveLog.info("Message:" + msgHdr.mime2DecodedSubject + " already exists in dest folder");
          duplicateHit.push(destHdr);
          return skipReason.duplicate++;
        } else if ( !autoArchiveUtil.folderExists(realDestFolder) && !offlineStream && !(realDestFolder.URI in this.missingFolders) ) { // sometime when TB has issue, folder.parent is null but getMsgHdrForMessageID can return hdr
          //autoArchiveLog.info("dest folder " + realDest + " not exists, need create");
          this.missingFolders[realDestFolder.URI] = additonalNames;
        }
        this.messagesDest[msgHdr.folder.URI + " _ " + msgHdr.messageId] = realDest;
      }
      //autoArchiveLog.info("add message:" + msgHdr.mime2DecodedSubject);
      self.totalSize += msgHdr.messageSize; actionSize += msgHdr.messageSize;
      self.numOfMessages ++;
      this.messages.push(msgHdr);
    };
    this.onSearchDone = function(status) {
      try {
        self._searchSession = null;
        autoArchiveLog.info("Total " + searchHit + " messages hit");
        autoArchiveLog.logObject(skipReason, 'skipReason', 0);
        let isMove = (rule.action == 'move');
        if ( duplicateHit.length ) autoArchiveLog.info(duplicateHit.length + " messages already exists in target folder", isMove, isMove);
        if ( !this.messages.length && !( isMove && autoArchivePref.options.delete_duplicate_in_src ) ) return self.doMoveOrArchiveOne();
        autoArchiveLog.info("will " + rule.action + " " + this.messages.length + " messages, total " + autoArchiveUtil.readablizeBytes(actionSize) + " bytes");
        // create missing folders first
        autoArchiveLog.logObject(this.missingFolders,'Need create these folders',0);
        if ( Object.keys(this.missingFolders).length ) { // for copy/move
          // rule.dest: imap://a@b.com/1/2
          // additionalNames:            [3,4,5]
          //                             [3,4,6]
          //                             [7,8]
          // =>
          // /3, /4, /5
          // (/3, /4), /6
          if ( autoArchivePref.options.dry_run || self.dry_run ) {
            Object.keys(this.missingFolders).forEach( function(uri) {
              self.dryRunLog(["create", uri]);
            } );
            delete this.missingFolders;
          } else {
            if ( autoArchiveUtil.createFolderAsync(destFolder) ) {
              autoArchiveLog.info("create folders async");
              //MailServices.mfn.addListener(this, MailServices.mfn.folderAdded);
              MailServices.mailSession.AddFolderListener(this, Ci.nsIFolderListener.added);
              self.folderListeners.push(this);
            } else autoArchiveLog.info("create folders sync");
            return this.OnItemAdded(); // OnItemAdded will chain to create next folder, CopyMessages can create the folder on the fly, but won't show it, and sometimes failed
          }
        }
        this.processHeaders();
      } catch(err) {
        autoArchiveLog.logException(err);
        return self.doMoveOrArchiveOne();
      }
    };
    this.onNewSearch = function() {};
    this.OnItemAdded = function(parentFolder, childFolder) {
      if ( childFolder && childFolder.QueryInterface ) {
        childFolder.QueryInterface(Ci.nsIMsgFolder);
        if ( childFolder.URI ) {
          autoArchiveLog.info("Folder " + childFolder.URI + " created");
          parentFolder.updateFolder(null);
        }
      }
      for ( let uri in this.missingFolders ) {
        let folder = destFolder, names = this.missingFolders[uri];
        let asyncCreating = names.some( function(folderName) {
          try {
            if ( !folder.containsChildNamed(folderName) ) {
              // if DB is messed-up, then the (wrong) folder might be invisible but there
              autoArchiveLog.info("Creating folder '" + folder.URI + "' => '" + folderName + "'");
              folder.createSubfolder(folderName, null); // 2nd parameter can be mail3PaneWindow.msgWindow to get alert when folder create failed
              if ( autoArchiveUtil.createFolderAsync(destFolder) ) return true; // if async, break 'some'
            }
          } catch(err) { autoArchiveLog.info("create folder '" + path + "' failed, " + err.toString(), "Error!", 1); }
          folder = folder.getChildNamed(folderName); // if folder exists, or sync creation
          return false; // try next level
        } );
        if ( asyncCreating ) return; // waiting for OnItemAdded
        else delete this.missingFolders[uri];
      }
      
      if ( ( 'missingFolders' in this ) && Object.keys(this.missingFolders).length == 0 ) {
        autoArchiveLog.info("All folders created");
        if ( autoArchiveUtil.createFolderAsync(destFolder) ) self.removeFolderListener(this);
        delete this.missingFolders;
        return this.processHeaders();
      }
    };
    this.processHeaders = function() {
      try {
        if ( rule.action != 'archive' ) {
          // group messages according to there src and dest
          self.copyGroups = [];
          let groups = {}; // { src => dest : 0, src2 => dest2: 1 }
          let removingDuplicate = ( rule.action == 'move' && autoArchivePref.options.delete_duplicate_in_src && duplicateHit.length ), actions = {};
          let messages = this.messages;
          if ( removingDuplicate ) messages.push.apply(messages, duplicateHit);
          messages.forEach( function(msgHdr) {
            let dest = listener.messagesDest[msgHdr.folder.URI + " _ " + msgHdr.messageId] || rule.dest || '', action = rule.action;
            if ( removingDuplicate && duplicateHit.indexOf(msgHdr) >= 0 ) {
              dest = '';
              action = 'delete';
            }
            if ( dest.length ) self.accessedFolders[dest] = true;
            let key = msgHdr.folder.URI + ( ["copy", "move"].indexOf(action) >= 0 ? " => " + dest : '' );
            if ( typeof(groups[key]) == 'undefined'  ) {
              groups[key] = self.copyGroups.length;
              self.copyGroups.push({src: msgHdr.folder.URI, dest: dest, action: action, messages: []});
            }
            self.copyGroups[groups[key]].messages.push(msgHdr);
            actions[action] = true;
          } );
          if ( self.copyGroups.length == 0 ) self.doMoveOrArchiveOne();
          else {
            autoArchiveLog.info("will do " + Object.keys(actions).join(' and ') + " in " + self.copyGroups.length + " steps");
            autoArchiveLog.logObject(groups, 'groups', 0);
            autoArchiveLog.logObject(self.copyGroups, 'self.copyGroups', 1);
            self.doCopyDeleteMoveOne(self.copyGroups.shift());
          }
        } else {
          if ( autoArchivePref.options.dry_run || self.dry_run ) {
            this.messages.forEach( function(msgHdr) {
              self.dryRunLog(["archive", msgHdr.mime2DecodedSubject, msgHdr.folder.URI]);
            } );
            return self.doMoveOrArchiveOne();
          }
          self.wait4Folders[rule.src] = true;
          self.summary[rule.action] = ( self.summary[rule.action] || 0 ) + this.messages.length;
          // from mailWindowOverlay.js
          let batchMover = new mail3PaneWindow.BatchMessageMover();
          let myFunc = function(result) {
            autoArchiveLog.info("BatchMessageMover OnStopCopy/OnStopRunningUrl");
            if ( !batchMover.awsome_auto_archive_done && ( batchMover._batches == null || Object.keys(batchMover._batches).length == 0 ) ) {
              autoArchiveLog.info("BatchMessageMover Done");
              batchMover.awsome_auto_archive_done = true; // prevent call doMoveOrArchiveOne twice
              self.hookedFunctions.forEach( function(hooked) {
                hooked.unweave();
              } );
              self.hookedFunctions = [];
              self.updateFolders(); // updateFolders will chain next doMoveOrArchiveOne
              //self.doMoveOrArchiveOne();
            }
            autoArchiveLog.info("BatchMessageMover OnStopCopy/OnStopRunningUrl exit");
            return result;
          }
          self.hookedFunctions.push( autoArchiveaop.after( {target: batchMover, method: 'OnStopCopy'}, myFunc )[0] );
          self.hookedFunctions.push( autoArchiveaop.after( {target: batchMover, method: 'OnStopRunningUrl'}, myFunc )[0] );
          self.hookedFunctions.push( autoArchiveaop.around( {target: batchMover, method: 'processNextBatch'}, function(invocation) {
            autoArchiveLog.info("BatchMessageMover processNextBatch");
            try {
              if ( !batchMover.awsome_auto_archive_getFolders ) {
                for ( let key in batchMover._batches ) {
                  // before https://bugzilla.mozilla.org/show_bug.cgi?id=975795, batchMover._batches = { key: [folder, URI, ..., monthFolderName, msghdr1, msghdr2,...], ... }
                  // after, batchMover._batches = { key: srcFolder: msgHdr.folder, ...,  monthFolderName: monthFolderName, messages: [...] }, ... }
                  let { srcFolder: srcFolder, archiveFolderURI: archiveFolderURI, granularity: granularity, keepFolderStructure: keepFolderStructure, yearFolderName: msgYear, monthFolderName: msgMonth } = batchMover._batches[key];
                  if ( !('srcFolder' in batchMover._batches[key]) ) // instanceof Array won't work, Array.isArray() should, see http://web.mit.edu/jwalden/www/isArray.html 
                    [srcFolder, archiveFolderURI, granularity, keepFolderStructure, msgYear, msgMonth] = batchMover._batches[key];
                  let archiveFolder = MailUtils.getFolderForURI(archiveFolderURI, false);
                  let forceSingle = !archiveFolder.canCreateSubfolders;
                  if (!forceSingle && (archiveFolder.server instanceof Ci.nsIImapIncomingServer)) forceSingle = archiveFolder.server.isGMailServer;
                  if (forceSingle) granularity = Ci.nsIMsgIncomingServer.singleArchiveFolder;
                  if (granularity >= Ci.nsIMsgIdentity.perYearArchiveFolders) archiveFolderURI += "/" + msgYear;
                  if (granularity >= Ci.nsIMsgIdentity.perMonthArchiveFolders) archiveFolderURI += "/" + msgMonth;
                  if (archiveFolder.canCreateSubfolders && keepFolderStructure) {
                    // .../Inbox/test/a/b/c => .../archive/2014/01/test/a/b/c
                    // .../another/test/d => .../archive/2014/01/another/test/d
                    let rootFolder = srcFolder.server.rootFolder;
                    let inboxFolder = srcFolder.server.rootMsgFolder.getFolderWithFlags(Ci.nsMsgFolderFlags.Inbox);
                    let folder = srcFolder, folderNames = [];
                    while (folder != rootFolder && folder != inboxFolder) {
                      folderNames.unshift(folder.name);
                      folder = folder.parent;
                    }
                    archiveFolderURI += "/" + folderNames.join('/');
                  }
                  autoArchiveLog.info("add update folders " + srcFolder.URI + " => " + archiveFolderURI);
                  self.wait4Folders[srcFolder.URI] = self.wait4Folders[archiveFolderURI] = self.accessedFolders[archiveFolderURI] = true;
                }
                batchMover.awsome_auto_archive_getFolders = true;
              }
            } catch(err) { autoArchiveLog.logException(err); }
            let result = invocation.proceed();
            myFunc(result);
            autoArchiveLog.info("BatchMessageMover processNextBatch exit");
            return result;
          } )[0] );
          if ( mail3PaneWindow.gFolderDisplay ) self.hookedFunctions.push( autoArchiveaop.after( {target: mail3PaneWindow.gFolderDisplay, method: 'hintMassMoveStarting'}, function(result) {
            mail3PaneWindow.gFolderDisplay._nextViewIndexAfterDelete = null; // prevent selection change
            return result;
          } )[0] );
          autoArchiveLog.info("Start doing archive");
          batchMover.archiveMessages(this.messages); // exceptions should be caught below
        }
      } catch(err) {
        autoArchiveLog.logException(err);
        return self.doMoveOrArchiveOne();
      }
    };
  },
  doCopyDeleteMoveOne: function(group) {
    function runNext() {
      if ( self.copyGroups.length ) self.doCopyDeleteMoveOne(self.copyGroups.shift());
      else self.doMoveOrArchiveOne();
    }
    if ( autoArchivePref.options.dry_run || this.dry_run ) {
      group.messages.forEach( function(msg) {
        self.dryRunLog([group.action, msg.mime2DecodedSubject, group.src , ( group.action != 'delete' ? group.dest : '' )])
      } );
      return runNext();
    }
    let srcFolder = MailUtils.getFolderForURI(group.src), destFolder = null;
    let busy = ( ['move', 'delete'].indexOf(group.action) >= 0 ) ? srcFolder.locked : false;
    if ( group.action != 'delete' ) {
      destFolder = MailUtils.getFolderForURI(group.dest);
      busy = busy || destFolder.locked;
    }
    if ( busy ) {
      autoArchiveLog.info("Folder busy, skip " + group.action + " " + group.src + " => " + group.dest);
      return runNext();
    }
    let xpcomHdrArray = toXPCOMArray(group.messages, Ci.nsIMutableArray);
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    let msgWindow = null;
    if ( mail3PaneWindow ) msgWindow = mail3PaneWindow.msgWindow;
    if ( group.action == 'delete' ) {
      // deleteMessages impacted by srcFolder.server.getIntValue('delete_model')
      // 0:mark as deleted, 1:move to trash, 2:remove it immediately
      let deleteModel = srcFolder.server.getIntValue('delete_model');
      autoArchiveLog.info('deleteModel ' + deleteModel);
      let isTrashFolder = srcFolder.getFlag(Ci.nsMsgFolderFlags.Trash); // sub folder of Trash is not Trash...
      if ( !isTrashFolder ) {
        let trashFolder = srcFolder.rootFolder.getFolderWithFlags(Ci.nsMsgFolderFlags.Trash);
        if ( trashFolder && (trashFolder == srcFolder || trashFolder.folderURL == srcFolder.folderURL || trashFolder.folderURL == srcFolder.URI || trashFolder.URI == srcFolder.folderURL || trashFolder.URI == srcFolder.URI) ) {
          isTrashFolder = true;
        }
      }
      autoArchiveLog.info('srcFolder:' + srcFolder.URI + " is trash? " + isTrashFolder);
      // http://code.google.com/p/reply-manager/source/browse/mailnews/base/test/unit/test_nsIMsgFolderListenerLocal.js
      // if (!isMove && (deleteStorage || isTrashFolder)) => msgsDeleted
      // else => msgsMoveCopyCompleted, onStopCopy
      let realDelete = deleteModel == 2 || isTrashFolder;
      if ( realDelete ) MailServices.mfn.addListener(new self.copyListener(group), MailServices.mfn.msgsDeleted );
      srcFolder.deleteMessages(xpcomHdrArray, null, /*deleteStorage*/realDelete, /*isMove*/false, realDelete ? null : new self.copyListener(group), /* allow undo */false);
      return;
    }
    let isMove = (group.action == 'move') && srcFolder.canDeleteMessages;
    try {
      MailServices.copy.CopyMessages(srcFolder, xpcomHdrArray, destFolder, isMove, new self.copyListener(group), /*msgWindow*/msgWindow, /* allow undo */false);
    } catch (err) {autoArchiveLog.logException(err);}
  },
  _searchSession: null,
  reportSummaryAndStartNext: function() {
    autoArchiveLog.info("Proposed to change " + this.numOfMessages + " messages, " + autoArchiveUtil.readablizeBytes(this.totalSize) + " bytes");
    let total = 0, report = [];
    for ( let action in this.summary ) {
      total += this.summary[action];
      report.push(this.summary[action] + " " + ( action == 'copy' ? 'copied' : action + "d"));
    }
    if ( autoArchivePref.options.dry_run || self.dry_run ) {
      let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
      // openDialog with dialog=no, as open can't have additional parameter, and dialog has no maximize button
      if ( mail3PaneWindow ) mail3PaneWindow.openDialog("chrome://awsomeAutoArchive/content/autoArchiveInfo.xul", "_blank", "chrome,modal,resizable,centerscreen,dialog=no", this.dryRunLogItems);
      else Services.prompt.select(null, 'Dry Run', 'These changes would be applied in real run:', this.dryRunLogItems.length, this.dryRunLogItems, {});
    } else if ( total != this.numOfMessages ) autoArchiveLog.info("Real change " + total + " messages, some folders might be busy.");
    autoArchiveLog.info( self.isExceed ? "Limitation reached, set next" : "auto archive done for all rules, set next");
    this.updateStatus(this.STATUS_FINISH, total == 0 ? "Archie: Nothing done" : "Archie: Processed " + total + " msgs (" + report.join(", ") + ")");
    let delay = this.isExceed ? autoArchivePref.options.start_exceed_delay : autoArchivePref.options.start_next_delay;
    this.clear();
    return this.preStart(delay);
  },
  
  doMoveOrArchiveOne: function() {
    this.timer.initWithCallback( function() { //
      return self._doMoveOrArchiveOne();
    }, 0, Ci.nsITimer.TYPE_ONE_SHOT );
  },
  
  _doMoveOrArchiveOne: function() {
    if ( this._searchSession ) { // updateFolder done, continue to search now
      autoArchiveLog.info(autoArchiveUtil.getSearchSessionString(this._searchSession));
      this._searchSession.search(null);
      return this._searchSession = null;
    }
    //[{"src": "xx", "dest": "yy", "action": "move", "age": 180, "sub": 1, "subject": /test/i, "from": who, "recipient": whom, "size": 100000, "tags": "!important", "enable": true}]
    if ( this.ruleIndex >= this.rules.length || self.isExceed ) {
      this.closeAllFoldersDB();
      return this.reportSummaryAndStartNext();
    }

    let rule = this.rules[this.ruleIndex];
    //autoArchiveLog.logObject(rule, 'running rule', 1);
    this.updateStatus(this.STATUS_RUN, "Running rule " + rule.action + " " + rule.src + ( ["move", "copy"].indexOf(rule.action)>=0 ? " to " + rule.dest : "" ) + " with filter { "
      + ["age", "subject", "from", "recipient", "size", "tags"].filter( function(item) {
        return typeof(rule[item]) != 'undefined';
      } ).map( function(item) {
        return item + " => " + rule[item];
      } ).join(", ")
      + " }", this.ruleIndex++, this.rules.length); // ruleIndex will ++ after this
    this.timer.initWithCallback( function() { // watch dog, will be reset by next doMoveOrArchiveOne watch dog or start
      autoArchiveLog.log("Timeout when " + self._status[1], 1);
      return self.doMoveOrArchiveOne(); // call doMoveOrArchiveOne might make me crazy, but it can make sure all rules have chance to run
      //return self.stop(); // this will be much safe, however, all rules below will not run.
    }, autoArchivePref.options.rule_timeout * 1000, Ci.nsITimer.TYPE_ONE_SHOT );
    let srcFolder = null, destFolder = null;
    try {
      srcFolder = MailUtils.getFolderForURI(rule.src);
      if ( ["move", "copy"].indexOf(rule.action) >= 0 ) {
        destFolder = MailUtils.getFolderForURI(rule.dest);
        if ( autoArchiveUtil.folderExists(destFolder) ) {
          self.wait4Folders[rule.dest] = self.accessedFolders[rule.dest] = true;
          if ( rule.sub == 2 ) {
            let folders = destFolder.descendants /* >=TB21 */;
            for (let folder in fixIterator(folders || [], Ci.nsIMsgFolder)) {
              self.wait4Folders[folder.URI] = self.accessedFolders[folder.URI] = true;
            }
          }
        }
      } else rule.dest = '';
    } catch (err) {
      autoArchiveLog.logException(err);
    }
    if ( !autoArchiveUtil.folderExists(srcFolder) || ( ["move", "copy"].indexOf(rule.action) >= 0 && !autoArchiveUtil.folderExists(destFolder) ) ) {
      autoArchiveLog.log("Error: Wrong rule because folder does not exist: " + rule.src + ( ["move", "copy"].indexOf(rule.action) >= 0 ? ' or ' + rule.dest : '' ), 'Error!');
      return this.doMoveOrArchiveOne();
    }
    //srcFolder.server.closeCachedConnections();
    if ( rule.action == 'archive' ) { // mare sure we have at least one folder show, or hintMassMoveStarting will throw exception
      let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
      if ( mail3PaneWindow && mail3PaneWindow.gFolderDisplay && mail3PaneWindow.gFolderDisplay.view && !mail3PaneWindow.gFolderDisplay.view.dbView ) {
        autoArchiveLog.info("no folders selected, hintMassMoveStarting might fail, so select source folder");
        mail3PaneWindow.SelectFolder(rule.src);
      }
    }

    let searchSession = Cc["@mozilla.org/messenger/searchSession;1"].createInstance(Ci.nsIMsgSearchSession);
    if (srcFolder.flags & Ci.nsMsgFolderFlags.Virtual) { // searchSession.addScopeTerm won't work for virtual folder 
      let virtFolder = VirtualFolderHelper.wrapVirtualFolder(srcFolder);
      let scope = virtFolder.onlineSearch ? Ci.nsMsgSearchScope.onlineMail : Ci.nsMsgSearchScope.offlineMail;
      virtFolder.searchFolders.forEach( function(folder) {
        if ( rule.action == 'archive' && self.folderIsOf(folder, Ci.nsMsgFolderFlags.Archive) ) return;
        autoArchiveLog.info("Add src folder " + folder.URI);
        searchSession.addScopeTerm(scope, folder);
        self.accessedFolders[folder.URI] = true;
        self.wait4Folders[folder.URI] = (rule.action == 'copy' ? 2 : true); // when copy, don't check if server is online or not
      } );
      let terms = virtFolder.searchTerms;
      //for (let term in fixIterator(terms, Ci.nsIMsgSearchTerm)) {
      let count = terms.Count();
      if ( count ) {
        for ( let i = 0; i < count; i++ ) {
          let term = terms.GetElementAt(i).QueryInterface(Ci.nsIMsgSearchTerm);
          if ( i == 0 ) term.beginsGrouping = true;
          if ( i+1 == count ) term.endsGrouping = true;
          searchSession.appendTerm(term);
        }
      }
    } else {
      searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, srcFolder);
      self.accessedFolders[rule.src] = true;
      self.wait4Folders[rule.src] = (rule.action == 'copy' ? 2 : true);
      if ( rule.sub ) {
        let folders = srcFolder.descendants /* >=TB21 */;
        for (let folder in fixIterator(folders || [], Ci.nsIMsgFolder)) {
          // We don't add special sub directories, same as AutoarchiveReloaded
          if ( folder.getFlag(Ci.nsMsgFolderFlags.Virtual) ) continue;
          if ( autoArchivePref.options.ignore_spam_folders && ["move", "archive", "copy"].indexOf(rule.action) >= 0 &&
            folder.getFlag(Ci.nsMsgFolderFlags.Trash | Ci.nsMsgFolderFlags.Junk| Ci.nsMsgFolderFlags.Queue | Ci.nsMsgFolderFlags.Drafts | Ci.nsMsgFolderFlags.Templates ) ) continue;
          if ( rule.action == 'archive' && self.folderIsOf(folder, Ci.nsMsgFolderFlags.Archive) ) continue;
          searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, folder);
          self.accessedFolders[folder.URI] = true;
          self.wait4Folders[folder.URI] = (rule.action == 'copy' ? 2 : true);
        }
      }
    }
    
    if ( rule.age ) autoArchiveUtil.addSearchTerm(searchSession, Ci.nsMsgSearchAttrib.AgeInDays, rule.age, Ci.nsMsgSearchOp.IsGreaterThan);

    // expressionsearch has logic to deal with Ci.nsMsgMessageFlags.HasRe, use it first
    let normal = { subject: Ci.nsMsgSearchAttrib.Subject, from: Ci.nsMsgSearchAttrib.Sender, recipient: Ci.nsMsgSearchAttrib.ToOrCC, size:Ci.nsMsgSearchAttrib.Size, tags: Ci.nsMsgSearchAttrib.Keywords };
    ["subject", "from", "recipient", "size", "tags"].forEach( function(filter) {
      if ( typeof(rule[filter]) != 'undefined' && rule[filter] != '' ) {
        // if subject in format ^/.*/[ismxpgc]*$ and have customTerm expressionsearch#subjectRegex or filtaquilla@mesquilla.com#subjectRegex
        let customId, positive = true, attribute = rule[filter];
        if ( attribute[0] == '!' ) {
          positive = false;
          attribute = attribute.substr(1);
        }
        if ( attribute.match(/^\/.*\/[ismxpgc]*$/) ) {
          self.advancedTerms[filter].some( function(term) { // .find need TB >=25
            if ( MailServices.filters.getCustomTerm(term) ) {
              customId = term;
              return true;
            } else return false;
          } );
          if ( !customId ) autoArchiveLog.log("Can't support regular expression search patterns '" + rule[filter] + "', 'FiltaQuilla' support RE search for subject, and 'Expression Search / GMailUI' support from/recipient/subject.", 1);
        }
        if ( customId ) autoArchiveUtil.addSearchTerm(searchSession, {type: Ci.nsMsgSearchAttrib.Custom, customId: customId}, attribute, positive ? Ci.nsMsgSearchOp.Matches : Ci.nsMsgSearchOp.DoesntMatch);
        else {
          if ( filter == 'subject' ) autoArchiveUtil.addSearchTerm(searchSession, normal[filter], attribute, positive ? Ci.nsMsgSearchOp.Contains : Ci.nsMsgSearchOp.DoesntContain);
          else if ( filter == 'size' ) {
            let value = autoArchiveUtil.sizeToKB(attribute);
            if ( value != -1 ) autoArchiveUtil.addSearchTerm(searchSession, normal[filter], value, positive ? Ci.nsMsgSearchOp.IsGreaterThan : Ci.nsMsgSearchOp.IsLessThan);
            else autoArchiveLog.log("Can't parse size " + attribute + " , ignore!", 1);
          } else { // from / recipient /tags normal patterns support multiple patterns like '!foo@bar.com, !bar@foo.com' or 'to do, work, !important'
            attribute.split(/[,;]+/).forEach( function(attr) {
              // first remove the leading & trailing blanks
              attr = attr.trim();
              positive = true;
              if ( attr[0] == '!' ) {
                positive = false;
                attr = attr.substr(1);
              }
              if ( filter == 'tags' ) attr = autoArchiveUtil.getKeyFromTag(attr);
              if ( attr != '' ) autoArchiveUtil.addSearchTerm(searchSession, normal[filter], attr, positive ? Ci.nsMsgSearchOp.Contains : Ci.nsMsgSearchOp.DoesntContain);
            } );
          }
        }
      }
    } );
    
    autoArchiveUtil.addSearchTerm(searchSession, Ci.nsMsgSearchAttrib.MsgStatus, Ci.nsMsgMessageFlags.IMAPDeleted, Ci.nsMsgSearchOp.Isnt);
    searchSession.registerListener(new self.searchListener(rule, srcFolder, destFolder));
    this._searchSession = searchSession;
    this.checkServers(); // when check done, call updateFolders if OK, or reset searcSession and call this function again;
    //this.updateFolders(); // when updateFolders done, will call this function again, but have this._searchSession
    //searchSession.search(null);
  },
  advancedTerms : { subject: ['expressionsearch#subjectRegex', 'filtaquilla@mesquilla.com#subjectRegex'], from: ['expressionsearch#fromRegex'], recipient: ['expressionsearch#toRegex'], size: [], tags: [] },
  addStatusListener: function(listener) {
    this.statusListeners.push(listener);
    listener.apply(null, self._status);
  },
  removeStatusListener: function(listener) {
    let index = this.statusListeners.indexOf(listener);
    if ( index >= 0 ) this.statusListeners.splice(index, 1);
  },
  updateStatus: function(status, detail) {
    if ( detail ) autoArchiveLog.info(detail);
    self._status = arguments;
    this.statusListeners.forEach( function(listener) {
      listener.apply(null, self._status);
    } );
  },
  folderIsOf: function(folder, flag) {
    do {
      if ( folder.getFlag(flag) ) return true;
    } while ( ( folder = folder.parent ) && folder && folder != folder.rootFolder );
    return false;
  },
  dryRunLogItems: [],
  dryRunLog: function(log) {
    this.dryRunLogItems.push(log);
    autoArchiveLog.info("Dry run: " + log.join(', '), false, true);
  },

};

// avoid error like 'Wrong rule because folder does not exist' if doArchive called very early
MailUtils.discoverFolders(); // https://bugzilla.mozilla.org/show_bug.cgi?id=502900
let self = autoArchiveService;
self.preStart(autoArchivePref.options.startup_delay);
