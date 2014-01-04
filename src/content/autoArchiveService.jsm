// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveService"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource:///modules/MailUtils.js");
Cu.import("resource:///modules/virtualFolderWrapper.js");
Cu.import("resource://gre/modules/XPCOMUtils.jsm");
Cu.import("resource:///modules/iteratorUtils.jsm"); // import toXPCOMArray
//Cu.import("resource://app/modules/activity/autosync.js");
//Cu.import("resource://app/modules/gloda/utils.js");
//Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("chrome://awsomeAutoArchive/content/aop.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");

let autoArchiveService = {
  timer: null,
  statusListeners: [],
  STATUS_INIT: 0, // not used actually
  STATUS_SLEEP: 1,
  STATUS_WAITIDLE: 2,
  STATUS_RUN: 3,
  STATUS_FINISH: 4,
  _status: [],
  start: function(time) {
    let date = new Date(Date.now() + time*1000);
    this.updateStatus(this.STATUS_SLEEP, "Will wakeup @ " + date.toLocaleDateString() + " " + date.toLocaleTimeString());
    if ( !this.timer ) this.timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
    this.timer.initWithCallback( function() {
      if ( autoArchiveLog && self ) self.waitTillIdle();
    }, time*1000, Ci.nsITimer.TYPE_ONE_SHOT );
  },
  idleService: Cc["@mozilla.org/widget/idleservice;1"].getService(Ci.nsIIdleService),
  idleObserver: {
    delay: null,
    observe: function(_idleService, topic, data) {
      // topic: idle, active
      if ( topic == 'idle' ) self.doArchive();
    }
  },
  waitTillIdle: function() {
    this.updateStatus(this.STATUS_WAITIDLE, "Wait for idle " + autoArchivePref.options.idle_delay + " seconds");
    this.idleObserver.delay = autoArchivePref.options.idle_delay;
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
    else autoArchiveLog.info("Can't remvoe FolderListener", 1);
  },
  clear: function() {
    this.cleanupIdleObserver();
    if ( this.timer ) this.timer.cancel();
    this.folderListeners.forEach( function(listener) {
      self.removeFolderListener(listener);
    } );
    this.hookedFunctions.forEach( function(hooked) {
      hooked.unweave();
    } );
    this.closeAllFoldersDB();
    this.hookedFunctions = [];
    this.rules = [];
    this.copyGroups = [];
    this.status = [];
    this.wait4Folders = {};
    this._searchSession = null;
    this.folderListeners = [];
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
  stop: function() {
    this.clear();
    this.start(autoArchivePref.options.start_next_delay);
  },
  doArchive: function() {
    autoArchiveLog.info("autoArchiveService doArchive");
    this.clear();
    this.rules = autoArchivePref.rules.filter( function(rule) {
      return rule.enable;
    } );
    this.updateStatus(this.STATUS_RUN, "Total " + this.rules.length + " rule(s)");
    autoArchiveLog.logObject(this.rules, 'this.rules',1);
    this.doMoveOrArchiveOne();
  },
  folderListeners: [], // may contain dynamic ones
  folderListener: {
    OnItemEvent: function(folder, event) {
      if ( event.toString() != "FolderLoaded" || !folder || !folder.URI ) return;
      autoArchiveLog.info("FolderLoaded " + folder.URI);
      if ( self.wait4Folders[folder.URI] ) delete self.wait4Folders[folder.URI];
      if ( Object.keys(self.wait4Folders).length == 0 ) {
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
        if ( mail3PaneWindow && mail3PaneWindow.gFolderDisplay && mail3PaneWindow.gFolderDisplay.view && mail3PaneWindow.gFolderDisplay.view.dbView ) mail3PaneWindow.gFolderDisplay.hintMassMoveStarting();
      } catch(err) { autoArchiveLog.logException(err); }
    };
    this.OnProgress = function(aProgress, aProgressMax) {
      autoArchiveLog.info("OnProgress " + aProgress + "/"+ aProgressMax);
    };
    this.OnStopCopy = function(aStatus) {
      autoArchiveLog.info("OnStop " + group.action);
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
    this.msgsDeleted = function(aMsgList) { // Ci.nsIMsgFolderListener, for realDelete message, thus can't get onStopCopy/msgsMoveCopyCompleted
      autoArchiveLog.info("msgsDeleted");
      self.wait4Folders[group.src] = true;
      autoArchiveLog.logObject(aMsgList,'aMsgList',1);
      for (let iMsgHdr = 0; iMsgHdr < aMsgList.length; iMsgHdr++) {
        let msgHdr = aMsgList.queryElementAt(iMsgHdr, Ci.nsIMsgDBHdr);
        let index = group.messages.indexOf(msgHdr);
        if ( index >= 0 ) group.messages.splice(index, 1);
      }
      if ( group.messages.length == 0 ) {
        autoArchiveLog.info("All msgsDeleted");
        MailServices.mfn.removeListener(this);
        if ( self.copyGroups.length ) self.doCopyDeleteMoveOne(self.copyGroups.shift());
        else self.updateFolders();
      }
    };
  },
  
  // updateFolders may get called before when we run search ( when _searchSession was set )
  // or get called after we doing one group of Move/Delete/Copy, or one Archive ( when _searchSession was null )
  // any case we will chain doMoveOrArchiveOne here or in folderListener, and let it to decide either start process a new rule, or continue to search
  updateFolders: function() {
    let folders = [];
    Object.keys(self.wait4Folders).forEach( function(uri) {
      let folder;
      try {
        folder = MailUtils.getFolderForURI(uri);
      } catch(err) { autoArchiveLog.logException(err); }
      if ( folder ) folders.push(folder);
      else delete self.wait4Folders[uri];
    } );
    if ( folders.length ) {
      MailServices.mailSession.AddFolderListener(self.folderListener, Ci.nsIFolderListener.event);
      self.folderListeners.push(self.folderListener);
      let failCount = 0;
      folders.forEach( function(folder) {
        try {
          autoArchiveLog.info("updateFolder " + folder.URI);
          folder.updateFolder(null);
        } catch(err) {
          autoArchiveLog.logException(err);
          failCount ++;
          delete self.wait4Folders[folder.URI];
        }
      } );
      if ( folders.length == failCount ) { // updateFolder fail
        autoArchiveLog.info("updateFolder fail all");
        self.removeFolderListener(self.folderListener);
        self.doMoveOrArchiveOne();
      }
    } else {
      autoArchiveLog.info("no folder to update");
      self.doMoveOrArchiveOne();
    }
  },
  copyGroups: [], // [ {src: src, dest: dest, action: move, messages[]}, ...]
  hookedFunctions: [],
  searchListener: function(rule, srcFolder, destFolder) {
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    if (!mail3PaneWindow) return self.doMoveOrArchiveOne();
    this.QueryInterface = XPCOMUtils.generateQI([Ci.nsISupports, Ci.nsIFolderListener, Ci.nsIMsgSearchNotify]);
    this.messages = [];
    this.missingFolders = {};
    this.sequenceCreateFolders = [];
    this.messagesDest = {};
    let allTags = {};
    let searchHit = 0;
    let duplicateHit = 0;
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
      if ( !msgHdr.messageId || !msgHdr.folder || !msgHdr.folder.URI || msgHdr.folder.URI == rule.dest ) return;
      if ( ['delete', 'move'].indexOf(rule.action) >= 0 && !msgHdr.folder.canDeleteMessages ) return;
      if ( msgHdr.flags & (Ci.nsMsgMessageFlags.Expunged|Ci.nsMsgMessageFlags.IMAPDeleted) ) return;
      let age = ( Date().now / 1000 - msgHdr.dateInSeconds ) / 3600 / 24;
      if ( ["move", "delete", "archive"].indexOf(rule.action) >= 0 && 
        ( ( msgHdr.isFlagged && ( !autoArchivePref.options.enable_flag || age < autoArchivePref.options.age_flag ) ) ||
          ( !msgHdr.isRead && ( !autoArchivePref.options.enable_unread || age < autoArchivePref.options.age_unread ) ) ||
          ( this.hasTag(msgHdr) && ( !autoArchivePref.options.enable_tag || age < autoArchivePref.options.age_tag ) ) ) ) return;
      if ( rule.action == 'archive' && ( msgHdr.folder.getFlag(Ci.nsMsgFolderFlags.Archive) || !mail3PaneWindow.getIdentityForHeader(msgHdr).archiveEnabled ) ) return;
      // check if dest folder has alreay has the message
      if ( ["copy", "move"].indexOf(rule.action) >= 0 ) {
        let realDest = rule.dest, additonal = '';
        let supportHierarchy = ( rule.sub == 2 ) && !srcFolder.getFlag(Ci.nsMsgFolderFlags.Virtual) && !destFolder.getFlag(Ci.nsMsgFolderFlags.Virtual) && destFolder.canCreateSubfolders;
        if ( supportHierarchy && (destFolder.server instanceof Ci.nsIImapIncomingServer)) supportHierarchy = !destFolder.server.isGMailServer;
        if ( supportHierarchy ) {
          let pos = msgHdr.folder.URI.indexOf(rule.src);
          if ( pos != 0 ) {
            autoArchiveLog.info("Message:" + msgHdr.mime2DecodedSubject + " not from src folder?");
            return;
          }
          additonal = msgHdr.folder.URI.substr(rule.src.length);
          realDest = rule.dest + additonal;
        }
        //autoArchiveLog.info(msgHdr.mime2DecodedSubject + ":" + msgHdr.folder.URI + " => " + realDest);
        let realDestFolder = MailUtils.getFolderForURI(realDest);
        // BatchMessageMover using createStorageIfMissing/createSubfolder
        // CopyFolders using createSubfolder
        // https://github.com/gark87/SmartFilters/blob/master/src/chrome/content/backend/imapfolders.jsm using createSubfolder
        // https://github.com/mozilla/releases-comm-central/blob/master/mailnews/imap/test/unit/test_localToImapFilter.js using CopyFolders, but it's empty folders
        // http://thunderbirddocs.blogspot.com/2005/12/mozilla-thunderbird-creating-folders.html
        // http://mxr.mozilla.org/comm-central/source/mailnews/imap/src/nsImapMailFolder.cpp
        // http://mxr.mozilla.org/comm-central/source/mailnews/local/src/nsLocalMailFolder.cpp
        // If target folder already exists but not subscribed, sometimes createStorageIfMissing will not trigger OnStopRunningUrl
        if ( !realDestFolder.parent ) {
          //autoArchiveLog.info("dest folder " + realDest + " not exists, need create");
          this.missingFolders[additonal] = true;
        } else {
          try {
            // msgDatabase is a getter that will always try and load the message database! so null it if not use if anymore
            let destHdr = realDestFolder.msgDatabase.getMsgHdrForMessageID(msgHdr.messageId);
            self.accessedFolders[realDest] = 1;
            if ( destHdr ) {
              //autoArchiveLog.info("Message:" + msgHdr.mime2DecodedSubject + " already exists in dest folder");
              duplicateHit ++;
              return;
            }
          } catch(err) { autoArchiveLog.logException(err); }
        }
        this.messagesDest[msgHdr.messageId] = realDest;
      }
      //autoArchiveLog.info("add message:" + msgHdr.mime2DecodedSubject);
      this.messages.push(msgHdr);
    };
    this.onSearchDone = function(status) {
      try {
        self._searchSession = null;
        autoArchiveLog.info("Total " + searchHit + " messages hit");
        if ( duplicateHit ) autoArchiveLog.info(duplicateHit + " messages already exists in target folder");
        if ( !this.messages.length ) return self.doMoveOrArchiveOne();
        autoArchiveLog.info("will " + rule.action + " " + this.messages.length + " messages");
        // create missing folders first
        if ( Object.keys(this.missingFolders).length ) { // for copy/move
          // rule.dest: imap://a@b.com/1/2
          // additonal:                   /3/4/5
          //                              /3/4/6
          //                              /7/8
          // => 
          // /3, /3/4, /3/4/5, /3/4/6, /7, /7/8
          let needCreateFolders = {};
          let isAsync = destFolder.server.protocolInfo.foldersCreatedAsync;
          Object.keys(this.missingFolders).forEach( function(path) {
            while ( path.length > 0 && path != "/" ) {
              let checkPath = rule.dest + path;
              if ( !needCreateFolders[checkPath] ) {
                let checkFolder = MailUtils.getFolderForURI(checkPath);
                if ( !checkFolder.parent ) needCreateFolders[checkPath] = true;
              }
              let index = path.lastIndexOf('/');
              path = path.substr(0, index);
            }
          } );
          
          this.sequenceCreateFolders = Object.keys(needCreateFolders).sort();
          autoArchiveLog.logObject(this.sequenceCreateFolders, 'this.sequenceCreateFolders', 1);
          if ( !isAsync ) {
            autoArchiveLog.info("create folders sync");
            this.sequenceCreateFolders.forEach( function(path) {
              let [, parent, child] = path.match(/(.*)\/([^\/]+)$/);
              let parentFolder = MailUtils.getFolderForURI(parent);
              parentFolder.createSubfolder(child, null);
            } );
            this.sequenceCreateFolders = [];
          } else {
            autoArchiveLog.info("create folders async");
            MailServices.mailSession.AddFolderListener(this, Ci.nsIFolderListener.added);
            self.folderListeners.push(this);
            return this.OnItemAdded(); // OnItemAdded will chain to create next folder
          }
          this.missingFolders = {};
        }
        this.processHeaders();
      } catch(err) {
        autoArchiveLog.logException(err);
        return self.doMoveOrArchiveOne();
      }
    };
    this.onNewSearch = function() {};
    this.OnItemAdded = function(parentFolder, childFolder) {
      if ( childFolder && childFolder.URI ) {
        autoArchiveLog.info("Folder " + childFolder.URI + " created");
        let index = this.sequenceCreateFolders.indexOf(childFolder.URI);
        if ( index >= 0 ) this.sequenceCreateFolders.splice(index, 1);
      }
      if ( this.sequenceCreateFolders.length ) {
        autoArchiveLog.info("Creating folder " + this.sequenceCreateFolders[0]);
        let destFolder = MailUtils.getFolderForURI(this.sequenceCreateFolders[0]);
        if ( destFolder.parent ) {
          autoArchiveLog.info("Folder " + destFolder.URI + " already exists");
          this.sequenceCreateFolders.splice(0, 1);
          return this.OnItemAdded();
        }
        let [, parent, child] = this.sequenceCreateFolders[0].match(/(.*)\/([^\/]+)$/);
        let parentFolder = MailUtils.getFolderForURI(parent);
        parentFolder.createSubfolder(child, null);
      } else {
        autoArchiveLog.info("All folders created");
        self.removeFolderListener(this);
        return this.processHeaders();
      }
    };
    this.processHeaders = function() {
      try {
        if ( rule.action != 'archive' ) {
          // group messages according to there src and dest
          self.copyGroups = [];
          let groups = {}; // { src => dest : 0, src2 => dest2: 1 }
          this.messages.forEach( function(msgHdr) {
            let dest = listener.messagesDest[msgHdr.messageId] || rule.dest || '';
            if ( dest.length ) self.accessedFolders[dest] = true;
            let key = msgHdr.folder.URI + ( ["copy", "move"].indexOf(rule.action) >= 0 ? " => " + dest : '' );
            if ( typeof(groups[key]) == 'undefined'  ) {
              groups[key] = self.copyGroups.length;
              self.copyGroups.push({src: msgHdr.folder.URI, dest: dest, action: rule.action, messages: []});
            }
            self.copyGroups[groups[key]].messages.push(msgHdr);
          } );
          autoArchiveLog.info("will do " + rule.action + " in " + self.copyGroups.length + " steps");
          autoArchiveLog.logObject(groups, 'groups', 0);
          self.doCopyDeleteMoveOne(self.copyGroups.shift());
        } else {
          // from mailWindowOverlay.js
          self.wait4Folders[rule.src] = true;
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
            for ( let key in batchMover._batches ) {
              autoArchiveLog.info("key " + key);
              autoArchiveLog.logObject(batchMover._batches[key], 'batchMover._batches[key] out', 0);
            }
            try {
              if ( !batchMover.awsome_auto_archive_getFolders ) {
                for ( let key in batchMover._batches ) {
                  autoArchiveLog.logObject(batchMover._batches[key], 'batchMover._batches[key]', 0);
                  let [srcFolder, archiveFolderUri, granularity, keepFolderStructure, msgYear, msgMonth] = batchMover._batches[key];
                  let archiveFolder = MailUtils.getFolderForURI(archiveFolderUri, false);
                  let forceSingle = !archiveFolder.canCreateSubfolders;
                  if (!forceSingle && (archiveFolder.server instanceof Ci.nsIImapIncomingServer)) forceSingle = archiveFolder.server.isGMailServer;
                  if (forceSingle) granularity = Ci.nsIMsgIncomingServer.singleArchiveFolder;
                  if (granularity >= Ci.nsIMsgIdentity.perYearArchiveFolders) archiveFolderUri += "/" + msgYear;
                  if (granularity >= Ci.nsIMsgIdentity.perMonthArchiveFolders) archiveFolderUri += "/" + msgMonth;
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
                    archiveFolderUri += "/" + folderNames.join('/');
                  }
                  autoArchiveLog.info("add update folders " + srcFolder.URI + " => " + archiveFolderUri);
                  self.wait4Folders[srcFolder.URI] = self.wait4Folders[archiveFolderUri] = self.accessedFolders[archiveFolderUri] = true;
                }
                batchMover.awsome_auto_archive_getFolders = true;
              }
            } catch(err) { autoArchiveLog.logException(err); }
            let result = invocation.proceed();
            myFunc(result);
            autoArchiveLog.info("BatchMessageMover processNextBatch exit");
            return result;
          } )[0] );
          autoArchiveLog.info("Start doing archive");
          batchMover.archiveMessages(this.messages);
        }
      } catch(err) {
        autoArchiveLog.logException(err);
        return self.doMoveOrArchiveOne();
      }
    };
  },
  doCopyDeleteMoveOne: function(group) {
    let xpcomHdrArray = toXPCOMArray(group.messages, Ci.nsIMutableArray);
    let srcFolder = MailUtils.getFolderForURI(group.src);
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    let msgWindow;
    if ( mail3PaneWindow ) msgWindow = mail3PaneWindow.msgWindow;
    if ( group.action == 'delete' ) {
      // deleteMessages impacted by srcFolder.server.getIntValue('delete_model')
      // 0:mark as deleted, 1:move to trash, 2:remove it immediately
      let deleteModel = srcFolder.server.getIntValue('delete_model');
      autoArchiveLog.info('deleteModel ' + deleteModel);
      let isTrashFolder = srcFolder.getFlag(Ci.nsMsgFolderFlags.Trash);
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
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    MailServices.copy.CopyMessages(srcFolder, xpcomHdrArray, MailUtils.getFolderForURI(group.dest), isMove, new self.copyListener(group), /*msgWindow*/msgWindow, /* allow undo */false);
  },
  _searchSession: null,
  doMoveOrArchiveOne: function() {
    if ( this._searchSession ) { // updateFolder done, continue to search now
      this._searchSession.search(null);
      return this._searchSession = null;
    }
    //[{"src": "xx", "dest": "yy", "action": "move", "age": 180, "sub": 1, "subject": /test/i, "enable": true}]
    if ( this.rules.length == 0 ) {
      this.closeAllFoldersDB();
      //if ( this.timer ) this.timer.cancel(); // no need to call cancel, start will init another one.
      autoArchiveLog.info("auto archive done for all rules, set next");
      this.updateStatus(this.STATUS_FINISH, '');
      return this.start(autoArchivePref.options.start_next_delay);
    }
    let rule = this.rules.shift();
    //autoArchiveLog.logObject(rule, 'running rule', 1);
    this.updateStatus(this.STATUS_RUN, "Running rule " + rule.action + " " + rule.src + ( ["move", "copy"].indexOf(rule.action)>=0 ? " to " + rule.dest : "" ) +
      " with filter { " + "age: " + ( typeof(rule.age) != 'undefined' ? rule.age : '' ) + " subject: " + ( typeof(rule.subject) ? rule.subject : '' ) + " }");
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
        self.wait4Folders[rule.dest] = self.accessedFolders[rule.dest] = true;
      } else rule.dest = '';
    } catch (err) {
      autoArchiveLog.logException(err);
    }
    if ( !srcFolder || ( ["move", "copy"].indexOf(rule.action) >= 0 && !destFolder ) ) {
      autoArchiveLog.log("Error: Wrong rule becase folder does not exist: " + rule.src + ( ["move", "copy"].indexOf(rule.action) >= 0 ? ' or ' + rule.dest : '' ), 1);
      return this.doMoveOrArchiveOne();
    }
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
        autoArchiveLog.info("Add src folder " + folder.URI);
        searchSession.addScopeTerm(scope, folder);
        self.wait4Folders[folder.URI] = self.accessedFolders[folder.URI] = true;
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
      self.wait4Folders[rule.src] = self.accessedFolders[rule.src] = true;
      if ( rule.sub ) {
        for (let folder in fixIterator(srcFolder.descendants /* >=TB21 */, Ci.nsIMsgFolder)) {
          // We don't add special sub directories, same as AutoarchiveReloaded
          if ( folder.getFlag(Ci.nsMsgFolderFlags.Virtual) ) continue;
          if ( ["move", "archive", "copy"].indexOf(rule.action) >= 0 &&
            folder.getFlag(Ci.nsMsgFolderFlags.Trash | Ci.nsMsgFolderFlags.Junk| Ci.nsMsgFolderFlags.Queue | Ci.nsMsgFolderFlags.Drafts | Ci.nsMsgFolderFlags.Templates ) ) continue;
          if ( rule.action == 'archive' && folder.getFlag(Ci.nsMsgFolderFlags.Archive) ) continue;
          searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, folder);
          self.wait4Folders[folder.URI] = self.accessedFolders[folder.URI] = true;
        }
      }
    }
    
    if ( rule.age ) self.addSearchTerm(searchSession, Ci.nsMsgSearchAttrib.AgeInDays, rule.age, Ci.nsMsgSearchOp.IsGreaterThan);

    let advanced = { subject: ['expressionsearch#subjectRegex', 'filtaquilla@mesquilla.com#subjectRegex'], from: ['expressionsearch#fromRegex'] };
    let normal = { subject: Ci.nsMsgSearchAttrib.Subject, from: Ci.nsMsgSearchAttrib.Sender };
    ["subject", "from"].forEach( function(filter) {
      if ( typeof(rule[filter]) != 'undefined' && rule[filter] != '' ) {
        // if subject in format ^/.*/[ismxpgc]*$ and have customTerm expressionsearch#subjectRegex or filtaquilla@mesquilla.com#subjectRegex
        let customId, positive = true, attribute = rule[filter];
        if ( attribute[0] == '!' ) {
          positive = false;
          attribute = attribute.substr(1);
        }
        if ( attribute.match(/^\/.*\/[ismxpgc]*$/) ) {
          // expressionsearch has logic to deal with Ci.nsMsgMessageFlags.HasRe, use it first
          advanced[filter].some( function(term) { // .find need TB >=25
            if ( MailServices.filters.getCustomTerm(term) ) {
              customId = term;
              return true;
            } else return false;
          } );
          if ( !customId ) autoArchiveLog.log("Can't support regular expression search patterns '" + rule[filter] + "' unless you installed addons like 'Expression Search / GMailUI' or 'FiltaQuilla'", 1);
        }
        if ( customId ) self.addSearchTerm(searchSession, {type: Ci.nsMsgSearchAttrib.Custom, customId: customId}, attribute, positive ? Ci.nsMsgSearchOp.Matches : Ci.nsMsgSearchOp.DoesntMatch);
        else self.addSearchTerm(searchSession, normal[filter], attribute, positive ? Ci.nsMsgSearchOp.Contains : Ci.nsMsgSearchOp.DoesntContain);
      }
    } );
    
    self.addSearchTerm(searchSession, Ci.nsMsgSearchAttrib.MsgStatus, Ci.nsMsgMessageFlags.IMAPDeleted, Ci.nsMsgSearchOp.Isnt);
    searchSession.registerListener(new self.searchListener(rule, srcFolder, destFolder));
    this._searchSession = searchSession;
    this.updateFolders(); // when updateFolders done, will call this function again, but have this._searchSession
    //searchSession.search(null);
  },

  addSearchTerm: function(searchSession, attr, str, op) { // simple version of the one in expression search
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
  },

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
    self._status = [status, detail];
    this.statusListeners.forEach( function(listener) {
      listener.apply(null, self._status);
    } );
  },

};
let self = autoArchiveService;
self.start(autoArchivePref.options.startup_delay);
