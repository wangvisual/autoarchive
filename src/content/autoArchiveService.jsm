// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveService"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource:///modules/MailUtils.js");
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
  clear: function() {
    this.cleanupIdleObserver();
    if ( this.timer ) this.timer.cancel();
    if ( Object.keys(self.wait4Folders).length ) MailServices.mailSession.RemoveFolderListener(self.folderListener);
    this.hookedFunctions.forEach( function(hooked) {
      hooked.unweave();
    } );
    this.hookedFunctions = [];
    this.rules = [];
    this.copyGroups = [];
    this.status = [];
    this.wait4Folders = {};
  },
  rules: [],
  wait4Folders: {},
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
  folderListener: {
    OnItemEvent: function(folder, event) {
      if ( event.toString() != "FolderLoaded" || !folder || !folder.URI ) return;
      autoArchiveLog.info("FolderLoaded " + folder.URI);
      if ( self.wait4Folders[folder.URI] ) delete self.wait4Folders[folder.URI];
      if ( Object.keys(self.wait4Folders).length == 0 ) {
        MailServices.mailSession.RemoveFolderListener(self.folderListener);
        autoArchiveLog.info("All FolderLoaded");
        self.doMoveOrArchiveOne();
      }
    },
  },
  copyListener: function(group) { // this listener is for Copy/Delete/Move actions
    this.QueryInterface = function(iid) {
      if ( !iid.equals(Ci.nsIMsgCopyServiceListener) && !iid.equals(Ci.nsIMsgFolderListener) && !iid.equals(Ci.nsISupports) )
        throw Components.results.NS_ERROR_NO_INTERFACE;
      return this;
    };
    this.OnStartCopy = function() {
      autoArchiveLog.info("OnStart " + group.action);
      let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
      if ( mail3PaneWindow && mail3PaneWindow.gFolderDisplay ) mail3PaneWindow.gFolderDisplay.hintMassMoveStarting();
    };
    this.OnProgress = function(aProgress, aProgressMax) {
      autoArchiveLog.info("OnProgress " + aProgress + "/"+ aProgressMax);
    };
    this.OnStopCopy = function(aStatus) {
      autoArchiveLog.info("OnStop " + group.action);
      if ( group.action == 'delete' || group.action == 'move' ) self.wait4Folders[group.src] = true;
      if ( group.action == 'copy' || group.action == 'move' ) self.wait4Folders[group.dest] = true;
      let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
      if ( mail3PaneWindow && mail3PaneWindow.gFolderDisplay ) mail3PaneWindow.gFolderDisplay.hintMassMoveCompleted();
      if ( self.copyGroups.length ) self.doCopyDeleteMoveOne(self.copyGroups.shift());
      else this.updateFolders();
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
        else this.updateFolders();
      }
    };
    this.updateFolders = function() {
      let folders = [];
      Object.keys(self.wait4Folders).forEach( function(uri) {
        let folder = MailUtils.getFolderForURI(uri);
        if ( folder ) folders.push(folder);
        else delete self.wait4Folders[uri];
      } );
      if ( folders.length ) {
        MailServices.mailSession.AddFolderListener(self.folderListener, Ci.nsIFolderListener.event);
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
          MailServices.mailSession.RemoveFolderListener(self.folderListener);
          self.doMoveOrArchiveOne();
        }
      } else {
        autoArchiveLog.info("no folder to update");
        self.doMoveOrArchiveOne();
      }
    };
  },
  copyGroups: [], // [ {src: src, dest: dest, action: move, messages[]}, ...]
  hookedFunctions: [],
  searchListener: function(rule, srcFolder, destFolder) {
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    if (!mail3PaneWindow) return self.doMoveOrArchiveOne();
    this.messages = [];
    this.createFolders = [];
    let messagesDest = {};
    let allTags = {};
    let searchHit = 0;
    let foldersGotDB = {}; // in onSearchHit, we check if the message exists in dest folder and open it's DB, need to null them later
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
        let realDest = rule.dest;
        let supportHierarchy = ( rule.sub == 2 ) && !srcFolder.getFlag(Ci.nsMsgFolderFlags.Virtual) && !destFolder.getFlag(Ci.nsMsgFolderFlags.Virtual) && destFolder.canCreateSubfolders;
        if ( supportHierarchy && (destFolder.server instanceof Ci.nsIImapIncomingServer)) supportHierarchy = !destFolder.server.isGMailServer;
        if ( supportHierarchy ) {
          let pos = msgHdr.folder.URI.indexOf(rule.src);
          if ( pos != 0 ) {
            autoArchiveLog.info("Message:" + msgHdr.mime2DecodedSubject + " not from src folder?");
            return;
          }
          realDest = rule.dest + msgHdr.folder.URI.substr(rule.src.length);
        }
        let realDestFolder = MailUtils.getFolderForURI(realDest);
        // BatchMessageMover using createStorageIfMissing/createSubfolder
        // CopyFolders using createSubfolder
        // https://github.com/gark87/SmartFilters/blob/master/src/chrome/content/backend/imapfolders.jsm using createSubfolder
        // https://github.com/mozilla/releases-comm-central/blob/master/mailnews/imap/test/unit/test_localToImapFilter.js using CopyFolders, but it's empty folders
        if ( !realDestFolder ) {
          //autoArchiveLog.info("dest folder " + realDest + "not exists, need create");
          let isAsync = destFolder.server.protocolInfo.foldersCreatedAsync;
          //return;
        }
        // msgDatabase is a getter that will always try and load the message database! so null it if not use if anymore
        try {
          let destHdr = realDestFolder.msgDatabase.getMsgHdrForMessageID(msgHdr.messageId);
          foldersGotDB[realDestFolder] = 1;
          if ( destHdr ) {
            autoArchiveLog.info("Message:" + msgHdr.mime2DecodedSubject + " already exists in dest folder");
            return;
          }
        } catch(err) { autoArchiveLog.logException(err); }
        messagesDest[msgHdr] = realDest;
      }
      //autoArchiveLog.info("add message:" + msgHdr.mime2DecodedSubject);
      this.messages.push(msgHdr);
    };
    this.onSearchDone = function(status) {
      try {
        // TODO: close opended DB
        if ( !this.messages.length ) return self.doMoveOrArchiveOne();
        autoArchiveLog.info("Total " + searchHit + " messages hit");
        autoArchiveLog.info("will " + rule.action + " " + this.messages.length + " messages");
        if ( rule.action != 'archive' ) {
          // group messages according to there src and dest
          self.copyGroups = [];
          let groups = {}; // { src-_|_-dest : 0, src2-_|_-dest2: 1 }
          this.messages.forEach( function(msgHdr) {
            let dest = messagesDest[msgHdr] || rule.dest|| '';
            let key = msgHdr.folder.URI + ( ["copy", "move"].indexOf(rule.action) >= 0 ? "-_|_-" + dest : '' );
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
          let batchMover = new mail3PaneWindow.BatchMessageMover();
          let myFunc = function(result) {
            autoArchiveLog.info("BatchMessageMover OnStopCopy/OnStopRunningUrl");
            autoArchiveLog.logObject(batchMover._batches,'batchMover._batches',1);
            if ( batchMover._batches == null || Object.keys(batchMover._batches).length == 0 ) {
              autoArchiveLog.info("BatchMessageMover Done");
              self.hookedFunctions.forEach( function(hooked) {
                hooked.unweave();
              } );
              self.hookedFunctions = [];
              // TODO: updateFolder
              self.doMoveOrArchiveOne();
            }
            return result;
          }
          self.hookedFunctions.push( autoArchiveaop.after( {target: batchMover, method: 'OnStopCopy'}, myFunc )[0] );
          self.hookedFunctions.push( autoArchiveaop.after( {target: batchMover, method: 'OnStopRunningUrl'}, myFunc )[0] );
          batchMover.archiveMessages(this.messages);
        }
      } catch(err) {
        autoArchiveLog.logException(err);
        return self.doMoveOrArchiveOne();
      }
    };
    this.onNewSearch = function() {};
    
  /*const nsMsgFolderFlags = Components.interfaces.nsMsgFolderFlags;
  if (!MailServices.mailSession.IsFolderOpenInWindow(folder) &&
      !(folder.flags & (nsMsgFolderFlags.Trash | nsMsgFolderFlags.Inbox)))
  {
    folder.msgDatabase = null;
  }*/
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
      srcFolder.msgDatabase = null; /* don't leak */
      return;
    }
    let isMove = (group.action == 'move') && srcFolder.canDeleteMessages;
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    MailServices.copy.CopyMessages(srcFolder, xpcomHdrArray, MailUtils.getFolderForURI(group.dest), isMove, new self.copyListener(group), /*msgWindow*/msgWindow, /* allow undo */false);
  },
  doMoveOrArchiveOne: function() {
    //[{"src": "xx", "dest": "yy", "action": "move", "age": 180, "sub": 1, "subject": /test/i, "enable": true}]
    if ( this.rules.length == 0 ) {
      autoArchiveLog.info("auto archive done for all rules, set next");
      this.updateStatus(this.STATUS_FINISH, '');
      return this.start(autoArchivePref.options.start_next_delay);
    }
    let rule = this.rules.shift();
    //autoArchiveLog.logObject(rule, 'running rule', 1);
    this.updateStatus(this.STATUS_RUN, "Running rule " + rule.action + " " + rule.src + ( ["move", "copy"].indexOf(rule.action)>=0 ? " to " + rule.dest : "" ) +
      " with filter { " + "age: " + rule.age + " subject: " + rule.subject + " }");
    let srcFolder = null, destFolder = null;
    try {
      srcFolder = MailUtils.getFolderForURI(rule.src);
      if ( ["move", "copy"].indexOf(rule.action) >= 0 ) destFolder = MailUtils.getFolderForURI(rule.dest);
    } catch (err) {
      autoArchiveLog.logException(err);
    }
    if ( !srcFolder || ( ["move", "copy"].indexOf(rule.action) >= 0 && !destFolder ) ) {
      autoArchiveLog.log("Error: Wrong rule becase folder does not exist: " + rule.src + ( ["move", "copy"].indexOf(rule.action) >= 0 ? ' or ' + rule.dest : '' ), 1);
      return this.doMoveOrArchiveOne();
    }
   
    //let array = toXPCOMArray([srcFolder], Ci.nsIMutableArray);
    //MailServices.copy.CopyFolders(array, destFolder, false, null, null);
    //return;
    
    let searchSession = Cc["@mozilla.org/messenger/searchSession;1"].createInstance(Ci.nsIMsgSearchSession);
    searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, srcFolder);
    if ( rule.sub ) {
      for (let folder in fixIterator(srcFolder.descendants /* >=TB21 */, Ci.nsIMsgFolder)) {
        // We don't add special sub directories, same as AutoarchiveReloaded
        if ( folder.getFlag(Ci.nsMsgFolderFlags.Virtual) ) continue;
        if ( ["move", "archive", "copy"].indexOf(rule.action) >= 0 &&
          folder.getFlag(Ci.nsMsgFolderFlags.Trash | Ci.nsMsgFolderFlags.Junk| Ci.nsMsgFolderFlags.Queue | Ci.nsMsgFolderFlags.Drafts | Ci.nsMsgFolderFlags.Templates ) ) continue;
        if ( rule.action == 'archive' && folder.getFlag(Ci.nsMsgFolderFlags.Archive) ) continue;
        searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, folder);
      }
    }
    self.addSearchTerm(searchSession, Ci.nsMsgSearchAttrib.AgeInDays, rule.age || 0, Ci.nsMsgSearchOp.IsGreaterThan);

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
    searchSession.search(null);
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
