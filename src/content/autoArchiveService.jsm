// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveService"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("resource:///modules/mailServices.js");
Cu.import("resource:///modules/MailUtils.js");
Cu.import("resource:///modules/iteratorUtils.jsm"); // import toXPCOMArray
//Cu.import("resource://app/modules/gloda/utils.js");
//Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("chrome://awsomeAutoArchive/content/aop.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchivePref.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");

let autoArchiveService = {
  timer: null,
  start: function(time) {
    autoArchiveLog.info('start: delay ' + time);
    if ( !this.timer ) this.timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
    this.timer.initWithCallback( function() {
      if ( autoArchiveLog && self ) {
        autoArchiveLog.info('autoArchiveService Timer reached');
        if ( 1 ) self.waitTillIdle();
        else self.doArchive();
      }
    }, time*1000, Ci.nsITimer.TYPE_ONE_SHOT );
  },
  idleService: Cc["@mozilla.org/widget/idleservice;1"].getService(Ci.nsIIdleService),
  idleObserver: {
    delay: null,
    observe: function(_idleService, topic, data) {
      // topic: idle, active
      if ( topic == 'idle' ) {
        self.cleanupIdleObserver();
        self.doArchive();
      }
    }
  },
  waitTillIdle: function() {
    autoArchiveLog.info("autoArchiveService waitTillIdle");
    this.idleObserver.delay = 2/*60*/;
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
    this.rules = [];
    this.copyGroups = [];
    if ( this.timer ) {
      this.timer.cancel();
      this.timer = null;
    }
    this.cleanupIdleObserver();
    autoArchiveLog.info("autoArchiveService cleanup done");
  },
  rules: [],
  doArchive: function() {
    autoArchiveLog.info("autoArchiveService doArchive");
    this.rules = autoArchivePref.rules.filter( function(rule) {
      return rule.enable;
    } );
    autoArchiveLog.logObject(this.rules, 'this.rules',1);
    this.doMoveOrArchiveOne();
  },
  copyListener: function(group) { // this listener is for Copy/Delete/Move actions
    this.OnStartCopy = function() {
      autoArchiveLog.info("OnStart " + group.action);
      //gFolderDisplay.hintMassMoveStarting();
    };
    this.OnProgress = function(aProgress, aProgressMax) {
      autoArchiveLog.info("OnProgress " + aProgress + "/"+ aProgressMax);
    };
    this.OnStopCopy = function(aStatus) {
      autoArchiveLog.info("OnStop " + group.action);
      //gFolderDisplay.hintMassMoveCompleted();
      if ( self.copyGroups.length ) self.doCopyDeleteMoveOne(self.copyGroups.shift());
      else self.doMoveOrArchiveOne();
    };
    this.SetMessageKey = function(aKey) {};
    this.GetMessageId = function() {};
  },
  copyGroups: [], // [ {src: src, dest: dest, action: move, messages[]}, ...]
  searchListener: function(rule) {
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    if (!mail3PaneWindow) return self.doMoveOrArchiveOne();
    this.messages = [];
    this.msgHdrsArchive = function() {
      try {
        if ( !this.messages.length ) return self.doMoveOrArchiveOne();
        autoArchiveLog.info("will " + rule.action + " " + this.messages.length + " messages");
        if ( rule.action != 'archive' ) {
          // group messages according to there src and dest
          self.copyGroups = [];
          let groups = {}; // { src-_|_-dest : 0, src2-_|_-dest2: 1 }
          this.messages.forEach( function(msgHdr) {
            let key = msgHdr.folder.folderURL + "-_|_-" + (rule.dest||'');
            if ( typeof(groups.key) == 'undefined'  ) {
              groups.key = self.copyGroups.length;
              self.copyGroups.push({src: msgHdr.folder.folderURL, dest: rule.dest, action: rule.action, messages: []});
            }
            self.copyGroups[groups.key].messages.push(msgHdr);
          } );
          autoArchiveLog.logObject(self.copyGroups, 'copyGroups', 0);
          self.doCopyDeleteMoveOne(self.copyGroups.shift());
        } else {
          let batchMover = new mail3PaneWindow.BatchMessageMover();
          autoArchiveaop.after( {target: batchMover, method: 'OnStopCopy'}, function(result) {
            if ( batchMover._batches == null ) self.doMoveOrArchiveOne();
            return result;
          } );
          batchMover.archiveMessages(this.messages);
        }
      } catch(err) {
        autoArchiveLog.logException(err);
        return self.doMoveOrArchiveOne();
      }
    };
    this.onSearchHit = function(msgHdr, folder) {
      //autoArchiveLog.info("search hit message:" + msgHdr.mime2DecodedSubject);
      if ( !msgHdr.folder || msgHdr.folder.folderURL == rule.dest ) return;
      if ( ( !autoArchivePref.options.enable_flag && msgHdr.isFlagged ) || ( !autoArchivePref.options.enable_tag && msgHdr.getStringProperty('keywords').contains('$label') ) ) return;
      //let str = ''; let e = msgHdr.propertyEnumerator; let str = "property:\n"; while ( e.hasMore() ) { let k = e.getNext(); str += k + ":" + msgHdr.getStringProperty(k) + "\n"; }; autoArchiveLog.info(str);
      if ( rule.action == 'archive' && ( msgHdr.folder.getFlag(Ci.nsMsgFolderFlags.Archive) || !mail3PaneWindow.getIdentityForHeader(msgHdr).archiveEnabled ) ) {
        autoArchiveLog.info('return:' + msgHdr.subject + "   :" + msgHdr.folder.customIdentity);
        autoArchiveLog.logObject(msgHdr.folder,'msgHdr.folder',0);
        return;
      }
      if ( ['delete', 'move'].indexOf(rule.action) >= 0 && !msgHdr.folder.canDeleteMessages ) return;
      if ( rule.action == 'delete' && ( msgHdr.flags & (Ci.nsMsgMessageFlags.Expunged|Ci.nsMsgMessageFlags.IMAPDeleted) ) ) return;
      //autoArchiveLog.info("add message:" + msgHdr.mime2DecodedSubject);
      this.messages.push(msgHdr);
    };
    this.onSearchDone = function(status) {
      this.msgHdrsArchive();
    };
    this.onNewSearch = function() {};
  },
  doCopyDeleteMoveOne: function(group) {
    let xpcomHdrArray = toXPCOMArray(group.messages, Ci.nsIMutableArray);
    let srcFolder = MailUtils.getFolderForURI(group.src, true);
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    let msgWindow;
    if ( mail3PaneWindow ) msgWindow = mail3PaneWindow.msgWindow;
    if ( group.action == 'delete' ) {
      // deleteMessages impacted by srcFolder.server.getIntValue('delete_model')
      // 0:mark as deleted, 1:move to trash, 2:remove it immediately
      let deleteModel = srcFolder.server.getIntValue('delete_model');
      let isTrash = srcFolder.getFlag(Ci.nsMsgFolderFlags.Trash);
      autoArchiveLog.info('srcFolder:' + srcFolder.URI + " is trash? " + isTrash);
      try {
        srcFolder.deleteMessages(xpcomHdrArray, null, /*deleteStorage*/(deleteModel == 2 || srcFolder.getFlag(Ci.nsMsgFolderFlags.Trash)), /*isMove*/false, new self.copyListener(group), /* allow undo */false);
        //srcFolder.updateFolder(msgWindow);
        srcFolder.msgDatabase = null; /* don't leak */
        return;
      } catch (err) {
        autoArchiveLog.logException(err);
      }
    }
    let isMove = (group.action == 'move') && srcFolder.canDeleteMessages;
    let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
    MailServices.copy.CopyMessages(srcFolder, xpcomHdrArray, MailUtils.getFolderForURI(group.dest, true), isMove, new self.copyListener(group), /*msgWindow*/msgWindow, /* allow undo */false);  
  },
  doMoveOrArchiveOne: function() {
    //[{"src": "xx", "dest": "yy", "action": "move", "age": 180, "sub": 1, "subject": /test/, "enable": true}]
    if ( this.rules.length == 0 ) {
      autoArchiveLog.info("auto archive done for all rules, set next");
      return this.start(300);
    }
    let rule = this.rules.shift();
    autoArchiveLog.logObject(rule, 'running rule', 1);
    let srcFolder = null, destFolder = null;
    try {
      srcFolder = MailUtils.getFolderForURI(rule.src, true);
      if ( ["move", "copy"].indexOf(rule.action) >= 0 ) destFolder = MailUtils.getFolderForURI(rule.dest, true);
    } catch (err) {
      autoArchiveLog.logException(err);
    }
    if ( !srcFolder || ( ["move", "copy"].indexOf(rule.action) >= 0 && !destFolder ) ) {
      autoArchiveLog.log("Error: Wrong rule becase folder does not exist: " + rule.src + ( ["move", "copy"].indexOf(rule.action) >= 0 ? ' or ' + rule.dest : '' ), 1);
      return this.doMoveOrArchiveOne();
    }
    let searchSession = Cc["@mozilla.org/messenger/searchSession;1"].createInstance(Ci.nsIMsgSearchSession);
    searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, srcFolder);
    if ( rule.sub ) {
      for (let folder in fixIterator(srcFolder.descendants /* >=TB21 */, Ci.nsIMsgFolder)) {
        searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, folder);
      }
    }
    let searchByAge = searchSession.createTerm();
    searchByAge.attrib = Ci.nsMsgSearchAttrib.AgeInDays;
    let value = searchByAge.value;
    value.attrib = searchByAge.attrib;
    value.age = rule.age;
    searchByAge.value = value;
    searchByAge.op = Ci.nsMsgSearchOp.IsGreaterThan;
    searchByAge.booleanAnd = true;
    searchSession.appendTerm(searchByAge);
    if ( rule.subject ) {
      //MailServices.filters.getCustomTerm('expressionsearch#subjectRegex');
      //attr = { type:nsMsgSearchAttrib.Custom, customId: 'expressionsearch#subjectRegex' };
      //nsMsgSearchOp.Contains;
      /*let term = aTermCreator.createTerm();
      term.attrib = Ci.nsMsgSearchAttrib.Custom;
      let value = term.value;
      value.attrib = term.attrib;
      term.value = value;
      term.op = Ci.nsMsgSearchOp.Contains;
      term.booleanAnd = true;
      term.customId = aCustomId;*/
    }
    searchSession.registerListener(new self.searchListener(rule));
    searchSession.search(null);
  },
};
let self = autoArchiveService;
self.start(2);
