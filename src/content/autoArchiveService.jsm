// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveService"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
//Cu.import("resource:///modules/mailServices.js");
//Cu.import("resource://app/modules/gloda/utils.js");
//Cu.import("resource://gre/modules/FileUtils.jsm");
Cu.import("chrome://autoArchive/content/log.jsm");

let autoArchiveService = {
  timer: null,
  start: function(time) {
    if ( !this.timer ) this.timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
    this.timer.initWithCallback( function() {
      if ( autoArchiveLog && autoArchiveService ) {
        autoArchiveLog.info('autoArchiveService Timer');
        if ( 1 ) autoArchiveService.waitTillIdle();
        else autoArchiveService.doArchive();
      }
    }, time*1000, Ci.nsITimer.TYPE_ONE_SHOT );
  },
  idleService: Cc["@mozilla.org/widget/idleservice;1"].getService(Ci.nsIIdleService),
  idleObserver: {
    delay: null,
    observe: function(_idleService, topic, data) {
      // topic: idle, active
      autoArchiveLog.info("topic: " + topic + "\ndata: " + data);
      if ( topic == 'idle' ) {
        autoArchiveService.cleanupIdleObserver();
        autoArchiveService.doArchive();
      }
    }
  },
  waitTillIdle: function() {
    this.idleObserver.delay = 1/*60*/;
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
    if ( this.timer ) {
      this.timer.cancel();
      this.timer = null;
    }
    this.cleanupIdleObserver();
    autoArchiveLog.info("autoArchiveService cleanup done");
  },
  
  doArchive: function() {
    autoArchiveLog.info("autoArchiveService doArchive");
    // get actions
    //autoArchiveService.start(86400/360);
  },
  searchListener: function(action, destFolder) {
    this.messages = [];
    //this.action = action;
    this.msgHdrsArchive = function () {
      let mail3PaneWindow = Services.wm.getMostRecentWindow("mail:3pane");
      if (!mail3PaneWindow) return;
      let batchMover = new mail3PaneWindow.BatchMessageMover();
      batchMover.archiveMessages(this.messages.filter( function (msgHdr) {
        // !msgHdr.folder.getFlag(0x00004000) && mail3PaneWindow.getIdentityForHeader(msgHdr).archiveEnabled && !star && !tag;
      } ));
    };
    this.onSearchHit = function (dbHdr, folder) {
      this.messages.push(dbHdr);
    };
    this.onSearchDone = function (status) {
      this.msgHdrsArchive();
    };
    this.onNewSearch = function () {};
  },
  doArchiveOne: function(action, folder, subFolder, age, destFolder) {
    let searchSession = Cc["@mozilla.org/messenger/searchSession;1"].createInstance(Ci.nsIMsgSearchSession);
    searchSession.addScopeTerm(Ci.nsMsgSearchScope.offlineMail, folder);
    let searchByAge = searchSession.createTerm();
    searchByAge.attrib = Ci.nsMsgSearchAttrib.AgeInDays;
    let value = searchByAge.value;
    value.attrib = Ci.nsMsgSearchAttrib.AgeInDays;
    value.age = age;
    searchByAge.value = value;
    searchByAge.op = Ci.nsMsgSearchOp.IsGreaterThan;
    searchByAge.booleanAnd = true;
    searchSession.appendTerm(searchByAge);
    searchSession.registerListener(new autoArchiveService.searchListener(action, destFolder));
    searchSession.search(null); 
  },
};
autoArchiveService.start(1);
