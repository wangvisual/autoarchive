// MPL/GPL
// Opera.Wang 2014/03/01
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveActivity"];

Cu.import("chrome://awsomeAutoArchive/content/autoArchiveService.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveUtil.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");

// SeaMonkey has no activity-manager
let gActivityManager = Cc["@mozilla.org/activity-manager;1"] ? Cc["@mozilla.org/activity-manager;1"].getService(Ci.nsIActivityManager) : null;
let autoArchiveActivity = {
  cleanup: function() {
    autoArchiveLog.info('autoArchiveActivity cleanup');
    if ( gActivityManager ) autoArchiveService.removeStatusListener(this.statusCallback);
    self.process = null;
    autoArchiveLog.info('autoArchiveActivity cleanup done');
  },
  process: null,
  statusCallback: function(status, detail, index, total) {
    try {
      if ( status == autoArchiveService.STATUS_RUN ) {
        if ( !self.process ) {
          self.process = Cc["@mozilla.org/activity-process;1"].createInstance(Ci.nsIActivityProcess);
          self.process.init("Total " + total + " rules", null);
          self.process.contextDisplayText = autoArchiveUtil.Name + " " + autoArchiveUtil.Version;
          gActivityManager.addActivity(self.process);
        }
        // if running the 2nd of 3 rules, that means only 1st finished.
        self.process.setProgress(detail + " (" + ( index + 1 ) + " of " + total + " rules)", index, total);
      } else if ( status == autoArchiveService.STATUS_FINISH && self.process ) {
        self.process.state = Ci.nsIActivityProcess.STATE_COMPLETED;
        gActivityManager.removeActivity(self.process.id);
        let event = Cc["@mozilla.org/activity-event;1"].createInstance(Ci.nsIActivityEvent);
        event.init(detail, null, autoArchiveUtil.Name + " " + autoArchiveUtil.Version, self.process.startTime, Date.now());
        gActivityManager.addActivity(event);
        self.process = null;
      }
    } catch(err) { autoArchiveLog.logException(err); }
  },
}
let self = autoArchiveActivity;
if ( gActivityManager ) autoArchiveService.addStatusListener(self.statusCallback);
