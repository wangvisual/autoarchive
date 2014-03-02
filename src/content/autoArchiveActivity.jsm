// MPL/GPL
// Opera.Wang 2014/03/01
"use strict";
var EXPORTED_SYMBOLS = ["autoArchiveActivity"];

const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveService.jsm");
Cu.import("chrome://awsomeAutoArchive/content/autoArchiveUtil.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");

let gActivityManager = Cc["@mozilla.org/activity-manager;1"].getService(Ci.nsIActivityManager);
let autoArchiveActivity = {
  cleanup: function() {
    autoArchiveService.removeStatusListener(this.statusCallback);
    self.process = null;
  },
  process: null,
  statusCallback: function(status, detail, index, total) {
    if ( status == autoArchiveService.STATUS_RUN && total > 1 ) {
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
  },
}
let self = autoArchiveActivity;
autoArchiveService.addStatusListener(self.statusCallback);
