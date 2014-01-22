// Opera Wang, 2013/5/1
// GPL V3 / MPL
"use strict";
var EXPORTED_SYMBOLS = ["autoArchivePref"];
const { classes: Cc, Constructor: CC, interfaces: Ci, utils: Cu, results: Cr, manager: Cm } = Components;
Cu.import("resource://gre/modules/Services.jsm");
Cu.import("chrome://awsomeAutoArchive/content/log.jsm");
const mozIJSSubScriptLoader = Cc["@mozilla.org/moz/jssubscript-loader;1"].getService(Ci.mozIJSSubScriptLoader);

let autoArchivePref = {
  path: null,
  // https://bugzilla.mozilla.org/show_bug.cgi?id=469673
  // https://groups.google.com/forum/#!topic/mozilla.dev.extensions/SBGIogdIiwE
  InstantApply: false,
  setInstantApply: function(instant) {
    this.InstantApply = instant;
    autoArchiveLog.info("autoArchivePref: set InstantApply to " + instant);
  },
  // TODO: When bug 564675 is implemented this will no longer be needed
  // Always set the default prefs, because they disappear on restart
  setDefaultPrefs: function () {
    let branch = Services.prefs.getDefaultBranch("");
    let prefLoaderScope = {
      pref: function(key, val) {
        switch (typeof val) {
          case "boolean":
            branch.setBoolPref(key, val);
            break;
          case "number":
            branch.setIntPref(key, val);
            break;
          case "string":
            branch.setCharPref(key, val);
            break;
        }
      }
    };
    let uri = Services.io.newURI("defaults/preferences/prefs.js", null, Services.io.newURI(this.path, null, null));
    try {
      mozIJSSubScriptLoader.loadSubScript(uri.spec, prefLoaderScope);
    } catch (err) {
      Cu.reportError(err);
    }
  },
  options: {},
  cleanup: function() {
    this.prefs.removeObserver("", this, false);
  },
  initPerf: function(spec) {
    this.path = spec.replace(/bootstrap\.js$/, '');
    this.setDefaultPrefs();
    this.prefs = Services.prefs.getBranch(this.prefPath);
    this.prefs.addObserver("", this, false);
    let self = this;
    this.allPrefs.forEach( function(key) {
      self.observe('', 'nsPref:changed', key); // we fake one
    } );
  },
  setPerf: function(key, value) {
    switch (typeof(value)) {
      case 'boolean':
        this.prefs.setBoolPref(key, value);
        break;
      case 'number':
        this.prefs.setIntPref(key, value);
        break;
      default:
        this.prefs.setCharPref(key, value);
    }
  },

  prefPath: "extensions.awsome_auto_archive.",
  allPrefs: ['enable_verbose_info', 'rules', 'enable_flag', 'enable_tag', 'enable_unread', 'age_flag', 'age_tag', 'age_unread', 'startup_delay', 'idle_delay', 'start_next_delay', 'rule_timeout', 'update_statusbartext', 'default_days', 'dry_run', 'messages_number_limit', 'messages_size_limit', 'start_exceed_delay', 'show_folder_as'],
  rules: [],
  observe: function(subject, topic, data) {
    try {
      if (topic != "nsPref:changed") return;
      switch(data) {
        case "enable_verbose_info":
        case "enable_flag":
        case "enable_tag":
        case 'enable_unread':
        case 'update_statusbartext':
        case 'dry_run':
          this.options[data] = this.prefs.getBoolPref(data);
          break;
        case "rules":
          this.options[data] = this.prefs.getCharPref(data);
          break;
        default:
          this.options[data] = this.prefs.getIntPref(data);
          break;
      }
      if ( data == 'enable_verbose_info' ) {
        autoArchiveLog.setVerbose(this.options.enable_verbose_info);
      } else if ( data == 'rules' ) {
        if ( !this.InstantApply ) this.validateRules();
      }
    } catch (err) {
      autoArchiveLog.logException(err);
    };
  },
  validateRules: function(rules) {
    if ( !rules ) rules = JSON.parse(this.options.rules);
    rules.forEach( function(rule) {
      let error = false;
      ["src", "action", "age", "sub", "enable"].forEach( function(att) {
        if ( typeof(rule[att]) == 'undefined' ) {
          autoArchiveLog.log("Error: rule lacks of property " + att, 1);
          error = true;
        }
      } );
      if ( ["move", "archive", "copy", "delete"].indexOf(rule.action) < 0 ) {
        autoArchiveLog.log("Error: rule action must be one of move or archive", 1);
        error = true;
      }
      if ( ["move", "copy"].indexOf(rule.action) >= 0 ) {
        if ( typeof(rule.dest) == 'undefined' ) {
          autoArchiveLog.log("Error: dest folder must be defined for copy/move action", 1);
          error = true;
        } else if ( rule.src == rule.dest ) {
          autoArchiveLog.log("Error: dest folder must be different from src folder", 1);
          error = true;
        }
      }
      if ( error ) rule.enable = false;
      // fix old config that save rule.sub as string, can be delete later
      if ( typeof(rule.sub) == 'string' ) rule.sub = Number(rule.sub);
    } );
    autoArchiveLog.logObject(rules,'rules',1);
    return rules;
  },
};
