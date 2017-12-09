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
  timer: Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer),
  // https://bugzilla.mozilla.org/show_bug.cgi?id=469673
  // https://groups.google.com/forum/#!topic/mozilla.dev.extensions/SBGIogdIiwE
  // https://bugzilla.mozilla.org/show_bug.cgi?id=1415567 Remove {get,set}ComplexValue use of nsISupportsString in Thunderbird
  oldAPI_58: Services.vc.compare(Services.appinfo.platformVersion, '58') < 0,
  InstantApply: false,
  setInstantApply: function(instant) {
    this.InstantApply = instant;
    autoArchiveLog.info("autoArchivePref: set InstantApply to " + instant);
  },
  // bootstrapped add-ons was obsoleted, Mozilla won't support read default any more
  // https://bugzilla.mozilla.org/show_bug.cgi?id=564675 Allow bootstrapped add-ons to have default preferences
  // https://bugzilla.mozilla.org/show_bug.cgi?id=1413413 Remove support for extensions having their own prefs file
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
    let uri = Services.io.newURI("content/defaults_prefs.js", null, Services.io.newURI(this.path, null, null));
    try {
      mozIJSSubScriptLoader.loadSubScript(uri.spec, prefLoaderScope);
    } catch (err) {
      Cu.reportError(err);
    }
  },
  options: {},
  prefListeners: [],
  cleanup: function() {
    this.prefs.removeObserver("", this, false);
    this.timer.cancel();
    delete this.prefListeners;
    delete this.options;
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
    this.options.key = value; // don't wait for observe callback
    switch (typeof(value)) {
      case 'boolean':
        this.prefs.setBoolPref(key, value);
        break;
      case 'number':
        this.prefs.setIntPref(key, value);
        break;
      default:
        if ( key in this.complexPrefs ) {
          if ( this.oldAPI_58) {
            let str = Cc["@mozilla.org/supports-string;1"].createInstance(Ci.nsISupportsString);
            str.data = value;
            this.prefs.setComplexValue(key, this.complexPrefs[key], str);
          } else {
            this.prefs.setStringPref(key, str);
          }
        }
        else this.prefs.setCharPref(key, value);
    }
  },
  addPrefListener: function(listener) {
    this.prefListeners.push(listener);
  },
  removePrefListener: function(listener) {
    let index = this.prefListeners.indexOf(listener);
    if ( index >= 0 ) this.prefListeners.splice(index, 1);
  },

  prefPath: "extensions.awsome_auto_archive.",
  allPrefs: ['enable_verbose_info', 'rules', 'rules_to_keep', 'enable_flag', 'enable_tag', 'enable_unread', 'age_flag', 'age_tag', 'age_unread', 'startup_delay', 'idle_delay', 'check_servers', 'update_folders',
             'start_next_delay', 'rule_timeout', 'generate_rule_use', 'show_from', 'show_recipient', 'show_subject', 'show_size', 'show_tags', 'show_age', 'delete_duplicate_in_src', 'ignore_spam_folders',
             'update_statusbartext', 'default_days', 'dry_run', 'messages_number_limit', 'messages_size_limit', 'start_exceed_delay', 'show_folder_as', 'add_context_munu_rule', 'alert_show_time', 'hibernate', 'archive_archive_folders'],
  complexPrefs: {'rules': Ci.nsISupportsString },
  observe: function(subject, topic, key) {
    try {
      if (topic != "nsPref:changed") return;
      switch(key) {
        case "enable_verbose_info":
        case "enable_flag":
        case "enable_tag":
        case 'enable_unread':
        case 'update_statusbartext':
        case 'dry_run':
        case 'add_context_munu_rule':
        case 'show_from':
        case 'show_recipient':
        case 'show_subject':
        case 'show_size':
        case 'show_tags':
        case 'show_age':
        case 'delete_duplicate_in_src':
        case 'ignore_spam_folders':
        case 'archive_archive_folders':
        case 'update_folders':
        case 'check_servers':
          this.options[key] = this.prefs.getBoolPref(key);
          break;
        default:
          if ( key in this.complexPrefs ) this.options[key] = this.oldAPI_58 ? this.prefs.getComplexValue(key, this.complexPrefs[key]).data : this.prefs.getStringPref(key);
          else this.options[key] = this.prefs.getIntPref(key);
          break;
      }
      if ( key == 'enable_verbose_info' ) {
        autoArchiveLog.setVerbose(this.options.enable_verbose_info);
      } else if ( key == 'alert_show_time' ) {
        autoArchiveLog.setPopupDelay(this.options.alert_show_time);
      } else if ( key == 'rules' ) {
        if ( !this.InstantApply ) this.validateRules();
      }
      this.prefListeners.forEach( function(listener) {
        listener.call(null, key);
      } );
    } catch (err) {
      autoArchiveLog.logException(err);
    };
  },
  realValidate: function(rules) {
    if ( !autoArchiveLog || !rules ) return;
    let count = 1;
    rules.forEach( rule => {
      let error = false;
      ["src", "action", "age", "sub", "enable"].forEach( function(att) {
        if ( typeof(rule[att]) == 'undefined' ) {
          autoArchiveLog.log("Error: rule " + count + " lacks of property " + att, 1);
          error = true;
        }
      } );
      if ( ["move", "archive", "copy", "delete"].indexOf(rule.action) < 0 ) {
        autoArchiveLog.log("Error: rule " + count + " action must be one of move or archive", 1);
        error = true;
      }
      if ( ["move", "copy"].indexOf(rule.action) >= 0 ) {
        if ( typeof(rule.dest) == 'undefined' ) {
          autoArchiveLog.log("Error: rule " + count + " dest folder must be defined for copy/move action", 1);
          error = true;
        } else if ( rule.src == rule.dest ) {
          autoArchiveLog.log("Error: rule " + count + " dest folder must be different from src folder", 1);
          error = true;
        }
      }
      if ( error ) rule.enable = false;
      // fix old config that save rule.sub as string, can be delete later
      if ( typeof(rule.sub) == 'string' ) rule.sub = Number(rule.sub);
      count++;
    } );
    autoArchiveLog.logObject(rules,'rules',1);
    return rules;
  },
  validateRules: function(rules) {
    let aSync = false;
    if ( !rules ) {
      rules = JSON.parse(this.options.rules);
      aSync = true;
    }
    if ( aSync ) {
      return this.timer.initWithCallback( () => {
        if ( autoArchivePref ) autoArchivePref.realValidate(rules);
      }, 0, Ci.nsITimer.TYPE_ONE_SHOT );
    } else return this.realValidate(rules);
  },
};
