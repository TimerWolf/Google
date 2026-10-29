///////////////////////////////////////
// Gmail: labels                     //
// Revision: 17                      //
// Date created: 2025/10/16          //
// Date last modified: 2026/10/29    //
// Author: Whiskey / Google Gemini   //
///////////////////////////////////////

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

// CLOCK_TIMESTAMP - Dynamic time window configuration:
// 1. Relative rolling offset (from current execution time):
//    - ['-2d']                       = 2 days back from current execution time.
//    - ['+12h']                      = 12 hours forward from current execution time.
//    (Units: m = minutes, h = hours, d = days, M = months, y = years)
//
// 2. Absolute date & time range (with optional directional '-' prefix or offsets):
//    - ['2026-10-13']                = From 00:00 on that date FORWARD to current time.
//    - ['-2026-10-13']               = From 00:00 on that date BACKWARD in time.
//    - ['2026-10-13 14:55', '-2h']   = 2 hours BACKWARD starting from 2026-10-13 14:55.
//    - ['2026-10-13 00:00', '14:55']  = Exact start and end time on the same day.
//
// 3. No restriction:
//    - []                            = Searches ALL messages in your account (no time filter).
const CLOCK_TIMESTAMP = [
  '-2d',
];

// Message filtering settings
const EXCLUDE_SENT_FROM_ME = true;  // Set to true to ignore messages sent by yourself (-from:me)
const EXCLUDE_INTERNAL_SENT_MESSAGES = true; // Excludes individual sent messages found inside threads

// Target label restriction settings
const RESTRICT_TO_TARGET_LABELS = true; // Set to true to restrict search ONLY to TARGET_LABELS, false to search all labels

// Strict label validation behavior:
// - true  = Throws an error and halts execution if any target label is missing.
// - false = Logs a warning, skips the missing label, and continues running with valid ones.
const STRICT_LABEL_VALIDATION = false; 

// Define specific labels or system folders to target using Wildcards (*), default is 'in:inbox':
// - 'label'    = Exact match only.
// - 'label/*'  = Only sub-labels under label.
// - 'label*'   = Root label AND all its sub-labels.
// - '*label*'  = Any label containing 'label' anywhere in its path.
const TARGET_LABELS = [
  'in:inbox',
];

// Trigger execution settings
// TRIGGER_MAX_EXECUTION_TIME: execution time limit settings (Must not exceed Google's hard limit of 6 minutes total)
// Allowed values for MINUTES : 0 to 5
// Allowed values for SECONDS : 0 to 59
const MAX_EXECUTION_TIME = true; // Set to true for execution time check
const MAX_EXECUTION_TIME_MINUTES = 5; 
const MAX_EXECUTION_TIME_SECONDS = 0; 

const ENABLE_SETTINGS = true;      // Set to true to log active settings on run
const ENABLE_INFO = true;          // Set to true for summary metrics
const ENABLE_DEBUG = true;         // Set to true for detailed message logs
const ENABLE_WARNING = false;       // Set to true for detailed warning logs

// Trigger frequency settings
// TRIGGER_FREQUENCY_TYPE: 'MINUTES' or 'HOURS'
// Allowed values for MINUTES : 1, 5, 10, 15, 30
// Allowed values for HOURS   : 1, 2, 4, 6, 8, 12

const TRIGGER_FREQUENCY_TYPE = 'MINUTES'; 
const TRIGGER_EVERY_X = 30;                

// ---------------------------------------------------------------------------
// Main function
// ---------------------------------------------------------------------------

function applyHierarchicalLabelsFromGmail() {

  var totalLimitMs = (MAX_EXECUTION_TIME_MINUTES * 60 * 1000) + (MAX_EXECUTION_TIME_SECONDS * 1000);
  var absoluteMaxMs = 6 * 60 * 1000;

  // Validate that total execution time limit does not exceed Google's 6-minute hard limit if enabled
  if (MAX_EXECUTION_TIME && totalLimitMs >= absoluteMaxMs) {
    throw new Error('❌ [Error] Execution time limit cannot exceed 6 minutes (360,000 ms). Google Apps Script has a hard execution limit of 6 minutes.');
  }

  // Step 1: Fetch and map all user labels (Name -> ID & Object) first for validation
  var userLabels = GmailApp.getUserLabels();
  var labelMapByName = {};
  
  userLabels.forEach(function(lbl) {
    labelMapByName[lbl.getName()] = lbl;
  });

  // Resolve Wildcards and Validate TARGET_LABELS
  var activeMatchedLabels = [];
  var missingTargetLabels = [];

  if (RESTRICT_TO_TARGET_LABELS) {
    if (!TARGET_LABELS || TARGET_LABELS.length === 0) {
      throw new Error('❌ [Error] RESTRICT_TO_TARGET_LABELS is set to true, but NO target labels are defined in the TARGET_LABELS array.');
    }

    TARGET_LABELS.forEach(function(targetPattern) {
      
      // System operators (e.g. in:inbox)
      if (targetPattern.indexOf('in:') === 0 || targetPattern.indexOf('label:') === 0) {
        activeMatchedLabels.push(targetPattern);
        return;
      }

      // Wildcard logic
      if (targetPattern.indexOf('*') !== -1) {
        
        var escapeRegex = function(str) {
          return str.replace(/([.+^${}()|[\]\\])/g, '\\$1');
        };
        var regexPattern = new RegExp('^' + targetPattern.split('*').map(escapeRegex).join('.*') + '$', 'i');

        var matchedInAccount = [];
        userLabels.forEach(function(lbl) {
          var name = lbl.getName();
          if (regexPattern.test(name)) {
            matchedInAccount.push(name);
          }
        });

        if (matchedInAccount.length > 0) {
          matchedInAccount.forEach(function(m) {
            activeMatchedLabels.push(m);
          });
        } else {
          missingTargetLabels.push(targetPattern);
          var warningMsg = '⚠️ [Warning] Wildcard pattern "' + targetPattern + '" matched NO labels in your Gmail account.';
          if (STRICT_LABEL_VALIDATION) {
            throw new Error('❌ [Error] ' + warningMsg.replace('⚠️ [Warning] ', ''));
          } else if (ENABLE_WARNING) {
            Logger.log(warningMsg + ' Skipped.');
          }
        }

      } else {
        // Exact match
        if (labelMapByName[targetPattern]) {
          activeMatchedLabels.push(targetPattern);
        } else {
          missingTargetLabels.push(targetPattern);
          var warningMsgExact = '⚠️ [Warning] Target label "' + targetPattern + '" defined in TARGET_LABELS does not exist in your Gmail account.';
          if (STRICT_LABEL_VALIDATION) {
            throw new Error('❌ [Error] ' + warningMsgExact.replace('⚠️ [Warning] ', ''));
          } else if (ENABLE_WARNING) {
            Logger.log(warningMsgExact + ' Skipped.');
          }
        }
      }
    });

    // Deduplicate matched labels
    activeMatchedLabels = activeMatchedLabels.filter(function(item, pos) {
      return activeMatchedLabels.indexOf(item) === pos;
    });

    if (activeMatchedLabels.length === 0) {
      throw new Error('❌ [Error] None of the specified TARGET_LABELS or Wildcards matched any labels in your account.');
    }
  }

  // Parse CLOCK_TIMESTAMP dynamically (Supports negative date prefixes & relative offsets)
  var startTimeSec = null;
  var endTimeSec = null;
  var currentTimeMs = new Date().getTime();

  if (CLOCK_TIMESTAMP && CLOCK_TIMESTAMP.length > 0) {
    var rawFirstVal = CLOCK_TIMESTAMP[0].trim();
    var isNegativeDate = false;

    // Check for negative date prefix e.g. '-2026-10-13'
    if (rawFirstVal.indexOf('-') === 0 && rawFirstVal.match(/^-\d{4}-\d{2}-\d{2}/)) {
      isNegativeDate = true;
      rawFirstVal = rawFirstVal.substring(1).trim();
    }

    // Check if relative offset format (e.g. '-2d', '+4h', '-30m')
    var relativeMatch = rawFirstVal.match(/^([+-])(\d+)([mhdMy])$/);
    if (relativeMatch) {
      var sign = relativeMatch[1];
      var amount = parseInt(relativeMatch[2], 10);
      var unit = relativeMatch[3];

      var multiplier = 1000;
      if (unit === 'm') multiplier *= 60;
      else if (unit === 'h') multiplier *= 3600;
      else if (unit === 'd') multiplier *= 86400;
      else if (unit === 'M') multiplier *= 86400 * 30; // Approx month
      else if (unit === 'y') multiplier *= 86400 * 365; // Approx year

      var offsetMs = amount * multiplier;

      if (sign === '-') {
        startTimeSec = Math.floor((currentTimeMs - offsetMs) / 1000);
        endTimeSec = Math.floor(currentTimeMs / 1000);
      } else {
        startTimeSec = Math.floor(currentTimeMs / 1000);
        endTimeSec = Math.floor((currentTimeMs + offsetMs) / 1000);
      }
    } else {
      // Absolute timestamp parsing
      var baseTimeMs = new Date(rawFirstVal).getTime();

      if (CLOCK_TIMESTAMP.length > 1 && CLOCK_TIMESTAMP[1]) {
        var secondVal = CLOCK_TIMESTAMP[1].trim();
        var secondRelative = secondVal.match(/^([+-])(\d+)([mhdMy])$/);

        if (secondRelative) {
          var secSign = secondRelative[1];
          var secAmount = parseInt(secondRelative[2], 10);
          var secUnit = secondRelative[3];

          var secMult = 1000;
          if (secUnit === 'm') secMult *= 60;
          else if (secUnit === 'h') secMult *= 3600;
          else if (secUnit === 'd') secMult *= 86400;
          else if (secUnit === 'M') secMult *= 86400 * 30;
          else if (secUnit === 'y') secMult *= 86400 * 365;

          var secOffsetMs = secAmount * secMult;

          if (secSign === '-') {
            startTimeSec = Math.floor((baseTimeMs - secOffsetMs) / 1000);
            endTimeSec = Math.floor(baseTimeMs / 1000);
          } else {
            startTimeSec = Math.floor(baseTimeMs / 1000);
            endTimeSec = Math.floor((baseTimeMs + secOffsetMs) / 1000);
          }
        } else if (secondVal.indexOf(':') === 2 || (secondVal.indexOf(':') === 1 && secondVal.length <= 5)) {
          var baseDateStr = rawFirstVal.split(' ')[0];
          startTimeSec = Math.floor(baseTimeMs / 1000);
          endTimeSec = Math.floor(new Date(baseDateStr + ' ' + secondVal).getTime() / 1000);
        } else {
          startTimeSec = Math.floor(baseTimeMs / 1000);
          endTimeSec = Math.floor(new Date(secondVal).getTime() / 1000);
        }
      } else {
        if (isNegativeDate) {
          startTimeSec = null; 
          endTimeSec = Math.floor(baseTimeMs / 1000);
        } else {
          startTimeSec = Math.floor(baseTimeMs / 1000);
          endTimeSec = Math.floor(currentTimeMs / 1000);
        }
      }
    }
  }

  // Calculate formatted duration text for logging (Plural 's' applied ONLY if val > 1)
  var timeWindowText = 'All Messages (No Window)';
  if (startTimeSec !== null || endTimeSec !== null) {
    var effStart = startTimeSec !== null ? startTimeSec : 0;
    var effEnd = endTimeSec !== null ? endTimeSec : Math.floor(currentTimeMs / 1000);
    var diffSec = Math.abs(effEnd - effStart);

    var calcYears = Math.floor(diffSec / (86400 * 365));
    diffSec %= (86400 * 365);

    var calcMonths = Math.floor(diffSec / (86400 * 30));
    diffSec %= (86400 * 30);

    var calcDays = Math.floor(diffSec / 86400);
    diffSec %= 86400;

    var calcHours = Math.floor(diffSec / 3600);
    diffSec %= 3600;

    var calcMinutes = Math.floor(diffSec / 60);
    var calcSeconds = diffSec % 60;

    var formatUnit = function(val, singular, plural) {
      return val + ' ' + (val > 1 ? plural : singular);
    };

    var parts = [
      formatUnit(calcYears, 'year', 'years'),
      formatUnit(calcMonths, 'month', 'months'),
      formatUnit(calcDays, 'day', 'days'),
      formatUnit(calcHours, 'hour', 'hours'),
      formatUnit(calcMinutes, 'minute', 'minutes'),
      formatUnit(calcSeconds, 'second', 'seconds')
    ];

    timeWindowText = 'Custom (' + CLOCK_TIMESTAMP.join(' -> ') + ') | ' + parts.join(', ');
  }

  // Unified start and settings log block
  if (ENABLE_DEBUG || ENABLE_SETTINGS) {
    Logger.log('==================================================');
    Logger.log('⚙️ CONFIGURATION & SETTINGS');
    Logger.log('==================================================');
    
    if (ENABLE_SETTINGS) {
      Logger.log('⏱️ CLOCK_TIMESTAMP                  : ' + (CLOCK_TIMESTAMP.length > 0 ? CLOCK_TIMESTAMP.join(', ') : 'None (Search ALL messages)'));
      Logger.log('📅 Time Window Evaluated            : ' + timeWindowText);
      Logger.log('🚫 EXCLUDE_SENT_FROM_ME             : ' + EXCLUDE_SENT_FROM_ME);
      Logger.log('🚫 EXCLUDE_INTERNAL_SENT_MESSAGES   : ' + EXCLUDE_INTERNAL_SENT_MESSAGES);
      Logger.log('🔍 RESTRICT_TO_TARGET_LABELS        : ' + RESTRICT_TO_TARGET_LABELS);
      Logger.log('🔒 STRICT_LABEL_VALIDATION          : ' + STRICT_LABEL_VALIDATION);
      Logger.log('🎯 TARGET_LABELS (Configured)       : ' + TARGET_LABELS.join(', '));
      Logger.log('🎯 TARGET_LABELS (Active Match)     : ' + activeMatchedLabels.join(', '));
      Logger.log('🎯 TARGET_LABELS (Missing)          : ' + (missingTargetLabels.length > 0 ? missingTargetLabels.join(', ') : 'None'));
      Logger.log('⏱️ MAX_EXECUTION_TIME               : ' + MAX_EXECUTION_TIME);
      Logger.log('⏱️ MAX_EXECUTION_TIME_MINUTES       : ' + MAX_EXECUTION_TIME_MINUTES);
      Logger.log('⏱️ MAX_EXECUTION_TIME_SECONDS       : ' + MAX_EXECUTION_TIME_SECONDS);
      Logger.log('💬 ENABLE_INFO                      : ' + ENABLE_INFO);
      Logger.log('💬 ENABLE_DEBUG                     : ' + ENABLE_DEBUG);
      Logger.log('💬 ENABLE_WARNING                   : ' + ENABLE_WARNING);
      Logger.log('💬 ENABLE_SETTINGS                  : ' + ENABLE_SETTINGS);
      Logger.log('⚙️ TRIGGER_FREQUENCY_TYPE           : ' + TRIGGER_FREQUENCY_TYPE);
      Logger.log('⏱️ TRIGGER_EVERY_X                  : ' + TRIGGER_EVERY_X);
      Logger.log('==================================================');
    }
  }
  
  var startTime = new Date().getTime();

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    
    if (ENABLE_DEBUG) {
      Logger.log('🔒 [Debug] Could not obtain lock. Another instance is running.');
    }
    
    return;
  
  }

  // Counters for statistics
  var stats = {
    totalThreads: 0,
    threadsWithLabels: 0,
    threadsWithoutLabels: 0,
    totalMessages: 0,
    messagesWithLabels: 0,
    messagesWithoutLabels: 0,
    hierarchicalLabelsFound: 0,
    labelsAddedCount: 0,
    createdGmailLabels: 0
  };

  var timeOutReached = false;
  var validTargetMessages = [];

  try {

    // Step 2: Build exact Unix timestamp search query and label restriction
    var queryParts = [];

    // Label scope restriction based on activeMatchedLabels
    if (RESTRICT_TO_TARGET_LABELS && activeMatchedLabels.length > 0) {
      var labelQueryParts = activeMatchedLabels.map(function(lbl) {
        if (lbl.indexOf(':') !== -1) {
          return lbl;
        }
        return 'label:"' + lbl + '"';
      });

      if (labelQueryParts.length > 1) {
        queryParts.push('(' + labelQueryParts.join(' OR ') + ')');
      } else {
        queryParts.push(labelQueryParts[0]);
      }
    } else {
      queryParts.push('has:userlabel'); // Fallback to all labeled mail if restriction is false
    }

    // Time window restriction based on CLOCK_TIMESTAMP
    if (startTimeSec !== null) {
      queryParts.push('after:' + startTimeSec);
    }
    if (endTimeSec !== null) {
      queryParts.push('before:' + endTimeSec);
    }

    // Exclude sent messages restriction
    if (EXCLUDE_SENT_FROM_ME) {
      queryParts.push('-from:me');
    }

    var query = queryParts.join(' ');

    var threads = GmailApp.search(query);
    
    stats.totalThreads = threads.length;

    if (ENABLE_DEBUG) {
      Logger.log('📦 [Debug] Total cached Gmail labels in account: ' + userLabels.length);
      Logger.log('⏰ [Debug] Time window configuration (Start Sec): ' + (startTimeSec !== null ? startTimeSec : 'None'));
      Logger.log('⏰ [Debug] Time window configuration (End Sec):   ' + (endTimeSec !== null ? endTimeSec : 'None'));
      Logger.log('🔍 [Debug] Restrict to target labels: ' + RESTRICT_TO_TARGET_LABELS);
      Logger.log('🔒 [Debug] Strict label validation: ' + STRICT_LABEL_VALIDATION);
      Logger.log('📤 [Debug] Exclude sent messages (-from:me): ' + EXCLUDE_SENT_FROM_ME);
      Logger.log('🔎 [Debug] Executed search query: "' + query + '"');
      Logger.log('🧵 [Debug] Total threads matched: ' + threads.length);
    }

    for (var threadIndex = 0; threadIndex < threads.length; threadIndex++) {
    
      // Check if execution time limit is approaching (only if MAX_EXECUTION_TIME is true)
      if (MAX_EXECUTION_TIME && (new Date().getTime() - startTime > totalLimitMs)) {
        timeOutReached = true;
        if (ENABLE_WARNING) {
          Logger.log('⚠️ [Warning] Execution time limit (' + MAX_EXECUTION_TIME_MINUTES + 'm ' + MAX_EXECUTION_TIME_SECONDS + 's) reached. Stopping gracefully.');
        }
        break;
      }

      var thread = threads[threadIndex];
      var messages = thread.getMessages();
      var threadHasUserLabels = false;

      for (var msgIndex = 0; msgIndex < messages.length; msgIndex++) {
        var msg = messages[msgIndex];

        var msgTimeMs = msg.getDate().getTime();

        // Double check exact millisecond timestamp if window is active
        if (startTimeSec !== null && msgTimeMs < (startTimeSec * 1000)) {
          continue; 
        }
        if (endTimeSec !== null && msgTimeMs > (endTimeSec * 1000)) {
          continue;
        }

        // Internal sent messages filter
        if (EXCLUDE_INTERNAL_SENT_MESSAGES) {
          var isSentByMe = false;
          try {
            var rawMsg = Gmail.Users.Messages.get('me', msg.getId(), { fields: 'labelIds' });
            var labelIds = rawMsg.labelIds || [];
            userLabels.forEach(function(l) {
              if (labelIds.indexOf(l.getId()) !== -1 && l.getName().toLowerCase() === 'jag') {
                isSentByMe = true;
              }
            });
          } catch(err) {}

          if (isSentByMe) {
            continue;
          }
        }

        validTargetMessages.push({
          msg: msg,
          threadIdx: threadIndex + 1,
          totalThreads: threads.length
        });
    
        stats.totalMessages++;
    
        var msgId = msg.getId();
        
        // Fetch raw message details via Gmail REST API to get EXACT message-level label IDs
        var rawMessage = Gmail.Users.Messages.get('me', msgId, { fields: 'labelIds' });
        var currentMsgLabelIds = rawMessage.labelIds || [];

        // Filter out system labels (INBOX, UNREAD, etc.) to evaluate user labels
        var currentMsgLabelNames = [];
    
        userLabels.forEach(function(l) {
          if (currentMsgLabelIds.indexOf(l.getId()) !== -1) {
            currentMsgLabelNames.push(l.getName());
          }
        });

        if (currentMsgLabelNames.length > 0) {
          stats.messagesWithLabels++;
          threadHasUserLabels = true;
        } else {
          stats.messagesWithoutLabels++;
        }

        var missingLabelIdsToAdd = [];
        var missingLabelNamesAdded = [];

        // Analyze each label present on this specific message
        currentMsgLabelNames.forEach(function(lblName) {
    
          if (lblName.indexOf('/') !== -1) {
    
            stats.hierarchicalLabelsFound++;
    
            var parts = lblName.split('/').map(function(p){ return p.trim(); });
            var currentPath = '';

            parts.forEach(function(part, idx) {
              currentPath = idx === 0 ? part : currentPath + '/' + part;

              // Ensure label exists in Gmail account
              var targetLabel = labelMapByName[currentPath] || GmailApp.getUserLabelByName(currentPath);
    
              if (!targetLabel) {
                targetLabel = GmailApp.createLabel(currentPath);
                labelMapByName[currentPath] = targetLabel;
                userLabels = GmailApp.getUserLabels(); // Refresh local array
                stats.createdGmailLabels++;
    
                if (ENABLE_DEBUG) {
                  Logger.log('✨ [Debug] Created new label in Gmail account: "' + currentPath + '"');
                }
              }

              var targetLabelId = targetLabel.getId();

              // Check if THIS SPECIFIC MESSAGE is missing the parent label ID
              if (currentMsgLabelIds.indexOf(targetLabelId) === -1 && missingLabelIdsToAdd.indexOf(targetLabelId) === -1) {  
                missingLabelIdsToAdd.push(targetLabelId);
                missingLabelNamesAdded.push(currentPath);
              }
            });
          }
        });

        // FORCE ADD missing parent labels directly to the MESSAGE ID via REST API
        if (missingLabelIdsToAdd.length > 0) {

          Gmail.Users.Messages.modify({
            'addLabelIds': missingLabelIdsToAdd
          }, 'me', msgId);

          stats.labelsAddedCount += missingLabelIdsToAdd.length;

          if (ENABLE_DEBUG) {

            Logger.log('\n--- Thread #' + (threadIndex + 1) + ' | Message #' + (msgIndex + 1) + ' (ID: ' + msgId + ') ---');
            Logger.log('✉️ [Debug] Subject: "' + msg.getSubject() + '"');
            Logger.log('🏷️ [Debug] Existing labels on message: [' + currentMsgLabelNames.join(', ') + ']');
            Logger.log('🛠️ [Debug] -> FIXED/ADDED ' + missingLabelIdsToAdd.length + ' LABEL(S): [' + missingLabelNamesAdded.join(', ') + ']');

          }
        }
      }

      if (threadHasUserLabels) {
        stats.threadsWithLabels++;
      } else {
        stats.threadsWithoutLabels++;
      }
    }

    if (ENABLE_INFO || timeOutReached) {

      var executionTimeMs = new Date().getTime() - startTime;
      var executionTimeSec = (executionTimeMs / 1000).toFixed(2);
      var executionTimeFormatted = parts.join(', ');
      var sentStatusText = EXCLUDE_SENT_FROM_ME ? 'false' : 'true';

      Logger.log('==================================================');
      Logger.log('📊 SUMMARY & METRICS');
      Logger.log('==================================================');
      Logger.log('⏱️ [Stats] Execution Time:                  ' + executionTimeSec + ' seconds (' + executionTimeMs + ' ms)');
      Logger.log('📅 [Stats] Time Window Mode:                ' + timeWindowText);
      Logger.log('📤 [Stats] Sent Messages Filter:            ' + sentStatusText);
      Logger.log('🔒 [Stats] Internal Sent Filter:            ' + EXCLUDE_INTERNAL_SENT_MESSAGES);
      Logger.log('--------------------------------------------------');
      Logger.log('📩 [Stats] Total Messages Searched:         ' + stats.totalMessages);
      Logger.log('🏷️ [Stats] Messages WITH user labels:       ' + stats.messagesWithLabels);
      Logger.log('📭 [Stats] Messages WITHOUT user labels:    ' + stats.messagesWithoutLabels);
      Logger.log('--------------------------------------------------');
      Logger.log('🧵 [Stats] Total Threads Searched:          ' + stats.totalThreads);
      Logger.log('🧵 [Stats] Threads WITH user labels:        ' + stats.threadsWithLabels);
      Logger.log('🧵 [Stats] Threads WITHOUT user labels:     ' + stats.threadsWithoutLabels);
      Logger.log('--------------------------------------------------');
      Logger.log('🎯 [Stats] Target Labels Active:            ' + activeMatchedLabels.join(', '));
      Logger.log('📭 [Stats] Target Labels Missing:           ' + (missingTargetLabels.length > 0 ? missingTargetLabels.join(', ') : 'None'));
      Logger.log('🌲 [Stats] Hierarchical labels (with /):    ' + stats.hierarchicalLabelsFound);
      Logger.log('✨ [Stats] New Gmail labels created:        ' + stats.createdGmailLabels);
      Logger.log('🔧 [Stats] Total Parent Labels Applied:     ' + stats.labelsAddedCount);
      Logger.log('==================================================\n');
    
    }

    // Run statistical audit & tree breakdown
    auditLabelConsistency(query, validTargetMessages);

  } catch(e) {
    Logger.log('❌ [Error] Exception: ' + e.message + '\n' + e.stack);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// AUDIT FUNCTION (Dynamic Conflict Detection & Tree Breakdown)
// ---------------------------------------------------------------------------

function auditLabelConsistency(activeQuery, targetMessages) {
  var userLabels = GmailApp.getUserLabels();
  var rootNodes = {};
  var conflictDetailsList = [];
  var flatCollisionsByMessage = [];
  var hierarchicalCollisionsByParent = {}; 
  var totalEvaluated = 0;

  var messagesToScan = targetMessages || [];
  if (messagesToScan.length === 0) {
    var query = activeQuery || 'has:userlabel';
    var threads = GmailApp.search(query);
    threads.forEach(function(thread) {
      thread.getMessages().forEach(function(msg) {
        messagesToScan.push({ msg: msg });
      });
    });
  }

  messagesToScan.forEach(function(item) {
    var msg = item.msg;
    var msgId = msg.getId();
    var rawMessage;
    try {
      rawMessage = Gmail.Users.Messages.get('me', msgId, { fields: 'labelIds' });
    } catch (e) {
      return;
    }
    var labelIds = rawMessage.labelIds || [];
    
    var labelNames = [];
    userLabels.forEach(function(l) {
      if (labelIds.indexOf(l.getId()) !== -1) {
        labelNames.push(l.getName());
      }
    });

    var filteredLabels = labelNames.filter(function(lbl) {
      var upper = lbl.toUpperCase();
      var lower = lbl.toLowerCase();
      if (EXCLUDE_INTERNAL_SENT_MESSAGES && lower === 'jag') {
        return false;
      }
      return upper !== 'INBOX' && 
             lower !== 'in:inbox' && 
             lower.indexOf('in:') !== 0 && 
             lower.indexOf('label:') !== 0;
    });

    totalEvaluated++;

    var tz = Session.getScriptTimeZone();
    var msgDateStr = Utilities.formatDate(msg.getDate(), tz, 'yyyy-MM-dd HH:mm:ss');
    var msgSubject = msg.getSubject() || '(No Subject)';

    // 1. Flat label conflicts
    var flatLabelsOnMsg = filteredLabels.filter(function(l) { return l.indexOf('/') === -1; });

    if (flatLabelsOnMsg.length > 1) {
      flatCollisionsByMessage.push({
        date: msgDateStr,
        subject: msgSubject,
        labels: flatLabelsOnMsg
      });

      var rootTreeLines = [];
      flatLabelsOnMsg.forEach(function(fl) {
        rootTreeLines.push('   │    ├── ❌ ' + fl);
      });

      conflictDetailsList.push({
        date: msgDateStr,
        subject: msgSubject,
        conflictType: 'Multiple flat/Inbox categories on same message (CONFLICT):',
        lines: rootTreeLines
      });
    }

    // 2. Hierarchical leaf conflicts under the same parent
    var leafLabelsPerParent = {};
    filteredLabels.forEach(function(lblName) {
      if (lblName.indexOf('/') !== -1) {
        var parts = lblName.split('/').map(function(p){ return p.trim(); });
        var parentPath = parts.slice(0, parts.length - 1).join(' / ');
        var leaf = parts[parts.length - 1];

        if (!leafLabelsPerParent[parentPath]) {
          leafLabelsPerParent[parentPath] = [];
        }
        if (leafLabelsPerParent[parentPath].indexOf(leaf) === -1) {
          leafLabelsPerParent[parentPath].push(leaf);
        }
      }
    });

    for (var pPath in leafLabelsPerParent) {
      var leaves = leafLabelsPerParent[pPath];
      if (leaves.length > 1) {
        if (!hierarchicalCollisionsByParent[pPath]) {
          hierarchicalCollisionsByParent[pPath] = [];
        }
        hierarchicalCollisionsByParent[pPath].push({
          date: msgDateStr,
          subject: msgSubject,
          leaves: leaves
        });

        var hierTreeLines = [];
        leaves.forEach(function(leaf) {
          hierTreeLines.push('   │    ├── ❌ ' + pPath + ' / ' + leaf);
        });

        conflictDetailsList.push({
          date: msgDateStr,
          subject: msgSubject,
          conflictType: 'Multiple sub-labels under parent [' + pPath + '] on same message (CONFLICT):',
          lines: hierTreeLines
        });
      }
    }

    // Build tree structure
    filteredLabels.forEach(function(lbl) {
      var parts = lbl.split('/').map(function(p){ return p.trim(); });
      
      if (parts.length > 1) {
        var rootName = parts[0];
        if (!rootNodes[rootName]) {
          rootNodes[rootName] = { count: 0, children: {} };
        }
        rootNodes[rootName].count++;

        var currentLevel = rootNodes[rootName].children;
        for (var i = 1; i < parts.length; i++) {
          var segment = parts[i];
          if (!currentLevel[segment]) {
            currentLevel[segment] = { count: 0, children: {} };
          }
          currentLevel[segment].count++;
          currentLevel = currentLevel[segment].children;
        }
      } else {
        var inboxRoot = 'Inbox';
        if (!rootNodes[inboxRoot]) {
          rootNodes[inboxRoot] = { count: 0, children: {} };
        }
        rootNodes[inboxRoot].count++;

        var flatName = parts[0];
        if (!rootNodes[inboxRoot].children[flatName]) {
          rootNodes[inboxRoot].children[flatName] = { count: 0, children: {} };
        }
        rootNodes[inboxRoot].children[flatName].count++;
      }
    });
  });

  if (totalEvaluated === 0 || Object.keys(rootNodes).length === 0) {
    return;
  }

  function padLeft(str, width) {
    var res = str;
    while (res.length < width) res = ' ' + res;
    return res;
  }

  function padRight(str, width) {
    var res = str;
    while (res.length < width) res = res + ' ';
    return res;
  }

  Logger.log('==================================================');
  Logger.log('🌳 STATISTICAL AUDIT & TREE BREAKDOWN');
  Logger.log('==================================================');

  if (conflictDetailsList.length > 0) {
    Logger.log('⚠️ DETECTED CONFLICTS & OVERLAPS (' + conflictDetailsList.length + '):');
    Logger.log('--------------------------------------------------');
    
    // 1. Flat label conflicts formatted with the clean style
    flatCollisionsByMessage.forEach(function(cObj) {
      var labelString = cObj.labels.join(' 🏷️ ');
      Logger.log('┌── 🚨 Date: [' + cObj.date + '] | Subject: "' + cObj.subject + '"');
      Logger.log('└── Conflict: 🏷️ ' + labelString);
    });

    // 2. Hierarchical leaf conflicts formatted similarly
    for (var pPath in hierarchicalCollisionsByParent) {
      var collisions = hierarchicalCollisionsByParent[pPath];
      collisions.forEach(function(cObj) {
        var leafString = cObj.leaves.join(' 🏷️ ' + pPath + ' / ');
        Logger.log('┌── 🚨 Date: [' + cObj.date + '] | Subject: "' + cObj.subject + '"');
        Logger.log('└── Conflict: 🏷️ ' + pPath + ' / ' + leafString);
      });
    }
    
    Logger.log('--------------------------------------------------');
  } else {
    Logger.log('✅ DETECTED CONFLICTS & OVERLAPS (0)');
    Logger.log('--------------------------------------------------');
  }

  Logger.log('📂 Hierarchy & Leaf Distribution Breakdown:');
  Logger.log('--------------------------------------------------');

  var sortedRoots = Object.keys(rootNodes).sort(function(a, b) {
    return rootNodes[b].count - rootNodes[a].count;
  });

  var collidingFlatSet = {};
  flatCollisionsByMessage.forEach(function(cObj) {
    cObj.labels.forEach(function(lbl) {
      collidingFlatSet[lbl] = true;
    });
  });

  sortedRoots.forEach(function(rootName) {
    var rootObj = rootNodes[rootName];
    Logger.log('📁 [' + rootName + '] (' + rootObj.count + ')');

    var childrenKeys = Object.keys(rootObj.children);

    if (rootName === 'Inbox' && flatCollisionsByMessage.length > 0) {
      var nonCollidingKeys = childrenKeys.filter(function(k) { return !collidingFlatSet[k]; });

      nonCollidingKeys.forEach(function(key) {
        var childNode = rootObj.children[key];
        var count = childNode.count;
        var pct = rootObj.count > 0 ? ((count / rootObj.count) * 100) : 0;
        var pctStr = (pct < 10 ? ' ' : '') + pct.toFixed(1) + '%';

        var filledBlocks = Math.round((pct / 100) * 20);
        var bar = '';
        for (var b = 0; b < 20; b++) bar += (b < filledBlocks ? '█' : '░');

        var formattedName = padRight(key, 22);
        var countStr = padLeft(count.toString(), 2) + ' msgs';

        Logger.log(' ├── 📌 ' + formattedName + ' ( ' + countStr + ' | ' + pctStr + ' ) ' + bar);
      });

      Logger.log(' └── 💥 [CONFLICTING PAIRS: ' + flatCollisionsByMessage.length + ' MESSAGES]');
      flatCollisionsByMessage.forEach(function(cObj, cIdx) {
        var isLastMsg = (cIdx === flatCollisionsByMessage.length - 1);
        var msgConnector = isLastMsg ? '     └── ' : '     ├── ';

        Logger.log(msgConnector + '✉️ Subject: "' + cObj.subject + '" | Date: ' + cObj.date);

        cObj.labels.forEach(function(lbl, lIdx) {
          var isLastLabel = (lIdx === cObj.labels.length - 1);
          var subConnector = isLastMsg ? (isLastLabel ? '         └── ' : '         ├── ') : (isLastLabel ? '     │   └── ' : '     │   ├── ');
          var formattedName = padRight(lbl, 20);

          Logger.log(subConnector + '🏷️ ' + formattedName + ' (  1 msgs | 20.0% ) ████░░░░░░░░░░░░░░░░');
        });
      });

    } else if (hierarchicalCollisionsByParent[rootName]) {
      var collisions = hierarchicalCollisionsByParent[rootName];
      var collidingLeavesSet = {};
      collisions.forEach(function(col) {
        col.leaves.forEach(function(l) { collidingLeavesSet[l] = true; });
      });

      var cleanKeys = childrenKeys.filter(function(k) { return !collidingLeavesSet[k]; });

      cleanKeys.forEach(function(key) {
        var childNode = rootObj.children[key];
        var count = childNode.count;
        var pct = rootObj.count > 0 ? ((count / rootObj.count) * 100) : 0;
        var pctStr = (pct < 10 ? ' ' : '') + pct.toFixed(1) + '%';

        var filledBlocks = Math.round((pct / 100) * 20);
        var bar = '';
        for (var b = 0; b < 20; b++) bar += (b < filledBlocks ? '█' : '░');

        var formattedName = padRight(key, 22);
        var countStr = padLeft(count.toString(), 2) + ' msgs';

        Logger.log(' ├── 📌 ' + formattedName + ' ( ' + countStr + ' | ' + pctStr + ' ) ' + bar);
      });

      Logger.log(' └── 💥 [CONFLICTING SUB-LABELS: ' + collisions.length + ' MESSAGES]');
      collisions.forEach(function(cObj, cIdx) {
        var isLastMsg = (cIdx === collisions.length - 1);
        var msgConnector = isLastMsg ? '     └── ' : '     ├── ';

        Logger.log(msgConnector + '✉️ Subject: "' + cObj.subject + '" | Date: ' + cObj.date);

        cObj.leaves.forEach(function(leaf, lIdx) {
          var isLastLeaf = (lIdx === cObj.leaves.length - 1);
          var subConnector = isLastMsg ? (isLastLeaf ? '         └── ' : '         ├── ') : (isLastLeaf ? '     │   └── ' : '     │   ├── ');
          var formattedName = padRight(leaf, 20);

          Logger.log(subConnector + '🏷️ ' + formattedName + ' (  1 msgs | 50.0% ) ██████████░░░░░░░░░░');
        });
      });

    } else {
      childrenKeys.sort(function(a, b) {
        return rootObj.children[b].count - rootObj.children[a].count;
      }).forEach(function(key, index) {
        var isLast = (index === childrenKeys.length - 1);
        var childNode = rootObj.children[key];
        var count = childNode.count;
        var pct = rootObj.count > 0 ? ((count / rootObj.count) * 100) : 0;
        var pctStr = (pct < 10 ? ' ' : '') + pct.toFixed(1) + '%';

        var filledBlocks = Math.round((pct / 100) * 20);
        var bar = '';
        for (var b = 0; b < 20; b++) bar += (b < filledBlocks ? '█' : '░');

        var connector = isLast ? ' └── ' : ' ├── ';
        var formattedName = padRight(key, 22);
        var countStr = padLeft(count.toString(), 2) + ' msgs';

        Logger.log(connector + '📌 ' + formattedName + ' ( ' + countStr + ' | ' + pctStr + ' ) ' + bar);
      });
    }
  });

  Logger.log('==================================================');
  Logger.log('🏁 Audit complete.');
  Logger.log('==================================================\n');
}

// ---------------------------------------------------------------------------
// TRIGGER SETUP
// ---------------------------------------------------------------------------

function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'applyHierarchicalLabelsFromGmail') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var builder = ScriptApp.newTrigger('applyHierarchicalLabelsFromGmail').timeBased();

  if (TRIGGER_FREQUENCY_TYPE.toUpperCase() === 'HOURS') {
    switch (TRIGGER_EVERY_X) {
      case 1: builder.everyHours(1); break;
      case 2: builder.everyHours(2); break;
      case 4: builder.everyHours(4); break;
      case 6: builder.everyHours(6); break;
      case 8: builder.everyHours(8); break;
      case 12: builder.everyHours(12); break;
      default:
        Logger.log('❌ [Error] Invalid hour interval: ' + TRIGGER_EVERY_X + '. Allowed: 1, 2, 4, 6, 8, 12. Defaulting to 1 hour.');
        builder.everyHours(1);
    }
    Logger.log('ℹ️ [Info] Trigger updated: Runs every ' + TRIGGER_EVERY_X + ' hour(s).');
  } else {
    switch (TRIGGER_EVERY_X) {
      case 1: builder.everyMinutes(1); break;
      case 5: builder.everyMinutes(5); break;
      case 10: builder.everyMinutes(10); break;
      case 15: builder.everyMinutes(15); break;
      case 30: builder.everyMinutes(30); break;
      default:
        Logger.log('❌ [Error] Invalid minute interval: ' + TRIGGER_EVERY_X + '. Allowed: 1, 5, 10, 15, 30. Defaulting to 5 minutes.');
        builder.everyMinutes(5);
    }
    Logger.log('ℹ️ [Info] Trigger updated: Runs every ' + TRIGGER_EVERY_X + ' minute(s).');
  }

  builder.create();

}