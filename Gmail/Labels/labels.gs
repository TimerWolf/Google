///////////////////////////////////////
// Gmail: labels                     //
// Revision: 11                      //
// Date created: 2025/10/16          //
// Date last modified: 2026/10/13    //
// Author: Whiskey / Google Gemini   //
///////////////////////////////////////

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

const NEW_EMAIL_WINDOW_DAYS = 2;    // Number of days back to check (0 = all messages)
const NEW_EMAIL_WINDOW_HOURS = 0;   // Number of hours back to check (0 = all messages)
const NEW_EMAIL_WINDOW_MINUTES = 0; // Number of minutes back to check (0 = all messages)

// Message filtering settings
const EXCLUDE_SENT_FROM_ME = true;  // Set to true to ignore messages sent by yourself (-from:me)

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
    throw new Error('[Error] Execution time limit cannot exceed 6 minutes (360,000 ms). Google Apps Script has a hard execution limit of 6 minutes.');
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
      throw new Error('[Error] RESTRICT_TO_TARGET_LABELS is set to true, but NO target labels are defined in the TARGET_LABELS array.');
    }

    TARGET_LABELS.forEach(function(targetPattern) {
      
      // System operators (e.g. in:inbox)
      if (targetPattern.indexOf('in:') === 0 || targetPattern.indexOf('label:') === 0) {
        activeMatchedLabels.push(targetPattern);
        return;
      }

      // Wildcard logic
      if (targetPattern.indexOf('*') !== -1) {
        
        // Convert wildcard string into Regex pattern
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
          var warningMsg = '[Warning] Wildcard pattern "' + targetPattern + '" matched NO labels in your Gmail account.';
          if (STRICT_LABEL_VALIDATION) {
            throw new Error('[Error] ' + warningMsg.replace('[Warning] ', ''));
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
          var warningMsgExact = '[Warning] Target label "' + targetPattern + '" defined in TARGET_LABELS does not exist in your Gmail account.';
          if (STRICT_LABEL_VALIDATION) {
            throw new Error('[Error] ' + warningMsgExact.replace('[Warning] ', ''));
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
      throw new Error('[Error] None of the specified TARGET_LABELS or Wildcards matched any labels in your account.');
    }
  }

  // Unified start and settings log block
  if (ENABLE_DEBUG || ENABLE_SETTINGS) {
    Logger.log('==================================================');
    Logger.log('=== Hierarchical labels run & settings         ===');
    Logger.log('==================================================');
    
    if (ENABLE_SETTINGS) {
      Logger.log('[Config] NEW_EMAIL_WINDOW_DAYS:        ' + NEW_EMAIL_WINDOW_DAYS);
      Logger.log('[Config] NEW_EMAIL_WINDOW_HOURS:       ' + NEW_EMAIL_WINDOW_HOURS);
      Logger.log('[Config] NEW_EMAIL_WINDOW_MINUTES:     ' + NEW_EMAIL_WINDOW_MINUTES);
      Logger.log('[Config] EXCLUDE_SENT_FROM_ME:         ' + EXCLUDE_SENT_FROM_ME);
      Logger.log('[Config] RESTRICT_TO_TARGET_LABELS:    ' + RESTRICT_TO_TARGET_LABELS);
      Logger.log('[Config] STRICT_LABEL_VALIDATION:      ' + STRICT_LABEL_VALIDATION);
      Logger.log('[Config] TARGET_LABELS (Configured):   ' + TARGET_LABELS.join(', '));
      Logger.log('[Config] TARGET_LABELS (Active):       ' + activeMatchedLabels.join(', '));
      Logger.log('[Config] TARGET_LABELS (Missing):      ' + (missingTargetLabels.length > 0 ? missingTargetLabels.join(', ') : 'None'));
      Logger.log('[Config] MAX_EXECUTION_TIME:           ' + MAX_EXECUTION_TIME);
      Logger.log('[Config] MAX_EXECUTION_TIME_MINUTES:   ' + MAX_EXECUTION_TIME_MINUTES);
      Logger.log('[Config] MAX_EXECUTION_TIME_SECONDS:   ' + MAX_EXECUTION_TIME_SECONDS);
      Logger.log('[Config] ENABLE_INFO:                  ' + ENABLE_INFO);
      Logger.log('[Config] ENABLE_DEBUG:                 ' + ENABLE_DEBUG);
      Logger.log('[Config] ENABLE_WARNING:               ' + ENABLE_WARNING);
      Logger.log('[Config] ENABLE_SETTINGS:              ' + ENABLE_SETTINGS);
      Logger.log('[Config] TRIGGER_FREQUENCY_TYPE:       ' + TRIGGER_FREQUENCY_TYPE);
      Logger.log('[Config] TRIGGER_EVERY_X:              ' + TRIGGER_EVERY_X);
      Logger.log('==================================================');
    }
  }
  
  var startTime = new Date().getTime();

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    
    if (ENABLE_DEBUG) {
      Logger.log('[Debug] Could not obtain lock. Another instance is running.');
    }
    
    return;
  
  }

  // Counters for statistics
  var stats = {
    totalThreads: 0,
    totalMessages: 0,
    messagesWithLabels: 0,
    messagesWithoutLabels: 0,
    hierarchicalLabelsFound: 0,
    labelsAddedCount: 0,
    createdGmailLabels: 0
  };

  var timeOutReached = false;

  try {

    // Step 2: Build exact Unix timestamp search query and label restriction
    var totalWindowMs = (NEW_EMAIL_WINDOW_DAYS * 24 * 60 * 60 * 1000) +
                        (NEW_EMAIL_WINDOW_HOURS * 60 * 60 * 1000) +
                        (NEW_EMAIL_WINDOW_MINUTES * 60 * 1000);

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

    // Time window restriction
    if (totalWindowMs > 0) {
      var afterSec = Math.floor((startTime - totalWindowMs) / 1000);
      var beforeSec = Math.floor(startTime / 1000);
      queryParts.push('after:' + afterSec);
      queryParts.push('before:' + beforeSec);
    }

    // Exclude sent messages restriction
    if (EXCLUDE_SENT_FROM_ME) {
      queryParts.push('-from:me');
    }

    var query = queryParts.join(' ');

    var threads = GmailApp.search(query);
    
    stats.totalThreads = threads.length;

    if (ENABLE_DEBUG) {
      Logger.log('[Debug] Total cached Gmail labels in account: ' + userLabels.length);
      Logger.log('[Debug] Time window configuration: ' + NEW_EMAIL_WINDOW_DAYS + ' day(s), ' + NEW_EMAIL_WINDOW_HOURS + ' hour(s), ' + NEW_EMAIL_WINDOW_MINUTES + ' minute(s)');
      Logger.log('[Debug] Restrict to target labels: ' + RESTRICT_TO_TARGET_LABELS);
      Logger.log('[Debug] Strict label validation: ' + STRICT_LABEL_VALIDATION);
      Logger.log('[Debug] Exclude sent messages (-from:me): ' + EXCLUDE_SENT_FROM_ME);
      Logger.log('[Debug] Executed search query: "' + query + '"');
      Logger.log('[Debug] Total threads matched: ' + threads.length);
    }

    for (var threadIndex = 0; threadIndex < threads.length; threadIndex++) {
    
      // Check if execution time limit is approaching (only if MAX_EXECUTION_TIME is true)
      if (MAX_EXECUTION_TIME && (new Date().getTime() - startTime > totalLimitMs)) {
        timeOutReached = true;
        if (ENABLE_WARNING) {
          Logger.log('[Warning] Execution time limit (' + MAX_EXECUTION_TIME_MINUTES + 'm ' + MAX_EXECUTION_TIME_SECONDS + 's) reached. Stopping gracefully.');
        }
        break;
      }

      var thread = threads[threadIndex];
      var messages = thread.getMessages();

      for (var msgIndex = 0; msgIndex < messages.length; msgIndex++) {
        var msg = messages[msgIndex];

        // Double check exact millisecond timestamp if window is active
        if (totalWindowMs > 0 && msg.getDate().getTime() < (startTime - totalWindowMs)) {
          continue; 
        }
    
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
                  Logger.log('[Debug] Created new label in Gmail account: "' + currentPath + '"');
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
            Logger.log('[Debug] Subject: "' + msg.getSubject() + '"');
            Logger.log('[Debug] Existing labels on message: [' + currentMsgLabelNames.join(', ') + ']');
            Logger.log('[Debug] -> FIXED/ADDED ' + missingLabelIdsToAdd.length + ' LABEL(S): [' + missingLabelNamesAdded.join(', ') + ']');

          }
        }
      }
    }

    if (ENABLE_INFO || timeOutReached) {

      var executionTimeMs = new Date().getTime() - startTime;
      var executionTimeSec = (executionTimeMs / 1000).toFixed(2);
      var sentStatusText = EXCLUDE_SENT_FROM_ME ? 'false' : 'true';

      Logger.log('==================================================');
      Logger.log('=== Summary & Metrics (timeout reached: ' + timeOutReached + ') ===');
      Logger.log('==================================================');
      Logger.log('[Stats] Execution Time:                  ' + executionTimeSec + ' seconds (' + executionTimeMs + ' ms)');
      Logger.log('[Stats] Search Scope Mode:               ' + (RESTRICT_TO_TARGET_LABELS ? 'Restricted (Wildcard Pattern)' : 'All User Labels'));
      Logger.log('[Stats] Target Labels Active:            ' + activeMatchedLabels.join(', '));
      Logger.log('[Stats] Target Labels Missing:           ' + (missingTargetLabels.length > 0 ? missingTargetLabels.join(', ') : 'None'));
      Logger.log('[Stats] Total Threads Searched:          ' + stats.totalThreads);
      Logger.log('[Stats] Total Messages Evaluated:        ' + stats.totalMessages);
      Logger.log('[Stats] Sent Messages Filter:            ' + sentStatusText);
      Logger.log('[Stats] Messages WITH user labels:       ' + stats.messagesWithLabels);
      Logger.log('[Stats] Messages WITHOUT user labels:    ' + stats.messagesWithoutLabels);
      Logger.log('[Stats] Hierarchical labels (with /):    ' + stats.hierarchicalLabelsFound);
      Logger.log('[Stats] New Gmail labels created:        ' + stats.createdGmailLabels);
      Logger.log('[Stats] Total Parent Labels Applied:     ' + stats.labelsAddedCount);
      Logger.log('==================================================\n');
    
    }
  } catch(e) {
    Logger.log('[Error] Exception: ' + e.message + '\n' + e.stack);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// TRIGGER SETUP
// ---------------------------------------------------------------------------

function setupTrigger() {
  
  // Remove previous triggers for this function to avoid duplicates
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
        Logger.log('[Error] Invalid hour interval: ' + TRIGGER_EVERY_X + '. Allowed: 1, 2, 4, 6, 8, 12. Defaulting to 1 hour.');
        builder.everyHours(1);
    
    }
  
    Logger.log('[Info] Trigger updated: Runs every ' + TRIGGER_EVERY_X + ' hour(s).');
  
  } else {
  
    switch (TRIGGER_EVERY_X) {
    
      case 1: builder.everyMinutes(1); break;
      case 5: builder.everyMinutes(5); break;
      case 10: builder.everyMinutes(10); break;
      case 15: builder.everyMinutes(15); break;
      case 30: builder.everyMinutes(30); break;
  
      default:
        Logger.log('[Error] Invalid minute interval: ' + TRIGGER_EVERY_X + '. Allowed: 1, 5, 10, 15, 30. Defaulting to 5 minutes.');
        builder.everyMinutes(5);
    
    }
  
    Logger.log('[Info] Trigger updated: Runs every ' + TRIGGER_EVERY_X + ' minute(s).');
  
  }

  builder.create();

}