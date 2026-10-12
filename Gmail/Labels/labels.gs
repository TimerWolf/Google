///////////////////////////////////////
// Gmail: labels                     //
// Revision: 3                       //
// Date created: 2025/10/16          //
// Date last modified: 2026/10/12    //
// Author: Whiskey / Google Gemini   //
///////////////////////////////////////

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

const NEW_EMAIL_WINDOW_DAYS = 2;    // Number of days back to check (0 = all messages)
const NEW_EMAIL_WINDOW_HOURS = 0;   // Number of hours back to check (0 = all messages)
const NEW_EMAIL_WINDOW_MINUTES = 0; // Number of minutes back to check (0 = all messages)
const ENABLE_INFO = false;          // Set to true for summary metrics
const ENABLE_DEBUG = false;         // Set to true for detailed message logs

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

  if (ENABLE_DEBUG) {
    Logger.log('==================================================');
    Logger.log('=== Hierarchical labels run                    ===');
    Logger.log('==================================================');
  }
  
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

  try {

    // Step 1: Fetch and map all user labels (Name -> ID & Object)
    var userLabels = GmailApp.getUserLabels();
    var labelMapByName = {};
    
    userLabels.forEach(function(lbl) {
      labelMapByName[lbl.getName()] = lbl;
    });

    // Step 2: Search threads
    var query = 'has:userlabel';
    if (NEW_EMAIL_WINDOW_DAYS > 0) {
      query = 'newer_than:' + NEW_EMAIL_WINDOW_DAYS + 'd';
    } else if (NEW_EMAIL_WINDOW_HOURS > 0) {
      query = 'newer_than:' + NEW_EMAIL_WINDOW_HOURS + 'h';
    } else if (NEW_EMAIL_WINDOW_MINUTES > 0) {
      query = 'newer_than:' + NEW_EMAIL_WINDOW_MINUTES + 'm';
    }

    var threads = GmailApp.search(query);
    
    stats.totalThreads = threads.length;

    if (ENABLE_DEBUG) {
      Logger.log('[Debug] Total cached Gmail labels in account: ' + userLabels.length);
      Logger.log('[Debug] Executed search query: "' + query + '"');
      Logger.log('[Debug] Total threads matched: ' + threads.length);
    }

    threads.forEach(function(thread, threadIndex) {
    
      var messages = thread.getMessages();

      messages.forEach(function(msg, msgIndex) {
    
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
      });
    });

    if (ENABLE_INFO) {

      Logger.log('==================================================');
      Logger.log('=== Summary & Metrics ===');
      Logger.log('==================================================');
      Logger.log('[Stats] Total Threads Searched:          ' + stats.totalThreads);
      Logger.log('[Stats] Total Messages Evaluated:        ' + stats.totalMessages);
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