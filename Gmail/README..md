# 🏷️ Gmail Label Engine

[![Google Apps Script](https://img.shields.io/badge/Platform-Google%20Apps%20Script-4285F4?logo=google)](https://developers.google.com/apps-script)
[![Gmail API](https://img.shields.io/badge/API-Gmail%20v1%20REST-EA4335?logo=gmail)](https://developers.google.com/gmail/api)
[![Architecture](https://img.shields.io/badge/Architecture-REST%20Message--Level-009688)](#-architecture-overview)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

An Google Apps Script solution that enforces taxonomy consistency across nested Gmail labels. By executing direct message-level modifications via the native Gmail REST API, it guarantees strict hierarchical integrity regardless of whether users operate in Threaded (Conversation View) or Unthreaded modes.

---

## 📋 Table of Contents
- [Architecture Overview](#-architecture-overview)
- [Enterprise Features](#-enterprise-features)
- [Prerequisites](#-prerequisites)
- [Deployment & Setup](#-deployment--setup)
  - [1. Script Initialization](#1-script-initialization)
  - [2. Enable Gmail REST API](#2-enable-gmail-rest-api)
  - [3. Configuration](#3-configuration)
  - [4. Authorization & Automation Setup](#4-authorization--automation-setup)
- [Configuration Matrix](#-configuration-matrix)
- [Execution Metrics & Auditing](#-execution-metrics--auditing)
- [Quotas & Performance Limits](#-quotas--performance-limits)
- [Troubleshooting](#-troubleshooting)

---

## 🏗️ Architecture Overview

The system replaces thread-level label actions with direct message-level REST execution to eliminate UI desynchronization in unthreaded views.

| Component | Class / API | Responsibility |
| :--- | :--- | :--- |
| **Search Engine** | `GmailApp.search()` | Queries messages matching lookback window (`NEW_EMAIL_WINDOW_DAYS`). |
| **Inspector** | `Gmail.Users.Messages.get` | Bypasses cache to fetch exact message-level `labelIds`. |
| **Taxonomy Resolver** | Local Parser | Splits label paths (`A/B/C`) and evaluates missing parents. |
| **Auto-Provisioner** | `GmailApp.createLabel()` | Instantiates missing intermediate parent labels dynamically. |
| **Modifier** | `Gmail.Users.Messages.modify` | Executes RFC822 metadata updates directly on message IDs. |
| **Concurrency Guard** | `LockService` | Enforces single-execution mutex across timed triggers. |

---

## ⚡ Enterprise Features

- **Message-Level Precision:** Modifies Gmail RFC822 metadata directly, overcoming native UI label-propagation lag.
- **Taxonomy Self-Healing:** Automatically instantiates missing intermediate parent labels in the user's account hierarchy.
- **Race Condition Prevention:** Implements `LockService` mutexes to prevent overlapping execution runs.
- **Idempotent Trigger Management:** Automatically cleans up duplicate project triggers upon re-configuration.
- **Zero-Trust Audit Logging:** Comprehensive diagnostic output featuring execution timestamps, memory states, and operational metrics.

---

## 🛠 Prerequisites

- A Google Workspace or Gmail Account.
- Pre-configured Gmail filters or automated workflows assigning nested child labels (e.g., `Finance/Invoices/2026`).

---

## 🚀 Deployment & Setup

### 1. Script Initialization
1. Navigate to Google Apps Script Dashboard: https://script.google.com/
2. Create a New Project and name it `Gmail-Label-Engine`.
3. Overwrite the contents of `Code.gs` with the production source code.

### 2. Enable Gmail REST API
> ⚠️ **CRITICAL:** Advanced Google Services must be enabled to expose the `Gmail` object.

1. In the left navigation pane, click the `+` icon next to Services.
2. Locate **Gmail API** in the API catalog.
3. Set the Identifier to `Gmail` and click **Add**.

---

### 3. Configuration
Define your operational parameters in the global scope at the top of `Code.gs`:

```javascript
// Lookback scope (0 = scan all labeled messages, >0 = scan recent N days)
const NEW_EMAIL_WINDOW_DAYS = 2;

// Logging level controls
const ENABLE_INFO = false;     // Summary statistics block
const ENABLE_DEBUG = false;    // Granular message trace output

// Scheduler settings (Time-driven execution)
// TYPE: 'MINUTES' (1, 5, 10, 15, 30) or 'HOURS' (1, 2, 4, 6, 8, 12)
const TRIGGER_FREQUENCY_TYPE = 'MINUTES'; 
const TRIGGER_EVERY_X = 30; 
```

---

### 4. Authorization & Automation Setup
1. Select `setupTrigger` from the function selector in the top toolbar.
2. Click **Run**.
3. Accept the OAuth consent prompt:
   - Click **Review Permissions**.
   - Choose your target workspace account.
   - Click **Advanced** -> **Go to Gmail-Label-Engine (unsafe)**.
   - Click **Allow**.
4. Verify execution success in the Execution log:
   `[Info] Trigger updated: Runs every 30 minute(s).`

---

## ⚙ Configuration Matrix

| Parameter | Type | Default | Valid Values | Description |
| :--- | :--- | :--- | :--- | :--- |
| `NEW_EMAIL_WINDOW_DAYS` | `Number` | `2` | `0` to `365` | Days to scan back. `0` disables time filtering. |
| `ENABLE_INFO` | `Boolean` | `false` | `true`, `false` | Enables summary metrics output at end of run. |
| `ENABLE_DEBUG` | `Boolean` | `false` | `true`, `false` | Enables verbose trace logs for every updated message. |
| `TRIGGER_FREQUENCY_TYPE`| `String` | `'MINUTES'`| `'MINUTES'`, `'HOURS'` | Frequency unit for automated time-driven triggers. |
| `TRIGGER_EVERY_X` | `Number` | `30` | `1, 5, 10, 15, 30` (min) / `1, 2, 4, 6, 8, 12` (hrs) | Frequency interval value enforced by GAS API. |

---

## 📈 Execution Metrics & Auditing

When `ENABLE_INFO = true`, the engine emits structured telemetry logs at execution termination:

```text
==================================================
=== RUN SUMMARY & METRICS ===
==================================================
[STATS] Total Threads Searched:          0
[STATS] Total Messages Evaluated:        9
[STATS] Messages WITH user labels:       0
[STATS] Messages WITHOUT user labels:    0
[STATS] Hierarchical labels (with /):    0
[STATS] New Gmail labels created:        0
[STATS] Total Parent Labels Applied:     0
==================================================
```

---

## ⚡ Quotas & Performance Limits

Be mindful of Google Workspace API quotas when deploying in high-volume environments:

- **Script Runtime Limit:** 6 minutes / execution (Consumer & Workspace).
- **Gmail API Quota:** 250 quota units / second per user.
- **Recommended Scan Window:** Set `NEW_EMAIL_WINDOW_DAYS = 1` or `2` for high-volume inboxes to remain well within execution time limits.

---

## 🔍 Troubleshooting

| Symptom / Error | Cause | Resolution |
| :--- | :--- | :--- |
| `ReferenceError: "Gmail" is not defined` | Gmail API Service is disabled | Re-visit Step 2 and enable the Gmail API service in Apps Script. |
| `LockService: Could not obtain lock` | Previous run still executing | Expected behavior. Concurrent instances are blocked to avoid race conditions. |
| `Invalid hour/minute interval` | Unsupported frequency value | Check the Configuration Matrix for allowed trigger values. |