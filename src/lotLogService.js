// src/lotLogService.js
/**
 * Factory Suite Pro - Centralized Multi-User Lot Change & Audit Logging Service
 * Connects directly to Google Sheets (Spreadsheet ID: 1dnukAAjyZy-W6oiRvT-6rcyTE5Rlu3u4GJqVXyN382I)
 * Automatically syncs & broadcasts real-time notifications across all computers and users.
 */

const STORAGE_KEY = "factory_lot_change_logs";
const BROADCAST_CHANNEL_NAME = "factory_lot_logs_channel";

const SPREADSHEET_ID =
  process.env.REACT_APP_LOT_LOGS_SPREADSHEET_ID ||
  "1dnukAAjyZy-W6oiRvT-6rcyTE5Rlu3u4GJqVXyN382I";
const SHEET_RANGE =
  process.env.REACT_APP_LOT_LOGS_SHEET_RANGE || "LotLogs!A1:I1000";
const API_KEY = process.env.REACT_APP_GOOGLE_API_KEY || "";

const DEFAULT_WEBHOOK_URL =
  "https://script.google.com/macros/s/AKfycby7j4mUEdmUM7TdFMPsonig_elsdYsRHb389ZPGka7yRYxDlOnZRFto_oI-Uqp0vVIJ/exec";

// Initial fallback seed data
const INITIAL_LOGS = [
  {
    id: "LOG-1001",
    lotNumber: "11028",
    changeDetails: "Updated Job Order status to Priority-High and adjusted quantity by +50 pcs",
    changedBy: "Gourav (Admin)",
    permissionBy: "Monu (Production Head)",
    category: "Job Order",
    priority: "High Alert",
    timestamp: new Date(Date.now() - 15 * 60 * 1000).toISOString(),
    timeFormatted: "15m ago",
    seenBy: []
  }
];

/**
 * Format relative time string (e.g., 5m ago, 2h ago)
 */
const formatRelativeTime = (timestampStr) => {
  if (!timestampStr) return "Recently";
  try {
    const date = new Date(timestampStr);
    const diffSeconds = Math.floor((Date.now() - date.getTime()) / 1000);
    if (diffSeconds < 60) return "Just now";
    const diffMinutes = Math.floor(diffSeconds / 60);
    if (diffMinutes < 60) return `${diffMinutes}m ago`;
    const diffHours = Math.floor(diffMinutes / 60);
    if (diffHours < 24) return `${diffHours}h ago`;
    const diffDays = Math.floor(diffHours / 24);
    return `${diffDays}d ago`;
  } catch (e) {
    return timestampStr;
  }
};

/**
 * Normalize username for robust comparison (strips role or timestamp in parentheses)
 */
const normalizeUserName = (name) => {
  if (!name) return "";
  return String(name).split("(")[0].trim().toLowerCase();
};

/**
 * Fetch live Lot Logs directly from Google Sheet (Accessible by ALL users on all computers)
 */
export const fetchLotLogsFromSheet = async () => {
  try {
    const encodedRange = encodeURIComponent(SHEET_RANGE);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}?key=${API_KEY}`;
    const response = await fetch(url);
    if (!response.ok) return [];

    const json = await response.json();
    const rows = json.values || [];
    if (rows.length <= 1) return [];

    const logs = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0 || !row[0]) continue;

      const rawTimestamp = row[0] || "";
      const logId = row[1] || `LOG-${i}`;
      const lotNum = row[2] || "";
      const details = row[3] || "";
      const cBy = row[4] || "User";
      const pBy = row[5] || "Authorized Head";
      const cat = row[6] || "Job Order";
      const prio = row[7] || "Standard Update";
      const seenByRaw = row[8] || "";
      const seenByList = seenByRaw
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);

      logs.push({
        id: logId,
        lotNumber: lotNum,
        changeDetails: details,
        changedBy: cBy,
        permissionBy: pBy,
        category: cat,
        priority: prio,
        timestamp: rawTimestamp,
        timeFormatted: formatRelativeTime(rawTimestamp),
        seenBy: seenByList
      });
    }

    // Newest logs first
    logs.reverse();

    // Cache in local storage for instant offline fallback
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(logs));
    } catch (e) {}

    return logs;
  } catch (err) {
    console.error("Error fetching LotLogs from Google Sheet:", err);
    return [];
  }
};

/**
 * Synchronous getter from local storage cache
 */
export const getLotLogs = () => {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
  } catch (err) {
    console.error("Error reading lot logs from localStorage:", err);
  }
  return INITIAL_LOGS;
};

/**
 * Check if a log has been read by a specific user (either in Sheet or Local Cache)
 */
export const isLogReadByUser = (log, userName) => {
  if (!log || !userName) return false;

  const targetName = normalizeUserName(userName);
  if (!targetName) return false;

  // 1. Check Google Sheet Seen By list
  if (Array.isArray(log.seenBy) && log.seenBy.some((u) => {
    const sheetUser = normalizeUserName(u);
    return sheetUser === targetName || sheetUser.includes(targetName) || targetName.includes(sheetUser);
  })) {
    return true;
  }

  // 2. Check local user read cache
  const storageKey = `user_read_logs_${targetName}`;
  try {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      const set = new Set(JSON.parse(saved));
      if (set.has(log.id)) return true;
    }
  } catch (e) {}

  return false;
};

/**
 * Mark a log as read by a specific user in Google Sheet & Local Cache
 */
export const markLogAsReadInSheet = async (logId, userName) => {
  if (!logId || !userName) return;

  const targetName = normalizeUserName(userName);
  const cleanUserName = userName.trim();

  // 1. Save to local storage cache for instant UI feedback
  const storageKey = `user_read_logs_${targetName}`;
  try {
    const saved = localStorage.getItem(storageKey);
    const set = new Set(saved ? JSON.parse(saved) : []);
    set.add(logId);
    localStorage.setItem(storageKey, JSON.stringify(Array.from(set)));
  } catch (e) {}

  // 2. Dispatch local event to update Navbar & UI components
  window.dispatchEvent(new CustomEvent("lot_log_marked_read", { detail: { logId, userName: cleanUserName } }));

  // 3. Send Webhook request to Google Apps Script to update Google Sheet column I (Seen By)
  const webhookUrl =
    (process.env.REACT_APP_LOT_LOGS_WEBHOOK_URL && process.env.REACT_APP_LOT_LOGS_WEBHOOK_URL.trim())
      ? process.env.REACT_APP_LOT_LOGS_WEBHOOK_URL.trim()
      : DEFAULT_WEBHOOK_URL;

  if (!webhookUrl) return;

  try {
    await fetch(webhookUrl, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify({
        action: "mark_read",
        logId: logId,
        userName: cleanUserName
      }),
      mode: "no-cors"
    });
    console.log(`Log ${logId} marked as read for user ${cleanUserName} in Google Sheet.`);
  } catch (err) {
    console.error("Error marking log as read in Apps Script Webhook:", err);
  }
};

/**
 * Asynchronously send new lot log entry to Google Sheets via Apps Script Webhook
 */
const syncToGoogleSheets = async (logEntry) => {
  const webhookUrl =
    (process.env.REACT_APP_LOT_LOGS_WEBHOOK_URL && process.env.REACT_APP_LOT_LOGS_WEBHOOK_URL.trim())
      ? process.env.REACT_APP_LOT_LOGS_WEBHOOK_URL.trim()
      : DEFAULT_WEBHOOK_URL;

  if (!webhookUrl) return;

  try {
    await fetch(webhookUrl, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(logEntry),
      mode: "no-cors"
    });
    console.log("Successfully sent lot log to Google Sheet via Apps Script:", logEntry.id);
  } catch (err) {
    console.error("Failed to sync lot log to Google Sheet Apps Script:", err);
  }
};

/**
 * Add a new Lot Change Log entry and broadcast to notification systems & Google Sheet
 */
export const addLotLog = ({ lotNumber, changeDetails, changedBy, permissionBy, category = "Job Order", priority = "Standard Update" }) => {
  const currentLogs = getLotLogs();
  
  const creator = String(changedBy || "System User").trim();

  const newLog = {
    id: `LOG-${Date.now().toString().slice(-5)}`,
    lotNumber: String(lotNumber || "").trim(),
    changeDetails: String(changeDetails || "").trim(),
    changedBy: creator,
    permissionBy: String(permissionBy || "Authorized Head").trim(),
    category: String(category || "Job Order").trim(),
    priority: String(priority || "Standard Update").trim(),
    timestamp: new Date().toISOString(),
    timeFormatted: "Just now",
    seenBy: [] // Empty until user clicks "Mark as Read & Acknowledge"
  };

  const updatedLogs = [newLog, ...currentLogs];
  
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(updatedLogs));
  } catch (err) {
    console.error("Error saving lot log to localStorage:", err);
  }

  // 1. Dispatch window event for same-tab subscribers
  window.dispatchEvent(new CustomEvent("lot_change_logged", { detail: newLog }));

  // 2. Broadcast across browser tabs via BroadcastChannel
  try {
    if ("BroadcastChannel" in window) {
      const bc = new BroadcastChannel(BROADCAST_CHANNEL_NAME);
      bc.postMessage(newLog);
      bc.close();
    }
  } catch (err) {
    console.warn("BroadcastChannel error:", err);
  }

  // 3. Trigger background sync to Google Sheets via Apps Script Webhook
  syncToGoogleSheets(newLog);

  return newLog;
};

// Track known log IDs to trigger notifications for newly created logs across computers
const knownLogIds = new Set();

/**
 * Subscribe to real-time lot change logs (for Navbar notifications, NotificationCenter & LotLogs page)
 * Includes fast 3-second background polling from Google Sheet for instant multi-user notifications.
 */
export const subscribeToLotLogs = (onNewLog) => {
  const handleCustomEvent = (event) => {
    if (event.detail && typeof onNewLog === "function") {
      onNewLog(event.detail);
    }
  };

  window.addEventListener("lot_change_logged", handleCustomEvent);

  let bc = null;
  if ("BroadcastChannel" in window) {
    try {
      bc = new BroadcastChannel(BROADCAST_CHANNEL_NAME);
      bc.onmessage = (event) => {
        if (event.data && typeof onNewLog === "function") {
          onNewLog(event.data);
        }
      };
    } catch (e) {}
  }

  const checkLogsAndNotify = (logs) => {
    logs.forEach((log) => {
      if (!knownLogIds.has(log.id)) {
        knownLogIds.add(log.id);
        if (typeof onNewLog === "function") {
          onNewLog(log);
        }
      }
    });
  };

  // Fast 3-second polling from Google Sheet to deliver instant notifications across all computers
  const pollInterval = setInterval(async () => {
    const liveLogs = await fetchLotLogsFromSheet();
    checkLogsAndNotify(liveLogs);
  }, 3000);

  // Initial fetch from Google Sheet
  fetchLotLogsFromSheet().then((logs) => {
    checkLogsAndNotify(logs);
  });

  return () => {
    window.removeEventListener("lot_change_logged", handleCustomEvent);
    if (bc) bc.close();
    clearInterval(pollInterval);
  };
};
