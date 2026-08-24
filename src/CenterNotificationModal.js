// src/CenterNotificationModal.js
import React, { useState, useEffect, useCallback } from "react";
import ReactDOM from "react-dom";
import { getCurrentUser } from "./auth";
import {
  fetchLotLogsFromSheet,
  markLogAsReadInSheet,
  isLogReadByUser,
  subscribeToLotLogs
} from "./lotLogService";

export default function CenterNotificationModal() {
  const currentUser = getCurrentUser();
  const userName = currentUser?.name || "User";

  const [activeLog, setActiveLog] = useState(null);
  const [unreadLogs, setUnreadLogs] = useState([]);
  const [isAcknowledging, setIsAcknowledging] = useState(false);

  // Play audio chime when center notification opens
  const playChime = useCallback(() => {
    try {
      const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();

      osc.type = "sine";
      osc.frequency.setValueAtTime(587.33, audioCtx.currentTime); // D5
      osc.frequency.exponentialRampToValueAtTime(880, audioCtx.currentTime + 0.15); // A5

      gain.gain.setValueAtTime(0.3, audioCtx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.35);

      osc.connect(gain);
      gain.connect(audioCtx.destination);

      osc.start();
      osc.stop(audioCtx.currentTime + 0.35);
    } catch (e) {}
  }, []);

  // Fetch unread logs and trigger modal if unread items exist
  const scanUnreadLogs = useCallback(async () => {
    const liveLogs = await fetchLotLogsFromSheet();
    const unread = liveLogs.filter((log) => !isLogReadByUser(log, userName));
    setUnreadLogs(unread);

    if (unread.length > 0) {
      setActiveLog((prev) => {
        if (!prev) {
          playChime();
          return unread[0];
        }
        return prev;
      });
    } else {
      setActiveLog(null);
    }
  }, [userName, playChime]);

  // Initial scan & real-time subscription
  useEffect(() => {
    scanUnreadLogs();

    const unsubscribe = subscribeToLotLogs((newLog) => {
      if (!isLogReadByUser(newLog, userName)) {
        setActiveLog(newLog);
        playChime();
        setUnreadLogs((prev) => [newLog, ...prev.filter((l) => l.id !== newLog.id)]);
      }
    });

    const handleLocalRead = (e) => {
      const { logId } = e.detail || {};
      if (logId) {
        setUnreadLogs((prev) => {
          const updated = prev.filter((l) => l.id !== logId);
          setActiveLog((curr) => (curr && curr.id === logId ? (updated[0] || null) : curr));
          return updated;
        });
      }
    };

    window.addEventListener("lot_log_marked_read", handleLocalRead);

    return () => {
      unsubscribe();
      window.removeEventListener("lot_log_marked_read", handleLocalRead);
    };
  }, [userName, scanUnreadLogs, playChime]);

  // 30-Minute Recurring Reminder Timer for Ignored / Unread Notifications
  useEffect(() => {
    const THIRTY_MINUTES_MS = 30 * 60 * 1000;
    const timer = setInterval(() => {
      console.log("30-Minute Reminder: Re-scanning unread lot notifications from Google Sheet...");
      scanUnreadLogs();
    }, THIRTY_MINUTES_MS);

    return () => clearInterval(timer);
  }, [scanUnreadLogs]);

  // Handle Mark as Read & Acknowledge Button Click
  const handleAcknowledge = async () => {
    if (!activeLog) return;
    setIsAcknowledging(true);

    try {
      // Update Google Sheet Column I ("Seen By") & Local Cache
      await markLogAsReadInSheet(activeLog.id, userName);

      const remaining = unreadLogs.filter((l) => l.id !== activeLog.id);
      setUnreadLogs(remaining);

      if (remaining.length > 0) {
        setActiveLog(remaining[0]);
        playChime();
      } else {
        setActiveLog(null);
      }
    } catch (err) {
      console.error("Error acknowledging notification:", err);
      setActiveLog(null);
    } finally {
      setIsAcknowledging(false);
    }
  };

  if (!activeLog) return null;

  return ReactDOM.createPortal(
    <div style={styles.overlay}>
      <div style={styles.modalCard}>
        {/* Header */}
        <div style={styles.header}>
          <div style={styles.badgeRow}>
            <span style={styles.alertBadge}>🔔 NEW LOT NOTIFICATION</span>
            <span style={styles.priorityBadge(activeLog.priority)}>
              {activeLog.priority || "Standard Update"}
            </span>
          </div>
          <h2 style={styles.title}>Lot #{activeLog.lotNumber} Action Alert</h2>
        </div>

        {/* Body Content */}
        <div style={styles.body}>
          <div style={styles.detailBox}>
            <p style={styles.changeText}>{activeLog.changeDetails}</p>
          </div>

          <div style={styles.metaGrid}>
            <div style={styles.metaItem}>
              <span style={styles.metaLabel}>Category</span>
              <span style={styles.metaVal}>{activeLog.category || "Job Order"}</span>
            </div>

            <div style={styles.metaItem}>
              <span style={styles.metaLabel}>Changed By</span>
              <span style={styles.metaVal}>{activeLog.changedBy || "System User"}</span>
            </div>

            <div style={styles.metaItem}>
              <span style={styles.metaLabel}>Authorized By</span>
              <span style={styles.metaVal}>{activeLog.permissionBy || "Production Head"}</span>
            </div>

            <div style={styles.metaItem}>
              <span style={styles.metaLabel}>Time</span>
              <span style={styles.metaVal}>{activeLog.timeFormatted || activeLog.timestamp}</span>
            </div>
          </div>
        </div>

        {/* Footer Actions */}
        <div style={styles.footer}>
          <p style={styles.infoNote}>
            ℹ️ Clicking <strong>Mark as Read</strong> updates the Google Sheet and ensures this notification will not trigger again on your window.
          </p>
          <button
            style={styles.ackBtn}
            onClick={handleAcknowledge}
            disabled={isAcknowledging}
          >
            {isAcknowledging ? "Updating Sheet..." : "✓ Mark as Read & Acknowledge (Update Sheet)"}
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}

const styles = {
  overlay: {
    position: "fixed",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(15, 23, 42, 0.75)",
    backdropFilter: "blur(8px)",
    WebkitBackdropFilter: "blur(8px)",
    zIndex: 999999,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "20px"
  },
  modalCard: {
    width: "100%",
    maxWidth: "540px",
    backgroundColor: "#ffffff",
    borderRadius: "16px",
    boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.35), 0 0 0 1px rgba(255, 255, 255, 0.1)",
    overflow: "hidden"
  },
  header: {
    background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #4338ca 100%)",
    padding: "24px 28px",
    color: "#ffffff"
  },
  badgeRow: {
    display: "flex",
    gap: "10px",
    alignItems: "center",
    marginBottom: "10px"
  },
  alertBadge: {
    background: "rgba(239, 68, 68, 0.2)",
    border: "1px solid rgba(248, 113, 113, 0.5)",
    color: "#fca5a5",
    padding: "4px 10px",
    borderRadius: "20px",
    fontSize: "11px",
    fontWeight: 800,
    letterSpacing: "0.5px"
  },
  priorityBadge: (prio) => ({
    background: prio === "High Alert" ? "rgba(234, 179, 8, 0.2)" : "rgba(255, 255, 255, 0.15)",
    border: "1px solid rgba(255, 255, 255, 0.2)",
    color: prio === "High Alert" ? "#fef08a" : "#e0e7ff",
    padding: "4px 10px",
    borderRadius: "20px",
    fontSize: "11px",
    fontWeight: 700
  }),
  title: {
    margin: 0,
    fontSize: "20px",
    fontWeight: 800,
    color: "#ffffff"
  },
  body: {
    padding: "24px 28px"
  },
  detailBox: {
    background: "#f8fafc",
    borderLeft: "4px solid #4338ca",
    borderRadius: "8px",
    padding: "16px",
    marginBottom: "20px"
  },
  changeText: {
    margin: 0,
    fontSize: "14px",
    lineHeight: 1.6,
    color: "#1e293b",
    fontWeight: 600
  },
  metaGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "14px"
  },
  metaItem: {
    display: "flex",
    flexDirection: "column",
    gap: "2px"
  },
  metaLabel: {
    fontSize: "11px",
    fontWeight: 700,
    color: "#64748b",
    textTransform: "uppercase",
    letterSpacing: "0.5px"
  },
  metaVal: {
    fontSize: "13px",
    fontWeight: 700,
    color: "#0f172a"
  },
  footer: {
    padding: "18px 28px 24px",
    background: "#f1f5f9",
    borderTop: "1px solid #e2e8f0",
    display: "flex",
    flexDirection: "column",
    gap: "12px"
  },
  infoNote: {
    margin: 0,
    fontSize: "12px",
    color: "#64748b",
    lineHeight: 1.4
  },
  ackBtn: {
    width: "100%",
    padding: "14px",
    background: "linear-gradient(135deg, #15803d 0%, #16a34a 100%)",
    color: "#ffffff",
    border: "none",
    borderRadius: "10px",
    fontSize: "15px",
    fontWeight: 800,
    cursor: "pointer",
    boxShadow: "0 4px 12px rgba(22, 163, 74, 0.35)",
    transition: "all 0.2s ease"
  }
};
