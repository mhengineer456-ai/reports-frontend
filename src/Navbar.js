// src/Navbar.js
import React, { useState, useEffect, useRef } from "react";
import { useHistory, useLocation } from "react-router-dom";
import { getCurrentUser, logoutUser } from "./auth";
import { motion, AnimatePresence } from "framer-motion";
import { subscribeToLotLogs, fetchLotLogsFromSheet, isLogReadByUser } from "./lotLogService";
import { fetchTodayMilestones } from "./productionMilestones";

import CenterNotificationModal from "./CenterNotificationModal";

const Navbar = () => {
  const history = useHistory();
  const location = useLocation();
  const currentUser = getCurrentUser();

  const [notificationsOpen, setNotificationsOpen] = useState(false);
  const [notifications, setNotifications] = useState([]);
  const [activeToast, setActiveToast] = useState(null);

  // Web Audio API Synthesized Crystal Notification Chime
  const playNotificationChime = () => {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = new AudioCtx();

      const osc1 = ctx.createOscillator();
      const gain1 = ctx.createGain();
      osc1.type = "sine";
      osc1.frequency.setValueAtTime(659.25, ctx.currentTime);
      gain1.gain.setValueAtTime(0.15, ctx.currentTime);
      gain1.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
      osc1.connect(gain1);
      gain1.connect(ctx.destination);

      const osc2 = ctx.createOscillator();
      const gain2 = ctx.createGain();
      osc2.type = "sine";
      osc2.frequency.setValueAtTime(987.77, ctx.currentTime + 0.08);
      gain2.gain.setValueAtTime(0.2, ctx.currentTime + 0.08);
      gain2.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.5);
      osc2.connect(gain2);
      gain2.connect(ctx.destination);

      osc1.start(ctx.currentTime);
      osc1.stop(ctx.currentTime + 0.35);
      osc2.start(ctx.currentTime + 0.08);
      osc2.stop(ctx.currentTime + 0.5);
    } catch (err) {
      console.warn("Audio playback issue:", err);
    }
  };

  const renderFormattedDesc = (desc, supervisor) => {
    if (!desc) return null;

    const supMatch = desc.match(/^(.*?)(?:issued to supervisor|under supervisor|supervisor)\s+([A-Za-z0-9\s\.\-_]+?)(?=\s*\[|\s*\||$)(.*)$/i);
    
    if (supMatch) {
      const [_, prefix, supName, suffix] = supMatch;
      return (
        <span>
          {prefix}{" "}
          <span
            style={{
              background: "linear-gradient(135deg, #fef3c7 0%, #fde68a 100%)",
              color: "#92400e",
              border: "1px solid #f59e0b",
              borderRadius: "8px",
              padding: "2px 8px",
              fontWeight: "800",
              fontSize: "0.78rem",
              display: "inline-flex",
              alignItems: "center",
              gap: "4px",
              margin: "0 4px",
              boxShadow: "0 2px 4px rgba(245, 158, 11, 0.15)"
            }}
          >
            👤 Supervisor: {supName.trim()}
          </span>{" "}
          {suffix}
        </span>
      );
    }

    if (supervisor) {
      return (
        <span>
          {desc}{" "}
          <span
            style={{
              background: "linear-gradient(135deg, #fef3c7 0%, #fde68a 100%)",
              color: "#92400e",
              border: "1px solid #f59e0b",
              borderRadius: "8px",
              padding: "2px 8px",
              fontWeight: "800",
              fontSize: "0.78rem",
              display: "inline-flex",
              alignItems: "center",
              gap: "4px",
              margin: "0 4px",
              boxShadow: "0 2px 4px rgba(245, 158, 11, 0.15)"
            }}
          >
            👤 Supervisor: {supervisor}
          </span>
        </span>
      );
    }

    return desc;
  };

  const notifRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (notifRef.current && !notifRef.current.contains(e.target)) {
        setNotificationsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  // Load today's live production milestones & unread Google Sheet lot logs on mount
  useEffect(() => {
    let isMounted = true;
    const userName = currentUser?.name || "User";

    const loadNotifs = async () => {
      const milestones = await fetchTodayMilestones();
      const sheetLogs = await fetchLotLogsFromSheet();

      const unreadSheetNotifs = (sheetLogs || [])
        .filter((log) => !isLogReadByUser(log, userName))
        .map((log) => ({
          id: `LOG-${log.id}`,
          title: `📝 Lot #${log.lotNumber} Changed`,
          desc: `${log.changeDetails} | By: ${log.changedBy} (Auth: ${log.permissionBy})`,
          time: log.timeFormatted || "Recently",
          unread: true,
          type: "warning",
          link: "/lot-logs"
        }));

      if (isMounted) {
        setNotifications([...unreadSheetNotifs, ...(milestones || [])]);
      }
    };

    loadNotifs();

    return () => {
      isMounted = false;
    };
  }, []);

  // Listen to Lot Change events and auto-add real-time notifications
  useEffect(() => {
    const userName = currentUser?.name || "User";

    const unsubscribe = subscribeToLotLogs((logItem) => {
      if (isLogReadByUser(logItem, userName)) return;

      const newNotif = {
        id: `LOG-${logItem.id}`,
        title: `📝 Lot #${logItem.lotNumber} Changed`,
        desc: `${logItem.changeDetails} | By: ${logItem.changedBy} (Auth: ${logItem.permissionBy})`,
        time: logItem.timeFormatted || "Just now",
        unread: true,
        type: "warning",
        link: "/lot-logs"
      };
      
      // Play Audio Chime Sound
      playNotificationChime();
      
      // Trigger Flashing Sidebar Toast Popup
      setActiveToast(newNotif);

      setNotifications((prev) => [newNotif, ...prev.filter((n) => n.id !== newNotif.id)]);
    });

    return () => unsubscribe();
  }, []);

  const unreadCount = notifications.filter(n => n.unread).length;

  const markAllRead = () => {
    setNotifications(notifications.map(n => ({ ...n, unread: false })));
  };

  const removeNotification = (id, e) => {
    e.stopPropagation();
    setNotifications(notifications.filter(n => n.id !== id));
  };

  const handleLogout = () => {
    logoutUser();
    history.push("/");
  };

  const isDashboard = location.pathname === "/dashboard";

  return (
    <nav
      style={{
        position: "sticky",
        top: 0,
        zIndex: 9999,
        background: "rgba(255, 255, 255, 0.92)",
        backdropFilter: "blur(16px)",
        WebkitBackdropFilter: "blur(16px)",
        borderBottom: "1px solid #e2e8f0",
        padding: "12px 32px",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        boxShadow: "0 4px 20px rgba(0, 0, 0, 0.04)",
        fontFamily: "'Plus Jakarta Sans', 'Inter', system-ui, sans-serif"
      }}
    >
      {/* Brand & Home Shortcut */}
      <div style={{ display: "flex", alignItems: "center", gap: "20px" }}>
        <div
          onClick={() => history.push("/dashboard")}
          style={{ display: "flex", alignItems: "center", gap: "12px", cursor: "pointer" }}
        >
          <div
            style={{
              width: "42px",
              height: "42px",
              borderRadius: "12px",
              background: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              color: "#ffffff",
              fontWeight: "800",
              fontSize: "1.2rem",
              boxShadow: "0 4px 12px rgba(49, 46, 129, 0.25)"
            }}
          >
            🏭
          </div>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <h2 style={{ margin: 0, fontSize: "1.15rem", fontWeight: "800", color: "#0f172a", letterSpacing: "-0.02em" }}>
                Factory Suite Pro
              </h2>
              <span style={{ background: "#dcfce7", color: "#15803d", fontSize: "0.65rem", fontWeight: "700", padding: "2px 6px", borderRadius: "8px" }}>
                LIVE
              </span>
            </div>
            <p style={{ margin: 0, fontSize: "0.75rem", color: "#64748b", fontWeight: "500" }}>
              Production Command Center
            </p>
          </div>
        </div>

        {!isDashboard && (
          <button
            onClick={() => history.push("/dashboard")}
            style={{
              display: "flex",
              alignItems: "center",
              gap: "6px",
              background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)",
              color: "#ffffff",
              border: "none",
              padding: "8px 16px",
              borderRadius: "10px",
              fontSize: "0.85rem",
              fontWeight: "700",
              cursor: "pointer",
              boxShadow: "0 4px 12px rgba(30, 27, 75, 0.2)"
            }}
          >
            🏠 Dashboard
          </button>
        )}

        <button
          onClick={() => history.push("/lot-logs")}
          style={{
            display: "flex",
            alignItems: "center",
            gap: "6px",
            background: location.pathname === "/lot-logs" ? "#4338ca" : "#f1f5f9",
            color: location.pathname === "/lot-logs" ? "#ffffff" : "#1e293b",
            border: "1px solid #cbd5e1",
            padding: "8px 16px",
            borderRadius: "10px",
            fontSize: "0.85rem",
            fontWeight: "700",
            cursor: "pointer",
            transition: "all 0.2s ease"
          }}
        >
          📋 Lot Logs
        </button>
      </div>

      {/* Right Controls */}
      <div style={{ display: "flex", alignItems: "center", gap: "16px" }}>
        {/* Notification Bell Dropdown */}
        <div ref={notifRef} style={{ position: "relative" }}>
          <button
            onClick={() => setNotificationsOpen(!notificationsOpen)}
            style={{
              position: "relative",
              background: notificationsOpen ? "#e0e7ff" : "#f1f5f9",
              border: "1px solid #cbd5e1",
              width: "40px",
              height: "40px",
              borderRadius: "12px",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "1.1rem",
              cursor: "pointer",
              transition: "all 0.2s ease"
            }}
            title="Notifications Center"
          >
            🔔
            {unreadCount > 0 && (
              <span
                style={{
                  position: "absolute",
                  top: "-4px",
                  right: "-4px",
                  background: "#ef4444",
                  color: "#ffffff",
                  fontSize: "0.7rem",
                  fontWeight: "800",
                  width: "18px",
                  height: "18px",
                  borderRadius: "50%",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  boxShadow: "0 2px 6px rgba(239, 68, 68, 0.4)",
                  border: "2px solid #ffffff"
                }}
              >
                {unreadCount}
              </span>
            )}
          </button>

          <AnimatePresence>
            {notificationsOpen && (
              <motion.div
                initial={{ opacity: 0, y: 15, scale: 0.96 }}
                animate={{ opacity: 1, y: 0, scale: 1 }}
                exit={{ opacity: 0, y: 15, scale: 0.96 }}
                transition={{ duration: 0.2, ease: "easeOut" }}
                style={{
                  position: "absolute",
                  right: 0,
                  top: "54px",
                  width: "390px",
                  maxHeight: "520px",
                  background: "#ffffff",
                  borderRadius: "20px",
                  boxShadow: "0 25px 50px -12px rgba(15, 23, 42, 0.25)",
                  border: "1px solid #e2e8f0",
                  overflow: "hidden",
                  zIndex: 10000,
                  display: "flex",
                  flexDirection: "column"
                }}
              >
                {/* Instagram Style Header */}
                <div
                  style={{
                    padding: "16px 20px",
                    borderBottom: "1px solid #f1f5f9",
                    background: "#ffffff",
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center"
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <h3 style={{ margin: 0, fontSize: "1.1rem", fontWeight: "800", color: "#0f172a", letterSpacing: "-0.02em" }}>
                      Activity
                    </h3>
                    {unreadCount > 0 && (
                      <span
                        style={{
                          background: "linear-gradient(45deg, #f09433, #e6683c, #dc2743, #cc2366, #bc1888)",
                          color: "#ffffff",
                          fontSize: "0.68rem",
                          fontWeight: "800",
                          padding: "2px 8px",
                          borderRadius: "12px",
                          boxShadow: "0 2px 6px rgba(220, 39, 67, 0.3)"
                        }}
                      >
                        {unreadCount} NEW
                      </span>
                    )}
                  </div>
                  {unreadCount > 0 && (
                    <button
                      onClick={markAllRead}
                      style={{
                        background: "#f1f5f9",
                        border: "none",
                        color: "#475569",
                        fontSize: "0.75rem",
                        fontWeight: "700",
                        padding: "6px 12px",
                        borderRadius: "10px",
                        cursor: "pointer",
                        transition: "all 0.2s ease"
                      }}
                    >
                      Mark read
                    </button>
                  )}
                </div>

                {/* Notifications Scroll Area */}
                <div style={{ flex: 1, overflowY: "auto", padding: "8px 0" }}>
                  {notifications.length === 0 ? (
                    <div style={{ padding: "40px 20px", textAlign: "center", color: "#94a3b8" }}>
                      <div style={{ fontSize: "2rem", marginBottom: "8px" }}>🔔</div>
                      <div style={{ fontWeight: "700", fontSize: "0.9rem", color: "#0f172a" }}>No Notifications Yet</div>
                      <div style={{ fontSize: "0.78rem" }}>Lot updates and changes will appear here instantly.</div>
                    </div>
                  ) : (
                    <>
                      {/* Unread / New Section */}
                      {notifications.filter(n => n.unread).length > 0 && (
                        <div>
                          <div style={{ padding: "8px 20px", fontSize: "0.75rem", fontWeight: "800", color: "#64748b", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                            New
                          </div>
                          {notifications.filter(n => n.unread).map((n) => (
                            <div
                              key={n.id}
                              onClick={() => {
                                if (n.link) {
                                  history.push(n.link);
                                  setNotificationsOpen(false);
                                }
                              }}
                              style={{
                                padding: "12px 20px",
                                background: "#f8fafc",
                                display: "flex",
                                alignItems: "center",
                                gap: "14px",
                                cursor: n.link ? "pointer" : "default",
                                transition: "background 0.15s ease",
                                borderBottom: "1px solid #f1f5f9"
                              }}
                            >
                              {/* Gradient Avatar Ring */}
                              <div
                                style={{
                                  width: "44px",
                                  height: "44px",
                                  borderRadius: "50%",
                                  background: "linear-gradient(45deg, #f09433, #e6683c, #dc2743, #cc2366, #bc1888)",
                                  padding: "2px",
                                  flexShrink: 0
                                }}
                              >
                                <div
                                  style={{
                                    width: "100%",
                                    height: "100%",
                                    borderRadius: "50%",
                                    background: "#ffffff",
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "center",
                                    fontSize: "1.1rem"
                                  }}
                                >
                                  {n.type === "warning" ? "📝" : n.type === "success" ? "✅" : "📢"}
                                </div>
                              </div>

                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontSize: "0.85rem", color: "#0f172a", lineHeight: "1.35", wordBreak: "break-word" }}>
                                  <strong style={{ fontWeight: "800" }}>{n.title}</strong>{" "}
                                  <span style={{ color: "#334155" }}>{n.desc}</span>
                                </div>
                                <div style={{ fontSize: "0.72rem", color: "#94a3b8", fontWeight: "600", marginTop: "3px" }}>
                                  {n.time}
                                </div>
                              </div>

                              {/* Instagram Unread Blue Dot */}
                              <div
                                style={{
                                  width: "8px",
                                  height: "8px",
                                  borderRadius: "50%",
                                  background: "#3b82f6",
                                  flexShrink: 0
                                }}
                              />
                            </div>
                          ))}
                        </div>
                      )}

                      {/* Earlier Section */}
                      {notifications.filter(n => !n.unread).length > 0 && (
                        <div>
                          <div style={{ padding: "12px 20px 6px 20px", fontSize: "0.75rem", fontWeight: "800", color: "#64748b", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                            Earlier
                          </div>
                          {notifications.filter(n => !n.unread).map((n) => (
                            <div
                              key={n.id}
                              onClick={() => {
                                if (n.link) {
                                  history.push(n.link);
                                  setNotificationsOpen(false);
                                }
                              }}
                              style={{
                                padding: "12px 20px",
                                background: "#ffffff",
                                display: "flex",
                                alignItems: "center",
                                gap: "14px",
                                cursor: n.link ? "pointer" : "default",
                                transition: "background 0.15s ease",
                                borderBottom: "1px solid #f1f5f9"
                              }}
                            >
                              <div
                                style={{
                                  width: "44px",
                                  height: "44px",
                                  borderRadius: "50%",
                                  background: "#e2e8f0",
                                  padding: "2px",
                                  flexShrink: 0
                                }}
                              >
                                <div
                                  style={{
                                    width: "100%",
                                    height: "100%",
                                    borderRadius: "50%",
                                    background: "#ffffff",
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "center",
                                    fontSize: "1.1rem"
                                  }}
                                >
                                  {n.type === "warning" ? "📝" : n.type === "success" ? "✅" : "ℹ️"}
                                </div>
                              </div>

                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontSize: "0.85rem", color: "#0f172a", lineHeight: "1.4", wordBreak: "break-word" }}>
                                  <strong style={{ fontWeight: "700" }}>{n.title}</strong>{" "}
                                  <span style={{ color: "#64748b" }}>{renderFormattedDesc(n.desc, n.supervisor)}</span>
                                </div>
                                <div style={{ fontSize: "0.72rem", color: "#94a3b8", fontWeight: "500", marginTop: "3px" }}>
                                  {n.time}
                                </div>
                              </div>

                              <button
                                onClick={(e) => removeNotification(n.id, e)}
                                style={{ background: "none", border: "none", color: "#cbd5e1", cursor: "pointer", fontSize: "0.8rem", padding: "4px" }}
                              >
                                ✕
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </>
                  )}
                </div>

                {/* Footer Link */}
                <div
                  onClick={() => {
                    history.push("/notifications");
                    setNotificationsOpen(false);
                  }}
                  style={{
                    padding: "12px",
                    background: "#f8fafc",
                    borderTop: "1px solid #f1f5f9",
                    textAlign: "center",
                    fontSize: "0.8rem",
                    fontWeight: "800",
                    color: "#4f46e5",
                    cursor: "pointer"
                  }}
                >
                  View All Activity & Notifications Screen →
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* User Profile */}
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: "12px",
            background: "#ffffff",
            padding: "5px 14px",
            borderRadius: "14px",
            border: "1px solid #cbd5e1"
          }}
        >
          <div
            style={{
              width: "36px",
              height: "36px",
              borderRadius: "50%",
              background: currentUser?.avatarColor || "#4f46e5",
              color: "#ffffff",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontWeight: "700",
              fontSize: "0.9rem"
            }}
          >
            {currentUser?.name ? currentUser.name.split(" ").map(n => n[0]).join("").slice(0, 2).toUpperCase() : "U"}
          </div>
          <div>
            <div style={{ fontSize: "0.85rem", fontWeight: "700", color: "#0f172a" }}>
              {currentUser?.name || "User"}
            </div>
            <span
              style={{
                fontSize: "0.65rem",
                fontWeight: "700",
                color: "#4f46e5",
                background: "#e0e7ff",
                padding: "1px 6px",
                borderRadius: "6px"
              }}
            >
              {currentUser?.role || "Member"}
            </span>
          </div>
        </div>

        {/* Logout Button */}
        <button
          onClick={handleLogout}
          style={{
            background: "#f1f5f9",
            color: "#475569",
            border: "1px solid #cbd5e1",
            padding: "8px 14px",
            borderRadius: "10px",
            fontSize: "0.85rem",
            fontWeight: "700",
            cursor: "pointer",
            display: "flex",
            alignItems: "center",
            gap: "6px"
          }}
        >
          🚪 Logout
        </button>
      </div>

      {/* Flashing Sidebar Toast Notification Popup */}
      <AnimatePresence>
        {activeToast && (
          <motion.div
            initial={{ opacity: 0, x: 100, scale: 0.9 }}
            animate={{ opacity: 1, x: 0, scale: 1 }}
            exit={{ opacity: 0, x: 100, scale: 0.9 }}
            transition={{ type: "spring", stiffness: 400, damping: 25 }}
            style={{
              position: "fixed",
              bottom: "24px",
              right: "24px",
              width: "360px",
              background: "rgba(15, 23, 42, 0.92)",
              backdropFilter: "blur(16px)",
              WebkitBackdropFilter: "blur(16px)",
              border: "1px solid rgba(99, 102, 241, 0.5)",
              borderRadius: "20px",
              padding: "16px 20px",
              color: "#ffffff",
              boxShadow: "0 20px 40px rgba(0, 0, 0, 0.3), 0 0 20px rgba(99, 102, 241, 0.4)",
              zIndex: 999999,
              display: "flex",
              alignItems: "center",
              gap: "14px"
            }}
          >
            {/* Story Ring Avatar Icon */}
            <div
              style={{
                width: "42px",
                height: "42px",
                borderRadius: "50%",
                background: "linear-gradient(45deg, #f09433, #e6683c, #dc2743, #cc2366, #bc1888)",
                padding: "2px",
                flexShrink: 0,
                boxShadow: "0 0 12px rgba(220, 39, 67, 0.6)"
              }}
            >
              <div
                style={{
                  width: "100%",
                  height: "100%",
                  borderRadius: "50%",
                  background: "#0f172a",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: "1.1rem"
                }}
              >
                🔔
              </div>
            </div>

            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: "0.85rem", fontWeight: "800", color: "#ffffff", marginBottom: "2px" }}>
                {activeToast.title}
              </div>
              <div style={{ fontSize: "0.78rem", color: "#cbd5e1", lineHeight: "1.3", overflow: "hidden", textOverflow: "ellipsis", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical" }}>
                {activeToast.desc}
              </div>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: "6px", alignItems: "center" }}>
              <button
                onClick={() => {
                  if (activeToast.link) history.push(activeToast.link);
                  setActiveToast(null);
                }}
                style={{
                  background: "linear-gradient(135deg, #6366f1 0%, #4338ca 100%)",
                  color: "#ffffff",
                  border: "none",
                  padding: "4px 10px",
                  borderRadius: "8px",
                  fontSize: "0.72rem",
                  fontWeight: "800",
                  cursor: "pointer"
                }}
              >
                View
              </button>
              <button
                onClick={() => setActiveToast(null)}
                style={{ background: "none", border: "none", color: "#94a3b8", fontSize: "0.75rem", cursor: "pointer" }}
              >
                ✕
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Global Center-of-Screen Notification Modal & 30-Min Reminder */}
      <CenterNotificationModal />
    </nav>
  );
};

export default Navbar;
