// src/NotificationCenter.js
import React, { useState, useEffect, useMemo } from "react";
import { subscribeToLotLogs, getLotLogs, fetchLotLogsFromSheet } from "./lotLogService";
import { fetchTodayMilestones } from "./productionMilestones";
import { useHistory } from "react-router-dom";

export default function NotificationCenter() {
  const history = useHistory();
  const [loading, setLoading] = useState(true);
  const [notifications, setNotifications] = useState([]);
  const [activeTab, setActiveTab] = useState("all");
  const [searchTerm, setSearchTerm] = useState("");

  const loadNotifications = async () => {
    setLoading(true);
    try {
      // 1. Load today's live production milestones
      const milestones = await fetchTodayMilestones();

      // 2. Load live lot change audit logs from Google Sheet
      const liveAuditLogs = await fetchLotLogsFromSheet();
      const auditLogs = liveAuditLogs.length > 0 ? liveAuditLogs : getLotLogs();

      const formattedAuditNotifs = auditLogs.map((logItem) => ({
        id: `LOG-${logItem.id}`,
        title: `📝 Lot #${logItem.lotNumber} Changed`,
        desc: `${logItem.changeDetails} | Changed by: ${logItem.changedBy} (Auth: ${logItem.permissionBy})`,
        time: logItem.timeFormatted || logItem.timestamp || "Recently",
        unread: true,
        type: "warning",
        link: "/lot-logs",
        category: "Lot Audit Logs",
        rawTimestamp: new Date(logItem.timestamp || Date.now()).getTime()
      }));

      // Combine and deduplicate
      const combined = [...formattedAuditNotifs, ...(milestones || [])];

      // Sort by newest first
      combined.sort((a, b) => (b.rawTimestamp || Date.now()) - (a.rawTimestamp || Date.now()));

      setNotifications(combined);
    } catch (err) {
      console.error("Error loading notification center data:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadNotifications();

    // Listen for incoming real-time lot logs
    const unsubscribe = subscribeToLotLogs((newLog) => {
      const formattedNotif = {
        id: `LOG-${newLog.id}`,
        title: `📝 Lot #${newLog.lotNumber} Changed`,
        desc: `${newLog.changeDetails} | Changed by: ${newLog.changedBy} (Auth: ${newLog.permissionBy})`,
        time: newLog.timeFormatted || "Just now",
        unread: true,
        type: "warning",
        link: "/lot-logs",
        category: "Lot Audit Logs",
        rawTimestamp: Date.now()
      };
      setNotifications((prev) => [formattedNotif, ...prev]);
    });

    return () => unsubscribe();
  }, []);

  // Filtered Notifications
  const filteredNotifications = useMemo(() => {
    return notifications.filter((item) => {
      // Search filter
      const matchesSearch =
        !searchTerm ||
        item.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
        item.desc.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (item.category && item.category.toLowerCase().includes(searchTerm.toLowerCase()));

      if (!matchesSearch) return false;

      // Tab filter
      if (activeTab === "all") return true;
      if (activeTab === "unread") return item.unread;
      if (activeTab === "logs") return item.category === "Lot Audit Logs";
      if (activeTab === "stitching_issue") return item.category === "Stitching Issue";
      if (activeTab === "stitching") return item.category === "Stitching";
      if (activeTab === "cutting") return item.category === "Cutting";
      if (activeTab === "embroidery") return item.category === "Embroidery";
      if (activeTab === "printing") return item.category === "Printing";

      return true;
    });
  }, [notifications, activeTab, searchTerm]);

  // Counts
  const counts = useMemo(() => {
    return {
      all: notifications.length,
      unread: notifications.filter((n) => n.unread).length,
      logs: notifications.filter((n) => n.category === "Lot Audit Logs").length,
      stitching_issue: notifications.filter((n) => n.category === "Stitching Issue").length,
      stitching: notifications.filter((n) => n.category === "Stitching").length,
      cutting: notifications.filter((n) => n.category === "Cutting").length,
      embroidery: notifications.filter((n) => n.category === "Embroidery").length,
      printing: notifications.filter((n) => n.category === "Printing").length
    };
  }, [notifications]);

  const markAllAsRead = () => {
    setNotifications((prev) => prev.map((n) => ({ ...n, unread: false })));
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

  return (
    <div style={{ minHeight: "100vh", backgroundColor: "#f8fafc", fontFamily: "'Plus Jakarta Sans', 'Inter', sans-serif" }}>
      <div style={{ maxWidth: "100%", margin: "0 auto", padding: "24px 32px" }}>
        {/* Instagram Header Banner */}
        <div
          style={{
            background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%)",
            borderRadius: "24px",
            padding: "32px 36px",
            color: "#ffffff",
            boxShadow: "0 20px 40px -15px rgba(30, 27, 75, 0.25)",
            marginBottom: "28px",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            flexWrap: "wrap",
            gap: "20px"
          }}
        >
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
              <div
                style={{
                  width: "48px",
                  height: "48px",
                  borderRadius: "50%",
                  background: "linear-gradient(45deg, #f09433, #e6683c, #dc2743, #cc2366, #bc1888)",
                  padding: "2px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center"
                }}
              >
                <div style={{ width: "100%", height: "100%", borderRadius: "50%", background: "#1e1b4b", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "1.3rem" }}>
                  🔔
                </div>
              </div>
              <div>
                <h1 style={{ margin: 0, fontSize: "1.8rem", fontWeight: 800, letterSpacing: "-0.02em" }}>
                  Activity & Notifications Center
                </h1>
                <p style={{ margin: "4px 0 0 0", color: "#c7d2fe", fontSize: "0.92rem", fontWeight: 500 }}>
                  Real-time factory lot changes, stitching, cutting & challans activity feed
                </p>
              </div>
            </div>
          </div>

          <div style={{ display: "flex", gap: "12px", alignItems: "center" }}>
            <button
              onClick={() => {
                if (history.length > 1) {
                  history.goBack();
                } else {
                  history.push("/dashboard");
                }
              }}
              style={{
                background: "rgba(255, 255, 255, 0.2)",
                backdropFilter: "blur(10px)",
                border: "1px solid rgba(255, 255, 255, 0.3)",
                color: "#ffffff",
                padding: "10px 18px",
                borderRadius: "12px",
                fontSize: "0.85rem",
                fontWeight: 700,
                cursor: "pointer",
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                transition: "all 0.2s ease"
              }}
            >
              ⬅️ Back
            </button>

            <button
              onClick={markAllAsRead}
              style={{
                background: "rgba(255, 255, 255, 0.15)",
                backdropFilter: "blur(10px)",
                border: "1px solid rgba(255, 255, 255, 0.2)",
                color: "#ffffff",
                padding: "10px 18px",
                borderRadius: "12px",
                fontSize: "0.85rem",
                fontWeight: 700,
                cursor: "pointer",
                transition: "all 0.2s ease"
              }}
            >
              ✓ Mark All Read
            </button>
            <button
              onClick={loadAllData => loadNotifications()}
              style={{
                background: "linear-gradient(135deg, #ec4899 0%, #be185d 100%)",
                border: "none",
                color: "#ffffff",
                padding: "10px 18px",
                borderRadius: "12px",
                fontSize: "0.85rem",
                fontWeight: 700,
                cursor: "pointer",
                boxShadow: "0 4px 12px rgba(236, 72, 153, 0.3)"
              }}
            >
              🔄 Refresh Feed
            </button>
          </div>
        </div>

        {/* Toolbar: Search + Filter Tabs */}
        <div style={{ background: "#ffffff", borderRadius: "20px", padding: "20px 24px", marginBottom: "24px", boxShadow: "0 10px 30px rgba(0,0,0,0.03)", border: "1px solid #e2e8f0" }}>
          {/* Search Box */}
          <div style={{ position: "relative", marginBottom: "16px" }}>
            <span style={{ position: "absolute", left: "14px", top: "50%", transform: "translateY(-50%)", color: "#64748b" }}>🔍</span>
            <input
              type="text"
              placeholder="Search notifications by lot number, supervisor, party, or details..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              style={{
                width: "100%",
                padding: "12px 14px 12px 42px",
                borderRadius: "12px",
                border: "1.5px solid #cbd5e1",
                fontSize: "0.9rem",
                fontWeight: 600,
                outline: "none",
                boxSizing: "border-box",
                color: "#0f172a"
              }}
            />
          </div>

          {/* Instagram Filter Pills */}
          <div style={{ display: "flex", gap: "8px", overflowX: "auto", paddingBottom: "4px" }}>
            {[
              { id: "all", label: "🌐 All Feed", count: counts.all },
              { id: "unread", label: "🔵 Unread", count: counts.unread },
              { id: "logs", label: "📝 Lot Logs", count: counts.logs },
              { id: "stitching_issue", label: "📍 Stitching Issued", count: counts.stitching_issue },
              { id: "stitching", label: "✨ Stitching Complete", count: counts.stitching },
              { id: "cutting", label: "✂️ Cutting", count: counts.cutting },
              { id: "embroidery", label: "📈 Embroidery", count: counts.embroidery },
              { id: "printing", label: "🖨️ Printing", count: counts.printing }
            ].map((tab) => {
              const active = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  style={{
                    padding: "8px 14px",
                    borderRadius: "20px",
                    border: active ? "none" : "1px solid #cbd5e1",
                    background: active ? "linear-gradient(135deg, #6366f1 0%, #4338ca 100%)" : "#f8fafc",
                    color: active ? "#ffffff" : "#475569",
                    fontSize: "0.82rem",
                    fontWeight: 700,
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                    display: "flex",
                    alignItems: "center",
                    gap: "6px",
                    boxShadow: active ? "0 4px 12px rgba(99, 102, 241, 0.25)" : "none",
                    transition: "all 0.2s ease"
                  }}
                >
                  {tab.label}
                  <span
                    style={{
                      background: active ? "rgba(255, 255, 255, 0.25)" : "#e2e8f0",
                      color: active ? "#ffffff" : "#475569",
                      padding: "2px 7px",
                      borderRadius: "10px",
                      fontSize: "0.72rem",
                      fontWeight: 800
                    }}
                  >
                    {tab.count}
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Notifications List Feed */}
        <div style={{ background: "#ffffff", borderRadius: "24px", border: "1px solid #e2e8f0", boxShadow: "0 10px 30px rgba(0,0,0,0.03)", overflow: "hidden" }}>
          {loading ? (
            <div style={{ padding: "60px", textAlign: "center", color: "#64748b" }}>
              <div style={{ fontSize: "2rem", marginBottom: "12px", animation: "spin 1s linear infinite" }}>⏳</div>
              <div style={{ fontWeight: 700, fontSize: "1rem", color: "#0f172a" }}>Fetching Real-Time Notifications...</div>
            </div>
          ) : filteredNotifications.length === 0 ? (
            <div style={{ padding: "60px", textAlign: "center", color: "#94a3b8" }}>
              <div style={{ fontSize: "3rem", marginBottom: "12px" }}>🔔</div>
              <div style={{ fontWeight: 800, fontSize: "1.1rem", color: "#0f172a" }}>No Notifications Found</div>
              <div style={{ fontSize: "0.85rem", marginTop: "4px" }}>There are no active notifications matching your selected filter.</div>
            </div>
          ) : (
            filteredNotifications.map((notif, idx) => (
              <div
                key={notif.id || idx}
                style={{
                  padding: "20px 24px",
                  borderBottom: idx === filteredNotifications.length - 1 ? "none" : "1px solid #f1f5f9",
                  background: notif.unread ? "#f8fafc" : "#ffffff",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: "16px",
                  transition: "background 0.15s ease"
                }}
              >
                {/* Left: Avatar ring + text */}
                <div style={{ display: "flex", alignItems: "center", gap: "16px", flex: 1, minWidth: 0 }}>
                  <div
                    style={{
                      width: "50px",
                      height: "50px",
                      borderRadius: "50%",
                      background: "linear-gradient(45deg, #f09433, #e6683c, #dc2743, #cc2366, #bc1888)",
                      padding: "2.5px",
                      flexShrink: 0,
                      boxShadow: notif.unread ? "0 0 12px rgba(220, 39, 67, 0.3)" : "none"
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
                        fontSize: "1.2rem"
                      }}
                    >
                      {notif.category === "Lot Audit Logs" ? "📝" :
                        notif.category === "Stitching Issue" ? "📍" :
                          notif.category === "Stitching" ? "✨" :
                            notif.category === "Cutting" ? "✂️" :
                              notif.category === "Embroidery" ? "📈" : "🖨️"}
                    </div>
                  </div>

                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap", marginBottom: "4px" }}>
                      <span style={{ fontSize: "0.95rem", fontWeight: 800, color: "#0f172a" }}>
                        {notif.title}
                      </span>
                      {notif.category && (
                        <span
                          style={{
                            fontSize: "0.68rem",
                            fontWeight: 800,
                            padding: "2px 8px",
                            borderRadius: "12px",
                            background: "#e0e7ff",
                            color: "#4338ca",
                            textTransform: "uppercase",
                            letterSpacing: "0.04em"
                          }}
                        >
                          {notif.category}
                        </span>
                      )}
                      {notif.unread && (
                        <span style={{ color: "#3b82f6", fontSize: "1rem", lineHeight: 1 }}>•</span>
                      )}
                    </div>

                    <div style={{ fontSize: "0.85rem", color: "#475569", lineHeight: 1.5, wordBreak: "break-word" }}>
                      {renderFormattedDesc(notif.desc, notif.supervisor)}
                    </div>

                    <div style={{ fontSize: "0.75rem", color: "#94a3b8", fontWeight: 600, marginTop: "4px" }}>
                      🕒 {notif.time}
                    </div>
                  </div>
                </div>

                {/* Right: View Button */}
                {notif.link && (
                  <button
                    onClick={() => history.push(notif.link)}
                    style={{
                      background: "linear-gradient(135deg, #6366f1 0%, #4338ca 100%)",
                      color: "#ffffff",
                      border: "none",
                      padding: "10px 18px",
                      borderRadius: "12px",
                      fontSize: "0.82rem",
                      fontWeight: 800,
                      cursor: "pointer",
                      whiteSpace: "nowrap",
                      boxShadow: "0 4px 12px rgba(99, 102, 241, 0.25)",
                      transition: "transform 0.15s ease"
                    }}
                  >
                    View Details →
                  </button>
                )}
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
}
