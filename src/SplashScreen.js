// src/SplashScreen.js
import React, { useEffect, useState } from "react";
import { motion, AnimatePresence } from "framer-motion";

const SplashScreen = ({ onFinish }) => {
  const [progress, setProgress] = useState(0);
  const [isComplete, setIsComplete] = useState(false);
  const [currentTime, setCurrentTime] = useState("");

  const systemChecklist = [
    { label: "Google Sheets Engine API", detail: "CONNECTED" },
    { label: "RAWPACK & Issues Synchronization", detail: "SYNCHRONIZED" },
    { label: "GARMENT IMAGE PROXY PIPELINE", detail: "ACTIVE" },
    { label: "SECURE USER PERMISSIONS & AUDIT LOGS", detail: "AUTHENTICATED" },
    { label: "PRODUCTION DASHBOARD & DAILY REPORTS", detail: "READY" }
  ];

  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      setCurrentTime(now.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit", second: "2-digit" }));
    };
    updateTime();
    const timer = setInterval(updateTime, 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    const interval = setInterval(() => {
      setProgress(prev => {
        if (prev >= 100) {
          clearInterval(interval);
          setIsComplete(true);
          setTimeout(() => {
            if (onFinish) onFinish();
          }, 1200);
          return 100;
        }

        let increment;
        if (prev < 65) {
          increment = 1.2 + Math.random() * 0.8;
        } else if (prev < 90) {
          increment = 0.6 + Math.random() * 0.4;
        } else {
          increment = 0.2 + Math.random() * 0.1;
        }
        return Math.min(prev + increment, 100);
      });
    }, 25);

    return () => clearInterval(interval);
  }, [onFinish]);

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Outfit:wght@400;600;700;800&display=swap');

        .splash-root {
          position: fixed;
          inset: 0;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #090d16;
          color: #f8fafc;
          font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
          overflow: hidden;
          z-index: 99999;
        }

        .splash-ambient-bg {
          position: absolute;
          inset: 0;
          background: 
            radial-gradient(circle at 15% 20%, rgba(99, 102, 241, 0.18) 0%, transparent 45%),
            radial-gradient(circle at 85% 75%, rgba(6, 182, 212, 0.15) 0%, transparent 45%),
            radial-gradient(circle at 50% 50%, rgba(168, 85, 247, 0.1) 0%, transparent 50%);
          filter: blur(40px);
          animation: pulseGlow 12s ease-in-out infinite alternate;
        }

        .grid-pattern {
          position: absolute;
          inset: 0;
          background-image: 
            linear-gradient(to right, rgba(255, 255, 255, 0.03) 1px, transparent 1px),
            linear-gradient(to bottom, rgba(255, 255, 255, 0.03) 1px, transparent 1px);
          background-size: 40px 40px;
          opacity: 0.6;
        }

        .splash-card {
          position: relative;
          z-index: 10;
          width: 90%;
          max-width: 580px;
          background: rgba(15, 23, 42, 0.75);
          backdrop-filter: blur(20px);
          -webkit-backdrop-filter: blur(20px);
          border: 1px solid rgba(255, 255, 255, 0.12);
          border-radius: 28px;
          padding: 40px;
          box-shadow: 0 25px 60px -15px rgba(0, 0, 0, 0.6), 0 0 30px rgba(99, 102, 241, 0.15);
          text-align: center;
        }

        .top-badge-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          margin-bottom: 28px;
        }

        .status-pill {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          background: rgba(16, 185, 129, 0.12);
          border: 1px solid rgba(16, 185, 129, 0.3);
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 11px;
          font-weight: 700;
          letter-spacing: 1px;
          color: #34d399;
          text-transform: uppercase;
        }

        .pulse-dot {
          width: 8px;
          height: 8px;
          background-color: #10b981;
          border-radius: 50%;
          box-shadow: 0 0 8px #10b981;
          animation: dotBlink 1.5s ease-in-out infinite;
        }

        .live-clock {
          font-size: 12px;
          font-weight: 600;
          color: #94a3b8;
          font-family: monospace;
          letter-spacing: 0.5px;
        }

        .brand-header {
          margin-bottom: 32px;
        }

        .brand-icon-wrapper {
          width: 72px;
          height: 72px;
          margin: 0 auto 20px;
          background: linear-gradient(135deg, #6366f1 0%, #a855f7 50%, #06b6d4 100%);
          border-radius: 22px;
          display: flex;
          align-items: center;
          justify-content: center;
          box-shadow: 0 12px 30px rgba(99, 102, 241, 0.4);
        }

        .brand-title {
          font-family: 'Outfit', sans-serif;
          font-size: 34px;
          font-weight: 800;
          letter-spacing: -0.5px;
          background: linear-gradient(135deg, #ffffff 0%, #cbd5e1 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
          margin: 0 0 6px 0;
        }

        .brand-subtitle {
          font-size: 12px;
          font-weight: 700;
          letter-spacing: 2px;
          color: #818cf8;
          text-transform: uppercase;
          margin: 0;
        }

        .progress-box {
          margin-bottom: 28px;
          background: rgba(30, 41, 59, 0.5);
          border: 1px solid rgba(255, 255, 255, 0.08);
          border-radius: 18px;
          padding: 20px;
        }

        .progress-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 12px;
        }

        .progress-text {
          font-size: 13px;
          font-weight: 600;
          color: #cbd5e1;
        }

        .progress-num {
          font-size: 20px;
          font-weight: 800;
          color: #6366f1;
          font-family: 'Outfit', sans-serif;
        }

        .track-bar {
          width: 100%;
          height: 8px;
          background: rgba(255, 255, 255, 0.08);
          border-radius: 10px;
          overflow: hidden;
          position: relative;
        }

        .fill-bar {
          height: 100%;
          background: linear-gradient(90deg, #6366f1 0%, #a855f7 50%, #06b6d4 100%);
          border-radius: 10px;
          box-shadow: 0 0 15px rgba(99, 102, 241, 0.8);
          transition: width 0.1s ease-out;
        }

        .checklist-wrapper {
          text-align: left;
          margin-bottom: 24px;
          display: flex;
          flex-direction: column;
          gap: 10px;
        }

        .checklist-item {
          display: flex;
          align-items: center;
          justify-content: space-between;
          background: rgba(30, 41, 59, 0.3);
          border: 1px solid rgba(255, 255, 255, 0.05);
          padding: 10px 14px;
          border-radius: 12px;
          font-size: 12px;
          transition: all 0.3s ease;
        }

        .checklist-item.done {
          background: rgba(99, 102, 241, 0.08);
          border-color: rgba(99, 102, 241, 0.25);
        }

        .item-label {
          color: #94a3b8;
          font-weight: 500;
        }

        .checklist-item.done .item-label {
          color: #f1f5f9;
          font-weight: 600;
        }

        .item-status {
          font-size: 10px;
          font-weight: 800;
          letter-spacing: 1px;
          padding: 3px 8px;
          border-radius: 6px;
          background: rgba(255, 255, 255, 0.05);
          color: #64748b;
        }

        .checklist-item.done .item-status {
          background: rgba(16, 185, 129, 0.15);
          color: #34d399;
        }

        .launch-btn {
          width: 100%;
          padding: 14px;
          border-radius: 14px;
          border: none;
          background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
          color: #ffffff;
          font-size: 14px;
          font-weight: 700;
          letter-spacing: 1px;
          text-transform: uppercase;
          cursor: pointer;
          box-shadow: 0 10px 25px rgba(79, 70, 229, 0.4);
          transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        .launch-btn:hover {
          transform: translateY(-2px);
          box-shadow: 0 14px 30px rgba(79, 70, 229, 0.5);
        }

        .footer-note {
          margin-top: 20px;
          font-size: 11px;
          color: #64748b;
          letter-spacing: 0.5px;
        }

        @keyframes pulseGlow {
          0% { transform: scale(1); opacity: 0.8; }
          100% { transform: scale(1.1); opacity: 1; }
        }

        @keyframes dotBlink {
          0%, 100% { opacity: 0.4; transform: scale(0.9); }
          50% { opacity: 1; transform: scale(1.2); }
        }
      `}</style>

      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0, transition: { duration: 0.4 } }}
        className="splash-root"
      >
        <div className="splash-ambient-bg" />
        <div className="grid-pattern" />

        <motion.div
          className="splash-card"
          initial={{ scale: 0.9, y: 20, opacity: 0 }}
          animate={{ scale: 1, y: 0, opacity: 1 }}
          transition={{ duration: 0.5, ease: "easeOut" }}
        >
          {/* Top Status Bar */}
          <div className="top-badge-row">
            <div className="status-pill">
              <span className="pulse-dot" />
              SYSTEM ONLINE • VER 2026.8
            </div>
            <div className="live-clock">{currentTime}</div>
          </div>

          {/* Brand Header */}
          <div className="brand-header">
            <div className="brand-icon-wrapper">
              <svg width="42" height="42" viewBox="0 0 512 512" fill="none">
                <defs>
                  <linearGradient id="splashLogoGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stopColor="#ffffff"/>
                    <stop offset="100%" stopColor="#e0e7ff"/>
                  </linearGradient>
                  <linearGradient id="splashBoltGrad" x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stopColor="#fbbf24"/>
                    <stop offset="100%" stopColor="#f59e0b"/>
                  </linearGradient>
                </defs>
                <path d="M 120 370 L 120 150 Q 120 140 130 140 L 155 140 Q 165 140 170 150 L 220 250 L 220 150 Q 220 140 230 140 L 250 140 Q 260 140 260 150 L 260 370 Q 260 380 245 380 L 225 380 Q 215 380 210 370 L 160 270 L 160 370 Q 160 380 145 380 L 135 380 Q 120 380 120 370 Z" fill="url(#splashLogoGrad)"/>
                <path d="M 290 150 Q 290 140 305 140 L 325 140 Q 340 140 340 150 L 340 230 L 390 230 L 390 150 Q 390 140 405 140 L 425 140 Q 440 140 440 150 L 440 370 Q 440 380 425 380 L 405 380 Q 390 380 390 370 L 390 280 L 340 280 L 340 370 Q 340 380 325 380 L 305 380 Q 290 380 290 370 Z" fill="url(#splashLogoGrad)"/>
                <polygon points="275,120 235,240 270,240 235,360 315,220 275,220" fill="url(#splashBoltGrad)"/>
              </svg>
            </div>
            <h1 className="brand-title">MH FACTORY SUITE PRO</h1>
            <p className="brand-subtitle">ENTERPRISE GARMENT PRODUCTION & LIVE REPORTING PORTAL</p>
          </div>

          {/* Progress Box */}
          <div className="progress-box">
            <div className="progress-row">
              <span className="progress-text">
                {isComplete ? "✨ All Systems Initialized Successfully" : "⚡ Synchronizing Live Sheets Data..."}
              </span>
              <span className="progress-num">{Math.round(progress)}%</span>
            </div>
            <div className="track-bar">
              <div className="fill-bar" style={{ width: `${progress}%` }} />
            </div>
          </div>

          {/* System Verification Checklist */}
          <div className="checklist-wrapper">
            {systemChecklist.map((item, index) => {
              const isItemDone = progress >= (index + 1) * 20;
              return (
                <div key={index} className={`checklist-item ${isItemDone ? "done" : ""}`}>
                  <span className="item-label">{item.label}</span>
                  <span className="item-status">
                    {isItemDone ? `✓ ${item.detail}` : "WAITING..."}
                  </span>
                </div>
              );
            })}
          </div>

          {/* Launch Button when Complete */}
          <AnimatePresence>
            {isComplete && (
              <motion.button
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0 }}
                onClick={() => onFinish && onFinish()}
                className="launch-btn"
              >
                🚀 ENTER PRODUCTION DASHBOARD
              </motion.button>
            )}
          </AnimatePresence>

          <p className="footer-note">MH ENGINEER • ADVANCED DAILY REPORTING ENGINE</p>
        </motion.div>
      </motion.div>
    </>
  );
};

export default SplashScreen;