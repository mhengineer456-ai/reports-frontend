import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { useHistory } from "react-router-dom";
import { loginUserAsync } from "./auth";

const LoginScreen = () => {
  const [id, setId] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [currentTime, setCurrentTime] = useState("");

  const history = useHistory();

  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      setCurrentTime(
        now.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit", second: "2-digit" })
      );
    };
    updateTime();
    const timer = setInterval(updateTime, 1000);
    return () => clearInterval(timer);
  }, []);

  const handleLogin = async (e) => {
    if (e) e.preventDefault();

    if (!id || !password) {
      setError("Please enter both User ID and Password.");
      return;
    }

    setIsLoading(true);
    setError("");

    try {
      const sessionUser = await loginUserAsync(id, password);

      if (sessionUser) {
        history.push("/dashboard");
      } else {
        setError("Invalid User ID or Password.");
        setIsLoading(false);
      }
    } catch (err) {
      console.error("Login authentication error:", err);
      setError("Authentication failed. Please check your connection.");
      setIsLoading(false);
    }
  };

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Outfit:wght@500;600;700;800&display=swap');

        .lgn-root {
          position: fixed;
          inset: 0;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #090d16;
          color: #f8fafc;
          font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
          overflow: hidden;
          padding: 1rem;
        }

        .lgn-ambient-bg {
          position: absolute;
          inset: 0;
          background: 
            radial-gradient(circle at 15% 20%, rgba(99, 102, 241, 0.22) 0%, transparent 45%),
            radial-gradient(circle at 85% 75%, rgba(6, 182, 212, 0.18) 0%, transparent 45%),
            radial-gradient(circle at 50% 50%, rgba(168, 85, 247, 0.12) 0%, transparent 50%);
          filter: blur(40px);
          animation: lgnPulse 12s ease-in-out infinite alternate;
        }

        .lgn-grid-pattern {
          position: absolute;
          inset: 0;
          background-image: 
            linear-gradient(to right, rgba(255, 255, 255, 0.03) 1px, transparent 1px),
            linear-gradient(to bottom, rgba(255, 255, 255, 0.03) 1px, transparent 1px);
          background-size: 40px 40px;
          opacity: 0.6;
        }

        @keyframes lgnPulse {
          0% { opacity: 0.7; transform: scale(1); }
          100% { opacity: 1; transform: scale(1.05); }
        }

        .lgn-card {
          position: relative;
          z-index: 10;
          width: 100%;
          max-width: 440px;
          background: rgba(15, 23, 42, 0.78);
          backdrop-filter: blur(24px);
          -webkit-backdrop-filter: blur(24px);
          border: 1px solid rgba(255, 255, 255, 0.12);
          border-radius: 24px;
          padding: 2.5rem 2rem;
          box-shadow: 
            0 25px 50px -12px rgba(0, 0, 0, 0.6),
            0 0 0 1px rgba(255, 255, 255, 0.05);
        }

        .lgn-header {
          text-align: center;
          margin-bottom: 2.25rem;
        }

        .lgn-logo-badge {
          width: 60px;
          height: 60px;
          background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 50%, #06b6d4 100%);
          border-radius: 18px;
          margin: 0 auto 1.25rem;
          display: grid;
          place-items: center;
          color: #ffffff;
          font-family: 'Outfit', sans-serif;
          font-weight: 800;
          font-size: 1.65rem;
          box-shadow: 0 10px 25px -5px rgba(79, 70, 229, 0.5);
          border: 1px solid rgba(255, 255, 255, 0.25);
        }

        .lgn-title {
          font-family: 'Outfit', sans-serif;
          font-size: 1.85rem;
          font-weight: 800;
          color: #ffffff;
          margin: 0 0 0.35rem 0;
          letter-spacing: -0.02em;
          background: linear-gradient(to right, #ffffff, #cbd5e1);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
        }

        .lgn-subtitle {
          color: #94a3b8;
          font-size: 0.9rem;
          margin: 0;
          font-weight: 500;
        }

        .lgn-status-pill {
          display: inline-flex;
          align-items: center;
          gap: 0.4rem;
          background: rgba(16, 185, 129, 0.12);
          color: #34d399;
          border: 1px solid rgba(16, 185, 129, 0.25);
          padding: 0.25rem 0.75rem;
          border-radius: 999px;
          font-size: 0.75rem;
          font-weight: 700;
          margin-top: 0.85rem;
          letter-spacing: 0.02em;
        }

        .lgn-field-group {
          margin-bottom: 1.35rem;
        }

        .lgn-label {
          display: block;
          font-size: 0.8rem;
          font-weight: 700;
          color: #cbd5e1;
          margin-bottom: 0.4rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }

        .lgn-input-wrapper {
          position: relative;
          display: flex;
          align-items: center;
        }

        .lgn-input-icon {
          position: absolute;
          left: 1rem;
          color: #64748b;
          font-size: 1.1rem;
          pointer-events: none;
          transition: color 0.2s ease;
        }

        .lgn-input {
          width: 100%;
          padding: 0.9rem 1rem 0.9rem 2.8rem;
          font-size: 0.95rem;
          background: rgba(30, 41, 59, 0.6);
          border: 1.5px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          outline: none;
          color: #ffffff;
          font-weight: 500;
          transition: all 0.2s ease;
          box-sizing: border-box;
        }

        .lgn-input::placeholder {
          color: #64748b;
        }

        .lgn-input:focus {
          background: rgba(30, 41, 59, 0.9);
          border-color: #6366f1;
          box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.25);
        }

        .lgn-input:focus + .lgn-input-icon,
        .lgn-input-wrapper:focus-within .lgn-input-icon {
          color: #818cf8;
        }

        .lgn-toggle-pass {
          position: absolute;
          right: 0.85rem;
          background: transparent;
          border: 0;
          color: #64748b;
          cursor: pointer;
          font-size: 1.1rem;
          padding: 0.25rem;
          display: flex;
          align-items: center;
          justify-content: center;
          transition: color 0.2s ease;
        }

        .lgn-toggle-pass:hover {
          color: #cbd5e1;
        }

        .lgn-error-box {
          background: rgba(239, 68, 68, 0.12);
          border: 1px solid rgba(239, 68, 68, 0.3);
          border-radius: 10px;
          padding: 0.7rem 0.85rem;
          margin-bottom: 1.25rem;
          color: #fca5a5;
          font-size: 0.85rem;
          font-weight: 600;
          display: flex;
          align-items: center;
          gap: 0.5rem;
        }

        .lgn-btn {
          width: 100%;
          padding: 1rem;
          font-size: 1rem;
          font-weight: 800;
          font-family: 'Outfit', sans-serif;
          background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 50%, #06b6d4 100%);
          color: #ffffff;
          border: 0;
          border-radius: 12px;
          cursor: pointer;
          box-shadow: 0 10px 25px -5px rgba(79, 70, 229, 0.4);
          transition: all 0.2s ease;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 0.5rem;
          margin-top: 0.5rem;
        }

        .lgn-btn:hover:not(:disabled) {
          transform: translateY(-1px);
          box-shadow: 0 15px 30px -5px rgba(79, 70, 229, 0.5);
        }

        .lgn-btn:disabled {
          opacity: 0.65;
          cursor: not-allowed;
          transform: none;
        }

        .lgn-spinner {
          width: 18px;
          height: 18px;
          border: 2px solid rgba(255, 255, 255, 0.3);
          border-top-color: #ffffff;
          border-radius: 50%;
          animation: lgnSpin 0.8s linear infinite;
        }

        @keyframes lgnSpin {
          to { transform: rotate(360deg); }
        }

        .lgn-footer {
          margin-top: 2rem;
          padding-top: 1.25rem;
          border-top: 1px solid rgba(255, 255, 255, 0.08);
          text-align: center;
          font-size: 0.75rem;
          color: #64748b;
          display: flex;
          align-items: center;
          justify-content: space-between;
        }
      `}</style>

      <div className="lgn-root">
        <div className="lgn-ambient-bg"></div>
        <div className="lgn-grid-pattern"></div>

        <motion.div
          className="lgn-card"
          initial={{ opacity: 0, y: 25, scale: 0.96 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          transition={{ duration: 0.5, ease: "easeOut" }}
        >
          {/* Header */}
          <div className="lgn-header">
            <motion.div
              className="lgn-logo-badge"
              initial={{ scale: 0.7, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              transition={{ delay: 0.1, duration: 0.4 }}
            >
              MH
            </motion.div>
            <h1 className="lgn-title">MH Factory Suite Pro</h1>
            <p className="lgn-subtitle">Enterprise Apparel Production System</p>
            <div className="lgn-status-pill">
              <span>🟢 System Online</span>
              <span style={{ opacity: 0.5 }}>|</span>
              <span>Secured Portal</span>
            </div>
          </div>

          {/* Form */}
          <form onSubmit={handleLogin}>
            <div className="lgn-field-group">
              <label className="lgn-label">User ID / Username</label>
              <div className="lgn-input-wrapper">
                <input
                  type="text"
                  className="lgn-input"
                  placeholder="Enter your Username..."
                  value={id}
                  onChange={(e) => {
                    setId(e.target.value);
                    setError("");
                  }}
                  autoFocus
                />
                <span className="lgn-input-icon">👤</span>
              </div>
            </div>

            <div className="lgn-field-group">
              <label className="lgn-label">Password</label>
              <div className="lgn-input-wrapper">
                <input
                  type={showPassword ? "text" : "password"}
                  className="lgn-input"
                  placeholder="Enter your Password..."
                  value={password}
                  onChange={(e) => {
                    setPassword(e.target.value);
                    setError("");
                  }}
                />
                <span className="lgn-input-icon">🔒</span>
                <button
                  type="button"
                  className="lgn-toggle-pass"
                  onClick={() => setShowPassword((prev) => !prev)}
                  title={showPassword ? "Hide Password" : "Show Password"}
                >
                  {showPassword ? "🙈" : "👁️"}
                </button>
              </div>
            </div>

            {/* Error message */}
            <AnimatePresence>
              {error && (
                <motion.div
                  className="lgn-error-box"
                  initial={{ opacity: 0, height: 0, y: -5 }}
                  animate={{ opacity: 1, height: "auto", y: 0 }}
                  exit={{ opacity: 0, height: 0, y: -5 }}
                  transition={{ duration: 0.2 }}
                >
                  <span>⚠️</span>
                  <span>{error}</span>
                </motion.div>
              )}
            </AnimatePresence>

            <motion.button
              type="submit"
              className="lgn-btn"
              disabled={isLoading}
              whileHover={{ scale: 1.01 }}
              whileTap={{ scale: 0.99 }}
            >
              {isLoading ? (
                <>
                  <div className="lgn-spinner"></div>
                  <span>Authenticating...</span>
                </>
              ) : (
                <>
                  <span>Sign In to Dashboard</span>
                  <span>➔</span>
                </>
              )}
            </motion.button>
          </form>

          {/* Footer */}
          <div className="lgn-footer">
            <span>© {new Date().getFullYear()} MH Factory Suite</span>
            <span>{currentTime || "System Ready"}</span>
          </div>
        </motion.div>
      </div>
    </>
  );
};

export default LoginScreen;