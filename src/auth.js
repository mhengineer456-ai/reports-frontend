// src/auth.js

const DEFAULT_USERS = [
  { id: "admin", name: "Administrator", password: "1234", role: "Admin", avatarColor: "#4f46e5" },
  { id: "manager", name: "Production Manager", password: "pass", role: "Manager", avatarColor: "#059669" },
  { id: "supervisor", name: "Stitching Supervisor", password: "123", role: "Supervisor", avatarColor: "#d97706" },
  { id: "viewer", name: "Guest Viewer", password: "view", role: "Viewer", avatarColor: "#2563eb" }
];

export const CREDENTIALS_SPREADSHEET_ID = "1iBDfsxA9XEC9nhQE-ALBYlyGRZWOaCYvWsnGfYYbr1I";
export const CREDENTIALS_SHEET_TAB = "ReportsLoginCredentials";

export const getRegisteredUsers = () => {
  try {
    const rawEnv = process.env.REACT_APP_USERS_CONFIG;
    if (rawEnv) {
      const parsed = JSON.parse(rawEnv);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
  } catch (err) {
    console.warn("Failed to parse REACT_APP_USERS_CONFIG from .env, using default users:", err);
  }
  return DEFAULT_USERS;
};

export const loginUser = (userId, password) => {
  const users = getRegisteredUsers();
  const trimmedId = (userId || "").trim().toLowerCase();
  const trimmedPass = (password || "").trim();

  const matchedUser = users.find(
    (u) => u.id.toLowerCase() === trimmedId && u.password === trimmedPass
  );

  if (matchedUser) {
    const sessionUser = {
      id: matchedUser.id,
      name: matchedUser.name || matchedUser.id,
      role: matchedUser.role || "User",
      avatarColor: matchedUser.avatarColor || "#4f46e5",
      loggedInAt: new Date().toISOString()
    };
    localStorage.setItem("factory_auth_user", JSON.stringify(sessionUser));
    return sessionUser;
  }
  return null;
};

export const loginUserAsync = async (userId, password) => {
  const trimmedId = (userId || "").trim().toLowerCase();
  const trimmedPass = (password || "").trim();

  // 1. Try Google Sheets Credentials First
  try {
    const apiKey = process.env.REACT_APP_GOOGLE_API_KEY || "AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk";
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CREDENTIALS_SPREADSHEET_ID}/values/${CREDENTIALS_SHEET_TAB}!A1:C200?key=${apiKey}`;

    const res = await fetch(url);
    if (res.ok) {
      const data = await res.json();
      const rows = data.values || [];

      if (rows.length > 0) {
        // Find header row or default to Col A=Username, B=Password, C=Role
        const headerRow = rows[0].map((h) => String(h || "").trim().toLowerCase());
        let userIdx = headerRow.findIndex((h) => h.includes("user") || h.includes("id") || h.includes("name"));
        let passIdx = headerRow.findIndex((h) => h.includes("pass"));
        let roleIdx = headerRow.findIndex((h) => h.includes("role"));

        if (userIdx === -1) userIdx = 0;
        if (passIdx === -1) passIdx = 1;
        if (roleIdx === -1) roleIdx = 2;

        const dataRows = rows.slice(1);
        const matchedRow = dataRows.find((r) => {
          const uName = String(r[userIdx] || "").trim().toLowerCase();
          const uPass = String(r[passIdx] || "").trim();
          return uName === trimmedId && uPass === trimmedPass;
        });

        if (matchedRow) {
          const uNameRaw = String(matchedRow[userIdx] || "").trim();
          const uRoleRaw = String(matchedRow[roleIdx] || "User").trim();

          const sessionUser = {
            id: uNameRaw,
            name: uNameRaw,
            role: uRoleRaw,
            avatarColor: uRoleRaw.toLowerCase().includes("admin") ? "#4f46e5" : "#059669",
            loggedInAt: new Date().toISOString()
          };
          localStorage.setItem("factory_auth_user", JSON.stringify(sessionUser));
          return sessionUser;
        }
      }
    }
  } catch (err) {
    console.warn("Sheet credentials fetch failed, checking local auth:", err);
  }

  // 2. Fallback to local / env users
  return loginUser(userId, password);
};

export const getCurrentUser = () => {
  try {
    const stored = localStorage.getItem("factory_auth_user");
    if (stored) {
      return JSON.parse(stored);
    }
  } catch (e) {
    console.error("Error reading current user session:", e);
  }
  return null;
};

export const logoutUser = () => {
  localStorage.removeItem("factory_auth_user");
};

export const isAuthenticated = () => {
  const currentUser = getCurrentUser();
  return !!currentUser && !!currentUser.id;
};
