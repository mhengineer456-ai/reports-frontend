// src/productionMilestones.js
import { GOOGLE_API_KEY, SPREADSHEET_IDS } from './config';

/**
 * Real Production Milestones Service
 * Dynamically fetches ACTUAL live production data from Google Sheets for TODAY.
 * Supports Index Sheet (Stitching Issues), Stitching_Issues Sheet, JobOrder Sheet, and Issues Sheet.
 */

export const isDateToday = (dateInput) => {
  if (!dateInput) return false;
  const s = String(dateInput).trim();
  if (!s) return false;

  let d = null;

  // Try ISO
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const parts = s.split("-");
    d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  } 
  // Try DD/MM/YYYY or DD-MM-YYYY
  else if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}/.test(s)) {
    const parts = s.split(/[\/\-\.]/);
    let y = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    d = new Date(y, parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
  }
  // Try DD MMM YYYY (e.g. 06 Aug 2026 or 06-Aug-2026)
  else if (/^\d{1,2}[-\s][A-Za-z]{3}[-\s]\d{2,4}/.test(s)) {
    const m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3})[-\s](\d{2,4})$/);
    if (m) {
      const monMap = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
      const dd = parseInt(m[1], 10);
      const mon = monMap[m[2].toLowerCase()];
      let y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      if (mon !== undefined) d = new Date(y, mon, dd);
    }
  } else {
    d = new Date(s);
  }

  if (!d || isNaN(d.getTime())) return false;

  const today = new Date();
  return (
    d.getDate() === today.getDate() &&
    d.getMonth() === today.getMonth() &&
    d.getFullYear() === today.getFullYear()
  );
};

/**
 * Fetch ACTUAL Live Notifications from Google Sheets for TODAY
 */
export const fetchTodayMilestones = async () => {
  const milestones = [];
  const processedLots = new Set();

  // 1. Fetch Daily Stitching Issues from Index Sheet (1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA)
  try {
    const indexUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_IDS.MAIN}/values/Index!A:O?key=${GOOGLE_API_KEY}`;
    const indexRes = await fetch(indexUrl);

    if (indexRes.ok) {
      const indexJson = await indexRes.json();
      const [headers, ...rows] = indexJson.values || [];

      if (headers && rows.length > 0) {
        const getCol = (name) => headers.findIndex(h => h && h.trim().toLowerCase().replace(/[^a-z0-9]/g, "") === name.toLowerCase().replace(/[^a-z0-9]/g, ""));

        const lotCol = getCol("Lot Number") >= 0 ? getCol("Lot Number") : getCol("Lot");
        const dateOfIssueCol = getCol("Date of Issue") >= 0 ? getCol("Date of Issue") : getCol("Date");
        const supervisorCol = getCol("Supervisor");
        const garmentCol = getCol("Garment Type") >= 0 ? getCol("Garment Type") : getCol("Garment");
        const brandCol = getCol("Brand");

        rows.forEach((row, idx) => {
          const lotNo = row[lotCol] ? String(row[lotCol]).trim() : "";
          const dateOfIssue = row[dateOfIssueCol] ? String(row[dateOfIssueCol]).trim() : "";
          const supervisor = row[supervisorCol] ? String(row[supervisorCol]).trim() : "";
          const garment = row[garmentCol] ? String(row[garmentCol]).trim() : "";
          const brand = row[brandCol] ? String(row[brandCol]).trim() : "";

          if (lotNo && isDateToday(dateOfIssue) && !processedLots.has(lotNo)) {
            processedLots.add(lotNo);
            milestones.push({
              id: `STITCH-ISSUE-INDEX-${idx}-${Date.now()}`,
              title: `📍 Stitching Issued Today: Lot #${lotNo}`,
              desc: `Lot #${lotNo} issued to supervisor ${supervisor || 'N/A'} ${garment ? `[${garment}]` : ''} ${brand ? `[${brand}]` : ''}`,
              supervisor: supervisor || 'N/A',
              time: "Today",
              unread: true,
              type: "warning",
              link: "/stitching",
              category: "Stitching Issue"
            });
          }
        });
      }
    }
  } catch (err) {
    console.error("Error fetching Index sheet for Daily Stitching Issue milestones:", err);
  }

  // 2. Fetch Daily Stitching Issues from Old Sheet (18FzakygM7DVD29IRbpe68pDeCFQhFLj7t4C-XQ1MWWc)
  try {
    const oldSheetId = "18FzakygM7DVD29IRbpe68pDeCFQhFLj7t4C-XQ1MWWc";
    const oldUrl = `https://sheets.googleapis.com/v4/spreadsheets/${oldSheetId}/values/Stitching_Issues!A:Q?key=${GOOGLE_API_KEY}`;
    const oldRes = await fetch(oldUrl);

    if (oldRes.ok) {
      const oldJson = await oldRes.json();
      const [headers, ...rows] = oldJson.values || [];

      if (headers && rows.length > 0) {
        const getCol = (name) => headers.findIndex(h => h && h.trim().toLowerCase().replace(/[^a-z0-9]/g, "") === name.toLowerCase().replace(/[^a-z0-9]/g, ""));

        const lotCol = getCol("Lot Number") >= 0 ? getCol("Lot Number") : getCol("Lot");
        const dateOfIssueCol = getCol("Date of Issue") >= 0 ? getCol("Date of Issue") : getCol("Date");
        const supervisorCol = getCol("Supervisor");
        const garmentCol = getCol("Garment Type");
        const brandCol = getCol("Brand");

        rows.forEach((row, idx) => {
          const lotNo = row[lotCol] ? String(row[lotCol]).trim() : "";
          const dateOfIssue = row[dateOfIssueCol] ? String(row[dateOfIssueCol]).trim() : "";
          const supervisor = row[supervisorCol] ? String(row[supervisorCol]).trim() : "";
          const garment = row[garmentCol] ? String(row[garmentCol]).trim() : "";
          const brand = row[brandCol] ? String(row[brandCol]).trim() : "";

          if (lotNo && isDateToday(dateOfIssue) && !processedLots.has(lotNo)) {
            processedLots.add(lotNo);
            milestones.push({
              id: `STITCH-ISSUE-OLD-${idx}-${Date.now()}`,
              title: `📍 Stitching Issued Today: Lot #${lotNo}`,
              desc: `Lot #${lotNo} issued to supervisor ${supervisor || 'N/A'} ${garment ? `[${garment}]` : ''} ${brand ? `[${brand}]` : ''}`,
              supervisor: supervisor || 'N/A',
              time: "Today",
              unread: true,
              type: "warning",
              link: "/stitching",
              category: "Stitching Issue"
            });
          }
        });
      }
    }
  } catch (err) {
    console.error("Error fetching Stitching_Issues sheet for Daily Stitching Issue milestones:", err);
  }

  // 3. Fetch Real Job Order Sheet (Embroidery & Printing Challans & Cutting)
  try {
    const jobOrderUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_IDS.JOBORDER}/values/JobOrder!A1:ZZZ?key=${GOOGLE_API_KEY}`;
    const jobRes = await fetch(jobOrderUrl);

    if (jobRes.ok) {
      const jobJson = await jobRes.json();
      const [headers, ...rows] = jobJson.values || [];

      if (headers && rows.length > 0) {
        const getCol = (name) => headers.findIndex(h => h && h.trim().toLowerCase() === name.toLowerCase());

        const lotCol = getCol("Lot Number");
        const challanCol = getCol("Challan No");
        const challanDateCol = getCol("Challan Date");
        const cuttingDateCol = getCol("Cutting Date");
        const totalQtyCol = getCol("Challan Total Qty");
        const styleCol = getCol("Style");

        rows.forEach((row, idx) => {
          const lotNo = row[lotCol] ? String(row[lotCol]).trim() : "";
          const challanNo = row[challanCol] ? String(row[challanCol]).trim() : "";
          const challanDate = row[challanDateCol] ? String(row[challanDateCol]).trim() : "";
          const cuttingDate = row[cuttingDateCol] ? String(row[cuttingDateCol]).trim() : "";
          const qty = row[totalQtyCol] ? String(row[totalQtyCol]).trim() : "";
          const style = row[styleCol] ? String(row[styleCol]).trim() : "";

          // Real Embroidery Challan Issued Today
          if (challanNo.startsWith("CH-EMB-") && isDateToday(challanDate)) {
            milestones.push({
              id: `EMB-REAL-${idx}-${Date.now()}`,
              title: `📈 Embroidery Challan Today: Lot #${lotNo}`,
              desc: `Challan #${challanNo} issued today (${qty || 'N/A'} pcs) ${style ? `[Style: ${style}]` : ''}`,
              time: "Today",
              unread: true,
              type: "warning",
              link: "/embroidery",
              category: "Embroidery"
            });
          }

          // Real Printing Challan Issued Today
          if ((challanNo.startsWith("CH-PRT-") || challanNo.startsWith("CH-PRINT-")) && isDateToday(challanDate)) {
            milestones.push({
              id: `PRT-REAL-${idx}-${Date.now()}`,
              title: `🖨️ Printing Challan Today: Lot #${lotNo}`,
              desc: `Challan #${challanNo} issued today (${qty || 'N/A'} pcs) ${style ? `[Style: ${style}]` : ''}`,
              time: "Today",
              unread: true,
              type: "info",
              link: "/printing",
              category: "Printing"
            });
          }

          // Real Cutting Done Today
          if (cuttingDate && isDateToday(cuttingDate)) {
            milestones.push({
              id: `CUT-REAL-${idx}-${Date.now()}`,
              title: `✂️ Cutting Completed Today: Lot #${lotNo}`,
              desc: `Cutting completed today for Lot #${lotNo} ${style ? `[Style: ${style}]` : ''}`,
              time: "Today",
              unread: true,
              type: "info",
              link: "/cutting-report",
              category: "Cutting"
            });
          }
        });
      }
    }
  } catch (err) {
    console.error("Error fetching real JobOrder milestones:", err);
  }

  // 4. Fetch Real Stitching Completion from Issues Sheet
  try {
    const issuesUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_IDS.ISSUES}/values/Issues!A1:ZZZ?key=${GOOGLE_API_KEY}`;
    const issuesRes = await fetch(issuesUrl);

    if (issuesRes.ok) {
      const issuesJson = await issuesRes.json();
      const [headers, ...rows] = issuesJson.values || [];

      if (headers && rows.length > 0) {
        const getCol = (name) => headers.findIndex(h => h && h.trim().toLowerCase() === name.toLowerCase());

        const lotCol = getCol("Lot Number");
        const completedStatusCol = getCol("Completed Status");
        const supervisorCol = getCol("Supervisor");

        rows.forEach((row, idx) => {
          const lotNo = row[lotCol] ? String(row[lotCol]).trim() : "";
          const completedStatusRaw = row[completedStatusCol] ? String(row[completedStatusCol]).trim() : "";
          const supervisor = row[supervisorCol] ? String(row[supervisorCol]).trim() : "";

          if (completedStatusRaw.startsWith("[")) {
            try {
              const statusArr = JSON.parse(completedStatusRaw);
              if (Array.isArray(statusArr)) {
                const completeEntry = statusArr.find(entry => entry.status && entry.status.toLowerCase().includes("complete"));
                if (completeEntry && completeEntry.timestamp && isDateToday(completeEntry.timestamp)) {
                  milestones.push({
                    id: `STITCH-REAL-${idx}-${Date.now()}`,
                    title: `✨ Stitching Completed Today: Lot #${lotNo}`,
                    desc: `Stitching marked completed today ${supervisor ? `under supervisor ${supervisor}` : ''}`,
                    time: "Today",
                    unread: true,
                    type: "success",
                    link: "/stitching-complete-lot",
                    category: "Stitching"
                  });
                }
              }
            } catch (e) {}
          }
        });
      }
    }
  } catch (err) {
    console.error("Error fetching real Stitching milestones:", err);
  }

  return milestones;
};
