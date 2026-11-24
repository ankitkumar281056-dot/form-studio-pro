// ============================================================================
// FORUM STUDIO PRO v13.3 — UNIVERSAL BULLETPROOF APPS SCRIPT (Code.gs)
// ============================================================================
// Works with ANY Google Sheet layout, ANY tab, and ANY number of rows (20, 50, 100+)!
// Zero hardcoded sheet IDs. Never clamps rows at 8-10 links and never overwrites in AUTO.
//
// SETUP / UPDATE STEPS:
// 1. Open https://script.new (or your existing script project).
// 2. Paste this entire code (replacing ALL old code) and press Ctrl+S.
// 3. Select "authorizeSheetAccess" at the top and click "▷ Run" once -> Allow.
// 4. Click Deploy -> New deployment (or Manage deployments -> Edit -> Version: New version)
//    -> Web app -> Who has access: Anyone -> Deploy.
// 5. Paste the Web App URL into "🔗 Apps Script URL" in Forum Studio Pro.
// ============================================================================

const POSTING_TAB_NAME = "posting site ";

function authorizeSheetAccess() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("✅ Google Sheets access authorized! Deploy -> New deployment -> Web app -> Anyone.");
}

function extractSpreadsheetId(idOrUrl) {
  const s = String(idOrUrl || "").trim();
  if (!s) return "";
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]{15,})/);
  return m ? m[1] : s;
}

function getTargetSheet(sheetIdOrUrl, tabName, gid) {
  const ssId = extractSpreadsheetId(sheetIdOrUrl);
  if (!ssId) {
    throw new Error("No Main Sheet Link provided. Paste your Google Sheet URL into 'Main Sheet Link' first.");
  }
  const ss = SpreadsheetApp.openById(ssId);
  if (gid !== undefined && gid !== null && String(gid).trim() !== "") {
    const numGid = Number(String(gid).trim());
    const allSheets = ss.getSheets();
    for (let i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getSheetId() === numGid) {
        return allSheets[i];
      }
    }
  }
  const reqTab = String(tabName || POSTING_TAB_NAME).trim();
  const byName = (
    ss.getSheetByName(tabName || POSTING_TAB_NAME) ||
    ss.getSheetByName(reqTab) ||
    ss.getSheetByName("posting site ") ||
    ss.getSheetByName("posting site")
  );
  if (byName) return byName;

  // Smart fallback: find the tab that has "S.no" or "keyword" or "link" headers (avoids blank "Untitled" tabs)
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    try {
      const topRows = sheets[i].getRange(1, 1, Math.min(4, Math.max(2, sheets[i].getLastRow())), Math.min(6, Math.max(4, sheets[i].getLastColumn()))).getValues();
      for (let r = 0; r < topRows.length; r++) {
        const line = topRows[r].map(v => String(v || "").trim().toLowerCase());
        if (line.indexOf("s.no") !== -1 || line.indexOf("sno") !== -1 || line.indexOf("keyword") !== -1 || line.indexOf("link") !== -1) {
          return sheets[i];
        }
      }
    } catch (e) {}
  }
  return ss.getSheetByName("Sheet1") || sheets[0];
}

function detectHeaderAndStartRow(sheet) {
  try {
    const scanR = Math.min(Math.max(sheet.getLastRow(), 2), 6);
    const scanC = Math.min(Math.max(sheet.getLastColumn(), 5), 10);
    const vals = sheet.getRange(1, 1, scanR, scanC).getValues();
    for (let r = 0; r < vals.length; r++) {
      const rowLow = vals[r].map(v => String(v || "").trim().toLowerCase());
      if (rowLow.indexOf("s.no") !== -1 || rowLow.indexOf("sno") !== -1 || rowLow.indexOf("keyword") !== -1 || rowLow.indexOf("link") !== -1) {
        return { headerRow: r + 1, startRow: r + 2 };
      }
    }
  } catch (e) {}
  return { headerRow: 2, startRow: 3 };
}

function getFirstLinkCol(sheet, headerRow) {
  const hdrR = headerRow || 2;
  try {
    const maxC = Math.min(Math.max(sheet.getLastColumn(), 8), 18);
    const hdrs = sheet.getRange(hdrR, 1, 1, maxC).getValues()[0];
    for (let c = 2; c < hdrs.length; c++) {
      const h = String(hdrs[c] || "").trim().toLowerCase();
      if (h.indexOf("s.no") !== -1 || h.indexOf("date") !== -1 || h.indexOf("keyword") !== -1 || h.indexOf("spanish") !== -1 || h.indexOf("topic") !== -1) {
        continue;
      }
      if (h === "link" || h === "links" || h === "url" || h === "") {
        return c + 1;
      }
    }
  } catch (e) {}
  return 4;
}

function getDynamicBlockEndRow(sheet, startRow, requestedEndRow) {
  const sRow = startRow || 3;
  const minTarget = Math.max(sRow + 19, Number(requestedEndRow || 0)); // At least 20 rows (e.g. Row 22)
  try {
    const maxScan = Math.min(Math.max(sheet.getLastRow(), minTarget), 200);
    const vals = sheet.getRange(sRow, 1, maxScan - sRow + 1, 4).getValues();
    let lastValidRow = sRow - 1;
    for (let i = 0; i < vals.length; i++) {
      const rowJoined = vals[i].map(v => String(v !== null && v !== undefined ? v : "").trim()).join("");
      const c0 = String(vals[i][0] || "").trim().toLowerCase();
      const c2 = String(vals[i][2] || "").trim().toLowerCase();
      if (i > 0 && (c0 === "s.no" || c0 === "sno" || c2 === "keyword")) {
        break;
      }
      if (i > 1 && c0 === "1") {
        break;
      }
      if (!rowJoined) {
        // Only break on a completely blank row AFTER we've seen at least 5 rows, or keep scanning up to minTarget
        if ((sRow + i) > minTarget) break;
        continue;
      }
      lastValidRow = sRow + i;
    }
    return Math.max(lastValidRow, minTarget);
  } catch (err) {}
  return minTarget;
}

function colLetterToIndex(letter) {
  if (!letter) return 0;
  const clean = String(letter).trim().toUpperCase().replace(/[^A-Z]/g, "");
  if (!clean || clean === "AUTO") return 0;
  let col = 0;
  for (let i = 0; i < clean.length; i++) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return col;
}

function colIndexToLetter(colIdx) {
  let temp, letter = "";
  while (colIdx > 0) {
    temp = (colIdx - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIdx = (colIdx - temp - 1) / 26;
  }
  return letter;
}

function findFirstEmptyLinkCol(sh, sCol, sRow, eRow) {
  const nRows = Math.max(1, eRow - sRow + 1);
  const maxC = Math.max(sh.getLastColumn() + 2, sCol + 14);
  const grid = sh.getRange(sRow, sCol, nRows, maxC - sCol + 1).getValues();
  let rightmostActive = sCol - 1;
  for (let c = 0; c < grid[0].length; c++) {
    for (let r = 0; r < nRows; r++) {
      if (String(grid[r][c] || "").trim().indexOf("http") === 0) {
        rightmostActive = sCol + c;
        break;
      }
    }
  }
  return rightmostActive >= sCol ? (rightmostActive + 1) : sCol;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = String(params.action || "").trim();

    if (action === "poll_remote_cmd") {
      const props = PropertiesService.getScriptProperties();
      const rawCmd = props.getProperty("REMOTE_CMD_QUEUE") || "";
      if (rawCmd) {
        props.deleteProperty("REMOTE_CMD_QUEUE");
        return jsonResponse({ status: "success", command: JSON.parse(rawCmd) });
      }
      return jsonResponse({ status: "empty" });
    }

    const ssId = extractSpreadsheetId(params.sheet_id || params.sheet_url);
    const sheet = getTargetSheet(ssId, params.tab_name, params.gid);
    const hdrInfo = detectHeaderAndStartRow(sheet);
    const startCol = getFirstLinkCol(sheet, hdrInfo.headerRow);
    const startRow = hdrInfo.startRow;
    const endRow = getDynamicBlockEndRow(sheet, startRow, Number(params.end_row || 0));
    const numRows = endRow - startRow + 1;

    const reqCol = String(params.col || params.column || "AUTO").trim().toUpperCase();
    let targetCol = colLetterToIndex(reqCol);

    if (targetCol < startCol) {
      const lastCol = Math.max(sheet.getLastColumn(), startCol);
      const grid = sheet.getRange(startRow, startCol, numRows, lastCol - startCol + 1).getValues();
      targetCol = startCol;
      for (let c = grid[0].length - 1; c >= 0; c--) {
        let hasLink = false;
        for (let r = 0; r < numRows; r++) {
          const cellVal = String(grid[r][c] || "").trim();
          if (cellVal.indexOf("http") === 0) {
            hasLink = true;
            break;
          }
        }
        if (hasLink) {
          targetCol = startCol + c;
          break;
        }
      }
    }

    const colVals = sheet.getRange(startRow, targetCol, numRows, 1).getValues();
    const links = [];
    for (let i = 0; i < colVals.length; i++) {
      const v = String(colVals[i][0] || "").trim();
      if (v.indexOf("http") === 0) {
        links.push(v);
      }
    }

    return jsonResponse({
      status: "success",
      ok: true,
      sheet_id: ssId,
      sheet_name: sheet.getParent().getName(),
      tab_name: sheet.getName(),
      column: colIndexToLetter(targetCol),
      start_row: startRow,
      end_row: endRow,
      keyword_count: numRows,
      links: links
    });
  } catch (err) {
    return jsonResponse({ status: "error", ok: false, message: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (lockErr) {}

  try {
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const data = JSON.parse(rawBody);
    const action = String(data.action || "save_post_link").trim();

    if (action === "queue_remote_cmd") {
      PropertiesService.getScriptProperties().setProperty("REMOTE_CMD_QUEUE", JSON.stringify(data));
      return jsonResponse({ status: "success", queued: true });
    }

    const ssId = extractSpreadsheetId(data.sheet_id || data.spreadsheet_id || data.sheet_url);
    const sheet = getTargetSheet(ssId, data.tab_name, data.gid);
    const hdrInfo = detectHeaderAndStartRow(sheet);
    const headerRow = hdrInfo.headerRow;
    const startRow = hdrInfo.startRow;
    const startCol = getFirstLinkCol(sheet, headerRow);

    if (action === "save_batch_column" && Array.isArray(data.links)) {
      const endRow = getDynamicBlockEndRow(sheet, startRow, startRow + data.links.length - 1);
      let colLetBatch = String(data.col_letter || data.column || data.col || "AUTO").trim().toUpperCase();
      let colIdxBatch = colLetterToIndex(colLetBatch);
      if (colIdxBatch < startCol) {
        colIdxBatch = findFirstEmptyLinkCol(sheet, startCol, startRow, endRow);
      }
      if (!String(sheet.getRange(headerRow, colIdxBatch).getValue() || "").trim()) {
        sheet.getRange(headerRow, colIdxBatch).setValue("link");
      }
      for (let i = 0; i < data.links.length; i++) {
        const u = String(data.links[i] || "").trim();
        if (u.indexOf("http") === 0) {
          sheet.getRange(startRow + i, colIdxBatch).setValue(u);
        }
      }
      SpreadsheetApp.flush();
      return jsonResponse({
        status: "success",
        ok: true,
        sheet_id: ssId,
        sheet_name: sheet.getParent().getName(),
        tab_name: sheet.getName(),
        column: colIndexToLetter(colIdxBatch),
        col: colIndexToLetter(colIdxBatch),
        written: data.links.length
      });
    }

    const liveUrl = String(data.link || data.url || data.live_url || "").trim();
    if (!liveUrl) {
      return jsonResponse({ status: "error", ok: false, message: "No link provided" });
    }

    // Exact targetRow: NEVER clamp to an early endRow!
    let targetRow = Number(data.target_row || data.row_index || data.row || 0);
    if (!targetRow || targetRow < startRow) {
      const itemIndex = Number(data.item_index || data.row_offset || data.s_no || data.id || 0);
      if (itemIndex >= 1) {
        targetRow = startRow + itemIndex - 1;
      } else {
        targetRow = startRow;
      }
    }

    const endRow = getDynamicBlockEndRow(sheet, startRow, Math.max(targetRow, Number(data.end_row || 0)));
    const reqColLetter = String(data.column || data.col || data.col_letter || "").trim().toUpperCase();
    let targetCol = colLetterToIndex(reqColLetter);
    if (targetCol < startCol) {
      targetCol = findFirstEmptyLinkCol(sheet, startCol, startRow, endRow);
    } else if (data.no_overwrite === true || reqColLetter === "AUTO") {
      const existingCell = String(sheet.getRange(targetRow, targetCol).getValue() || "").trim();
      if (existingCell.indexOf("http") === 0) {
        targetCol = findFirstEmptyLinkCol(sheet, startCol, startRow, endRow);
      }
    }

    if (!String(sheet.getRange(headerRow, targetCol).getValue() || "").trim()) {
      sheet.getRange(headerRow, targetCol).setValue("link");
    }
    sheet.getRange(targetRow, targetCol).setValue(liveUrl);
    SpreadsheetApp.flush();

    return jsonResponse({
      status: "success",
      ok: true,
      sheet_id: ssId,
      sheet_name: sheet.getParent().getName(),
      tab_name: sheet.getName(),
      column: colIndexToLetter(targetCol),
      col: colIndexToLetter(targetCol),
      row: targetRow,
      start_row: startRow,
      end_row: endRow,
      link: liveUrl
    });
  } catch (err) {
    return jsonResponse({ status: "error", ok: false, message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
