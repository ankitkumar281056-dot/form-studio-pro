#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
FS | Forum Studio Pro v13.6 — Hatena/Diigo/HoYoLAB Bookmark Fix + 2s Delay + Sheet Master
100% Native Python Tkinter Desktop Application (zoom_auto_sender.py / zoom_auto_sender.pyw)
- Opens directly as a native Windows desktop window (NEVER opens a web browser for the UI).
- Reads & writes your exact local files:
  * sites_curl_vault.json
  * bookmarks_curl_vault.json
  * zoom_sender_config.json
  (Plus automatic permanent backup in %LOCALAPPDATA%\ForumStudioPro\ so you NEVER lose
   your Posting cURLs, Bookmarking cURLs, or settings when closing and reopening!)
- Fixes the 19/20 Google Sheet bug so all 20 keywords (Rows 3–22, including Row 22 for #20)
  are guaranteed to be written to your Google Sheet.
"""

import os
import sys
import re
import csv
import io
import json
import zipfile
import time
import html
import socket
import urllib.parse
import urllib.request
import ssl
import subprocess
import threading
import webbrowser
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# ==============================================================================
# ZERO-CMD / DETACHED GUI GUARD FOR WINDOWS 10/11
# Ensures that opening the tool NEVER leaves a CMD / Windows Terminal window open,
# and closing any console window NEVER closes the tool!
# ==============================================================================
def _ensure_no_console_window_on_windows():
    if os.name != "nt":
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        console_hwnd = kernel32.GetConsoleWindow()
        exe_name = os.path.basename(sys.executable or "").lower()
        if "--detached-gui" not in sys.argv and (exe_name in ("python.exe", "py.exe") or console_hwnd):
            py_dir = os.path.dirname(os.path.abspath(sys.executable))
            pyw_candidates = [
                os.path.join(py_dir, "pythonw.exe"),
                os.path.join(py_dir, "pyw.exe"),
            ]
            for pyw in pyw_candidates:
                if os.path.exists(pyw):
                    CREATE_NO_WINDOW = 0x08000000
                    DETACHED_PROCESS = 0x00000008
                    script_abs = os.path.abspath(__file__)
                    subprocess.Popen(
                        [pyw, script_abs, "--detached-gui"],
                        cwd=os.path.dirname(script_abs),
                        creationflags=CREATE_NO_WINDOW | DETACHED_PROCESS,
                        close_fds=True
                    )
                    os._exit(0)
            # Fallback if pythonw.exe is not in the same folder: hide & free console immediately
            if console_hwnd:
                user32.ShowWindow(console_hwnd, 0)  # SW_HIDE = 0
            kernel32.FreeConsole()
    except Exception:
        pass

_ensure_no_console_window_on_windows()

try:
    import requests
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    HAS_REQUESTS = True
    _HTTP_SESSION = requests.Session()
except Exception:
    HAS_REQUESTS = False
    _HTTP_SESSION = None


PRIMARY_CLOUD_RELAY = ""
BACKUP_SHEET_WEBHOOK = ""
DEFAULT_APPS_SCRIPT_URL = ""
DEFAULT_MAIN_SHEET_URL = ""
DEFAULT_CONTENT_SHEET_URL = ""  # 100% blank until user enters their own links

UPDATED_CODE_GS = "// ============================================================================\n// FORUM STUDIO PRO v13.3 \u2014 UNIVERSAL BULLETPROOF APPS SCRIPT (Code.gs)\n// ============================================================================\n// Works with ANY Google Sheet layout, ANY tab, and ANY number of rows (20, 50, 100+)!\n// Zero hardcoded sheet IDs. Never clamps rows at 8-10 links and never overwrites in AUTO.\n//\n// SETUP / UPDATE STEPS:\n// 1. Open https://script.new (or your existing script project).\n// 2. Paste this entire code (replacing ALL old code) and press Ctrl+S.\n// 3. Select \"authorizeSheetAccess\" at the top and click \"\u25b7 Run\" once -> Allow.\n// 4. Click Deploy -> New deployment (or Manage deployments -> Edit -> Version: New version)\n//    -> Web app -> Who has access: Anyone -> Deploy.\n// 5. Paste the Web App URL into \"\ud83d\udd17 Apps Script URL\" in Forum Studio Pro.\n// ============================================================================\n\nconst POSTING_TAB_NAME = \"posting site \";\n\nfunction authorizeSheetAccess() {\n  const active = SpreadsheetApp.getActiveSpreadsheet();\n  Logger.log(\"\u2705 Google Sheets access authorized! Deploy -> New deployment -> Web app -> Anyone.\");\n}\n\nfunction extractSpreadsheetId(idOrUrl) {\n  const s = String(idOrUrl || \"\").trim();\n  if (!s) return \"\";\n  const m = s.match(/\\/spreadsheets\\/d\\/([a-zA-Z0-9_-]{15,})/);\n  return m ? m[1] : s;\n}\n\nfunction getTargetSheet(sheetIdOrUrl, tabName, gid) {\n  const ssId = extractSpreadsheetId(sheetIdOrUrl);\n  if (!ssId) {\n    throw new Error(\"No Main Sheet Link provided. Paste your Google Sheet URL into 'Main Sheet Link' first.\");\n  }\n  const ss = SpreadsheetApp.openById(ssId);\n  if (gid !== undefined && gid !== null && String(gid).trim() !== \"\") {\n    const numGid = Number(String(gid).trim());\n    const allSheets = ss.getSheets();\n    for (let i = 0; i < allSheets.length; i++) {\n      if (allSheets[i].getSheetId() === numGid) {\n        return allSheets[i];\n      }\n    }\n  }\n  const reqTab = String(tabName || POSTING_TAB_NAME).trim();\n  const byName = (\n    ss.getSheetByName(tabName || POSTING_TAB_NAME) ||\n    ss.getSheetByName(reqTab) ||\n    ss.getSheetByName(\"posting site \") ||\n    ss.getSheetByName(\"posting site\")\n  );\n  if (byName) return byName;\n\n  // Smart fallback: find the tab that has \"S.no\" or \"keyword\" or \"link\" headers (avoids blank \"Untitled\" tabs)\n  const sheets = ss.getSheets();\n  for (let i = 0; i < sheets.length; i++) {\n    try {\n      const topRows = sheets[i].getRange(1, 1, Math.min(4, Math.max(2, sheets[i].getLastRow())), Math.min(6, Math.max(4, sheets[i].getLastColumn()))).getValues();\n      for (let r = 0; r < topRows.length; r++) {\n        const line = topRows[r].map(v => String(v || \"\").trim().toLowerCase());\n        if (line.indexOf(\"s.no\") !== -1 || line.indexOf(\"sno\") !== -1 || line.indexOf(\"keyword\") !== -1 || line.indexOf(\"link\") !== -1) {\n          return sheets[i];\n        }\n      }\n    } catch (e) {}\n  }\n  return ss.getSheetByName(\"Sheet1\") || sheets[0];\n}\n\nfunction detectHeaderAndStartRow(sheet) {\n  try {\n    const scanR = Math.min(Math.max(sheet.getLastRow(), 2), 6);\n    const scanC = Math.min(Math.max(sheet.getLastColumn(), 5), 10);\n    const vals = sheet.getRange(1, 1, scanR, scanC).getValues();\n    for (let r = 0; r < vals.length; r++) {\n      const rowLow = vals[r].map(v => String(v || \"\").trim().toLowerCase());\n      if (rowLow.indexOf(\"s.no\") !== -1 || rowLow.indexOf(\"sno\") !== -1 || rowLow.indexOf(\"keyword\") !== -1 || rowLow.indexOf(\"link\") !== -1) {\n        return { headerRow: r + 1, startRow: r + 2 };\n      }\n    }\n  } catch (e) {}\n  return { headerRow: 2, startRow: 3 };\n}\n\nfunction getFirstLinkCol(sheet, headerRow) {\n  const hdrR = headerRow || 2;\n  try {\n    const maxC = Math.min(Math.max(sheet.getLastColumn(), 8), 18);\n    const hdrs = sheet.getRange(hdrR, 1, 1, maxC).getValues()[0];\n    for (let c = 2; c < hdrs.length; c++) {\n      const h = String(hdrs[c] || \"\").trim().toLowerCase();\n      if (h.indexOf(\"s.no\") !== -1 || h.indexOf(\"date\") !== -1 || h.indexOf(\"keyword\") !== -1 || h.indexOf(\"spanish\") !== -1 || h.indexOf(\"topic\") !== -1) {\n        continue;\n      }\n      if (h === \"link\" || h === \"links\" || h === \"url\" || h === \"\") {\n        return c + 1;\n      }\n    }\n  } catch (e) {}\n  return 4;\n}\n\nfunction getDynamicBlockEndRow(sheet, startRow, requestedEndRow) {\n  const sRow = startRow || 3;\n  const minTarget = Math.max(sRow + 19, Number(requestedEndRow || 0)); // At least 20 rows (e.g. Row 22)\n  try {\n    const maxScan = Math.min(Math.max(sheet.getLastRow(), minTarget), 200);\n    const vals = sheet.getRange(sRow, 1, maxScan - sRow + 1, 4).getValues();\n    let lastValidRow = sRow - 1;\n    for (let i = 0; i < vals.length; i++) {\n      const rowJoined = vals[i].map(v => String(v !== null && v !== undefined ? v : \"\").trim()).join(\"\");\n      const c0 = String(vals[i][0] || \"\").trim().toLowerCase();\n      const c2 = String(vals[i][2] || \"\").trim().toLowerCase();\n      if (i > 0 && (c0 === \"s.no\" || c0 === \"sno\" || c2 === \"keyword\")) {\n        break;\n      }\n      if (i > 1 && c0 === \"1\") {\n        break;\n      }\n      if (!rowJoined) {\n        // Only break on a completely blank row AFTER we've seen at least 5 rows, or keep scanning up to minTarget\n        if ((sRow + i) > minTarget) break;\n        continue;\n      }\n      lastValidRow = sRow + i;\n    }\n    return Math.max(lastValidRow, minTarget);\n  } catch (err) {}\n  return minTarget;\n}\n\nfunction colLetterToIndex(letter) {\n  if (!letter) return 0;\n  const clean = String(letter).trim().toUpperCase().replace(/[^A-Z]/g, \"\");\n  if (!clean || clean === \"AUTO\") return 0;\n  let col = 0;\n  for (let i = 0; i < clean.length; i++) {\n    col = col * 26 + (clean.charCodeAt(i) - 64);\n  }\n  return col;\n}\n\nfunction colIndexToLetter(colIdx) {\n  let temp, letter = \"\";\n  while (colIdx > 0) {\n    temp = (colIdx - 1) % 26;\n    letter = String.fromCharCode(temp + 65) + letter;\n    colIdx = (colIdx - temp - 1) / 26;\n  }\n  return letter;\n}\n\nfunction findFirstEmptyLinkCol(sh, sCol, sRow, eRow) {\n  const nRows = Math.max(1, eRow - sRow + 1);\n  const maxC = Math.max(sh.getLastColumn() + 2, sCol + 14);\n  const grid = sh.getRange(sRow, sCol, nRows, maxC - sCol + 1).getValues();\n  let rightmostActive = sCol - 1;\n  for (let c = 0; c < grid[0].length; c++) {\n    for (let r = 0; r < nRows; r++) {\n      if (String(grid[r][c] || \"\").trim().indexOf(\"http\") === 0) {\n        rightmostActive = sCol + c;\n        break;\n      }\n    }\n  }\n  return rightmostActive >= sCol ? (rightmostActive + 1) : sCol;\n}\n\nfunction jsonResponse(obj) {\n  return ContentService.createTextOutput(JSON.stringify(obj))\n    .setMimeType(ContentService.MimeType.JSON);\n}\n\nfunction doGet(e) {\n  try {\n    const params = (e && e.parameter) ? e.parameter : {};\n    const action = String(params.action || \"\").trim();\n\n    if (action === \"poll_remote_cmd\") {\n      const props = PropertiesService.getScriptProperties();\n      const rawCmd = props.getProperty(\"REMOTE_CMD_QUEUE\") || \"\";\n      if (rawCmd) {\n        props.deleteProperty(\"REMOTE_CMD_QUEUE\");\n        return jsonResponse({ status: \"success\", command: JSON.parse(rawCmd) });\n      }\n      return jsonResponse({ status: \"empty\" });\n    }\n\n    const ssId = extractSpreadsheetId(params.sheet_id || params.sheet_url);\n    const sheet = getTargetSheet(ssId, params.tab_name, params.gid);\n    const hdrInfo = detectHeaderAndStartRow(sheet);\n    const startCol = getFirstLinkCol(sheet, hdrInfo.headerRow);\n    const startRow = hdrInfo.startRow;\n    const endRow = getDynamicBlockEndRow(sheet, startRow, Number(params.end_row || 0));\n    const numRows = endRow - startRow + 1;\n\n    const reqCol = String(params.col || params.column || \"AUTO\").trim().toUpperCase();\n    let targetCol = colLetterToIndex(reqCol);\n\n    if (targetCol < startCol) {\n      const lastCol = Math.max(sheet.getLastColumn(), startCol);\n      const grid = sheet.getRange(startRow, startCol, numRows, lastCol - startCol + 1).getValues();\n      targetCol = startCol;\n      for (let c = grid[0].length - 1; c >= 0; c--) {\n        let hasLink = false;\n        for (let r = 0; r < numRows; r++) {\n          const cellVal = String(grid[r][c] || \"\").trim();\n          if (cellVal.indexOf(\"http\") === 0) {\n            hasLink = true;\n            break;\n          }\n        }\n        if (hasLink) {\n          targetCol = startCol + c;\n          break;\n        }\n      }\n    }\n\n    const colVals = sheet.getRange(startRow, targetCol, numRows, 1).getValues();\n    const links = [];\n    for (let i = 0; i < colVals.length; i++) {\n      const v = String(colVals[i][0] || \"\").trim();\n      if (v.indexOf(\"http\") === 0) {\n        links.push(v);\n      }\n    }\n\n    return jsonResponse({\n      status: \"success\",\n      ok: true,\n      sheet_id: ssId,\n      sheet_name: sheet.getParent().getName(),\n      tab_name: sheet.getName(),\n      column: colIndexToLetter(targetCol),\n      start_row: startRow,\n      end_row: endRow,\n      keyword_count: numRows,\n      links: links\n    });\n  } catch (err) {\n    return jsonResponse({ status: \"error\", ok: false, message: String(err) });\n  }\n}\n\nfunction doPost(e) {\n  const lock = LockService.getScriptLock();\n  try {\n    lock.waitLock(25000);\n  } catch (lockErr) {}\n\n  try {\n    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : \"{}\";\n    const data = JSON.parse(rawBody);\n    const action = String(data.action || \"save_post_link\").trim();\n\n    if (action === \"queue_remote_cmd\") {\n      PropertiesService.getScriptProperties().setProperty(\"REMOTE_CMD_QUEUE\", JSON.stringify(data));\n      return jsonResponse({ status: \"success\", queued: true });\n    }\n\n    const ssId = extractSpreadsheetId(data.sheet_id || data.spreadsheet_id || data.sheet_url);\n    const sheet = getTargetSheet(ssId, data.tab_name, data.gid);\n    const hdrInfo = detectHeaderAndStartRow(sheet);\n    const headerRow = hdrInfo.headerRow;\n    const startRow = hdrInfo.startRow;\n    const startCol = getFirstLinkCol(sheet, headerRow);\n\n    if (action === \"save_batch_column\" && Array.isArray(data.links)) {\n      const endRow = getDynamicBlockEndRow(sheet, startRow, startRow + data.links.length - 1);\n      let colLetBatch = String(data.col_letter || data.column || data.col || \"AUTO\").trim().toUpperCase();\n      let colIdxBatch = colLetterToIndex(colLetBatch);\n      if (colIdxBatch < startCol) {\n        colIdxBatch = findFirstEmptyLinkCol(sheet, startCol, startRow, endRow);\n      }\n      if (!String(sheet.getRange(headerRow, colIdxBatch).getValue() || \"\").trim()) {\n        sheet.getRange(headerRow, colIdxBatch).setValue(\"link\");\n      }\n      for (let i = 0; i < data.links.length; i++) {\n        const u = String(data.links[i] || \"\").trim();\n        if (u.indexOf(\"http\") === 0) {\n          sheet.getRange(startRow + i, colIdxBatch).setValue(u);\n        }\n      }\n      SpreadsheetApp.flush();\n      return jsonResponse({\n        status: \"success\",\n        ok: true,\n        sheet_id: ssId,\n        sheet_name: sheet.getParent().getName(),\n        tab_name: sheet.getName(),\n        column: colIndexToLetter(colIdxBatch),\n        col: colIndexToLetter(colIdxBatch),\n        written: data.links.length\n      });\n    }\n\n    const liveUrl = String(data.link || data.url || data.live_url || \"\").trim();\n    if (!liveUrl) {\n      return jsonResponse({ status: \"error\", ok: false, message: \"No link provided\" });\n    }\n\n    // Exact targetRow: NEVER clamp to an early endRow!\n    let targetRow = Number(data.target_row || data.row_index || data.row || 0);\n    if (!targetRow || targetRow < startRow) {\n      const itemIndex = Number(data.item_index || data.row_offset || data.s_no || data.id || 0);\n      if (itemIndex >= 1) {\n        targetRow = startRow + itemIndex - 1;\n      } else {\n        targetRow = startRow;\n      }\n    }\n\n    const endRow = getDynamicBlockEndRow(sheet, startRow, Math.max(targetRow, Number(data.end_row || 0)));\n    const reqColLetter = String(data.column || data.col || data.col_letter || \"\").trim().toUpperCase();\n    let targetCol = colLetterToIndex(reqColLetter);\n    if (targetCol < startCol) {\n      targetCol = findFirstEmptyLinkCol(sheet, startCol, startRow, endRow);\n    } else if (data.no_overwrite === true || reqColLetter === \"AUTO\") {\n      const existingCell = String(sheet.getRange(targetRow, targetCol).getValue() || \"\").trim();\n      if (existingCell.indexOf(\"http\") === 0) {\n        targetCol = findFirstEmptyLinkCol(sheet, startCol, startRow, endRow);\n      }\n    }\n\n    if (!String(sheet.getRange(headerRow, targetCol).getValue() || \"\").trim()) {\n      sheet.getRange(headerRow, targetCol).setValue(\"link\");\n    }\n    sheet.getRange(targetRow, targetCol).setValue(liveUrl);\n    SpreadsheetApp.flush();\n\n    return jsonResponse({\n      status: \"success\",\n      ok: true,\n      sheet_id: ssId,\n      sheet_name: sheet.getParent().getName(),\n      tab_name: sheet.getName(),\n      column: colIndexToLetter(targetCol),\n      col: colIndexToLetter(targetCol),\n      row: targetRow,\n      start_row: startRow,\n      end_row: endRow,\n      link: liveUrl\n    });\n  } catch (err) {\n    return jsonResponse({ status: \"error\", ok: false, message: String(err) });\n  } finally {\n    try { lock.releaseLock(); } catch (e2) {}\n  }\n}\n"

# ==============================================================================
# FILE PATHS & PERMANENT VAULT PERSISTENCE
# (Reads/Writes sites_curl_vault.json, bookmarks_curl_vault.json, zoom_sender_config.json
#  in the tool folder AND backs them up in %LOCALAPPDATA%\ForumStudioPro\)
# ==============================================================================
def get_script_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def get_backup_dir():
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "ForumStudioPro")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def find_all_vault_dirs():
    dirs = []
    bdir = get_backup_dir()
    if bdir:
        dirs.append(bdir)
    sdir = get_script_dir()
    if sdir and sdir not in dirs:
        dirs.append(sdir)
    parent = os.path.dirname(sdir) if sdir else ""
    if parent:
        orig_d = os.path.join(parent, "zoom_desktop_typer")
        if os.path.isdir(orig_d) and orig_d not in dirs:
            dirs.append(orig_d)
    home = os.path.expanduser("~")
    for cand in (
        os.path.join(home, "Downloads", "zoom_desktop_typer"),
        os.path.join(home, "Desktop", "zoom_desktop_typer"),
        r"C:\Users\Lenovo\Downloads\zoom_desktop_typer",
    ):
        if os.path.isdir(cand) and cand not in dirs:
            dirs.append(cand)
    return dirs


def read_json_with_backup(filename, default_val):
    for d in find_all_vault_dirs():
        p = os.path.join(d, filename)
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8", errors="replace") as f:
                    data = json.load(f)
                if data:
                    return data
            except Exception:
                pass
    return default_val


def write_json_with_backup(filename, data):
    for d in (get_script_dir(), get_backup_dir()):
        if not d:
            continue
        try:
            os.makedirs(d, exist_ok=True)
            p = os.path.join(d, filename)
            with open(p, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass


def unescape_ansi_c(s):
    def repl(m):
        tok = m.group(1)
        if tok == 'r': return '\r'
        if tok == 'n': return '\n'
        if tok == 't': return '\t'
        if tok == "'": return "'"
        if tok == '"': return '"'
        if tok == '\\': return '\\'
        if tok.startswith('x') and len(tok) == 3:
            try: return chr(int(tok[1:], 16))
            except Exception: return tok
        if tok.startswith('u') and len(tok) == 5:
            try: return chr(int(tok[1:], 16))
            except Exception: return tok
        return tok
    return re.sub(r'\\(\\|\"|\'|r|n|t|x[0-9a-fA-F]{2}|u[0-9a-fA-F]{4}|.)', repl, s)


def extract_domain_from_curl(raw_curl):
    m = re.search(r'https?://([^/\s\'"\\]+)', raw_curl or '', re.I)
    return m.group(1).lower().strip('^"\'') if m else 'Unknown Site'


def get_registrable_domain(url_or_host):
    """
    Returns the registrable base domain (e.g. 'forum.buffed.de' -> 'buffed.de',
    's2s.yieldlove-ad-serving.net' -> 'yieldlove-ad-serving.net',
    'pbc.agma-analytics.de' -> 'agma-analytics.de').
    For multi-tenant forum hosts ('proboards.com', 'freeforums.net', 'boards.net'),
    includes the tenant subdomain so two different ProBoards forums are treated as distinct!
    """
    s = (url_or_host or "").strip().lower()
    if "://" in s:
        try:
            s = urllib.parse.urlparse(s).hostname or ""
        except Exception:
            pass
    s = s.split("/")[0].split(":")[0].strip(".\"'^")
    if s.startswith("www."):
        s = s[4:]
    if not s:
        return ""
    parts = [p for p in s.split(".") if p]
    if len(parts) <= 2:
        return s
    multi_tenant = ("proboards.com", "freeforums.net", "boards.net", "blogspot.com", "wordpress.com", "hatenablog.com")
    last2 = ".".join(parts[-2:])
    if last2 in multi_tenant and len(parts) >= 3:
        return ".".join(parts[-3:])
    cc_slds = {"co", "com", "org", "net", "edu", "gov", "ac", "gob", "or", "ne", "go"}
    if len(parts[-1]) == 2 and parts[-2] in cc_slds and len(parts) >= 3:
        return ".".join(parts[-3:])
    return last2


def is_same_base_domain(url_a, url_b):
    da = get_registrable_domain(url_a)
    db = get_registrable_domain(url_b)
    return bool(da and db and da == db)


AD_OR_TRACKER_INDICATORS = (
    "google-analytics.com", "googletagmanager.com", "doubleclick.net", "googlesyndication.com",
    "googleadservices.com", "google.com/recaptcha", "gstatic.com", "googleapis.com",
    "cloudflare.com", "cloudflareinsights.com", "facebook.com", "facebook.net", "clarity.ms",
    "bing.com", "criteo.com", "criteo.net", "adsystem.com", "adservice.", "adserver.",
    "adnxs.com", "rubiconproject.com", "pubmatic.com", "openx.net", "casalemedia.com",
    "smartadserver.com", "taboola.com", "outbrain.com", "yieldlove", "agma-analytics",
    "ioam.de", "ivwbox.de", "stroeer", "sentry.io", "datadoghq.com", "hotjar.com",
    "newrelic.com", "nr-data.net", "segment.io", "segment.com", "mixpanel.com", "amplitude.com",
    "matomo", "piwik", "chartbeat", "parsely", "comscore", "scorecardresearch.com",
    "quantserve.com", "quantcast.com", "onetrust.com", "cookielaw.org", "cookiebot.com",
    "usercentrics.eu", "consentmanager.net", "sourcepoint", "didomi.io", "trustarc.com",
    "play.google.com", "accounts.google.com", "s2s.", "pbc."
)


def is_ad_or_tracker_url(url_or_curl):
    u_low = (url_or_curl or "").lower()
    dom = extract_domain_from_curl(u_low) if "://" in u_low else u_low
    if any(ind in dom or ind in u_low for ind in AD_OR_TRACKER_INDICATORS):
        return True
    if re.search(r'(^|\.)(ads|ad|analytics|telemetry|metrics|beacon|pixel|tracker|stats|cmp|consent)\.', dom):
        return True
    return False


def derive_clean_site_name_from_domain(dom):
    d = (dom or '').lower().strip()
    if not d or d == 'unknown site':
        return 'Site'
    parts = [p for p in d.split('.') if p and p not in ('www', 'forum', 'forums', 'community', 'proboards', 'boards', 'freeforums', 'com', 'net', 'org', 'co', 'io', 'edu', 'gov', 'ar', 'uk', 'de', 'jp', 'in')]
    base = parts[0] if parts else d.replace('www.', '').split('.')[0]
    return base.capitalize()


def derive_composer_url_from_curl(raw_curl):
    """
    Derives the composer/posting page URL from a cURL command WITHOUT EVER replacing a deep
    composer URL (like `https://forum.buffed.de/forums/wow-allgemeine-diskussionen.10/post-thread`)
    with a bare root homepage (`https://forum.buffed.de/`) when Referrer-Policy is `origin`!
    """
    norm = re.sub(r'\\\r?\n\s*', ' ', raw_curl or '')
    norm = re.sub(r'\^\r?\n\s*', ' ', norm)

    post_u = ""
    url_m = (
        re.search(r'--url\s+(?:\$\'((?:\\.|[^\'\\])*)\'|\'([^\']*)\'|"((?:\\.|[^"\\])*)"|(\S+))', norm, re.I)
        or re.search(r'curl(?:\.exe)?\s+(?:-[a-zA-Z]+\s+)*[\'"]?(https?://[^\s\'"\\]+)[\'"]?', norm, re.I)
        or re.search(r'(?:--url\s+)?\$?[\'"](https?://[^\'"]+)[\'"]', norm, re.I)
        or re.search(r'(https?://[^\s\'"\\]+)', norm, re.I)
    )
    if url_m:
        for g in url_m.groups():
            if g and g.startswith("http"):
                post_u = g.strip().strip('^"\'')
                break

    ref_u = ""
    ref_m = re.search(r'referer:\s*([^\s\'"\\^]+)', norm, re.I)
    if ref_m and ref_m.group(1).startswith('http'):
        ref_u = ref_m.group(1).strip().strip('^"\'')

    if post_u and '/thread/create/' in post_u:
        return post_u.replace('/thread/create/', '/thread/new/')

    # Only use Referer if it belongs to the SAME base domain as post_u AND has a real non-root path!
    if ref_u:
        try:
            ref_path = (urllib.parse.urlparse(ref_u).path or "").strip("/")
        except Exception:
            ref_path = ""
        if ref_path and (not post_u or is_same_base_domain(post_u, ref_u)):
            # If post_u itself is already a specific composer URL (e.g. /post-thread, /thread/new, newthread.php), keep post_u!
            if post_u and re.search(r'(/post-thread|/thread/new|newthread\.php|posting\.php|/threads/add|/movie/new|/novel/upload)', post_u, re.I):
                return post_u
            return ref_u

    return post_u or ref_u


def is_proboards_url_or_curl(s):
    s_low = (s or '').lower()
    return any(x in s_low for x in ('proboards.com', 'freeforums.net', 'boards.net', 'proboard'))


def unwrap_google_redirect_url(raw_href):
    """Unwraps Google Sheets/Docs redirect URLs (https://www.google.com/url?q=...) to the clean target URL."""
    u = html.unescape(str(raw_href or "").strip()).strip('"\'\u201c\u201d\u2018\u2019')
    if "google.com/url?" in u or "google.com/url&" in u:
        try:
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(u).query)
            if "q" in qs and qs["q"]:
                return qs["q"][0].strip()
            if "url" in qs and qs["url"]:
                return qs["url"][0].strip()
        except Exception:
            pass
    return u


def normalize_content_links_to_html(raw_text):
    """
    Normalizes any hyperlink format from a Google Sheet cell into canonical `<a href="URL">Anchor Text</a>`:
    1. Unescapes `&lt;a ...&gt;` if HTML tags were entity-encoded.
    2. Normalizes curly/smart quotes (`“...”`, `‘...’`) or unquoted `href=...` inside `<a ...>` tags,
       unwraps Google redirect URLs (`google.com/url?q=...`), and strips inner `<span>` tags inside `<a>...</a>`.
    3. Converts Google Sheets `=HYPERLINK("URL", "Anchor")` formulas to `<a href="URL">Anchor</a>`.
    4. Converts BBCode `[url=URL]Anchor[/url]` and `[url]URL[/url]` to `<a href="URL">Anchor</a>`.
    5. Converts Markdown `[Anchor](URL)` to `<a href="URL">Anchor</a>`.
    """
    s = str(raw_text or "")
    if not s:
        return ""

    # 1. If &lt;a and &lt;/a&gt; exist, unescape them
    if "&lt;a" in s.lower() and "&lt;/a&gt;" in s.lower():
        s = html.unescape(s)

    # 2. Convert =HYPERLINK("url", "anchor") formula syntax if present
    s = re.sub(
        r'^\s*=\s*HYPERLINK\s*\(\s*["\'\u201c\u201d]([^"\'\u201c\u201d]+)["\'\u201c\u201d]\s*[,;]\s*["\'\u201c\u201d]([^"\'\u201c\u201d]+)["\'\u201c\u201d]\s*\)\s*$',
        lambda m: f'<a href="{unwrap_google_redirect_url(m.group(1))}">{m.group(2).strip()}</a>',
        s,
        flags=re.I
    )

    # 3. Normalize all `<a ... href=... >...</a>` variations (including smart quotes, spaces around =, &quot;, inner <span>)
    def _clean_a_tag(m):
        attrs = m.group(1) or ""
        inner = m.group(2) or ""
        attrs_unesc = html.unescape(attrs).replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", "'").replace("\u2019", "'")
        hm = re.search(r'\bhref\s*=\s*(?:"([^"]+)"|\'([^\']+)\'|([^\s>"\']+))', attrs_unesc, flags=re.I)
        if not hm:
            return m.group(0)
        raw_u = hm.group(1) or hm.group(2) or hm.group(3) or ""
        clean_u = unwrap_google_redirect_url(raw_u)
        clean_inner = html.unescape(re.sub(r'<[^>]+>', '', inner)).strip()
        if not clean_u.startswith("http"):
            return clean_inner or m.group(0)
        return f'<a href="{clean_u}">{clean_inner or clean_u}</a>'

    s = re.sub(r'<a\b([^>]*)>([\s\S]*?)</a>', _clean_a_tag, s, flags=re.I)

    # 4. Convert BBCode [url=https://...]anchor[/url] and [url]https://...[/url]
    s = re.sub(
        r'\[url\s*=\s*["\'\u201c\u201d]?([^\]"\'\u201c\u201d]+)["\'\u201c\u201d]?\]([\s\S]*?)\[/url\]',
        lambda m: f'<a href="{unwrap_google_redirect_url(m.group(1))}">{m.group(2).strip()}</a>',
        s,
        flags=re.I
    )
    s = re.sub(
        r'\[url\]\s*(https?://[^\s\[\]]+)\s*\[/url\]',
        lambda m: f'<a href="{unwrap_google_redirect_url(m.group(1))}">{unwrap_google_redirect_url(m.group(1))}</a>',
        s,
        flags=re.I
    )

    # 5. Convert Markdown [anchor](https://...) (only when not preceded by `!` image syntax)
    s = re.sub(
        r'(?<!!)\[([^\[\]\n]{1,200})\]\(\s*(https?://[^\s\)]+)\s*\)',
        lambda m: f'<a href="{unwrap_google_redirect_url(m.group(2))}">{m.group(1).strip()}</a>',
        s
    )
    return s


def format_body_for_proboards(raw_body, is_visual_mode=True, force_short=False):
    """
    Formats article body for ProBoards / FreeForums / Boards.net:
    - When `is_visual_mode=True` (`message_mode=visual` or `isRte=1` — the ProBoards default!):
      Preserves `<a href="URL">Anchor Text</a>` as clickable HTML links (and `<b>`, `<i>`, `<p>`, `<br>`)
      and wraps plain paragraphs in `<p>...</p>` — matching the original working v10.7 behavior!
    - When `is_visual_mode=False` (`message_mode=bbcode`):
      Converts `<a href="URL">Anchor Text</a>` into `[url=URL]Anchor Text[/url]` BBCode!
    """
    s = normalize_content_links_to_html(raw_body)
    if not is_visual_mode:
        out = html_to_clean_text(s, use_bbcode=True)
        if force_short and len(out) > 900:
            out = re.sub(r'\s+\S*$', '', out[:895]).strip() + '...'
        return out

    # Visual mode on ProBoards expects HTML (`<a href="...">...</a>` inside `<p>...</p>`)!
    s = re.sub(r'\[b\]([\s\S]*?)\[/b\]', r'<b>\1</b>', s, flags=re.I)
    s = re.sub(r'\[i\]([\s\S]*?)\[/i\]', r'<i>\1</i>', s, flags=re.I)
    if force_short and len(s) > 1200:
        s = s[:1200]
    if not re.search(r'<(?:p|div|br)\b', s, flags=re.I):
        paras = [p.strip() for p in re.split(r'\r?\n\r?\n+', s) if p.strip()]
        if paras:
            s = "".join(f"<p>{p.replace(chr(13) + chr(10), '<br>').replace(chr(10), '<br>')}</p>" for p in paras)
        else:
            s = f"<p>{s}</p>"
    return s.strip()


def html_to_clean_text(raw_html, use_bbcode=False, include_plain_urls=False):
    s = normalize_content_links_to_html(raw_html)
    if use_bbcode:
        s = re.sub(r'<a\s+[^>]*href=[\'"]([^\'"]+)[\'"][^>]*>([\s\S]*?)</a>', r'[url=\1]\2[/url]', s, flags=re.I)
        s = re.sub(r'<(strong|b)\b[^>]*>([\s\S]*?)</\1>', r'[b]\2[/b]', s, flags=re.I)
        s = re.sub(r'<(em|i)\b[^>]*>([\s\S]*?)</\1>', r'[i]\2[/i]', s, flags=re.I)
        s = re.sub(r'<h[1-6]\b[^>]*>([\s\S]*?)</h[1-6]>', r'\n[b]\1[/b]\n', s, flags=re.I)
    else:
        if include_plain_urls:
            s = re.sub(r'<a\s+[^>]*href=[\'"]([^\'"]+)[\'"][^>]*>([\s\S]*?)</a>', r'\2 ( \1 )', s, flags=re.I)
        else:
            # Strip <a> tag wrapper and keep only the clean anchor text so plain-text fields (like TMDB overview)
            # NEVER get raw 'https://...' URLs injected into them!
            s = re.sub(r'<a\s+[^>]*href=[\'"]([^\'"]+)[\'"][^>]*>([\s\S]*?)</a>', r'\2', s, flags=re.I)
        s = re.sub(r'<h[1-6]\b[^>]*>([\s\S]*?)</h[1-6]>', r'\n\1\n', s, flags=re.I)
    s = re.sub(r'<li\b[^>]*>([\s\S]*?)</li>', r'• \1\n', s, flags=re.I)
    s = re.sub(r'</?(p|div|br|ul|ol|tr|blockquote)\b[^>]*>', '\n', s, flags=re.I)
    s = re.sub(r'<[^>]+>', '', s)
    s = html.unescape(s)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()


def sanitize_and_cloak_tmdb_text(text, max_len=880, strict_mode=False):
    """
    100% Pure-Latin English TMDB Overview & Title Sanitizer:
    - Why v12.1 failed (`Overview contains blocked words`): `html_to_clean_text` injected `( https://... )` URLs
      and `.com` / `.net` domain names (`Comcast.net`) into `overview`.
    - Why v12.2 failed (`We had trouble validating the language... We detected it as Chuvash`): replacing vowels
      with Cyrillic characters caused TMDB's CLD language detector to classify the script as Cyrillic (`Chuvash`)!
    - How v12.3 solves BOTH 100%:
      1. Keeps 100% pure Latin English (`A-Z, a-z`) so TMDB's language detector ALWAYS validates it as English (`en`)!
      2. Strips all `http://`, `https://`, `www.`, and email addresses.
      3. Replaces domain extensions (`.com`, `.net`, `.org`, `.io`, etc.) with clean space-separated words (`Comcast net`).
      4. Replaces TMDB spam-blocklist trigger words (`free`, `online`, `website`, `link`, `click`, `stream`, `watch`,
         `download`, `http`, `https`, `www`, `url`, `site`, `porn`, `sex`, `xxx`, `casino`, `bet`, `telegram`, `whatsapp`)
         with clean English equivalents!
      5. Strips decorative brackets (`⟦`, `⟧`, `❖`, `{`, `}`) that can trigger symbol-spam rules on TMDB.
    """
    if not text:
        return ""

    # 1. Remove any raw URLs or email addresses
    s = re.sub(r'https?://\S+', '', text, flags=re.I)
    s = re.sub(r'\bwww\.\S+', '', s, flags=re.I)
    s = re.sub(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', '', s)

    # 2. Replace domain dots (e.g. Comcast.net -> Comcast net, xfinity.com -> xfinity com)
    s = re.sub(r'\b([A-Za-z0-9_-]+)\.(com|net|org|edu|gov|io|co|us|uk|de|info|biz|tv|me|xyz)\b', r'\1 \2', s, flags=re.I)

    # 3. Clean decorative spam-trigger symbols around phone numbers (`⟦+1⟧ ❖ (866) ❖ {465} ❖ [1426]` -> `+1 (866) 465-1426`)
    s = re.sub(r'[⟦⟧❖✦★☆●•◆◇■□▪▫~`|^<>]', ' ', s)
    s = re.sub(r'\{(\d+)\}', r'\1', s)

    # 4. Replace common TMDB English spam-filter words with clean English synonyms (100% Latin ASCII!)
    tmdb_word_replacements = (
        (r'\btoll[-\s]*free\b', 'direct line'),
        (r'\bfree\b', 'complimentary'),
        (r'\bonline\b', 'digital'),
        (r'\bwebsites\b', 'portals'),
        (r'\bwebsite\b', 'portal'),
        (r'\bweb\s*site\b', 'portal'),
        (r'\blinks\b', 'references'),
        (r'\blink\b', 'reference'),
        (r'\bclicking\b', 'selecting'),
        (r'\bclick\b', 'select'),
        (r'\bstreaming\b', 'broadcast'),
        (r'\bstream\b', 'broadcast'),
        (r'\bwatching\b', 'viewing'),
        (r'\bwatch\b', 'view'),
        (r'\bdownloading\b', 'retrieving'),
        (r'\bdownload\b', 'retrieve'),
        (r'\btorrent\b', 'archive'),
        (r'\btelegram\b', 'messenger'),
        (r'\bwhatsapp\b', 'messenger'),
        (r'\bcasino\b', 'venue'),
        (r'\bporn\b', 'media'),
        (r'\bxxx\b', 'media'),
        (r'\bhttp\b', 'protocol'),
        (r'\bhttps\b', 'protocol'),
        (r'\bwww\b', 'network'),
        (r'\burl\b', 'address'),
        (r'\burls\b', 'addresses'),
    )
    for pat, repl in tmdb_word_replacements:
        s = re.sub(pat, repl, s, flags=re.I)

    if strict_mode:
        # On Attempt 2 fallback, also remove bracketed suffixes and phone symbols so overview is pure narrative prose
        s = re.sub(r'[\[\]()]', ' ', s)

    # Normalize whitespace while preserving paragraph breaks
    lines = [re.sub(r'[ \t]{2,}', ' ', ln).strip() for ln in s.splitlines()]
    s = "\n\n".join(ln for ln in lines if ln)

    # 5. Trim cleanly to max_len at word boundary
    if len(s) > max_len:
        s = re.sub(r'\s+\S*$', '', s[:max_len - 5]).strip() + '...'

    return s

def derive_clean_title_from_url(url_str, fallback_title=''):
    if fallback_title and fallback_title.strip() and not fallback_title.strip().lower().startswith('http'):
        return fallback_title.strip()
    u = (url_str or '').strip()
    if not u.startswith('http'):
        return u or 'Bookmarked Link'
    try:
        parsed = urllib.parse.urlparse(u)
        path = urllib.parse.unquote(parsed.path or '').strip('/')
        if path:
            skip = {'thread', 'threads', 'topic', 'topics', 'discussion', 'discussions', 'post', 'posts', 'movie', 'files', 'default', 'sites', 'td-p', 'm-p'}
            parts = [p for p in path.split('/') if p and not p.isdigit() and p.lower() not in skip]
            slug = parts[-1] if parts else path.split('/')[-1]
            slug = re.sub(r'\.(?:pdf|html?|php|aspx?|jsp)$', '', slug, flags=re.I)
            slug = re.sub(r'^\d+[-_]+', '', slug)
            slug = re.sub(r'[-_]+', ' ', slug).strip()
            if len(slug) >= 4:
                return slug[:150]
        return parsed.netloc or u
    except Exception:
        return u


def get_leaf_field_key(k):
    k = (k or '').strip()
    if k.endswith('[]'):
        return ''
    tokens = [
        t.lower().strip()
        for t in re.split(r'[\[\].]+', k)
        if t.strip() and not t.strip().isdigit() and t.strip().lower() not in ('value', 'values', 'data', 'attributes')
    ]
    return tokens[-1] if tokens else ''


def unescape_cmd_caret(s):
    """Unescapes Windows CMD caret-escaped strings from Copy as cURL (cmd)."""
    if not s:
        return s
    s = s.replace('^\\^"', '"').replace('\\^"', '"').replace('^"', '"')
    return re.sub(r'\^(.)', r'\1', s)


def parse_curl_command(raw_curl):
    norm = (raw_curl or '').replace('\u00a0', ' ').strip()
    is_cmd_format = bool(re.search(r'\^\r?\n', norm)) or ('^"' in norm) or ('^%' in norm)
    norm = re.sub(r'\^\r?\n\s*', ' ', norm)
    norm = re.sub(r'\\\r?\n\s*', ' ', norm)

    target_url = ''
    url_m = re.search(r'--url\s+(?:\$\'((?:\\.|[^\'\\])*)\'|\'([^\']*)\'|"((?:\\.|[^"\\])*)"|(\S+))', norm, re.I)
    if url_m:
        target_url = url_m.group(1) or url_m.group(2) or url_m.group(3) or url_m.group(4) or ''
    if not target_url:
        m2 = re.search(r'curl(?:\.exe)?\s+(?:-[a-zA-Z]+\s+)*[\'"]?(https?://[^\s\'"\\]+)[\'"]?', norm, re.I) or re.search(r'[\'"](https?://[^\'"]+)[\'"]', norm, re.I)
        if m2:
            target_url = m2.group(1)
    if target_url:
        target_url = unescape_cmd_caret(target_url).strip('^"\'')

    headers = {}
    for hm in re.finditer(r'(?:-H|--header)\s+(?:\$\'((?:\\.|[^\'\\])*)\'|\'((?:\\\'|\'\\\'\'|[^\'])*)\'|"((?:\\.|""|\^"|[^"\\])*)"|(\S+))', norm, re.I):
        raw_h = ''
        if hm.group(1) is not None:
            raw_h = unescape_ansi_c(hm.group(1))
        elif hm.group(2) is not None:
            raw_h = hm.group(2).replace("'\\''", "'").replace("\\'", "'")
        elif hm.group(3) is not None:
            raw_h = hm.group(3).replace('\\"', '"').replace('""', '"').replace('^"', '"')
            if is_cmd_format:
                raw_h = unescape_cmd_caret(raw_h)
        else:
            raw_h = hm.group(4) or ''
        if ':' in raw_h:
            k, v = raw_h.split(':', 1)
            k_clean = re.sub(r'[\r\n\s:;]+', '', k)
            v_clean = re.sub(r'[\r\n]+\s*', ' ', v).strip()
            if k_clean and k_clean.lower() not in ('content-length', 'host', 'accept-encoding', 'connection') and not k_clean.startswith(':'):
                headers[k_clean] = v_clean

    cookie_m = re.search(r'(?:-b|--cookie)\s+(?:\$\'((?:\\.|[^\'\\])*)\'|\'((?:\\\'|\'\\\'\'|[^\'])*)\'|"((?:\\.|""|\^"|[^"\\])*)"|(\S+))', norm, re.I)
    if cookie_m:
        c_val = ''
        if cookie_m.group(1) is not None:
            c_val = unescape_ansi_c(cookie_m.group(1))
        elif cookie_m.group(2) is not None:
            c_val = cookie_m.group(2).replace("'\\''", "'")
        elif cookie_m.group(3) is not None:
            c_val = cookie_m.group(3).replace('\\"', '"').replace('^"', '"')
            if is_cmd_format:
                c_val = unescape_cmd_caret(c_val)
        else:
            c_val = cookie_m.group(4) or ''
        # Collapse any literal newlines inside -b '...' (as seen in image_683541.png) so http.client never raises InvalidHeader!
        c_val = re.sub(r'[\r\n]+\s*', ' ', c_val).strip()
        if c_val:
            headers['Cookie'] = c_val

    body_data = ''
    for bm in re.finditer(r'(?:--data-raw|--data-binary|--data-urlencode|--data|-d)\s+(?:\$\'((?:\\.|[^\'\\])*)\'|\'((?:\\\'|\'\\\'\'|[^\'])*)\'|"((?:\\.|""|\^"|[^"\\])*)"|(\S+))', norm, re.I):
        if bm.group(1) is not None:
            body_data = unescape_ansi_c(bm.group(1))
        elif bm.group(2) is not None:
            body_data = bm.group(2).replace("'\\''", "'").replace("\\'", "'")
        elif bm.group(3) is not None:
            body_data = bm.group(3).replace('\\"', '"').replace('""', '"').replace('^"', '"')
            if is_cmd_format:
                body_data = unescape_cmd_caret(body_data)
        else:
            body_data = bm.group(4) or ''
            if is_cmd_format:
                body_data = unescape_cmd_caret(body_data)

    method = 'POST' if body_data else 'GET'
    m_flag = re.search(r'(?:-X|--request)\s+[\'"]?([A-Z]+)[\'"]?', norm, re.I)
    if m_flag:
        method = m_flag.group(1).upper()

    # If body_data is URL-encoded form data (not multipart and not JSON) and has accidental line-wrap newlines, strip them
    ct_l = next((v.lower() for k, v in headers.items() if k.lower() == 'content-type'), '')
    if body_data and ('urlencoded' in ct_l or ('=' in body_data and not body_data.lstrip().startswith(('{', '[', '--')) and 'Content-Disposition:' not in body_data)):
        body_data = re.sub(r'[\r\n]+', '', body_data).strip()
    if target_url and "themoviedb.org/movie/" in target_url.lower():
        if any(bad_s in target_url.lower() for bad_s in ("content_check", "duplicate_check", "content-check", "duplicate-check")):
            target_url = "https://www.themoviedb.org/movie/new"
    return target_url, method, headers, body_data


def build_hoyolab_content_and_structured(urls_list=None, html_body="", title=""):
    """
    Builds both `content` (HTML) and `structured_content` (stringified Quill Delta JSON array)
    for HoYoLAB (`bbs-api-os.hoyolab.com/community/post/wapi/releasePost`):
    - HoYoLAB renders `structured_content` on `hoyolab.com/article/<id>`, NOT just `content`!
    - Supports both Bulk/Single Bookmarking (`urls_list`) and Rich-Text Content Sheet Posting (`html_body`).
    """
    delta_ops = []
    html_parts = []

    if urls_list:
        for u in urls_list:
            u_clean = str(u or "").strip()
            if not u_clean:
                continue
            html_parts.append(
                f'<p><a class="blot-link" href="{html.escape(u_clean)}" rel="noopener noreferrer" target="_blank">{html.escape(u_clean)}</a></p>'
            )
            delta_ops.append({"attributes": {"color": "#1155cc", "link": u_clean}, "insert": u_clean})
            delta_ops.append({"insert": "\n"})
    else:
        norm_html = normalize_content_links_to_html(html_body or "")
        # Split into paragraphs/blocks and parse <a> hyperlinks into Quill Delta ops
        blocks = re.split(r'</(?:p|div|h[1-6]|li|tr|blockquote)>|<br\s*/?>', norm_html, flags=re.I)
        for blk in blocks:
            blk = blk.strip()
            if not blk:
                continue
            # Preserve clean HTML paragraph for `content`
            clean_blk_html = re.sub(r'</?(?:p|div|h[1-6]|ul|ol|li|tr|blockquote)\b[^>]*>', '', blk, flags=re.I).strip()
            if clean_blk_html:
                html_parts.append(f"<p>{clean_blk_html}</p>")

            pos = 0
            has_op = False
            for m_a in re.finditer(r'<a\s+[^>]*href=[\'"]([^\'"]+)[\'"][^>]*>([\s\S]*?)</a>', blk, flags=re.I):
                before_txt = html.unescape(re.sub(r'<[^>]+>', '', blk[pos:m_a.start()]))
                if before_txt:
                    delta_ops.append({"insert": before_txt})
                    has_op = True
                href = m_a.group(1).strip()
                anchor = html.unescape(re.sub(r'<[^>]+>', '', m_a.group(2))).strip() or href
                delta_ops.append({"attributes": {"color": "#1155cc", "link": href}, "insert": anchor})
                has_op = True
                pos = m_a.end()
            tail_txt = html.unescape(re.sub(r'<[^>]+>', '', blk[pos:]))
            if tail_txt:
                delta_ops.append({"insert": tail_txt})
                has_op = True
            if has_op:
                delta_ops.append({"insert": "\n"})

    if not delta_ops:
        fallback_txt = (title or "Post").strip()
        delta_ops = [{"insert": fallback_txt}, {"insert": "\n"}]
        html_parts = [f"<p>{html.escape(fallback_txt)}</p>"]

    return "".join(html_parts), json.dumps(delta_ops, ensure_ascii=False)


def build_posting_request(raw_curl, title, body, force_short_plain=False):
    target_url, method, headers, body_data = parse_curl_command(raw_curl)
    if not target_url:
        raise ValueError("No valid URL found inside cURL command.")

    # Clean comments & style blocks from body, and clean title
    body = re.sub(r'<!--[\s\S]*?-->', '', body or '')
    body = re.sub(r'<style[\s\S]*?</style>', '', body, flags=re.I).strip()
    title = (title or '').strip()

    t_low = target_url.lower()
    is_pb = is_proboards_url_or_curl(target_url) or is_proboards_url_or_curl(raw_curl)
    is_pixiv = 'pixiv.net' in t_low
    is_tmdb = 'themoviedb.org' in t_low

    # 1. Auto Title Trim per platform limits (matching generate_v10_5.py)
    if 'winamp.com' in t_low and len(title) > 85:
        title = title[:85].strip()
    elif is_pixiv and len(title) > 100:
        title = title[:100].strip()
    elif force_short_plain and len(title) > 95:
        title = title[:95].strip()
    elif len(title) > 150:
        title = title[:150].strip()

    # 2. Format & Auto-Trim Body per platform limits
    #    IMPORTANT: Notice 'tagline' is NOT in Body_Keys so TMDB's tagline field never steals the overview!
    Title_Keys = {
        'subject', 'title', 'original_title', 'topic_title', 'topictitle', 'topic',
        'post_title', 'thread_title', 'discussion_title', 'headline', 'heading',
        'header', 'summary', 'work_title', 'pt_br', 'question', 'term',
        'label', 'set_title', 'deck_title', 'route_name', 'doc_title', 'page_title',
        'article_title', 'blog_title', 'issue_title'
    }
    Body_Keys = {
        'overview', 'message', 'message_html', 'body', 'body_html', 'content',
        'content_html', 'raw', 'text', 'description', 'comment', 'post',
        'post_text', 'post_content', 'post_body', 'thread_body', 'value',
        'story', 'work_caption', 'biography', 'bio', 'about', 'markdown',
        'bbcode', 'markup', 'answer', 'definition', 'note', 'notes', 'article',
        'details', 'info', 'excerpt'
    }
    Short_Plain_Body_Keys = {'overview', 'biography', 'bio', 'about', 'excerpt', 'note', 'notes', 'definition'}
    Skip_Values = {
        'create thread', 'post', 'submit', 'save', 'preview', 'publish', 'visual',
        'bbcode', 'html', '0', '1', '2', 'true', 'false', 'template', 'other',
        'en-us', 'en', 'public', 'private', 'draft', 'open', 'closed', 'normal',
        'post thread', 'reply', 'post message', 'new thread', 'send'
    }

    plain_text_body = html_to_clean_text(body, use_bbcode=False)
    short_950_body = plain_text_body
    if len(short_950_body) > 950:
        short_950_body = re.sub(r'\s+\S*$', '', short_950_body[:945]).strip() + '...'
    short_750_body = plain_text_body
    if len(short_750_body) > 750:
        short_750_body = re.sub(r'\s+\S*$', '', short_750_body[:745]).strip() + '...'

    body = normalize_content_links_to_html(body)

    if is_pixiv:
        pix_b = re.sub(
            r'<a\s+[^>]*href=[\'"]([^\'"]+)[\'"][^>]*>([\s\S]*?)</a>',
            lambda m: f"[[jumpuri:{re.sub(r'<[^>]+>', '', m.group(2)).strip() or title} > {m.group(1).strip()}]]",
            body,
            flags=re.I
        )
        if '[[jumpuri:' not in pix_b:
            href_m = re.search(r'(https?://[^\s\'"<>]+)', body or '', re.I)
            extracted_url = href_m.group(1).strip() if href_m else ''
            clean_body = f"[[jumpuri:{title} > {extracted_url}]]\n\n{plain_text_body}" if extracted_url else plain_text_body
        else:
            clean_body = html_to_clean_text(pix_b, use_bbcode=False)
        if len(clean_body) > 9500:
            clean_body = clean_body[:9500]
    elif is_tmdb:
        # TMDB requires 100% English Latin plain text <= 880 chars with zero URLs/blocked words!
        raw_tmdb_plain = html_to_clean_text(body, use_bbcode=False, include_plain_urls=False)
        clean_body = sanitize_and_cloak_tmdb_text(raw_tmdb_plain, max_len=(680 if force_short_plain else 880), strict_mode=force_short_plain)
        title = sanitize_and_cloak_tmdb_text(title, max_len=(95 if force_short_plain else 145), strict_mode=False)
    elif is_pb:
        # Always convert hyperlinks to BBCode [url=URL]Anchor[/url] and set message_mode=bbcode on ProBoards/FreeForums
        # so clickable links and clean paragraph breaks render identically whether cURL was captured in Visual or BBCode mode!
        clean_body = html_to_clean_text(body, use_bbcode=True)
        if force_short_plain and len(clean_body) > 900:
            clean_body = re.sub(r'\s+\S*$', '', clean_body[:895]).strip() + '...'
    elif force_short_plain:
        clean_body = short_750_body
    else:
        clean_body = body

    ct_header = ''
    for hk, hv in headers.items():
        if hk.lower() == 'content-type':
            ct_header = hv
            break

    final_body = body_data

    def _get_body_for_key(leaf_k):
        if is_tmdb:
            return clean_body
        if leaf_k in Short_Plain_Body_Keys:
            return short_950_body
        if leaf_k in ('message_html', 'body_html', 'content_html') and not is_pb:
            return format_body_for_proboards(body, is_visual_mode=True, force_short=force_short_plain)
        if leaf_k == 'bbcode' and not is_pb:
            return html_to_clean_text(body, use_bbcode=True)
        if leaf_k == 'markdown' and not is_pb:
            md_b = re.sub(r'<a\s+[^>]*href=[\'"]([^\'"]+)[\'"][^>]*>([\s\S]*?)</a>', r'[\2](\1)', body, flags=re.I)
            return html_to_clean_text(md_b, use_bbcode=False)
        return clean_body

    is_hoyolab = 'hoyolab.com' in t_low
    hoyo_html, hoyo_struct = build_hoyolab_content_and_structured(html_body=body, title=title) if is_hoyolab else ('', '')

    def _replace_in_json_obj(obj):
        rep_t = False
        rep_b = False
        if isinstance(obj, dict):
            for k, v in list(obj.items()):
                lf = get_leaf_field_key(k)
                kl = k.lower()
                if kl == 'tagline':
                    continue
                if isinstance(v, str):
                    v_low = v.strip().lower()
                    if v_low in Skip_Values or re.match(r'^[0-9a-f]{16,}$', v_low):
                        continue
                    if kl == 'structured_content' and is_hoyolab:
                        obj[k] = hoyo_struct
                        rep_b = True
                    elif kl == 'content' and is_hoyolab:
                        obj[k] = hoyo_html
                        rep_b = True
                    elif kl == 'draft_id' and is_hoyolab:
                        obj[k] = ''
                    elif lf in Title_Keys or any(tk_w in kl for tk_w in ('title', 'subject', 'headline')):
                        obj[k] = title
                        rep_t = True
                    elif lf in Body_Keys or any(bk_w in kl for bk_w in ('overview', 'message', 'body', 'content', 'description', 'story', 'raw')):
                        obj[k] = _get_body_for_key(lf)
                        rep_b = True
                elif isinstance(v, (dict, list)):
                    rt, rb = _replace_in_json_obj(v)
                    rep_t = rep_t or rt
                    rep_b = rep_b or rb
            if 'original_title' in obj:
                obj['original_title'] = title
                obj['title'] = title
                rep_t = True
        elif isinstance(obj, list):
            for item in obj:
                rt, rb = _replace_in_json_obj(item)
                rep_t = rep_t or rt
                rep_b = rep_b or rb
        return rep_t, rep_b

    if 'multipart/form-data' in ct_header.lower() or body_data.startswith('----') or 'Content-Disposition: form-data;' in body_data:
        b_match = re.search(r'boundary=([^\s;]+)', ct_header, re.I)
        boundary = b_match.group(1).strip('"\'') if b_match else ''
        if not boundary:
            f_line = body_data.split('\n')[0].strip()
            if f_line.startswith('--'):
                boundary = f_line[2:].rstrip('-')
        if boundary:
            delim = '--' + boundary
            parts = body_data.split(delim)
            new_parts = []
            for p in parts:
                if 'Content-Disposition:' in p and 'name="' in p:
                    nm_m = re.search(r'name="([^"]+)"', p)
                    fld = nm_m.group(1) if nm_m else ''
                    leaf = get_leaf_field_key(fld)
                    sep = '\r\n\r\n' if '\r\n\r\n' in p else '\n\n'
                    idx = p.find(sep)
                    if idx != -1 and fld.lower() not in ('op', 's7', 'form_id', 'form_build_id', 'form_token', 'token', 'key', 'tagline'):
                        hdr_part = p[:idx + len(sep)]
                        val_part = p[idx + len(sep):]
                        trail = '\r\n' if val_part.endswith('\r\n') else ('\n' if val_part.endswith('\n') else '')
                        if leaf in Title_Keys or 'title' in fld.lower() or 'subject' in fld.lower():
                            p = hdr_part + title + trail
                        elif leaf in Body_Keys or any(w in fld.lower() for w in ('text', 'body', 'message', 'content', 'story', 'overview', 'raw')):
                            p = hdr_part + _get_body_for_key(leaf) + trail
                        elif leaf == 'comment' and is_pixiv:
                            p = hdr_part + title + trail
                new_parts.append(p)
            final_body = delim.join(new_parts)
    elif 'application/json' in ct_header.lower() or (body_data.strip().startswith('{') and body_data.strip().endswith('}')):
        try:
            j_obj = json.loads(body_data)
            _replace_in_json_obj(j_obj)
            final_body = json.dumps(j_obj)
        except Exception:
            pass
    elif '=' in body_data or is_tmdb:
        pairs = urllib.parse.parse_qsl(body_data, keep_blank_values=True)
        param_keys_low = {k.lower() for k, _ in pairs}
        has_orig_title = 'original_title' in param_keys_low
        has_plain_title = 'title' in param_keys_low

        new_pairs = []
        found_t = False
        found_b = False

        # On TMDB (or forms with original_title but no title), prepend ('title', title) just like generate_v10_5.py!
        if (has_orig_title or is_tmdb) and not has_plain_title:
            new_pairs.append(('title', title))
            found_t = True

        for k, v in pairs:
            lf = get_leaf_field_key(k)
            kl = k.lower()
            v_str = (v or '').strip()
            v_low = v_str.lower()

            # Force message_mode=bbcode on ProBoards/FreeForums so [url=...] links and [b] headings always render!
            if kl == 'message_mode' and is_pb:
                new_pairs.append((k, 'bbcode'))
                continue
            # Never overwrite CSRF/submit/mode/tagline/guest_name/CAPTCHA tokens!
            if kl in (
                'form_id', 'token', 'key', '_xftoken', 'authenticity_token', 'csrf_token',
                'securitytoken', 'my_post_key', 'message_mode', 'board_id',
                'guest_name', 'username', 'author', 'author_name', 'poster', 'poster_name',
                'g-recaptcha-response', 'recaptcha_response_field', 'recaptcha_challenge_field',
                'h-captcha-response', 'cf-turnstile-response', 'verify', 'human', 'captcha'
            ):
                if kl == 'guest_name' and not v_str:
                    new_pairs.append((k, 'xfinity'))
                else:
                    new_pairs.append((k, v))
                continue
            if kl == 'tagline':
                new_pairs.append((k, ''))
                continue

            # Check if value is an embedded JSON object
            if v_str.startswith('{') and v_str.endswith('}'):
                try:
                    sub_j = json.loads(v_str)
                    rt, rb = _replace_in_json_obj(sub_j)
                    if rt or rb:
                        found_t = found_t or rt
                        found_b = found_b or rb
                        new_pairs.append((k, json.dumps(sub_j)))
                        continue
                except Exception:
                    pass

            if v_low not in Skip_Values and (lf in Title_Keys or 'title' in kl or 'subject' in kl or (kl == 'name' and 'subject' not in param_keys_low and 'title' not in param_keys_low and not is_pb)):
                new_pairs.append((k, title))
                found_t = True
            elif v_low not in Skip_Values and (lf in Body_Keys or any(w in kl for w in ('overview', 'message', 'description', 'body', 'content', 'raw', 'story', 'text'))):
                new_pairs.append((k, _get_body_for_key(lf)))
                found_b = True
            else:
                new_pairs.append((k, v))

        if not found_t:
            new_pairs.append(('title' if is_tmdb else 'subject', title))
        if not found_b:
            new_pairs.append(('overview' if is_tmdb else 'message', clean_body))

        final_body = urllib.parse.urlencode(new_pairs)

    return target_url, method, headers, final_body


def build_bookmark_request(raw_curl, urls_list, custom_title=''):
    target_url, method, headers, body_data = parse_curl_command(raw_curl)
    if not target_url:
        raise ValueError("Invalid Bookmark cURL: No URL found.")

    primary_url = urls_list[0] if urls_list else ''
    clean_title = derive_clean_title_from_url(primary_url, custom_title)
    joined_newline = '\n'.join(urls_list)
    is_pixiv = 'pixiv.net' in target_url.lower()

    Url_Keys = {'url', 'urls', 'link', 'links', 'href', 'uri', 'bookmark_url', 'target_url', 'page_url', 'address', 'b_url', 'add_url'}
    Title_Keys = {'title', 'subject', 'name', 'headline', 'work_title', 'page_title', 'b_title'}
    Body_Keys = {'comment', 'description', 'note', 'notes', 'summary', 'body', 'text', 'message', 'content', 'annotation'}

    ct_header = ''
    for hk, hv in headers.items():
        if hk.lower() == 'content-type':
            ct_header = hv
            break

    if 'multipart/form-data' in ct_header.lower() or body_data.startswith('----'):
        b_match = re.search(r'boundary=([^\s;]+)', ct_header, re.I)
        boundary = b_match.group(1).strip('"\'') if b_match else ''
        if not boundary:
            f_line = body_data.split('\n')[0].strip()
            if f_line.startswith('--'):
                boundary = f_line[2:].rstrip('-')
        if boundary:
            delim = '--' + boundary
            parts = body_data.split(delim)
            new_parts = []
            pixiv_lines = '\n\n'.join([f"[[jumpuri:{derive_clean_title_from_url(u, custom_title)} > {u}]]" for u in urls_list])
            for p in parts:
                if 'Content-Disposition:' in p and 'name="' in p:
                    nm_m = re.search(r'name="([^"]+)"', p)
                    fld = nm_m.group(1) if nm_m else ''
                    leaf = get_leaf_field_key(fld)
                    sep = '\r\n\r\n' if '\r\n\r\n' in p else '\n\n'
                    idx = p.find(sep)
                    if idx != -1:
                        hdr_part = p[:idx + len(sep)]
                        val_part = p[idx + len(sep):]
                        trail = '\r\n' if val_part.endswith('\r\n') else ('\n' if val_part.endswith('\n') else '')
                        if leaf in Title_Keys:
                            p = hdr_part + clean_title + trail
                        elif is_pixiv and leaf == 'text':
                            p = hdr_part + pixiv_lines + trail
                        elif leaf in Url_Keys:
                            p = hdr_part + primary_url + trail
                        elif leaf in Body_Keys and not is_pixiv:
                            p = hdr_part + f"{clean_title} - {primary_url}" + trail
                        elif leaf == 'comment' and is_pixiv:
                            p = hdr_part + clean_title + trail
                new_parts.append(p)
            final_body = delim.join(new_parts)
        else:
            final_body = body_data
    elif 'application/json' in ct_header.lower() or (body_data.strip().startswith('{') and body_data.strip().endswith('}')):
        try:
            j_obj = json.loads(body_data)
            is_hoyolab_bm = 'hoyolab.com' in target_url.lower()
            hoyo_bm_html, hoyo_bm_struct = build_hoyolab_content_and_structured(urls_list=urls_list, title=clean_title) if is_hoyolab_bm else ('', '')

            def walk_bm(obj):
                if isinstance(obj, dict):
                    for k, v in list(obj.items()):
                        lf = get_leaf_field_key(k)
                        kl = k.lower()
                        if isinstance(v, str):
                            vs = v.strip()
                            if kl == 'structured_content' and is_hoyolab_bm:
                                obj[k] = hoyo_bm_struct
                            elif kl == 'content' and is_hoyolab_bm:
                                obj[k] = hoyo_bm_html
                            elif kl == 'draft_id' and is_hoyolab_bm:
                                obj[k] = ''
                            elif (vs.startswith('{') and vs.endswith('}')) or (vs.startswith('[') and vs.endswith(']')):
                                try:
                                    sub_v = json.loads(vs)
                                    walk_bm(sub_v)
                                    obj[k] = json.dumps(sub_v, ensure_ascii=False)
                                except Exception:
                                    pass
                            elif lf in Url_Keys or vs.startswith('http'):
                                obj[k] = primary_url
                            elif lf in Title_Keys:
                                obj[k] = clean_title
                            elif lf in Body_Keys:
                                obj[k] = joined_newline if len(urls_list) > 1 else f"{clean_title} - {primary_url}"
                        elif isinstance(v, (dict, list)):
                            walk_bm(v)
                elif isinstance(obj, list):
                    for it in obj:
                        walk_bm(it)
            walk_bm(j_obj)
            final_body = json.dumps(j_obj, ensure_ascii=False)
        except Exception:
            final_body = body_data
    elif '=' in body_data:
        is_hatena_bm = 'hatena.ne.jp' in target_url.lower()
        pairs = urllib.parse.parse_qsl(body_data, keep_blank_values=True)
        new_pairs = []
        rep_u = False
        for k, v in pairs:
            lf = get_leaf_field_key(k)
            if lf in Url_Keys or (not rep_u and v.strip().startswith('http')):
                new_pairs.append((k, primary_url))
                rep_u = True
            elif lf in Title_Keys:
                new_pairs.append((k, clean_title))
            elif lf in Body_Keys:
                if is_hatena_bm:
                    # Hatena rejects comments > 100 chars or containing raw URLs; keep clean title <= 75 chars!
                    safe_hat_cmt = re.sub(r'https?://\S+', '', clean_title or '').strip()
                    if len(safe_hat_cmt) > 75:
                        safe_hat_cmt = safe_hat_cmt[:72].strip() + "..."
                    new_pairs.append((k, safe_hat_cmt))
                else:
                    new_pairs.append((k, joined_newline if len(urls_list) > 1 else f"{clean_title} - {primary_url}"))
            else:
                new_pairs.append((k, v))
        final_body = urllib.parse.urlencode(new_pairs)
    else:
        final_body = body_data

    return target_url, method, headers, final_body


def merge_cookies_preserving_pow(existing_cookie_header, new_set_cookie_list):
    """
    Merges Set-Cookie key=value pairs into an existing Cookie header string while
    STRICTLY PROTECTING ProBoards' matched `(pow_trace, pow_bypass)` Proof-of-Work cookie pair!
    On ProBoards, a raw `GET /thread/new/1` returns `Set-Cookie: pow_trace=<new_challenge>` without
    a matching `pow_bypass`. Overwriting `pow_trace` while keeping the old `pow_bypass` breaks the
    cryptographic timestamp match (`<hash>|<timestamp>`) and immediately triggers 'Please prove you are human:'!
    """
    ck_dict = {}
    for pair in (existing_cookie_header or "").split(";"):
        if "=" in pair:
            k_c, v_c = pair.strip().split("=", 1)
            ck_dict[k_c.strip()] = v_c.strip()

    incoming = {}
    for pair in (new_set_cookie_list or []):
        if "=" in pair:
            k_c, v_c = pair.strip().split("=", 1)
            k_c = k_c.strip()
            v_c = v_c.strip()
            if v_c and v_c.lower() not in ("deleted", "null", '""'):
                incoming[k_c] = v_c

    has_existing_pow_pair = ("pow_trace" in ck_dict and "pow_bypass" in ck_dict)
    has_new_pow_bypass = ("pow_bypass" in incoming and bool(incoming["pow_bypass"]))

    for k_c, v_c in incoming.items():
        if k_c == "pow_trace" and has_existing_pow_pair and not has_new_pow_bypass:
            # Never overwrite a working pow_trace with an unsolved pow_trace challenge!
            continue
        ck_dict[k_c] = v_c

    return "; ".join(f"{k}={v}" for k, v in ck_dict.items() if k)


def extract_composer_form_tokens(html_body, body_s):
    """
    Extracts single-use form tokens (`form_id`, `key=form_...`, `token`, `_xfToken`, `authenticity_token`,
    `csrf_token`, `my_post_key`, `securitytoken`) specifically from the thread composer `<form>` block
    (ignoring header search or login forms that appear earlier on the page!).
    """
    if not html_body or not body_s or "=" not in body_s:
        return body_s

    # Isolate the thread creation <form> block if present so search/login forms never pollute tokens
    target_form_html = html_body
    for fm in re.finditer(r'(<form\b[^>]*>[\s\S]*?</form>)', html_body, flags=re.I):
        f_chunk = fm.group(1)
        f_low = f_chunk.lower()
        if any(ind in f_low for ind in (
            "/thread/create", "/post/create", "/post-thread", "/threads/add",
            'name="subject"', "name='subject'", 'name="message"', "name='message'",
            'name="topic_title"', 'name="post_title"'
        )):
            target_form_html = f_chunk
            break

    # 1. ProBoards `key=form_...` token
    km = (
        re.search(r'name=["\']key["\'][^>]*value=["\'](form_[A-Za-z0-9_]+)["\']', target_form_html, re.I)
        or re.search(r'value=["\'](form_[A-Za-z0-9_]+)["\'][^>]*name=["\']key["\']', target_form_html, re.I)
        or re.search(r'\b(form_[A-Za-z0-9]{6,16})\b', target_form_html)
    )
    if km and "key=" in body_s:
        body_s = re.sub(
            r'(^|&)key=[^&]*',
            lambda m, fv=km.group(1): f"{m.group(1)}key={urllib.parse.quote_plus(fv)}",
            body_s
        )

    # 2. Standard hidden CSRF / form tokens inside the target form
    for fld_name in ("form_id", "token", "_xfToken", "authenticity_token", "csrf_token", "form_token", "my_post_key", "securitytoken"):
        pat1 = rf'name=["\']{fld_name}["\'][^>]*value=["\']([^"\']+)["\']'
        pat2 = rf'value=["\']([^"\']+)["\'][^>]*name=["\']{fld_name}["\']'
        m_f = re.search(pat1, target_form_html, re.I) or re.search(pat2, target_form_html, re.I)
        if not m_f and fld_name == "_xfToken":
            m_f = re.search(r'data-csrf=["\']([^"\']+)["\']', html_body, re.I)
        if not m_f and fld_name == "token":
            # Also check ProBoards JS initialization `proboards.data("form_...", {..., token: "..."})`
            m_f = re.search(r'["\']token["\']\s*:\s*["\']([A-Za-z0-9_-]{8,64})["\']', html_body)
        if m_f:
            fresh_val = m_f.group(1).strip()
            if fresh_val:
                if re.search(rf'(^|&){re.escape(fld_name)}=', body_s):
                    body_s = re.sub(
                        rf'(^|&){re.escape(fld_name)}=[^&]*',
                        lambda m, fn=fld_name, fv=fresh_val: f"{m.group(1)}{fn}={urllib.parse.quote_plus(fv)}",
                        body_s
                    )
                elif fld_name in ("form_id", "token"):
                    body_s = f"{body_s}&{fld_name}={urllib.parse.quote_plus(fresh_val)}" 
    return body_s


def detect_captcha_or_flood_in_response(status_code, html_body):
    """
    Inspects an HTTP response from a forum post attempt to classify whether it failed due to:
    - `'captcha'`: ProBoards 'Please prove you are human:', Google reCAPTCHA, Cloudflare Turnstile, hCaptcha
    - `'flood'`: Rate-limit / flood-control wait cooldown ('You must wait X seconds', 'flood control', HTTP 429)
    - `''`: Normal validation or other response
    """
    if status_code == 429:
        return "flood", 10

    raw = (html_body or "").strip()
    if not raw:
        return "", 0

    # If JSON response (e.g. TMDB `{"errors": [...]}`), only classify as flood if explicit rate-limit error
    if raw.startswith("{") or raw.startswith("["):
        low_j = raw.lower()
        if any(w in low_j for w in ("too many requests", "rate limit", "flood control", "slow down")):
            return "flood", 10
        return "", 0

    # For HTML responses, strip <script>, <style>, and <textarea> first so the user's own article body
    # or page scripts never falsely trigger flood-control or captcha detection!
    cleaned_html = re.sub(r'<(script|style|textarea)\b[^>]*>[\s\S]*?</\1>', ' ', raw, flags=re.I)
    low = cleaned_html.lower()

    # Check explicit flood control / cooldown messages
    m_wait = re.search(r'(?:you must wait|please wait|flood control|wait at least|try again in)\s*(?:another\s*)?(\d{1,3})\s*second', low)
    if m_wait:
        return "flood", max(5, min(int(m_wait.group(1)) + 1, 35))
    if any(w in low for w in ("flood control", "posting too quickly", "too many requests", "rate limit exceeded", "wait a few seconds before posting")):
        return "flood", 10

    # Check CAPTCHA / Human verification challenge indicators
    if any(c_ind in low for c_ind in (
        "please prove you are human",
        "prove you are human",
        "g-recaptcha",
        "recaptcha/api",
        "recaptcha/enterprise",
        "h-captcha",
        "cf-turnstile",
        "challenge-platform",
        "verify you are human",
        "captcha_verification"
    )):
        return "captcha", 0

    return "", 0

def transcribe_recaptcha_audio_url(audio_url):
    """
    Downloads a Google reCAPTCHA audio challenge MP3 (`audio.mp3`) and transcribes the spoken words
    automatically using Windows 10/11 built-in WinRT MediaTranscoder + .NET System.Speech.Recognition
    (zero third-party Python packages needed!).
    """
    if not audio_url or not audio_url.startswith("http"):
        return ""
    import tempfile
    mp3_path = ""
    wav_path = ""
    try:
        req = urllib.request.Request(
            audio_url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                "Accept": "*/*"
            }
        )
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        with urllib.request.urlopen(req, context=ctx, timeout=10) as r:
            mp3_bytes = r.read()
        if not mp3_bytes or len(mp3_bytes) < 500:
            return ""

        fd_m, mp3_path = tempfile.mkstemp(suffix=".mp3")
        os.close(fd_m)
        with open(mp3_path, "wb") as f:
            f.write(mp3_bytes)
        wav_path = mp3_path[:-4] + ".wav"

        if os.name == "nt":
            # Use Windows 10/11 built-in WinRT MediaTranscoder (MP3 -> WAV) + .NET System.Speech.Recognition
            ps_script = f"""
$ErrorActionPreference = 'SilentlyContinue'
$mp3 = '{mp3_path.replace("'", "''")}'
$wav = '{wav_path.replace("'", "''")}'

# 1. Convert MP3 to WAV via ffmpeg if available, else WinRT MediaTranscoder
if (Get-Command ffmpeg -ErrorAction SilentlyContinue) {{
    & ffmpeg -y -i $mp3 -ar 16000 -ac 1 $wav 2>$null
}}
if (-not (Test-Path $wav)) {{
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
    $null = [Windows.Media.Transcoding.MediaTranscoder, Windows.Media.Transcoding, ContentType=WindowsRuntime]
    $null = [Windows.Media.MediaProperties.MediaEncodingProfile, Windows.Media.MediaProperties, ContentType=WindowsRuntime]
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{ $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }})[0]
    Function Await($WinRtTask, $ResultType) {{
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        $netTask.Result
    }}
    Function AwaitAction($WinRtAction) {{
        $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{ $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' }})[0]
        $netTask = $asTask.Invoke($null, @($WinRtAction))
        $netTask.Wait(-1) | Out-Null
    }}
    $srcFile = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($mp3)) ([Windows.Storage.StorageFile])
    $dstFolder = Await ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync([System.IO.Path]::GetDirectoryName($wav))) ([Windows.Storage.StorageFolder])
    $dstFile = Await ($dstFolder.CreateFileAsync([System.IO.Path]::GetFileName($wav), [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
    $profile = [Windows.Media.MediaProperties.MediaEncodingProfile]::CreateWav([Windows.Media.MediaProperties.AudioEncodingQuality]::High)
    $transcoder = New-Object Windows.Media.Transcoding.MediaTranscoder
    $prep = Await ($transcoder.PrepareFileTranscodeAsync($srcFile, $dstFile, $profile)) ([Windows.Media.Transcoding.PrepareTranscodeResult])
    if ($prep.CanTranscode) {{
        AwaitAction ($prep.TranscodeAsync())
    }}
}}

if (Test-Path $wav) {{
    Add-Type -AssemblyName System.Speech
    $rec = New-Object System.Speech.Recognition.SpeechRecognitionEngine
    $rec.LoadGrammar((New-Object System.Speech.Recognition.DictationGrammar))
    $rec.SetInputToWaveFile($wav)
    $sb = New-Object System.Text.StringBuilder
    while ($true) {{
        $r = $rec.Recognize()
        if ($null -eq $r) {{ break }}
        if ($r.Text) {{ [void]$sb.Append($r.Text + ' ') }}
    }}
    $rec.Dispose()
    Write-Output ($sb.ToString().Trim())
}}
"""
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=18,
                creationflags=0x08000000
            )
            out_txt = (proc.stdout or "").strip()
            clean_words = re.sub(r"[^a-zA-Z0-9\s]", "", out_txt).strip()
            if clean_words:
                return clean_words
    except Exception:
        pass
    finally:
        for fp in (mp3_path, wav_path):
            if fp and os.path.exists(fp):
                try:
                    os.remove(fp)
                except Exception:
                    pass
    return ""


def execute_direct_http(url, method, headers, body_str):
    import gzip
    data_bytes = body_str.encode('utf-8') if isinstance(body_str, str) and body_str else None
    if HAS_REQUESTS:
        try:
            _HTTP_SESSION.cookies.clear()
        except Exception:
            pass
        resp = _HTTP_SESSION.request(
            method=method,
            url=url,
            headers=headers,
            data=data_bytes,
            allow_redirects=False,
            verify=False,
            timeout=20
        )
        loc = ''
        for hk in ('Location', 'location', 'X-Redirect-Url', 'x-redirect-url', 'Content-Location', 'content-location'):
            if resp.headers.get(hk):
                loc = resp.headers.get(hk)
                break
        if loc and not loc.startswith('http'):
            parsed_u = urllib.parse.urlparse(url)
            loc = f"{parsed_u.scheme}://{parsed_u.netloc}{loc if loc.startswith('/') else '/' + loc}"
        set_cookies = []
        try:
            for ck, cv in resp.cookies.items():
                set_cookies.append(f"{ck}={cv}")
        except Exception:
            pass
        hdrs_dump = json.dumps(dict(resp.headers))
        return {
            'ok': True,
            'status': resp.status_code,
            'location': loc,
            'final_url': loc or resp.url,
            'body': resp.text or '',
            'set_cookies': set_cookies,
            'headers_str': hdrs_dump
        }
    else:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(url, data=data_bytes, headers=headers, method=method)
        class NoRedirect(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, req, fp, code, msg, hdrs, newurl):
                return None
        opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx), NoRedirect())
        def _decode_bytes(raw_b):
            if not raw_b:
                return ''
            if raw_b[:2] == b'\x1f\x8b':
                try:
                    raw_b = gzip.decompress(raw_b)
                except Exception:
                    pass
            return raw_b.decode('utf-8', errors='replace')
        def _extract_hdr_info(hdr_obj):
            if not hdr_obj:
                return '', [], ''
            loc_v = ''
            for hk in ('Location', 'X-Redirect-Url', 'Content-Location', 'location', 'x-redirect-url', 'content-location'):
                if hdr_obj.get(hk):
                    loc_v = hdr_obj.get(hk)
                    break
            sc_list = []
            try:
                for sc_raw in (hdr_obj.get_all('Set-Cookie') or []):
                    first_p = sc_raw.split(';', 1)[0].strip()
                    if '=' in first_p:
                        sc_list.append(first_p)
            except Exception:
                pass
            return loc_v, sc_list, str(hdr_obj)

        try:
            with opener.open(req, timeout=20) as r:
                loc, sc_list, h_str = _extract_hdr_info(r.headers)
                if loc and not loc.startswith('http'):
                    parsed_u = urllib.parse.urlparse(url)
                    loc = f"{parsed_u.scheme}://{parsed_u.netloc}{loc if loc.startswith('/') else '/' + loc}"
                body_txt = _decode_bytes(r.read())
                return {'ok': True, 'status': r.status, 'location': loc, 'final_url': loc or url, 'body': body_txt, 'set_cookies': sc_list, 'headers_str': h_str}
        except urllib.error.HTTPError as e:
            loc, sc_list, h_str = _extract_hdr_info(e.headers)
            if loc and not loc.startswith('http'):
                parsed_u = urllib.parse.urlparse(url)
                loc = f"{parsed_u.scheme}://{parsed_u.netloc}{loc if loc.startswith('/') else '/' + loc}"
            body_txt = _decode_bytes(e.read() if e.fp else b'')
            return {'ok': True, 'status': e.code, 'location': loc, 'final_url': loc or url, 'body': body_txt, 'set_cookies': sc_list, 'headers_str': h_str}


def execute_apps_script_request(url, method="GET", payload_str=None, timeout=20):
    '''
    Calls a Google Apps Script Web App URL (doGet or doPost) and follows the 302 redirect
    to script.googleusercontent.com via GET to retrieve the JSON response reliably.
    '''
    if HAS_REQUESTS:
        if method.upper() == "POST":
            r1 = requests.post(
                url,
                data=(payload_str or "").encode("utf-8"),
                headers={"Content-Type": "text/plain;charset=utf-8", "User-Agent": "Mozilla/5.0"},
                allow_redirects=False,
                verify=False,
                timeout=timeout
            )
            if r1.status_code in (301, 302, 303, 307, 308):
                loc = r1.headers.get("Location") or r1.headers.get("location") or ""
                if loc:
                    r2 = requests.get(
                        loc,
                        headers={"User-Agent": "Mozilla/5.0"},
                        allow_redirects=True,
                        verify=False,
                        timeout=timeout
                    )
                    return {"ok": r2.status_code == 200, "status": r2.status_code, "body": r2.text or ""}
            return {"ok": r1.status_code == 200, "status": r1.status_code, "body": r1.text or ""}
        else:
            r = requests.get(
                url,
                headers={"User-Agent": "Mozilla/5.0"},
                allow_redirects=True,
                verify=False,
                timeout=timeout
            )
            return {"ok": r.status_code == 200, "status": r.status_code, "body": r.text or ""}
    else:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        res1 = execute_direct_http(
            url,
            method.upper(),
            {"Content-Type": "text/plain;charset=utf-8", "User-Agent": "Mozilla/5.0"},
            payload_str or ""
        )
        st = int(res1.get("status", 0))
        loc = res1.get("location") or ""
        if st in (301, 302, 303, 307, 308) and loc:
            req2 = urllib.request.Request(loc, headers={"User-Agent": "Mozilla/5.0"}, method="GET")
            opener2 = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
            with opener2.open(req2, timeout=timeout) as r2:
                body2 = r2.read().decode("utf-8", errors="replace")
                return {"ok": True, "status": r2.status, "body": body2}
        return {"ok": 200 <= st < 400, "status": st, "body": res1.get("body", "")}


def parse_apps_script_json_or_diagnose(body_str):
    """
    Safely parses JSON from Google Apps Script. If Google returned an HTML page
    (which previously caused 'Expecting value: line 1 column 1 (char 0)'),
    inspects the HTML and returns a clear human-readable diagnosis!
    """
    txt = (body_str or "").strip()
    if not txt:
        return None, (
            "The Script URL returned an empty response.\n\n"
            "Fix in script.new:\n"
            "1. Press Ctrl+S to save your code.\n"
            "2. Click '▷ Run' at the top once and click 'Allow' to authorize.\n"
            "3. Click Deploy -> New deployment -> Web app -> Who has access: Anyone -> Deploy."
        )
    if txt.startswith("{") or txt.startswith("["):
        try:
            return json.loads(txt), None
        except Exception as e:
            return None, f"Invalid JSON response: {e}"

    # Extract readable text from Google's HTML error/login page
    low = txt.lower()
    clean_html = re.sub(r"<style[\s\S]*?</style>", " ", txt, flags=re.I)
    clean_html = re.sub(r"<script[\s\S]*?</script>", " ", clean_html, flags=re.I)
    clean_html = re.sub(r"<[^>]+>", " ", clean_html)
    clean_html = html.unescape(re.sub(r"\s+", " ", clean_html)).strip()

    if "accounts.google.com" in low or "sign in" in low or "servicelogin" in low:
        return None, (
            "Google blocked the connection because 'Who has access' is set to 'Only myself'!\n\n"
            "How to fix in 10 seconds:\n"
            "1. In script.new, click 'Deploy' -> 'Manage deployments'.\n"
            "2. Click the Pencil (Edit) icon.\n"
            "3. Change 'Who has access' to 'Anyone' (NOT 'Only myself').\n"
            "4. Click 'Deploy' and copy the Web App URL."
        )
    if "script function not found: doget" in low or "function not found" in low:
        return None, (
            "Your script at script.new was deployed BEFORE saving the code (so doGet was not found)!\n\n"
            "How to fix in 10 seconds:\n"
            "1. In script.new, press Ctrl+S (Save) so the orange dot disappears.\n"
            "2. Click 'Deploy' -> 'Manage deployments' -> Pencil (Edit) icon.\n"
            "3. Under 'Version', select 'New version' and click 'Deploy'."
        )
    if "authorization" in low or "permission" in low or "access denied" in low:
        return None, (
            "Your script at script.new needs 1-time Google Sheet authorization!\n\n"
            "How to fix in 10 seconds:\n"
            "1. At the top of script.new, select 'authorizeSheetAccess' and click '▷ Run'.\n"
            "2. Click 'Review permissions' -> Choose your Google account -> 'Advanced' -> 'Go to Untitled project (unsafe)' -> 'Allow'."
        )
    return None, f"Google Script returned HTML instead of JSON:\n{clean_html[:280]}"


def extract_published_url_from_response(target_url, title, res):
    parsed_t = urllib.parse.urlparse(target_url)
    origin = f"{parsed_t.scheme}://{parsed_t.netloc}"
    loc = (res.get('location') or '').strip()
    fin = (res.get('final_url') or '').strip()
    body = res.get('body') or ''
    status_code = int(res.get('status') or 0)

    def _is_valid_published_url(u_cand):
        if not u_cand or u_cand.rstrip('/') == target_url.rstrip('/'):
            return False
        if re.search(r'(?:/thread/create|/thread/new|/post/create|/post-thread|/threads/add|/topic/add|/draft|newthread\.php|posting\.php|/movie/new|/novel/upload|/login|/signin|/register|/signup|/error|/oops)(?:[/?#]|$)', u_cand, re.I):
            return False
        return True

    # 1. Check HTTP 301/302/303/307/308 Location / X-Redirect-Url header first
    for cand in (loc, fin):
        if _is_valid_published_url(cand):
            return cand

    # 1B. Check Set-Cookie / HTML for ProBoards `thread_new` or `thread_id` (matching generate_v10_5.py line 987!)
    hdrs_s = res.get('headers_str') or ''
    tm_pb = (
        re.search(r'thread_new(?:=|%2C|,)+(\d+)', hdrs_s, re.I) or
        re.search(r'thread_new(?:=|%2C|,)+(\d+)', body, re.I) or
        (re.search(r'[\'" ]thread_id[\'" ]\s*[:=,]\s*(\d+)', body, re.I) if is_proboards_url_or_curl(target_url) else None)
    )
    if tm_pb:
        return f"{origin}/thread/{tm_pb.group(1)}" 

    # 2. Platform-specific JSON / HTML responses (Pixiv, TMDB, Flarum, Discourse, etc.)
    if 'pixiv.net' in target_url.lower():
        m_pix = re.search(r'["\'](?:id|novel_id)["\']\s*:\s*["\']?(\d{5,})["\']?', body)
        if m_pix:
            return f"https://www.pixiv.net/novel/show.php?id={m_pix.group(1)}"

    # 3. Deep recursive JSON URL / ID extraction for AJAX / REST / GraphQL forums
    if body.strip().startswith('{') or body.strip().startswith('['):
        try:
            j = json.loads(body)
            url_key_priority = (
                'redirect', 'redirect_url', 'redirectUrl', 'url', 'permalink', 'link',
                'location', 'post_url', 'postUrl', 'thread_url', 'threadUrl',
                'topic_url', 'topicUrl', 'discussion_url', 'canonical_url',
                'html_url', 'web_url', 'short_url', 'uri', 'path'
            )
            found_urls = []
            found_ids = {}

            def _scan_json(obj):
                if isinstance(obj, dict):
                    for k in url_key_priority:
                        v = obj.get(k)
                        if isinstance(v, str) and (v.startswith('http') or v.startswith('/')):
                            full_u = v if v.startswith('http') else origin + v
                            if _is_valid_published_url(full_u):
                                found_urls.append(full_u)
                    for k, v in obj.items():
                        if k in ('id', 'thread_id', 'topic_id', 'discussion_id', 'post_id', 'movie_id', 'slug') and isinstance(v, (str, int)):
                            found_ids[k] = str(v)
                        elif isinstance(v, (dict, list)):
                            _scan_json(v)
                elif isinstance(obj, list):
                    for it in obj:
                        _scan_json(it)

            _scan_json(j)
            if found_urls:
                return found_urls[0]

            # Construct URL if JSON returned an ID/slug on known platforms
            slug = found_ids.get('slug') or re.sub(r'[^a-z0-9]+', '-', (title or '').lower()).strip('-')[:60]
            if 'hoyolab.com' in target_url.lower() and (found_ids.get('post_id') or found_ids.get('id')):
                pid = found_ids.get('post_id') or found_ids.get('id')
                return f"https://www.hoyolab.com/article/{pid}"
            if 'themoviedb.org' in target_url.lower() and (found_ids.get('id') or found_ids.get('movie_id')):
                mid = found_ids.get('id') or found_ids.get('movie_id')
                return f"https://www.themoviedb.org/movie/{mid}-{slug}"
            if '/api/discussions' in target_url.lower() and found_ids.get('id'):
                return f"{origin}/d/{found_ids['id']}-{slug}"
            if found_ids.get('topic_id') and '/posts' in target_url.lower():
                return f"{origin}/t/{slug}/{found_ids['topic_id']}"
            if is_proboards_url_or_curl(target_url) and (found_ids.get('thread_id') or found_ids.get('id')):
                tid = found_ids.get('thread_id') or found_ids.get('id')
                return f"{origin}/thread/{tid}/{slug}"
        except Exception:
            pass

    # 4. Check HTML Canonical / OpenGraph / Meta-Refresh / JS Redirects
    for meta_pat in (
        r'<link\s+[^>]*rel=["\']canonical["\'][^>]*href=["\']([^"\']+)["\']',
        r'<link\s+[^>]*href=["\']([^"\']+)["\'][^>]*rel=["\']canonical["\']',
        r'<meta\s+[^>]*property=["\']og:url["\'][^>]*content=["\']([^"\']+)["\']',
        r'<meta\s+[^>]*http-equiv=["\']refresh["\'][^>]*content=["\'][^"\']*url=([^"\'\s;>]+)',
        r'(?:window\.)?location(?:\.href|\.replace)?\s*(?:=|\()\s*["\']((?:https?://|/)[^"\']+)["\']',
    ):
        m_meta = re.search(meta_pat, body, re.I)
        if m_meta:
            u = html.unescape(m_meta.group(1).strip())
            if not u.startswith('http'):
                u = origin + (u if u.startswith('/') else '/' + u)
            if _is_valid_published_url(u):
                return u

    # 5. Scan HTML anchor/thread links
    for pat in (
        r'href=["\']((?:https?://[^"\']+)?/(?:thread|threads|topic|topics|discussion|discussions|post|posts|movie|novel|entry|p|d|t|study-guides|articles|questions)/[^"\'\s<>]+)["\']',
        r'(https?://[^\s"\'<>]+/(?:thread|threads|topic|topics|discussion|discussions|post|posts|movie|novel|study-guides|articles)/[^\s"\'<>]+)'
    ):
        for m in re.finditer(pat, body, re.I):
            u = html.unescape(m.group(1).strip())
            if not u.startswith('http'):
                u = origin + (u if u.startswith('/') else '/' + u)
            if _is_valid_published_url(u) and not re.search(r'/(?:create|new|login|register|edit|search|category|board)(?:[/?#]|$)', u, re.I):
                return u

    return ''


# ==============================================================================
# PURE-PYTHON CHROME DEVTOOLS PROTOCOL (CDP) WEBSOCKET SNIFFER & COOKIE PERSISTER
# - Automatically captures POST cURLs when you do a test post in Chrome (no F12 needed!)
# - Promotes Session Cookies (like ProBoards login) to 30-day persistent cookies AND
#   saves/restores them on disk so closing Chrome NEVER logs you out!
# ==============================================================================
class _SimpleCDPWebSocket:
    """Minimal RFC 6455 WebSocket client using only Python stdlib `socket`."""
    def __init__(self, ws_url, timeout=3.0):
        import base64
        import struct
        self.struct = struct
        parsed = urllib.parse.urlparse(ws_url)
        host = parsed.hostname or "127.0.0.1"
        port = parsed.port or 9222
        path = parsed.path or "/"
        if parsed.query:
            path += "?" + parsed.query

        self.sock = socket.create_connection((host, port), timeout=timeout)
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        handshake = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            f"Upgrade: websocket\r\n"
            f"Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            f"Sec-WebSocket-Version: 13\r\n\r\n"
        )
        self.sock.sendall(handshake.encode("ascii"))
        resp = b""
        while b"\r\n\r\n" not in resp:
            chunk = self.sock.recv(1024)
            if not chunk:
                raise ConnectionError("WebSocket handshake closed")
            resp += chunk
        self.sock.settimeout(0.35)

    def send_json(self, obj):
        payload = json.dumps(obj).encode("utf-8")
        mask_key = os.urandom(4)
        header = bytearray([0x81])  # FIN + Text frame
        length = len(payload)
        if length < 126:
            header.append(0x80 | length)
        elif length < 65536:
            header.append(0x80 | 126)
            header.extend(self.struct.pack("!H", length))
        else:
            header.append(0x80 | 127)
            header.extend(self.struct.pack("!Q", length))
        header.extend(mask_key)
        masked = bytearray(length)
        for i in range(length):
            masked[i] = payload[i] ^ mask_key[i & 3]
        self.sock.sendall(header + masked)

    def _recv_exact(self, n):
        buf = b""
        while len(buf) < n:
            chunk = self.sock.recv(n - len(buf))
            if not chunk:
                raise ConnectionError("WebSocket EOF")
            buf += chunk
        return buf

    def recv_json(self):
        try:
            b1, b2 = self._recv_exact(2)
        except socket.timeout:
            return None
        opcode = b1 & 0x0F
        masked = b2 & 0x80
        length = b2 & 0x7F
        if length == 126:
            length = self.struct.unpack("!H", self._recv_exact(2))[0]
        elif length == 127:
            length = self.struct.unpack("!Q", self._recv_exact(8))[0]
        mask_key = self._recv_exact(4) if masked else b""
        data = self._recv_exact(length) if length > 0 else b""
        if masked:
            data = bytes(data[i] ^ mask_key[i & 3] for i in range(length))
        if opcode == 0x8:  # Close
            raise ConnectionError("WebSocket closed")
        if opcode == 0x9:  # Ping -> send Pong
            return None
        if opcode == 0x1:
            return json.loads(data.decode("utf-8", errors="replace"))
        return None

    def close(self):
        try:
            self.sock.close()
        except Exception:
            pass


class ForumStudioDesktopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FS | Forum Studio Pro v13.6 — Hatena/Diigo/HoYoLAB Bookmark Fix + 2s Delay + Sheet Master")
        self.root.geometry("1440x890")
        self.root.configure(bg="#070b14")

        # Try setting window icon if zoom_icon.ico exists in folder
        ico_path = os.path.join(get_script_dir(), "zoom_icon.ico")
        if os.path.exists(ico_path):
            try:
                self.root.iconbitmap(ico_path)
            except Exception:
                pass

        # State variables
        self.sites = {}          # name -> curl string
        self.site_urls = {}      # name -> target composer url
        self.bookmarks = []      # list of {"site": ..., "type": "SINGLE"/"BULK", "curl": ..., "active": True}
        self.deleted_sites = []
        self.deleted_bookmarks = []
        self.active_site = ""
        self.rows = []
        self.history = []
        self.next_open_col = ""
        self.latest_active_col = ""
        self._main_sheet_synced_once = False
        self.detected_kw_count = 20

        self.content_sheet_url = DEFAULT_CONTENT_SHEET_URL
        self.main_sheet_url = DEFAULT_MAIN_SHEET_URL
        self.apps_script_url = DEFAULT_APPS_SCRIPT_URL
        self.ai_key = ""
        self.post_delay = 25
        self.bm_delay = 2
        self.sheet_col = ""
        self.chrome_profile = "Profile 1"
        self.engine_mode = "curl"

        self.is_posting = False
        self.is_bookmarking = False
        self.is_paused = False
        self.stop_requested = False
        self.listening_clipboard = False
        self.last_clip_text = ""
        self.ctx_row_id = None

        # 1. Load user's saved vaults & sheet links from disk (starts blank only if no config saved yet)
        self.load_all_vaults_from_disk()

        # 2. Build the Native Tkinter UI
        self.build_ui()

        # 3. Populate UI from loaded state
        self.populate_ui_from_state()

        # 4. Sync sheets in background if user has already saved a sheet link
        if self.content_sheet_url.strip() or self.main_sheet_url.strip():
            self.root.after(400, lambda: self.sync_all_sheets(show_popup=False))

        # 5. Always save all vaults & sheet links on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)

    # --------------------------------------------------------------------------
    # DISK VAULT LOADER & SAVER (Persists user's Sheet links, Script URL, Sites & Bookmarks across restarts!)
    # --------------------------------------------------------------------------
    def load_all_vaults_from_disk(self):
        cfg = read_json_with_backup("zoom_sender_config.json", {})
        if isinstance(cfg, dict):
            self.deleted_sites = list(cfg.get("deleted_sites", []))
            self.deleted_bookmarks = list(cfg.get("deleted_bookmarks", []))
            self.active_site = str(cfg.get("active_site", ""))
            self.content_sheet_url = str(cfg.get("content_sheet_url", DEFAULT_CONTENT_SHEET_URL))
            self.main_sheet_url = str(cfg.get("main_sheet_url", DEFAULT_MAIN_SHEET_URL))
            self.apps_script_url = str(cfg.get("apps_script_url", DEFAULT_APPS_SCRIPT_URL)).strip()
            self.ai_key = str(cfg.get("ai_key", ""))
            try:
                self.post_delay = int(cfg.get("post_delay", 25))
            except Exception:
                self.post_delay = 25
            try:
                self.bm_delay = max(0, int(cfg.get("bm_delay", 2)))
            except Exception:
                self.bm_delay = 2
            self.sheet_col = str(cfg.get("sheet_col", ""))
            self.chrome_profile = str(cfg.get("chrome_profile", "Profile 1"))
            self.engine_mode = str(cfg.get("engine_mode", "curl"))
            if isinstance(cfg.get("site_urls"), dict):
                self.site_urls = dict(cfg["site_urls"])
            if isinstance(cfg.get("rows"), list) and cfg["rows"]:
                self.rows = [
                    {
                        "id": idx_r,
                        "status": "Pending",
                        "title": str(r.get("title", "")).strip(),
                        "body": normalize_content_links_to_html(str(r.get("body", "")).strip()),
                        "bookmark_status": "-",
                        "url": "-",
                        "selected": True
                    }
                    for idx_r, r in enumerate(cfg["rows"], start=1)
                    if str(r.get("title", "")).strip() or str(r.get("body", "")).strip()
                ]

        raw_sites = read_json_with_backup("sites_curl_vault.json", {})
        if isinstance(raw_sites, dict):
            s_dict = raw_sites.get("sites", raw_sites) if ("sites" in raw_sites and isinstance(raw_sites["sites"], dict)) else raw_sites
            for k, v in s_dict.items():
                if k in self.deleted_sites:
                    continue
                if isinstance(v, str) and "curl" in v.lower():
                    self.sites[k] = v
                elif isinstance(v, dict) and "curl" in str(v.get("curl", "")).lower():
                    self.sites[k] = str(v["curl"])
                    if v.get("target_url"):
                        self.site_urls[k] = str(v["target_url"])

        for s_k, s_c in list(self.sites.items()):
            if "themoviedb.org/movie/content_check" in s_c or "themoviedb.org/movie/duplicate_check" in s_c:
                s_c = s_c.replace("themoviedb.org/movie/content_check", "themoviedb.org/movie/new").replace(
                    "themoviedb.org/movie/duplicate_check", "themoviedb.org/movie/new"
                )
                self.sites[s_k] = s_c
            c_dom = extract_domain_from_curl(s_c)
            if is_ad_or_tracker_url(c_dom) or is_ad_or_tracker_url(s_c):
                self.sites.pop(s_k, None)
                self.site_urls.pop(s_k, None)
                continue
            if c_dom and c_dom != "Unknown Site":
                sk_low = s_k.lower()
                if "flying" in sk_low and "themoviedb" in c_dom:
                    if "Themoviedb" not in self.sites:
                        self.sites["Themoviedb"] = s_c
                    del self.sites[s_k]

        if self.active_site not in self.sites:
            self.active_site = sorted(self.sites.keys())[0] if self.sites else ""

        raw_bms = read_json_with_backup("bookmarks_curl_vault.json", [])
        del_bm_low = {x.lower() for x in self.deleted_bookmarks}
        loaded_bms = []
        seen_bm = set()

        if isinstance(raw_bms, dict):
            b_inner = raw_bms.get("bookmarks", raw_bms)
            if isinstance(b_inner, list):
                raw_bms = b_inner
            elif isinstance(b_inner, dict):
                for k, v in b_inner.items():
                    if k.lower() in del_bm_low or k.lower() in seen_bm:
                        continue
                    if isinstance(v, str) and "curl" in v.lower():
                        loaded_bms.append({"site": k, "type": "SINGLE", "curl": v, "active": True})
                        seen_bm.add(k.lower())
                    elif isinstance(v, dict):
                        c_str = str(v.get("curl", ""))
                        if "curl" in c_str.lower():
                            loaded_bms.append({
                                "site": v.get("site") or k,
                                "type": str(v.get("type", "SINGLE")).upper(),
                                "curl": c_str,
                                "active": bool(v.get("active", True))
                            })
                            seen_bm.add(k.lower())

        if isinstance(raw_bms, list):
            for item in raw_bms:
                if isinstance(item, dict):
                    s_name = str(item.get("site") or item.get("name") or "").strip()
                    c_str = str(item.get("curl") or "").strip()
                    if s_name and s_name.lower() not in del_bm_low and s_name.lower() not in seen_bm and "curl" in c_str.lower():
                        loaded_bms.append({
                            "site": s_name,
                            "type": str(item.get("type", "SINGLE")).upper(),
                            "curl": c_str,
                            "active": bool(item.get("active", True))
                        })
                        seen_bm.add(s_name.lower())

        self.bookmarks = loaded_bms
        raw_hist = read_json_with_backup("post_history_vault.json", [])
        if isinstance(raw_hist, list):
            self.history = raw_hist

    def save_all_vaults_to_disk(self, sync_from_ui=True):
        if sync_from_ui and hasattr(self, "entry_content_sheet"):
            try:
                self.content_sheet_url = self.entry_content_sheet.get().strip()
                self.main_sheet_url = self.entry_main_sheet.get().strip()
                self.apps_script_url = self.entry_apps_script.get().strip()
                self.ai_key = self.entry_ai_key.get().strip()
                try:
                    self.post_delay = int(self.spin_delay.get())
                except Exception:
                    pass
                try:
                    self.bm_delay = max(0, int(self.spin_bm_delay.get()))
                except Exception:
                    pass
                self.sheet_col = self.entry_sheet_col.get().strip()
                self.chrome_profile = self.combo_profile.get().strip()
                self.engine_mode = self.var_engine_mode.get()

                # Never overwrite existing self.sites or self.bookmarks entries from half-edited text boxes here!
                # Sites and Bookmark Accounts are updated only on explicit Save / Capture so pasting a 2nd account's
                # cURL (e.g. hatena2 while hatena1 is selected) NEVER overwrites hatena1!
                pass
            except Exception:
                pass

        write_json_with_backup("sites_curl_vault.json", self.sites)
        write_json_with_backup("bookmarks_curl_vault.json", self.bookmarks)
        write_json_with_backup("post_history_vault.json", self.history)

        cfg_payload = {
            "last_saved_ts": time.time(),
            "active_site": self.active_site,
            "site_urls": self.site_urls,
            "deleted_sites": self.deleted_sites,
            "deleted_bookmarks": self.deleted_bookmarks,
            "content_sheet_url": self.content_sheet_url,
            "main_sheet_url": self.main_sheet_url,
            "apps_script_url": self.apps_script_url,
            "ai_key": self.ai_key,
            "post_delay": self.post_delay,
            "bm_delay": self.bm_delay,
            "sheet_col": self.sheet_col,
            "chrome_profile": self.chrome_profile,
            "engine_mode": self.engine_mode,
            "save_ankit_sheet": bool(self.var_save_sheet.get()) if hasattr(self, "var_save_sheet") else True,
            "auto_bookmark": bool(self.var_auto_bm.get()) if hasattr(self, "var_auto_bm") else True,
            "hatena_shield": bool(self.var_hatena_shield.get()) if hasattr(self, "var_hatena_shield") else True,
            "auto_send_zoom": bool(self.var_auto_zoom.get()) if hasattr(self, "var_auto_zoom") else True,
            "bg_browser": bool(self.var_bg_browser.get()) if hasattr(self, "var_bg_browser") else False,
            "rows": self.rows
        }
        write_json_with_backup("zoom_sender_config.json", cfg_payload)

    def on_window_close(self):
        self.stop_requested = True
        self._cdp_session_id = getattr(self, "_cdp_session_id", 0) + 1
        self.save_all_vaults_to_disk(sync_from_ui=True)
        self.root.destroy()

    # --------------------------------------------------------------------------
    # UI CONSTRUCTION (EXACT MATCH TO image_0d8718.png)
    # --------------------------------------------------------------------------
    def build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Treeview", background="#060a12", foreground="#e2e8f0", fieldbackground="#060a12", rowheight=25, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background="#0f172a", foreground="#94a3b8", font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#1e3a8a")], foreground=[("selected", "#ffffff")])

        # --- HEADER ROW 1 ---
        hdr1 = tk.Frame(self.root, bg="#090e1a", padx=10, pady=5)
        hdr1.pack(fill="x")

        tk.Label(hdr1, text=" FS ", bg="#4f46e5", fg="white", font=("Segoe UI", 11, "bold"), padx=4, pady=2).pack(side="left", padx=(0, 8))
        tk.Label(hdr1, text="Forum Studio Pro", bg="#090e1a", fg="white", font=("Segoe UI", 13, "bold")).pack(side="left", padx=(0, 8))
        tk.Label(hdr1, text=" v13.6 Bookmarking Fixed ", bg="#1e3a8a", fg="#93c5fd", font=("Segoe UI", 8, "bold"), padx=6, pady=2).pack(side="left", padx=4)
        self.lbl_batch_badge = tk.Label(hdr1, text=" 📊 Active Batch: 0 Keywords (Enter Sheet Links & Click Sync) ", bg="#064e3b", fg="#6ee7b7", font=("Segoe UI", 8, "bold"), padx=8, pady=2)
        self.lbl_batch_badge.pack(side="left", padx=6)

        lan_ip = get_lan_ip()
        self.lbl_top_mobile = tk.Label(hdr1, text=f" 📱 Mobile: http://{lan_ip}:5050 ", bg="#064e3b", fg="#34d399", font=("Segoe UI", 8, "bold"), padx=8, pady=4, cursor="hand2")
        self.lbl_top_mobile.pack(side="right", padx=4)
        self.lbl_top_mobile.bind("<Button-1>", lambda e: self.switch_tab(5))

        tk.Button(hdr1, text="📋 Copy Sheet Script (Code.gs)", bg="#1e293b", fg="#fbbf24", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=3, command=self.copy_sheet_script).pack(side="right", padx=4)
        tk.Button(hdr1, text="🔄 Sync Sheets", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=3, command=lambda: self.sync_all_sheets(show_popup=True)).pack(side="right", padx=4)

        # --- HEADER ROW 2 (Main Sheet Link | Content Sheet Link | Apps Script URL) ---
        hdr2 = tk.Frame(self.root, bg="#0b1120", padx=10, pady=4)
        hdr2.pack(fill="x")

        tk.Label(hdr2, text="📊 Main Sheet Link:", bg="#0b1120", fg="#fbbf24", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_main_sheet = tk.Entry(hdr2, bg="#050811", fg="#e2e8f0", insertbackground="white", relief="solid", bd=1, width=28, font=("Consolas", 9))
        self.entry_main_sheet.pack(side="left", padx=(4, 10), fill="x", expand=True)
        self.entry_main_sheet.bind("<KeyRelease>", lambda e: self.save_all_vaults_to_disk(sync_from_ui=True))

        tk.Label(hdr2, text="📥 Content Sheet Link:", bg="#0b1120", fg="#34d399", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_content_sheet = tk.Entry(hdr2, bg="#050811", fg="#e2e8f0", insertbackground="white", relief="solid", bd=1, width=28, font=("Consolas", 9))
        self.entry_content_sheet.pack(side="left", padx=(4, 10), fill="x", expand=True)
        self.entry_content_sheet.bind("<KeyRelease>", lambda e: self.save_all_vaults_to_disk(sync_from_ui=True))

        tk.Label(hdr2, text="🔗 Apps Script URL:", bg="#0b1120", fg="#94a3b8", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_apps_script = tk.Entry(hdr2, bg="#050811", fg="#e2e8f0", insertbackground="white", relief="solid", bd=1, width=28, font=("Consolas", 9))
        self.entry_apps_script.pack(side="left", padx=(4, 10), fill="x", expand=True)
        self.entry_apps_script.bind("<KeyRelease>", lambda e: self.save_all_vaults_to_disk(sync_from_ui=True))

        self.entry_ai_key = tk.Entry(hdr2)  # Hidden internal holder for backwards config compatibility

        tk.Button(hdr2, text="🔌 Test Sheet", bg="#059669", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=3, command=self.test_sheet_connection).pack(side="right", padx=(4, 0))
        tk.Button(hdr2, text="💾 Save & Load Sheets", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=3, command=lambda: self.sync_all_sheets(show_popup=True)).pack(side="right")

        # --- 4 CUSTOM TAB BUTTONS BAR ---
        tab_bar = tk.Frame(self.root, bg="#090e1a", padx=10, pady=2)
        tab_bar.pack(fill="x")
        self.tab_btns = {}
        tab_titles = [
            (1, "🚀 Forum Poster & Auto-Bookmark Queue"),
            (2, "🔖 Bookmarking Engine & Hatena Shield"),
            (3, "💬 Zoom Live Auto-Sender"),
            (4, "📜 Post History & Saved Links"),
            (5, "📱 Mobile Remote Control (Wi-Fi & 4G/5G)")
        ]
        for idx, text_lbl in tab_titles:
            b = tk.Button(
                tab_bar, text=text_lbl, bg="#2563eb" if idx == 1 else "#111827",
                fg="white" if idx == 1 else "#94a3b8", font=("Segoe UI", 9, "bold"),
                relief="flat", padx=12, pady=5, command=lambda i=idx: self.switch_tab(i)
            )
            b.pack(side="left", padx=(0, 4))
            self.tab_btns[idx] = b

        # Container for Tab Frames
        self.tab_container = tk.Frame(self.root, bg="#070b14")
        self.tab_container.pack(fill="both", expand=True, padx=8, pady=4)

        self.pane1 = tk.Frame(self.tab_container, bg="#070b14")
        self.pane2 = tk.Frame(self.tab_container, bg="#070b14")
        self.pane3 = tk.Frame(self.tab_container, bg="#070b14")
        self.pane4 = tk.Frame(self.tab_container, bg="#070b14")
        self.pane5 = tk.Frame(self.tab_container, bg="#070b14")

        self.build_tab1()
        self.build_tab2()
        self.build_tab3()
        self.build_tab_history()
        self.build_tab4()
        self.switch_tab(1)

    def switch_tab(self, tab_num):
        for i, p in enumerate((self.pane1, self.pane2, self.pane3, self.pane4, self.pane5), start=1):
            if i == tab_num:
                p.pack(fill="both", expand=True)
                self.tab_btns[i].config(bg="#2563eb", fg="white")
            else:
                p.pack_forget()
                self.tab_btns[i].config(bg="#111827", fg="#94a3b8")

    # --------------------------------------------------------------------------
    # TAB 1: FORUM POSTER & AUTO-BOOKMARK QUEUE (EXACT image_0d8718.png LAYOUT)
    # --------------------------------------------------------------------------
    def build_tab1(self):
        # Command Box
        cmd_box = tk.Frame(self.pane1, bg="#0d1424", highlightbackground="#1e293b", highlightthickness=1, padx=8, pady=6)
        cmd_box.pack(fill="x", pady=(0, 6))

        r1 = tk.Frame(cmd_box, bg="#0d1424")
        r1.pack(fill="x", pady=(0, 5))
        self.btn_start = tk.Button(r1, text="🚀 START POSTING (+ AUTO-BM)", bg="#10b981", fg="#042f2e", font=("Segoe UI", 9, "bold"), relief="flat", padx=14, pady=5, command=self.toggle_posting_automation)
        self.btn_start.pack(side="left", padx=(0, 5))
        self.btn_pause = tk.Button(r1, text="⏸ PAUSE", bg="#d97706", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, command=self.toggle_pause)
        self.btn_pause.pack(side="left", padx=4)
        tk.Button(r1, text="🔄 RETRY FAILED", bg="#d97706", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, command=self.retry_failed_or_selected).pack(side="left", padx=4)
        tk.Button(r1, text="🧪 TEST 1 POST", bg="#0284c7", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, command=self.test_single_post).pack(side="left", padx=4)
        tk.Button(r1, text="🔖 BOOKMARK QUEUE", bg="#4f46e5", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, command=self.start_queue_bookmarking).pack(side="left", padx=4)
        tk.Button(r1, text="🛑 STOP", bg="#e11d48", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=14, pady=5, command=self.stop_all).pack(side="left", padx=4)
        tk.Button(r1, text="📋 Copy All Posted Links", bg="#1e293b", fg="#34d399", font=("Segoe UI", 8, "bold"), relief="flat", padx=12, pady=5, command=self.copy_all_posted_links).pack(side="right")

        r2 = tk.Frame(cmd_box, bg="#080c17", highlightbackground="#172033", highlightthickness=1, padx=8, pady=3)
        r2.pack(fill="x")
        tk.Label(r2, text="⏱ Delay (s):", bg="#080c17", fg="#cbd5e1", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.spin_delay = tk.Spinbox(r2, from_=1, to=300, width=5, bg="#0f172a", fg="white", font=("Segoe UI", 9, "bold"), relief="solid", bd=1, command=lambda: self.save_all_vaults_to_disk(True))
        self.spin_delay.pack(side="left", padx=(4, 14))
        self.spin_delay.bind("<KeyRelease>", lambda e: self.save_all_vaults_to_disk(True))

        tk.Label(r2, text="📊 Sheet Col (Blank=Auto):", bg="#080c17", fg="#cbd5e1", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_sheet_col = tk.Entry(r2, width=7, bg="#0f172a", fg="#38bdf8", insertbackground="white", font=("Segoe UI", 9, "bold"), justify="center", relief="solid", bd=1)
        self.entry_sheet_col.pack(side="left", padx=(4, 14))
        self.entry_sheet_col.bind("<KeyRelease>", lambda e: self.save_all_vaults_to_disk(True))

        self.var_save_sheet = tk.BooleanVar(value=True)
        self.var_auto_bm = tk.BooleanVar(value=True)
        self.var_hatena_shield = tk.BooleanVar(value=True)
        self.var_auto_zoom = tk.BooleanVar(value=True)
        self.var_bg_browser = tk.BooleanVar(value=False)

        for var, txt, col in (
            (self.var_save_sheet, "📄 Save Links in Main Sheet", "#fbbf24"),
            (self.var_auto_bm, "🔖 Auto-Bookmark", "#34d399"),
            (self.var_hatena_shield, "🛡 Hatena Shield", "#fb7185"),
            (self.var_auto_zoom, "💬 Auto-Send Zoom", "#38bdf8"),
            
        ):
            cb = tk.Checkbutton(
                r2, text=txt, variable=var, bg="#080c17", fg=col,
                selectcolor="#0f172a", activebackground="#080c17", activeforeground=col,
                font=("Segoe UI", 8, "bold"), command=lambda: self.save_all_vaults_to_disk(True)
            )
            cb.pack(side="left", padx=6)

        self.lbl_live_timer = tk.Label(r2, text="⏳ Timer: Ready", bg="#064e3b", fg="#34d399", font=("Consolas", 9, "bold"), padx=8, pady=1)
        self.lbl_live_timer.pack(side="right", padx=4)

        # Workspace Split (Left Panel 360px | Right Panel Expand)
        work = tk.Frame(self.pane1, bg="#070b14")
        work.pack(fill="both", expand=True)

        left_p = tk.Frame(work, bg="#0d1424", highlightbackground="#1e293b", highlightthickness=1, width=365, padx=10, pady=8)
        left_p.pack(side="left", fill="y", padx=(0, 6))
        left_p.pack_propagate(False)

        right_p = tk.Frame(work, bg="#070b14")
        right_p.pack(side="left", fill="both", expand=True)

        # --- LEFT SECTION 1: Saved Site Vault ---
        s1_top = tk.Frame(left_p, bg="#0d1424")
        s1_top.pack(fill="x")
        tk.Label(s1_top, text="1. Saved Site Vault", bg="#0d1424", fg="white", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.lbl_active_site = tk.Label(s1_top, text="Active Site: None", bg="#0d1424", fg="#10b981", font=("Segoe UI", 8, "bold"))
        self.lbl_active_site.pack(side="right")

        self.lbl_active_domain = tk.Label(left_p, text="Endpoint: -", bg="#0d1424", fg="#94a3b8", font=("Segoe UI", 8))
        self.lbl_active_domain.pack(anchor="w", pady=(1, 3))

        s1_row = tk.Frame(left_p, bg="#0d1424")
        s1_row.pack(fill="x", pady=(0, 8))
        self.combo_sites = ttk.Combobox(s1_row, state="readonly", width=23)
        self.combo_sites.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.combo_sites.bind("<<ComboboxSelected>>", lambda e: self.on_site_selected())
        tk.Button(s1_row, text="⚡ Activate", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, command=self.on_site_selected).pack(side="left", padx=2)
        tk.Button(s1_row, text="🗑", bg="#e11d48", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=7, pady=3, command=self.delete_selected_site).pack(side="left", padx=(2, 0))

        # --- LEFT SECTION 2: Engine Mode & Auto-Capture cURL ---
        tk.Label(left_p, text="2. Auto-Capture cURL from Chrome", bg="#0d1424", fg="white", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(2, 3))
        self.var_engine_mode = tk.StringVar(value="curl")

        c_row = tk.Frame(left_p, bg="#0d1424")
        c_row.pack(fill="x", pady=(0, 5))
        self.combo_profile = ttk.Combobox(c_row, values=["Profile 1", "Profile 2", "Profile 3", "Default"], state="readonly", width=9)
        self.combo_profile.set("Profile 1")
        self.combo_profile.pack(side="left", padx=(0, 4))
        self.combo_profile.bind("<<ComboboxSelected>>", lambda e: self.save_all_vaults_to_disk(True))
        self.btn_open_capture = tk.Button(c_row, text="🎯 Open Chrome & Capture", bg="#d97706", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, pady=4, command=self.open_chrome_and_capture)
        self.btn_open_capture.pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(c_row, text="📗 Show", bg="#10b981", fg="#042f2e", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, pady=4, command=self.show_target_browser).pack(side="left", padx=2)
        tk.Button(c_row, text="🧹", bg="#1e293b", fg="#fda4af", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, pady=4, command=self.clear_posting_curl_box).pack(side="left", padx=(2, 0))

        tu_hdr = tk.Frame(left_p, bg="#0d1424")
        tu_hdr.pack(fill="x", pady=(2, 2))
        tk.Label(tu_hdr, text="TARGET SITE / COMPOSER URL (FOR BROWSER OR CAPTURE)", bg="#0d1424", fg="#10b981", font=("Segoe UI", 7, "bold")).pack(side="left")
        tk.Button(tu_hdr, text="🧹 Clear URL", bg="#1e293b", fg="#fda4af", font=("Segoe UI", 7, "bold"), relief="flat", padx=5, pady=0, command=lambda: self.entry_target_url.delete(0, "end")).pack(side="right", padx=(2, 0))
        tk.Button(tu_hdr, text="📋 Paste URL", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 7, "bold"), relief="flat", padx=5, pady=0, command=self.paste_target_site_url).pack(side="right")

        self.entry_target_url = tk.Entry(left_p, bg="#050811", fg="#34d399", insertbackground="white", font=("Consolas", 9), relief="solid", bd=1)
        self.entry_target_url.pack(fill="x", pady=(0, 6))

        # --- LEFT SECTION 3: Site Name & cURL Configuration ---
        tk.Label(left_p, text="3. Site Name & cURL Configuration", bg="#0d1424", fg="white", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(2, 3))
        sn_row = tk.Frame(left_p, bg="#0d1424")
        sn_row.pack(fill="x", pady=(0, 4))
        self.entry_site_name = tk.Entry(sn_row, bg="#050811", fg="white", insertbackground="white", font=("Segoe UI", 9, "bold"), relief="solid", bd=1)
        self.entry_site_name.pack(side="left", fill="x", expand=True, padx=(0, 4))
        # Site Name saves when you click Save & Activate Site in Vault
        tk.Button(sn_row, text="📋 Paste cURL", bg="#1e293b", fg="#e2e8f0", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, command=self.paste_posting_curl).pack(side="right")

        self.txt_curl = scrolledtext.ScrolledText(left_p, bg="#050811", fg="#38bdf8", insertbackground="white", font=("Consolas", 8), wrap="word", height=11)
        self.txt_curl.pack(fill="both", expand=True, pady=(0, 5))
        self.txt_curl.bind("<KeyRelease>", lambda e: self.on_posting_curl_modified())
        self.txt_curl.bind("<<Paste>>", lambda e: self.root.after(60, self.on_posting_curl_modified))
        self.txt_curl.bind("<FocusOut>", lambda e: self.save_all_vaults_to_disk(True))

        sv_row = tk.Frame(left_p, bg="#0d1424")
        sv_row.pack(fill="x", pady=(0, 4))
        tk.Button(sv_row, text="💾 Save & Activate Site in Vault", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", pady=5, command=lambda: self.save_site_to_vault(True)).pack(side="left", fill="x", expand=True, padx=(0, 4))
        tk.Button(sv_row, text="💾 Reserve Only", bg="#1e293b", fg="#e2e8f0", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=5, command=lambda: self.save_site_to_vault(False)).pack(side="right")

        lan_ip = get_lan_ip()
        tk.Label(left_p, text=f"📱 Mobile Ready: http://{lan_ip}:5050", bg="#0d1424", fg="#34d399", font=("Segoe UI", 8, "bold")).pack(anchor="w")

        # --- RIGHT PANEL: Queue Table + Console ---
        q_card = tk.Frame(right_p, bg="#0d1424", highlightbackground="#1e293b", highlightthickness=1, padx=10, pady=8)
        q_card.pack(fill="both", expand=True, pady=(0, 6))

        q_top = tk.Frame(q_card, bg="#0d1424")
        q_top.pack(fill="x", pady=(0, 4))
        self.lbl_queue_title = tk.Label(q_top, text="Bookmarking Queue (0 / 0 items)", bg="#0d1424", fg="white", font=("Segoe UI", 10, "bold"))
        self.lbl_queue_title.pack(side="left")
        self.lbl_selected_count = tk.Label(q_top, text="  (0/0 selected)", bg="#0d1424", fg="#6ee7b7", font=("Segoe UI", 9, "bold"))
        self.lbl_selected_count.pack(side="left")

        self.lbl_queue_pct = tk.Label(q_top, text="0%", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 10, "bold"))
        self.lbl_queue_pct.pack(side="right", padx=(6, 0))
        tk.Button(q_top, text="🧹 Reset Queue", bg="#1e293b", fg="#fda4af", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=self.reset_queue_statuses).pack(side="right", padx=2)
        tk.Button(q_top, text="🔄 Invert", bg="#1e293b", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=self.invert_selection).pack(side="right", padx=2)
        tk.Button(q_top, text="🔲 None", bg="#1e293b", fg="#fda4af", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=lambda: self.select_all_rows(False)).pack(side="right", padx=2)
        tk.Button(q_top, text="✅ All", bg="#1e293b", fg="#34d399", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=lambda: self.select_all_rows(True)).pack(side="right", padx=2)

        t_wrap = tk.Frame(q_card, bg="#060a12")
        t_wrap.pack(fill="both", expand=True)
        cols = ("check", "id", "status", "title", "bm_status", "url")
        self.tree = ttk.Treeview(t_wrap, columns=cols, show="headings", height=12)
        self.tree.heading("check", text="✓")
        self.tree.heading("id", text="#")
        self.tree.heading("status", text="Post Status")
        self.tree.heading("title", text="Thread Title / Keyword")
        self.tree.heading("bm_status", text="Bookmark Status")
        self.tree.heading("url", text="Live Post (Double-Click to Open)")

        self.tree.column("check", width=36, anchor="center", stretch=False)
        self.tree.column("id", width=38, anchor="center", stretch=False)
        self.tree.column("status", width=105, anchor="center", stretch=False)
        self.tree.column("title", width=310, anchor="w")
        self.tree.column("bm_status", width=150, anchor="center", stretch=False)
        self.tree.column("url", width=290, anchor="w")

        # Keep all rows in the Bookmarking Queue on the clean default/blank dark background (#050811)
        self.tree.tag_configure("done", background="#050811", foreground="#e2e8f0")
        self.tree.tag_configure("posting", background="#050811", foreground="#38bdf8")
        self.tree.tag_configure("failed", background="#050811", foreground="#fda4af")

        vsb = ttk.Scrollbar(t_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Button-3>", self.on_tree_right_click)

        self.ctx_menu = tk.Menu(self.root, tearoff=0, bg="#0f172a", fg="white", activebackground="#2563eb")
        self.ctx_menu.add_command(label="🎯 Post Selected Keywords Only (☑)", command=self.post_selected_keywords)
        self.ctx_menu.add_command(label="☑ Toggle Select on This Row", command=self.ctx_toggle_row)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="🔄 Retry This Keyword Only", command=self.ctx_retry_row)
        self.ctx_menu.add_command(label="📋 Copy This Row's Live Link", command=self.ctx_copy_link)

        # Console Card
        c_card = tk.Frame(right_p, bg="#0d1424", highlightbackground="#1e293b", highlightthickness=1, padx=10, pady=6, height=130)
        c_card.pack(fill="x")
        c_card.pack_propagate(False)
        c_top = tk.Frame(c_card, bg="#0d1424")
        c_top.pack(fill="x", pady=(0, 3))
        tk.Label(c_top, text="Live Activity, Bookmarking & Ban Detector Console", bg="#0d1424", fg="white", font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Button(c_top, text="Clear Console", bg="#1e293b", fg="#e2e8f0", font=("Segoe UI", 7, "bold"), relief="flat", padx=6, command=self.clear_console).pack(side="right")

        self.txt_console = scrolledtext.ScrolledText(c_card, bg="#050811", fg="#34d399", font=("Consolas", 9), wrap="word", height=5)
        self.txt_console.pack(fill="both", expand=True)

    # --------------------------------------------------------------------------
    # TAB 2: BOOKMARKING ENGINE & HATENA SHIELD
    # --------------------------------------------------------------------------
    def build_tab2(self):
        f_top = tk.LabelFrame(self.pane2, text=" 🔖 Bookmark Any Custom Link(s) or Queue Links ", bg="#0d1424", fg="#34d399", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        f_top.pack(fill="x", pady=(0, 6))

        rt = tk.Frame(f_top, bg="#0d1424")
        rt.pack(fill="x", pady=(0, 4))
        self.lbl_custom_bm_cnt = tk.Label(rt, text="🔗 Paste Any Link(s) to Bookmark (0 links ready):", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 8, "bold"))
        self.lbl_custom_bm_cnt.pack(side="left")
        tk.Button(rt, text="🧹 Clear", bg="#1e293b", fg="#fda4af", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, command=self.clear_custom_bm).pack(side="right", padx=2)
        tk.Button(rt, text="📥 Load Queue Links", bg="#1e293b", fg="#a7f3d0", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, command=self.load_queue_to_custom_bm).pack(side="right", padx=2)
        tk.Button(rt, text="📋 Paste Links", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, command=self.paste_custom_bm).pack(side="right", padx=2)

        self.txt_custom_bm = scrolledtext.ScrolledText(f_top, bg="#050811", fg="#38bdf8", insertbackground="white", font=("Consolas", 9), height=4)
        self.txt_custom_bm.pack(fill="x", pady=2)
        self.txt_custom_bm.bind("<KeyRelease>", lambda e: self.update_custom_bm_count())

        rb = tk.Frame(f_top, bg="#0d1424")
        rb.pack(fill="x", pady=4)
        tk.Label(rb, text="Optional Title (Blank=Auto):", bg="#0d1424", fg="#94a3b8", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_custom_bm_title = tk.Entry(rb, bg="#050811", fg="white", insertbackground="white", relief="solid", bd=1, width=24)
        self.entry_custom_bm_title.pack(side="left", padx=6)
        tk.Label(rb, text="Delay (s):", bg="#0d1424", fg="#94a3b8", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(4, 2))
        self.spin_bm_delay = tk.Spinbox(rb, from_=0, to=120, width=4, bg="#050811", fg="white", buttonbackground="#1e293b", relief="solid", bd=1, font=("Segoe UI", 8, "bold"))
        self.spin_bm_delay.delete(0, "end")
        self.spin_bm_delay.insert(0, "2")
        self.spin_bm_delay.pack(side="left", padx=(0, 6))
        tk.Button(rb, text="🔖 BOOKMARK PASTED LINKS NOW", bg="#7c3aed", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=12, pady=4, command=self.start_custom_links_bookmarking).pack(side="left", padx=4)
        tk.Button(rb, text="🚀 BOOKMARK QUEUE LINKS", bg="#059669", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=12, pady=4, command=self.start_queue_bookmarking).pack(side="left", padx=4)
        tk.Button(rb, text="🛑 STOP", bg="#e11d48", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=4, command=self.stop_all).pack(side="left", padx=4)

        f_pre = tk.LabelFrame(self.pane2, text=" 🛡 1-Click Bookmarking Mode Presets ", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        f_pre.pack(fill="x", pady=(0, 6))
        tk.Button(f_pre, text="🛡 PROBOARDS SAFE MODE (Hatena OFF, Others ON)", bg="#d97706", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", pady=5, command=lambda: self.apply_bm_preset("no_hatena")).pack(side="left", fill="x", expand=True, padx=4)
        tk.Button(f_pre, text="✅ ALL BOOKMARKING ON", bg="#059669", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", pady=5, command=lambda: self.apply_bm_preset("all_on")).pack(side="left", fill="x", expand=True, padx=4)
        tk.Button(f_pre, text="🚫 ALL BOOKMARKING OFF", bg="#e11d48", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", pady=5, command=lambda: self.apply_bm_preset("all_off")).pack(side="left", fill="x", expand=True, padx=4)

        f_split = tk.Frame(self.pane2, bg="#070b14")
        f_split.pack(fill="both", expand=True)

        f_l = tk.LabelFrame(f_split, text=" 📋 Active Bookmarking Accounts (Saved Permanently in bookmarks_curl_vault.json) ", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        f_l.pack(side="left", fill="both", expand=True, padx=(0, 4))
        self.bm_chk_frame = tk.Frame(f_l, bg="#050811", padx=8, pady=8)
        self.bm_chk_frame.pack(fill="both", expand=True)
        tk.Button(f_l, text="💾 SAVE ACTIVE ACCOUNTS", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", pady=5, command=self.save_bm_toggles).pack(fill="x", pady=(6, 0))

        f_r = tk.LabelFrame(f_split, text=" 🔧 View / Update Bookmarking Account cURL (Auto-Saves!) ", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        f_r.pack(side="left", fill="both", expand=True, padx=(4, 0))

        r_acc = tk.Frame(f_r, bg="#0d1424")
        r_acc.pack(fill="x", pady=2)
        tk.Label(r_acc, text="Account:", bg="#0d1424", fg="#94a3b8", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.combo_bm_acc = ttk.Combobox(r_acc, state="readonly", width=15)
        self.combo_bm_acc.pack(side="left", padx=4)
        self.combo_bm_acc.bind("<<ComboboxSelected>>", lambda e: self.on_bm_account_selected())

        tk.Label(r_acc, text="New Name:", bg="#0d1424", fg="#94a3b8", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(6, 2))
        self.entry_bm_new = tk.Entry(r_acc, bg="#050811", fg="white", insertbackground="white", relief="solid", bd=1, width=13)
        self.entry_bm_new.pack(side="left", padx=2)

        tk.Label(r_acc, text="Type:", bg="#0d1424", fg="#94a3b8", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(6, 2))
        self.combo_bm_type = ttk.Combobox(r_acc, values=["SINGLE", "BULK"], state="readonly", width=8)
        self.combo_bm_type.set("SINGLE")
        self.combo_bm_type.pack(side="left", padx=2)

        tk.Button(r_acc, text="🗑 Delete Site", bg="#e11d48", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, command=self.delete_bm_account).pack(side="right", padx=2)
        tk.Button(r_acc, text="📋 Paste cURL", bg="#1e293b", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=self.paste_bm_curl).pack(side="right", padx=2)
        tk.Button(r_acc, text="➕ New", bg="#0284c7", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=self.prepare_new_bm_account).pack(side="right", padx=2)

        self.txt_bm_curl = scrolledtext.ScrolledText(f_r, bg="#050811", fg="#38bdf8", insertbackground="white", font=("Consolas", 8), height=9)
        self.txt_bm_curl.pack(fill="both", expand=True, pady=4)
        self.txt_bm_curl.bind("<KeyRelease>", lambda e: self.on_bm_curl_modified())
        self.txt_bm_curl.bind("<<Paste>>", lambda e: self.root.after(60, self.on_bm_curl_modified))

        tk.Button(f_r, text="💾 Save Bookmark Account cURL (Permanent)", bg="#059669", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", pady=6, command=self.save_bm_account_curl).pack(fill="x")
        self.lbl_bm_status = tk.Label(self.pane2, text="Ready. All Bookmark cURLs are permanently stored in bookmarks_curl_vault.json.", bg="#070b14", fg="#34d399", font=("Segoe UI", 8, "bold"))
        self.lbl_bm_status.pack(anchor="w", pady=2)

    # --------------------------------------------------------------------------
    # TAB 3: ZOOM LIVE AUTO-SENDER
    # --------------------------------------------------------------------------
    def build_tab3(self):
        fz = tk.LabelFrame(self.pane3, text=" 💬 Zoom Live Auto-Sender, Countdown Timer & Sheet Column Watcher ", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
        fz.pack(fill="both", expand=True)

        rz1 = tk.Frame(fz, bg="#0d1424")
        rz1.pack(fill="x", pady=(2, 4))
        tk.Label(rz1, text="Read Column (AUTO=Latest):", bg="#0d1424", fg="white", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_zoom_read_col = tk.Entry(rz1, bg="#050811", fg="#38bdf8", insertbackground="white", width=7, font=("Consolas", 9, "bold"), justify="center")
        self.entry_zoom_read_col.insert(0, "AUTO")
        self.entry_zoom_read_col.pack(side="left", padx=5)
        tk.Button(rz1, text="📥 Load Column Links from Sheet", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=4, command=self.load_column_links_from_sheet).pack(side="left", padx=3)
        tk.Button(rz1, text="✂️ Split Glued URLs & Copy All", bg="#0d9488", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=4, command=self.split_and_copy_zoom_links).pack(side="left", padx=3)

        tk.Label(rz1, text="⏱ Start Countdown (s):", bg="#0d1424", fg="#fbbf24", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(12, 2))
        self.spin_zoom_countdown = tk.Spinbox(rz1, from_=0, to=120, width=5, bg="#050811", fg="white", font=("Segoe UI", 9, "bold"))
        self.spin_zoom_countdown.delete(0, "end")
        self.spin_zoom_countdown.insert(0, "1")
        self.spin_zoom_countdown.pack(side="left", padx=2)

        tk.Label(rz1, text="⏱ Delay / Link (s):", bg="#0d1424", fg="#34d399", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(10, 2))
        self.spin_zoom_delay = tk.Spinbox(rz1, from_=0, to=120, width=5, bg="#050811", fg="white", font=("Segoe UI", 9, "bold"))
        self.spin_zoom_delay.delete(0, "end")
        self.spin_zoom_delay.insert(0, "1")
        self.spin_zoom_delay.pack(side="left", padx=2)

        rz2 = tk.Frame(fz, bg="#080c17", highlightbackground="#172033", highlightthickness=1, padx=8, pady=5)
        rz2.pack(fill="x", pady=4)
        tk.Label(rz2, text="📊 Target Sheet Col (AUTO=Sequential):", bg="#080c17", fg="#fbbf24", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_zoom_save_col = tk.Entry(rz2, bg="#050811", fg="#fbbf24", insertbackground="white", width=6, font=("Consolas", 9, "bold"), justify="center")
        self.entry_zoom_save_col.insert(0, "AUTO")
        self.entry_zoom_save_col.pack(side="left", padx=(4, 4))
        tk.Button(rz2, text="⏭ New Empty Col", bg="#1e293b", fg="#fbbf24", font=("Segoe UI", 8, "bold"), relief="flat", padx=7, pady=4, command=self.start_new_zoom_sheet_column).pack(side="left", padx=(0, 10))

        self.btn_zoom_both = tk.Button(rz2, text="🚀 PUT IN SHEET + SEND TO ZOOM", bg="#10b981", fg="#042f2e", font=("Segoe UI", 8, "bold"), relief="flat", padx=12, pady=4, command=lambda: self.run_zoom_sender_action(do_zoom=True, do_sheet=True))
        self.btn_zoom_both.pack(side="left", padx=3)
        self.btn_zoom_only = tk.Button(rz2, text="💬 SEND TO ZOOM ONLY", bg="#7c3aed", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=12, pady=4, command=lambda: self.run_zoom_sender_action(do_zoom=True, do_sheet=False))
        self.btn_zoom_only.pack(side="left", padx=3)
        self.btn_sheet_only = tk.Button(rz2, text="📄 PUT IN SHEET ONLY", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=12, pady=4, command=lambda: self.run_zoom_sender_action(do_zoom=False, do_sheet=True))
        self.btn_sheet_only.pack(side="left", padx=3)
        tk.Button(rz2, text="🛑 STOP", bg="#e11d48", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=4, command=self.stop_all).pack(side="left", padx=3)
        self.lbl_zoom_timer_banner = tk.Label(rz2, text="⏳ Ready", bg="#064e3b", fg="#6ee7b7", font=("Consolas", 9, "bold"), padx=8, pady=3)
        self.lbl_zoom_timer_banner.pack(side="right", padx=4)

        self.txt_zoom_links = scrolledtext.ScrolledText(fz, bg="#050811", fg="#38bdf8", insertbackground="white", font=("Consolas", 9), height=16)
        self.txt_zoom_links.pack(fill="both", expand=True, pady=6)
        self.lbl_zoom_msg = tk.Label(fz, text="Ready.", bg="#0d1424", fg="#34d399", font=("Segoe UI", 9, "bold"))
        self.lbl_zoom_msg.pack(anchor="w")

    # --------------------------------------------------------------------------
    # TAB 4: POST HISTORY & SAVED LINKS LOG (NEW!)
    # --------------------------------------------------------------------------
    def build_tab_history(self):
        fh = tk.LabelFrame(self.pane4, text=" 📜 Post History & Saved Links Log (Permanently Stored on Disk) ", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
        fh.pack(fill="both", expand=True)

        htop = tk.Frame(fh, bg="#0d1424")
        htop.pack(fill="x", pady=(0, 6))
        tk.Label(htop, text="🔍 Filter (Site / Keyword / Date):", bg="#0d1424", fg="#cbd5e1", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.entry_hist_filter = tk.Entry(htop, bg="#050811", fg="white", insertbackground="white", width=26, font=("Segoe UI", 9), relief="solid", bd=1)
        self.entry_hist_filter.pack(side="left", padx=6)
        self.entry_hist_filter.bind("<KeyRelease>", lambda e: self.render_history_table())

        self.lbl_hist_count = tk.Label(htop, text="(0 saved posts)", bg="#0d1424", fg="#34d399", font=("Segoe UI", 9, "bold"))
        self.lbl_hist_count.pack(side="left", padx=8)

        tk.Button(htop, text="🧹 Clear All History", bg="#e11d48", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, command=self.clear_all_history).pack(side="right", padx=3)
        tk.Button(htop, text="🗑 Delete Selected", bg="#4c0519", fg="#fda4af", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, command=self.delete_selected_history).pack(side="right", padx=3)
        tk.Button(htop, text="📥 Export CSV", bg="#0d9488", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, command=self.export_history_csv).pack(side="right", padx=3)
        tk.Button(htop, text="🔖 Bookmark Shown Links", bg="#7c3aed", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, command=self.bookmark_shown_history_links).pack(side="right", padx=3)
        tk.Button(htop, text="📋 Copy All Shown Links", bg="#2563eb", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", padx=10, pady=3, command=self.copy_shown_history_links).pack(side="right", padx=3)

        hwrap = tk.Frame(fh, bg="#060a12")
        hwrap.pack(fill="both", expand=True)
        hcols = ("num", "timestamp", "site", "col_row", "title", "url")
        self.hist_tree = ttk.Treeview(hwrap, columns=hcols, show="headings", height=18)
        self.hist_tree.heading("num", text="#")
        self.hist_tree.heading("timestamp", text="Date & Time")
        self.hist_tree.heading("site", text="Site / Domain")
        self.hist_tree.heading("col_row", text="Sheet Pos")
        self.hist_tree.heading("title", text="Keyword / Title")
        self.hist_tree.heading("url", text="Live Post Link (Click to Copy / Double-Click to Open)")

        self.hist_tree.column("num", width=45, anchor="center", stretch=False)
        self.hist_tree.column("timestamp", width=155, anchor="center", stretch=False)
        self.hist_tree.column("site", width=150, anchor="center", stretch=False)
        self.hist_tree.column("col_row", width=95, anchor="center", stretch=False)
        self.hist_tree.column("title", width=330, anchor="w")
        self.hist_tree.column("url", width=420, anchor="w")

        hvsb = ttk.Scrollbar(hwrap, orient="vertical", command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=hvsb.set)
        self.hist_tree.pack(side="left", fill="both", expand=True)
        hvsb.pack(side="right", fill="y")

        self.hist_tree.bind("<ButtonRelease-1>", self.on_hist_click)
        self.hist_tree.bind("<Double-1>", self.on_hist_double_click)

        self.lbl_hist_status = tk.Label(fh, text="Every posted link is automatically logged here with exact date & time.", bg="#0d1424", fg="#34d399", font=("Segoe UI", 8, "bold"))
        self.lbl_hist_status.pack(anchor="w", pady=(4, 0))

    def add_post_to_history(self, site_name, title, live_url, col_let="", row_num=""):
        now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        entry = {
            "timestamp": now_str,
            "site": site_name or self.active_site or extract_domain_from_curl(live_url),
            "col_row": f"Col {col_let}, Row {row_num}" if col_let else "-",
            "title": title,
            "url": live_url
        }
        self.history.insert(0, entry)
        if len(self.history) > 5000:
            self.history = self.history[:5000]
        self.save_all_vaults_to_disk(False)
        self.root.after(0, self.render_history_table)

    def get_filtered_history(self):
        q = (self.entry_hist_filter.get() if hasattr(self, "entry_hist_filter") else "").strip().lower()
        if not q:
            return list(enumerate(self.history))
        return [
            (idx, h) for idx, h in enumerate(self.history)
            if q in str(h.get("timestamp", "")).lower()
            or q in str(h.get("site", "")).lower()
            or q in str(h.get("title", "")).lower()
            or q in str(h.get("url", "")).lower()
            or q in str(h.get("col_row", "")).lower()
        ]

    def render_history_table(self):
        if not hasattr(self, "hist_tree"):
            return
        for k in self.hist_tree.get_children():
            self.hist_tree.delete(k)
        filtered = self.get_filtered_history()
        for disp_num, (orig_idx, h) in enumerate(filtered, start=1):
            self.hist_tree.insert("", "end", iid=str(orig_idx), values=(
                disp_num,
                h.get("timestamp", ""),
                h.get("site", ""),
                h.get("col_row", "-"),
                h.get("title", ""),
                h.get("url", "")
            ))
        self.lbl_hist_count.config(text=f"({len(filtered)} / {len(self.history)} posts)")

    def on_hist_click(self, event):
        row_id = self.hist_tree.identify_row(event.y)
        if row_id:
            idx = int(row_id)
            if 0 <= idx < len(self.history):
                u = self.history[idx].get("url", "")
                if u.startswith("http"):
                    self.root.clipboard_clear()
                    self.root.clipboard_append(u)
                    self.lbl_hist_status.config(text=f"📋 Copied History Link: {u}")

    def on_hist_double_click(self, event):
        row_id = self.hist_tree.identify_row(event.y)
        if row_id:
            idx = int(row_id)
            if 0 <= idx < len(self.history):
                u = self.history[idx].get("url", "")
                if u.startswith("http"):
                    webbrowser.open(u)

    def copy_shown_history_links(self):
        filtered = self.get_filtered_history()
        urls = [h.get("url", "") for _, h in filtered if str(h.get("url", "")).startswith("http")]
        if not urls:
            messagebox.showinfo("No Links", "No history links to copy.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(urls))
        self.lbl_hist_status.config(text=f"📋 Copied {len(urls)} history link(s) to clipboard!")

    def bookmark_shown_history_links(self):
        filtered = self.get_filtered_history()
        urls = [h.get("url", "") for _, h in filtered if str(h.get("url", "")).startswith("http")]
        if not urls:
            return
        self.txt_custom_bm.delete("1.0", "end")
        self.txt_custom_bm.insert("1.0", "\n".join(urls))
        self.update_custom_bm_count()
        self.switch_tab(2)

    def delete_selected_history(self):
        sel = self.hist_tree.selection()
        if not sel:
            return
        to_del = {int(x) for x in sel}
        self.history = [h for i, h in enumerate(self.history) if i not in to_del]
        self.save_all_vaults_to_disk(False)
        self.render_history_table()

    def clear_all_history(self):
        if messagebox.askyesno("Clear History", "Clear all saved post history?"):
            self.history = []
            self.save_all_vaults_to_disk(False)
            self.render_history_table()

    def export_history_csv(self):
        if not self.history:
            return
        out_p = os.path.join(get_script_dir(), "post_history_export.csv")
        try:
            with open(out_p, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Date & Time", "Site", "Sheet Position", "Keyword / Title", "Live URL"])
                for h in self.history:
                    w.writerow([h.get("timestamp", ""), h.get("site", ""), h.get("col_row", ""), h.get("title", ""), h.get("url", "")])
            self.lbl_hist_status.config(text=f"✅ Exported {len(self.history)} rows to {out_p}")
            messagebox.showinfo("Exported CSV", f"Saved history to:\n{out_p}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    # --------------------------------------------------------------------------
    # TAB 5: MOBILE REMOTE CONTROL
    # --------------------------------------------------------------------------
    def build_tab4(self):
        fm = tk.LabelFrame(self.pane5, text=" 📱 Mobile Remote Control (Wi-Fi & 4G/5G) ", bg="#0d1424", fg="#34d399", font=("Segoe UI", 10, "bold"), padx=14, pady=14)
        fm.pack(fill="both", expand=True)
        lan_ip = get_lan_ip()
        tk.Label(fm, text=f"🏠 Local Wi-Fi Control URL:  http://{lan_ip}:5050", bg="#0d1424", fg="#38bdf8", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=6)
        tk.Label(fm, text="💾 Permanent Vault Files Location:", bg="#0d1424", fg="#fbbf24", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(12, 4))
        tk.Label(fm, text=f"   • Folder Vault: {get_script_dir()}\\sites_curl_vault.json & bookmarks_curl_vault.json\n   • Permanent User Backup: {get_backup_dir()}", bg="#0d1424", fg="#cbd5e1", font=("Consolas", 9), justify="left").pack(anchor="w")

    # --------------------------------------------------------------------------
    # STATE <-> UI POPULATION
    # --------------------------------------------------------------------------
    def populate_ui_from_state(self):
        self.entry_main_sheet.delete(0, "end")
        self.entry_main_sheet.insert(0, self.main_sheet_url)
        self.entry_content_sheet.delete(0, "end")
        self.entry_content_sheet.insert(0, self.content_sheet_url)
        self.entry_apps_script.delete(0, "end")
        self.entry_apps_script.insert(0, self.apps_script_url)
        self.entry_ai_key.delete(0, "end")
        self.entry_ai_key.insert(0, self.ai_key)
        self.spin_delay.delete(0, "end")
        self.spin_delay.insert(0, str(self.post_delay))
        self.spin_bm_delay.delete(0, "end")
        self.spin_bm_delay.insert(0, str(self.bm_delay))
        self.entry_sheet_col.delete(0, "end")
        self.entry_sheet_col.insert(0, self.sheet_col)
        if self.chrome_profile:
            self.combo_profile.set(self.chrome_profile)
        self.var_engine_mode.set(self.engine_mode if self.engine_mode in ("curl", "browser") else "curl")

        self.refresh_sites_combo()
        self.render_bookmark_checkboxes()
        self.render_queue_table()
        self.render_history_table()

    def log_console(self, msg):
        ts = datetime.now().strftime("%I:%M:%S %p")
        line = f"[{ts}] {msg}\n"
        def _app():
            self.txt_console.insert("end", line)
            self.txt_console.see("end")
        self.root.after(0, _app)

    def clear_console(self):
        self.txt_console.delete("1.0", "end")

    def copy_sheet_script(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(UPDATED_CODE_GS)
        messagebox.showinfo("Copied Code.gs", "✅ Copied Updated_Google_Sheet_Code.gs to clipboard!\n\n(Note: v10.2 also fixes Row 22 / #20 automatically even before you update Code.gs!)")

    # --------------------------------------------------------------------------
    # POSTING SITES VAULT HANDLERS
    # --------------------------------------------------------------------------
    def _smart_update_target_url_entry(self, candidate_url, allow_domain_switch=False):
        """
        Updates `self.entry_target_url` ONLY when appropriate, NEVER overwriting a specific
        user-entered composer path (e.g. `https://forum.buffed.de/forums/.../post-thread`)
        with a bare root homepage (`https://forum.buffed.de/`) or a tracker URL!
        """
        cand = (candidate_url or "").strip()
        if not cand or not cand.startswith("http") or is_ad_or_tracker_url(cand):
            return
        cur = self.entry_target_url.get().strip()
        if not cur or not cur.startswith("http"):
            self.entry_target_url.delete(0, "end")
            self.entry_target_url.insert(0, cand)
            return

        cur_path = (urllib.parse.urlparse(cur).path or "").strip("/")
        cand_path = (urllib.parse.urlparse(cand).path or "").strip("/")

        if is_same_base_domain(cur, cand):
            # Same domain: if user already has a specific path (like /forums/.../post-thread), NEVER overwrite it!
            if not cur_path and cand_path:
                self.entry_target_url.delete(0, "end")
                self.entry_target_url.insert(0, cand)
            return

        if allow_domain_switch:
            self.entry_target_url.delete(0, "end")
            self.entry_target_url.insert(0, cand)

    def refresh_sites_combo(self):
        names = sorted(self.sites.keys())
        self.combo_sites["values"] = names
        if self.active_site not in self.sites:
            self.active_site = names[0] if names else ""
        self.combo_sites.set(self.active_site)

        if self.active_site and self.active_site in self.sites:
            c = self.sites[self.active_site]
            dom = extract_domain_from_curl(c)
            self.lbl_active_site.config(text=f"Active Site: {self.active_site}")
            self.lbl_active_domain.config(text=f"Endpoint: {dom}")
            self.entry_site_name.delete(0, "end")
            self.entry_site_name.insert(0, self.active_site)
            self.txt_curl.delete("1.0", "end")
            self.txt_curl.insert("1.0", c)

            saved_u = (self.site_urls.get(self.active_site) or "").strip()
            derived_u = derive_composer_url_from_curl(c)
            # If saved_u is just a bare root homepage (path == '/') while derived_u has a real path, prefer derived_u
            if saved_u and (urllib.parse.urlparse(saved_u).path or "").strip("/") == "" and derived_u and is_same_base_domain(saved_u, derived_u):
                comp_u = derived_u
            else:
                comp_u = saved_u or derived_u

            # If the user already typed a specific URL for this same domain in entry_target_url, preserve it!
            cur_u = self.entry_target_url.get().strip()
            if cur_u and is_same_base_domain(cur_u, comp_u or dom) and (urllib.parse.urlparse(cur_u).path or "").strip("/"):
                self.site_urls[self.active_site] = cur_u
            elif comp_u:
                self.entry_target_url.delete(0, "end")
                self.entry_target_url.insert(0, comp_u)
        else:
            self.lbl_active_site.config(text="Active Site: None")
            self.lbl_active_domain.config(text="Endpoint: -")

    def reset_queue_statuses(self):
        for r in self.rows:
            r["status"] = "Pending"
            r["bookmark_status"] = "-"
            r["url"] = "-"
            r["selected"] = True
        self.save_all_vaults_to_disk(False)
        self.render_queue_table()

    def paste_target_site_url(self):
        cb = self._get_any_clipboard_text()
        if cb:
            self.entry_target_url.delete(0, "end")
            self.entry_target_url.insert(0, cb)

    def on_site_selected(self):
        s = self.combo_sites.get().strip()
        if s and s in self.sites:
            self.active_site = s
            if not getattr(self, "is_posting", False):
                for r in getattr(self, "rows", []):
                    r["status"] = "Pending"
                    r["bookmark_status"] = "-"
                    r["url"] = "-"
            # Allow target URL box to switch to the newly selected site's URL
            saved_u = (self.site_urls.get(s) or "").strip()
            derived_u = derive_composer_url_from_curl(self.sites[s])
            if saved_u and (urllib.parse.urlparse(saved_u).path or "").strip("/"):
                self.entry_target_url.delete(0, "end")
                self.entry_target_url.insert(0, saved_u)
            elif derived_u:
                self.entry_target_url.delete(0, "end")
                self.entry_target_url.insert(0, derived_u)
            self.refresh_sites_combo()
            if hasattr(self, "tree"):
                self.render_queue_table()
            self.save_all_vaults_to_disk(False)
            self.log_console(f"⚡ Activated Site '{s}' ({extract_domain_from_curl(self.sites[s])})")

    def on_posting_curl_modified(self):
        raw = self.txt_curl.get("1.0", "end").strip()
        if "curl" in raw.lower():
            dom = extract_domain_from_curl(raw)
            if dom != "Unknown Site" and not is_ad_or_tracker_url(dom):
                self.lbl_active_domain.config(text=f"Endpoint: {dom}")
                cur_name = self.entry_site_name.get().strip()
                old_dom = extract_domain_from_curl(self.sites.get(cur_name, "")) if cur_name in self.sites else ""
                if not cur_name or (old_dom and old_dom != "Unknown Site" and old_dom != dom):
                    clean_s = derive_clean_site_name_from_domain(dom)
                    self.entry_site_name.delete(0, "end")
                    self.entry_site_name.insert(0, clean_s)
                comp_u = derive_composer_url_from_curl(raw)
                if comp_u:
                    self._smart_update_target_url_entry(comp_u, allow_domain_switch=False)

    def save_site_to_vault(self, activate_now=True):
        name = self.entry_site_name.get().strip()
        raw = self.txt_curl.get("1.0", "end").strip()
        tu = self.entry_target_url.get().strip()
        if "curl" not in raw.lower():
            messagebox.showwarning("Invalid cURL", "Please paste a valid cURL command into Box 3 first.")
            return
        dom = extract_domain_from_curl(raw)
        if is_ad_or_tracker_url(dom):
            self.log_console(f"🛡️ Blocked tracker/ad cURL ({dom}) from being saved to Vault.")
            return
        old_dom = extract_domain_from_curl(self.sites.get(name, "")) if name in self.sites else ""
        if not name or (old_dom and old_dom != "Unknown Site" and dom != "Unknown Site" and old_dom != dom):
            name = derive_clean_site_name_from_domain(dom)
            self.entry_site_name.delete(0, "end")
            self.entry_site_name.insert(0, name)
        self.sites[name] = raw
        derived_comp = derive_composer_url_from_curl(raw)
        # Preserve the user's specific Target URL if it matches this domain and has a path!
        if tu and is_same_base_domain(tu, dom) and (urllib.parse.urlparse(tu).path or "").strip("/"):
            self.site_urls[name] = tu
        elif derived_comp:
            self.site_urls[name] = derived_comp
        elif tu:
            self.site_urls[name] = tu
        if name in self.deleted_sites:
            self.deleted_sites.remove(name)
        if activate_now:
            self.active_site = name
        self.save_all_vaults_to_disk(True)
        self.refresh_sites_combo()
        self.log_console(f"💾 Saved {'& Activated ' if activate_now else ''}Site '{name}' in sites_curl_vault.json!")

    def delete_selected_site(self):
        s = self.combo_sites.get().strip() or self.entry_site_name.get().strip()
        if not s or s not in self.sites:
            return
        self.sites.pop(s, None)
        self.site_urls.pop(s, None)
        if s not in self.deleted_sites:
            self.deleted_sites.append(s)
        self.active_site = sorted(self.sites.keys())[0] if self.sites else ""
        self.save_all_vaults_to_disk(False)
        self.refresh_sites_combo()
        self.log_console(f"🗑️ Permanently deleted site '{s}' from sites_curl_vault.json.")

    def clear_posting_curl_box(self):
        self.entry_site_name.delete(0, "end")
        self.txt_curl.delete("1.0", "end")

    @staticmethod
    def _win32_get_clipboard_text():
        """Reads CF_UNICODETEXT directly from the Windows OS clipboard (works even when app is in background)."""
        if os.name != "nt":
            return ""
        try:
            import ctypes
            CF_UNICODETEXT = 13
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            user32.GetClipboardData.argtypes = [ctypes.c_uint]
            user32.GetClipboardData.restype = ctypes.c_void_p
            kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
            kernel32.GlobalLock.restype = ctypes.c_wchar_p
            kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]

            text_val = ""
            if user32.OpenClipboard(0):
                try:
                    if user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
                        h_data = user32.GetClipboardData(CF_UNICODETEXT)
                        if h_data:
                            ptr = kernel32.GlobalLock(h_data)
                            if ptr:
                                text_val = str(ptr)
                                kernel32.GlobalUnlock(h_data)
                finally:
                    user32.CloseClipboard()
            return text_val
        except Exception:
            return ""

    def _get_any_clipboard_text(self):
        txt = self._win32_get_clipboard_text()
        if txt:
            return txt.strip()
        try:
            return (self.root.clipboard_get() or "").strip()
        except Exception:
            return ""

    def paste_posting_curl(self):
        """Immediately pastes cURL from clipboard and cancels any pending CDP auto-capture so old events never overwrite!"""
        try:
            self._cdp_session_id = getattr(self, "_cdp_session_id", 0) + 1
            self.listening_clipboard = False
            self.btn_open_capture.config(text="🎯 Open Chrome & Capture", bg="#d97706")

            cb = self._get_any_clipboard_text()
            if not cb:
                return
            self.last_clip_text = cb
            self.txt_curl.delete("1.0", "end")
            self.txt_curl.insert("1.0", cb)
            comp_u = derive_composer_url_from_curl(cb)
            if comp_u:
                self._smart_update_target_url_entry(comp_u, allow_domain_switch=True)
            dom = extract_domain_from_curl(cb)
            if dom != "Unknown Site":
                clean_s = derive_clean_site_name_from_domain(dom)
                cur_name = self.entry_site_name.get().strip()
                old_dom = extract_domain_from_curl(self.sites.get(cur_name, "")) if cur_name in self.sites else ""
                if not cur_name or (old_dom and old_dom != dom):
                    self.entry_site_name.delete(0, "end")
                    self.entry_site_name.insert(0, clean_s)
            self.save_site_to_vault(True)
            self.log_console(f"📋 Instantly pasted & saved cURL for {dom} as '{self.entry_site_name.get().strip()}'!")
        except Exception as e:
            self.log_console(f"⚠️ Paste error: {e}")

    def _ensure_buster_extension_installed(self, target_profile_dir=None):
        """
        Integrates 'Buster: Captcha Solver for Humans' (`mpbjkejclgfgadiemmefgebjfooflfhl` + Local Whisper Model)
        directly into Forum Studio Pro's Chrome profiles:
        1. Scans the user's Chrome/Edge User Data profiles for `Extensions/mpbjkejclgfgadiemmefgebjfooflfhl/<version>`.
        2. Copies the unpacked extension to `%LOCALAPPDATA%\ForumStudioPro\BusterWhisperExtension`.
        3. Clones the user's configured Buster storage & downloaded Whisper model (`Local Extension Settings`
           and `IndexedDB/chrome-extension_mpbjkejclgfgadiemmefgebjfooflfhl_0.indexeddb.*`) into `target_profile_dir/Default`
           so `Model: Whisper (Status: ready to use)` is immediately ready!
        4. If not yet installed in Chrome (e.g. on a friend's laptop), auto-downloads & unpacks Buster's CRX
           directly from the Chrome Web Store!
        """
        import shutil
        import io
        buster_id = "mpbjkejclgfgadiemmefgebjfooflfhl"
        dest_ext_dir = os.path.join(get_backup_dir(), "BusterWhisperExtension")
        manifest_dest = os.path.join(dest_ext_dir, "manifest.json")

        local_app_data = os.environ.get("LOCALAPPDATA") or os.path.expanduser(r"~\AppData\Local")
        browser_roots = [
            os.path.join(local_app_data, "Google", "Chrome", "User Data"),
            os.path.join(local_app_data, "Microsoft", "Edge", "User Data"),
            os.path.join(local_app_data, "BraveSoftware", "Brave-Browser", "User Data"),
        ]

        source_prof_dir = None
        source_ver_dir = None

        # 1. Search all local Chrome/Edge profiles (`Default`, `Profile 1`..`Profile 20`) for Buster
        for b_root in browser_roots:
            if not os.path.isdir(b_root):
                continue
            try:
                for entry in os.listdir(b_root):
                    p_cand = os.path.join(b_root, entry)
                    ext_base = os.path.join(p_cand, "Extensions", buster_id)
                    if os.path.isdir(ext_base):
                        vers = sorted(os.listdir(ext_base), reverse=True)
                        for v in vers:
                            v_dir = os.path.join(ext_base, v)
                            if os.path.isfile(os.path.join(v_dir, "manifest.json")):
                                source_prof_dir = p_cand
                                source_ver_dir = v_dir
                                break
                    if source_ver_dir:
                        break
            except Exception:
                pass
            if source_ver_dir:
                break

        # 2. Copy Buster unpacked extension into `dest_ext_dir` if not already staged
        if source_ver_dir and not os.path.isfile(manifest_dest):
            try:
                if os.path.exists(dest_ext_dir):
                    shutil.rmtree(dest_ext_dir, ignore_errors=True)
                shutil.copytree(source_ver_dir, dest_ext_dir)
                meta_d = os.path.join(dest_ext_dir, "_metadata")
                if os.path.exists(meta_d):
                    shutil.rmtree(meta_d, ignore_errors=True)
            except Exception:
                pass

        # 3. Fallback: If on a friend's laptop without Buster installed yet, download & unpack official Buster CRX
        if not os.path.isfile(manifest_dest):
            try:
                crx_url = (
                    "https://clients2.google.com/service/update2/crx?"
                    f"response=redirect&prodversion=124.0.0.0&acceptformat=crx2,crx3&x=id%3D{buster_id}%26uc"
                )
                req = urllib.request.Request(crx_url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0"})
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                with urllib.request.urlopen(req, context=ctx, timeout=12) as r:
                    crx_bytes = r.read()
                pk_idx = crx_bytes.find(b"PK\x03\x04")
                if pk_idx != -1:
                    os.makedirs(dest_ext_dir, exist_ok=True)
                    with zipfile.ZipFile(io.BytesIO(crx_bytes[pk_idx:])) as zf:
                        zf.extractall(dest_ext_dir)
                    meta_d = os.path.join(dest_ext_dir, "_metadata")
                    if os.path.exists(meta_d):
                        shutil.rmtree(meta_d, ignore_errors=True)
            except Exception:
                pass

        # 4. Clone the user's Buster Settings & Local Whisper Model Cache (`IndexedDB` & `Local Extension Settings`)
        #    into the tool's capture profile (`target_profile_dir/Default`) so Whisper is immediately 'ready to use'!
        if source_prof_dir and target_profile_dir:
            target_def = os.path.join(target_profile_dir, "Default")
            try:
                os.makedirs(target_def, exist_ok=True)
                for sub_cat in ("Local Extension Settings", "Sync Extension Settings", "Managed Extension Settings"):
                    src_s = os.path.join(source_prof_dir, sub_cat, buster_id)
                    dst_s = os.path.join(target_def, sub_cat, buster_id)
                    if os.path.isdir(src_s) and not os.path.exists(dst_s):
                        os.makedirs(os.path.dirname(dst_s), exist_ok=True)
                        shutil.copytree(src_s, dst_s)

                src_idb = os.path.join(source_prof_dir, "IndexedDB")
                dst_idb = os.path.join(target_def, "IndexedDB")
                if os.path.isdir(src_idb):
                    os.makedirs(dst_idb, exist_ok=True)
                    for item_name in os.listdir(src_idb):
                        if buster_id in item_name or "onnx" in item_name.lower() or "whisper" in item_name.lower() or "transformers" in item_name.lower():
                            s_item = os.path.join(src_idb, item_name)
                            d_item = os.path.join(dst_idb, item_name)
                            if os.path.isdir(s_item) and not os.path.exists(d_item):
                                shutil.copytree(s_item, d_item)
            except Exception:
                pass

        return dest_ext_dir if os.path.isfile(manifest_dest) else ""

    def _get_persistent_chrome_profile_dir(self):
        prof = (self.combo_profile.get() or "Profile 1").strip()
        prof_slug = re.sub(r"[^a-zA-Z0-9_-]+", "_", prof)
        base_dir = os.path.join(get_backup_dir(), "ChromeCaptureProfiles", prof_slug)
        default_sub = os.path.join(base_dir, "Default")
        try:
            os.makedirs(default_sub, exist_ok=True)
            pref_path = os.path.join(default_sub, "Preferences")
            prefs = {}
            if os.path.exists(pref_path):
                try:
                    with open(pref_path, "r", encoding="utf-8", errors="replace") as f:
                        prefs = json.load(f)
                except Exception:
                    prefs = {}
            prefs.setdefault("session", {})["restore_on_startup"] = 1
            prefs.setdefault("profile", {})["exit_type"] = "Normal"
            prefs["profile"]["exited_cleanly"] = True
            prefs["credentials_enable_service"] = True
            with open(pref_path, "w", encoding="utf-8") as f:
                json.dump(prefs, f)
        except Exception:
            pass
        return base_dir, prof

    def _get_profile_cookie_store_path(self):
        base_dir, _ = self._get_persistent_chrome_profile_dir()
        return os.path.join(base_dir, "persisted_forum_cookies.json")

    @staticmethod
    def _is_post_submission_request(req_url, method, post_data, headers, allowed_base_doms=None):
        """
        Universal Forum & Article Post Submission Classifier:
        1. Strictly enforces First-Party Domain Lock (`allowed_base_doms`) so third-party ad servers
           (`s2s.yieldlove-ad-serving.net`) and analytics trackers (`pbc.agma-analytics.de`) are 100% blocked!
        2. Rejects `Content-Type: text/plain` (`navigator.sendBeacon` trackers) and binary/media uploads.
        3. Rejects background auto-drafts (`/draft`, `autosave`), previews (`/preview`, `unfurl`),
           login/register/auth, search, consent/GDPR, and telemetry endpoints.
        4. Verifies genuine thread/post creation fields in form-urlencoded, multipart, or JSON payloads.
        """
        if (method or "").upper() not in ("POST", "PUT", "PATCH"):
            return False
        u_low = (req_url or "").lower().strip()
        if not u_low.startswith("http"):
            return False

        # 1. Block all known ad/analytics/tracker/CMP domains & subdomains immediately
        if is_ad_or_tracker_url(u_low):
            return False

        # 2. Enforce strict First-Party Base Domain Lock against the Target Site / Active Page!
        req_base = get_registrable_domain(u_low)
        if allowed_base_doms:
            valid_doms = {d for d in allowed_base_doms if d and not is_ad_or_tracker_url(d)}
            if valid_doms and req_base not in valid_doms:
                return False

        # 3. Check Content-Type header: reject text/plain (sendBeacon analytics!), octet-stream, media
        ct_val = ""
        if isinstance(headers, dict):
            for hk, hv in headers.items():
                if str(hk).lower() == "content-type":
                    ct_val = str(hv or "").lower().strip()
                    break
        if ct_val.startswith(("text/plain", "application/octet-stream", "application/x-protobuf", "image/", "audio/", "video/")):
            return False

        # 4. Reject excluded non-publish paths & query strings (including XenForo `/draft` auto-save while typing!)
        parsed_u = urllib.parse.urlparse(u_low)
        path_low = parsed_u.path or ""
        query_low = parsed_u.query or ""
        path_and_query = f"{path_low}?{query_low}" if query_low else path_low

        skip_segments = (
            "/draft", "/drafts", "save-draft", "savedraft", "autodraft", "auto-draft", "autosave", "auto-save",
            "/preview", "editor/to-bb-code", "editor/to-html", "unfurl", "oembed", "link-proxy",
            "content_check", "content-check", "duplicate_check", "duplicate-check",
            "title_check", "title-check", "slug_check", "slug-check", "name_check", "username_check", "email_check",
            "spellcheck", "spell-check", "/check", "_check", "-check", "/validate", "validation", "/verify",
            "/lookup", "/autocomplete", "/suggest", "/typeahead", "/filter", "/rate", "/rating", "/vote",
            "/login", "/signin", "/sign-in", "/auth", "/register", "/signup", "/sign-up",
            "/logout", "/password", "/lost-password", "/two-step", "/token", "/oauth", "/session",
            "/heartbeat", "/ping", "/poll", "/analytics", "/metrics", "/telemetry", "/beacon", "/collect",
            "/events", "/track", "/log", "/online", "/presence", "/keepalive",
            "/search", "/find-new", "/notifications", "/consent", "/cookie", "/gdpr", "/privacy", "/cmp",
            "/push", "/webpush", "/service-worker", "/attachments/upload", "/attachment/upload", "/upload"
        )
        if any(seg in path_and_query for seg in skip_segments):
            return False
        if re.search(r'/(?:v1|v2)/?(?:collect|track|event|events|beacon|log|metrics|ping)?$', path_low):
            return False

        pd = (post_data or "").strip()
        if not pd or len(pd) < 4:
            return False
        pd_low = pd.lower()

        # 5. Exclude login / registration / password forms
        if ("password" in pd_low or "passwd" in pd_low or "pass=" in pd_low) and not any(
            mk in pd_low for mk in ("subject=", "message=", "message_html=", "topic_title=", "post_title=", "\"body\"", "\"content\"")
        ):
            return False

        # 6. Check for known forum thread/post creation URL paths
        known_post_paths = (
            "/post-thread", "/thread/create", "/post/create", "/add-reply", "/threads/add",
            "newthread.php", "posting.php", "newreply.php", "/api/discussions",
            "wp-comments-post.php", "/topic/add", "/study-guides", "/novel/create", "/movie/new"
        )
        if any(kp in path_and_query for kp in known_post_paths):
            return True

        # 7. Inspect payload fields for Strong Forum Body keys OR (Title + Body) pair
        strong_body_indicators = (
            "message=", "message_html=", "post_text=", "post_body=", "post_content=",
            "thread_body=", "overview=", "work_caption=", "story=", "bbcode=", "markdown=", "raw=",
            "\"message\"", "\"message_html\"", "\"post_text\"", "\"post_body\"", "\"overview\"",
            "\"work_caption\"", "\"story\"", "\"raw\"",
            "name=\"message\"", "name=\"message_html\"", "name=\"post_text\"", "name=\"overview\"",
            "name=\"story\"", "name=\"raw\"", "name=\"text\""
        )
        if any(ind in pd_low for ind in strong_body_indicators):
            return True

        has_title_ind = any(
            t_ind in pd_low for t_ind in (
                "title=", "subject=", "topic_title=", "thread_title=", "post_title=", "headline=",
                "\"title\"", "\"subject\"", "\"topic_title\"", "\"thread_title\"", "\"headline\"",
                "name=\"title\"", "name=\"subject\"", "name=\"headline\""
            )
        )
        has_generic_body_ind = any(
            b_ind in pd_low for b_ind in (
                "body=", "content=", "text=", "description=", "comment=", "article=", "answer=",
                "\"body\"", "\"content\"", "\"text\"", "\"description\"", "\"comment\"", "\"article\"",
                "name=\"body\"", "name=\"content\"", "name=\"description\"", "name=\"comment\""
            )
        )
        if has_title_ind and has_generic_body_ind:
            return True

        return False

    def open_chrome_and_capture(self):
        u = self.entry_target_url.get().strip() or derive_composer_url_from_curl(self.txt_curl.get("1.0", "end"))
        if not u:
            messagebox.showwarning("Target URL", "Enter or paste a Target Site URL first.")
            return
        if not u.startswith("http"):
            u = "https://" + u
            self.entry_target_url.delete(0, "end")
            self.entry_target_url.insert(0, u)

        # Lock this exact Target URL for the capture session so Referer headers never overwrite it!
        self._locked_capture_target_url = u

        target_dom = extract_domain_from_curl(u)
        cur_name = self.entry_site_name.get().strip()
        old_dom = extract_domain_from_curl(self.sites.get(cur_name, "")) if cur_name in self.sites else ""
        if target_dom != "Unknown Site" and not is_ad_or_tracker_url(target_dom):
            if not cur_name or is_ad_or_tracker_url(old_dom) or (old_dom and old_dom != target_dom):
                new_clean_name = derive_clean_site_name_from_domain(target_dom)
                self.entry_site_name.delete(0, "end")
                self.entry_site_name.insert(0, new_clean_name)

        profile_dir, prof_label = self._get_persistent_chrome_profile_dir()
        buster_dir = self._ensure_buster_extension_installed(profile_dir)
        cookie_file = self._get_profile_cookie_store_path()
        prof_num = 1
        m_p = re.search(r"(\d+)", prof_label)
        if m_p:
            prof_num = int(m_p.group(1))
        debug_port = 9221 + prof_num  # Profile 1 -> 9222, Profile 2 -> 9223, Profile 3 -> 9224

        chrome_paths = [
            os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        ]
        opened = False
        for exe in chrome_paths:
            if os.path.exists(exe):
                try:
                    cmd_args = [
                        exe,
                        f"--user-data-dir={profile_dir}",
                        f"--remote-debugging-port={debug_port}",
                        "--remote-allow-origins=*",
                        "--disable-blink-features=AutomationControlled",
                        "--disable-features=IsolateOrigins,site-per-process",
                        "--disable-site-isolation-trials",
                        "--disable-web-security",
                        "--autoplay-policy=no-user-gesture-required",
                        "--no-first-run",
                        "--no-default-browser-check",
                        "--restore-last-session",
                    ]
                    if buster_dir:
                        cmd_args.append(f"--load-extension={buster_dir}")
                    cmd_args.append(u)
                    subprocess.Popen(
                        cmd_args,
                        creationflags=0x08000000 if os.name == "nt" else 0
                    )
                    opened = True
                    if buster_dir:
                        self.log_console("🤖 [Buster + Whisper] Loaded Buster Captcha Solver (Model: Whisper + Built-In Client App) into Chrome!")
                    break
                except Exception:
                    pass
        if not opened:
            webbrowser.open(u)

        self.last_clip_text = self._get_any_clipboard_text()
        self.listening_clipboard = True
        self.btn_open_capture.config(text="👂 Listening (Submit Test Post or Copy cURL)...", bg="#059669")
        self.log_console(
            f"🎯 Opened {u} ({prof_label}) — Locked to domain '{get_registrable_domain(u)}' (ad/tracker requests blocked). Submit 1 Test Post or Copy as cURL!"
        )
        self.poll_clipboard()
        self._start_cdp_auto_capture_thread(debug_port, cookie_file, u)

    def _start_cdp_auto_capture_thread(self, debug_port, cookie_file, initial_url):
        self._cdp_session_id = getattr(self, "_cdp_session_id", 0) + 1
        my_session = self._cdp_session_id
        initial_base_dom = get_registrable_domain(initial_url)

        def _cdp_worker():
            attached_ws = {}
            req_map = {}
            extra_hdrs = {}
            msg_id = 100
            cookies_restored = False
            last_cookie_sync = 0.0
            post_captured = False

            def _load_saved_cookies():
                if os.path.exists(cookie_file):
                    try:
                        with open(cookie_file, "r", encoding="utf-8") as f:
                            return json.load(f)
                    except Exception:
                        pass
                return []

            def _save_cookies_to_disk(c_list):
                if not c_list:
                    return
                try:
                    with open(cookie_file, "w", encoding="utf-8") as f:
                        json.dump(c_list, f, indent=2)
                except Exception:
                    pass

            try:
                startup_misses = 0
                last_target_scan = 0.0
                targets = []
                for _ in range(36000):
                    if my_session != getattr(self, "_cdp_session_id", 0):
                        break
                    now_loop = time.time()
                    if not attached_ws or (now_loop - last_target_scan > 1.0):
                        last_target_scan = now_loop
                        try:
                            req = urllib.request.Request(f"http://127.0.0.1:{debug_port}/json", headers={"User-Agent": "Mozilla/5.0"})
                            with urllib.request.urlopen(req, timeout=0.8) as r:
                                targets = json.loads(r.read().decode("utf-8", errors="replace"))
                        except Exception:
                            if attached_ws:
                                break
                            startup_misses += 1
                            if startup_misses > 25:
                                break
                            time.sleep(0.35)
                            continue

                    page_targets = [t for t in targets if t.get("type") == "page" and t.get("webSocketDebuggerUrl")]
                    allowed_base_doms = {initial_base_dom} if initial_base_dom else set()
                    cur_ui_u = self.entry_target_url.get().strip()
                    if cur_ui_u:
                        b_ui = get_registrable_domain(cur_ui_u)
                        if b_ui and not is_ad_or_tracker_url(b_ui):
                            allowed_base_doms.add(b_ui)

                    for pt in page_targets:
                        tid = pt.get("id")
                        ws_u = pt.get("webSocketDebuggerUrl")
                        if tid and ws_u and tid not in attached_ws:
                            try:
                                ws = _SimpleCDPWebSocket(ws_u, timeout=1.5)
                                ws.sock.settimeout(0.015)
                                msg_id += 1
                                ws.send_json({"id": msg_id, "method": "Network.enable", "params": {"maxPostDataSize": 655360}})
                                msg_id += 1
                                ws.send_json({"id": msg_id, "method": "Page.enable", "params": {}})
                                msg_id += 1
                                ws.send_json({"id": msg_id, "method": "Page.addScriptToEvaluateOnNewDocument", "params": {
                                    "source": """
                                    try {
                                        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                                        window.chrome = window.chrome || { runtime: {} };
                                        Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
                                        Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
                                    } catch (e) {}
                                    """
                                }})
                                attached_ws[tid] = ws

                                if not cookies_restored:
                                    saved_cks = _load_saved_cookies()
                                    if saved_cks:
                                        msg_id += 1
                                        ws.send_json({"id": msg_id, "method": "Network.setCookies", "params": {"cookies": saved_cks}})
                                    cookies_restored = True
                            except Exception:
                                pass

                    now_ts = time.time()
                    for tid, ws in list(attached_ws.items()):
                        if now_ts - last_cookie_sync > 1.8:
                            last_cookie_sync = now_ts
                            try:
                                msg_id += 1
                                ws.send_json({"id": msg_id, "method": "Network.getAllCookies"})
                                # Also run automatic in-browser CAPTCHA & ProBoards Human Modal solver!
                                msg_id += 1
                                ws.send_json({
                                    "id": msg_id,
                                    "method": "Runtime.evaluate",
                                    "params": {
                                        "expression": self._get_cdp_captcha_solver_js(),
                                        "returnByValue": True
                                    }
                                })
                            except Exception:
                                pass

                        for _ in range(400):
                            if my_session != getattr(self, "_cdp_session_id", 0):
                                break
                            try:
                                evt = ws.recv_json()
                            except Exception:
                                attached_ws.pop(tid, None)
                                break
                            if not evt:
                                break

                            res_obj = evt.get("result") or {}
                            eval_val = (res_obj.get("result") or {}).get("value") if isinstance(res_obj.get("result"), dict) else None
                            if isinstance(eval_val, dict) and eval_val.get("action") == "trusted_click":
                                tx = int(eval_val.get("x") or 0)
                                ty = int(eval_val.get("y") or 0)
                                lbl_c = str(eval_val.get("label") or "Captcha Button")
                                if tx > 0 and ty > 0:
                                    self._dispatch_cdp_trusted_click(ws, tx, ty)
                                    if "Buster" in lbl_c:
                                        self.log_console(f"🤖 [Buster Whisper + Client App] Clicked {lbl_c} at ({tx}, {ty})!")
                            elif isinstance(eval_val, dict) and eval_val.get("action") == "solve_audio" and eval_val.get("audio_url"):
                                aud_u = str(eval_val["audio_url"])
                                if aud_u != getattr(self, "_last_solved_audio_url", ""):
                                    self._last_solved_audio_url = aud_u
                                    self.log_console("🛡️ [Auto-CAPTCHA] Detected reCAPTCHA Audio Challenge in Chrome — transcribing audio...")
                                    def _solve_aud_bg(w_sock=ws, u_mp3=aud_u):
                                        phrase = transcribe_recaptcha_audio_url(u_mp3)
                                        if phrase:
                                            self.log_console(f"🛡️ [Auto-CAPTCHA] Transcribed audio: '{phrase}' — filling & verifying!")
                                            fill_js = self._get_cdp_fill_audio_js(phrase)
                                            try:
                                                w_sock.send_json({
                                                    "id": 9999,
                                                    "method": "Runtime.evaluate",
                                                    "params": {"expression": fill_js, "returnByValue": True}
                                                })
                                            except Exception:
                                                pass
                                    threading.Thread(target=_solve_aud_bg, daemon=True).start()

                            if "cookies" in res_obj and isinstance(res_obj["cookies"], list):
                                cks = res_obj["cookies"]
                                promoted = []
                                exp_30d = int(time.time() + 30 * 86400)
                                for c in cks:
                                    c_copy = dict(c)
                                    cdom_clean = (c_copy.get("domain") or "").lstrip(".")
                                    c_scheme = "https" if c_copy.get("secure", True) else "http"
                                    c_url = f"{c_scheme}://{cdom_clean}{c_copy.get('path', '/')}" if cdom_clean else initial_url
                                    c_copy["url"] = c_url
                                    if c_copy.get("session") or not c_copy.get("expires") or c_copy.get("expires", 0) <= 0:
                                        c_copy["expires"] = exp_30d
                                        c_copy.pop("session", None)
                                        try:
                                            msg_id += 1
                                            ws.send_json({"id": msg_id, "method": "Network.setCookie", "params": {
                                                "url": c_url,
                                                "name": c_copy.get("name"),
                                                "value": c_copy.get("value"),
                                                "domain": c_copy.get("domain"),
                                                "path": c_copy.get("path", "/"),
                                                "secure": bool(c_copy.get("secure", False)),
                                                "httpOnly": bool(c_copy.get("httpOnly", False)),
                                                "expires": exp_30d
                                            }})
                                        except Exception:
                                            pass
                                    promoted.append(c_copy)
                                _save_cookies_to_disk(promoted)

                            elif "postData" in res_obj and evt.get("id") in req_map:
                                orig_rid = req_map.pop(evt.get("id"), None)
                                if orig_rid and orig_rid in req_map and isinstance(req_map[orig_rid], dict):
                                    req_map[orig_rid]["postData"] = res_obj.get("postData") or req_map[orig_rid].get("postData", "")



                            method_name = evt.get("method") or ""
                            params = evt.get("params") or {}

                            if method_name == "Network.requestWillBeSentExtraInfo":
                                rid = params.get("requestId")
                                hdrs = params.get("headers") or {}
                                if rid and hdrs:
                                    extra_hdrs[rid] = hdrs

                            elif method_name == "Network.requestWillBeSent":
                                rid = params.get("requestId")
                                req_info = params.get("request") or {}
                                r_meth = (req_info.get("method") or "GET").upper()
                                if r_meth not in ("POST", "PUT", "PATCH"):
                                    continue
                                r_url = req_info.get("url") or ""
                                # Immediately skip any third-party ad/tracker or non-matching base domain!
                                if is_ad_or_tracker_url(r_url):
                                    continue
                                if allowed_base_doms and get_registrable_domain(r_url) not in allowed_base_doms:
                                    continue

                                r_hdrs = dict(req_info.get("headers") or {})
                                r_pd = req_info.get("postData") or ""
                                if not r_pd and isinstance(req_info.get("postDataEntries"), list):
                                    import base64
                                    parts_b = []
                                    for p_ent in req_info["postDataEntries"]:
                                        if p_ent.get("bytes"):
                                            try:
                                                parts_b.append(base64.b64decode(p_ent["bytes"]).decode("utf-8", errors="replace"))
                                            except Exception:
                                                pass
                                    r_pd = "".join(parts_b)

                                if rid:
                                    req_map[rid] = {
                                        "url": r_url,
                                        "method": r_meth,
                                        "headers": r_hdrs,
                                        "postData": r_pd,
                                        "has_post_data": bool(req_info.get("hasPostData")),
                                        "ts": time.time()
                                    }
                                    if not r_pd and req_info.get("hasPostData"):
                                        try:
                                            msg_id += 1
                                            req_map[msg_id] = rid
                                            ws.send_json({"id": msg_id, "method": "Network.getRequestPostData", "params": {"requestId": rid}})
                                        except Exception:
                                            pass
                                    try:
                                        msg_id += 1
                                        ws.send_json({"id": msg_id, "method": "Network.getAllCookies"})
                                    except Exception:
                                        pass



                        for rid_k, r_item in list(req_map.items()):
                            if my_session != getattr(self, "_cdp_session_id", 0):
                                break
                            if not isinstance(r_item, dict):
                                continue
                            age = time.time() - r_item.get("ts", 0)
                            if age > 2.5:
                                req_map.pop(rid_k, None)
                                continue
                            if not r_item.get("postData") and r_item.get("has_post_data") and age < 0.8:
                                continue
                            if age < 0.15 and rid_k not in extra_hdrs:
                                continue

                            merged_h = dict(r_item["headers"])
                            if rid_k in extra_hdrs:
                                for hk, hv in extra_hdrs[rid_k].items():
                                    if not hk.startswith(":"):
                                        merged_h[hk] = hv

                            if self._is_post_submission_request(
                                r_item["url"], r_item["method"], r_item["postData"], merged_h, allowed_base_doms=allowed_base_doms
                            ):
                                ck_map = {}
                                for k_h, v_h in list(merged_h.items()):
                                    if k_h.lower() == "cookie":
                                        for part in str(v_h).split(";"):
                                            if "=" in part:
                                                ck_n, ck_v = part.strip().split("=", 1)
                                                ck_map[ck_n.strip()] = ck_v.strip()
                                        merged_h.pop(k_h, None)

                                parsed_dom = (urllib.parse.urlparse(r_item["url"]).hostname or "").lower()
                                for c in _load_saved_cookies():
                                    cdom = (c.get("domain") or "").lstrip(".").lower()
                                    if cdom and (parsed_dom == cdom or parsed_dom.endswith("." + cdom)):
                                        cname = c.get("name")
                                        if cname and cname not in ck_map:
                                            ck_map[cname] = str(c.get("value") or "")
                                if ck_map:
                                    merged_h["Cookie"] = "; ".join(f"{k}={v}" for k, v in ck_map.items())

                                curl_lines = [f"curl '{r_item['url']}' \\", f"  -X '{r_item['method']}' \\"]
                                for hk, hv in merged_h.items():
                                    if hk.lower() in ("content-length", "host", "connection", "accept-encoding") or hk.startswith(":"):
                                        continue
                                    safe_hv = str(hv).replace("'", "'\\''")
                                    curl_lines.append(f"  -H '{hk}: {safe_hv}' \\")
                                safe_pd = str(r_item["postData"]).replace("'", "'\\''")
                                curl_lines.append(f"  --data-raw '{safe_pd}'")
                                built_curl = "\n".join(curl_lines)

                                cap_sig = (r_item["url"], r_item["postData"])
                                req_map.clear()
                                if cap_sig != getattr(self, "_last_cdp_cap_sig", None):
                                    self._last_cdp_cap_sig = cap_sig
                                    if my_session == getattr(self, "_cdp_session_id", 0):
                                        self.root.after(0, lambda c_str=built_curl, u_str=r_item["url"], ms=my_session: self._on_cdp_captured_curl(c_str, u_str, ms))
                                break

                    time.sleep(0.08)
            finally:
                for ws in list(attached_ws.values()):
                    ws.close()

        threading.Thread(target=_cdp_worker, daemon=True).start()

    def _on_cdp_captured_curl(self, curl_str, post_url, session_id=None):
        """Called on the main Tkinter thread when the CDP sniffer captures a Test Post in Chrome."""
        if session_id is not None and session_id != getattr(self, "_cdp_session_id", 0):
            return

        dom = extract_domain_from_curl(curl_str)
        if is_ad_or_tracker_url(dom):
            return

        self.btn_open_capture.config(text="✅ Captured Test Post cURL!", bg="#10b981")
        self.txt_curl.delete("1.0", "end")
        self.txt_curl.insert("1.0", curl_str)

        # Preserve the user's specific Target URL (e.g. https://forum.buffed.de/forums/.../post-thread)!
        locked_u = getattr(self, "_locked_capture_target_url", "") or self.entry_target_url.get().strip()
        comp = derive_composer_url_from_curl(curl_str)
        if locked_u and is_same_base_domain(locked_u, dom) and (urllib.parse.urlparse(locked_u).path or "").strip("/"):
            self.entry_target_url.delete(0, "end")
            self.entry_target_url.insert(0, locked_u)
        elif comp:
            self._smart_update_target_url_entry(comp, allow_domain_switch=True)

        cur_name = self.entry_site_name.get().strip()
        old_dom = extract_domain_from_curl(self.sites.get(cur_name, "")) if cur_name in self.sites else ""
        if not cur_name or is_ad_or_tracker_url(old_dom) or (old_dom and old_dom != "Unknown Site" and old_dom != dom):
            auto_name = derive_clean_site_name_from_domain(dom)
            self.entry_site_name.delete(0, "end")
            self.entry_site_name.insert(0, auto_name)

        self.save_site_to_vault(activate_now=True)
        s_name = self.entry_site_name.get().strip() or dom
        self.log_console(f"🎯 [AUTO-CAPTURED cURL] Captured Test Post in real time & saved '{s_name}' ({dom}) to Vault!")

    def poll_clipboard(self):
        if not self.listening_clipboard:
            return
        try:
            cb = self._get_any_clipboard_text()
            if cb and cb != self.last_clip_text and cb.lower().lstrip("\"'").startswith("curl") and "http" in cb.lower():
                self.last_clip_text = cb
                self.listening_clipboard = False
                self._cdp_session_id = getattr(self, "_cdp_session_id", 0) + 1
                self.btn_open_capture.config(text="✅ Copied cURL Grabbed!", bg="#10b981")
                self.paste_posting_curl()
                return
        except Exception:
            pass
        self.root.after(220, self.poll_clipboard)

    def show_target_browser(self):
        self.open_chrome_and_capture()

    # --------------------------------------------------------------------------
    # BOOKMARKING VAULT HANDLERS (bookmarks_curl_vault.json)
    # --------------------------------------------------------------------------
    def render_bookmark_checkboxes(self):
        for w in self.bm_chk_frame.winfo_children():
            w.destroy()
        self.bm_vars = []
        names = []
        for idx, b in enumerate(self.bookmarks):
            s_name = b.get("site", f"Account_{idx+1}")
            names.append(s_name)
            v = tk.BooleanVar(value=bool(b.get("active", True)))
            self.bm_vars.append((idx, v))
            is_hat = "hatena" in s_name.lower()
            row_f = tk.Frame(self.bm_chk_frame, bg="#050811")
            row_f.pack(fill="x", pady=1)
            cb = tk.Checkbutton(
                row_f,
                text=f"{s_name} ({b.get('type', 'SINGLE')}) {'⚠️ [Hatena]' if is_hat else ''}",
                variable=v,
                bg="#050811",
                fg="#fb7185" if is_hat else "#e2e8f0",
                selectcolor="#0f172a",
                activebackground="#050811",
                font=("Segoe UI", 9, "bold"),
                command=self.on_bm_toggle_changed
            )
            cb.pack(side="left", anchor="w")
            tk.Button(
                row_f, text="🗑 Delete", bg="#4c0519", fg="#fda4af",
                font=("Segoe UI", 7, "bold"), relief="flat", padx=6, pady=1,
                command=lambda sn=s_name: self.delete_bm_account_by_name(sn)
            ).pack(side="right", padx=4)

        prev = self.combo_bm_acc.get()
        self.combo_bm_acc["values"] = names
        if names:
            self.combo_bm_acc.set(prev if prev in names else names[0])
            self.on_bm_account_selected()
        else:
            self.combo_bm_acc.set("")
            self.txt_bm_curl.delete("1.0", "end")

    def on_bm_toggle_changed(self):
        for idx, v in self.bm_vars:
            if idx < len(self.bookmarks):
                self.bookmarks[idx]["active"] = bool(v.get())
        self.save_all_vaults_to_disk(False)

    def _suggest_next_bm_account_name(self, raw_curl):
        """
        Generates a clean, non-colliding bookmark account name (e.g. 'hatena1', 'hatena2', 'diigo1')
        when the user pastes a new bookmark cURL.
        """
        dom = extract_domain_from_curl(raw_curl).lower()
        if "hatena" in dom:
            base = "hatena"
        elif "diigo" in dom:
            base = "diigo"
        elif "instapaper" in dom:
            base = "instapaper"
        elif "plurk" in dom:
            base = "plurk"
        elif "folkd" in dom:
            base = "folkd"
        elif dom and dom != "unknown site":
            base = derive_clean_site_name_from_domain(dom).lower()
        else:
            base = "bookmark"

        existing_low = {str(b.get("site", "")).strip().lower() for b in self.bookmarks}
        if base not in existing_low and f"{base}1" not in existing_low:
            return f"{base}1"
        idx = 1
        while f"{base}{idx}".lower() in existing_low:
            idx += 1
        return f"{base}{idx}"

    def prepare_new_bm_account(self):
        """Clears the Bookmark cURL editor and prepares a fresh account name so existing accounts are untouched."""
        self.txt_bm_curl.delete("1.0", "end")
        self.entry_bm_new.delete(0, "end")
        self.entry_bm_new.focus_set()
        self.lbl_bm_status.config(text="➕ Ready for new account: type a name in 'New Name:', paste your cURL, and click Save!")

    def on_bm_account_selected(self):
        sel = self.combo_bm_acc.get().strip()
        for b in self.bookmarks:
            if b.get("site") == sel:
                self.entry_bm_new.delete(0, "end")
                self.combo_bm_type.set(b.get("type", "SINGLE"))
                self.txt_bm_curl.delete("1.0", "end")
                self.txt_bm_curl.insert("1.0", str(b.get("curl", "")))
                break

    def on_bm_curl_modified(self):
        """
        Inspects edited/pasted Bookmark cURL WITHOUT overwriting the currently selected account in `self.bookmarks`!
        If the user pasted a different cURL while `New Name:` is still empty, automatically pre-fills `New Name:`
        with the next available name (e.g. `hatena2`) so `hatena1` is NEVER overwritten!
        """
        raw = self.txt_bm_curl.get("1.0", "end").strip()
        if "curl" not in raw.lower():
            return

        new_nm = self.entry_bm_new.get().strip()
        sel_nm = self.combo_bm_acc.get().strip()

        # Check if this cURL is already identical to an existing account
        for b in self.bookmarks:
            if str(b.get("curl", "")).strip() == raw:
                return

        if not new_nm:
            # If no account exists yet, or if the selected account already has a DIFFERENT saved cURL,
            # automatically suggest a new name (e.g. hatena2) in `New Name:` so hatena1 is never overwritten!
            sel_curl = ""
            for b in self.bookmarks:
                if str(b.get("site", "")).lower() == sel_nm.lower():
                    sel_curl = str(b.get("curl", "")).strip()
                    break
            if not sel_nm or (sel_curl and sel_curl != raw):
                suggested = self._suggest_next_bm_account_name(raw)
                self.entry_bm_new.delete(0, "end")
                self.entry_bm_new.insert(0, suggested)
                self.lbl_bm_status.config(
                    text=f"✨ New cURL detected! Pre-filled New Name '{suggested}' (keeps '{sel_nm}' untouched). Click Save below!"
                )

    def paste_bm_curl(self):
        """
        Pastes clipboard cURL using Win32 + PowerShell fallback (`_get_any_clipboard_text`) so large cURLs
        never fail to paste, and never overwrites an existing account when adding a new account!
        """
        cb = self._get_any_clipboard_text()
        if not cb or "curl" not in cb.lower():
            messagebox.showwarning("Clipboard Empty", "Copy a valid 'Copy as cURL (bash)' command to your clipboard first.")
            return

        new_nm = self.entry_bm_new.get().strip()
        sel_nm = self.combo_bm_acc.get().strip()

        self.txt_bm_curl.delete("1.0", "end")
        self.txt_bm_curl.insert("1.0", cb)

        if not new_nm:
            sel_curl = ""
            for b in self.bookmarks:
                if str(b.get("site", "")).lower() == sel_nm.lower():
                    sel_curl = str(b.get("curl", "")).strip()
                    break
            if not sel_nm or (sel_curl and sel_curl != cb.strip()):
                suggested = self._suggest_next_bm_account_name(cb)
                self.entry_bm_new.delete(0, "end")
                self.entry_bm_new.insert(0, suggested)

        self.save_bm_account_curl()

    def save_bm_account_curl(self):
        new_nm = self.entry_bm_new.get().strip()
        sel_nm = self.combo_bm_acc.get().strip()
        acc_type = self.combo_bm_type.get().strip() or "SINGLE"
        acc_curl = str(self.txt_bm_curl.get("1.0", "end")).strip()
        if "curl" not in acc_curl.lower():
            messagebox.showwarning("Missing cURL", "Paste a valid Bookmark cURL command first.")
            return

        # If `New Name:` was left blank, check if `sel_nm` already has a different Hatena username/cURL;
        # if it's a different Hatena user URL (e.g. b.hatena.ne.jp/user2/ vs b.hatena.ne.jp/user1/), auto-assign next name!
        if not new_nm and sel_nm:
            for b in self.bookmarks:
                if str(b.get("site", "")).lower() == sel_nm.lower():
                    old_c = str(b.get("curl", "")).strip()
                    m_old_u = re.search(r'hatena\.ne\.jp/([^/\s\'"\\]+)/', old_c, re.I)
                    m_new_u = re.search(r'hatena\.ne\.jp/([^/\s\'"\\]+)/', acc_curl, re.I)
                    if m_old_u and m_new_u and m_old_u.group(1).lower() != m_new_u.group(1).lower():
                        new_nm = self._suggest_next_bm_account_name(acc_curl)
                    break

        acc_name = new_nm or sel_nm
        if not acc_name:
            acc_name = self._suggest_next_bm_account_name(acc_curl)

        self.deleted_bookmarks = [x for x in self.deleted_bookmarks if x.lower() != acc_name.lower()]
        found = False
        for b in self.bookmarks:
            if str(b.get("site", "")).lower() == acc_name.lower():
                b["site"] = acc_name
                b["type"] = str(acc_type)
                b["curl"] = str(acc_curl)
                b["active"] = True
                found = True
                break
        if not found:
            self.bookmarks.append({
                "site": str(acc_name),
                "type": str(acc_type),
                "curl": str(acc_curl),
                "active": True
            })
        self.entry_bm_new.delete(0, "end")
        self.combo_bm_acc["values"] = [x["site"] for x in self.bookmarks]
        self.combo_bm_acc.set(acc_name)
        self.save_all_vaults_to_disk(False)
        self.render_bookmark_checkboxes()
        self.lbl_bm_status.config(text=f"✅ Permanently saved '{acc_name}' ({acc_type}) to bookmarks_curl_vault.json!")
        self.log_console(f"💾 Saved Bookmark Account '{acc_name}' in bookmarks_curl_vault.json.")

    def delete_bm_account_by_name(self, site_name):
        sel = (site_name or "").strip()
        if not sel:
            return
        self.bookmarks = [b for b in self.bookmarks if b.get("site", "").lower() != sel.lower()]
        if sel not in self.deleted_bookmarks:
            self.deleted_bookmarks.append(sel)
        self.save_all_vaults_to_disk(False)
        self.render_bookmark_checkboxes()
        self.lbl_bm_status.config(text=f"🗑 Permanently deleted Bookmark Site '{sel}' from bookmarks_curl_vault.json!")
        self.log_console(f"🗑 Permanently deleted Bookmark Site '{sel}'.")

    def delete_bm_account(self):
        sel = self.combo_bm_acc.get().strip() or self.entry_bm_new.get().strip()
        self.delete_bm_account_by_name(sel)

    def apply_bm_preset(self, mode):
        for b in self.bookmarks:
            if mode == "all_on": b["active"] = True
            elif mode == "all_off": b["active"] = False
            elif mode == "no_hatena": b["active"] = "hatena" not in b.get("site", "").lower()
        self.save_all_vaults_to_disk(False)
        self.render_bookmark_checkboxes()

    def save_bm_toggles(self):
        self.on_bm_toggle_changed()
        cnt = sum(1 for b in self.bookmarks if b.get("active"))
        self.lbl_bm_status.config(text=f"✅ Saved {cnt} active bookmark accounts to bookmarks_curl_vault.json.")

    # --------------------------------------------------------------------------
    # SHEET SYNC & 20/20 SHEET SAVER + AUDIT (FIXES ROW 22 / #20)
    # --------------------------------------------------------------------------
    def extract_sheet_id_and_gid(self, url_str):
        m_id = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]{15,})', url_str or '')
        m_gid = re.search(r'[?#&]gid=(\d+)', url_str or '')
        return (m_id.group(1) if m_id else (url_str or '').strip()), (m_gid.group(1) if m_gid else '')

    def fetch_csv_grid(self, csv_url):
        try:
            if HAS_REQUESTS:
                r = _HTTP_SESSION.get(csv_url, headers={"User-Agent": "Mozilla/5.0"}, allow_redirects=True, verify=False, timeout=4.5)
                if r.status_code != 200:
                    return []
                txt = r.text or ""
            else:
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(csv_url, headers={"User-Agent": "Mozilla/5.0"}, method="GET")
                opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
                with opener.open(req, timeout=4.5) as r:
                    if r.status != 200:
                        return []
                    txt = r.read().decode("utf-8", errors="replace")
            t_low = txt[:300].lower()
            if not txt or "<!doctype" in t_low or "<html" in t_low or "servicelogin" in t_low:
                return []
            return list(csv.reader(io.StringIO(txt)))
        except Exception:
            return []

    @staticmethod
    def _col_to_num(col_str):
        n = 0
        for ch in (col_str or "").strip().upper():
            if "A" <= ch <= "Z":
                n = n * 26 + (ord(ch) - 64)
        return n

    @staticmethod
    def _num_to_col(num):
        s = ""
        while num > 0:
            rem = (num - 1) % 26
            s = chr(65 + rem) + s
            num = (num - rem - 1) // 26
        return s or "E"

    def _is_valid_keyword_grid(self, grid):
        if not grid or len(grid) < 3:
            return False
        for r_idx in range(min(6, len(grid))):
            row = grid[r_idx]
            c0 = (row[0] if len(row) > 0 else "").strip().lower()
            c1 = (row[1] if len(row) > 1 else "").strip().lower()
            c2 = (row[2] if len(row) > 2 else "").strip().lower()
            c3 = (row[3] if len(row) > 3 else "").strip().lower()
            if c0 in ("s.no", "sno", "s.no.", "1") or "keyword" in c2 or "keyword" in c3 or "link" in c3 or (c0.isdigit() and (c1 or c2)):
                return True
        return False

    def _score_posting_tab_grid(self, grid):
        """
        Distinguishes the real 'posting site ' tab from Tab 0 ('keyword' tab) in multi-tab sheets like 'Ankit singh':
        - Returns 0 if not a valid keyword grid.
        - Returns 1 if it only has keywords (like Tab 0 'keyword' tab with no 'link' header and no 'http' links).
        - Returns 2 if Row 1/2 has 'link' column headers (Columns D/E+).
        - Returns 3 + link_count if it has 'link' column headers AND 'http' links in Today's rows!
        """
        if not self._is_valid_keyword_grid(grid):
            return 0
        has_link_hdr = False
        for r_idx in range(min(5, len(grid))):
            for c_idx in range(3, min(len(grid[r_idx]), 18)):
                cell_low = (grid[r_idx][c_idx] or "").strip().lower()
                if cell_low in ("link", "links", "url"):
                    has_link_hdr = True
                    break
            if has_link_hdr:
                break
        http_cnt = 0
        for r_idx in range(1, min(26, len(grid))):
            for c_idx in range(3, min(len(grid[r_idx]), 25)):
                if (grid[r_idx][c_idx] or "").strip().startswith("http"):
                    http_cnt += 1
        if http_cnt > 0:
            return 3 + http_cnt
        if has_link_hdr:
            return 2
        return 1

    def get_todays_table_info_from_csv(self, force_refresh=False):
        """
        Reads the Main Sheet CSV and strictly isolates TODAY'S top table (Rows 3..22) on the 'posting site ' / 'Sheet1' tab.
        Never gets fooled by Tab 0 ('keyword' tab or blank 'Untitled' tab) even when the URL is '/edit?usp=sharing' (no gid)!
        """
        self.main_sheet_url = self.entry_main_sheet.get().strip() or self.main_sheet_url
        sid, gid = self.extract_sheet_id_and_gid(self.main_sheet_url)
        if not sid:
            return None

        now_t = time.time()
        if not force_refresh and getattr(self, "_cached_today_info_sid", None) == (sid, gid):
            if now_t - getattr(self, "_cached_today_info_ts", 0) < 3.5 and getattr(self, "_cached_today_info_val", None):
                return self._cached_today_info_val

        resolved_tab, resolved_gid = self._resolve_sheet_tab_and_gid(sid, gid)
        use_gid = gid or resolved_gid

        ts = int(now_t * 1000)
        csv_candidates = []
        if use_gid:
            csv_candidates.append((resolved_tab or "posting site ", f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&gid={use_gid}&_t={ts}"))
            csv_candidates.append((resolved_tab or "posting site ", f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={use_gid}&_t={ts}"))

        # Always prioritize 'posting site ' (with trailing space!) BEFORE 'Sheet1' or 'posting site' (no space)
        for tab_cand in ("posting site ", resolved_tab, "Sheet1", "posting site"):
            if not tab_cand:
                continue
            u_cand = f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote(tab_cand)}&_t={ts}"
            if not any(u_cand == existing_u for _, existing_u in csv_candidates):
                csv_candidates.append((tab_cand, u_cand))

        csv_candidates.append((resolved_tab or "Sheet1", f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&_t={ts}"))

        best_grid = []
        best_score = 0
        winning_tab = resolved_tab or "posting site "

        for tab_label, cu in csv_candidates:
            g = self.fetch_csv_grid(cu)
            sc = self._score_posting_tab_grid(g)
            if sc > best_score:
                best_score = sc
                best_grid = g
                if tab_label:
                    winning_tab = tab_label
            # If we found a tab with 'link' headers (score >= 2), stop immediately!
            if best_score >= 2:
                break

        grid = best_grid
        if not grid:
            return None

        if winning_tab:
            disc_map = getattr(self, "_discovered_tab_for_sid", {})
            disc_map[sid] = winning_tab
            self._discovered_tab_for_sid = disc_map

        # 1. Dynamically locate Today's header row (Row 2 in Sheet, which is grid[0] in gviz or grid[1] in export)
        hdr_row_idx = 1 if len(grid) > 1 else 0
        for r_cand in range(min(5, len(grid))):
            c0 = (grid[r_cand][0] if len(grid[r_cand]) > 0 else "").strip().lower()
            c2 = (grid[r_cand][2] if len(grid[r_cand]) > 2 else "").strip().lower()
            if c0 in ("s.no", "sno", "s.no.") or c2 == "keyword":
                hdr_row_idx = r_cand
                break

        # 2. Detect first 'link' column in hdr_row_idx (skip S.no=0, DATE=1, keyword=2, keyword in spanish=3)
        first_link_col_idx = 4  # Default Column E (0-based index 4)
        for c in range(3, min(len(grid[hdr_row_idx]), 12)):
            h = (grid[hdr_row_idx][c] or "").strip().lower()
            if any(w in h for w in ("keyword", "spanish", "date", "s.no", "topic")):
                continue
            if h in ("link", "links", "url", ""):
                first_link_col_idx = c
                break

        # 3. Strictly isolate TODAY'S top keyword block (starting right after hdr_row_idx)
        today_start_idx = hdr_row_idx + 1
        today_end_idx = today_start_idx
        kw_count = 0
        for r in range(today_start_idx, min(len(grid), today_start_idx + 60)):
            sno = (grid[r][0] if len(grid[r]) > 0 else "").strip()
            kw = (grid[r][2] if len(grid[r]) > 2 else "").strip()
            if kw_count > 0 and (sno.lower() in ("s.no", "sno", "s.no.") or kw.lower() == "keyword"):
                break  # Reached Yesterday's table header (Row 26)!
            if kw_count > 1 and sno == "1":
                break  # S.no reset back to 1 (Yesterday's Row 27)!
            if not sno and not kw:
                break  # Reached the blank separator row at the bottom of Today's table (Row 23)!
            today_end_idx = r
            kw_count += 1

        if kw_count == 0:
            kw_count = 20
            today_end_idx = min(len(grid) - 1, today_start_idx + 19)

        # 4. Scan columns from first_link_col_idx onwards ONLY across Today's rows (today_start_idx .. today_end_idx)
        max_c = max((len(grid[r]) for r in range(hdr_row_idx, min(len(grid), today_end_idx + 1))), default=first_link_col_idx + 8)
        rightmost_active_idx = first_link_col_idx
        has_any_active = False
        col_links_map = {}  # col_letter -> list of (sheet_row_3_to_22, url) in Today's table

        for c in range(first_link_col_idx, max_c + 6):
            col_let = self._num_to_col(c + 1)
            links_in_c = []
            for r in range(today_start_idx, min(len(grid), today_end_idx + 1)):
                val = (grid[r][c] if c < len(grid[r]) else "").strip()
                if val.startswith("http"):
                    sheet_row_num = 3 + (r - today_start_idx)  # Always 3..22 for Today's 20 keywords!
                    links_in_c.append((sheet_row_num, val))
            col_links_map[col_let] = links_in_c
            if len(links_in_c) > 0:
                rightmost_active_idx = c
                has_any_active = True

        latest_active_col = self._num_to_col(rightmost_active_idx + 1)
        next_open_idx = (rightmost_active_idx + 1) if has_any_active else first_link_col_idx
        while col_links_map.get(self._num_to_col(next_open_idx + 1)):
            next_open_idx += 1
        next_open_col = self._num_to_col(next_open_idx + 1)

        res_info = {
            "grid": grid,
            "hdr_row_idx": hdr_row_idx,
            "today_start_idx": today_start_idx,
            "today_end_idx": today_end_idx,
            "first_link_col_idx": first_link_col_idx,
            "kw_count": kw_count,
            "latest_active_col": latest_active_col,
            "next_open_col": next_open_col,
            "col_links_map": col_links_map
        }
        self._cached_today_info_sid = (sid, gid)
        self._cached_today_info_ts = time.time()
        self._cached_today_info_val = res_info
        return res_info

    def _fetch_raw_url_bytes(self, url, timeout=6.0):
        try:
            if HAS_REQUESTS:
                r = _HTTP_SESSION.get(url, headers={"User-Agent": "Mozilla/5.0"}, allow_redirects=True, verify=False, timeout=timeout)
                if r.status_code == 200:
                    return r.content or b""
                return b""
            else:
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}, method="GET")
                opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
                with opener.open(req, timeout=timeout) as r:
                    if r.status == 200:
                        return r.read() or b""
                return b""
        except Exception:
            return b""

    def fetch_content_sheet_rich_links(self, sid, gid):
        """
        Extracts Google Sheets Rich-Text (`Ctrl+K`) hyperlinks and cell hyperlinks from a Content Sheet (`sid`, `gid`)
        using `/htmlview`, `/pubhtml`, `/gviz/tq?tqx=out:html`, and `/export?format=xlsx` (in-memory ZIP/XML),
        returning:
        - `html_cells`: list of `(alnum_fingerprint, [(clean_url, anchor_text), ...])` for every cell containing `<a>` links
        - `xlsx_row_links`: dict `row_1based -> [clean_url, ...]` from `.xlsx` `<hyperlinks>` relationships
        """
        html_cells = []
        xlsx_row_links = {}
        ts = int(time.time() * 1000)

        # 1. Fetch `/htmlview`, `/pubhtml`, and `/gviz/tq?tqx=out:html` to capture partial in-cell Rich-Text (`Ctrl+K`) links!
        html_urls = []
        if gid:
            html_urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/htmlview?gid={gid}&_t={ts}")
            html_urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:html&gid={gid}&_t={ts}")
            html_urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/pubhtml?gid={gid}&single=true&_t={ts}")
        else:
            html_urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/htmlview?_t={ts}")
            html_urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:html&sheet=Content&_t={ts}")
            html_urls.append(f"https://docs.google.com/spreadsheets/d/{sid}/pubhtml&_t={ts}")

        for hu in html_urls:
            raw_b = self._fetch_raw_url_bytes(hu, timeout=5.5)
            if not raw_b:
                continue
            htxt = raw_b.decode("utf-8", errors="replace")
            if "<a " not in htxt.lower() and "<a\t" not in htxt.lower():
                continue
            for m_td in re.finditer(r'<td\b[^>]*>([\s\S]*?)</td>', htxt, flags=re.I):
                td_inner = m_td.group(1)
                if "<a" not in td_inner.lower():
                    continue
                links_in_td = []
                for m_a in re.finditer(r'<a\b([^>]*)>([\s\S]*?)</a>', td_inner, flags=re.I):
                    attrs_a = html.unescape(m_a.group(1) or "")
                    inner_a = m_a.group(2) or ""
                    hm = re.search(r'\bhref\s*=\s*(?:"([^"]+)"|\'([^\']+)\'|([^\s>"\']+))', attrs_a, flags=re.I)
                    if not hm:
                        continue
                    raw_href = hm.group(1) or hm.group(2) or hm.group(3) or ""
                    clean_u = unwrap_google_redirect_url(raw_href)
                    clean_anc = html.unescape(re.sub(r'<[^>]+>', '', inner_a)).strip()
                    if clean_u.startswith("http") and clean_anc and not any(
                        sk in clean_u.lower() for sk in ("accounts.google.com", "support.google.com", "docs.google.com/spreadsheets")
                    ):
                        links_in_td.append((clean_u, clean_anc))
                if links_in_td:
                    td_plain = html.unescape(re.sub(r'<[^>]+>', '', td_inner))
                    td_alnum = re.sub(r'[^a-z0-9]+', '', td_plain.lower())
                    if td_alnum:
                        html_cells.append((td_alnum, links_in_td))
            if html_cells:
                break

        # 2. Also check `.xlsx` export in memory for whole-cell `<hyperlink>` relationships
        xlsx_u = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx&gid={gid}" if gid else f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
        xb = self._fetch_raw_url_bytes(xlsx_u, timeout=6.0)
        if xb and xb[:2] == b"PK":
            try:
                with zipfile.ZipFile(io.BytesIO(xb)) as zf:
                    names = zf.namelist()
                    rels_map = {}
                    for nm in names:
                        if nm.startswith("xl/worksheets/_rels/") and nm.endswith(".rels"):
                            r_xml = zf.read(nm).decode("utf-8", errors="replace")
                            for m_rel in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="([^"]+)"', r_xml):
                                rid = m_rel.group(1) or m_rel.group(4)
                                targ = html.unescape(m_rel.group(2) or m_rel.group(3) or "")
                                if rid and targ.startswith("http"):
                                    rels_map[(nm, rid)] = unwrap_google_redirect_url(targ)
                    for nm in names:
                        if nm.startswith("xl/worksheets/sheet") and nm.endswith(".xml"):
                            s_xml = zf.read(nm).decode("utf-8", errors="replace")
                            rel_nm = nm.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
                            for m_hl in re.finditer(r'<hyperlink\b([^>]+)/?>', s_xml, flags=re.I):
                                at = m_hl.group(1)
                                ref_m = re.search(r'\bref="([A-Z]+)(\d+)"', at, flags=re.I)
                                rid_m = re.search(r'\br:id="([^"]+)"', at, flags=re.I)
                                if ref_m and rid_m:
                                    r_num = int(ref_m.group(2))
                                    u_val = rels_map.get((rel_nm, rid_m.group(1)))
                                    if u_val:
                                        xlsx_row_links.setdefault(r_num, []).append(u_val)
            except Exception:
                pass

        return html_cells, xlsx_row_links

    @staticmethod
    def _inject_rich_links_into_body(body_text, title_text, links_list):
        """
        Injects `(clean_url, anchor_text)` pairs into `body_text` as `<a href="clean_url">anchor_text</a>`
        without disturbing paragraph breaks (`\n\n`).
        """
        s = normalize_content_links_to_html(body_text)
        for clean_u, clean_anc in links_list:
            if not clean_u or not clean_u.startswith("http"):
                continue
            if clean_u in s:
                continue  # Already linked!
            anc = (clean_anc or "").strip()
            if anc:
                # Match anchor text allowing flexible whitespace
                words = [re.escape(w) for w in anc.split() if w]
                if words:
                    pat = r'(?<!["\'>/])\b' + r'\s+'.join(words) if re.match(r'^\w', anc) else r'\s+'.join(words)
                    new_s, cnt = re.subn(pat, lambda m: f'<a href="{clean_u}">{m.group(0)}</a>', s, count=1, flags=re.I)
                    if cnt > 0:
                        s = new_s
                        continue
            # Fallback: link the base keyword from title_text inside body_text, or append anchor link
            base_kw = re.split(r'\s*[\[\(]', (title_text or "").strip())[0].strip()
            if base_kw and len(base_kw) >= 4:
                kw_words = [re.escape(w) for w in base_kw.split() if w]
                pat_kw = r'\s+'.join(kw_words)
                new_s, cnt = re.subn(pat_kw, lambda m: f'<a href="{clean_u}">{m.group(0)}</a>', s, count=1, flags=re.I)
                if cnt > 0:
                    s = new_s
                    continue
            if anc:
                s = f'{s}\n\n<a href="{clean_u}">{anc}</a>'.strip()
        return s

    def _sync_content_sheet_now(self, cu):
        """Synchronously loads Content Sheet rows and enriches them with all Rich-Text / HTML / BBCode / Markdown hyperlinks."""
        sid, gid = self.extract_sheet_id_and_gid(cu)
        if not sid:
            return []
        ts = int(time.time() * 1000)
        csv_urls = [
            f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&gid={gid}&_t={ts}"
            if gid else
            f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet=Content&_t={ts}",
            f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}&_t={ts}"
            if gid else
            f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&_t={ts}"
        ]
        grid = []
        for u_c in csv_urls:
            grid = self.fetch_csv_grid(u_c)
            if grid:
                break
        if not grid:
            return []

        html_cells, xlsx_row_links = self.fetch_content_sheet_rich_links(sid, gid)

        first_cell = (grid[0][0] if grid[0] else "").strip().lower()
        second_cell = (grid[0][1] if len(grid[0]) > 1 else "").strip().lower()
        has_hdr = (
            first_cell in ("title", "keyword", "keywords", "topic", "subject", "heading", "s.no", "sno", "#")
            or second_cell in ("title", "keyword", "keywords", "content", "body", "article", "description")
        )
        start_idx = 1 if has_hdr else 0
        new_rows = []
        for i in range(start_idx, len(grid)):
            row_vals = grid[i]
            if not row_vals:
                continue
            # Support both [Title, Body, (Optional URL)] and [S.No, Title, Body, (Optional URL)]
            c0 = (row_vals[0] if len(row_vals) > 0 else "").strip()
            c1 = (row_vals[1] if len(row_vals) > 1 else "").strip()
            c2 = (row_vals[2] if len(row_vals) > 2 else "").strip()
            c3 = (row_vals[3] if len(row_vals) > 3 else "").strip()
            if c0.isdigit() and len(c0) <= 3 and c1 and c2 and not c2.startswith("http"):
                t, b, extra_u = c1, c2, c3
            else:
                t, b, extra_u = c0, c1, c2

            if not t and not b:
                continue

            b = normalize_content_links_to_html(b)

            # If `b` does not yet contain `<a href=`, check Rich-Text links from `/htmlview` / `.xlsx` / `extra_u`!
            if "<a href=" not in b.lower():
                b_alnum = re.sub(r'[^a-z0-9]+', '', html.unescape(re.sub(r'<[^>]+>', '', b)).lower())
                t_alnum = re.sub(r'[^a-z0-9]+', '', t.lower())
                matched_links = []
                if b_alnum and html_cells:
                    probe = b_alnum[:220]
                    for td_alnum, td_links in html_cells:
                        if (probe and probe in td_alnum) or (td_alnum[:220] and td_alnum[:220] in b_alnum):
                            matched_links.extend(td_links)
                            break
                if not matched_links and t_alnum and html_cells:
                    for td_alnum, td_links in html_cells:
                        if t_alnum == td_alnum or (len(t_alnum) >= 12 and t_alnum in td_alnum):
                            matched_links.extend(td_links)
                            break
                if not matched_links and (i + 1) in xlsx_row_links:
                    for xu in xlsx_row_links[i + 1]:
                        matched_links.append((xu, ""))
                if not matched_links and extra_u.startswith("http"):
                    matched_links.append((unwrap_google_redirect_url(extra_u), ""))

                if matched_links:
                    b = self._inject_rich_links_into_body(b, t, matched_links)

            rid = len(new_rows) + 1
            new_rows.append({
                "id": rid,
                "status": "Pending",
                "title": t,
                "body": b,
                "bookmark_status": "-",
                "url": "-",
                "selected": True
            })
        return new_rows

    def sync_all_sheets(self, show_popup=False):
        self.save_all_vaults_to_disk(True)
        cu = self.content_sheet_url.strip()
        mu = self.main_sheet_url.strip()

        def _bg():
            link_cnt = 0
            if cu:
                try:
                    new_rows = self._sync_content_sheet_now(cu)
                    if new_rows:
                        if self.rows and len(self.rows) == len(new_rows):
                            # Update existing row dicts in-place so active/completed posts in flight NEVER lose their status or live URL!
                            for old_r, fresh_r in zip(self.rows, new_rows):
                                old_r["title"] = fresh_r["title"]
                                old_r["body"] = fresh_r["body"]
                                if str(old_r.get("url", "-")).startswith("http") or "Done" in str(old_r.get("status", "")) or "Posting" in str(old_r.get("status", "")):
                                    pass
                                else:
                                    old_r["status"] = fresh_r["status"]
                                    old_r["url"] = fresh_r["url"]
                                    old_r["bookmark_status"] = fresh_r["bookmark_status"]
                        else:
                            old_by_id = {r.get("id"): r for r in (self.rows or [])}
                            for fresh_r in new_rows:
                                prev = old_by_id.get(fresh_r["id"])
                                if prev and (str(prev.get("url", "-")).startswith("http") or "Done" in str(prev.get("status", "")) or "Posting" in str(prev.get("status", ""))):
                                    fresh_r["status"] = prev["status"]
                                    fresh_r["url"] = prev["url"]
                                    fresh_r["bookmark_status"] = prev.get("bookmark_status", "-")
                                    fresh_r["selected"] = prev.get("selected", True)
                            self.rows = new_rows
                        link_cnt = sum(1 for r in self.rows if "<a href=" in str(r.get("body", "")).lower())
                        self.save_all_vaults_to_disk(False)
                except Exception as e:
                    self.log_console(f"⚠️ Content Sheet sync error: {e}")

            if mu:
                try:
                    info = self.get_todays_table_info_from_csv(force_refresh=True)
                    if info:
                        self.detected_kw_count = info["kw_count"]
                        self.latest_active_col = info["latest_active_col"]
                        self.next_open_col = info["next_open_col"]
                        self._main_sheet_synced_once = True
                        if not cu and not self.rows:
                            grid = info["grid"]
                            main_kw_rows = []
                            for r in range(info["today_start_idx"], info["today_end_idx"] + 1):
                                kw_t = (grid[r][2] if len(grid[r]) > 2 else "").strip()
                                if kw_t:
                                    main_kw_rows.append({
                                        "id": len(main_kw_rows) + 1,
                                        "status": "Pending",
                                        "title": kw_t,
                                        "body": f"<p>{kw_t}</p>",
                                        "bookmark_status": "-",
                                        "url": "-",
                                        "selected": True
                                    })
                            if main_kw_rows:
                                self.rows = main_kw_rows
                    else:
                        col_let, filled_cnt, kw_cnt, _ = self.query_rightmost_sheet_column_state("AUTO")
                        self.detected_kw_count = kw_cnt
                        self.latest_active_col = col_let
                        self.next_open_col = self._num_to_col(self._col_to_num(col_let) + 1) if filled_cnt > 0 else col_let
                        self._main_sheet_synced_once = True
                except Exception as e:
                    self.log_console(f"⚠️ Main Sheet check note: {e}")

            self.save_all_vaults_to_disk(False)

            def _ui():
                self.lbl_batch_badge.config(
                    text=f" 📊 Active Batch: {len(self.rows)} Keywords (Rows 3–{2 + len(self.rows)}) | Today's Col: {self.latest_active_col} | Next Empty: {self.next_open_col} "
                )
                self.render_queue_table()
                hl_note = f" ({link_cnt}/{len(self.rows)} rows with hyperlinks)" if link_cnt else ""
                self.log_console(f"✅ Synced {len(self.rows)} Content Rows{hl_note} | Today's Active Col: {self.latest_active_col} | Next Empty Col: {self.next_open_col}")
                if show_popup:
                    messagebox.showinfo(
                        "Sheets Synced",
                        f"Synced {len(self.rows)} Content Rows{hl_note}!\n\n"
                        f"• Today's Active Column: {self.latest_active_col}\n"
                        f"• Today's Next Empty Column: {self.next_open_col}"
                    )
            self.root.after(0, _ui)

        threading.Thread(target=_bg, daemon=True).start()

    def query_rightmost_sheet_column_state(self, req_col="AUTO"):
        """
        Determines the target column and links in TODAY'S table (Rows 3..22).
        Always uses Today's CSV table first; if falling back to Apps Script doGet,
        strictly passes &end_row=22 and caps returned links to Today's rows!
        """
        clean_req = (req_col or "AUTO").strip().upper()
        info = self.get_todays_table_info_from_csv()
        if info:
            kw_cnt = info["kw_count"]
            col_map = info["col_links_map"]
            if clean_req != "AUTO" and clean_req:
                pairs = col_map.get(clean_req, [])
                return clean_req, len(pairs), kw_cnt, [u for _, u in pairs]
            else:
                lat_col = info["latest_active_col"]
                pairs = col_map.get(lat_col, [])
                return lat_col, len(pairs), kw_cnt, [u for _, u in pairs]

        # Fallback to Apps Script doGet with explicit end_row=22 so Yesterday's table (Row 27+) is NEVER scanned
        self.main_sheet_url = self.entry_main_sheet.get().strip() or self.main_sheet_url
        self.apps_script_url = self.entry_apps_script.get().strip() or self.apps_script_url
        sid, gid = self.extract_sheet_id_and_gid(self.main_sheet_url)
        resolved_tab, resolved_gid = self._resolve_sheet_tab_and_gid(sid, gid)
        for ep in self._get_candidate_sheet_endpoints():
            if not sid:
                break
            try:
                q_url = (
                    f"{ep}?sheet_id={urllib.parse.quote(sid)}"
                    f"&tab_name={urllib.parse.quote(resolved_tab)}"
                    f"&gid={urllib.parse.quote(resolved_gid)}"
                    f"&col={urllib.parse.quote(clean_req)}&end_row=22&_t={int(time.time() * 1000)}"
                )
                res = execute_apps_script_request(q_url, "GET")
                j, _ = parse_apps_script_json_or_diagnose(res.get("body", ""))
                if j and j.get("status") == "success":
                    col_let = str(j.get("column") or "E").strip().upper()
                    if col_let in ("A", "B", "C"):
                        col_let = "E"
                    kw_cnt = min(int(j.get("keyword_count") or 20), 20)
                    lnks = (j.get("links") or [])[:kw_cnt]
                    return col_let, len(lnks), kw_cnt, lnks
            except Exception:
                pass
        return (self.latest_active_col or "E"), 0, 20, []

    def start_new_zoom_sheet_column(self):
        """Advances the Tab 3 Target Sheet Column to the next completely empty column to the right in Today's table."""
        def _adv():
            info = self.get_todays_table_info_from_csv()
            if info:
                nxt_col = info["next_open_col"]
            else:
                col_let, filled_cnt, _, _ = self.query_rightmost_sheet_column_state("AUTO")
                nxt_col = self._num_to_col(self._col_to_num(col_let) + 1) if filled_cnt > 0 else col_let
            self.next_open_col = nxt_col
            def _u():
                self.entry_zoom_save_col.delete(0, "end")
                self.entry_zoom_save_col.insert(0, nxt_col)
                self.lbl_zoom_msg.config(text=f"🔒 Locked Target Column to new clean Column {nxt_col} (starts at Row 3)!")
            self.root.after(0, _u)
        threading.Thread(target=_adv, daemon=True).start()

    def _resolve_sheet_tab_and_gid(self, sid, gid):
        """
        Resolves `(tab_name, tab_gid)` for ANY Google Sheet URL — even when copied from the
        Share dialog (`/edit?usp=sharing` without `?gid=...`) or when Tab 0 is a blank `Untitled` tab
        and Tab 1 is `Sheet1` (as in image_cdc264.png)!
        """
        if not sid:
            return ("posting site ", "")

        cache_map = getattr(self, "_gid_tab_name_cache", {})
        if (sid, gid) in cache_map:
            return cache_map[(sid, gid)]

        disc_map = getattr(self, "_discovered_tab_for_sid", {})

        tab_name = ""
        found_gid = str(gid or "").strip()

        # 1. Inspect `/htmlview` to see all tabs `(t_gid, t_name)` and which tab holds the keyword table
        try:
            hb = self._fetch_raw_url_bytes(f"https://docs.google.com/spreadsheets/d/{sid}/htmlview", timeout=4.5)
            if hb:
                htxt = hb.decode("utf-8", errors="replace")
                tabs = []
                for m_btn in re.finditer(
                    r'id=["\']sheet-button-([0-9]+)["\'][^>]*>[\s\S]*?<a[^>]*>([\s\S]*?)</a>',
                    htxt,
                    flags=re.I
                ):
                    tg = m_btn.group(1).strip()
                    tn = html.unescape(re.sub(r'<[^>]+>', '', m_btn.group(2))).strip()
                    if tn:
                        tabs.append((tg, tn))
                if not tabs:
                    for m_js in re.finditer(
                        r'name:\s*["\']([^"\']+)["\'][^}]*?gid:\s*["\']?([0-9]+)["\']?',
                        htxt
                    ):
                        tn = html.unescape(m_js.group(1)).strip()
                        tg = m_js.group(2).strip()
                        if tn:
                            tabs.append((tg, tn))

                if tabs:
                    # If caller gave an explicit gid in URL, match it first
                    if found_gid:
                        for tg, tn in tabs:
                            if tg == found_gid:
                                tab_name = tn
                                break

                    # If no gid was in the URL (`/edit?usp=sharing`), intelligently pick the active posting tab:
                    if not tab_name:
                        # (a) Prefer 'posting site ' or 'posting site' if present in this workbook
                        for tg, tn in tabs:
                            if tn.strip().lower() == "posting site":
                                tab_name = "posting site "
                                found_gid = tg
                                break

                    if not tab_name:
                        # (b) Check which `<div id="{tg}">` section in htmlview contains 'S.no' and 'keyword'
                        for tg, tn in tabs:
                            m_sec = re.search(
                                rf'<div[^>]+id=["\']{re.escape(tg)}["\'][^>]*>([\s\S]*?)(?:<div[^>]+id=["\'][0-9]+["\']|</body>)',
                                htxt,
                                flags=re.I
                            )
                            if m_sec:
                                sec_low = m_sec.group(1).lower()
                                if ("s.no" in sec_low or "keyword" in sec_low) and tn.strip().lower() not in ("untitled", "keyword", "keywords", "admin", "compititor", "bookmarking site"):
                                    tab_name = "posting site " if tn.strip().lower() == "posting site" else tn
                                    found_gid = tg
                                    break

                    if not tab_name:
                        # (c) Prefer 'Sheet1' or any non-'Untitled' tab over a blank 'Untitled' tab!
                        for tg, tn in tabs:
                            if tn.strip().lower() == "sheet1":
                                tab_name = tn
                                found_gid = tg
                                break
                    if not tab_name:
                        for tg, tn in tabs:
                            if "untitled" not in tn.strip().lower():
                                tab_name = tn
                                found_gid = tg
                                break
        except Exception:
            pass

        # 2. If `/htmlview` didn't resolve `tab_name`, check discovered tab from CSV or probe `gviz`
        if not tab_name and sid in disc_map:
            tab_name = disc_map[sid]

        if not tab_name:
            best_t = ""
            best_sc = 0
            for cand_t in ("posting site ", "Sheet1", "posting site"):
                try:
                    u_test = f"https://docs.google.com/spreadsheets/d/{sid}/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote(cand_t)}"
                    g_test = self.fetch_csv_grid(u_test)
                    sc = self._score_posting_tab_grid(g_test)
                    if sc > best_sc:
                        best_sc = sc
                        best_t = cand_t
                    if best_sc >= 2:
                        break
                except Exception:
                    pass
            if best_t:
                tab_name = best_t
                disc_map[sid] = best_t
                self._discovered_tab_for_sid = disc_map

        if not tab_name:
            tab_name = "Sheet1"

        res_pair = (tab_name, found_gid)
        cache_map[(sid, gid)] = res_pair
        self._gid_tab_name_cache = cache_map
        return res_pair

    def _resolve_tab_name_for_gid(self, sid, gid):
        tab_name, _ = self._resolve_sheet_tab_and_gid(sid, gid)
        return tab_name

    def _get_candidate_sheet_endpoints(self):
        """
        Strictly returns ONLY the user's own `🔗 Apps Script URL` entered in the UI.
        Zero built-in or hidden fallback scripts are ever used, ensuring complete isolation
        between different users' sheets.
        """
        ep_user = (self.entry_apps_script.get().strip() if hasattr(self, "entry_apps_script") else "") or (self.apps_script_url or "").strip()
        if ep_user and "script.google.com" in ep_user:
            return [ep_user]
        return []

    def save_link_to_sheet_sync(self, item_id, title, live_url, col_let, is_auto_mode=False):
        """
        Guaranteed Google Sheet link writer compatible with ALL deployed Code.gs versions:
        - Automatically resolves `(resolved_tab, resolved_gid)` even when the sheet URL is `/edit?usp=sharing`
          (no `gid` in URL) and even when Tab 0 is a blank `Untitled` tab while Tab 1 is `Sheet1`!
        - Sends `action: "save_post_link"` (plus fallback to `action: "write_sheet_url"`).
        - Strictly uses ONLY the user's own `🔗 Apps Script URL` (zero hidden fallback relays).
        - Verifies that the response matches the requested sheet and row!
        """
        if self.stop_requested:
            return False

        self.main_sheet_url = self.entry_main_sheet.get().strip() or self.main_sheet_url
        self.apps_script_url = self.entry_apps_script.get().strip() or self.apps_script_url
        sid, gid = self.extract_sheet_id_and_gid(self.main_sheet_url)
        exact_row = 2 + int(item_id)  # #1 -> Row 3, ..., #20 -> Row 22
        max_kw = max(20, len(self.rows) or 20)

        use_col = (col_let or "").strip().upper()
        if not use_col or use_col in ("AUTO", "A", "B", "C"):
            is_auto_mode = True
            info_now = self.get_todays_table_info_from_csv(force_refresh=True)
            if info_now:
                self.latest_active_col = info_now["latest_active_col"]
                self.next_open_col = info_now["next_open_col"]
                use_col = info_now["next_open_col"]
            else:
                # Let the Apps Script (`BACKUP_SHEET_WEBHOOK` with `write_sheet_url`) scan the live sheet
                # server-side for the first completely empty column when `use_col == "AUTO"`!
                use_col = "AUTO" 
        elif is_auto_mode:
            # Zero-Overwrite Guard in AUTO mode: verify `(use_col, exact_row)` does not already contain a link!
            cached_info = getattr(self, "_cached_today_info_val", None)
            if cached_info and isinstance(cached_info.get("col_links_map"), dict):
                existing_in_col = dict(cached_info["col_links_map"].get(use_col, []))
                if existing_in_col.get(exact_row, "").startswith("http"):
                    fresh_info = self.get_todays_table_info_from_csv(force_refresh=True)
                    if fresh_info:
                        fresh_in_col = dict(fresh_info["col_links_map"].get(use_col, []))
                        if fresh_in_col.get(exact_row, "").startswith("http"):
                            use_col = fresh_info["next_open_col"]

        resolved_tab, resolved_gid = self._resolve_sheet_tab_and_gid(sid, gid)
        use_gid = gid or resolved_gid

        def _make_payload(act_name):
            return json.dumps({
                "action": act_name,
                "sheet_id": sid,
                "spreadsheet_id": sid,
                "id_spreadsheet": sid,
                "sheet_url": self.main_sheet_url,
                "gid": use_gid,
                "tab_name": resolved_tab,
                "title": title,
                "keyword": title,
                "link": live_url,
                "live_url": live_url,
                "url": live_url,
                "item_index": int(item_id),
                "id": int(item_id),
                "s_no": int(item_id),
                "row_offset": int(item_id),
                "row": exact_row,
                "row_index": exact_row,
                "target_row": exact_row,
                "column": use_col,
                "col": use_col,
                "col_letter": use_col,
                "start_row": 3,
                "end_row": 4 + max_kw,
                "total_keywords": max_kw,
                "no_overwrite": bool(is_auto_mode)
            })

        if not sid:
            self.log_console("⚠️ Sheet Write Skipped: Paste your Google Sheet link into '📊 Main Sheet Link' first.")
            return False

        endpoints = self._get_candidate_sheet_endpoints()
        if not endpoints:
            self.log_console("⚠️ Sheet Write Skipped: Paste your own Web App URL from https://script.new into '🔗 Apps Script URL' first (click '📋 Copy Sheet Script' to get the code).")
            return False

        last_err = ""
        for attempt in range(1, 3):
            if self.stop_requested:
                return False
            for ep in endpoints:
                if self.stop_requested:
                    return False
                pref_act = getattr(self, "_working_sheet_action", None)
                if use_col == "AUTO":
                    act_order = ("write_sheet_url", "save_post_link")
                else:
                    act_order = (pref_act, "save_post_link", "write_sheet_url") if pref_act else ("save_post_link", "write_sheet_url")
                seen_acts = set()
                for act_str in [a for a in act_order if a and not (a in seen_acts or seen_acts.add(a))]:
                    try:
                        res = execute_apps_script_request(ep, "POST", _make_payload(act_str), timeout=20)
                        body_txt = (res.get("body") or "").strip()
                        if res.get("ok") and body_txt:
                            j, diag_err = parse_apps_script_json_or_diagnose(body_txt)
                            if j:
                                has_real_write = bool(
                                    j.get("status") == "success"
                                    or j.get("row") is not None
                                    or j.get("col") is not None
                                    or j.get("column") is not None
                                )
                                if has_real_write and j.get("ok") is not False and j.get("status") != "error":
                                    ret_ss_id = str(j.get("sheet_id") or j.get("spreadsheet_id") or "").strip()
                                    if ret_ss_id and sid and ret_ss_id != sid:
                                        last_err = f"Endpoint is bound to another sheet ({ret_ss_id[:8]}...), trying next relay..."
                                        continue
                                    written_col = str(j.get("column") or j.get("col") or use_col).strip().upper()
                                    written_row = int(j.get("row") or exact_row)
                                    # Reject if an old bound script ignored our requested row/column and jumped to a different row
                                    if written_row != exact_row and abs(written_row - exact_row) > 1:
                                        last_err = f"Outdated Apps Script wrote to Row {written_row} instead of Row {exact_row} — click Copy Sheet Script (Code.gs) and deploy a New version in script.new!"
                                        continue
                                    self._working_sheet_ep = ep
                                    self._working_sheet_action = act_str
                                    self._cached_today_info_ts = 0
                                    self.log_console(
                                        f"📄 Saved #{item_id} to Main Sheet [{resolved_tab}] (Column {written_col}, Row {written_row})"
                                    )
                                    return written_col
                                elif j.get("ok") is True and not has_real_write:
                                    last_err = f"Endpoint ignored action '{act_str}'"
                                    continue
                                elif j.get("message") or j.get("error"):
                                    last_err = str(j.get("message") or j.get("error"))
                                    break
                            elif diag_err:
                                last_err = diag_err.splitlines()[0]
                                break
                    except Exception as e:
                        last_err = str(e)
                        break
            time.sleep(0.3)

        self.log_console(f"⚠️ Sheet Write Failed for #{item_id} (Tab '{resolved_tab}', Col {use_col}, Row {exact_row}): {last_err or 'Make sure sheet General Access is set to Anyone with the link -> Editor'}")
        return False

    def test_sheet_connection(self):
        self.save_all_vaults_to_disk(True)
        sid, gid = self.extract_sheet_id_and_gid(self.main_sheet_url)
        if not sid:
            messagebox.showwarning("Main Sheet Link Needed", "Paste your Google Sheet link into '📊 Main Sheet Link' first.")
            return

        self.log_console(f"🔌 Testing connection to Main Sheet ({sid[:12]}...)...")

        def _test():
            try:
                resolved_tab, resolved_gid = self._resolve_sheet_tab_and_gid(sid, gid)
                info = self.get_todays_table_info_from_csv(force_refresh=True)
                if info:
                    self.latest_active_col = info["latest_active_col"]
                    self.next_open_col = info["next_open_col"]
                    kw_cnt = info["kw_count"]
                    lat_links = info["col_links_map"].get(self.latest_active_col, [])
                    msg = (
                        f"✅ Sheet Connected Successfully!\n\n"
                        f"• Active Worksheet Tab: {resolved_tab}" + (f" (gid={resolved_gid})\n" if resolved_gid else "\n") +
                        f"• Today's Keywords (Rows 3–{2 + kw_cnt}): {kw_cnt}\n"
                        f"• Today's Active Column: Column {self.latest_active_col} ({len(lat_links)}/{kw_cnt} links)\n"
                        f"• Today's Next Empty Column: Column {self.next_open_col}\n\n"
                        f"Tip: Ensure Share -> 'Anyone with the link: Editor' is enabled on this sheet so links write instantaneously!"
                    )
                    self.log_console(
                        f"✅ Sheet Connected [{resolved_tab}]! Today's Active Col: {self.latest_active_col} ({len(lat_links)}/{kw_cnt} links) | Next Empty Col: {self.next_open_col}"
                    )
                    self.root.after(0, lambda: messagebox.showinfo("Sheet Connected!", msg))
                    return

                # Fallback to Apps Script doGet across candidate endpoints
                last_diag = ""
                for ep in self._get_candidate_sheet_endpoints():
                    q_url = (
                        f"{ep}?sheet_id={urllib.parse.quote(sid)}"
                        f"&tab_name={urllib.parse.quote(resolved_tab)}"
                        f"&gid={urllib.parse.quote(resolved_gid)}"
                        f"&col=AUTO&end_row=22&_t={int(time.time() * 1000)}"
                    )
                    res = execute_apps_script_request(q_url, "GET")
                    j, diag_err = parse_apps_script_json_or_diagnose(res.get("body", ""))
                    if j and j.get("status") == "success":
                        col = j.get("column", "D")
                        kw_cnt = min(int(j.get("keyword_count") or 20), 20)
                        lat_links = (j.get("links") or [])[:kw_cnt]
                        self.latest_active_col = col
                        self.next_open_col = self._num_to_col(self._col_to_num(col) + 1) if len(lat_links) > 0 else col
                        msg = (
                            f"✅ Sheet & Script Connected Successfully!\n\n"
                            f"• Active Worksheet Tab: {resolved_tab}\n"
                            f"• Today's Keywords (Rows 3–{2 + kw_cnt}): {kw_cnt}\n"
                            f"• Today's Active Column: Column {self.latest_active_col} ({len(lat_links)}/{kw_cnt} links)\n"
                            f"• Today's Next Empty Column: Column {self.next_open_col}"
                        )
                        self.log_console(f"✅ Sheet Connected [{resolved_tab}]! Active Col: {self.latest_active_col} | Next Empty Col: {self.next_open_col}")
                        self.root.after(0, lambda: messagebox.showinfo("Sheet Connected!", msg))
                        return
                    elif diag_err:
                        last_diag = diag_err
                if last_diag:
                    self.log_console(f"🚨 Script Diagnosis: {last_diag}")
                    self.root.after(0, lambda: messagebox.showwarning("How to Fix Sheet Connection", last_diag))
            except Exception as e:
                self.log_console(f"🚨 Sheet Connection Failed: {e}")
                self.root.after(0, lambda: messagebox.showerror("Connection Error", str(e)))

        threading.Thread(target=_test, daemon=True).start()

    def verify_and_repair_sheet_column(self, col_let):
        posted_rows = [r for r in self.rows if str(r.get("url", "")).startswith("http")]
        if not posted_rows:
            return
        self.log_console(f"🔍 [Sheet Audit] Verifying all {len(posted_rows)} links in Today's Column {col_let} (Rows 3–{2 + len(self.rows)})...")
        time.sleep(1.5)
        try:
            info = self.get_todays_table_info_from_csv()
            if not info:
                return
            pairs = dict(info["col_links_map"].get(col_let, []))
            missing = []
            for idx, item in enumerate(self.rows):
                u = str(item.get("url", "")).strip()
                if not u.startswith("http"):
                    continue
                sheet_row = 3 + idx
                if not pairs.get(sheet_row, "").startswith("http"):
                    missing.append((item["id"], sheet_row, item["title"], u))

            if missing:
                rows_str = ", ".join(str(m[1]) for m in missing)
                self.log_console(f"🔧 [Sheet Auto-Repair] Filling missing Row(s) {rows_str} in Column {col_let}...")
                for item_id, sheet_row, title, u in missing:
                    self.save_link_to_sheet_sync(item_id, title, u, col_let)
                    time.sleep(0.5)
            else:
                self.log_console(f"✅ [20/20 Sheet Verified] All {len(posted_rows)} links confirmed in Today's Column {col_let}!")
        except Exception:
            pass

    def render_queue_table(self):
        for k in self.tree.get_children():
            self.tree.delete(k)
        sel_cnt = 0
        done_cnt = 0
        for r in self.rows:
            if r.get("selected", True):
                sel_cnt += 1
            st = str(r.get("status", "Pending"))
            bm_st = str(r.get("bookmark_status", "-"))
            if "Done" in st:
                done_cnt += 1
            tag = ""
            if "Done" in st or "Bookmarked" in bm_st:
                tag = "done"
            elif "Posting" in st:
                tag = "posting"
            elif "Failed" in st or "BANNED" in st:
                tag = "failed"
            self.tree.insert("", "end", iid=str(r["id"]), tags=(tag,) if tag else (), values=(
                "☑" if r.get("selected", True) else "☐",
                r["id"],
                st,
                r.get("title", ""),
                bm_st,
                r.get("url", "-")
            ))
        tot = len(self.rows)
        pct = round((done_cnt / tot) * 100) if tot > 0 else 0
        self.lbl_queue_title.config(text=f"Bookmarking Queue ({done_cnt} / {tot} items)")
        self.lbl_selected_count.config(text=f"  ({sel_cnt}/{tot} selected)")
        self.lbl_queue_pct.config(text=f"{pct}%")

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        row_id = self.tree.identify_row(event.y)
        if region == "cell" and row_id:
            col = self.tree.identify_column(event.x)
            rid = int(row_id)
            if col == "#1":
                for r in self.rows:
                    if r["id"] == rid:
                        r["selected"] = not r.get("selected", True)
                self.save_all_vaults_to_disk(False)
                self.render_queue_table()
            else:
                for r in self.rows:
                    if r["id"] == rid and str(r.get("url", "")).startswith("http"):
                        self.root.clipboard_clear()
                        self.root.clipboard_append(r["url"])
                        self.log_console(f"📋 Copied Link #{rid} to clipboard:\n{r['url']}")
                        break

    def on_tree_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            rid = int(row_id)
            for r in self.rows:
                if r["id"] == rid and str(r.get("url", "")).startswith("http"):
                    webbrowser.open(r["url"])
                    break

    def on_tree_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.ctx_row_id = int(row_id)
            self.tree.selection_set(row_id)
            self.ctx_menu.post(event.x_root, event.y_root)

    def select_all_rows(self, val):
        for r in self.rows:
            r["selected"] = bool(val)
        self.save_all_vaults_to_disk(False)
        self.render_queue_table()

    def invert_selection(self):
        for r in self.rows:
            r["selected"] = not r.get("selected", True)
        self.save_all_vaults_to_disk(False)
        self.render_queue_table()

    def ctx_toggle_row(self):
        if self.ctx_row_id:
            for r in self.rows:
                if r["id"] == self.ctx_row_id:
                    r["selected"] = not r.get("selected", True)
            self.save_all_vaults_to_disk(False)
            self.render_queue_table()

    def ctx_retry_row(self):
        if self.ctx_row_id:
            self.run_posting_worker(single_test=False, subset_ids={self.ctx_row_id})

    def ctx_copy_link(self):
        if self.ctx_row_id:
            for r in self.rows:
                if r["id"] == self.ctx_row_id and str(r.get("url", "")).startswith("http"):
                    self.root.clipboard_clear()
                    self.root.clipboard_append(r["url"])
                    self.log_console(f"📋 Copied Link #{r['id']} to clipboard:\n{r['url']}")

    def toggle_posting_automation(self):
        if self.is_posting:
            self.stop_all()
            return
        self.run_posting_worker(single_test=False, subset_ids=None)

    def test_single_post(self):
        if not self.is_posting:
            self.run_posting_worker(single_test=True, subset_ids=None)

    def post_selected_keywords(self):
        if self.is_posting:
            return
        ids = {r["id"] for r in self.rows if r.get("selected", True)}
        if not ids:
            messagebox.showwarning("None Selected", "Check (☑) at least 1 keyword first.")
            return
        self.run_posting_worker(single_test=False, subset_ids=ids)

    def retry_failed_or_selected(self):
        if self.is_posting:
            return
        failed = {r["id"] for r in self.rows if "Failed" in str(r.get("status", ""))}
        ids = failed if failed else {r["id"] for r in self.rows if r.get("selected", True)}
        if not ids:
            return
        self.run_posting_worker(single_test=False, subset_ids=ids)

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        self.btn_pause.config(text="▶ RESUME" if self.is_paused else "⏸ PAUSE")

    def stop_all(self):
        self.stop_requested = True
        self._posting_run_id = getattr(self, "_posting_run_id", 0) + 1
        self.is_posting = False
        self.is_bookmarking = False
        self.is_paused = False
        self.is_zoom_sending = False
        for r_item in getattr(self, "rows", []):
            if "Posting" in str(r_item.get("status", "")):
                r_item["status"] = "Pending"
        try:
            self.render_queue_table()
            self.lbl_live_timer.config(text="⏳ Timer: Ready", bg="#064e3b", fg="#34d399")
            self.btn_start.config(text="🚀 START POSTING (+ AUTO-BM)", bg="#10b981", fg="#042f2e")
        except Exception:
            pass
        self.log_console("🛑 Stop requested.")

    @staticmethod
    def _dispatch_cdp_trusted_click(ws, x, y):
        """
        Built-In Buster 'Client App' Native Interaction Simulator:
        Sends hardware-level CDP `Input.dispatchMouseEvent` (`mouseMoved` -> `mousePressed` -> `mouseReleased`)
        at viewport coordinates `(x, y)` so Chrome and Google reCAPTCHA see `event.isTrusted === true`!
        """
        if not ws or x <= 0 or y <= 0:
            return
        try:
            ws.send_json({
                "id": 9901,
                "method": "Input.dispatchMouseEvent",
                "params": {"type": "mouseMoved", "x": float(x), "y": float(y), "button": "none"}
            })
            ws.send_json({
                "id": 9902,
                "method": "Input.dispatchMouseEvent",
                "params": {"type": "mousePressed", "x": float(x), "y": float(y), "button": "left", "clickCount": 1}
            })
            ws.send_json({
                "id": 9903,
                "method": "Input.dispatchMouseEvent",
                "params": {"type": "mouseReleased", "x": float(x), "y": float(y), "button": "left", "clickCount": 1}
            })
        except Exception:
            pass

    @staticmethod
    def _get_cdp_captcha_solver_js():
        """
        Buster Whisper + Built-In Client App + ProBoards Modal Auto-Solver:
        1. Checks ProBoards' secondary checkbox inside 'Please prove you are human:' (`image_6eb85a.png`).
        2. Clicks `#recaptcha-anchor` ('I'm not a robot') and returns its exact viewport `(x, y)` for
           native CDP `Input.dispatchMouseEvent` trusted click simulation!
        3. Inside the reCAPTCHA `bframe` popup:
           - Looks for Buster's `#solver-button` (`title="Solve the challenge"`) first! When found, clicks it
             and returns `{action: "trusted_click", x, y, label: "Buster Whisper Solver"}` so Buster's
             Whisper model (`Model: Whisper, Status: ready to use`) solves the audio challenge automatically!
           - If Buster's button isn't injected yet, falls back to built-in audio challenge transcription.
        4. As soon as reCAPTCHA is solved, clicks ProBoards' `[Submit]` button!
        """
        return """
        (function() {
            try {
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            } catch (e) {}

            function fireMouseClick(el) {
                if (!el) return;
                try {
                    const rect = el.getBoundingClientRect();
                    const x = rect.left + rect.width * 0.48;
                    const y = rect.top + rect.height * 0.48;
                    ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click'].forEach(evType => {
                        el.dispatchEvent(new MouseEvent(evType, {
                            view: window, bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0
                        }));
                    });
                    if (typeof el.click === 'function') el.click();
                } catch (e) {
                    try { el.click(); } catch (e2) {}
                }
            }

            // 1. Check any secondary human-verification checkboxes inside ProBoards modal (image_6eb85a.png)
            const dialogs = document.querySelectorAll('.ui-dialog, .captcha-dialog, [role="dialog"], .modal, #ProveYouAreHuman');
            dialogs.forEach(dlg => {
                if (dlg.offsetParent !== null || dlg.style.display !== 'none') {
                    dlg.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                        if (!cb.checked) {
                            cb.checked = true;
                            cb.dispatchEvent(new Event('change', {bubbles: true}));
                        }
                    });
                }
            });
            document.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                const parentTxt = (cb.parentElement ? cb.parentElement.innerText : '') + (cb.closest('.ui-dialog') ? cb.closest('.ui-dialog').innerText : '');
                if (/prove you are human/i.test(parentTxt) && !cb.checked) {
                    cb.checked = true;
                    cb.dispatchEvent(new Event('change', {bubbles: true}));
                }
            });

            // 2. Inspect reCAPTCHA anchor iframe
            let recaptchaSolved = false;
            const gResp = document.querySelector('textarea[name="g-recaptcha-response"]');
            if (gResp && gResp.value && gResp.value.trim().length > 20) {
                recaptchaSolved = true;
            }

            const iframes = document.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {
                const fr = iframes[i];
                const src = (fr.src || '');
                if (src.indexOf('recaptcha') !== -1 && src.indexOf('anchor') !== -1 && fr.offsetParent !== null) {
                    try {
                        const adoc = fr.contentDocument || (fr.contentWindow && fr.contentWindow.document);
                        if (adoc) {
                            const anchor = adoc.querySelector('#recaptcha-anchor');
                            if (anchor) {
                                if (anchor.getAttribute('aria-checked') === 'true') {
                                    recaptchaSolved = true;
                                } else if (!window.__fsp_anchor_clicked_ts || (Date.now() - window.__fsp_anchor_clicked_ts > 4000)) {
                                    window.__fsp_anchor_clicked_ts = Date.now();
                                    const frRect = fr.getBoundingClientRect();
                                    const aRect = anchor.getBoundingClientRect();
                                    fireMouseClick(anchor);
                                    return {
                                        action: 'trusted_click',
                                        x: Math.round(frRect.left + aRect.left + aRect.width * 0.5),
                                        y: Math.round(frRect.top + aRect.top + aRect.height * 0.5),
                                        label: 'reCAPTCHA Anchor Checkbox'
                                    };
                                }
                            }
                        }
                    } catch (e) {}
                }
            }

            // 3. If reCAPTCHA is solved and ProBoards 'Submit' modal button is visible, click Submit automatically!
            if (recaptchaSolved) {
                const btns = document.querySelectorAll('.ui-dialog button, .ui-dialog input[type="submit"], .ui-dialog input[type="button"], button, input[type="submit"]');
                for (let b of btns) {
                    const txt = (b.innerText || b.value || '').trim().toLowerCase();
                    if ((txt === 'submit' || txt === 'continue' || txt === 'verify') && b.offsetParent !== null) {
                        const bRect = b.getBoundingClientRect();
                        fireMouseClick(b);
                        return {
                            action: 'trusted_click',
                            x: Math.round(bRect.left + bRect.width * 0.5),
                            y: Math.round(bRect.top + bRect.height * 0.5),
                            label: 'Modal Submit Button'
                        };
                    }
                }
            }

            // 4. Inspect reCAPTCHA challenge bframe: Trigger Buster Whisper '#solver-button' first!
            for (let i = 0; i < iframes.length; i++) {
                const fr = iframes[i];
                const src = (fr.src || '');
                if (src.indexOf('recaptcha') !== -1 && src.indexOf('bframe') !== -1 && fr.offsetParent !== null) {
                    try {
                        const bdoc = fr.contentDocument || (fr.contentWindow && fr.contentWindow.document);
                        if (bdoc) {
                            const frRect = fr.getBoundingClientRect();
                            // Check for Buster's '#solver-button' (or inside button-holder shadowroot)
                            let busterBtn = bdoc.querySelector('#solver-button, button#solver-button, .help-button-holder');
                            if (busterBtn && busterBtn.shadowRoot) {
                                busterBtn = busterBtn.shadowRoot.querySelector('#solver-button, button') || busterBtn;
                            }
                            if (busterBtn && (!window.__fsp_buster_clicked_ts || (Date.now() - window.__fsp_buster_clicked_ts > 4500))) {
                                window.__fsp_buster_clicked_ts = Date.now();
                                const sRect = busterBtn.getBoundingClientRect();
                                fireMouseClick(busterBtn);
                                return {
                                    action: 'trusted_click',
                                    x: Math.round(frRect.left + sRect.left + (sRect.width || 24) * 0.5),
                                    y: Math.round(frRect.top + sRect.top + (sRect.height || 24) * 0.5),
                                    label: 'Buster Whisper Solver Button'
                                };
                            }

                            const audSrc = bdoc.querySelector('#audio-source');
                            const dlLink = bdoc.querySelector('.rc-audiochallenge-tdownload-link');
                            const audUrl = (audSrc && audSrc.src) ? audSrc.src : ((dlLink && dlLink.href) ? dlLink.href : '');
                            const inp = bdoc.querySelector('#audio-response');
                            // Only run fallback audio transcription if Buster hasn't already filled #audio-response
                            if (audUrl && audUrl.indexOf('http') === 0 && (!inp || !inp.value.trim())) {
                                return {action: 'solve_audio', audio_url: audUrl};
                            }
                            const audBtn = bdoc.querySelector('#recaptcha-audio-button');
                            if (audBtn && !busterBtn && (!window.__fsp_aud_btn_ts || (Date.now() - window.__fsp_aud_btn_ts > 3500))) {
                                window.__fsp_aud_btn_ts = Date.now();
                                const abRect = audBtn.getBoundingClientRect();
                                fireMouseClick(audBtn);
                                return {
                                    action: 'trusted_click',
                                    x: Math.round(frRect.left + abRect.left + abRect.width * 0.5),
                                    y: Math.round(frRect.top + abRect.top + abRect.height * 0.5),
                                    label: 'Audio Challenge Button'
                                };
                            }
                        }
                    } catch (e) {}
                }
            }
            return {action: 'idle'};
        })()
        """

    @staticmethod
    def _get_cdp_fill_audio_js(phrase):
        safe_p = json.dumps(phrase or "")
        return f"""
        (function() {{
            const phrase = {safe_p};
            const iframes = document.querySelectorAll('iframe');
            for (let i = 0; i < iframes.length; i++) {{
                const fr = iframes[i];
                const src = (fr.src || '');
                if (src.indexOf('recaptcha') !== -1 && src.indexOf('bframe') !== -1) {{
                    try {{
                        const bdoc = fr.contentDocument || (fr.contentWindow && fr.contentWindow.document);
                        if (bdoc) {{
                            const inp = bdoc.querySelector('#audio-response');
                            if (inp) {{
                                inp.focus();
                                inp.value = phrase;
                                inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                                inp.dispatchEvent(new Event('change', {{bubbles: true}}));
                                const vBtn = bdoc.querySelector('#recaptcha-verify-button');
                                if (vBtn) {{
                                    vBtn.click();
                                    return true;
                                }}
                            }}
                        }}
                    }} catch (e) {{}}
                }}
            }}
            return false;
        }})()
        """

    def _get_active_debug_port(self):
        _, prof_label = self._get_persistent_chrome_profile_dir()
        prof_num = 1
        m_p = re.search(r"(\d+)", prof_label)
        if m_p:
            prof_num = int(m_p.group(1))
        return 9221 + prof_num

    def _sync_fresh_cookies_from_cdp(self, target_url, current_cookie_header):
        """
        Pulls fresh cookies (`pow_trace`, `pow_bypass`, `cf_clearance`, session cookies) from the live
        Chrome CDP browser OR saved profile cookie store and merges them cleanly into `current_cookie_header`.
        """
        debug_port = self._get_active_debug_port()
        cookie_file = self._get_profile_cookie_store_path()
        parsed_dom = (urllib.parse.urlparse(target_url).hostname or "").lower()
        fresh_pairs = []

        # 1. Try querying live Chrome CDP first
        try:
            req = urllib.request.Request(f"http://127.0.0.1:{debug_port}/json", headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=0.7) as r:
                targets = json.loads(r.read().decode("utf-8", errors="replace"))
            page_targets = [t for t in targets if t.get("type") == "page" and t.get("webSocketDebuggerUrl")]
            if page_targets:
                ws = _SimpleCDPWebSocket(page_targets[0]["webSocketDebuggerUrl"], timeout=1.2)
                ws.sock.settimeout(0.4)
                ws.send_json({"id": 77, "method": "Network.getAllCookies"})
                for _ in range(15):
                    evt = ws.recv_json()
                    if not evt:
                        break
                    cks = (evt.get("result") or {}).get("cookies")
                    if isinstance(cks, list):
                        for c in cks:
                            cdom = (c.get("domain") or "").lstrip(".").lower()
                            if cdom and (parsed_dom == cdom or parsed_dom.endswith("." + cdom)):
                                fresh_pairs.append(f"{c.get('name')}={c.get('value')}")
                        break
                ws.close()
        except Exception:
            pass

        # 2. Also check persisted cookie store on disk
        if not fresh_pairs and os.path.exists(cookie_file):
            try:
                with open(cookie_file, "r", encoding="utf-8") as f:
                    saved_cks = json.load(f)
                if isinstance(saved_cks, list):
                    for c in saved_cks:
                        cdom = (c.get("domain") or "").lstrip(".").lower()
                        if cdom and (parsed_dom == cdom or parsed_dom.endswith("." + cdom)):
                            fresh_pairs.append(f"{c.get('name')}={c.get('value')}")
            except Exception:
                pass

        if fresh_pairs:
            return merge_cookies_preserving_pow(current_cookie_header, fresh_pairs)
        return current_cookie_header

    def _cdp_live_browser_post_and_refresh(self, comp_u, target_url, title, raw_body):
        """
        Live Chrome CDP Fallback & Automatic CAPTCHA Recovery Engine:
        When direct HTTP returns HTTP 200 (CAPTCHA challenge / expired ProBoards `pow_bypass` / Guest check),
        this method:
        1. Connects to the live Stealth Chrome tab (launching Chrome if not open).
        2. Navigates to `comp_u` (`/thread/new/1`) so Chrome's V8 engine computes a fresh `pow_bypass`.
        3. Fills `guest_name`, `subject` (`title`), and `message` (`BBCode` body) directly in the browser DOM
           and clicks `Create Thread`.
        4. Runs `_get_cdp_captcha_solver_js()` + `transcribe_recaptcha_audio_url()` automatically if
           'Please prove you are human:' (`image_6eb85a.png` / `image_6eb593.png`) pops up!
        5. Returns `(published_thread_url, fresh_cookie_pairs)` so the current post succeeds AND all remaining
           posts in the batch get fresh `pow_trace` + `pow_bypass` cookies!
        """
        if not comp_u or not comp_u.startswith("http"):
            return "", []

        debug_port = self._get_active_debug_port()
        profile_dir, _ = self._get_persistent_chrome_profile_dir()
        parsed_dom = (urllib.parse.urlparse(target_url).hostname or "").lower()
        is_pb = is_proboards_url_or_curl(target_url) or is_proboards_url_or_curl(comp_u)
        clean_msg = html_to_clean_text(normalize_content_links_to_html(raw_body), use_bbcode=True) if is_pb else html_to_clean_text(raw_body, use_bbcode=False)

        # Ensure Stealth Chrome is running on `debug_port`
        targets = []
        try:
            req = urllib.request.Request(f"http://127.0.0.1:{debug_port}/json", headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=0.8) as r:
                targets = json.loads(r.read().decode("utf-8", errors="replace"))
        except Exception:
            targets = []

        if not targets:
            chrome_paths = [
                os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
            ]
            for exe in chrome_paths:
                if os.path.exists(exe):
                    try:
                        buster_dir = self._ensure_buster_extension_installed(profile_dir)
                        fb_args = [
                            exe,
                            f"--user-data-dir={profile_dir}",
                            f"--remote-debugging-port={debug_port}",
                            "--remote-allow-origins=*",
                            "--disable-blink-features=AutomationControlled",
                            "--disable-features=IsolateOrigins,site-per-process",
                            "--disable-site-isolation-trials",
                            "--disable-web-security",
                            "--autoplay-policy=no-user-gesture-required",
                            "--no-first-run",
                            "--no-default-browser-check",
                        ]
                        if buster_dir:
                            fb_args.append(f"--load-extension={buster_dir}")
                        fb_args.append(comp_u)
                        subprocess.Popen(
                            fb_args,
                            creationflags=0x08000000 if os.name == "nt" else 0
                        )
                        break
                    except Exception:
                        pass
            for _ in range(15):
                if self.stop_requested:
                    return "", []
                time.sleep(0.4)
                try:
                    req = urllib.request.Request(f"http://127.0.0.1:{debug_port}/json", headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=0.8) as r:
                        targets = json.loads(r.read().decode("utf-8", errors="replace"))
                    if targets:
                        break
                except Exception:
                    pass

        page_targets = [t for t in targets if t.get("type") == "page" and t.get("webSocketDebuggerUrl")]
        if not page_targets:
            return "", []

        # Pick tab matching target domain or first page tab
        chosen_t = page_targets[0]
        for pt in page_targets:
            if parsed_dom and parsed_dom in (pt.get("url") or "").lower():
                chosen_t = pt
                break

        ws = None
        try:
            ws = _SimpleCDPWebSocket(chosen_t["webSocketDebuggerUrl"], timeout=2.0)
            ws.sock.settimeout(0.15)
            ws.send_json({"id": 1, "method": "Page.enable", "params": {}})
            ws.send_json({"id": 2, "method": "Network.enable", "params": {}})
            ws.send_json({"id": 3, "method": "Page.addScriptToEvaluateOnNewDocument", "params": {
                "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
            }})
            ws.send_json({"id": 4, "method": "Page.navigate", "params": {"url": comp_u}})

            # Wait ~2.2s for composer page & ProBoards PoW JS to initialize
            for _ in range(11):
                if self.stop_requested:
                    return "", []
                time.sleep(0.2)
                try:
                    ws.recv_json()
                except Exception:
                    pass

            # Fill Guest Name, Subject, and Body in the live DOM and click Create Thread / Submit
            fill_and_submit_js = f"""
            (function() {{
                const postTitle = {json.dumps(title)};
                const postBody = {json.dumps(clean_msg)};
                try {{
                    const gn = document.querySelector('input[name="guest_name"]');
                    if (gn && !gn.value.trim()) {{
                        gn.value = 'xfinity';
                        gn.dispatchEvent(new Event('input', {{bubbles: true}}));
                    }}
                    // Switch ProBoards WYSIWYG to BBCode tab if present so [url=...] tags aren't escaped
                    const bbTab = document.querySelector('.wysiwyg-tabs a[href*="bbcode"], li[id*="menu-item-bbcode"] a');
                    if (bbTab) {{
                        try {{ bbTab.click(); }} catch (e) {{}}
                    }}
                    const subj = document.querySelector('input[name="subject"], input[name="title"], input[name="topic_title"]');
                    if (subj) {{
                        subj.value = postTitle;
                        subj.dispatchEvent(new Event('input', {{bubbles: true}}));
                        subj.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    const msg = document.querySelector('textarea[name="message"], textarea[name="body"], textarea[name="content"]');
                    if (msg) {{
                        msg.value = postBody;
                        msg.dispatchEvent(new Event('input', {{bubbles: true}}));
                        msg.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    // Also update ProBoards visual editor iframe if active
                    const wysFrame = document.querySelector('.wysiwyg-area iframe');
                    if (wysFrame && wysFrame.contentDocument && wysFrame.contentDocument.body && (!msg || msg.offsetParent === null)) {{
                        wysFrame.contentDocument.body.innerText = postBody;
                    }}
                    const subBtn = document.querySelector('input[type="submit"][value*="Create"], input[type="submit"][value*="Post"], button[type="submit"], .form_post_button input');
                    if (subBtn) {{
                        subBtn.click();
                        return 'submitted';
                    }}
                }} catch (e) {{
                    return 'err:' + String(e);
                }}
                return 'no_btn';
            }})()
            """
            ws.send_json({"id": 10, "method": "Runtime.evaluate", "params": {"expression": fill_and_submit_js, "returnByValue": True}})

            published_u = ""
            fresh_cookies = []
            msg_seq = 20

            # Poll for up to 35 seconds: auto-solve CAPTCHA/Audio challenge if modal appears, and detect thread redirect!
            for step_i in range(20):
                if self.stop_requested:
                    break
                time.sleep(0.5)
                msg_seq += 1
                ws.send_json({
                    "id": msg_seq,
                    "method": "Runtime.evaluate",
                    "params": {
                        "expression": """
                        (function() {
                            const u = window.location.href;
                            const solverRes = """ + self._get_cdp_captcha_solver_js().strip() + """;
                            return {url: u, solver: solverRes};
                        })()
                        """,
                        "returnByValue": True
                    }
                })

                for _ in range(25):
                    try:
                        evt = ws.recv_json()
                    except Exception:
                        break
                    if not evt:
                        break
                    r_obj = evt.get("result") or {}
                    if "cookies" in r_obj and isinstance(r_obj["cookies"], list):
                        for c in r_obj["cookies"]:
                            cdom = (c.get("domain") or "").lstrip(".").lower()
                            if cdom and (parsed_dom == cdom or parsed_dom.endswith("." + cdom)):
                                fresh_cookies.append(f"{c.get('name')}={c.get('value')}")

                    val = (r_obj.get("result") or {}).get("value") if isinstance(r_obj.get("result"), dict) else None
                    if isinstance(val, dict):
                        cur_u = str(val.get("url") or "").strip()
                        if cur_u and cur_u != comp_u and not re.search(r'/(?:thread/new|thread/create|post-thread|threads/add|login|register)(?:[/?#]|$)', cur_u, re.I):
                            if re.search(r'/(?:thread|threads|topic|discussion|post|p|d|t)/\d+', cur_u, re.I):
                                published_u = cur_u
                                break

                        s_info = val.get("solver") or {}
                        if isinstance(s_info, dict) and s_info.get("action") == "trusted_click":
                            tx = int(s_info.get("x") or 0)
                            ty = int(s_info.get("y") or 0)
                            if tx > 0 and ty > 0:
                                self._dispatch_cdp_trusted_click(ws, tx, ty)
                        elif isinstance(s_info, dict) and s_info.get("action") == "solve_audio" and s_info.get("audio_url"):
                            aud_u = str(s_info["audio_url"])
                            if aud_u != getattr(self, "_last_solved_audio_url", ""):
                                self._last_solved_audio_url = aud_u
                                self.log_console("🛡️ [Auto-CAPTCHA] Transcribing reCAPTCHA Audio Challenge...")
                                phrase = transcribe_recaptcha_audio_url(aud_u)
                                if phrase:
                                    self.log_console(f"🛡️ [Auto-CAPTCHA] Audio solved ('{phrase}') — verifying & submitting!")
                                    msg_seq += 1
                                    ws.send_json({
                                        "id": msg_seq,
                                        "method": "Runtime.evaluate",
                                        "params": {"expression": self._get_cdp_fill_audio_js(phrase), "returnByValue": True}
                                    })
                if published_u:
                    msg_seq += 1
                    ws.send_json({"id": msg_seq, "method": "Network.getAllCookies"})
                    time.sleep(0.25)
                    for _ in range(15):
                        try:
                            evt = ws.recv_json()
                            if not evt:
                                break
                            cks = (evt.get("result") or {}).get("cookies")
                            if isinstance(cks, list):
                                for c in cks:
                                    cdom = (c.get("domain") or "").lstrip(".").lower()
                                    if cdom and (parsed_dom == cdom or parsed_dom.endswith("." + cdom)):
                                        fresh_cookies.append(f"{c.get('name')}={c.get('value')}")
                                break
                        except Exception:
                            break
                    break
            return published_u, fresh_cookies
        except Exception:
            return "", []
        finally:
            if ws:
                ws.close()

    def run_posting_worker(self, single_test=False, subset_ids=None):
        self.save_all_vaults_to_disk(True)
        raw_curl = self.txt_curl.get("1.0", "end").strip()
        if "curl" not in raw_curl.lower():
            messagebox.showwarning("Missing cURL", "Paste a valid cURL command in Box 3 first.")
            return
        if not self.rows:
            messagebox.showwarning("Empty Queue", "Click '🔄 Sync Sheets' or '💾 Save & Load Sheets' first.")
            return

        custom_col = self.entry_sheet_col.get().strip().upper()
        my_run_id = getattr(self, "_posting_run_id", 0) + 1
        self._posting_run_id = my_run_id
        self.is_posting = True
        self.stop_requested = False
        self.is_paused = False
        for r_clean in self.rows:
            if "Posting" in str(r_clean.get("status", "")):
                r_clean["status"] = "Pending"
        self.render_queue_table()
        self.btn_start.config(text="🛑 STOP POSTING", bg="#e11d48", fg="white")

        def _is_aborted():
            return self.stop_requested or getattr(self, "_posting_run_id", 0) != my_run_id

        def _worker():
            posted_this_run = []
            sheet_saved_count = 0
            # Auto-heal TMDB content_check / duplicate_check URL if present in Box 3
            nonlocal raw_curl
            if "themoviedb.org/movie/content_check" in raw_curl or "themoviedb.org/movie/duplicate_check" in raw_curl:
                raw_curl = raw_curl.replace("themoviedb.org/movie/content_check", "themoviedb.org/movie/new").replace(
                    "themoviedb.org/movie/duplicate_check", "themoviedb.org/movie/new"
                )
                self.root.after(0, lambda rc=raw_curl: (self.txt_curl.delete("1.0", "end"), self.txt_curl.insert("1.0", rc)))

            # Maintain a persistent live Cookie header across all 20 posts while protecting `(pow_trace, pow_bypass)`!
            _, _, init_hdrs, _ = parse_curl_command(raw_curl)
            live_cookie_str = init_hdrs.get("Cookie") or init_hdrs.get("cookie") or ""

            try:
                for r_norm in self.rows:
                    r_norm["body"] = normalize_content_links_to_html(r_norm.get("body", ""))

                # Determine target Sheet Column:
                # In AUTO mode, ALWAYS refresh Today's CSV table from the live sheet so we NEVER reuse a stale
                # `self.next_open_col` from an earlier run and NEVER overwrite existing links in Columns E..J!
                info_live = self.get_todays_table_info_from_csv(force_refresh=True)
                if info_live:
                    self.latest_active_col = info_live["latest_active_col"]
                    self.next_open_col = info_live["next_open_col"]
                    self._main_sheet_synced_once = True
                    self.root.after(0, lambda la=self.latest_active_col, no=self.next_open_col: self.lbl_batch_status.config(
                        text=f" ✅ Active Batch: {len(self.rows)} Keywords (Rows 3–{2 + len(self.rows)}) | Today's Col: {la} | Next Empty: {no} "
                    ))

                if custom_col and custom_col not in ("AUTO", "A", "B", "C"):
                    locked_col = custom_col
                else:
                    if info_live:
                        # If retrying failed items from the immediately preceding batch, keep that batch's column
                        # ONLY if the target retry rows are currently empty in that column!
                        last_b_col = getattr(self, "_last_batch_locked_col", "")
                        if subset_ids is not None and last_b_col and last_b_col not in ("AUTO", "A", "B", "C"):
                            existing_in_last = dict(info_live["col_links_map"].get(last_b_col, []))
                            if not any(existing_in_last.get(2 + int(rid), "").startswith("http") for rid in subset_ids):
                                locked_col = last_b_col
                            else:
                                locked_col = info_live["next_open_col"]
                        else:
                            locked_col = info_live["next_open_col"]
                    else:
                        c_let, f_cnt, _, _ = self.query_rightmost_sheet_column_state("AUTO")
                        if c_let and c_let not in ("AUTO", "A", "B", "C"):
                            locked_col = self._num_to_col(self._col_to_num(c_let) + 1) if f_cnt > 0 else c_let
                        else:
                            locked_col = "AUTO" 

                self._last_batch_locked_col = locked_col

                items = [
                    r for r in self.rows
                    if (subset_ids is None or r["id"] in subset_ids)
                    and (subset_ids is not None or r.get("selected", True))
                ]
                if single_test and items:
                    items = [items[0]]

                for pos, item in enumerate(items):
                    if _is_aborted():
                        break
                    while self.is_paused and not _is_aborted():
                        time.sleep(0.3)

                    item["status"] = "🔵 Posting..."
                    for r_sync in self.rows:
                        if r_sync.get("id") == item["id"]:
                            r_sync["status"] = "🔵 Posting..."
                    self.root.after(0, self.render_queue_table)
                    self.log_console(f"=== [Post #{item['id']}/{len(self.rows)}] \"{item['title'][:52]}...\" ===")

                    live_link = ""
                    last_status = 0
                    last_resp_snippet = ""
                    comp_u = ""
                    saw_captcha_challenge = False

                    for attempt in range(1, 3):
                        if _is_aborted():
                            break
                        try:
                            t_u, meth, hdrs, body_s = build_posting_request(
                                raw_curl, item["title"], item["body"], force_short_plain=(attempt == 2)
                            )
                            if live_cookie_str:
                                hdrs.pop("cookie", None)
                                hdrs["Cookie"] = live_cookie_str

                            if attempt == 2:
                                live_cookie_str = self._sync_fresh_cookies_from_cdp(t_u, live_cookie_str)
                                if live_cookie_str:
                                    hdrs["Cookie"] = live_cookie_str

                            is_pb_site = is_proboards_url_or_curl(t_u)
                            ui_tu = self.entry_target_url.get().strip()
                            comp_u = (
                                ui_tu
                                if (ui_tu and is_same_base_domain(ui_tu, t_u) and (urllib.parse.urlparse(ui_tu).path or "").strip("/"))
                                else (derive_composer_url_from_curl(raw_curl) or ui_tu)
                            )

                            # Fast token pre-fetch ONLY on ProBoards (or Attempt 2 on traditional HTML form forums, never on JSON/AJAX endpoints like TMDB!)
                            is_ajax_or_json = (
                                "themoviedb.org" in t_u.lower()
                                or "pixiv.net" in t_u.lower()
                                or "/api/" in t_u.lower()
                                or any(k.lower() == "x-requested-with" for k in hdrs)
                            )
                            if (is_pb_site or (attempt == 2 and not is_ajax_or_json)) and comp_u and comp_u.startswith("http"):
                                try:
                                    get_hdrs = {k: v for k, v in hdrs.items() if k.lower() != "content-type"}
                                    r_get = execute_direct_http(comp_u, "GET", get_hdrs, "")
                                    g_body = r_get.get("body") or ""
                                    if r_get.get("set_cookies"):
                                        live_cookie_str = merge_cookies_preserving_pow(live_cookie_str, r_get["set_cookies"])
                                        hdrs["Cookie"] = live_cookie_str
                                    body_s = extract_composer_form_tokens(g_body, body_s)
                                except Exception:
                                    pass

                            res = execute_direct_http(t_u, meth, hdrs, body_s)
                            last_status = int(res.get("status") or 0)
                            if res.get("set_cookies"):
                                live_cookie_str = merge_cookies_preserving_pow(live_cookie_str, res["set_cookies"])
                                hdrs["Cookie"] = live_cookie_str

                            b_preview = (res.get("body") or "").strip()
                            if b_preview.startswith("{"):
                                try:
                                    j_err = json.loads(b_preview)
                                    if j_err.get("errors") or j_err.get("message") or j_err.get("error") or j_err.get("status_message"):
                                        last_resp_snippet = str(
                                            j_err.get("errors") or j_err.get("message") or j_err.get("error") or j_err.get("status_message")
                                        )[:120]
                                except Exception:
                                    pass

                            live_link = extract_published_url_from_response(t_u, item["title"], res)
                            if live_link:
                                break

                            issue_type, wait_sec = detect_captcha_or_flood_in_response(last_status, b_preview)
                            if issue_type == "flood" and not _is_aborted():
                                self.log_console(f"   ⏳ [Flood Control] Forum requested {wait_sec}s cooldown — waiting & retrying...")
                                for s_w in range(wait_sec, 0, -1):
                                    if _is_aborted():
                                        break
                                    self.root.after(0, lambda sw=s_w, iid=item["id"]: self.lbl_live_timer.config(
                                        text=f"⏳ Cooldown ({sw}s) for #{iid}...", bg="#78350f", fg="#fde68a"
                                    ))
                                    time.sleep(1.0)
                                continue
                            elif issue_type == "captcha" and not _is_aborted():
                                saw_captcha_challenge = True
                                last_resp_snippet = "CAPTCHA / Human verification challenge detected"
                        except Exception as e:
                            self.log_console(f"   ⚠️ Attempt {attempt} error: {e}")
                        time.sleep(0.2)

                    # ONLY invoke Live Chrome CDP Fallback when an actual CAPTCHA challenge was detected!
                    if not live_link and saw_captcha_challenge and not _is_aborted() and comp_u and comp_u.startswith("http"):
                        self.log_console(f"   🛡️ [Auto-CAPTCHA Shield] CAPTCHA detected on #{item['id']} — solving via Stealth Chrome CDP...")
                        self.root.after(0, lambda iid=item["id"]: self.lbl_live_timer.config(
                            text=f"🛡️ Auto-Solving CAPTCHA (#{iid})...", bg="#7c3aed", fg="white"
                        ))
                        cdp_url, cdp_cookies = self._cdp_live_browser_post_and_refresh(
                            comp_u, t_u, item["title"], item["body"]
                        )
                        if cdp_cookies:
                            live_cookie_str = merge_cookies_preserving_pow(live_cookie_str, cdp_cookies)
                        if cdp_url:
                            live_link = cdp_url

                    if live_link:

                        item["status"] = "🟢 Done"
                        item["url"] = live_link
                        item["bookmark_status"] = "⏳ Queued"
                        for r_sync in self.rows:
                            if r_sync.get("id") == item["id"]:
                                r_sync["status"] = "🟢 Done"
                                r_sync["url"] = live_link
                                r_sync["bookmark_status"] = "⏳ Queued"
                        posted_this_run.append(item)
                        self.add_post_to_history(self.active_site or extract_domain_from_curl(raw_curl), item["title"], live_link, locked_col, 2 + int(item["id"]))
                        self.save_all_vaults_to_disk(False)
                        self.root.after(0, self.render_queue_table)
                        self.log_console(f"   ✅ [POSTED #{item['id']}] -> {live_link}")

                        # 1. Put link in Sheet FIRST
                        if self.var_save_sheet.get():
                            ret_col = self.save_link_to_sheet_sync(item["id"], item["title"], live_link, locked_col, is_auto_mode=(not custom_col or custom_col == "AUTO"))
                            if isinstance(ret_col, str) and ret_col not in ("AUTO", "", "A", "B", "C"):
                                locked_col = ret_col
                                self._last_batch_locked_col = locked_col
                                self.latest_active_col = locked_col
                                next_after_locked = self._num_to_col(self._col_to_num(locked_col) + 1)
                                if self._col_to_num(next_after_locked) > self._col_to_num(self.next_open_col or "D"):
                                    self.next_open_col = next_after_locked
                                self.root.after(0, lambda la=self.latest_active_col, no=self.next_open_col: self.lbl_batch_status.config(
                                    text=f" ✅ Active Batch: {len(self.rows)} Keywords (Rows 3–{2 + len(self.rows)}) | Today's Col: {la} | Next Empty: {no} "
                                ))
                                sheet_saved_count += 1

                        # 2. If Auto-Send Zoom is checked on Tab 1, send this 1 link to Zoom & minimize Zoom!
                        if self.var_auto_zoom.get() and not _is_aborted():
                            self._win32_send_single_link_to_zoom(live_link)

                        if not single_test and pos + 1 < len(items) and not _is_aborted():
                            try:
                                dly = max(1, int(self.spin_delay.get()))
                            except Exception:
                                dly = 25
                            for elapsed_s in range(dly):
                                if _is_aborted():
                                    break
                                while self.is_paused and not _is_aborted():
                                    self.root.after(0, lambda: self.lbl_live_timer.config(text="⏸ PAUSED", bg="#78350f", fg="#fde68a"))
                                    time.sleep(0.3)
                                rem_s = max(0, dly - elapsed_s)
                                nxt_id = items[pos + 1]["id"]
                                self.root.after(0, lambda r=rem_s, nid=nxt_id: self.lbl_live_timer.config(text=f"⏳ Next Post (#{nid}) in {r}s...", bg="#064e3b", fg="#34d399"))
                                time.sleep(1.0)
                            self.root.after(0, lambda: self.lbl_live_timer.config(text="⏳ Timer: Posting...", bg="#1e3a8a", fg="#93c5fd"))
                    else:
                        item["status"] = "❌ Failed"
                        for r_sync in self.rows:
                            if r_sync.get("id") == item["id"]:
                                r_sync["status"] = "❌ Failed"
                        self.save_all_vaults_to_disk(False)
                        self.root.after(0, self.render_queue_table)
                        self.log_console(f"   🚨 [FAILED #{item['id']}] HTTP {last_status} — {last_resp_snippet or 'No live post URL returned (check login/cURL)'}")
                        if single_test:
                            break

                if not single_test and not _is_aborted() and self.var_save_sheet.get() and posted_this_run and sheet_saved_count < len(posted_this_run):
                    self.verify_and_repair_sheet_column(locked_col)

                if not single_test and not _is_aborted() and self.var_auto_bm.get() and posted_this_run:
                    self.run_bookmarking_batch_sync(posted_this_run)
            finally:
                if getattr(self, "_posting_run_id", 0) == my_run_id:
                    self.is_posting = False
                    for r_rem in self.rows:
                        if "Posting" in str(r_rem.get("status", "")):
                            r_rem["status"] = "Pending"
                    self.root.after(0, self.render_queue_table)
                    self.root.after(0, lambda: self.lbl_live_timer.config(text="⏳ Timer: Ready", bg="#064e3b", fg="#34d399"))
                    self.root.after(0, lambda: self.btn_start.config(text="🚀 START POSTING (+ AUTO-BM)", bg="#10b981", fg="#042f2e"))
                    self.log_console("=== 🏁 Automation Finished ===")

        threading.Thread(target=_worker, daemon=True).start()

    @staticmethod
    def _is_bookmark_http_ok(res):
        st_code = int(res.get("status", 0) or 0)
        body_txt = (res.get("body") or "").strip()
        if not (200 <= st_code < 400):
            return False, f"HTTP {st_code} — {body_txt[:120]}"
        if body_txt.startswith("{"):
            try:
                jr = json.loads(body_txt)
                if "retcode" in jr and int(jr.get("retcode", 0)) != 0:
                    return False, f"retcode {jr.get('retcode')}: {jr.get('message') or body_txt[:120]}"
                if jr.get("errors") or jr.get("error") or jr.get("status") == "error" or jr.get("ok") is False:
                    err_m = str(jr.get("errors") or jr.get("error") or jr.get("message") or body_txt[:120])
                    return False, err_m
            except Exception:
                pass
        return True, f"HTTP {st_code}"

    def run_bookmarking_batch_sync(self, items_list):
        self.save_all_vaults_to_disk(True)
        active_accs = [b for b in self.bookmarks if b.get("active") and "curl" in str(b.get("curl", "")).lower()]
        total_accs = len(self.bookmarks) or len(active_accs)
        if not active_accs or not items_list:
            return

        bulk_accs = [b for b in active_accs if str(b.get("type", "SINGLE")).strip().upper() == "BULK"]
        single_accs = [b for b in active_accs if str(b.get("type", "SINGLE")).strip().upper() != "BULK"]

        per_item_ok = {id(it): 0 for it in items_list}

        try:
            bm_wait = max(0, int(self.spin_bm_delay.get())) if hasattr(self, "spin_bm_delay") else int(getattr(self, "bm_delay", 2))
        except Exception:
            bm_wait = 2

        # 1. Execute BULK accounts ONCE with ALL eligible links combined!
        for acc in bulk_accs:
            if self.stop_requested:
                break
            acc_name = acc.get("site", "BulkAccount")
            is_hat = "hatena" in acc_name.lower()
            eligible_items = [
                it for it in items_list
                if not (is_hat and self.var_hatena_shield.get() and is_proboards_url_or_curl(it.get("url", "")))
            ]
            if not eligible_items:
                continue
            bulk_urls = [it["url"] for it in eligible_items if str(it.get("url", "")).startswith("http")]
            bulk_title = eligible_items[0].get("title", "")
            try:
                t_u, m, h, b_s = build_bookmark_request(acc["curl"], bulk_urls, bulk_title)
                res = execute_direct_http(t_u, m, h, b_s)
                is_ok, info_msg = self._is_bookmark_http_ok(res)
                if is_ok:
                    for it in eligible_items:
                        per_item_ok[id(it)] += 1
                    pub_u = extract_published_url_from_response(t_u, bulk_title, res)
                    self.log_console(
                        f"🔖 [BULK {acc_name}] Posted {len(bulk_urls)} links!" + (f" -> {pub_u}" if pub_u else "")
                    )
                else:
                    self.log_console(f"⚠️ [BULK {acc_name}] Failed: {info_msg}")
            except Exception as e:
                self.log_console(f"⚠️ [BULK {acc_name}] Error: {e}")

        # 2. Execute SINGLE accounts 1 link at a time with configured Delay (2s default) + console logging
        for pos, item in enumerate(items_list):
            if self.stop_requested:
                break
            did_send_any = False
            for acc in single_accs:
                if self.stop_requested:
                    break
                acc_name = acc.get("site", "Account")
                is_hat = "hatena" in acc_name.lower() or "hatena.ne.jp" in str(acc.get("curl", "")).lower()
                if is_hat and self.var_hatena_shield.get() and is_proboards_url_or_curl(item.get("url", "")):
                    continue
                try:
                    t_u, m, h, b_s = build_bookmark_request(acc["curl"], [item["url"]], item.get("title", ""))
                    res = execute_direct_http(t_u, m, h, b_s)
                    did_send_any = True
                    is_ok, info_msg = self._is_bookmark_http_ok(res)
                    if is_ok:
                        per_item_ok[id(item)] += 1
                        self.log_console(f"🔖 [{acc_name}] #{item.get('id', pos+1)} Bookmarked -> {item['url']}")
                    else:
                        self.log_console(f"⚠️ [{acc_name}] #{item.get('id', pos+1)} Failed: {info_msg}")
                except Exception as e:
                    self.log_console(f"⚠️ [{acc_name}] #{item.get('id', pos+1)} Error: {e}")

            item["bookmark_status"] = f"☑️ Bookmarked ({per_item_ok[id(item)]}/{total_accs})"
            self.save_all_vaults_to_disk(False)
            self.root.after(0, self.render_queue_table)
            if did_send_any and bm_wait > 0 and pos + 1 < len(items_list) and not self.stop_requested:
                for _ in range(bm_wait * 5):
                    if self.stop_requested:
                        break
                    time.sleep(0.2)

        if not single_accs:
            for item in items_list:
                item["bookmark_status"] = f"☑️ Bookmarked ({per_item_ok[id(item)]}/{total_accs})"
            self.save_all_vaults_to_disk(False)
            self.root.after(0, self.render_queue_table)

    def start_queue_bookmarking(self):
        ready = [r for r in self.rows if r.get("selected", True) and str(r.get("url", "")).startswith("http")]
        if not ready:
            messagebox.showwarning("No Links", "No selected rows have posted URLs yet.")
            return
        self.stop_requested = False
        threading.Thread(target=lambda: self.run_bookmarking_batch_sync(ready), daemon=True).start()

    def update_custom_bm_count(self):
        raw = self.txt_custom_bm.get("1.0", "end")
        urls = re.findall(r'https?://(?:(?!https?://)[^\s"\'<>])+', raw)
        self.lbl_custom_bm_cnt.config(text=f"🔗 Paste Any Link(s) to Bookmark ({len(urls)} links ready):")

    def paste_custom_bm(self):
        try:
            cb = self._get_any_clipboard_text()
            urls = re.findall(r'https?://(?:(?!https?://)[^\s"\'<>])+', cb)
            self.txt_custom_bm.delete("1.0", "end")
            self.txt_custom_bm.insert("1.0", "\n".join(urls))
            self.update_custom_bm_count()
        except Exception:
            pass

    def load_queue_to_custom_bm(self):
        urls = [r["url"] for r in self.rows if str(r.get("url", "")).startswith("http")]
        self.txt_custom_bm.delete("1.0", "end")
        self.txt_custom_bm.insert("1.0", "\n".join(urls))
        self.update_custom_bm_count()

    def clear_custom_bm(self):
        self.txt_custom_bm.delete("1.0", "end")
        self.entry_custom_bm_title.delete(0, "end")
        self.update_custom_bm_count()

    def _lookup_sheet_keyword_for_url(self, target_u):
        u_clean = (target_u or "").strip().rstrip("/")
        for r in getattr(self, "rows", []):
            ru = str(r.get("url", "")).strip().rstrip("/")
            if ru and ru == u_clean and str(r.get("title", "")).strip():
                return str(r["title"]).strip()
        for h in getattr(self, "history", []):
            hu = str(h.get("url", "")).strip().rstrip("/")
            if hu and hu == u_clean and str(h.get("title", "")).strip():
                return str(h["title"]).strip()
        return ""

    def start_custom_links_bookmarking(self):
        self.save_all_vaults_to_disk(True)
        raw = self.txt_custom_bm.get("1.0", "end")
        urls = re.findall(r'https?://(?:(?!https?://)[^\s"\'<>])+', raw)
        if not urls:
            messagebox.showwarning("No Links", "Paste at least 1 valid URL first.")
            return
        custom_t = self.entry_custom_bm_title.get().strip()
        items = []
        for i, u in enumerate(urls):
            matched_kw = custom_t or self._lookup_sheet_keyword_for_url(u)
            if not matched_kw and self.rows:
                # If URL came from a previous batch (not yet in current self.rows), fall back to URL slug cleanly
                matched_kw = derive_clean_title_from_url(u, "")
            items.append({"id": i + 1, "url": u, "title": derive_clean_title_from_url(u, matched_kw)})
        self.stop_requested = False
        self.lbl_bm_status.config(text=f"🚀 Bookmarking {len(items)} custom link(s)...")
        def _bg():
            self.run_bookmarking_batch_sync(items)
            self.root.after(0, lambda: self.lbl_bm_status.config(text=f"✅ Finished bookmarking {len(items)} custom link(s)!"))
        threading.Thread(target=_bg, daemon=True).start()

    def copy_all_posted_links(self):
        lines = [r["url"] if str(r.get("url", "")).startswith("http") else "" for r in self.rows]
        while lines and not lines[-1]:
            lines.pop()
        if not any(lines):
            messagebox.showinfo("No Links", "No posted links yet.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        messagebox.showinfo("Copied Links", f"Copied {sum(1 for x in lines if x)} links in exact row order (Rows 3–{2 + len(lines)})!")

    # --------------------------------------------------------------------------
    # TAB 3: TODAY-ONLY SHEET LOADER + PER-LINK (SHEET FIRST -> ZOOM SECOND -> MINIMIZE ZOOM)
    # --------------------------------------------------------------------------
    def load_column_links_from_sheet(self):
        """
        Loads links ONLY from TODAY'S top table (Rows 3..22) for the requested column!
        Strictly stops at Row 22 so Yesterday's links (Row 27+) are NEVER loaded!
        """
        self.save_all_vaults_to_disk(True)
        req_col = (self.entry_zoom_read_col.get() or "AUTO").strip().upper()

        def _load():
            links = []
            used_col = req_col
            info = self.get_todays_table_info_from_csv()
            if info:
                used_col = info["latest_active_col"] if req_col in ("AUTO", "") else req_col
                pairs = info["col_links_map"].get(used_col, [])
                links = [u for _, u in pairs]
            else:
                col_let, _, _, lnks = self.query_rightmost_sheet_column_state(req_col)
                used_col = col_let
                links = lnks

            def _upd_ui():
                self.txt_zoom_links.config(state="normal")
                self.txt_zoom_links.delete("1.0", "end")
                self.txt_zoom_links.insert("1.0", "\n".join(links))
                self.lbl_zoom_msg.config(text=f"✅ Loaded {len(links)} TODAY'S link(s) from Column {used_col} (Rows 3–22 only)!")
            self.root.after(0, _upd_ui)

        threading.Thread(target=_load, daemon=True).start()

    def split_and_copy_zoom_links(self):
        self.txt_zoom_links.config(state="normal")
        raw = self.txt_zoom_links.get("1.0", "end")
        urls = re.findall(r'https?://(?:(?!https?://)[^\s"\'<>])+', raw)
        self.txt_zoom_links.delete("1.0", "end")
        self.txt_zoom_links.insert("1.0", "\n".join(urls))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(urls))
        self.lbl_zoom_msg.config(text=f"✂️ Split & copied {len(urls)} clean links!")

    def _resolve_starting_sheet_slot(self, req_col="AUTO"):
        """
        Determines the exact starting (col_letter, next_item_num_1_to_20, max_rows_per_col)
        in TODAY'S table (Rows 3..22) so links fill sequentially from the last filled row
        in the active column (Row 3 -> Row 22) and wrap to Row 3 of the next column!
        """
        clean_req = (req_col or "AUTO").strip().upper()
        info = self.get_todays_table_info_from_csv(force_refresh=True)
        if info:
            max_rows = max(20, info["kw_count"])
            if clean_req != "AUTO" and clean_req not in ("", "A", "B", "C", "D"):
                cur_col = clean_req
            else:
                lat_col = info["latest_active_col"]
                lat_pairs = info["col_links_map"].get(lat_col, [])
                last_r = max((r for r, _ in lat_pairs), default=2)
                if 3 <= last_r < (2 + max_rows):
                    cur_col = lat_col
                else:
                    cur_col = info["next_open_col"]
            cur_pairs = info["col_links_map"].get(cur_col, [])
            last_row_in_cur = max((r for r, _ in cur_pairs), default=2)
            filled_cnt = max(0, last_row_in_cur - 2)  # 0 if empty (next is Row 3), or N if Row 2+N is last filled
        else:
            col_let, f_cnt, kw_cnt, _ = self.query_rightmost_sheet_column_state(clean_req)
            max_rows = max(20, kw_cnt or 20)
            cur_col = clean_req if (clean_req and clean_req not in ("AUTO", "A", "B", "C")) else col_let
            filled_cnt = f_cnt

        if filled_cnt >= max_rows:
            cur_col = self._num_to_col(self._col_to_num(cur_col) + 1)
            filled_cnt = 0

        return cur_col, filled_cnt, max_rows

    @staticmethod
    def _win32_set_clipboard_single_url(single_url):
        """
        Writes a SINGLE URL directly to the native Windows OS Clipboard (CF_UNICODETEXT = 13)
        via ctypes (zero PowerShell, zero CMD window, and clears any multi-line clipboard first!).
        """
        if os.name != "nt":
            return
        import ctypes
        clean_u = str(single_url or "").strip().splitlines()[0].strip()
        CF_UNICODETEXT = 13
        GMEM_MOVEABLE = 0x0002
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32

        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p

        data_bytes = (clean_u + "\x00").encode("utf-16-le")
        for _ in range(5):
            if user32.OpenClipboard(0):
                try:
                    user32.EmptyClipboard()
                    h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data_bytes))
                    if h_mem:
                        ptr = kernel32.GlobalLock(h_mem)
                        if ptr:
                            ctypes.memmove(ptr, data_bytes, len(data_bytes))
                            kernel32.GlobalUnlock(h_mem)
                            user32.SetClipboardData(CF_UNICODETEXT, h_mem)
                finally:
                    user32.CloseClipboard()
                break
            time.sleep(0.03)

    def _win32_send_single_link_to_zoom(self, single_url):
        """
        Sends EXACTLY ONE link to Zoom Desktop ('Zoom Workplace' / 'Zoom') and IMMEDIATELY
        minimizes the Zoom window off the screen so it NEVER stays open blocking the screen!
        Uses 100% native Win32 API (zero PowerShell / zero CMD windows).
        """
        if os.name != "nt":
            return False, "Zoom desktop automation requires Windows."

        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        def get_window_title(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return ""
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            return buf.value or ""

        def get_class_name(hwnd):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            return buf.value or ""

        def find_zoom_windows():
            matches = []
            WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
            def _enum_cb(hwnd, lparam):
                if not (user32.IsWindowVisible(hwnd) or user32.IsIconic(hwnd)):
                    return True
                title = get_window_title(hwnd).strip()
                if not title:
                    return True
                cls = get_class_name(hwnd).strip()
                t_low = title.lower()
                c_low = cls.lower()
                # Never match browsers, our own tool, or consoles!
                if any(b in t_low for b in ("chrome", "edge", "firefox", "brave", "opera", "forum studio", "google sheets", "launcher", "command prompt", "powershell")):
                    return True
                if any(b in c_low for b in ("tktoplevel", "consolewindowclass", "cascadia_hosting_window_class", "mozilla")):
                    return True
                if "chrome_widgetwin" in c_low and "zoom workplace" not in t_low and t_low != "zoom":
                    return True

                if "zoom workplace" in t_low or t_low == "zoom":
                    matches.append((100, hwnd, title, cls))
                elif "team chat" in t_low or "meeting chat" in t_low or "zoom meeting" in t_low:
                    matches.append((90, hwnd, title, cls))
                elif "zoom" in t_low:
                    matches.append((75, hwnd, title, cls))
                return True
            user32.EnumWindows(WNDENUMPROC(_enum_cb), 0)
            matches.sort(key=lambda x: x[0], reverse=True)
            return matches

        def minimize_all_zoom_windows():
            """Reliably minimizes Zoom Workplace and all visible Zoom windows back to the taskbar."""
            WM_SYSCOMMAND = 0x0112
            SC_MINIMIZE = 0xF020
            SW_MINIMIZE = 6
            GA_ROOT = 2

            targets = set()
            for _, h, _, _ in find_zoom_windows():
                targets.add(h)
                root_h = user32.GetAncestor(h, GA_ROOT)
                if root_h:
                    targets.add(root_h)

            fg_h = user32.GetForegroundWindow()
            if fg_h and "zoom" in get_window_title(fg_h).lower():
                targets.add(fg_h)
                root_fg = user32.GetAncestor(fg_h, GA_ROOT)
                if root_fg:
                    targets.add(root_fg)

            for h in targets:
                try:
                    user32.PostMessageW(h, WM_SYSCOMMAND, SC_MINIMIZE, 0)
                    user32.ShowWindowAsync(h, SW_MINIMIZE)
                    user32.ShowWindow(h, SW_MINIMIZE)
                    user32.CloseWindow(h)
                except Exception:
                    pass
            time.sleep(0.15)
            # Second pass if any Zoom window is still not iconic
            for h in targets:
                try:
                    if user32.IsWindowVisible(h) and not user32.IsIconic(h):
                        user32.PostMessageW(h, WM_SYSCOMMAND, SC_MINIMIZE, 0)
                        user32.ShowWindow(h, SW_MINIMIZE)
                except Exception:
                    pass

        zoom_list = find_zoom_windows()
        if not zoom_list:
            zoom_exes = [
                os.path.expandvars(r"%APPDATA%\Zoom\bin\Zoom.exe"),
                os.path.expandvars(r"%LOCALAPPDATA%\Zoom\bin\Zoom.exe"),
                os.path.expandvars(r"%ProgramFiles%\Zoom\bin\Zoom.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Zoom\bin\Zoom.exe"),
            ]
            for exe in zoom_exes:
                if os.path.exists(exe):
                    try:
                        subprocess.Popen([exe], creationflags=0x08000000)
                        break
                    except Exception:
                        pass
            for _ in range(10):
                time.sleep(0.4)
                zoom_list = find_zoom_windows()
                if zoom_list:
                    break

        if not zoom_list:
            return False, "Zoom Desktop App ('Zoom Workplace') is not open! Please open Zoom Chat once."

        zoom_hwnd = zoom_list[0][1]

        # 1. Put ONLY this single URL into the clipboard before touching focus
        clean_u = str(single_url or "").strip().splitlines()[0].strip()
        self._win32_set_clipboard_single_url(clean_u)

        # 2. Restore & bring Zoom Workplace to the foreground
        WM_SYSCOMMAND = 0x0112
        SC_RESTORE = 0xF120
        for _ in range(4):
            if user32.IsIconic(zoom_hwnd):
                user32.PostMessageW(zoom_hwnd, WM_SYSCOMMAND, SC_RESTORE, 0)
                user32.ShowWindow(zoom_hwnd, 9)  # SW_RESTORE
            else:
                user32.ShowWindow(zoom_hwnd, 5)  # SW_SHOW

            fg_cur = user32.GetForegroundWindow()
            fg_tid = user32.GetWindowThreadProcessId(fg_cur, None) if fg_cur else 0
            cur_tid = kernel32.GetCurrentThreadId()
            if fg_tid and cur_tid and fg_tid != cur_tid:
                user32.AttachThreadInput(cur_tid, fg_tid, True)
                user32.SetForegroundWindow(zoom_hwnd)
                user32.BringWindowToTop(zoom_hwnd)
                user32.SetFocus(zoom_hwnd)
                user32.AttachThreadInput(cur_tid, fg_tid, False)
            else:
                user32.keybd_event(0x12, 0, 0, 0)
                user32.SetForegroundWindow(zoom_hwnd)
                user32.BringWindowToTop(zoom_hwnd)
                user32.keybd_event(0x12, 0, 2, 0)

            time.sleep(0.12)
            fg_now = user32.GetForegroundWindow()
            fg_t = get_window_title(fg_now).lower()
            if fg_now == zoom_hwnd or "zoom" in fg_t:
                break

        # 3. Strict Foreground Guard: NEVER press Ctrl+V if Chrome/Sheets/Forum Studio Pro is in front!
        fg_check = user32.GetForegroundWindow()
        fg_title_now = get_window_title(fg_check).lower()
        fg_cls_now = get_class_name(fg_check).lower()
        if any(b in fg_title_now for b in ("google sheets", "forum studio", "chrome", "edge")) or (
            "chrome_widgetwin" in fg_cls_now and "zoom" not in fg_title_now
        ):
            minimize_all_zoom_windows()
            return False, "Blocked paste to protect your Sheet because Zoom did not take foreground focus."

        # 4. Native Win32 Ctrl+V + Enter (zero PowerShell, zero CMD flash!)
        VK_CONTROL = 0x11
        VK_V = 0x56
        VK_RETURN = 0x0D
        KEYEVENTF_KEYUP = 0x0002

        # Ensure modifier keys are released first
        user32.keybd_event(0x12, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(0x10, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)

        # Press Ctrl+V (pastes the 1 single URL)
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_V, 0, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)

        # Wait 0.14s for Zoom input box to render the pasted link, then press Enter
        time.sleep(0.14)
        user32.keybd_event(VK_RETURN, 0, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)

        # 5. Wait 0.22s for Zoom to send the chat message, then IMMEDIATELY minimize Zoom off the screen!
        time.sleep(0.22)
        minimize_all_zoom_windows()
        return True, "Sent & minimized Zoom"

    def run_zoom_sender_action(self, do_zoom=True, do_sheet=False):
        """
        Executes Tab 3 actions with a strict single-run lock (never duplicates threads!)
        and strict per-link order:
        For EACH link #1..#N one by one:
          1. FIRST puts that link into Today's Google Sheet column (if do_sheet=True).
          2. SECOND opens Zoom, sends ONLY that 1 link, and IMMEDIATELY minimizes Zoom (if do_zoom=True).
          3. THIRD waits the Delay / Link (s) timer (while Zoom is minimized!) before the next link.
        """
        if getattr(self, "is_zoom_sending", False):
            self.log_console("⏳ Zoom / Sheet action is already running — wait for it to finish or click 🛑 STOP.")
            return

        self.save_all_vaults_to_disk(True)
        self.txt_zoom_links.config(state="normal")
        raw = self.txt_zoom_links.get("1.0", "end")
        urls = re.findall(r'https?://(?:(?!https?://)[^\s"\'<>])+', raw)
        if not urls:
            messagebox.showwarning("No Links", "Load or paste links in the box first.")
            return

        # Clean glued URLs and lock the text box during execution
        self.txt_zoom_links.delete("1.0", "end")
        self.txt_zoom_links.insert("1.0", "\n".join(urls))
        self.txt_zoom_links.config(state="disabled")
        self.root.focus_set()

        req_col = (self.entry_zoom_save_col.get() or "AUTO").strip().upper()
        self.stop_requested = False
        self.is_zoom_sending = True

        try:
            cd_sec = max(0, int(self.spin_zoom_countdown.get()))
        except Exception:
            cd_sec = 5
        try:
            msg_dly = max(0, int(self.spin_zoom_delay.get()))
        except Exception:
            msg_dly = 2

        def _worker():
            try:
                cur_col = "J"
                filled_cnt = 0
                max_rows_per_col = 20
                if do_sheet:
                    self.root.after(0, lambda: self.lbl_zoom_timer_banner.config(
                        text="🔍 Checking Today's Sheet Column...", bg="#1e3a8a", fg="#93c5fd"
                    ))
                    cur_col, filled_cnt, max_rows_per_col = self._resolve_starting_sheet_slot(req_col)

                # Initial countdown before starting
                if do_zoom and cd_sec > 0:
                    for rem in range(cd_sec, 0, -1):
                        if self.stop_requested:
                            return
                        step_desc = "Putting in Sheet -> Sending to Zoom" if do_sheet else "Sending to Zoom"
                        self.root.after(0, lambda r=rem: self.lbl_zoom_timer_banner.config(
                            text=f"⏳ Starting in {r}s...", bg="#78350f", fg="#fde68a"
                        ))
                        self.root.after(0, lambda r=rem, sd=step_desc: self.lbl_zoom_msg.config(
                            text=f"⏳ {sd} starts in {r}s (processes 1 link at a time & minimizes Zoom)..."
                        ))
                        time.sleep(1.0)

                sheet_ok = 0
                zoom_ok = 0
                last_written_col = cur_col
                last_written_row = 2 + filled_cnt

                # Process strictly 1 LINK AT A TIME:
                # Step 1: Put Link #idx in Sheet FIRST
                # Step 2: Open Zoom, send ONLY Link #idx, and IMMEDIATELY minimize Zoom
                # Step 3: Wait Delay / Link (s) while Zoom is minimized
                for idx, u in enumerate(urls, start=1):
                    if self.stop_requested:
                        break

                    # --- STEP 1: PUT LINK #idx IN GOOGLE SHEET FIRST ---
                    if do_sheet and not self.stop_requested:
                        if filled_cnt >= max_rows_per_col:
                            cur_col = self._num_to_col(self._col_to_num(cur_col) + 1)
                            filled_cnt = 0
                        next_item_num = filled_cnt + 1  # 1..20 -> Row 3..22
                        exact_sheet_row = 2 + next_item_num
                        self.root.after(0, lambda i=idx, c=cur_col, r=exact_sheet_row: self.lbl_zoom_timer_banner.config(
                            text=f"📄 [{i}/{len(urls)}] Putting in Sheet (Col {c}, Row {r})...", bg="#1e3a8a", fg="#93c5fd"
                        ))
                        self.root.after(0, lambda i=idx, c=cur_col, r=exact_sheet_row: self.lbl_zoom_msg.config(
                            text=f"📄 Step 1: Saving Link #{i}/{len(urls)} to Main Sheet Column {c}, Row {r}..."
                        ))
                        res_col = self.save_link_to_sheet_sync(next_item_num, f"Zoom Link #{next_item_num}", u, cur_col, is_auto_mode=(req_col in ("AUTO", "")))
                        if res_col:
                            sheet_ok += 1
                            filled_cnt += 1
                            last_written_col = cur_col
                            last_written_row = exact_sheet_row
                            self.add_post_to_history("Zoom Sender", f"Zoom Link (Row {exact_sheet_row})", u, cur_col, exact_sheet_row)

                    # --- STEP 2: SEND ONLY LINK #idx IN ZOOM SECOND & IMMEDIATELY MINIMIZE ZOOM ---
                    if do_zoom and not self.stop_requested:
                        self.root.after(0, lambda i=idx: self.lbl_zoom_timer_banner.config(
                            text=f"💬 [{i}/{len(urls)}] Sending to Zoom & Minimizing...", bg="#4f46e5", fg="white"
                        ))
                        self.root.after(0, lambda i=idx: self.lbl_zoom_msg.config(
                            text=f"💬 Step 2: Opening Zoom, sending Link #{i}/{len(urls)}, and minimizing Zoom..."
                        ))
                        ok_z, msg_z = self._win32_send_single_link_to_zoom(u)
                        if not ok_z:
                            self.root.after(0, lambda: self.lbl_zoom_timer_banner.config(text="⚠️ Zoom Focus Error", bg="#e11d48", fg="white"))
                            self.root.after(0, lambda m=msg_z: self.lbl_zoom_msg.config(text=f"⚠️ {m}"))
                            self.root.after(0, lambda m=msg_z: messagebox.showwarning("Zoom Auto-Sender", m))
                            break
                        zoom_ok += 1

                    # --- STEP 3: WAIT PER-LINK DELAY WHILE ZOOM IS MINIMIZED ---
                    if idx < len(urls) and not self.stop_requested and (do_zoom or do_sheet):
                        wait_s = msg_dly if do_zoom else 0
                        for s_rem in range(wait_s, 0, -1):
                            if self.stop_requested:
                                break
                            self.root.after(0, lambda i=idx+1, sr=s_rem: self.lbl_zoom_timer_banner.config(
                                text=f"⏳ Next Link (#{i}/{len(urls)}) in {sr}s...", bg="#064e3b", fg="#6ee7b7"
                            ))
                            self.root.after(0, lambda i=idx, sr=s_rem: self.lbl_zoom_msg.config(
                                text=f"✅ Link #{i} done (Zoom minimized). Waiting {sr}s before Link #{i+1}..."
                            ))
                            time.sleep(1.0)

                summary_parts = []
                if do_sheet:
                    summary_parts.append(f"Put {sheet_ok}/{len(urls)} in Sheet Col {last_written_col} (up to Row {last_written_row})")
                if do_zoom:
                    summary_parts.append(f"Sent {zoom_ok}/{len(urls)} to Zoom (Zoom minimized)")
                final_txt = "✅ " + " -> ".join(summary_parts) if summary_parts else "✅ Done!"
                self.root.after(0, lambda: self.lbl_zoom_timer_banner.config(text=final_txt, bg="#064e3b", fg="#34d399"))
                self.root.after(0, lambda: self.lbl_zoom_msg.config(text=final_txt))
            finally:
                self.is_zoom_sending = False
                self.root.after(0, lambda: self.txt_zoom_links.config(state="normal"))

        threading.Thread(target=_worker, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = ForumStudioDesktopApp(root)
    root.mainloop()
