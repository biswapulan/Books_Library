/**
 * BPUTNotes Apps Script — Backend for site + intern book-assignment panel
 * Paste this ENTIRE file into script.google.com, replacing what's there now,
 * then Deploy → Manage Deployments → Edit → New Version.
 *
 * Tabs this script touches:
 *   Notes, PYQs, Books, Scholarships, Popup   ← existing site data (unchanged)
 *   Users                                      ← intern login + row ranges
 *   Submissions                                ← intern book-submission sheet
 *
 * IMPORTANT: Do NOT include the Users or Submissions tabs when you do
 * File → Share → Publish to web. This script reads/writes them directly
 * with full permissions regardless of publish settings, so they never need
 * to go through the public read API — keeping passwords and raw
 * submissions off the public API key.
 */

const SHEET_ID = '17o1aktSV1oM_EpWkSPkx7kVhGw4ZJChi1BuhMIl3jnk'; // main site sheet: Notes/PYQs/Books/Scholarships/Popup

// Your SEPARATE sheet that has the Users + Submissions tabs.
// Get this ID from its URL: docs.google.com/spreadsheets/d/<THIS PART>/edit
const SUBMISSIONS_SHEET_ID = '13E9vnsnErH8r-GM8JvmxMAOoir-Ys56f_s8g_bin7II';

const USERS_TAB       = 'Users';
const SUBMISSIONS_TAB = 'Submissions';
const BOOKS_TAB        = 'Books';

function getSubmissionsSS() {
  return SpreadsheetApp.openById(SUBMISSIONS_SHEET_ID);
}

// Folder in YOUR Drive where assigned book files get copied to, so links
// never break even if the original intern's file gets deleted/unshared later.
// Get this ID from the folder's URL: drive.google.com/drive/folders/<THIS PART>
const BOOK_LIBRARY_FOLDER_ID = '1Q7GpTSXzPcV9CZBihdzIXQtjNtwENmrl';

// Same grouping already used in the admin panel's fan-out feature — kept in sync on purpose.
const BRANCH_GROUPS = {
  circuit:       ['electrical', 'civil', 'mining', 'mechanical'],
  'non-circuit': ['cse', 'mineral', 'metallurgy'],
};

const SESSION_TTL_SECONDS = 21600; // 6 hours (CacheService max) — interns re-login after that

// Skip copying (keep original link + warn) if a file is bigger than this.
// Note: this cap applies cleanly to Drive-to-Drive copies (no size limit there
// besides your Drive storage quota). For files downloaded from an external
// URL, Google's UrlFetchApp itself caps responses at roughly 50MB — a file
// bigger than that will fail the fetch (caught below) before this check ever
// runs, regardless of what MAX_COPY_BYTES is set to.
const MAX_COPY_BYTES = 100 * 1024 * 1024; // 100MB

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SHEET_ID);

    // ── Intern assignment-panel actions ──
    if (data.action === 'login')          return handleLogin(data);
    if (data.action === 'getSubmissions') return handleGetSubmissions(data);
    if (data.action === 'assignBook')     return handleAssignBook(ss, data); // ss = main sheet, for the Books tab

    // ── Existing admin-panel actions (unchanged) ──
    const sheet = ss.getSheetByName(data.tab);
    if (!sheet) throw new Error('Tab not found: ' + data.tab);

    if (data.action === 'append') {
      sheet.appendRow(data.row);
    } else if (data.action === 'update') {
      sheet.getRange(data.rowIndex, 1, 1, data.row.length).setValues([data.row]);
    } else if (data.action === 'delete') {
      sheet.deleteRow(data.rowIndex);
    } else if (data.action === 'setPopupKey') {
      const vals = sheet.getDataRange().getValues();
      for (let i = 0; i < vals.length; i++) {
        if (vals[i][0] === data.key) {
          sheet.getRange(i + 1, 2).setValue(data.value);
          break;
        }
      }
    } else if (data.action === 'replacePopup') {
      sheet.clearContents();
      sheet.getRange(1, 1, data.rows.length, 2).setValues(data.rows);
    } else {
      throw new Error('Unknown action: ' + data.action);
    }

    return jsonOut({ success: true });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

function doGet() {
  return ContentService
    .createTextOutput('BPUTNotes Apps Script OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ════════════════ AUTH ════════════════ */

function handleLogin(data) {
  const sheet = getSubmissionsSS().getSheetByName(USERS_TAB);
  if (!sheet) return jsonOut({ success: false, error: 'Users tab not found' });

  const rows = sheet.getDataRange().getValues();
  const username = String(data.username || '').trim();
  const password = String(data.password || '');

  for (let i = 1; i < rows.length; i++) { // skip header row
    const [u, p, name, rowStart, rowEnd, active] = rows[i];
    if (String(u).trim() === username && String(p) === password) {
      if (String(active).toUpperCase() === 'FALSE') {
        return jsonOut({ success: false, error: 'This account has been disabled' });
      }
      const session = {
        username,
        name: name || username,
        rowStart: Number(rowStart) || 2,
        rowEnd: Number(rowEnd) || 2,
      };
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put('sess_' + token, JSON.stringify(session), SESSION_TTL_SECONDS);
      return jsonOut({
        success: true, token,
        username: session.username,
        name: session.name, rowStart: session.rowStart, rowEnd: session.rowEnd,
      });
    }
  }
  return jsonOut({ success: false, error: 'Invalid username or password' });
}

function getSession(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('sess_' + token);
  return raw ? JSON.parse(raw) : null;
}

/* ════════════════ FETCH ASSIGNABLE ROWS (server-side range enforced) ════════════════ */

function handleGetSubmissions(data) {
  const session = getSession(data.token);
  if (!session) return jsonOut({ success: false, error: 'Session expired. Please log in again.' });

  const sheet = getSubmissionsSS().getSheetByName(SUBMISSIONS_TAB);
  if (!sheet) return jsonOut({ success: false, error: 'Submissions tab not found' });

  const lastRow = sheet.getLastRow();
  // rowStart/rowEnd come ONLY from the server-stored session — never from the client.
  const start = Math.max(2, session.rowStart);
  const end   = Math.min(lastRow, session.rowEnd);
  if (end < start) return jsonOut({ success: true, rows: [] });

  const vals = sheet.getRange(start, 1, end - start + 1, 12).getValues(); // A:L (L pads blank if missing)
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    const assigned = r[9]; // column J
    if (String(assigned).toUpperCase() === 'TRUE') continue; // already assigned — skip
    out.push({
      rowIndex:    start + i,
      timestamp:   r[0], internName: r[1], department: r[2], semester: r[3],
      subject:     r[4], bookTitle:  r[5], bookLink:   r[6], sourceType: r[7], note: r[8],
    });
  }
  return jsonOut({ success: true, rows: out });
}

/* ════════════════ COPY BOOK FILE INTO YOUR OWN DRIVE ════════════════ */

// Pulls a Drive file ID out of the common URL shapes people paste.
function extractDriveFileId(url) {
  if (!url) return null;
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]{20,})/,   // .../file/d/ID/view
    /[?&]id=([a-zA-Z0-9_-]{20,})/,        // ...?id=ID
    /\/d\/([a-zA-Z0-9_-]{20,})/,          // ...docs.google.com/.../d/ID
  ];
  for (const re of patterns) {
    const m = url.match(re);
    if (m) return m[1];
  }
  return null;
}

// Tries to save a durable copy into BOOK_LIBRARY_FOLDER_ID and return a stable
// link. Handles two cases: (1) a Google Drive file link — copied directly via
// DriveApp, no download needed; (2) any other direct link — downloaded and
// inspected; if it's actually a PDF (by content-type OR file signature, not
// just URL extension), it's re-uploaded. Anything else (an HTML page, a
// login-walled host, a folder link, a file that's too large, etc.) falls
// back to the original link with a warning.
function tryCopyToOwnDrive(originalUrl, suggestedName) {
  if (!originalUrl || originalUrl === '#') return { link: originalUrl, warning: null };
  if (!BOOK_LIBRARY_FOLDER_ID || BOOK_LIBRARY_FOLDER_ID === 'PUT_YOUR_FOLDER_ID_HERE') {
    return { link: originalUrl, warning: 'BOOK_LIBRARY_FOLDER_ID not set in Code.gs — original link kept.' };
  }
  const folder = DriveApp.getFolderById(BOOK_LIBRARY_FOLDER_ID);
  const safeName = (suggestedName ? String(suggestedName).replace(/[\\/:*?"<>|]/g, '').trim() : '') || 'book';

  // ── Case 1: Google Drive link — copy directly, no download needed ──
  if (/drive\.google\.com|docs\.google\.com/.test(originalUrl)) {
    const fileId = extractDriveFileId(originalUrl);
    if (!fileId) {
      return { link: originalUrl, warning: 'Looked like a Drive folder link, not a single file — kept original link.' };
    }
    try {
      const original = DriveApp.getFileById(fileId);

      // Guard against huge files here too, before copying.
      const size = original.getSize();
      if (size && size > MAX_COPY_BYTES) {
        return {
          link: originalUrl,
          warning: 'File too large to copy (' + Math.round(size / (1024 * 1024)) + 'MB) — original link kept.'
        };
      }

      const copy = original.makeCopy(safeName, folder);
      copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return { link: copy.getUrl(), warning: null };
    } catch (err) {
      return { link: originalUrl, warning: 'Could not copy from Drive (' + err.message + ') — original link kept, check its sharing settings.' };
    }
  }

  // ── Case 2: any other direct link — fetch it and inspect what actually comes back ──
  try {
    const resp = UrlFetchApp.fetch(originalUrl, { muteHttpExceptions: true, followRedirects: true });
    const code = resp.getResponseCode();
    if (code !== 200) {
      return { link: originalUrl, warning: 'Could not download that link (HTTP ' + code + ') — original link kept.' };
    }

    const blob = resp.getBlob();
    const headers = resp.getAllHeaders();
    // Header names can come back in different cases depending on the server.
    const contentTypeHeader = String(
      headers['Content-Type'] || headers['content-type'] || ''
    ).toLowerCase();

    const bytes = blob.getBytes();

    if (bytes.length > MAX_COPY_BYTES) {
      return {
        link: originalUrl,
        warning: 'File too large to copy (' + Math.round(bytes.length / (1024 * 1024)) + 'MB) — original link kept.'
      };
    }

    // Check the actual file signature too ("%PDF"), since some hosts serve
    // PDFs with a wrong or generic content-type (e.g. application/octet-stream).
    const first4 = bytes.slice(0, 4).map(b => String.fromCharCode((b + 256) % 256)).join('');
    const isPdfBySignature = first4 === '%PDF';
    const isPdfByContentType = contentTypeHeader.indexOf('pdf') !== -1;

    if (!isPdfBySignature && !isPdfByContentType) {
      return {
        link: originalUrl,
        warning: 'That link did not return a PDF (got "' + (contentTypeHeader || 'unknown content-type') + '") — kept original link, not copied.'
      };
    }

    blob.setName(safeName + '.pdf');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { link: file.getUrl(), warning: null };
  } catch (err) {
    return { link: originalUrl, warning: 'Could not download that link (' + err.message + ') — original link kept.' };
  }
}

/* ════════════════ ASSIGN A BOOK (writes to Books tab, with sem1/2 fan-out) ════════════════ */

function handleAssignBook(ss, data) {
  const session = getSession(data.token);
  if (!session) return jsonOut({ success: false, error: 'Session expired. Please log in again.' });

  const rowIndex = Number(data.sourceRow);
  if (!rowIndex || rowIndex < session.rowStart || rowIndex > session.rowEnd) {
    return jsonOut({ success: false, error: 'That row is outside your assigned range' });
  }

  const subSheet = getSubmissionsSS().getSheetByName(SUBMISSIONS_TAB);
  if (!subSheet) return jsonOut({ success: false, error: 'Submissions tab not found' });

  const rowVals = subSheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
  if (String(rowVals[9]).toUpperCase() === 'TRUE') {
    return jsonOut({ success: false, error: 'This entry was already assigned by someone else' });
  }

  const semester      = Number(data.semester) || 1;
  const subjectNumber = Number(data.subjectNumber) || 1;
  const subjectName   = String(data.subjectName || '').trim();
  const bookName       = String(data.bookName || '').trim();
  const tags             = String(data.tags || 'Book').trim();
  const status            = data.status || 'active';

  if (!subjectName || !bookName) {
    return jsonOut({ success: false, error: 'Subject name and book name are required' });
  }

  // Copy the file into your own Drive so the link stays alive long-term.
  const originalLink = String(data.driveLink || '').trim() || '#';
  const copyResult = tryCopyToOwnDrive(originalLink, bookName);
  const driveLink = copyResult.link;

  let branches = [];
  if ((semester === 1 || semester === 2) && data.group && BRANCH_GROUPS[data.group]) {
    branches = BRANCH_GROUPS[data.group];
  } else if (data.branch) {
    branches = [String(data.branch)];
  } else {
    return jsonOut({ success: false, error: 'Pick a branch or a group' });
  }

  const booksSheet = ss.getSheetByName(BOOKS_TAB);
  if (!booksSheet) return jsonOut({ success: false, error: 'Books tab not found' });

  branches.forEach(function (b) {
    booksSheet.appendRow([b, semester, subjectNumber, subjectName, bookName, driveLink, tags, status]);
  });

  // Mark the submission row as used so it can't be double-assigned.
  subSheet.getRange(rowIndex, 10).setValue(true);                          // J: Assigned
  subSheet.getRange(rowIndex, 11).setValue(session.name || session.username); // K: AssignedBy
  subSheet.getRange(rowIndex, 12).setValue(new Date());                    // L: AssignedAt

  return jsonOut({ success: true, branchesAssigned: branches, driveWarning: copyResult.warning });
}
