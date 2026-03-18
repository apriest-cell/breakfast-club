# Breakfast Club — Project Reference

## Functional Summary

Student sign-in/out kiosk for a school breakfast club. Staff search for a student by name and tap their card to toggle between IN and OUT states. Attendance is logged to Google Sheets. An admin panel (PIN-protected) shows who is currently signed in and allows retry of failed writes. Visit totals are maintained per student.

---

## Architecture

### Frontend (`Index.html`, 715 lines)
- Single-file HTML + CSS + vanilla JS, served via `HtmlService`
- No framework; state held in:
  - `allStudents[]` — full roster
  - `statusById{}` — current IN/OUT state per student
  - `updatingIds` (Set) — students with in-flight RPC calls
  - `failedOps[]` — queue of writes that errored, available for retry
- Optimistic UI: status flips immediately on tap, then confirmed/rolled back by RPC callback

### Backend (`Code.js`, 260 lines)
Exposed functions (called via `google.script.run`):

| Function | Purpose |
|---|---|
| `doGet` | Serves the HTML page |
| `getInitialData` | Returns full student roster + today's status map |
| `logSignAction` | Appends a sign event to the log sheet, updates visit count |
| `validateAdminPin` | Checks PIN string, returns boolean |
| `archiveOldLogs` | Moves rows older than N days to an archive sheet |

### Storage (Google Sheets)
- **`Students`** sheet — roster with visit count column
- **`BreakfastLog`** sheet — timestamped IN/OUT events
- **`BreakfastLog_Archive`** sheet (optional) — archived old rows

### Concurrency
`LockService.getScriptLock()` in `logSignAction` — waits up to 20 s, always released in `finally`.

### Polling
15-second client interval; skipped if `updatingIds` is non-empty or within 5 s of the last sign action.

---

## Identified Issues

### Security (Critical)

1. **Hardcoded PIN** (`Code.js:4`)
   `const ADMIN_PIN = '1522'` is visible to anyone with script editor access.
   **Fix:** `PropertiesService.getScriptProperties().getProperty('ADMIN_PIN')`

2. **No rate-limiting on `validateAdminPin`**
   A caller can loop the RPC and brute-force a 4-digit PIN trivially.
   **Fix:** Track failed attempts + timestamp in `PropertiesService` or `CacheService`; enforce exponential back-off or a lockout period.

3. **No server-side auth on `logSignAction`**
   Any authenticated Google user can call it with arbitrary `studentId`/`newStatus`.
   **Fix:** For internal-only deployments, restrict to domain. For public deployments, add a per-session CSRF token or shared secret.

4. **`XFrameOptionsMode.ALLOWALL`** (`Code.js:32`)
   Exposes the app to clickjacking.
   **Fix:** Change to `SAMEORIGIN` unless cross-origin embedding is explicitly required.

---

### Data Integrity

5. **`getTodayStatusMap_` early-break assumption** (`Code.js:115`)
   Breaks the log scan the moment *any* row has a date before today, assuming strict chronological order. Manual edits or partial archiving can violate this, causing stale status reads.
   **Fix:** Remove the early break, or sort the lookback range defensively before scanning.

6. **`LOG_LOOKBACK_ROWS = 800`** (`Code.js:12`)
   On a very busy day (500+ students, multiple sign-ins/outs) earlier records may fall outside the lookback window.
   **Fix:** Increase to 1500, or scan from a date boundary rather than a fixed row count.

7. **`incrementVisitCount_` full-sheet scan** (`Code.js:209–218`)
   Reads the entire Students sheet on every sign-in to locate one row — O(n) per action.
   **Fix:** Read only column 1 to find the row index, then update the specific cell.

8. **`failedOps` is in-memory only**
   A page refresh permanently drops any queued retries, potentially losing attendance records.
   **Fix:** Persist `failedOps` to `localStorage` and restore on page load.

---

### Logic Bugs

9. **`getLastActionToday_` date short-circuit** (`Code.js:145–152`)
   The function never short-circuits on a date boundary unless the current student's row appears first. For students with no activity today, it scans the full lookback window unnecessarily. Low impact but wasteful.

10. **`archiveOldLogs` opens its own spreadsheet** (`Code.js:229`)
    Re-opens the spreadsheet independently instead of using `getSheets_()`, wasting a quota call. Minor inconsistency.

---

### UX

11. **Admin panel status is stale**
    `renderAdminList()` reads in-memory `statusById`, not a fresh server query. Sign-ins from other devices won't appear until the admin refreshes.
    **Fix:** Call `loadData()` when opening the admin panel.

12. **No end-of-day sign-out / daily report**
    Students still marked IN at close are never automatically signed out; the log accumulates open sessions indefinitely.
    **Fix:** Add a time-driven Apps Script trigger (e.g. 09:30) to sign out all remaining IN students and optionally email a summary.

13. **Search fires at 1 character**
    In a large school, typing "A" matches many students and produces a noisy list.
    **Fix:** Raise the threshold to 2 characters.

14. **Phone warning always visible**
    The "please use a tablet" banner is shown to all users including staff at a fixed kiosk.
    **Fix:** Make it collapsible, or suppress it after a user dismissal stored in `localStorage`.

---

### Minor / Code Quality

15. **Spreadsheet ID hardcoded** (`Code.js:2`)
    Should move to `PropertiesService` alongside the PIN for easier config management without touching code.

16. **`getSheets_()` caching comment overstated**
    Apps Script already caches `SpreadsheetApp.openById` within an execution; the comment implies performance savings that don't exist in a meaningful way. Remove or correct the comment.

---

## Prioritised Improvements

| Priority | Issue | Effort |
|---|---|---|
| P0 | Move PIN to PropertiesService (#1) | Low |
| P0 | Move Spreadsheet ID to PropertiesService (#15) | Low |
| P0 | Fix `XFrameOptionsMode` to `SAMEORIGIN` (#4) | Low |
| P1 | Add PIN brute-force rate limiting (#2) | Medium |
| P1 | Persist `failedOps` to `localStorage` (#8) | Medium |
| P1 | Fix `getTodayStatusMap_` early-break (#5) | Low |
| P1 | Increase `LOG_LOOKBACK_ROWS` (#6) | Low |
| P2 | Refresh data on admin panel open (#11) | Low |
| P2 | End-of-day sign-out trigger (#12) | Medium |
| P2 | Optimise `incrementVisitCount_` (#7) | Low |
| P3 | Raise search threshold to 2 chars (#13) | Low |
| P3 | Collapsible/dismissible phone warning (#14) | Low |
| P3 | Fix `archiveOldLogs` to use `getSheets_()` (#10) | Low |
| P3 | Fix `getLastActionToday_` scan efficiency (#9) | Low |
| P3 | Correct `getSheets_()` caching comment (#16) | Low |

---

## Quick Wins (low-risk, ship immediately)

1. **PropertiesService for PIN + Spreadsheet ID** — two-line changes, eliminates two hardcoded secrets.
2. **`XFrameOptionsMode.SAMEORIGIN`** — one-word change, eliminates clickjacking surface.
3. **`LOG_LOOKBACK_ROWS` → 1500** — one-number change, improves correctness on busy days.
4. **Remove early break in `getTodayStatusMap_`** — three-line change, prevents stale status from manual sheet edits.
5. **`localStorage` persistence for `failedOps`** — ~10 lines of JS, prevents silent attendance data loss on refresh.
