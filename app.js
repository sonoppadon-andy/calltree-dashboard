(() => {
  "use strict";
  const cfg = window.CALLTREE_CONFIG;
  let msalApp;
  let allRows = [];
  let allPhoneBook = []; // Master employee list from the "Phone Book" SharePoint site (see config.js)
  let statusChart;
  let drillChart;
  let missingChart;
  // Cross-linking charts -> table: currentRows/currentBucketRowIds are (re)filled on every
  // render() and read by the chart click handlers (which only call renderTable(), never the
  // full render(), so the charts themselves aren't destroyed/rebuilt on every click).
  let currentRows = [];
  let currentBucketRowIds = {};
  // {key, label, predicate(row)=>bool} for a filter over currentRows (most slices/bars), OR
  // {key, label, rows:[...]} for a filter that supplies its own row set instead of predicating
  // over currentRows (used by "Pending Employer", added 2026-09-24 — see getFilteredRows()), OR
  // null (no chart filter active).
  let activeFilter = null;
  // Latest Responses table: per-column text filters, pagination (30 rows/page).
  const PAGE_SIZE = 30;
  let currentPage = 1;
  let columnFilters = {id:"", refid:"", email:"", created:"", status:"", drill:""};
  // "รายงานผู้ที่ยังไม่ตอบรับ/ตอบรับ" tab (Phone Book master list): currentReportRows holds
  // EVERY master-list person (not just the not-yet-responded ones — widened 2026-09-24 to
  // support the ตอบ/ไม่ตอบ status filter/column below), each tagged with a `Responded`
  // boolean; currentPhoneBookTotal is the master headcount. Both are (re)filled on every
  // render(). Five filters combine (AND) to filter the table: missingBuFilter (BU dropdown,
  // added 2026-09-24 as the filter-bar section — cascades into Department's own options via
  // populateReportFilterSelects() and rescopes missingChart, both updated 2026-09-24),
  // missingDeptFilter (Department dropdown — driven by BOTH the #missingDeptFilterSelect
  // dropdown and clicking a missingChart bar as of 2026-09-24 when the chart switched from
  // BU-grouped to Department-grouped, kept in sync both ways), reportStatusFilter (ตอบ/ไม่ตอบ
  // dropdown), reportTypeFilter (Type dropdown, added 2026-09-24 — see classifyJobType()
  // below), and reportBucketFilter (จำนวนผู้ตอบตามช่วงเวลา 15 นาที dropdown, added 2026-09-24 —
  // matches each report row's `ResponseBucket`, see the ResponseTime/ResponseBucket computation
  // in render()). missingPage
  // (same PAGE_SIZE as the Member table) paginates the result. All of these only call
  // renderMissingTable() (never render()), same pattern as the Member table's
  // activeFilter/columnFilters/currentPage.
  let currentReportRows = [];
  let currentPhoneBookTotal = 0;
  let missingBuFilter = null;
  let missingDeptFilter = null;
  let reportStatusFilter = null; // null (all) | "responded" | "missing"
  let reportTypeFilter = null; // null (all) | "management" | "staff"
  let reportBucketFilter = null; // null (all) | bucket-start minute (number, e.g. 0, 15, 30...)
  let missingPage = 1;
  // Dashboard's OWN filter bar (added 2026-09-25) — same 5 dimensions as the Report tab's filter
  // bar above, but scopes the WHOLE Dashboard view (every KPI card, both charts, and the Latest
  // Responses table) instead of just one table, per the user's explicit request. Kept as entirely
  // separate state from the Report tab's missing*/report* variables above — the two filter bars
  // are independent (picking a BU here does not affect the Report tab's own BU filter or vice
  // versa). See render()'s "Dashboard filter bar" section for how these are applied (a join
  // against the Phone Book by email, since Member rows themselves don't carry BU/Department/Job
  // Title — only the Phone Book does).
  let dashBuFilter = null;
  let dashDeptFilter = null;
  let dashStatusFilter = null; // null (all) | "responded" | "missing"
  let dashTypeFilter = null; // null (all) | "management" | "staff"
  let dashBucketFilter = null; // null (all) | bucket-start minute (number, e.g. 0, 15, 30...)
  // "รายงานขอความช่วยเหลือ" tab (added 2026-09-27, revised 3rd pass same day) — own independent
  // filter bar (same 5 dimensions as Dashboard's). TWO base arrays, per the user's explicit
  // correction that the table must show every real SharePoint row (not collapsed) while the chart
  // counts must still collapse to one per (RefID, Email):
  //   - currentHelpRows: EVERY individual row (own `ID`) with HelpNote content — used by the
  //     table/CSV export, undeduped.
  //   - currentHelpUniqueRows: the same rows collapsed to one per (RefID, Email) — EARLIEST
  //     Created wins — used ONLY by renderHelpChart()/renderHelpReasonChart()'s counts.
  // Both tagged with BU/Department/JobTitle/Type (joined from the Phone Book) and Reason (see
  // classifyHelpReason(), built from real HelpNote sample text per "Never Guess"). Rebuilt every
  // render(); helpChart/helpPage mirror missingChart/missingPage's own pattern (own chart
  // instance, own pagination). helpReasonFilter (set only by clicking a bar of helpReasonChart,
  // not a filter-bar dropdown — same relationship as quizCorrectFilter below) narrows the table by
  // the derived Reason bucket.
  let currentHelpRows = [];
  let currentHelpUniqueRows = [];
  let helpBuFilter = null;
  let helpDeptFilter = null;
  let helpStatusFilter = null; // null (all) | "responded" | "missing"
  let helpTypeFilter = null; // null (all) | "management" | "staff"
  let helpBucketFilter = null; // null (all) | bucket-start minute
  let helpReasonFilter = null; // null (all) | "injury" | "trapped" | "lost" | "other"
  let helpPage = 1;
  let helpChart;
  let helpReasonChart;
  // Help tab now has 2 sub-tabs (added 2026-09-27, per user request): "ข้อมูลสรุปตาม KPI Dashboard"
  // (KPI cards + the 2 existing charts) and "ตารางข้อมูลทั้งหมด" (the full undeduped table, unchanged).
  // Both sub-tabs share the SAME filter-bar above them — helpSubTab only toggles which content is
  // visible, it never changes what's counted/filtered.
  let helpSubTab = "kpi"; // "kpi" | "table"
  // "รายงานผลการตอบคำถาม" tab (added 2026-09-27) — own independent filter bar (same 5
  // dimensions), own base row set. currentQuizRows holds ONE entry per (RefID, person) drill-
  // answer instance within the current RefID scope — same "first non-blank Drill Response wins"
  // rule as the existing "% ผู้ที่ตอบผิด" KPI — tagged with BU/Department/Type (joined from the
  // Phone Book) and whether that answer was Correct (matches the RefID's Admin iMsg text) or not.
  // quizCorrectFilter (set by clicking a pie slice) narrows the table to only-correct or
  // only-wrong answers; quizPage mirrors missingPage's own pagination pattern.
  let currentQuizRows = [];
  let quizBuFilter = null;
  let quizDeptFilter = null;
  let quizStatusFilter = null; // null (all) | "responded" | "missing"
  let quizTypeFilter = null; // null (all) | "management" | "staff"
  let quizBucketFilter = null; // null (all) | bucket-start minute
  let quizCorrectFilter = null; // null (all) | "correct" | "wrong"
  let quizPage = 1;
  let quizChart;
  const $ = id => document.getElementById(id);
  const show = (id, visible=true) => $(id).classList.toggle("hidden", !visible);
  const clean = value => String(value ?? "").trim();
  const escapeHtml = value => clean(value).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
  const isSafe = row => clean(row.ResponseSafe).toLowerCase() === "iamsafe";
  const hasResponse = row => Boolean(clean(row.ClickDateTime) || clean(row.ResponseSafe) || clean(row.DrillResponse));
  // SharePoint Choice/Lookup/Person fields can come back as objects instead of plain strings.
  // fieldText() extracts readable text from those shapes so "empty" checks (Mode/iMsg) work correctly.
  function fieldText(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "object") {
      if (Array.isArray(value)) return value.map(fieldText).filter(Boolean).join(", ");
      return clean(value.Value ?? value.LookupValue ?? value.Title ?? value.DisplayName ?? value.Label ?? value.email ?? "");
    }
    return clean(value);
  }
  // The Admin broadcast/announcement row carries the Mode + iMsg text for the whole RefID batch;
  // individual member response rows do not have Mode/iMsg populated. So "has Mode or iMsg" = notification row.
  const isNotificationRow = row => Boolean(fieldText(row.Mode) || fieldText(row.iMsg));
  // "Type" filter on the Phone Book report table: groups Job Title into Management (VP level
  // and above) vs. Staff, per the user's explicit rule ("Management คือ Job Title ที่เป็น VP
  // ขึ้นไป"). This is a best-effort KEYWORD match, not exact data — the Phone Book's actual Job
  // Title values weren't available to design this against directly (per "Never Guess", nothing
  // here was invented from assumed org-chart data), so it's built to be conservative and
  // explicit rather than guessy:
  //  - Matches "Vice President"/VP/AVP/SVP/EVP (word-boundary, so it won't match inside another
  //    word) and unambiguous C-suite/top titles (President, Chairman, CEO/CFO/COO/CTO, Managing
  //    Director) as Management. AVP (Assistant Vice President) was explicitly added to
  //    Management 2026-09-24 per the user's follow-up request — it was excluded when this
  //    function was first built (reasoned as one rank below full VP), but the user confirmed
  //    AVP should count as "VP ขึ้นไป" too.
  //  - Still explicitly EXCLUDES support/staff roles that happen to contain one of those words
  //    as a substring but are not themselves that rank — e.g. "CEO Driver" (a driver role, not
  //    the CEO), "Secretary to MD", "Personal Assistant to VP". Found via a real Phone Book
  //    screenshot in this conversation (RefID BU=HC&GA had a "CEO Driver" job title) — without
  //    this exclusion list a naive substring match would have wrongly classified that row as
  //    Management.
  // If any Job Title in the real data is still misclassified, tell me the exact title text and
  // which bucket it should be in — I'll add it to the keyword/exclude lists below rather than
  // guess further ones preemptively.
  function classifyJobType(jobTitle) {
    const t = clean(jobTitle).toLowerCase();
    if (!t) return null;
    const excludePatterns = [/driver/, /secretary/, /personal assistant/, /\bpa\b/, /assistant to/, /housekeeper/];
    if (excludePatterns.some(p => p.test(t))) return "staff";
    const managementPatterns = [
      /vice president/, /\bvp\b/, /\bavp\b/, /\bsvp\b/, /\bevp\b/,
      /\bpresident\b/, /\bchairman\b/,
      /chief executive officer/, /\bceo\b/,
      /chief financial officer/, /\bcfo\b/,
      /chief operating officer/, /\bcoo\b/,
      /chief technology officer/, /\bcto\b/,
      /chief human/, /\bchro\b/,
      /managing director/, /\bmd\b/,
    ];
    return managementPatterns.some(p => p.test(t)) ? "management" : "staff";
  }
  // "เหตุผล" classifier for the Help Request Report's reason-breakdown chart (added 2026-09-27),
  // built AFTER the user supplied real HelpNote sample text (per "Never Guess" — a keyword
  // classifier was deliberately withheld until then, see the notes doc). The real data showed
  // HelpNote is a multi-select checkbox field: its 3 fixed option strings — "บาดเจ็บ", "หลงทาง",
  // "ติดในอาคาร" — appear verbatim as comma-joined substrings, often followed by free text and/or
  // a ":โทร <phone> <name>" callback suffix the user appends themselves (e.g. "บาดเจ็บ, หลงทาง,
  // ติดในอาคาร อยู่ชั้น40:โทร 086-897-5806 จอย"). A single person can tick more than one option in
  // the same HelpNote.
  //
  // REVISED 2026-09-27 (2nd pass) per the user's explicit correction: each entry is classified by
  // WHICHEVER of the 3 keywords appears FIRST (leftmost) in the actual HelpNote text — NOT a fixed
  // severity ranking (the original "บาดเจ็บ > ติดในอาคาร > หลงทาง" priority order from the first
  // pass was wrong whenever หลงทาง was ticked before ติดในอาคาร in the text, e.g. "หลงทาง,
  // ติดในอาคาร ทดสอบ..." should classify as หลงทาง — the keyword that's actually first in the
  // sentence — not ติดในอาคาร). Ties can't occur (each keyword is checked by its own text
  // position; the earliest index wins). No keyword found at all (test/garbled entries like
  // "ทดสอบ...", "Test Show log...", "หลงๆลืมๆ" seen in the real sample data — the last one does
  // NOT literally contain the "หลงทาง" substring) → "อื่นๆ".
  function classifyHelpReason(helpNote) {
    const t = clean(helpNote);
    const candidates = [
      {key:"injury", idx:t.indexOf("บาดเจ็บ")},
      {key:"lost", idx:t.indexOf("หลงทาง")},
      {key:"trapped", idx:t.indexOf("ติดในอาคาร")},
    ].filter(c => c.idx !== -1);
    if (!candidates.length) return "other";
    candidates.sort((a,b) => a.idx - b.idx);
    return candidates[0].key;
  }
  const HELP_REASON_LABELS = {injury:"บาดเจ็บ", lost:"หลงทาง", trapped:"ติดในอาคาร", other:"อื่นๆ"};
  // Display order for the reason chart follows the order the user originally asked for
  // ("บาดเจ็บ,หลงทาง ติดในอาคาร อื่นๆ") — unrelated to classifyHelpReason()'s own "first keyword in
  // the text" logic above (that's about picking ONE bucket per entry, not about how the bars are
  // laid out).
  const HELP_REASON_ORDER = ["injury","lost","trapped","other"];

  function setMessage(message, kind="info") {
    const el=$("statusMessage"); el.textContent=message; el.className=`status ${kind}`; show("statusMessage", Boolean(message));
  }
  function validateConfig() {
    const invalid = !cfg || cfg.clientId.includes("REPLACE_") || cfg.redirectUri.includes("REPLACE_");
    if (invalid) {
      $("configMessage").textContent = "กรุณาแก้ไข clientId, GitHub username และ redirectUri ใน config.js ก่อนใช้งาน";
      $("loginButton").disabled = true;
    }
    return !invalid;
  }
  async function graphGet(url, token) {
    const endpoint = url.startsWith("https://") ? url : `https://graph.microsoft.com/v1.0${url}`;
    const response = await fetch(endpoint, {headers:{Authorization:`Bearer ${token}`,Accept:"application/json"}});
    if (!response.ok) throw new Error(`Microsoft Graph ${response.status}: ${await response.text()}`);
    return response.json();
  }
  async function acquireToken() {
    const account = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
    if (!account) throw new Error("ไม่พบบัญชีที่เข้าสู่ระบบ");
    try { return (await msalApp.acquireTokenSilent({account, scopes:cfg.graphScopes})).accessToken; }
    catch { return (await msalApp.acquireTokenPopup({account, scopes:cfg.graphScopes})).accessToken; }
  }
  async function readSharePointList(token) {
    const site = await graphGet(`/sites/${cfg.sharePointHost}:${cfg.sitePath}?$select=id,displayName`, token);
    const lists = await graphGet(`/sites/${site.id}/lists?$select=id,displayName`, token);
    const list = (lists.value || []).find(x => x.displayName === cfg.listName);
    if (!list) throw new Error(`ไม่พบ SharePoint List ชื่อ ${cfg.listName}`);
    // NOTE: the SharePoint column shown to users as "Mode" has internal (API) field name "Team" —
    // its Display Name was changed after the column was created, and SharePoint never updates the
    // internal name to match. Confirmed via /_layouts/15/FldEdit.aspx?...&Field=Team for that column.
    const fields = "Team,Member,EMail,ClickDateTime,Created,iMsg,ResponseSafe,HelpNote,RefID,DrillChoice,DrillResponse";
    let next = `/sites/${site.id}/lists/${list.id}/items?$expand=fields($select=${fields})&$select=id,fields&$top=999`;
    const rows=[];
    while (next) {
      const page=await graphGet(next, token);
      // Map Team -> Mode so the rest of the app can keep using row.Mode as before.
      rows.push(...(page.value || []).map(item => ({ID:item.id, ...item.fields, Mode:item.fields.Team})));
      next=page["@odata.nextLink"] || null;
    }
    return rows;
  }
  // Reads the "Phone Book" master list from the /sites/snet SharePoint site (separate site
  // from cfg.sitePath/"Member"). Same pagination pattern as readSharePointList(). Field names
  // come from cfg.phoneBook.fields — internal SharePoint names (field_15/field_16/field_14/
  // field_10/field_12), confirmed via each column's FldEdit.aspx URL — never guess these, see
  // config.js. Department (field_12) added 2026-09-24, table-only per the user's request.
  async function readPhoneBookList(token) {
    const pb = cfg.phoneBook;
    if (!pb) return [];
    const site = await graphGet(`/sites/${cfg.sharePointHost}:${pb.sitePath}?$select=id,displayName`, token);
    const lists = await graphGet(`/sites/${site.id}/lists?$select=id,displayName`, token);
    const list = (lists.value || []).find(x => x.displayName === pb.listName);
    if (!list) throw new Error(`ไม่พบ SharePoint List ชื่อ ${pb.listName} ที่ไซต์ ${pb.sitePath}`);
    const f = pb.fields;
    const fields = `${f.email},${f.bu},${f.division},${f.jobTitle},${f.department}`;
    let next = `/sites/${site.id}/lists/${list.id}/items?$expand=fields($select=${fields})&$select=id,fields&$top=999`;
    const rows = [];
    while (next) {
      const page = await graphGet(next, token);
      rows.push(...(page.value || []).map(item => ({
        ID: item.id,
        Email: fieldText(item.fields[f.email]),
        BU: fieldText(item.fields[f.bu]),
        Division: fieldText(item.fields[f.division]),
        JobTitle: fieldText(item.fields[f.jobTitle]),
        Department: fieldText(item.fields[f.department]),
      })));
      next = page["@odata.nextLink"] || null;
    }
    return rows;
  }
  function fillRefFilter(rows) {
    const select=$("refFilter"), previous=select.value;
    const refMsg=new Map();
    rows.forEach(r=>{
      const ref=clean(r.RefID);
      if (!ref) return;
      if (!refMsg.has(ref)) refMsg.set(ref, "");
      const msg=fieldText(r.iMsg);
      if (msg && !refMsg.get(ref)) refMsg.set(ref, msg);
    });
    const refs=[...refMsg.keys()].sort((a,b)=>Number(b)-Number(a));
    select.innerHTML='<option value="all">All RefID</option>' + refs.map(r=>{
      const msg=refMsg.get(r);
      const msgShort=msg ? (msg.length>50 ? msg.slice(0,50)+"…" : msg) : "(ไม่มีข้อความ iMsg)";
      const label=`${r} — ${msgShort}`;
      return `<option value="${escapeHtml(r)}" title="${escapeHtml(msg||r)}">${escapeHtml(label)}</option>`;
    }).join("");
    if (refs.includes(previous)) select.value=previous;
  }
  // Drops every per-table/per-view filter state back to "no filter" — chart-click table filter,
  // column filters, page position, the Report tab's 5 filters, and the Dashboard's own 5 filters
  // (added 2026-09-25) — plus resets every filter <select>/<input> in the DOM to match. Called
  // when the DATA SCOPE itself changes (RefID filter change, Refresh) — a filter selection made
  // against the old scope may not even make sense against the new one. Deliberately NOT called
  // from render() itself (that used to reset everything on every call) or from any individual
  // filter dropdown's own change handler — picking one filter should never clear the others.
  function resetAllFilters() {
    activeFilter=null;
    columnFilters={id:"", refid:"", email:"", created:"", status:"", drill:""};
    document.querySelectorAll(".col-filter").forEach(input=>{ input.value=""; });
    currentPage=1;
    missingBuFilter=null;
    missingDeptFilter=null;
    reportStatusFilter=null;
    reportTypeFilter=null;
    reportBucketFilter=null;
    missingPage=1;
    dashBuFilter=null;
    dashDeptFilter=null;
    dashStatusFilter=null;
    dashTypeFilter=null;
    dashBucketFilter=null;
    helpBuFilter=null;
    helpDeptFilter=null;
    helpStatusFilter=null;
    helpTypeFilter=null;
    helpBucketFilter=null;
    helpReasonFilter=null;
    helpPage=1;
    quizBuFilter=null;
    quizDeptFilter=null;
    quizStatusFilter=null;
    quizTypeFilter=null;
    quizBucketFilter=null;
    quizCorrectFilter=null;
    quizPage=1;
    ["missingBuFilterSelect","missingDeptFilterSelect","missingStatusFilter","missingTypeFilterSelect","missingBucketFilterSelect",
     "dashBuFilterSelect","dashDeptFilterSelect","dashStatusFilterSelect","dashTypeFilterSelect","dashBucketFilterSelect",
     "helpBuFilterSelect","helpDeptFilterSelect","helpStatusFilterSelect","helpTypeFilterSelect","helpBucketFilterSelect",
     "quizBuFilterSelect","quizDeptFilterSelect","quizStatusFilterSelect","quizTypeFilterSelect","quizBucketFilterSelect"].forEach(id=>{
      const el=$(id); if (el) el.value="";
    });
  }
  function render() {
    const ref=$("refFilter").value;
    const scoped=ref === "all" ? allRows : allRows.filter(r=>clean(r.RefID)===ref);
    // Exclude the Admin announcement row (has Mode/iMsg) from every metric and the table — only member responses stay.
    const rows=scoped.filter(r=>!isNotificationRow(r));
    console.info(`[CallTree Dashboard] RefID=${ref}: ทั้งหมด ${scoped.length} แถว, ตัดรายการแจ้งเหตุการณ์ของ Admin (มี Mode/iMsg) ออก ${scoped.length-rows.length} แถว, เหลือ ${rows.length} แถว`);
    // SECURITY (fixed 2026-09-23): this used to also console.warn() the raw row object
    // (scoped[0]) when no row got excluded, to help debug Mode/iMsg field names. That dumped
    // real employee data (email, response text, etc.) into the browser console, where it could
    // leak via a shared screenshot. Keep only the aggregate counts above — no raw row objects
    // are logged anywhere in this file.
    if (scoped.length && rows.length===scoped.length) {
      console.warn(`[CallTree Dashboard] RefID=${ref}: ไม่มีแถวใดถูกตัดออกเลย — ถ้าคาดว่าควรมีรายการแจ้งเหตุการณ์ถูกกรองออก ให้ตรวจสอบค่า Mode/iMsg ของ List ใน SharePoint โดยตรง`);
    }
    const emailOf=r=>clean(r.EMail).toLowerCase();

    // Reference time for "how long did this member take to respond" is the Admin/notification
    // row's own Created (the moment the event was announced) for that RefID — NOT each member
    // row's ClickDateTime, which is unreliable/blank in this list.
    const adminCreatedByRef={};
    scoped.forEach(r=>{
      if (!isNotificationRow(r) || !r.Created) return;
      const key=clean(r.RefID)||"Unknown";
      if (!adminCreatedByRef[key]) adminCreatedByRef[key]=new Date(r.Created);
    });
    // Each person's EARLIEST response Created time within the current RefID scope, plus the
    // resulting 15-minute bucket (responseBucketByEmail) — computed ONCE here (unaffected by any
    // filter, since it's intrinsic to when the person actually responded) and reused by: the
    // Report tab's "เวลาที่ตอบรับ" column/filter (allMasterRows below), and the Dashboard's own
    // "จำนวนผู้ตอบตามช่วงเวลา 15 นาที" filter (added 2026-09-25, see the dashboard filter-bar
    // section below).
    const earliestResponseByEmail={};
    rows.forEach(r=>{
      const email=emailOf(r);
      if (!email || !r.Created) return;
      const t=new Date(r.Created);
      if (!earliestResponseByEmail[email] || t<earliestResponseByEmail[email].time) {
        earliestResponseByEmail[email]={time:t, ref:clean(r.RefID)||"Unknown"};
      }
    });
    const responseBucketByEmail={};
    Object.entries(earliestResponseByEmail).forEach(([email,resp])=>{
      const adminTime=adminCreatedByRef[resp.ref];
      if (!adminTime) return;
      const mins=(resp.time-adminTime)/60000;
      if (Number.isFinite(mins) && mins>=0) responseBucketByEmail[email]=Math.floor(mins/15)*15;
    });

    // Master (Phone Book) list. Dedupe by email first (keep the first row per email) in case the
    // same person has more than one row there.
    const phoneBookByEmail={};
    allPhoneBook.forEach(p=>{
      const email=clean(p.Email).toLowerCase();
      if (!email || phoneBookByEmail[email]) return;
      phoneBookByEmail[email]=p;
    });
    const phoneBookEmails=Object.keys(phoneBookByEmail);

    // TRUE (unfiltered) responded set — every unique email with at least one member row in the
    // current RefID scope. NOT filtered through hasResponse() (ClickDateTime/ResponseSafe/
    // DrillResponse are legitimately blank for some Modes — see the notes doc's row-types
    // section). This is what the Report tab and the Dashboard filter bar's own ตอบ/ไม่ตอบ filter
    // (below) both check against.
    const respondedEmailsAll=new Set(rows.map(emailOf).filter(Boolean));

    // Admin announcement text (iMsg) per RefID — used by both the existing "% ผู้ที่ตอบผิด" KPI
    // (below) and the new "รายงานผลการตอบคำถาม" tab's correct/wrong comparison (moved up here,
    // 2026-09-27, so both can share the same computation instead of duplicating it).
    const adminIMsgByRef={};
    scoped.forEach(r=>{
      if (!isNotificationRow(r)) return;
      const key=clean(r.RefID)||"Unknown";
      if (!(key in adminIMsgByRef)) adminIMsgByRef[key]=fieldText(r.iMsg).trim();
    });

    // ---- "รายงานขอความช่วยเหลือ" tab base data — REVISED 2026-09-27 (3rd pass) per the user's
    // explicit correction, with a real SharePoint screenshot showing multiple distinct HelpNote
    // rows (different `ID`s) for the SAME email within the SAME RefID (e.g. sarinya.n@... has
    // BOTH ID 884 and ID 962 under RefID 876 — two separate, real help requests, not a duplicate).
    // This means TWO different things were being conflated before and now need TWO different base
    // arrays:
    //   - `currentHelpRows` — EVERY individual SharePoint list row (by its real `ID`) that has
    //     HelpNote content, completely UNDEDUPED — this is what the TABLE/CSV export show, per the
    //     user's explicit "ในตารางแสดงข้อมูลที่มี HelpNoted ทั้งหมด ไม่ตัด" (show ALL HelpNote
    //     records in the table, don't cut/collapse them) — filtered by the tab's own filter bar
    //     only, nothing else.
    //   - `currentHelpUniqueRows` — the SAME rows collapsed to one per (RefID, Email) — keeping
    //     only the EARLIEST (`Created` ascending — "ID แรก", interpreted as earliest submission
    //     time, same "first/earliest wins" convention as `firstDrillByRefEmail`/
    //     `earliestResponseByEmail` elsewhere in this file; flag it if "ID แรก" was meant as
    //     literally the smallest SharePoint item ID instead — in practice these should agree,
    //     since item IDs are assigned in insertion order) — used ONLY by the 2 charts'
    //     unique-person counts (`renderHelpChart()`/`renderHelpReasonChart()`), per the user's
    //     earlier explicit "นับการตอบ ID แรกของ Email" counting rule. The table is NEVER built
    //     from this deduped set — only the charts are.
    currentHelpRows=[];
    rows.forEach(r=>{
      const isHelpRow=Boolean(clean(r.HelpNote) || clean(r.ResponseSafe).toLowerCase()==="seehelpnote");
      if (!isHelpRow) return;
      const email=emailOf(r);
      if (!email || !r.Created) return;
      const pb=phoneBookByEmail[email];
      currentHelpRows.push({
        ID: r.ID,
        RefID: clean(r.RefID)||"Unknown",
        Email: pb ? (clean(pb.Email)||email) : (clean(r.EMail)||email),
        BU: clean(pb && pb.BU)||"ไม่ระบุ BU",
        Department: clean(pb && pb.Department)||"ไม่ระบุ Department",
        JobTitle: pb ? clean(pb.JobTitle) : "",
        Type: pb ? classifyJobType(pb.JobTitle) : null,
        HelpNote: fieldText(r.HelpNote) || fieldText(r.ResponseSafe),
        Reason: classifyHelpReason(fieldText(r.HelpNote)),
        Created: new Date(r.Created),
        ResponseBucket: responseBucketByEmail[email] ?? null,
      });
    });
    const helpUniqueByRefEmail={};
    currentHelpRows.forEach(p=>{
      const key=p.RefID+"|"+p.Email.toLowerCase();
      if (!helpUniqueByRefEmail[key] || p.Created<helpUniqueByRefEmail[key].Created) helpUniqueByRefEmail[key]=p;
    });
    currentHelpUniqueRows=Object.values(helpUniqueByRefEmail);

    // ---- "รายงานผลการตอบคำถาม" tab base data (added 2026-09-27) — one entry per (RefID, person)
    // drill-answer instance within the current RefID scope, using the SAME "first non-blank Drill
    // Response wins" rule as the existing "% ผู้ที่ตอบผิด" KPI further below (built from `rows`,
    // i.e. NOT scoped by the Dashboard's own filter bar — this tab has its own independent one,
    // applied later inside renderQuizChart()/renderQuizTable(), same unfiltered-base pattern).
    const quizFirstDrillByRefEmail={};
    rows.forEach(r=>{
      const drill=fieldText(r.DrillResponse).trim();
      const email=emailOf(r);
      if (!drill || !email || !r.Created) return;
      const key=(clean(r.RefID)||"Unknown")+"|"+email;
      const t=new Date(r.Created);
      if (!quizFirstDrillByRefEmail[key] || t<quizFirstDrillByRefEmail[key].time) {
        quizFirstDrillByRefEmail[key]={time:t, drill, ref:clean(r.RefID)||"Unknown", email};
      }
    });
    currentQuizRows=Object.values(quizFirstDrillByRefEmail)
      .filter(({ref})=>adminIMsgByRef[ref]!==undefined)
      .map(({time,drill,ref,email})=>{
        const pb=phoneBookByEmail[email];
        return {
          EmailKey: email,
          Email: pb ? (clean(pb.Email)||email) : email,
          RefID: ref,
          BU: clean(pb && pb.BU)||"ไม่ระบุ BU",
          Department: clean(pb && pb.Department)||"ไม่ระบุ Department",
          Type: pb ? classifyJobType(pb.JobTitle) : null,
          Correct: drill===adminIMsgByRef[ref],
          DrillResponse: drill,
          Created: time,
          ResponseBucket: responseBucketByEmail[email] ?? null,
        };
      });
    // ------------------------------------------------------------------------------------------

    // ---- Dashboard filter bar (BU / Department / ตอบ-ไม่ตอบ / Type / 15-minute bucket, added
    // 2026-09-25) — same 5 filters as the "รายงานผู้ที่ยังตอบรับ" tab's filter bar, but this one
    // scopes the WHOLE Dashboard view (every KPI card, both charts, and the Latest Responses
    // table), per the user's explicit request, instead of just one table. BU/Department/Job
    // Title live on the Phone Book, not the Member list, so matching them means joining each
    // Member row's email against phoneBookByEmail.
    function passesDashFilters(email) {
      if (dashBuFilter || dashDeptFilter || dashTypeFilter) {
        const pb=phoneBookByEmail[email];
        // A responder who isn't in the Phone Book has no BU/Department/Job Title to check
        // against — excluded whenever any of these 3 filters is active. If a known responder
        // unexpectedly disappears from the Dashboard after picking a BU/Department/Type, this is
        // why — check whether their email actually exists in the Phone Book list.
        if (!pb) return false;
        if (dashBuFilter && (clean(pb.BU)||"ไม่ระบุ BU")!==dashBuFilter) return false;
        if (dashDeptFilter && (clean(pb.Department)||"ไม่ระบุ Department")!==dashDeptFilter) return false;
        if (dashTypeFilter && classifyJobType(pb.JobTitle)!==dashTypeFilter) return false;
      }
      if (dashStatusFilter) {
        const responded=respondedEmailsAll.has(email);
        if (dashStatusFilter==="responded" && !responded) return false;
        if (dashStatusFilter==="missing" && responded) return false;
      }
      if (dashBucketFilter!==null && responseBucketByEmail[email]!==dashBucketFilter) return false;
      return true;
    }
    // dashRows = Member response rows for an email that passes the filter bar above — feeds
    // Safe/พนักงานที่ตอบทั้งหมด/% ผู้ที่ตอบผิด/the Response Status doughnut/the 15-minute
    // histogram/the Latest Responses table. dashPhoneBookEmails = master-list emails passing the
    // same filter — feeds Total Records/ยังไม่ตอบ/Pending/"Pending Employer". Both replace `rows`/
    // `phoneBookEmails` everywhere BELOW this point — the Report tab (currentReportRows/
    // allMasterRows, built next) intentionally stays UNFILTERED by this bar, since it has its own
    // separate filter bar and the user's request was specifically about the Dashboard ("หน้าแรก").
    const dashRows=rows.filter(r=>passesDashFilters(emailOf(r)));
    const dashPhoneBookEmails=phoneBookEmails.filter(passesDashFilters);
    // ------------------------------------------------------------------------------------------

    // Safe / Responded count unique employees (by email) within the dash-filtered scope, not raw
    // rows — one member can appear on multiple rows for the same RefID (re-submits, drill +
    // safe-check, etc.).
    const safeEmails=new Set(dashRows.filter(isSafe).map(emailOf).filter(Boolean));
    const respondedEmails=new Set(dashRows.map(emailOf).filter(Boolean));
    const safe=safeEmails.size;
    const responded=respondedEmails.size;

    // "ยังไม่แจ้งเหตุ" (has not reported at all), within the dash-filtered master-list scope.
    const missingRows=dashPhoneBookEmails.filter(e=>!respondedEmailsAll.has(e)).map(e=>phoneBookByEmail[e]);

    // "รายงานผู้ที่ยังไม่ตอบรับ/ตอบรับ" (Report tab) data — EVERY master-list person, UNFILTERED
    // by the Dashboard's own filter bar above (see the note there).
    const allMasterRows=phoneBookEmails.map(e=>({
      ...phoneBookByEmail[e],
      Responded:respondedEmailsAll.has(e),
      ResponseTime:earliestResponseByEmail[e]?earliestResponseByEmail[e].time:null,
      ResponseBucket:responseBucketByEmail[e] ?? null,
    }));
    currentReportRows=allMasterRows;
    currentPhoneBookTotal=phoneBookEmails.length;
    // Show "–" (not "0") when the Phone Book hasn't loaded at all, so a genuine "0 after
    // filtering" is never confused with "no master data".
    $("kpiMissing").textContent=phoneBookEmails.length ? missingRows.length.toLocaleString("th-TH") : "–";
    // "Pending" = dash-filtered master employees who have not reported at all — same set as
    // `missingRows` above, so the doughnut's "Pending Employer" slice, the kpiPending KPI (still
    // hidden), and this number all agree by construction.
    const pending=missingRows.length;
    // Pseudo-rows so clicking "Pending Employer" can list people in the Latest Responses table
    // who have NO row there at all (they never submitted anything) — built from the dash-filtered
    // missingRows, so it respects the filter bar too. Only Email is real; every other column is
    // blank ("-") — the existing "||'-'" fallbacks in rowColumnText()/renderTable() already
    // handle that.
    const pendingMasterRows=missingRows.map(p=>({
      ID:`PB-${p.Email}`, RefID:"", EMail:p.Email, Created:null, ResponseSafe:"", HelpNote:"", DrillResponse:"",
    }));
    const help=dashRows.filter(r=>clean(r.HelpNote) || clean(r.ResponseSafe).toLowerCase()==="seehelpnote").length;
    // "% ผู้ที่ตอบผิด": compare each RefID's Admin announcement text (iMsg) against each member's
    // Drill Response, within the dash-filtered scope (dashRows). Counted per UNIQUE email — if a
    // person answered more than once, only their FIRST Drill Response (by earliest Created,
    // among rows that actually have a Drill Response) is used. A blank Drill Response means the
    // member never answered the drill question, so they are excluded from both the numerator and
    // denominator (not counted as "wrong").
    // adminIMsgByRef is now computed earlier in render() (moved 2026-09-27 so the Quiz Answer
    // Report tab's base data can reuse it too) — reused here as-is.
    const firstDrillByRefEmail={};
    dashRows.forEach(r=>{
      const drill=fieldText(r.DrillResponse).trim();
      const email=emailOf(r);
      if (!drill || !email || !r.Created) return;
      const key=(clean(r.RefID)||"Unknown")+"|"+email;
      const t=new Date(r.Created);
      if (!firstDrillByRefEmail[key] || t<firstDrillByRefEmail[key].time) {
        firstDrillByRefEmail[key]={time:t, drill, ref:clean(r.RefID)||"Unknown"};
      }
    });
    const drillEntries=Object.values(firstDrillByRefEmail).filter(({ref})=>adminIMsgByRef[ref]!==undefined);
    const wrongCount=drillEntries.filter(({drill,ref})=>drill!==adminIMsgByRef[ref]).length;
    const wrongPct=drillEntries.length ? (wrongCount/drillEntries.length*100) : 0;
    // "Total Records" shows the (dash-filtered) Phone Book master headcount — "–" (not "0") if
    // Phone Book hasn't loaded, same reasoning as kpiMissing above.
    $("kpiTotal").textContent=phoneBookEmails.length ? dashPhoneBookEmails.length.toLocaleString("th-TH") : "–";
    $("kpiSafe").textContent=safe.toLocaleString("th-TH");
    $("kpiResponded").textContent=responded.toLocaleString("th-TH");
    $("kpiPending").textContent=pending.toLocaleString("th-TH");
    $("kpiAverage").textContent=drillEntries.length ? `${wrongPct.toFixed(1)}%` : "–";

    if (ref === "all") {
      $("modeSubtitle").textContent="";
    } else {
      const info=scoped.find(isNotificationRow) || {};
      const modeVal=fieldText(info.Mode);
      const createdVal=info.Created ? new Date(info.Created).toLocaleString('th-TH') : "";
      const msgVal=fieldText(info.iMsg);
      // SECURITY (fixed 2026-09-23): this used to also console.warn() the full raw `info`
      // object (the Admin/notification row from Graph API) to help debug field-name mismatches.
      // That row can carry sensitive text (iMsg content, etc.), so it must never be dumped to
      // the console in production. Log only that the condition happened, no row data.
      if (!modeVal && info.ID) {
        console.warn(`[CallTree Dashboard] RefID=${ref}: อ่านค่า Mode จาก field "Mode" ไม่ได้ (ได้ค่าว่าง) ทั้งที่ควรมีข้อมูล (แถวประกาศ Admin ID ${info.ID}) — ตรวจสอบชื่อ Internal Field ของคอลัมน์ Mode ใน SharePoint`);
      }
      $("modeSubtitle").textContent=`Mode: ${modeVal||'-'} · Created: ${createdVal||'-'} · iMsg: ${msgVal||'-'}`;
    }

    // Row-level predicates for each doughnut slice, reused by its click handler so the table
    // filter shows exactly the rows that make up that slice. "Safe" mirrors the `safe` number
    // above; "Need help" mirrors `help` above (which is already row-based, so clicking it
    // filters to exactly `help` rows). "Pending Employer" (renamed + redefined 2026-09-24) has
    // no predicate — it has no Member rows to filter at all (these people never submitted
    // anything), so its click handler below swaps in `pendingMasterRows` wholesale instead of
    // filtering `currentRows`. See `getFilteredRows()`'s `activeFilter.rows` branch.
    const statusPredicates={
      safe: isSafe,
      help: r=>Boolean(clean(r.HelpNote) || clean(r.ResponseSafe).toLowerCase()==="seehelpnote"),
    };
    const statusLabels={safe:"Safe", pending:"Pending Employer", help:"Need help"};
    if(statusChart) statusChart.destroy();
    statusChart=new Chart($("statusChart"),{type:"doughnut",data:{labels:["Safe","Pending Employer","Need help"],datasets:[{data:[safe,pending,help],backgroundColor:["#16845b","#C5A153","#d64545"],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      const kinds=["safe","pending","help"];
      const kind=kinds[elements[0].index];
      const key=`status:${kind}`;
      if (activeFilter && activeFilter.key===key) { activeFilter=null; }
      else if (kind==="pending") { activeFilter={key, label:statusLabels.pending, rows:pendingMasterRows}; }
      else { activeFilter={key, label:statusLabels[kind], predicate:statusPredicates[kind]}; }
      currentPage=1;
      renderTable();
    }}});

    // Response-time histogram: count UNIQUE emails within the dash-filtered scope (dashRows), not
    // raw rows. A member with multiple response rows for the same RefID is counted once, using
    // the EARLIEST of their own Created times as their response moment — bucketed against the
    // Admin row's Created for that RefID, in 15-minute buckets from 0.
    const earliestByRefEmail={};
    dashRows.forEach(r=>{
      const email=emailOf(r);
      const key=(clean(r.RefID)||"Unknown")+"|"+email;
      if (!email || !r.Created) return;
      const t=new Date(r.Created);
      if (!earliestByRefEmail[key] || t<earliestByRefEmail[key].time) {
        earliestByRefEmail[key]={time:t, ref:clean(r.RefID)||"Unknown", row:r};
      }
    });
    const bucketCounts={};
    const bucketRowIds={};
    Object.values(earliestByRefEmail).forEach(({time,ref,row})=>{
      const adminTime=adminCreatedByRef[ref];
      if (!adminTime) return;
      const mins=(time-adminTime)/60000;
      if (!Number.isFinite(mins) || mins<0) return;
      const bucketStart=Math.floor(mins/15)*15;
      bucketCounts[bucketStart]=(bucketCounts[bucketStart]||0)+1;
      (bucketRowIds[bucketStart] ??= []).push(row.ID);
    });
    currentBucketRowIds=bucketRowIds;
    const maxBucket=Object.keys(bucketCounts).length ? Math.max(...Object.keys(bucketCounts).map(Number)) : 0;
    const bucketLabels=[], bucketData=[];
    for (let b=0; b<=maxBucket; b+=15) { bucketLabels.push(`${b}–${b+15} นาที`); bucketData.push(bucketCounts[b]||0); }
    if(drillChart) drillChart.destroy();
    drillChart=new Chart($("drillChart"),{type:"bar",data:{labels:bucketLabels,datasets:[{label:"จำนวนผู้ตอบ",data:bucketData,backgroundColor:"#202B49",borderRadius:4,maxBarThickness:56}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{title:items=>items[0].label,label:item=>`${item.parsed.y.toLocaleString('th-TH')} คน`}}},scales:{x:{title:{display:true,text:'นาทีหลังจาก Admin แจ้งเหตุ'},grid:{display:false}},y:{beginAtZero:true,ticks:{precision:0},title:{display:true,text:'จำนวนผู้ตอบ (คนไม่ซ้ำ)'}}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      const bucketStart=elements[0].index*15;
      const key=`bucket:${bucketStart}`;
      if (activeFilter && activeFilter.key===key) { activeFilter=null; }
      else {
        const ids=new Set(currentBucketRowIds[bucketStart]||[]);
        activeFilter={key, label:`ตอบภายใน ${bucketStart}–${bucketStart+15} นาที`, predicate:r=>ids.has(r.ID)};
      }
      currentPage=1;
      renderTable();
    }}});

    populateReportFilterSelects(allMasterRows);
    renderMissingChart();
    populateDashFilterSelects(allMasterRows);
    show("clearDashAllFiltersButton", Boolean(dashBuFilter||dashDeptFilter||dashStatusFilter||dashTypeFilter||dashBucketFilter!==null));
    currentRows=dashRows;
    renderTable();
    renderMissingTable();
    populateHelpFilterSelects(currentHelpRows);
    renderHelpChart();
    renderHelpReasonChart();
    renderHelpKpi();
    show("clearHelpAllFiltersButton", Boolean(helpBuFilter||helpDeptFilter||helpStatusFilter||helpTypeFilter||helpBucketFilter!==null||helpReasonFilter));
    renderHelpTable();
    populateQuizFilterSelects(currentQuizRows);
    renderQuizChart();
    show("clearQuizAllFiltersButton", Boolean(quizBuFilter||quizDeptFilter||quizStatusFilter||quizTypeFilter||quizBucketFilter!==null||quizCorrectFilter));
    renderQuizTable();
  }
  // Builds/rebuilds missingChart: both series (responded vs. not-yet-responded) grouped by
  // Department (changed 2026-09-24, was grouped by BU) so the two counts sit on the same
  // categories for easy comparison — per the user's explicit request. Reads directly from
  // currentReportRows (already tagged BU/Department/Responded per person) rather than the
  // render()-local phoneBookEmails/respondedEmails/etc., so it can be called on its own —
  // NOT just from render() — whenever missingBuFilter changes, since the chart is now scoped
  // to the currently selected BU (cascading, added 2026-09-24, same day): pick a BU and the
  // chart narrows to that BU's Department breakdown instead of showing every Department in
  // the whole Master list at once (which could be dozens of bars — see the notes doc for why).
  // "Responded" here means "in the Phone Book AND has a response in the current RefID scope" —
  // a different (smaller-or-equal) population than the kpiResponded KPI, which doesn't require
  // Phone Book membership.
  function renderMissingChart() {
    let rowsForChart=currentReportRows;
    if (missingBuFilter) rowsForChart=rowsForChart.filter(p=>(clean(p.BU)||"ไม่ระบุ BU")===missingBuFilter);
    const respondedByDept={};
    const missingByDept={};
    rowsForChart.forEach(p=>{
      const dept=clean(p.Department)||"ไม่ระบุ Department";
      if (p.Responded) respondedByDept[dept]=(respondedByDept[dept]||0)+1;
      else missingByDept[dept]=(missingByDept[dept]||0)+1;
    });
    const deptLabels=[...new Set([...Object.keys(respondedByDept), ...Object.keys(missingByDept)])]
      .sort((a,b)=>((respondedByDept[b]||0)+(missingByDept[b]||0)) - ((respondedByDept[a]||0)+(missingByDept[a]||0)));
    const respondedData=deptLabels.map(d=>respondedByDept[d]||0);
    const missingData=deptLabels.map(d=>missingByDept[d]||0);
    if(missingChart) missingChart.destroy();
    missingChart=new Chart($("missingChart"),{type:"bar",data:{labels:deptLabels,datasets:[
      {label:"ตอบรับแล้ว",data:respondedData,backgroundColor:"#16845b",borderRadius:4,maxBarThickness:22},
      {label:"ยังไม่ตอบรับ",data:missingData,backgroundColor:"#202B49",borderRadius:4,maxBarThickness:22},
    ]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"},tooltip:{callbacks:{label:item=>`${item.dataset.label}: ${item.parsed.x.toLocaleString('th-TH')} คน`}}},scales:{x:{beginAtZero:true,ticks:{precision:0}},y:{grid:{display:false}}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      // Both datasets share the same category (Department) labels, so any bar clicked in either
      // series resolves to the same Department via its category index — filters the table to
      // that Department (clicking a "ตอบรับแล้ว" bar filters the table to that Department too,
      // which can legitimately show only responded people, or come back empty under the ตอบ/
      // ไม่ตอบ filter, depending on what else is selected).
      const dept=deptLabels[elements[0].index];
      missingDeptFilter=(missingDeptFilter===dept) ? null : dept;
      const deptSelect=$("missingDeptFilterSelect");
      if (deptSelect) deptSelect.value=missingDeptFilter||"";
      missingPage=1;
      renderMissingTable();
    }}});
  }
  // Populates the BU / Department / จำนวนผู้ตอบตามช่วงเวลา 15 นาที filter-bar dropdowns
  // (added 2026-09-24) from whatever's actually in the current Phone Book data, preserving each
  // dropdown's current selection if it's still a valid option (same "preserve previous" pattern
  // as fillRefFilter()). BU/Department blanks are bucketed as "ไม่ระบุ BU"/"ไม่ระบุ Department"
  // — matching how renderMissingTable()'s own filtering already treats blanks, so a value shown
  // here always matches something below. The bucket dropdown only lists buckets that actually
  // have at least one person in them (via each row's `ResponseBucket`, computed in render()).
  // CASCADING (added 2026-09-24): Department is scoped to the currently-selected BU (missingBuFilter)
  // so the dropdown never offers a Department that doesn't exist within the chosen BU. If the
  // previously-selected Department falls outside the new BU scope, it's cleared back to "ทั้งหมด"
  // (both the module-level missingDeptFilter and the <select>'s value) rather than silently kept.
  function populateReportFilterSelects(rows) {
    const buSelect=$("missingBuFilterSelect"), deptSelect=$("missingDeptFilterSelect"), bucketSelect=$("missingBucketFilterSelect");
    const buValues=[...new Set(rows.map(p=>clean(p.BU)||"ไม่ระบุ BU"))].sort((a,b)=>a.localeCompare(b));
    const deptScope = missingBuFilter ? rows.filter(p=>(clean(p.BU)||"ไม่ระบุ BU")===missingBuFilter) : rows;
    const deptValues=[...new Set(deptScope.map(p=>clean(p.Department)||"ไม่ระบุ Department"))].sort((a,b)=>a.localeCompare(b));
    const bucketValues=[...new Set(rows.map(p=>p.ResponseBucket).filter(b=>b!==null&&b!==undefined))].sort((a,b)=>a-b);
    const prevBu=buSelect.value, prevDept=deptSelect.value, prevBucket=bucketSelect.value;
    buSelect.innerHTML='<option value="">ทั้งหมด</option>'+buValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    deptSelect.innerHTML='<option value="">ทั้งหมด</option>'+deptValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    bucketSelect.innerHTML='<option value="">ทั้งหมด</option>'+bucketValues.map(b=>`<option value="${b}">${b}–${b+15} นาที</option>`).join("");
    if (buValues.includes(prevBu)) buSelect.value=prevBu;
    if (deptValues.includes(prevDept)) {
      deptSelect.value=prevDept;
    } else if (prevDept) {
      deptSelect.value="";
      missingDeptFilter=null;
    }
    if (bucketValues.some(b=>String(b)===prevBucket)) bucketSelect.value=prevBucket;
  }
  // Populates the Dashboard's OWN BU / Department / 15-minute-bucket filter dropdowns (added
  // 2026-09-25) — same pattern/cascading behavior as populateReportFilterSelects() above (BU's
  // option list is always full; Department's option list is scoped to whichever BU is currently
  // selected, dashBuFilter, and the previous Department selection is cleared if it no longer fits
  // the new scope), but reading dashBu*/dashDept* state and writing to the dash*Select> elements
  // instead of the Report tab's. Always called with the FULL (unfiltered) allMasterRows — Type/
  // ตอบ-ไม่ตอบ never narrow BU/Department's own option lists, only BU narrows Department's.
  function populateDashFilterSelects(rows) {
    const buSelect=$("dashBuFilterSelect"), deptSelect=$("dashDeptFilterSelect"), bucketSelect=$("dashBucketFilterSelect");
    const buValues=[...new Set(rows.map(p=>clean(p.BU)||"ไม่ระบุ BU"))].sort((a,b)=>a.localeCompare(b));
    const deptScope = dashBuFilter ? rows.filter(p=>(clean(p.BU)||"ไม่ระบุ BU")===dashBuFilter) : rows;
    const deptValues=[...new Set(deptScope.map(p=>clean(p.Department)||"ไม่ระบุ Department"))].sort((a,b)=>a.localeCompare(b));
    const bucketValues=[...new Set(rows.map(p=>p.ResponseBucket).filter(b=>b!==null&&b!==undefined))].sort((a,b)=>a-b);
    const prevBu=buSelect.value, prevDept=deptSelect.value, prevBucket=bucketSelect.value;
    buSelect.innerHTML='<option value="">ทั้งหมด</option>'+buValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    deptSelect.innerHTML='<option value="">ทั้งหมด</option>'+deptValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    bucketSelect.innerHTML='<option value="">ทั้งหมด</option>'+bucketValues.map(b=>`<option value="${b}">${b}–${b+15} นาที</option>`).join("");
    if (buValues.includes(prevBu)) buSelect.value=prevBu;
    if (deptValues.includes(prevDept)) {
      deptSelect.value=prevDept;
    } else if (prevDept) {
      deptSelect.value="";
      dashDeptFilter=null;
    }
    if (bucketValues.some(b=>String(b)===prevBucket)) bucketSelect.value=prevBucket;
  }
  // Renders the "รายงานผู้ที่ยังไม่ตอบรับ/ตอบรับ" table: currentReportRows (EVERY master-list
  // person, tagged Responded true/false — widened 2026-09-24), filtered by missingBuFilter,
  // missingDeptFilter, reportStatusFilter, reportTypeFilter and reportBucketFilter (the
  // filter-bar's 5 dropdowns — BU is also settable by clicking a missingChart bar, kept in sync
  // both ways), sorted by Email, paginated at PAGE_SIZE (30/page — same as the Member table, so
  // a large Master list is never "crammed" into one scrolling box). Called on its own by the
  // chart's onClick, the filter-bar dropdowns, the pagination buttons, and the clear-filter
  // link/button — never destroys/rebuilds missingChart.
  function renderMissingTable() {
    let list=currentReportRows;
    if (missingBuFilter) list=list.filter(p=>(clean(p.BU)||"ไม่ระบุ BU")===missingBuFilter);
    if (missingDeptFilter) list=list.filter(p=>(clean(p.Department)||"ไม่ระบุ Department")===missingDeptFilter);
    if (reportTypeFilter) list=list.filter(p=>classifyJobType(p.JobTitle)===reportTypeFilter);
    if (reportBucketFilter!==null) list=list.filter(p=>p.ResponseBucket===reportBucketFilter);
    if (reportStatusFilter==="responded") list=list.filter(p=>p.Responded);
    else if (reportStatusFilter==="missing") list=list.filter(p=>!p.Responded);
    const sorted=[...list].sort((a,b)=>clean(a.Email).localeCompare(clean(b.Email)));

    const total=sorted.length;
    const totalPages=Math.max(1, Math.ceil(total/PAGE_SIZE));
    if (missingPage>totalPages) missingPage=totalPages;
    if (missingPage<1) missingPage=1;
    const startIdx=(missingPage-1)*PAGE_SIZE;
    const pageRows=sorted.slice(startIdx, startIdx+PAGE_SIZE);

    const filterParts=[];
    if (missingBuFilter) filterParts.push(`BU: <strong>${escapeHtml(missingBuFilter)}</strong>`);
    if (missingDeptFilter) filterParts.push(`Department: <strong>${escapeHtml(missingDeptFilter)}</strong>`);
    if (reportStatusFilter) filterParts.push(`สถานะ: <strong>${reportStatusFilter==="responded"?"ตอบรับแล้ว":"ยังไม่ตอบรับ"}</strong>`);
    if (reportTypeFilter) filterParts.push(`Type: <strong>${reportTypeFilter==="management"?"Management":"Staff"}</strong>`);
    if (reportBucketFilter!==null) filterParts.push(`ช่วงเวลาตอบ: <strong>${reportBucketFilter}–${reportBucketFilter+15} นาที</strong>`);
    const anyFilterActive=Boolean(missingBuFilter||missingDeptFilter||reportStatusFilter||reportTypeFilter||reportBucketFilter!==null);
    const filterNote=filterParts.length ? ` — กรองตาม ${filterParts.join(", ")}` : "";
    const rangeText=total ? `${startIdx+1}–${Math.min(startIdx+PAGE_SIZE,total)}` : "0";
    $("missingSummary").innerHTML=currentPhoneBookTotal
      ? `แสดง ${rangeText} จาก ${total.toLocaleString("th-TH")} รายการ (Master ทั้งหมด ${currentPhoneBookTotal.toLocaleString("th-TH")} คน)${filterNote}`
      : `ยังไม่ได้โหลดข้อมูล Phone Book (Master List)`;
    show("clearMissingAllFiltersButton", anyFilterActive);
    $("missingTable").innerHTML=pageRows.map(p=>{
      const cls=p.Responded?"responded":"pending";
      const label=p.Responded?"ตอบรับแล้ว":"ยังไม่ตอบรับ";
      const responseTimeText=p.ResponseTime?escapeHtml(new Date(p.ResponseTime).toLocaleString("th-TH")):'-';
      return `<tr><td>${escapeHtml(p.Email)||'-'}</td><td>${escapeHtml(p.BU)||'-'}</td><td>${escapeHtml(p.Division)||'-'}</td><td>${escapeHtml(p.JobTitle)||'-'}</td><td>${escapeHtml(p.Department)||'-'}</td><td><span class="badge ${cls}">${label}</span></td><td>${responseTimeText}</td></tr>`;
    }).join("") || `<tr><td colspan="7">${currentPhoneBookTotal ? (anyFilterActive?'ไม่พบข้อมูลที่ตรงกับตัวกรอง':'ไม่พบข้อมูล') : '-'}</td></tr>`;

    $("missingPageIndicator").textContent=`หน้า ${missingPage} / ${totalPages}`;
    $("missingPrevPageButton").disabled=missingPage<=1;
    $("missingNextPageButton").disabled=missingPage>=totalPages;
  }
  // Exports the FULL current match set (all 5 filter-bar filters applied — BU, Department,
  // ตอบ/ไม่ตอบ, Type, ช่วงเวลาตอบ — ALL pages, not just the visible page) — same
  // csvEscape/BOM/downloadBlob pattern as exportCsv(). Includes Status and เวลาที่ตอบรับ columns.
  function exportMissingCsv() {
    let list=currentReportRows;
    if (missingBuFilter) list=list.filter(p=>(clean(p.BU)||"ไม่ระบุ BU")===missingBuFilter);
    if (missingDeptFilter) list=list.filter(p=>(clean(p.Department)||"ไม่ระบุ Department")===missingDeptFilter);
    if (reportTypeFilter) list=list.filter(p=>classifyJobType(p.JobTitle)===reportTypeFilter);
    if (reportBucketFilter!==null) list=list.filter(p=>p.ResponseBucket===reportBucketFilter);
    if (reportStatusFilter==="responded") list=list.filter(p=>p.Responded);
    else if (reportStatusFilter==="missing") list=list.filter(p=>!p.Responded);
    if (!list.length) { setMessage("ไม่มีรายชื่อตามตัวกรองปัจจุบัน (หรือยังไม่ได้โหลด Phone Book)", "error"); return; }
    const csvEscape=v=>{ const s=String(v ?? ""); return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s; };
    const headers=["Email","BU","Division","JobTitle","Department","Status","ResponseTime"];
    const lines=[headers.join(",")];
    list.forEach(p=>{ lines.push([clean(p.Email), clean(p.BU), clean(p.Division), clean(p.JobTitle), clean(p.Department), p.Responded?"ตอบรับแล้ว":"ยังไม่ตอบรับ", p.ResponseTime?new Date(p.ResponseTime).toLocaleString("th-TH"):""].map(csvEscape).join(",")); });
    const blob=new Blob(["﻿"+lines.join("\r\n")], {type:"text/csv;charset=utf-8;"});
    downloadBlob(blob, `phonebook-report-${exportTimestamp()}.csv`);
  }
  // Shared base filter for the Help tab's 2 charts + table (added 2026-09-27, refactored out of
  // what used to be renderHelpChart()'s/filterHelpRows()'s own separate inline filtering when the
  // reason-breakdown chart was added). BU/Type/15-min-bucket/ตอบ-ไม่ตอบ always apply; Department
  // and Reason each apply UNLESS that dimension is the chart currently being built (same
  // "an axis dimension doesn't narrow itself" rule as renderMissingChart() only applying BU, never
  // Department, to itself) — excludeDept=true for helpDeptChart, excludeReason=true for
  // helpReasonChart, both false for the table/CSV export (which apply every filter, including the
  // chart-click-only helpDeptFilter/helpReasonFilter). `sourceRows` is which base array to filter —
  // callers MUST pass `currentHelpUniqueRows` for the 2 charts (one count per RefID+Email) and
  // `currentHelpRows` for the table/export (every real row, undeduped) — see the module-level
  // comment above `currentHelpRows`'s declaration for why these are two different arrays.
  function baseHelpFilterList(sourceRows, excludeDept, excludeReason) {
    let list=sourceRows;
    if (helpBuFilter) list=list.filter(p=>p.BU===helpBuFilter);
    if (!excludeDept && helpDeptFilter) list=list.filter(p=>p.Department===helpDeptFilter);
    if (helpTypeFilter) list=list.filter(p=>p.Type===helpTypeFilter);
    if (helpBucketFilter!==null) list=list.filter(p=>p.ResponseBucket===helpBucketFilter);
    if (helpStatusFilter==="missing") list=[];
    if (!excludeReason && helpReasonFilter) list=list.filter(p=>p.Reason===helpReasonFilter);
    return list;
  }
  // Fills the 3 KPI cards on the "ข้อมูลสรุปตาม KPI Dashboard" sub-tab (added 2026-09-27). Applies
  // ALL of the tab's own filters (like the table/export — no axis exclusion, since these cards
  // aren't a chart axis): จำนวนคำขอความช่วยเหลือ uses currentHelpUniqueRows (1 ต่อ RefID+Email,
  // matching the 2 charts' own counting basis); จำนวนรายการ HelpNote ทั้งหมด and จำนวน RefID ใช้
  // currentHelpRows (ทุกแถวจริง ไม่ deduplicate — matching the table's own row count).
  function renderHelpKpi() {
    const uniqueList=baseHelpFilterList(currentHelpUniqueRows, false, false);
    const rawList=baseHelpFilterList(currentHelpRows, false, false);
    $("helpKpiUnique").textContent=uniqueList.length.toLocaleString("th-TH");
    $("helpKpiRaw").textContent=rawList.length.toLocaleString("th-TH");
    $("helpKpiRefCount").textContent=new Set(rawList.map(p=>p.RefID)).size.toLocaleString("th-TH");
  }
  // Builds/rebuilds helpDeptChart (added 2026-09-27): unique help-requesters grouped by
  // Department, same "BU narrows / Department is the axis so it's excluded from the chart's own
  // narrowing" pattern as renderMissingChart() — helpDeptFilter is NOT applied here (only to the
  // table below), or picking a bar would immediately empty every other bar. ตอบ/ไม่ตอบ is included
  // for filter-set consistency with the Dashboard's own bar, even though everyone in
  // currentHelpRows has, by definition, already responded (so "missing" always yields an empty
  // chart — expected, not a bug). helpReasonFilter DOES narrow this chart (Reason isn't this
  // chart's own axis).
  function renderHelpChart() {
    const list=baseHelpFilterList(currentHelpUniqueRows, true, false);
    const countByDept={};
    list.forEach(p=>{ countByDept[p.Department]=(countByDept[p.Department]||0)+1; });
    const deptLabels=Object.keys(countByDept).sort((a,b)=>countByDept[b]-countByDept[a]);
    const data=deptLabels.map(d=>countByDept[d]);
    if(helpChart) helpChart.destroy();
    helpChart=new Chart($("helpDeptChart"),{type:"bar",data:{labels:deptLabels,datasets:[
      {label:"จำนวนคำขอความช่วยเหลือ (1 คนต่อ 1 ครั้งต่อ RefID)",data,backgroundColor:"#d64545",borderRadius:4,maxBarThickness:28},
    ]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:item=>`${item.parsed.x.toLocaleString('th-TH')} คน`}}},scales:{x:{beginAtZero:true,ticks:{precision:0}},y:{grid:{display:false}}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      const dept=deptLabels[elements[0].index];
      helpDeptFilter=(helpDeptFilter===dept) ? null : dept;
      const deptSelect=$("helpDeptFilterSelect");
      if (deptSelect) deptSelect.value=helpDeptFilter||"";
      helpPage=1;
      renderHelpKpi();
      renderHelpTable();
    }}});
  }
  // Builds/rebuilds helpReasonChart (added 2026-09-27, once real HelpNote sample text was
  // available — see classifyHelpReason() for the derivation rule and the severity-priority
  // decision behind it): unique help-requesters grouped by Reason (บาดเจ็บ/หลงทาง/ติดในอาคาร/อื่นๆ,
  // in HELP_REASON_ORDER — the order the user originally asked for, not the severity-priority
  // order used to pick each person's single bucket). helpReasonFilter is NOT applied here (Reason
  // is this chart's own axis — same rule as helpDeptFilter/helpDeptChart above); helpDeptFilter
  // DOES narrow this chart.
  function renderHelpReasonChart() {
    const list=baseHelpFilterList(currentHelpUniqueRows, false, true);
    const countByReason={injury:0, trapped:0, lost:0, other:0};
    list.forEach(p=>{ countByReason[p.Reason]=(countByReason[p.Reason]||0)+1; });
    const reasonKeys=HELP_REASON_ORDER;
    const labels=reasonKeys.map(k=>HELP_REASON_LABELS[k]);
    const data=reasonKeys.map(k=>countByReason[k]||0);
    if(helpReasonChart) helpReasonChart.destroy();
    helpReasonChart=new Chart($("helpReasonChart"),{type:"bar",data:{labels,datasets:[
      {label:"จำนวนคำขอความช่วยเหลือ (1 คนต่อ 1 ครั้งต่อ RefID)",data,backgroundColor:["#d64545","#C5A153","#202B49","#8b8e91"],borderRadius:4,maxBarThickness:48},
    ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:item=>`${item.parsed.y.toLocaleString('th-TH')} คน`}}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,ticks:{precision:0}}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      const reasonKey=reasonKeys[elements[0].index];
      helpReasonFilter=(helpReasonFilter===reasonKey) ? null : reasonKey;
      helpPage=1;
      renderHelpKpi();
      renderHelpTable();
    }}});
  }
  // Populates the Help tab's BU / Department / 15-minute-bucket dropdowns from currentHelpRows
  // (help-requesters only, not the whole Phone Book) — same cascading-Department-by-BU pattern as
  // populateReportFilterSelects()/populateDashFilterSelects().
  function populateHelpFilterSelects(rows) {
    const buSelect=$("helpBuFilterSelect"), deptSelect=$("helpDeptFilterSelect"), bucketSelect=$("helpBucketFilterSelect");
    const buValues=[...new Set(rows.map(p=>p.BU))].sort((a,b)=>a.localeCompare(b));
    const deptScope = helpBuFilter ? rows.filter(p=>p.BU===helpBuFilter) : rows;
    const deptValues=[...new Set(deptScope.map(p=>p.Department))].sort((a,b)=>a.localeCompare(b));
    const bucketValues=[...new Set(rows.map(p=>p.ResponseBucket).filter(b=>b!==null&&b!==undefined))].sort((a,b)=>a-b);
    const prevBu=buSelect.value, prevDept=deptSelect.value, prevBucket=bucketSelect.value;
    buSelect.innerHTML='<option value="">ทั้งหมด</option>'+buValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    deptSelect.innerHTML='<option value="">ทั้งหมด</option>'+deptValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    bucketSelect.innerHTML='<option value="">ทั้งหมด</option>'+bucketValues.map(b=>`<option value="${b}">${b}–${b+15} นาที</option>`).join("");
    if (buValues.includes(prevBu)) buSelect.value=prevBu;
    if (deptValues.includes(prevDept)) {
      deptSelect.value=prevDept;
    } else if (prevDept) {
      deptSelect.value="";
      helpDeptFilter=null;
    }
    if (bucketValues.some(b=>String(b)===prevBucket)) bucketSelect.value=prevBucket;
  }
  // Renders the Help Request Report table: currentHelpRows filtered by ALL 5 of the tab's own
  // filters (unlike the chart above, helpDeptFilter DOES apply here — matching renderMissingTable()
  // applying all 5 of its own filters while renderMissingChart() only applies BU).
  function filterHelpRows() {
    return baseHelpFilterList(currentHelpRows, false, false);
  }
  function renderHelpTable() {
    const list=filterHelpRows();
    // Sort by Email, then RefID, then Created — currentHelpRows is now completely undeduped (one
    // row per real SharePoint `ID`), so the same person can have several rows even within a single
    // RefID (see the module-level comment above currentHelpRows' declaration) — this keeps a
    // person's entries grouped together and chronological within that grouping.
    const sorted=[...list].sort((a,b)=>clean(a.Email).localeCompare(clean(b.Email)) || clean(a.RefID).localeCompare(clean(b.RefID)) || (a.Created-b.Created));
    const total=sorted.length;
    const totalPages=Math.max(1, Math.ceil(total/PAGE_SIZE));
    if (helpPage>totalPages) helpPage=totalPages;
    if (helpPage<1) helpPage=1;
    const startIdx=(helpPage-1)*PAGE_SIZE;
    const pageRows=sorted.slice(startIdx, startIdx+PAGE_SIZE);

    const filterParts=[];
    if (helpBuFilter) filterParts.push(`BU: <strong>${escapeHtml(helpBuFilter)}</strong>`);
    if (helpDeptFilter) filterParts.push(`Department: <strong>${escapeHtml(helpDeptFilter)}</strong>`);
    if (helpStatusFilter) filterParts.push(`สถานะ: <strong>${helpStatusFilter==="responded"?"ตอบรับแล้ว":"ยังไม่ตอบรับ"}</strong>`);
    if (helpTypeFilter) filterParts.push(`Type: <strong>${helpTypeFilter==="management"?"Management":"Staff"}</strong>`);
    if (helpBucketFilter!==null) filterParts.push(`ช่วงเวลาตอบ: <strong>${helpBucketFilter}–${helpBucketFilter+15} นาที</strong>`);
    if (helpReasonFilter) filterParts.push(`เหตุผล: <strong>${HELP_REASON_LABELS[helpReasonFilter]}</strong>`);
    const anyFilterActive=Boolean(helpBuFilter||helpDeptFilter||helpStatusFilter||helpTypeFilter||helpBucketFilter!==null||helpReasonFilter);
    const filterNote=filterParts.length ? ` — กรองตาม ${filterParts.join(", ")}` : "";
    const rangeText=total ? `${startIdx+1}–${Math.min(startIdx+PAGE_SIZE,total)}` : "0";
    // "รายการ" (records), not "คน" (people) — since RefID+Email grouping (not Email-only) means
    // the same person can appear more than once under "All RefID" (see currentHelpRows above).
    $("helpSummary").innerHTML=`แสดง ${rangeText} จาก ${total.toLocaleString("th-TH")} รายการขอความช่วยเหลือ${filterNote}`;
    show("clearHelpAllFiltersButton", anyFilterActive);
    $("helpTable").innerHTML=pageRows.map(p=>{
      const created=p.Created?escapeHtml(new Date(p.Created).toLocaleString("th-TH")):'-';
      return `<tr><td>${escapeHtml(p.ID)}</td><td>${escapeHtml(p.Email)||'-'}</td><td>${escapeHtml(p.RefID)||'-'}</td><td>${escapeHtml(p.BU)||'-'}</td><td>${escapeHtml(p.Department)||'-'}</td><td>${escapeHtml(p.JobTitle)||'-'}</td><td>${HELP_REASON_LABELS[p.Reason]}</td><td>${escapeHtml(p.HelpNote)||'-'}</td><td>${created}</td></tr>`;
    }).join("") || `<tr><td colspan="9">${anyFilterActive?'ไม่พบข้อมูลที่ตรงกับตัวกรอง':'ไม่พบข้อมูล'}</td></tr>`;
    $("helpPageIndicator").textContent=`หน้า ${helpPage} / ${totalPages}`;
    $("helpPrevPageButton").disabled=helpPage<=1;
    $("helpNextPageButton").disabled=helpPage>=totalPages;
  }
  function exportHelpCsv() {
    const list=filterHelpRows();
    if (!list.length) { setMessage("ไม่มีรายชื่อผู้ขอความช่วยเหลือตามตัวกรองปัจจุบัน", "error"); return; }
    const csvEscape=v=>{ const s=String(v ?? ""); return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s; };
    const headers=["ID","Email","RefID","BU","Department","JobTitle","Reason","HelpNote","Created"];
    const lines=[headers.join(",")];
    list.forEach(p=>{ lines.push([p.ID, clean(p.Email), clean(p.RefID), clean(p.BU), clean(p.Department), clean(p.JobTitle), HELP_REASON_LABELS[p.Reason], clean(p.HelpNote), p.Created?new Date(p.Created).toLocaleString("th-TH"):""].map(csvEscape).join(",")); });
    const blob=new Blob(["﻿"+lines.join("\r\n")], {type:"text/csv;charset=utf-8;"});
    downloadBlob(blob, `help-request-report-${exportTimestamp()}.csv`);
  }
  // Builds/rebuilds quizChart (added 2026-09-27): correct-vs-wrong pie over currentQuizRows,
  // narrowed by all 5 of the Quiz tab's own filters (unlike the Help tab's chart, there's no
  // "axis" here to exclude — Correct/Wrong is what's plotted, not Department — so quizDeptFilter
  // DOES narrow this chart too, same as every other filter).
  function filterQuizRowsForChart() {
    let list=currentQuizRows;
    if (quizBuFilter) list=list.filter(p=>p.BU===quizBuFilter);
    if (quizDeptFilter) list=list.filter(p=>p.Department===quizDeptFilter);
    if (quizTypeFilter) list=list.filter(p=>p.Type===quizTypeFilter);
    if (quizBucketFilter!==null) list=list.filter(p=>p.ResponseBucket===quizBucketFilter);
    if (quizStatusFilter==="missing") list=[];
    return list;
  }
  function renderQuizChart() {
    const list=filterQuizRowsForChart();
    const correctCount=list.filter(p=>p.Correct).length;
    const wrongCount=list.length-correctCount;
    if(quizChart) quizChart.destroy();
    quizChart=new Chart($("quizChart"),{type:"pie",data:{labels:["ตอบถูก","ตอบผิด"],datasets:[{data:[correctCount,wrongCount],backgroundColor:["#16845b","#d64545"],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"},tooltip:{callbacks:{label:item=>`${item.label}: ${item.parsed.toLocaleString('th-TH')} คน`}}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      const kind=elements[0].index===0 ? "correct" : "wrong";
      quizCorrectFilter=(quizCorrectFilter===kind) ? null : kind;
      quizPage=1;
      renderQuizTable();
    }}});
  }
  // Populates the Quiz tab's BU / Department / 15-minute-bucket dropdowns from currentQuizRows —
  // same cascading pattern as populateHelpFilterSelects()/populateDashFilterSelects().
  function populateQuizFilterSelects(rows) {
    const buSelect=$("quizBuFilterSelect"), deptSelect=$("quizDeptFilterSelect"), bucketSelect=$("quizBucketFilterSelect");
    const buValues=[...new Set(rows.map(p=>p.BU))].sort((a,b)=>a.localeCompare(b));
    const deptScope = quizBuFilter ? rows.filter(p=>p.BU===quizBuFilter) : rows;
    const deptValues=[...new Set(deptScope.map(p=>p.Department))].sort((a,b)=>a.localeCompare(b));
    const bucketValues=[...new Set(rows.map(p=>p.ResponseBucket).filter(b=>b!==null&&b!==undefined))].sort((a,b)=>a-b);
    const prevBu=buSelect.value, prevDept=deptSelect.value, prevBucket=bucketSelect.value;
    buSelect.innerHTML='<option value="">ทั้งหมด</option>'+buValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    deptSelect.innerHTML='<option value="">ทั้งหมด</option>'+deptValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    bucketSelect.innerHTML='<option value="">ทั้งหมด</option>'+bucketValues.map(b=>`<option value="${b}">${b}–${b+15} นาที</option>`).join("");
    if (buValues.includes(prevBu)) buSelect.value=prevBu;
    if (deptValues.includes(prevDept)) {
      deptSelect.value=prevDept;
    } else if (prevDept) {
      deptSelect.value="";
      quizDeptFilter=null;
    }
    if (bucketValues.some(b=>String(b)===prevBucket)) bucketSelect.value=prevBucket;
  }
  // Renders the Quiz Answer Report table: currentQuizRows filtered by all 5 of the tab's own
  // filters PLUS quizCorrectFilter (set by clicking a pie slice — not part of the filter bar, so
  // it's applied on top, same relationship as activeFilter/columnFilters on the Latest Responses
  // table).
  function filterQuizRows() {
    let list=filterQuizRowsForChart();
    if (quizCorrectFilter==="correct") list=list.filter(p=>p.Correct);
    else if (quizCorrectFilter==="wrong") list=list.filter(p=>!p.Correct);
    return list;
  }
  function renderQuizTable() {
    const list=filterQuizRows();
    const sorted=[...list].sort((a,b)=>new Date(b.Created||0)-new Date(a.Created||0));
    const total=sorted.length;
    const totalPages=Math.max(1, Math.ceil(total/PAGE_SIZE));
    if (quizPage>totalPages) quizPage=totalPages;
    if (quizPage<1) quizPage=1;
    const startIdx=(quizPage-1)*PAGE_SIZE;
    const pageRows=sorted.slice(startIdx, startIdx+PAGE_SIZE);

    const filterParts=[];
    if (quizBuFilter) filterParts.push(`BU: <strong>${escapeHtml(quizBuFilter)}</strong>`);
    if (quizDeptFilter) filterParts.push(`Department: <strong>${escapeHtml(quizDeptFilter)}</strong>`);
    if (quizStatusFilter) filterParts.push(`สถานะ: <strong>${quizStatusFilter==="responded"?"ตอบรับแล้ว":"ยังไม่ตอบรับ"}</strong>`);
    if (quizTypeFilter) filterParts.push(`Type: <strong>${quizTypeFilter==="management"?"Management":"Staff"}</strong>`);
    if (quizBucketFilter!==null) filterParts.push(`ช่วงเวลาตอบ: <strong>${quizBucketFilter}–${quizBucketFilter+15} นาที</strong>`);
    if (quizCorrectFilter) filterParts.push(`ตอบ: <strong>${quizCorrectFilter==="correct"?"ถูก":"ผิด"}</strong>`);
    const anyFilterActive=Boolean(quizBuFilter||quizDeptFilter||quizStatusFilter||quizTypeFilter||quizBucketFilter!==null||quizCorrectFilter);
    const filterNote=filterParts.length ? ` — กรองตาม ${filterParts.join(", ")}` : "";
    const rangeText=total ? `${startIdx+1}–${Math.min(startIdx+PAGE_SIZE,total)}` : "0";
    $("quizSummary").innerHTML=`แสดง ${rangeText} จาก ${total.toLocaleString("th-TH")} รายการ${filterNote}`;
    show("clearQuizAllFiltersButton", anyFilterActive);
    $("quizTable").innerHTML=pageRows.map(p=>{
      const created=p.Created?escapeHtml(new Date(p.Created).toLocaleString("th-TH")):'-';
      const typeLabel=p.Type==="management"?"Management":(p.Type==="staff"?"Staff":"-");
      const correctBadge=p.Correct?'<span class="badge safe">ถูก</span>':'<span class="badge pending">ผิด</span>';
      return `<tr><td>${escapeHtml(p.Email)||'-'}</td><td>${escapeHtml(p.RefID)||'-'}</td><td>${escapeHtml(p.BU)||'-'}</td><td>${escapeHtml(p.Department)||'-'}</td><td>${typeLabel}</td><td>${correctBadge}</td><td>${escapeHtml(p.DrillResponse)||'-'}</td><td>${created}</td></tr>`;
    }).join("") || `<tr><td colspan="8">${anyFilterActive?'ไม่พบข้อมูลที่ตรงกับตัวกรอง':'ไม่พบข้อมูล'}</td></tr>`;
    $("quizPageIndicator").textContent=`หน้า ${quizPage} / ${totalPages}`;
    $("quizPrevPageButton").disabled=quizPage<=1;
    $("quizNextPageButton").disabled=quizPage>=totalPages;
  }
  function exportQuizCsv() {
    const list=filterQuizRows();
    if (!list.length) { setMessage("ไม่มีข้อมูลผลการตอบคำถามตามตัวกรองปัจจุบัน", "error"); return; }
    const csvEscape=v=>{ const s=String(v ?? ""); return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s; };
    const headers=["Email","RefID","BU","Department","Type","ตอบ","DrillResponse","Created"];
    const lines=[headers.join(",")];
    list.forEach(p=>{ lines.push([clean(p.Email), clean(p.RefID), clean(p.BU), clean(p.Department), p.Type==="management"?"Management":(p.Type==="staff"?"Staff":""), p.Correct?"ถูก":"ผิด", clean(p.DrillResponse), p.Created?new Date(p.Created).toLocaleString("th-TH"):""].map(csvEscape).join(",")); });
    const blob=new Blob(["﻿"+lines.join("\r\n")], {type:"text/csv;charset=utf-8;"});
    downloadBlob(blob, `quiz-answer-report-${exportTimestamp()}.csv`);
  }
  // Text shown/matched for a row in a given Latest Responses column — shared by the per-column
  // filter inputs and by CSV/Excel export, so "what you filtered on" and "what you exported"
  // always agree.
  function rowColumnText(r, col) {
    switch (col) {
      case "id": return String(r.ID ?? "");
      case "refid": return clean(r.RefID);
      case "email": return clean(r.EMail);
      case "created": return r.Created ? new Date(r.Created).toLocaleString("th-TH") : "";
      case "status": return [clean(r.ResponseSafe), clean(r.HelpNote)].filter(Boolean).join(" · ");
      case "drill": return clean(r.DrillResponse);
      default: return "";
    }
  }
  // currentRows -> chart click filter (activeFilter) -> per-column text filters (columnFilters)
  // -> sorted newest-first. This is the full matching set (every page), used both to slice out
  // the current page for display and, unsliced, for CSV/Excel export.
  function getFilteredRows() {
    // "Pending Employer" (added 2026-09-24) has no underlying Member-list rows to filter —
    // it supplies its own pseudo-rows (activeFilter.rows) built from the Phone Book master
    // list instead of a predicate over currentRows. Every other slice/bar still uses predicate.
    let list = activeFilter
      ? (activeFilter.rows ? activeFilter.rows : currentRows.filter(activeFilter.predicate))
      : currentRows;
    Object.entries(columnFilters).forEach(([col, needle]) => {
      if (!needle) return;
      list = list.filter(r => rowColumnText(r, col).toLowerCase().includes(needle));
    });
    return [...list].sort((a,b)=>new Date(b.Created||0)-new Date(a.Created||0));
  }
  // Renders the Latest Responses table: applies activeFilter (set by clicking a slice of the
  // Response Status chart or a bar of the 15-minute histogram) and the per-column text filters,
  // then shows the current page (30 rows max — see PAGE_SIZE). Called on its own by chart click
  // handlers and column-filter/pagination controls so none of those ever destroy/rebuild the
  // charts themselves — only render() (RefID change / Refresh) does that.
  function renderTable() {
    const sorted=getFilteredRows();
    const total=sorted.length;
    const totalPages=Math.max(1, Math.ceil(total/PAGE_SIZE));
    if (currentPage>totalPages) currentPage=totalPages;
    if (currentPage<1) currentPage=1;
    const startIdx=(currentPage-1)*PAGE_SIZE;
    const pageRows=sorted.slice(startIdx, startIdx+PAGE_SIZE);

    const filterNote=activeFilter
      ? ` — กรองตามกราฟ: <strong>${escapeHtml(activeFilter.label)}</strong> <a href="#" id="clearTableFilter" style="color:#202B49;text-decoration:underline;">(ล้างตัวกรอง)</a>`
      : "";
    const rangeText=total ? `${startIdx+1}–${Math.min(startIdx+PAGE_SIZE,total)}` : "0";
    $("recordSummary").innerHTML=`แสดง ${rangeText} จาก ${total} รายการ${filterNote}`;
    const clearLink=$("clearTableFilter");
    if (clearLink) clearLink.addEventListener("click", e=>{ e.preventDefault(); activeFilter=null; currentPage=1; renderTable(); });

    $("responseTable").innerHTML=pageRows.map(r=>{
      const cls=isSafe(r)?"safe":(clean(r.ResponseSafe)||clean(r.HelpNote))?"responded":"pending";
      const safeNote=[clean(r.ResponseSafe),clean(r.HelpNote)].filter(Boolean).join(" · ")||'-';
      const created=r.Created?escapeHtml(new Date(r.Created).toLocaleString('th-TH')):'-';
      return `<tr><td>${escapeHtml(r.ID)}</td><td>${escapeHtml(r.RefID)||'-'}</td><td>${escapeHtml(r.EMail)||'-'}</td><td>${created}</td><td><span class="badge ${cls}">${escapeHtml(safeNote)}</span></td><td>${escapeHtml(r.DrillResponse)||'-'}</td></tr>`;
    }).join("") || `<tr><td colspan="6">${(activeFilter||Object.values(columnFilters).some(Boolean))?'ไม่พบข้อมูลที่ตรงกับตัวกรอง':'ไม่พบข้อมูล'}</td></tr>`;

    $("pageIndicator").textContent=`หน้า ${currentPage} / ${totalPages}`;
    $("prevPageButton").disabled=currentPage<=1;
    $("nextPageButton").disabled=currentPage>=totalPages;
    show("clearColumnFiltersButton", Object.values(columnFilters).some(Boolean));
  }
  function exportTimestamp() {
    const d=new Date(), p=n=>String(n).padStart(2,"0");
    return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
  }
  function downloadBlob(blob, filename) {
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url; a.download=filename; document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 1000);
  }
  // Export the FULL current match set (activeFilter + column filters applied, all pages) —
  // not just the 30 rows currently visible on screen.
  function exportCsv() {
    const rows=getFilteredRows();
    if (!rows.length) { setMessage("ไม่มีข้อมูลให้ Export ตามตัวกรองปัจจุบัน", "error"); return; }
    const csvEscape=v=>{ const s=String(v ?? ""); return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s; };
    const headers=["ID","RefID","Email","Created","ResponseSafe","HelpNote","DrillResponse"];
    const lines=[headers.join(",")];
    rows.forEach(r=>{
      lines.push([r.ID, clean(r.RefID), clean(r.EMail), r.Created?new Date(r.Created).toLocaleString("th-TH"):"", clean(r.ResponseSafe), clean(r.HelpNote), clean(r.DrillResponse)].map(csvEscape).join(","));
    });
    // Leading BOM so Excel opens the Thai text as UTF-8 instead of mangling it.
    const blob=new Blob(["﻿"+lines.join("\r\n")], {type:"text/csv;charset=utf-8;"});
    downloadBlob(blob, `call-tree-responses-${exportTimestamp()}.csv`);
  }
  function exportExcel() {
    if (typeof XLSX === "undefined") { setMessage("ไม่สามารถโหลดไลบรารี Export Excel ได้ กรุณาตรวจสอบการเชื่อมต่ออินเทอร์เน็ตแล้วลองใหม่", "error"); return; }
    const rows=getFilteredRows();
    if (!rows.length) { setMessage("ไม่มีข้อมูลให้ Export ตามตัวกรองปัจจุบัน", "error"); return; }
    const data=rows.map(r=>({
      ID:r.ID, RefID:clean(r.RefID), Email:clean(r.EMail),
      Created:r.Created?new Date(r.Created).toLocaleString("th-TH"):"",
      ResponseSafe:clean(r.ResponseSafe), HelpNote:clean(r.HelpNote), DrillResponse:clean(r.DrillResponse),
    }));
    const ws=XLSX.utils.json_to_sheet(data);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Responses");
    XLSX.writeFile(wb, `call-tree-responses-${exportTimestamp()}.xlsx`);
  }
  async function loadData() {
    show("loading",true); show("content",false); setMessage("");
    try {
      const token=await acquireToken();
      allRows=await readSharePointList(token);
      // Phone Book (Master List) is a separate SharePoint site/list from Member — load it in its
      // own try/catch so a problem there (e.g. no access to /sites/snet, or the list/columns
      // change) never breaks the Member dashboard itself; only the "ยังไม่ตอบ" KPI/chart/table
      // degrade (kpiMissing shows "–", the panel shows an empty state).
      try { allPhoneBook=await readPhoneBookList(token); }
      catch(pbError) {
        console.error(pbError);
        allPhoneBook=[];
        setMessage(`โหลด Phone Book (Master List) ไม่สำเร็จ: ${pbError.message||pbError} — ข้อมูล Response อื่นใช้งานได้ตามปกติ แต่ KPI/กราฟ/ตาราง "ยังไม่ตอบ" จะไม่แสดงผล`, "error");
      }
      fillRefFilter(allRows);
      resetAllFilters();
      render();
      show("content",true);
    }
    catch(error){ console.error(error); setMessage(error.message || "ไม่สามารถโหลดข้อมูลได้", "error"); }
    finally { show("loading",false); }
  }
  async function login() {
    if (!msalApp) { $("configMessage").textContent="ระบบยังไม่พร้อมเข้าสู่ระบบ (MSAL ยังไม่ถูกโหลด) กรุณารีเฟรชหน้านี้ หากยังไม่หาย ให้ตรวจสอบว่าเครือข่าย/Proxy บล็อก alcdn.msauth.net หรือไม่"; return; }
    try { const result=await msalApp.loginPopup({scopes:cfg.graphScopes,prompt:"select_account"}); msalApp.setActiveAccount(result.account); show("loginView",false); show("dashboardView",true); await loadData(); }
    catch(error){ console.error(error); $("configMessage").textContent=error.message || "เข้าสู่ระบบไม่สำเร็จ"; }
  }
  // Singha Estate CI branding (added 2026-09-27): the login page's brand-mark shows the company
  // logo from images/singhaestate-logo.svg (kept in its own folder, per the user's explicit
  // request, so it's easy to find/replace/publish separately from the code files) if that file
  // exists; otherwise it falls back to a plain "SE" initials mark in the brand colors so the page
  // never shows a broken-image icon. Wired via a real `error` event listener (not an inline
  // onerror="..." attribute) because the page's CSP (script-src 'self', no 'unsafe-inline')
  // would silently block an inline event-handler attribute.
  function initBrandLogo() {
    const img=$("brandLogoImg"), fallback=$("brandLogoFallback");
    if (!img || !fallback) return;
    img.addEventListener("error", () => {
      img.style.display="none";
      fallback.style.display="flex";
    }, {once:true});
  }
  async function init() {
    initBrandLogo();
    if (!validateConfig()) return;
    try {
      if (typeof msal === "undefined") throw new Error("ไม่สามารถโหลดไลบรารี Microsoft Sign-in (MSAL) ได้ กรุณาตรวจสอบการเชื่อมต่ออินเทอร์เน็ต, Proxy/Firewall ขององค์กร หรือ Ad-blocker ที่อาจบล็อก alcdn.msauth.net แล้วรีเฟรชหน้าใหม่");
      msalApp=new msal.PublicClientApplication({auth:{clientId:cfg.clientId,authority:`https://login.microsoftonline.com/${cfg.tenantId}`,redirectUri:cfg.redirectUri,postLogoutRedirectUri:cfg.redirectUri,navigateToLoginRequestUrl:false},cache:{cacheLocation:"sessionStorage"}});
      await msalApp.initialize();
      $("configMessage").textContent="";
      $("loginButton").disabled=false;
      const accounts=msalApp.getAllAccounts();
      if(accounts.length){msalApp.setActiveAccount(accounts[0]);show("loginView",false);show("dashboardView",true);await loadData();}
    } catch(error) {
      console.error(error);
      $("loginButton").disabled=true;
      $("configMessage").textContent=error.message || "เกิดข้อผิดพลาดระหว่างเริ่มต้นระบบเข้าสู่ระบบ";
    }
  }
  // Top-level menu: "Dashboard" (KPIs/status/histogram/Latest Responses) vs. "รายงานผู้ที่ยัง
  // ตอบรับ" (the responded-vs-not-responded BU chart + table, moved into its own menu item per
  // the user's request 2026-09-23). Both views read the same allRows/allPhoneBook and are kept
  // in sync by the same render() call regardless of which one is visible — switching tabs never
  // re-fetches or re-renders data, it only toggles which <div> is shown.
  function showView(view) {
    show("viewDashboard", view==="dashboard");
    show("viewReport", view==="report");
    show("viewHelp", view==="help");
    show("viewQuiz", view==="quiz");
    $("navDashboard").classList.toggle("active", view==="dashboard");
    $("navReport").classList.toggle("active", view==="report");
    $("navHelp").classList.toggle("active", view==="help");
    $("navQuiz").classList.toggle("active", view==="quiz");
    // Chart.js sizes a canvas from its container at creation time. missingChart/helpChart/
    // quizChart are all (re)built by render() even while their own tab is hidden (display:none →
    // 0×0), so each must be told to recalculate its size once its container actually becomes
    // visible, or it stays blank/tiny.
    if (view==="report" && missingChart) missingChart.resize();
    if (view==="help" && helpChart) helpChart.resize();
    if (view==="help" && helpReasonChart) helpReasonChart.resize();
    if (view==="quiz" && quizChart) quizChart.resize();
  }
  $("navDashboard").addEventListener("click", ()=>showView("dashboard"));
  $("navReport").addEventListener("click", ()=>showView("report"));
  $("navHelp").addEventListener("click", ()=>showView("help"));
  $("navQuiz").addEventListener("click", ()=>showView("quiz"));
  $("loginButton").addEventListener("click",login);
  $("refreshButton").addEventListener("click",loadData);
  $("logoutButton").addEventListener("click",()=>{ if(msalApp) msalApp.logoutPopup({mainWindowRedirectUri:cfg.redirectUri}); });
  $("refFilter").addEventListener("change", ()=>{ resetAllFilters(); render(); });
  document.querySelectorAll(".col-filter").forEach(input=>{
    input.addEventListener("input", ()=>{
      columnFilters[input.dataset.col]=input.value.trim().toLowerCase();
      currentPage=1;
      renderTable();
    });
  });
  $("clearColumnFiltersButton").addEventListener("click", ()=>{
    document.querySelectorAll(".col-filter").forEach(input=>{ input.value=""; });
    Object.keys(columnFilters).forEach(k=>{ columnFilters[k]=""; });
    currentPage=1;
    renderTable();
  });
  $("prevPageButton").addEventListener("click", ()=>{ currentPage=Math.max(1,currentPage-1); renderTable(); });
  $("nextPageButton").addEventListener("click", ()=>{ currentPage=currentPage+1; renderTable(); });
  $("exportCsvButton").addEventListener("click", exportCsv);
  $("exportExcelButton").addEventListener("click", exportExcel);
  $("exportMissingCsvButton").addEventListener("click", exportMissingCsv);
  // BU changes cascade: re-scope the Department dropdown to the new BU (clearing Department if it
  // no longer fits) and rebuild missingChart so its Department bars only cover this BU (2026-09-24).
  $("missingBuFilterSelect").addEventListener("change", e=>{
    missingBuFilter=e.target.value||null;
    missingPage=1;
    populateReportFilterSelects(currentReportRows);
    renderMissingChart();
    renderMissingTable();
  });
  $("missingDeptFilterSelect").addEventListener("change", e=>{ missingDeptFilter=e.target.value||null; missingPage=1; renderMissingTable(); });
  $("missingStatusFilter").addEventListener("change", e=>{ reportStatusFilter=e.target.value||null; missingPage=1; renderMissingTable(); });
  $("missingTypeFilterSelect").addEventListener("change", e=>{ reportTypeFilter=e.target.value||null; missingPage=1; renderMissingTable(); });
  $("missingBucketFilterSelect").addEventListener("change", e=>{ reportBucketFilter=e.target.value===""?null:Number(e.target.value); missingPage=1; renderMissingTable(); });
  $("clearMissingAllFiltersButton").addEventListener("click", ()=>{
    missingBuFilter=null; missingDeptFilter=null; reportStatusFilter=null; reportTypeFilter=null; reportBucketFilter=null; missingPage=1;
    ["missingBuFilterSelect","missingDeptFilterSelect","missingStatusFilter","missingTypeFilterSelect","missingBucketFilterSelect"].forEach(id=>{ $(id).value=""; });
    populateReportFilterSelects(currentReportRows);
    renderMissingChart();
    renderMissingTable();
  });
  $("missingPrevPageButton").addEventListener("click", ()=>{ missingPage=Math.max(1,missingPage-1); renderMissingTable(); });
  $("missingNextPageButton").addEventListener("click", ()=>{ missingPage=missingPage+1; renderMissingTable(); });
  // Dashboard's own filter bar (added 2026-09-25) — every change re-runs the FULL render() (not
  // a narrow renderX() like the Report tab's filters use) because these 5 filters touch every
  // KPI card and both charts, not just one table. resetAllFilters() is deliberately NOT called
  // here — picking one Dashboard filter must not clear the others, or the Report tab's filters.
  $("dashBuFilterSelect").addEventListener("change", e=>{ dashBuFilter=e.target.value||null; currentPage=1; render(); });
  $("dashDeptFilterSelect").addEventListener("change", e=>{ dashDeptFilter=e.target.value||null; currentPage=1; render(); });
  $("dashStatusFilterSelect").addEventListener("change", e=>{ dashStatusFilter=e.target.value||null; currentPage=1; render(); });
  $("dashTypeFilterSelect").addEventListener("change", e=>{ dashTypeFilter=e.target.value||null; currentPage=1; render(); });
  $("dashBucketFilterSelect").addEventListener("change", e=>{ dashBucketFilter=e.target.value===""?null:Number(e.target.value); currentPage=1; render(); });
  $("clearDashAllFiltersButton").addEventListener("click", ()=>{
    dashBuFilter=null; dashDeptFilter=null; dashStatusFilter=null; dashTypeFilter=null; dashBucketFilter=null;
    currentPage=1;
    ["dashBuFilterSelect","dashDeptFilterSelect","dashStatusFilterSelect","dashTypeFilterSelect","dashBucketFilterSelect"].forEach(id=>{ $(id).value=""; });
    render();
  });
  // "รายงานขอความช่วยเหลือ" tab's own filter bar (added 2026-09-27) — same "narrow renderX(), not
  // the full render()" pattern as the Report tab's own listeners (this tab has its own base data
  // already computed once per render(), so re-filtering it doesn't need a full re-render).
  $("helpBuFilterSelect").addEventListener("change", e=>{
    helpBuFilter=e.target.value||null;
    helpPage=1;
    populateHelpFilterSelects(currentHelpRows);
    renderHelpChart();
    renderHelpReasonChart();
    renderHelpKpi();
    renderHelpTable();
  });
  // Department narrows renderHelpChart's own axis (excluded there) but DOES narrow
  // renderHelpReasonChart (Reason isn't that chart's axis) — see baseHelpFilterList().
  $("helpDeptFilterSelect").addEventListener("change", e=>{ helpDeptFilter=e.target.value||null; helpPage=1; renderHelpReasonChart(); renderHelpKpi(); renderHelpTable(); });
  $("helpStatusFilterSelect").addEventListener("change", e=>{ helpStatusFilter=e.target.value||null; helpPage=1; renderHelpChart(); renderHelpReasonChart(); renderHelpKpi(); renderHelpTable(); });
  $("helpTypeFilterSelect").addEventListener("change", e=>{ helpTypeFilter=e.target.value||null; helpPage=1; renderHelpChart(); renderHelpReasonChart(); renderHelpKpi(); renderHelpTable(); });
  $("helpBucketFilterSelect").addEventListener("change", e=>{ helpBucketFilter=e.target.value===""?null:Number(e.target.value); helpPage=1; renderHelpChart(); renderHelpReasonChart(); renderHelpKpi(); renderHelpTable(); });
  $("clearHelpAllFiltersButton").addEventListener("click", ()=>{
    helpBuFilter=null; helpDeptFilter=null; helpStatusFilter=null; helpTypeFilter=null; helpBucketFilter=null; helpReasonFilter=null; helpPage=1;
    ["helpBuFilterSelect","helpDeptFilterSelect","helpStatusFilterSelect","helpTypeFilterSelect","helpBucketFilterSelect"].forEach(id=>{ $(id).value=""; });
    populateHelpFilterSelects(currentHelpRows);
    renderHelpChart();
    renderHelpReasonChart();
    renderHelpKpi();
    renderHelpTable();
  });
  $("helpPrevPageButton").addEventListener("click", ()=>{ helpPage=Math.max(1,helpPage-1); renderHelpTable(); });
  $("helpNextPageButton").addEventListener("click", ()=>{ helpPage=helpPage+1; renderHelpTable(); });
  // Sub-tab toggle (added 2026-09-27): "kpi" shows the KPI cards + 2 charts, "table" shows the full
  // undeduped table. Filters/base data are shared and unaffected by which sub-tab is active.
  function showHelpSubTab(tab) {
    helpSubTab=tab;
    show("helpSubViewKpi", tab==="kpi");
    show("helpSubViewTable", tab==="table");
    $("helpSubTabKpi").classList.toggle("active", tab==="kpi");
    $("helpSubTabTable").classList.toggle("active", tab==="table");
    // Chart.js sizes a canvas from its container at creation time — same 0×0-while-hidden gotcha as
    // showView()'s own chart .resize() calls — so re-size both charts whenever the KPI sub-tab
    // (which holds them) becomes visible again.
    if (tab==="kpi") {
      if (helpChart) helpChart.resize();
      if (helpReasonChart) helpReasonChart.resize();
    }
  }
  $("helpSubTabKpi").addEventListener("click", ()=>showHelpSubTab("kpi"));
  $("helpSubTabTable").addEventListener("click", ()=>showHelpSubTab("table"));
  $("exportHelpCsvButton").addEventListener("click", exportHelpCsv);
  // "รายงานผลการตอบคำถาม" tab's own filter bar (added 2026-09-27) — same pattern; every filter
  // (including BU/Department, unlike the Help tab's chart) narrows the pie itself, so each change
  // re-renders both the chart and the table.
  $("quizBuFilterSelect").addEventListener("change", e=>{
    quizBuFilter=e.target.value||null;
    quizPage=1;
    populateQuizFilterSelects(currentQuizRows);
    renderQuizChart();
    renderQuizTable();
  });
  $("quizDeptFilterSelect").addEventListener("change", e=>{ quizDeptFilter=e.target.value||null; quizPage=1; renderQuizChart(); renderQuizTable(); });
  $("quizStatusFilterSelect").addEventListener("change", e=>{ quizStatusFilter=e.target.value||null; quizPage=1; renderQuizChart(); renderQuizTable(); });
  $("quizTypeFilterSelect").addEventListener("change", e=>{ quizTypeFilter=e.target.value||null; quizPage=1; renderQuizChart(); renderQuizTable(); });
  $("quizBucketFilterSelect").addEventListener("change", e=>{ quizBucketFilter=e.target.value===""?null:Number(e.target.value); quizPage=1; renderQuizChart(); renderQuizTable(); });
  $("clearQuizAllFiltersButton").addEventListener("click", ()=>{
    quizBuFilter=null; quizDeptFilter=null; quizStatusFilter=null; quizTypeFilter=null; quizBucketFilter=null; quizCorrectFilter=null; quizPage=1;
    ["quizBuFilterSelect","quizDeptFilterSelect","quizStatusFilterSelect","quizTypeFilterSelect","quizBucketFilterSelect"].forEach(id=>{ $(id).value=""; });
    populateQuizFilterSelects(currentQuizRows);
    renderQuizChart();
    renderQuizTable();
  });
  $("quizPrevPageButton").addEventListener("click", ()=>{ quizPage=Math.max(1,quizPage-1); renderQuizTable(); });
  $("quizNextPageButton").addEventListener("click", ()=>{ quizPage=quizPage+1; renderQuizTable(); });
  $("exportQuizCsvButton").addEventListener("click", exportQuizCsv);
  window.addEventListener("DOMContentLoaded",init);
})();
