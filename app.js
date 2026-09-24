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
  // render(). Four filters combine (AND) to filter the table: missingBuFilter (BU — driven by
  // BOTH the #missingBuFilterSelect dropdown and clicking a missingChart bar, added 2026-09-24
  // as the filter-bar section, kept in sync both ways), missingDeptFilter (Department dropdown,
  // added 2026-09-24), reportStatusFilter (ตอบ/ไม่ตอบ dropdown) and reportTypeFilter (Type
  // dropdown, added 2026-09-24 — see classifyJobType() below). missingPage (same PAGE_SIZE as
  // the Member table) paginates the result. All of these only call renderMissingTable() (never
  // render()), same pattern as the Member table's activeFilter/columnFilters/currentPage.
  let currentReportRows = [];
  let currentPhoneBookTotal = 0;
  let missingBuFilter = null;
  let missingDeptFilter = null;
  let reportStatusFilter = null; // null (all) | "responded" | "missing"
  let reportTypeFilter = null; // null (all) | "management" | "staff"
  let missingPage = 1;
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
  function render() {
    // Scope changed (RefID filter / refresh) — any chart-click table filter, column filters,
    // and page position from before no longer apply to the new data, so drop them all.
    activeFilter=null;
    columnFilters={id:"", refid:"", email:"", created:"", status:"", drill:""};
    document.querySelectorAll(".col-filter").forEach(input=>{ input.value=""; });
    currentPage=1;
    missingBuFilter=null;
    missingDeptFilter=null;
    reportStatusFilter=null;
    reportTypeFilter=null;
    missingPage=1;
    ["missingBuFilterSelect","missingDeptFilterSelect","missingStatusFilter","missingTypeFilterSelect"].forEach(id=>{
      const el=$(id); if (el) el.value="";
    });
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
    // Safe / Responded now count unique employees (by email), not raw rows — one member can
    // appear on multiple rows for the same RefID (re-submits, drill + safe-check, etc.).
    const emailOf=r=>clean(r.EMail).toLowerCase();
    const safeEmails=new Set(rows.filter(isSafe).map(emailOf).filter(Boolean));
    // "Responded" = every unique email with at least one member row — NOT filtered through
    // hasResponse() (ClickDateTime/ResponseSafe/DrillResponse). Bug fixed 2026-09-22: for some
    // Modes (e.g. "ActivateCallTree", confirmed via RefID 645 screenshot: 4 rows / 3 unique
    // emails, all with ResponseSafe and DrillResponse blank) those three fields are legitimately
    // blank on every member row for that Mode, so hasResponse() was always false and Responded
    // showed 0 even though the rows themselves ARE the responses. `rows` already excludes the
    // Admin/notification row (see isNotificationRow above), so anything left in it with an EMail
    // is by construction a real member response — same "the row's existence is the signal"
    // reasoning already used for the 15-minute histogram (which is why the histogram was already
    // showing the correct count of 3 for RefID 645 while this KPI wrongly showed 0).
    const respondedEmails=new Set(rows.map(emailOf).filter(Boolean));
    const safe=safeEmails.size;
    const responded=respondedEmails.size;
    // Master (Phone Book) list — moved up from further below (2026-09-24) because "Total
    // Records" and the Response Status doughnut's "Pending Employer" now both depend on it, per
    // the user's explicit request. Dedupe the Phone Book itself by email first (keep the first
    // row per email) in case the same person has more than one row there.
    const phoneBookByEmail={};
    allPhoneBook.forEach(p=>{
      const email=clean(p.Email).toLowerCase();
      if (!email || phoneBookByEmail[email]) return;
      phoneBookByEmail[email]=p;
    });
    const phoneBookEmails=Object.keys(phoneBookByEmail);
    // "ยังไม่แจ้งเหตุ" (has not reported at all) = every unique email in the master list that is
    // NOT in respondedEmails (already scoped to the current RefID filter, same as everything
    // else in this function).
    const missingRows=phoneBookEmails.filter(e=>!respondedEmails.has(e)).map(e=>phoneBookByEmail[e]);
    // "รายงานผู้ที่ยังไม่ตอบรับ/ตอบรับ" table (widened 2026-09-24, was missingRows only) — EVERY
    // master-list person, tagged with whether they're in respondedEmails, so the tab's own
    // ตอบ/ไม่ตอบ filter and status column can show either or both. `missingRows` above (the
    // not-yet-responded subset) is still computed as-is for kpiMissing/kpiTotal/pending/the
    // doughnut's "Pending Employer" — this is a separate, additional array just for this table.
    const allMasterRows=phoneBookEmails.map(e=>({...phoneBookByEmail[e], Responded:respondedEmails.has(e)}));
    currentReportRows=allMasterRows;
    currentPhoneBookTotal=phoneBookEmails.length;
    // Show "–" (not "0") when the Phone Book hasn't loaded at all, so a genuine "everyone
    // responded" (0 missing, Phone Book loaded fine) is never confused with "no master data".
    $("kpiMissing").textContent=phoneBookEmails.length ? missingRows.length.toLocaleString("th-TH") : "–";
    // "Pending" (fixed 2026-09-24, was Math.max(rows.length-responded,0) — mixed a row-count
    // with a people-count, which is why the kpiPending card was hidden in the first place, see
    // the KPI-semantics notes) = master employees who have not reported at all. Same set as
    // `missingRows` above — reusing it keeps the doughnut's "Pending Employer" slice, the
    // kpiPending KPI (still hidden), and the "รายงานผู้ที่ยังตอบรับ" tab all in agreement.
    const pending=missingRows.length;
    // Pseudo-rows so clicking "Pending Employer" can list people in the Latest Responses table
    // who have NO row there at all (they never submitted anything) — only Email is real; every
    // other column is blank ("-") since there is no response to show. `rowColumnText()` and
    // `renderTable()` already handle blank ResponseSafe/HelpNote/DrillResponse/Created/RefID
    // gracefully (existing "||'-'" fallbacks), so no changes were needed there.
    const pendingMasterRows=missingRows.map(p=>({
      ID:`PB-${p.Email}`, RefID:"", EMail:p.Email, Created:null, ResponseSafe:"", HelpNote:"", DrillResponse:"",
    }));
    const help=rows.filter(r=>clean(r.HelpNote) || clean(r.ResponseSafe).toLowerCase()==="seehelpnote").length;
    // Reference time for "how long did this member take to respond" is the Admin/notification
    // row's own Created (the moment the event was announced) for that RefID — NOT each member
    // row's ClickDateTime, which is unreliable/blank in this list. Each member response row's own
    // Created is when THAT member's response was recorded (see Latest Responses table's "Created"
    // column), so elapsed = memberRow.Created − adminRow.Created for the same RefID.
    const adminCreatedByRef={};
    scoped.forEach(r=>{
      if (!isNotificationRow(r) || !r.Created) return;
      const key=clean(r.RefID)||"Unknown";
      if (!adminCreatedByRef[key]) adminCreatedByRef[key]=new Date(r.Created);
    });
    // "% ผู้ที่ตอบผิด": compare each RefID's Admin announcement text (iMsg) against each member's
    // Drill Response. Counted per UNIQUE email — if a person answered more than once, only their
    // FIRST Drill Response (by earliest Created, among rows that actually have a Drill Response) is
    // used. A blank Drill Response means the member never answered the drill question, so they are
    // excluded from both the numerator and denominator (not counted as "wrong").
    const adminIMsgByRef={};
    scoped.forEach(r=>{
      if (!isNotificationRow(r)) return;
      const key=clean(r.RefID)||"Unknown";
      if (!(key in adminIMsgByRef)) adminIMsgByRef[key]=fieldText(r.iMsg).trim();
    });
    const firstDrillByRefEmail={};
    rows.forEach(r=>{
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
    // "Total Records" (changed 2026-09-24, was rows.length — a row count of Member responses)
    // now shows the Phone Book master headcount instead, per the user's explicit request. This
    // is a genuine meaning change: it used to answer "how many response rows exist", it now
    // answers "how many employees are in the master list" — "–" (not "0") if Phone Book hasn't
    // loaded, same reasoning as kpiMissing above.
    $("kpiTotal").textContent=phoneBookEmails.length ? phoneBookEmails.length.toLocaleString("th-TH") : "–";
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
    statusChart=new Chart($("statusChart"),{type:"doughnut",data:{labels:["Safe","Pending Employer","Need help"],datasets:[{data:[safe,pending,help],backgroundColor:["#16845b","#e9a23b","#d64545"],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
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

    // Response-time histogram: count UNIQUE emails, not rows. A member with multiple response
    // rows for the same RefID is counted once, using the EARLIEST of their own Created times as
    // their response moment — bucketed against the Admin row's Created for that RefID, in
    // 15-minute buckets from 0.
    const earliestByRefEmail={};
    rows.forEach(r=>{
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
    drillChart=new Chart($("drillChart"),{type:"bar",data:{labels:bucketLabels,datasets:[{label:"จำนวนผู้ตอบ",data:bucketData,backgroundColor:"#2563eb",borderRadius:4,maxBarThickness:56}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{title:items=>items[0].label,label:item=>`${item.parsed.y.toLocaleString('th-TH')} คน`}}},scales:{x:{title:{display:true,text:'นาทีหลังจาก Admin แจ้งเหตุ'},grid:{display:false}},y:{beginAtZero:true,ticks:{precision:0},title:{display:true,text:'จำนวนผู้ตอบ (คนไม่ซ้ำ)'}}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
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

    // Both series (responded vs. not-yet-responded) grouped by BU, so the two counts sit on the
    // same categories for easy comparison — per the user's explicit request (2026-09-23).
    // "Responded" here means "in the Phone Book AND has a response in the current RefID scope" —
    // a different (smaller-or-equal) population than the kpiResponded KPI, which doesn't require
    // Phone Book membership. A person who responded but isn't in the Phone Book master list is
    // invisible to this chart (no BU to group them by) — flag this to the user if the two totals
    // are ever compared and don't match.
    const respondedByBU={};
    const missingByBU={};
    phoneBookEmails.forEach(e=>{
      const bu=clean(phoneBookByEmail[e].BU)||"ไม่ระบุ BU";
      if (respondedEmails.has(e)) respondedByBU[bu]=(respondedByBU[bu]||0)+1;
      else missingByBU[bu]=(missingByBU[bu]||0)+1;
    });
    const buLabels=[...new Set([...Object.keys(respondedByBU), ...Object.keys(missingByBU)])]
      .sort((a,b)=>((respondedByBU[b]||0)+(missingByBU[b]||0)) - ((respondedByBU[a]||0)+(missingByBU[a]||0)));
    const respondedData=buLabels.map(b=>respondedByBU[b]||0);
    const missingData=buLabels.map(b=>missingByBU[b]||0);
    if(missingChart) missingChart.destroy();
    missingChart=new Chart($("missingChart"),{type:"bar",data:{labels:buLabels,datasets:[
      {label:"ตอบรับแล้ว",data:respondedData,backgroundColor:"#16845b",borderRadius:4,maxBarThickness:22},
      {label:"ยังไม่ตอบรับ",data:missingData,backgroundColor:"#0e7490",borderRadius:4,maxBarThickness:22},
    ]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"},tooltip:{callbacks:{label:item=>`${item.dataset.label}: ${item.parsed.x.toLocaleString('th-TH')} คน`}}},scales:{x:{beginAtZero:true,ticks:{precision:0}},y:{grid:{display:false}}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      // Both datasets share the same category (BU) labels, so any bar clicked in either series
      // resolves to the same BU via its category index — the table filter always shows that BU's
      // "ยังไม่ตอบรับ" list (clicking a "ตอบรับแล้ว" bar filters the not-yet-responded table to
      // that BU too, which can legitimately come back empty if everyone in that BU responded).
      const bu=buLabels[elements[0].index];
      missingBuFilter=(missingBuFilter===bu) ? null : bu;
      const buSelect=$("missingBuFilterSelect");
      if (buSelect) buSelect.value=missingBuFilter||"";
      missingPage=1;
      renderMissingTable();
    }}});

    populateReportFilterSelects(allMasterRows);
    currentRows=rows;
    renderTable();
    renderMissingTable();
  }
  // Populates the BU / Department filter-bar dropdowns (added 2026-09-24) from whatever's
  // actually in the current Phone Book data, preserving each dropdown's current selection if
  // it's still a valid option (same "preserve previous" pattern as fillRefFilter()). Blanks
  // are bucketed as "ไม่ระบุ BU"/"ไม่ระบุ Department" — matching how renderMissingTable()'s own
  // filtering already treats blanks, so a value shown here always matches something below.
  function populateReportFilterSelects(rows) {
    const buSelect=$("missingBuFilterSelect"), deptSelect=$("missingDeptFilterSelect");
    const buValues=[...new Set(rows.map(p=>clean(p.BU)||"ไม่ระบุ BU"))].sort((a,b)=>a.localeCompare(b));
    const deptValues=[...new Set(rows.map(p=>clean(p.Department)||"ไม่ระบุ Department"))].sort((a,b)=>a.localeCompare(b));
    const prevBu=buSelect.value, prevDept=deptSelect.value;
    buSelect.innerHTML='<option value="">ทั้งหมด</option>'+buValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    deptSelect.innerHTML='<option value="">ทั้งหมด</option>'+deptValues.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    if (buValues.includes(prevBu)) buSelect.value=prevBu;
    if (deptValues.includes(prevDept)) deptSelect.value=prevDept;
  }
  // Renders the "รายงานผู้ที่ยังไม่ตอบรับ/ตอบรับ" table: currentReportRows (EVERY master-list
  // person, tagged Responded true/false — widened 2026-09-24), filtered by missingBuFilter,
  // missingDeptFilter, reportStatusFilter and reportTypeFilter (the filter-bar's 4 dropdowns,
  // added 2026-09-24 — BU is also settable by clicking a missingChart bar, kept in sync both
  // ways), sorted by Email, paginated at PAGE_SIZE (30/page — same as the Member table, so a
  // large Master list is never "crammed" into one scrolling box). Called on its own by the
  // chart's onClick, the filter-bar dropdowns, the pagination buttons, and the clear-filter
  // link/button — never destroys/rebuilds missingChart.
  function renderMissingTable() {
    let list=currentReportRows;
    if (missingBuFilter) list=list.filter(p=>(clean(p.BU)||"ไม่ระบุ BU")===missingBuFilter);
    if (missingDeptFilter) list=list.filter(p=>(clean(p.Department)||"ไม่ระบุ Department")===missingDeptFilter);
    if (reportTypeFilter) list=list.filter(p=>classifyJobType(p.JobTitle)===reportTypeFilter);
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
    const anyFilterActive=Boolean(missingBuFilter||missingDeptFilter||reportStatusFilter||reportTypeFilter);
    const filterNote=filterParts.length ? ` — กรองตาม ${filterParts.join(", ")}` : "";
    const rangeText=total ? `${startIdx+1}–${Math.min(startIdx+PAGE_SIZE,total)}` : "0";
    $("missingSummary").innerHTML=currentPhoneBookTotal
      ? `แสดง ${rangeText} จาก ${total.toLocaleString("th-TH")} รายการ (Master ทั้งหมด ${currentPhoneBookTotal.toLocaleString("th-TH")} คน)${filterNote}`
      : `ยังไม่ได้โหลดข้อมูล Phone Book (Master List)`;
    show("clearMissingAllFiltersButton", anyFilterActive);
    $("missingTable").innerHTML=pageRows.map(p=>{
      const cls=p.Responded?"responded":"pending";
      const label=p.Responded?"ตอบรับแล้ว":"ยังไม่ตอบรับ";
      return `<tr><td>${escapeHtml(p.Email)||'-'}</td><td>${escapeHtml(p.BU)||'-'}</td><td>${escapeHtml(p.Division)||'-'}</td><td>${escapeHtml(p.JobTitle)||'-'}</td><td>${escapeHtml(p.Department)||'-'}</td><td><span class="badge ${cls}">${label}</span></td></tr>`;
    }).join("") || `<tr><td colspan="6">${currentPhoneBookTotal ? (anyFilterActive?'ไม่พบข้อมูลที่ตรงกับตัวกรอง':'ไม่พบข้อมูล') : '-'}</td></tr>`;

    $("missingPageIndicator").textContent=`หน้า ${missingPage} / ${totalPages}`;
    $("missingPrevPageButton").disabled=missingPage<=1;
    $("missingNextPageButton").disabled=missingPage>=totalPages;
  }
  // Exports the FULL current match set (all 4 filter-bar filters applied — BU, Department,
  // ตอบ/ไม่ตอบ, Type — ALL pages, not just the visible page) — same csvEscape/BOM/downloadBlob
  // pattern as exportCsv(). Includes the Status column so an "all" export can still tell
  // responded/not-responded apart.
  function exportMissingCsv() {
    let list=currentReportRows;
    if (missingBuFilter) list=list.filter(p=>(clean(p.BU)||"ไม่ระบุ BU")===missingBuFilter);
    if (missingDeptFilter) list=list.filter(p=>(clean(p.Department)||"ไม่ระบุ Department")===missingDeptFilter);
    if (reportTypeFilter) list=list.filter(p=>classifyJobType(p.JobTitle)===reportTypeFilter);
    if (reportStatusFilter==="responded") list=list.filter(p=>p.Responded);
    else if (reportStatusFilter==="missing") list=list.filter(p=>!p.Responded);
    if (!list.length) { setMessage("ไม่มีรายชื่อตามตัวกรองปัจจุบัน (หรือยังไม่ได้โหลด Phone Book)", "error"); return; }
    const csvEscape=v=>{ const s=String(v ?? ""); return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s; };
    const headers=["Email","BU","Division","JobTitle","Department","Status"];
    const lines=[headers.join(",")];
    list.forEach(p=>{ lines.push([clean(p.Email), clean(p.BU), clean(p.Division), clean(p.JobTitle), clean(p.Department), p.Responded?"ตอบรับแล้ว":"ยังไม่ตอบรับ"].map(csvEscape).join(",")); });
    const blob=new Blob(["﻿"+lines.join("\r\n")], {type:"text/csv;charset=utf-8;"});
    downloadBlob(blob, `phonebook-report-${exportTimestamp()}.csv`);
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
      ? ` — กรองตามกราฟ: <strong>${escapeHtml(activeFilter.label)}</strong> <a href="#" id="clearTableFilter" style="color:#2563eb;text-decoration:underline;">(ล้างตัวกรอง)</a>`
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
  async function init() {
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
    $("navDashboard").classList.toggle("active", view==="dashboard");
    $("navReport").classList.toggle("active", view==="report");
    // Chart.js sizes a canvas from its container at creation time. missingChart is (re)built by
    // render() even while viewReport is hidden (display:none → 0×0), so it must be told to
    // recalculate its size once the container actually becomes visible, or it stays blank/tiny.
    if (view==="report" && missingChart) missingChart.resize();
  }
  $("navDashboard").addEventListener("click", ()=>showView("dashboard"));
  $("navReport").addEventListener("click", ()=>showView("report"));
  $("loginButton").addEventListener("click",login);
  $("refreshButton").addEventListener("click",loadData);
  $("logoutButton").addEventListener("click",()=>{ if(msalApp) msalApp.logoutPopup({mainWindowRedirectUri:cfg.redirectUri}); });
  $("refFilter").addEventListener("change",render);
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
  $("missingBuFilterSelect").addEventListener("change", e=>{ missingBuFilter=e.target.value||null; missingPage=1; renderMissingTable(); });
  $("missingDeptFilterSelect").addEventListener("change", e=>{ missingDeptFilter=e.target.value||null; missingPage=1; renderMissingTable(); });
  $("missingStatusFilter").addEventListener("change", e=>{ reportStatusFilter=e.target.value||null; missingPage=1; renderMissingTable(); });
  $("missingTypeFilterSelect").addEventListener("change", e=>{ reportTypeFilter=e.target.value||null; missingPage=1; renderMissingTable(); });
  $("clearMissingAllFiltersButton").addEventListener("click", ()=>{
    missingBuFilter=null; missingDeptFilter=null; reportStatusFilter=null; reportTypeFilter=null; missingPage=1;
    ["missingBuFilterSelect","missingDeptFilterSelect","missingStatusFilter","missingTypeFilterSelect"].forEach(id=>{ $(id).value=""; });
    renderMissingTable();
  });
  $("missingPrevPageButton").addEventListener("click", ()=>{ missingPage=Math.max(1,missingPage-1); renderMissingTable(); });
  $("missingNextPageButton").addEventListener("click", ()=>{ missingPage=missingPage+1; renderMissingTable(); });
  window.addEventListener("DOMContentLoaded",init);
})();
