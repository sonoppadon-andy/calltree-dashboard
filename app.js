(() => {
  "use strict";
  const cfg = window.CALLTREE_CONFIG;
  let msalApp;
  let allRows = [];
  let statusChart;
  let drillChart;
  // Cross-linking charts -> table: currentRows/currentBucketRowIds are (re)filled on every
  // render() and read by the chart click handlers (which only call renderTable(), never the
  // full render(), so the charts themselves aren't destroyed/rebuilt on every click).
  let currentRows = [];
  let currentBucketRowIds = {};
  let activeFilter = null; // {key, label, predicate(row)=>bool} or null (set by a chart click)
  // Latest Responses table: per-column text filters, pagination (30 rows/page).
  const PAGE_SIZE = 30;
  let currentPage = 1;
  let columnFilters = {id:"", refid:"", email:"", created:"", status:"", drill:""};
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
    const ref=$("refFilter").value;
    const scoped=ref === "all" ? allRows : allRows.filter(r=>clean(r.RefID)===ref);
    // Exclude the Admin announcement row (has Mode/iMsg) from every metric and the table — only member responses stay.
    const rows=scoped.filter(r=>!isNotificationRow(r));
    console.info(`[CallTree Dashboard] RefID=${ref}: ทั้งหมด ${scoped.length} แถว, ตัดรายการแจ้งเหตุการณ์ของ Admin (มี Mode/iMsg) ออก ${scoped.length-rows.length} แถว, เหลือ ${rows.length} แถว`);
    if (scoped.length && rows.length===scoped.length) {
      console.warn("[CallTree Dashboard] ไม่มีแถวใดถูกตัดออกเลย — ถ้าคาดว่าควรมีรายการแจ้งเหตุการณ์ถูกกรองออก ให้ตรวจสอบค่าจริงของ Mode/iMsg ในแถวตัวอย่างนี้:", scoped[0]);
    }
    // Safe / Responded now count unique employees (by email), not raw rows — one member can
    // appear on multiple rows for the same RefID (re-submits, drill + safe-check, etc.).
    const emailOf=r=>clean(r.EMail).toLowerCase();
    const safeEmails=new Set(rows.filter(isSafe).map(emailOf).filter(Boolean));
    const respondedEmails=new Set(rows.filter(hasResponse).map(emailOf).filter(Boolean));
    const safe=safeEmails.size;
    const responded=respondedEmails.size;
    const pending=Math.max(rows.length-responded,0);
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
    $("kpiTotal").textContent=rows.length.toLocaleString("th-TH");
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
      if (!modeVal && info.ID) {
        console.warn(`[CallTree Dashboard] RefID=${ref}: อ่านค่า Mode จาก field "Mode" ไม่ได้ (ได้ค่าว่าง) ทั้งที่ควรมีข้อมูล — นี่คือ raw object ทั้งหมดที่ Graph API ส่งกลับมาสำหรับแถวประกาศ Admin (ID ${info.ID}) ลองหาว่าค่า "Drill" ที่คาดไว้อยู่ภายใต้ property ชื่ออะไรจริง ๆ:`, info);
      }
      $("modeSubtitle").textContent=`Mode: ${modeVal||'-'} · Created: ${createdVal||'-'} · iMsg: ${msgVal||'-'}`;
    }

    // Row-level predicates for each doughnut slice, reused by its click handler so the table
    // filter shows exactly the rows that make up that slice. "Safe"/"Pending" mirror the
    // definitions of the `safe`/`pending` numbers above; "Need help" mirrors `help` above
    // (which is already row-based, so clicking it filters to exactly `help` rows).
    const statusPredicates={
      safe: isSafe,
      pending: r=>!hasResponse(r),
      help: r=>Boolean(clean(r.HelpNote) || clean(r.ResponseSafe).toLowerCase()==="seehelpnote"),
    };
    const statusLabels={safe:"Safe", pending:"Pending", help:"Need help"};
    if(statusChart) statusChart.destroy();
    statusChart=new Chart($("statusChart"),{type:"doughnut",data:{labels:["Safe","Pending","Need help"],datasets:[{data:[safe,pending,help],backgroundColor:["#16845b","#e9a23b","#d64545"],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"}},onHover:(evt,elements)=>{ if(evt.native) evt.native.target.style.cursor=elements.length?"pointer":"default"; },onClick:(evt,elements)=>{
      if (!elements.length) return;
      const kinds=["safe","pending","help"];
      const kind=kinds[elements[0].index];
      const key=`status:${kind}`;
      activeFilter=(activeFilter && activeFilter.key===key) ? null : {key, label:statusLabels[kind], predicate:statusPredicates[kind]};
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

    currentRows=rows;
    renderTable();
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
    let list = activeFilter ? currentRows.filter(activeFilter.predicate) : currentRows;
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
    try { const token=await acquireToken(); allRows=await readSharePointList(token); fillRefFilter(allRows); render(); show("content",true); }
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
  window.addEventListener("DOMContentLoaded",init);
})();
