(() => {
  "use strict";
  const cfg = window.CALLTREE_CONFIG;
  let msalApp;
  let allRows = [];
  let statusChart;
  let drillChart;
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
    const times=rows.filter(r=>r.Created&&r.ClickDateTime).map(r=>(new Date(r.ClickDateTime)-new Date(r.Created))/60000).filter(v=>Number.isFinite(v)&&v>=0);
    const avg=times.length?times.reduce((a,b)=>a+b,0)/times.length:0;
    $("kpiTotal").textContent=rows.length.toLocaleString("th-TH");
    $("kpiSafe").textContent=safe.toLocaleString("th-TH");
    $("kpiResponded").textContent=responded.toLocaleString("th-TH");
    $("kpiPending").textContent=pending.toLocaleString("th-TH");
    $("kpiAverage").textContent=`${avg.toFixed(1)} min`;

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

    const byRef={}; rows.forEach(r=>{const key=clean(r.RefID)||"Unknown";byRef[key]??={total:0,responded:0,safe:0};byRef[key].total++;if(hasResponse(r))byRef[key].responded++;if(isSafe(r))byRef[key].safe++;});
    if(statusChart) statusChart.destroy();
    statusChart=new Chart($("statusChart"),{type:"doughnut",data:{labels:["Safe","Pending","Need help"],datasets:[{data:[safe,pending,help],backgroundColor:["#16845b","#e9a23b","#d64545"],borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom"}}}});
    if(drillChart) drillChart.destroy();
    drillChart=new Chart($("drillChart"),{type:"bar",data:{labels:Object.keys(byRef),datasets:[{label:"Total",data:Object.values(byRef).map(x=>x.total),backgroundColor:"#94a3b8"},{label:"Responded",data:Object.values(byRef).map(x=>x.responded),backgroundColor:"#2563eb"},{label:"Safe",data:Object.values(byRef).map(x=>x.safe),backgroundColor:"#16845b"}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true,ticks:{precision:0}}}}});

    $("recordSummary").textContent=`แสดง ${Math.min(rows.length,20)} จาก ${rows.length} รายการ`;
    const latest=[...rows].sort((a,b)=>new Date(b.Created||0)-new Date(a.Created||0)).slice(0,20);
    $("responseTable").innerHTML=latest.map(r=>{
      const cls=isSafe(r)?"safe":(clean(r.ResponseSafe)||clean(r.HelpNote))?"responded":"pending";
      const safeNote=[clean(r.ResponseSafe),clean(r.HelpNote)].filter(Boolean).join(" · ")||'-';
      const created=r.Created?escapeHtml(new Date(r.Created).toLocaleString('th-TH')):'-';
      return `<tr><td>${escapeHtml(r.ID)}</td><td>${escapeHtml(r.RefID)||'-'}</td><td>${escapeHtml(r.EMail)||'-'}</td><td>${created}</td><td><span class="badge ${cls}">${escapeHtml(safeNote)}</span></td><td>${escapeHtml(r.DrillResponse)||'-'}</td></tr>`;
    }).join("") || '<tr><td colspan="6">ไม่พบข้อมูล</td></tr>';
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
  window.addEventListener("DOMContentLoaded",init);
})();
