"use strict";
/* Casos Hemodinâmica — versão web (OneDrive via Microsoft Graph).
   Mesmo formato de dados do app local: <ROOT>/casos.json, <ROOT>/Casos_Hemodinamica.xlsx, <ROOT>/Imagens/<pasta do caso>/ */

const CFG = window.APP_CONFIG || {};
const PROCS = ["Cateterismo diagnóstico","Angioplastia coronária","TAVI","Valve-in-valve","Valvoplastia aórtica",
 "Valvoplastia mitral","Reparo mitral/tricúspide (TEER)","Oclusão de apêndice atrial esquerdo","Fechamento de CIA/FOP",
 "Cateterismo direito","Biópsia endomiocárdica","Intervenção periférica","Alcoolização septal","Pericardiocentese",
 "Suporte circulatório (BIA/Impella/ECMO)"];
const DATA_FILE = "casos.json", XLSX_FILE = "Casos_Hemodinamica.xlsx", IMG_DIR = "Imagens";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Files.ReadWrite"];
const SMALL = 4 * 1024 * 1024, CHUNK = 32 * 327680; // 10 MiB (múltiplo de 320 KiB)
const $ = s => document.querySelector(s);

let msalApp = null, account = null;
let db = {versao:1, casos:[], procedimentosExtras:[]}, pending = [], editing = null, busy = false;

/* ---------- Utilidades ---------- */
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const slug = s => s.normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^A-Za-z0-9]+/g,"-").replace(/^-|-$/g,"").slice(0,40);
const fmtDate = iso => iso ? iso.split("-").reverse().join("/") : "";
const today = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset()*6e4).toISOString().slice(0,10); };
const fmtSize = b => b > 1e9 ? (b/1e9).toFixed(1)+" GB" : b > 1e6 ? (b/1e6).toFixed(1)+" MB" : Math.max(1, Math.round(b/1e3))+" KB";
const allProcs = () => [...PROCS, ...(db.procedimentosExtras||[]).filter(p => !PROCS.includes(p))];
const sleep = ms => new Promise(r => setTimeout(r, ms));
function lsGet(k){ try { return localStorage.getItem(k); } catch(e){ return null; } }
function lsSet(k,v){ try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k,v); } catch(e){} }
const rootFolder = () => (lsGet("repo-hemo-root") || CFG.ROOT_FOLDER || "Casos Hemodinâmica").replace(/^\/+|\/+$/g,"");

let msgT;
function msg(text, type="ok", ms=6000){ const m=$("#msg"); m.className="msg show "+type; m.innerHTML=text; clearTimeout(msgT); if(ms) msgT=setTimeout(()=>m.className="msg", ms); window.scrollTo({top:0,behavior:"smooth"}); }

/* ---------- Login (MSAL) ---------- */
const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, "");

async function initAuth(){
  if(!CFG.CLIENT_ID || CFG.CLIENT_ID.startsWith("COLE")){ $("#cfgMissing").classList.remove("hidden"); $("#btnLogin").disabled = true; return; }
  msalApp = new msal.PublicClientApplication({
    auth: { clientId: CFG.CLIENT_ID, authority: "https://login.microsoftonline.com/consumers", redirectUri: redirectUri(), postLogoutRedirectUri: redirectUri() },
    cache: { cacheLocation: "sessionStorage" } // nada persiste depois de fechar a aba
  });
  await msalApp.initialize();
  try {
    const res = await msalApp.handleRedirectPromise();
    if(res && res.account) account = res.account;
  } catch(e){ msg("Falha no login: " + esc(e.message), "err", 0); }
  if(!account){ const a = msalApp.getAllAccounts(); if(a.length) account = a[0]; }
  if(account) start();
}
$("#btnLogin").onclick = () => msalApp.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
$("#btnLogout").onclick = logout;
function logout(){
  if(busy && !confirm("Há um envio em andamento. Sair mesmo assim?")) return;
  busy = false;
  msalApp.logoutRedirect({ account, postLogoutRedirectUri: redirectUri() });
}
async function token(){
  try { return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
  catch(e){
    if(e instanceof msal.InteractionRequiredAuthError){ await msalApp.acquireTokenRedirect({ scopes: SCOPES, account }); }
    throw e;
  }
}

/* ---------- Encerrar por inatividade ---------- */
let lastAct = Date.now();
["mousemove","keydown","touchstart","scroll","click"].forEach(ev => addEventListener(ev, () => lastAct = Date.now(), {passive:true}));
setInterval(() => { if(account && !busy && Date.now() - lastAct > (CFG.IDLE_MINUTES||30)*60000) logout(); }, 30000);

/* ---------- Microsoft Graph ---------- */
const enc = parts => parts.filter(Boolean).map(encodeURIComponent).join("/");
const itemPath = rel => `/me/drive/root:/${enc([...rootFolder().split("/"), ...rel.split("/")])}`;

async function g(path, opts={}){
  for(let i=0;;i++){
    const r = await fetch(GRAPH + path, { ...opts, headers: { Authorization: "Bearer " + await token(), ...(opts.headers||{}) } });
    if((r.status === 429 || r.status >= 500) && i < 3){ await sleep((+r.headers.get("Retry-After") || 2**i) * 1000); continue; }
    return r;
  }
}
async function gErr(r){ let t = ""; try { const j = await r.json(); t = j.error && j.error.message; } catch(e){} return new Error(`${r.status} ${t||r.statusText}`); }

async function loadDb(){
  const r = await g(itemPath(DATA_FILE));
  if(r.status === 404) return { data: { versao:1, casos:[], procedimentosExtras:[] }, etag: null };
  if(!r.ok) throw await gErr(r);
  const meta = await r.json();
  const c = await fetch(meta["@microsoft.graph.downloadUrl"], { cache: "no-store" });
  if(!c.ok) throw new Error("Não consegui baixar " + DATA_FILE);
  const data = Object.assign({ versao:1, casos:[], procedimentosExtras:[] }, await c.json());
  return { data, etag: meta.eTag };
}
// Lê a versão mais recente, aplica a alteração e grava; se outro computador gravou no meio-tempo, tenta de novo.
async function commit(mutate){
  for(let i = 0; i < 5; i++){
    const { data, etag } = await loadDb();
    const result = mutate(data);
    const r = await g(itemPath(DATA_FILE) + ":/content", {
      method: "PUT", headers: { "Content-Type": "application/json", ...(etag ? { "If-Match": etag } : {}) },
      body: JSON.stringify(data, null, 2)
    });
    if(r.ok){ db = data; return result; }
    if(r.status === 412 || r.status === 409){ await sleep(400 + Math.random()*600); continue; }
    if(r.status === 423) throw new Error("o arquivo casos.json está bloqueado por outro programa; tente em instantes");
    throw await gErr(r);
  }
  throw new Error("conflito ao gravar casos.json; tente novamente");
}
async function ensureFolder(rel){
  const parts = [...rootFolder().split("/"), ...rel.split("/")].filter(Boolean);
  let parent = [];
  for(const name of parts){
    const url = parent.length ? `/me/drive/root:/${enc(parent)}:/children` : `/me/drive/root/children`;
    const r = await g(url, { method:"POST", headers:{ "Content-Type":"application/json" },
      body: JSON.stringify({ name, folder:{}, "@microsoft.graph.conflictBehavior":"fail" }) });
    if(!r.ok && r.status !== 409) throw await gErr(r);
    parent.push(name);
  }
  const m = await g(`/me/drive/root:/${enc(parts)}`);
  return m.ok ? await m.json() : null;
}
async function uploadFile(rel, file, onBytes, mode="rename"){
  if(file.size <= SMALL){
    const r = await g(itemPath(rel) + `:/content?@microsoft.graph.conflictBehavior=${mode}`, { method:"PUT", body:file });
    if(!r.ok) throw await gErr(r);
    onBytes(file.size); return r.json();
  }
  const s = await g(itemPath(rel) + ":/createUploadSession", { method:"POST", headers:{ "Content-Type":"application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": mode } }) });
  if(!s.ok) throw await gErr(s);
  const { uploadUrl } = await s.json();
  let start = 0, res = null;
  while(start < file.size){
    const end = Math.min(start + CHUNK, file.size);
    for(let i = 0;; i++){
      try {
        res = await fetch(uploadUrl, { method:"PUT", headers:{ "Content-Range": `bytes ${start}-${end-1}/${file.size}` }, body: file.slice(start, end) });
        if(res.ok) break;
        if(res.status < 500 || i >= 4) throw new Error(`envio interrompido (${res.status})`);
      } catch(e){ if(i >= 4) throw e; }
      await sleep(1000 * 2**i);
    }
    onBytes(end - start); start = end;
  }
  return res.json();
}

/* ---------- Início ---------- */
async function start(){
  $("#vLogin").classList.add("hidden");
  $("#user").classList.remove("hidden"); $("#userName").textContent = account.name || account.username;
  $("#tabs").classList.remove("hidden"); $("#rootName").textContent = rootFolder();
  try { db = (await loadDb()).data; }
  catch(e){ msg("Não consegui ler os casos do OneDrive: " + esc(e.message), "err", 0); }
  fillProcSelects(); newCase(); renderList(); show("novo");
}
$("#btnRoot").onclick = async () => {
  const v = prompt("Caminho da pasta a partir da raiz do OneDrive:", rootFolder());
  if(v == null) return;
  lsSet("repo-hemo-root", v.trim() && v.trim() !== CFG.ROOT_FOLDER ? v.trim() : null);
  $("#rootName").textContent = rootFolder();
  await reload();
};
async function reload(){ try { db = (await loadDb()).data; fillProcSelects(); renderList(); msg("Casos recarregados do OneDrive."); } catch(e){ msg("Erro ao recarregar: " + esc(e.message), "err", 0); } }
$("#btnReload").onclick = reload;

/* ---------- Navegação ---------- */
document.querySelectorAll(".tab").forEach(t => t.onclick = () => show(t.dataset.v));
function show(v){
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("on", t.dataset.v === v));
  $("#vForm").classList.toggle("hidden", v !== "novo"); $("#vList").classList.toggle("hidden", v !== "lista");
  if(v === "novo" && editing) newCase();
  if(v === "lista") renderList();
}

/* ---------- Procedimentos ---------- */
function fillProcSelects(){
  $("#proc").innerHTML = '<option value="">Selecione…</option>' + allProcs().map(p => `<option>${esc(p)}</option>`).join("") + '<option value="__outro">Outro…</option>';
  $("#fProc").innerHTML = '<option value="">Todos os procedimentos</option>' + [...new Set(db.casos.map(c => c.procedimento))].sort().map(p => `<option>${esc(p)}</option>`).join("");
  $("#count").textContent = `(${db.casos.length})`;
}
$("#proc").onchange = () => $("#procOutroWrap").classList.toggle("hidden", $("#proc").value !== "__outro");

/* ---------- Arquivos ---------- */
function addFiles(list){ for(const it of list){ if(!pending.some(p => p.path === it.path && p.file.size === it.file.size)) pending.push(it); } renderFiles(); }
function renderFiles(){
  $("#fileList").innerHTML = pending.map((p,i) => `<li><span class="n" title="${esc(p.path)}">${esc(p.path)}</span><span class="s">${fmtSize(p.file.size)}</span><button type="button" class="x" data-i="${i}" title="Remover">✕</button></li>`).join("");
  const tot = pending.reduce((a,p) => a + p.file.size, 0);
  if(pending.length > 3) $("#fileList").insertAdjacentHTML("beforeend", `<li class="small">${pending.length} arquivos · ${fmtSize(tot)}</li>`);
}
$("#fileList").onclick = e => { if(e.target.dataset.i !== undefined){ pending.splice(+e.target.dataset.i, 1); renderFiles(); } };
$("#pickFiles").onclick = () => $("#inFiles").click();
$("#pickDir").onclick = () => $("#inDir").click();
$("#inFiles").onchange = e => { addFiles([...e.target.files].map(f => ({ file:f, path:f.name }))); e.target.value = ""; };
$("#inDir").onchange = e => { addFiles([...e.target.files].map(f => ({ file:f, path:f.webkitRelativePath || f.name }))); e.target.value = ""; };
const drop = $("#drop");
["dragenter","dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave","drop"].forEach(ev => drop.addEventListener(ev, () => drop.classList.remove("over")));
drop.addEventListener("drop", async e => {
  e.preventDefault();
  const entries = [...e.dataTransfer.items].map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
  const out = [];
  async function walk(en, prefix){
    if(en.isFile){ const f = await new Promise((r,j) => en.file(r,j)); out.push({ file:f, path:prefix+f.name }); }
    else { const rd = en.createReader(); let batch; do { batch = await new Promise((r,j) => rd.readEntries(r,j)); for(const c of batch) await walk(c, prefix+en.name+"/"); } while(batch.length); }
  }
  for(const en of entries) await walk(en, "");
  addFiles(out);
});

/* ---------- Formulário ---------- */
function newCase(){
  editing = null; pending = []; $("#form").reset(); renderFiles();
  $("#data").value = today(); $("#evoData").value = today();
  $("#formTitle").textContent = "Novo caso"; $("#formSub").textContent = "Preencha os campos e anexe as imagens. A planilha é atualizada ao salvar.";
  $("#btnSave").textContent = "Salvar caso"; $("#btnCancel").classList.add("hidden");
  $("#procOutroWrap").classList.add("hidden"); $("#evoList").innerHTML = ""; $("#existing").classList.add("hidden");
}
async function openCase(id){
  const c = db.casos.find(x => x.id === id); if(!c) return;
  show("novo"); editing = c; pending = []; renderFiles();
  $("#nome").value = c.nome; $("#idade").value = c.idade ?? ""; $("#prontuario").value = c.prontuario || ""; $("#data").value = c.data;
  fillProcSelects(); $("#proc").value = c.procedimento; $("#procOutroWrap").classList.add("hidden");
  $("#resumo").value = c.resumo || ""; $("#evoTexto").value = ""; $("#evoData").value = today();
  $("#evoList").innerHTML = (c.evolucoes||[]).map(e => `<div class="evo"><div class="dt">${fmtDate(e.data)}</div><div style="white-space:pre-wrap">${esc(e.texto)}</div></div>`).join("");
  $("#formTitle").textContent = c.id + " · " + c.nome;
  $("#formSub").textContent = "Edite os campos, acrescente imagens ou registre uma nova evolução.";
  $("#btnSave").textContent = "Salvar alterações"; $("#btnCancel").classList.remove("hidden");
  const ex = $("#existing"); ex.classList.remove("hidden"); ex.textContent = "Consultando a pasta…";
  try {
    const r = await g(itemPath(`${IMG_DIR}/${c.pasta}`));
    if(r.ok){ const m = await r.json(); ex.innerHTML = `Já na pasta <a href="${esc(m.webUrl)}" target="_blank" rel="noopener">${esc(IMG_DIR+"/"+c.pasta)}</a>: ${m.folder ? m.folder.childCount : 0} item(ns), ${fmtSize(m.size||0)}. Novos arquivos serão adicionados a ela.`; }
    else ex.textContent = `Pasta ${IMG_DIR}/${c.pasta} ainda sem arquivos.`;
  } catch(e){ ex.textContent = ""; }
}
$("#btnCancel").onclick = () => { newCase(); show("lista"); };

$("#form").addEventListener("submit", async e => {
  e.preventDefault();
  if(busy) return;
  let proc = $("#proc").value, novoProc = null;
  if(proc === "__outro"){ proc = $("#procOutro").value.trim(); if(!proc){ msg("Informe qual é o procedimento.", "warn"); return; } if(!allProcs().includes(proc)) novoProc = proc; }
  const now = new Date().toISOString();
  const fields = { nome:$("#nome").value.trim(), idade:$("#idade").value === "" ? null : +$("#idade").value, prontuario:$("#prontuario").value.trim(),
    data:$("#data").value, procedimento:proc, resumo:$("#resumo").value.trim() };
  const evo = $("#evoTexto").value.trim() ? { data:$("#evoData").value || today(), texto:$("#evoTexto").value.trim() } : null;
  const files = pending.slice(), editId = editing && editing.id;

  const btn = $("#btnSave"), prog = $("#prog"), bar = prog.firstElementChild, ptxt = $("#progTxt");
  busy = true; btn.disabled = true; btn.textContent = "Salvando…";
  let c;
  try {
    // 1) grava o registro (reserva o ID)
    c = await commit(data => {
      if(novoProc && !(data.procedimentosExtras ||= []).includes(novoProc)) data.procedimentosExtras.push(novoProc);
      let x;
      if(editId){
        x = data.casos.find(k => k.id === editId);
        if(!x) throw new Error("caso não encontrado no OneDrive (foi apagado?)");
        Object.assign(x, fields, { atualizado: now });
      } else {
        const n = data.casos.reduce((m,k) => Math.max(m, parseInt(String(k.id).split("-")[1]) || 0), data.ultimoNumero || 0) + 1;
        data.ultimoNumero = n; // números de casos excluídos não são reaproveitados
        const id = "HEMO-" + String(n).padStart(4, "0");
        x = { id, ...fields, pasta: `${id}_${fields.data}_${slug(proc)}`, evolucoes: [], criado: now, atualizado: now };
        data.casos.push(x);
      }
      if(evo) (x.evolucoes ||= []).push(evo);
      return JSON.parse(JSON.stringify(x));
    });
    // 2) pasta + arquivos
    const folder = await ensureFolder(`${IMG_DIR}/${c.pasta}`);
    if(files.length){
      prog.style.display = "block";
      const total = files.reduce((a,p) => a + p.file.size, 0) || 1; let done = 0;
      for(const [i,p] of files.entries()){
        ptxt.textContent = `Enviando ${i+1}/${files.length}: ${p.path} (${fmtSize(p.file.size)})`;
        await uploadFile(`${IMG_DIR}/${c.pasta}/${p.path}`, p.file, b => { done += b; bar.style.width = (100*done/total).toFixed(1) + "%"; });
        pending = pending.filter(q => q !== p);
      }
    }
    // 3) link da pasta para a planilha
    if(folder && folder.webUrl && c.pastaUrl !== folder.webUrl){
      await commit(data => { const x = data.casos.find(k => k.id === c.id); if(x) x.pastaUrl = folder.webUrl; });
    }
    // 4) planilha
    ptxt.textContent = "Atualizando planilha…";
    const ok = await writeXlsx();
    if(ok) msg(`<b>${c.id}</b> salvo${files.length ? ` com ${files.length} arquivo(s)` : ""}. Planilha atualizada.`);
    else msg(`<b>${c.id}</b> salvo, mas a planilha não pôde ser atualizada (provavelmente está aberta no Excel). Feche-a e clique em <b>Casos › Atualizar planilha</b>.`, "warn", 0);
    fillProcSelects();
    if(editId) openCase(c.id); else newCase();
  } catch(err){
    const left = pending.length;
    msg(c ? `<b>${c.id}</b> foi registrado, mas houve erro no envio: ${esc(err.message)}. ${left ? `Faltaram ${left} arquivo(s): abra o caso em <b>Casos</b> e anexe de novo.` : ""}`
          : `Erro ao salvar: ${esc(err.message)}. Nada foi gravado; tente novamente.`, "err", 0);
    if(c){ fillProcSelects(); if(!editId){ const keep = pending; await openCase(c.id); pending = keep; renderFiles(); } }
  } finally {
    busy = false; btn.disabled = false; btn.textContent = editing ? "Salvar alterações" : "Salvar caso";
    prog.style.display = "none"; bar.style.width = "0"; ptxt.textContent = "";
  }
});
addEventListener("beforeunload", e => { if(busy){ e.preventDefault(); e.returnValue = ""; } });

/* ---------- Planilha ---------- */
async function buildXlsx(){
  const wb = new ExcelJS.Workbook(); wb.creator = "Casos Hemodinâmica";
  const ws = wb.addWorksheet("Casos", { views: [{ state:"frozen", xSplit:2, ySplit:1 }] });
  ws.columns = [
    {header:"ID do caso",key:"id",width:12},{header:"Nome do paciente",key:"nome",width:28},{header:"Idade",key:"idade",width:7},
    {header:"Prontuário",key:"prontuario",width:14},{header:"Data do exame",key:"data",width:13,style:{numFmt:"dd/mm/yyyy"}},
    {header:"Tipo de procedimento",key:"procedimento",width:26},{header:"Resumo do caso",key:"resumo",width:60},
    {header:"Imagens (pasta)",key:"pasta",width:34},{header:"Evolução",key:"evolucao",width:55},
    {header:"Data da última evolução",key:"ultEvo",width:14,style:{numFmt:"dd/mm/yyyy"}}];
  const toDate = iso => { if(!iso) return null; const [y,m,d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m-1, d)); };
  for(const c of [...db.casos].sort((a,b) => String(a.id).localeCompare(String(b.id)))){
    const ev = c.evolucoes || [];
    const r = ws.addRow({ id:c.id, nome:c.nome, idade:c.idade, prontuario:c.prontuario, data:toDate(c.data), procedimento:c.procedimento,
      resumo:c.resumo, pasta:{ text:c.pasta, hyperlink: c.pastaUrl || `${IMG_DIR}/${c.pasta}` },
      evolucao: ev.map(e => `${fmtDate(e.data)}: ${e.texto}`).join("\n"), ultEvo: ev.length ? toDate(ev[ev.length-1].data) : null });
    r.getCell("prontuario").numFmt = "@";
    r.eachCell(cell => cell.alignment = { vertical:"top", wrapText:true });
    r.getCell("pasta").font = { color:{ argb:"FF1F5FBF" }, underline:true };
  }
  const h = ws.getRow(1); h.height = 30;
  h.eachCell(cell => { cell.font = { bold:true, color:{ argb:"FFFFFFFF" } }; cell.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FF8B1E2D" } }; cell.alignment = { vertical:"middle", wrapText:true }; });
  ws.autoFilter = { from:"A1", to:"J1" };
  return new Blob([await wb.xlsx.writeBuffer()], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
async function writeXlsx(){
  try { await uploadFile(XLSX_FILE, await buildXlsx(), () => {}, "replace"); return true; }
  catch(e){ console.error(e); return false; }
}
$("#btnXlsx").onclick = async () => { (await writeXlsx()) ? msg("Planilha atualizada.") : msg("Não consegui gravar a planilha. Ela está aberta no Excel? Feche e tente de novo.", "warn", 0); };

/* ---------- Lista ---------- */
function renderList(){
  const norm = s => String(s ?? "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  const q = norm($("#q").value), fp = $("#fProc").value;
  const items = [...db.casos].reverse().filter(c => (!fp || c.procedimento === fp) && (!q || norm([c.id,c.nome,c.prontuario,c.procedimento,c.resumo,...(c.evolucoes||[]).map(e => e.texto)].join(" ")).includes(q)));
  $("#list").innerHTML = items.length ? items.map(c => `<div class="item" data-id="${esc(c.id)}">
    <div><div class="id">${esc(c.id)}</div><div class="d">${fmtDate(c.data)}</div></div>
    <div><div class="t">${esc(c.nome)}${c.idade != null ? `, ${c.idade} a` : ""}</div><div class="d">${c.prontuario ? "Pront. " + esc(c.prontuario) + " · " : ""}${esc((c.resumo||"").slice(0,110))}${(c.resumo||"").length > 110 ? "…" : ""}</div></div>
    <div style="text-align:right"><span class="chip">${esc(c.procedimento)}</span>${(c.evolucoes||[]).length ? `<div class="d" style="margin-top:4px">${c.evolucoes.length} evolução(ões)</div>` : ""}</div>
    <button type="button" class="del" data-del="${esc(c.id)}" title="Excluir caso" aria-label="Excluir ${esc(c.id)}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6M10 11v6M14 11v6"/></svg></button>
  </div>`).join("") : `<div class="empty">${db.casos.length ? "Nenhum caso encontrado." : "Nenhum caso ainda. Cadastre o primeiro em “Novo caso”."}</div>`;
}
$("#q").oninput = renderList; $("#fProc").onchange = renderList;
$("#list").onclick = e => { const d = e.target.closest(".del"); if(d){ askDelete(d.dataset.del); return; } const it = e.target.closest(".item"); if(it) openCase(it.dataset.id); };

async function deleteCase(id, imgs){
  const c = await commit(data => {
    const i = data.casos.findIndex(k => k.id === id); if(i < 0) return null;
    data.ultimoNumero = Math.max(data.ultimoNumero || 0, ...data.casos.map(k => parseInt(String(k.id).split("-")[1]) || 0));
    return data.casos.splice(i, 1)[0];
  });
  if(!c) return { removed:false, xlsx:true };
  if(imgs && c.pasta){
    const r = await g(itemPath(`${IMG_DIR}/${c.pasta}`), { method:"DELETE" });
    if(!r.ok && r.status !== 404) throw new Error(`caso removido, mas a pasta de imagens não pôde ser apagada (${(await gErr(r)).message})`);
  }
  return { removed:true, xlsx: await writeXlsx() };
}

/* ---------- Excluir ---------- */
let delId = null;
function askDelete(id){
  const c = db.casos.find(x => x.id === id); if(!c) return;
  delId = id;
  $("#delTxt").innerHTML = `<b>${esc(c.id)}</b> · ${esc(c.nome)} — ${esc(c.procedimento)}, ${fmtDate(c.data)}.<br>O caso sai da lista e da planilha.`;
  $("#delImgs").checked = true; $("#delYes").disabled = false; $("#delYes").textContent = "Excluir";
  $("#delModal").classList.remove("hidden"); $("#delNo").focus();
}
function closeDelete(){ $("#delModal").classList.add("hidden"); delId = null; }
$("#delNo").onclick = closeDelete;
$("#delModal").addEventListener("click", e => { if(e.target.id === "delModal" && !$("#delYes").disabled) closeDelete(); });
addEventListener("keydown", e => { if(e.key === "Escape" && !$("#delModal").classList.contains("hidden") && !$("#delYes").disabled) closeDelete(); });
$("#delYes").onclick = async () => {
  const id = delId, imgs = $("#delImgs").checked, b = $("#delYes");
  b.disabled = true; b.textContent = "Excluindo…";
  try {
    const { removed, xlsx } = await deleteCase(id, imgs);
    if(editing && editing.id === id) newCase();
    msg(removed ? `<b>${esc(id)}</b> excluído.` + (xlsx ? "" : " A planilha não pôde ser atualizada (aberta no Excel?): use <b>Atualizar planilha</b> depois.") : `${esc(id)} já não existia.`, xlsx ? "ok" : "warn", xlsx ? 6000 : 0);
  } catch(err){ msg("Erro ao excluir: " + esc(err.message), "err", 0); }
  finally { closeDelete(); fillProcSelects(); renderList(); }
};

initAuth();
