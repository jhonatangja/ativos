"use strict";

const COLUNAS = {
  codigo: "Número de pedido JMS",
  motorista: "Responsável pela entrega",
  cidade: "Cidade Destino",
  motivo: "Motivos dos pacotes problemáticos",
  cliente: "Destinatário",
  endereco: "Complemento",
  bairro: "Distrito destinatário",
  cep: "CEP destino",
};
const MODELO_PADRAO =
  "Olá {primeiro_nome}, tudo bem? Aqui é da entrega do seu pedido (código {codigo}). " +
  "Não conseguimos localizar o endereço informado ({endereco}, {bairro}, {cidade}). " +
  "Pode nos enviar o endereço correto, com rua, número e ponto de referência? " +
  "Assim reagendamos a entrega. Obrigado!";

const novoEstado = () => ({ pacotes: [], totais: {}, res: {}, feitos: {}, modelo: MODELO_PADRAO, arquivo: "", atual: null });
let estado = novoEstado();
let aba = "resumo";
let busca = { rodando: false, parar: false, feitos: 0, total: 0, erro: "" };
let abriuWhats = false;
let waTab = null;

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const norm = (s) => String(s ?? "").normalize("NFKD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* ---------- armazenamento (chrome.storage; localStorage só como reserva em testes) ---------- */
const store = {
  async ler() {
    try {
      if (globalThis.chrome?.storage?.local) return (await chrome.storage.local.get("estado")).estado;
      return JSON.parse(localStorage.getItem("ativo_estado") || "null");
    } catch (e) { return null; }
  },
  async gravar(v) {
    try {
      if (globalThis.chrome?.storage?.local) await chrome.storage.local.set({ estado: v });
      else localStorage.setItem("ativo_estado", JSON.stringify(v));
    } catch (e) { /* ignora */ }
  },
};
let salvarT;
function salvar() { clearTimeout(salvarT); salvarT = setTimeout(() => store.gravar(estado), 150); }

function toast(msg) {
  const t = $("toast");
  t.textContent = msg; t.classList.add("show");
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove("show"), 1500);
}
async function copiar(txt, msg = "Copiado") {
  try { await navigator.clipboard.writeText(txt); }
  catch (e) {
    const ta = document.createElement("textarea");
    ta.value = txt; document.body.appendChild(ta); ta.select();
    document.execCommand("copy"); ta.remove();
  }
  toast(msg);
}

/* ---------- telefone ---------- */
function telNacional(bruto) {
  let d = String(bruto ?? "").replace(/\D/g, "");
  if (d.startsWith("55") && d.length >= 12) d = d.slice(2);
  return d;
}
function telBonito(bruto) {
  const d = telNacional(bruto);
  if (d.length === 11) return `(${d.slice(0, 2)}) ${d.slice(2, 7)}-${d.slice(7)}`;
  if (d.length === 10) return `(${d.slice(0, 2)}) ${d.slice(2, 6)}-${d.slice(6)}`;
  return String(bruto ?? "");
}
function linkWhats(bruto, texto) {
  const d = telNacional(bruto);
  return d.length === 10 || d.length === 11 ? `https://wa.me/55${d}?text=${encodeURIComponent(texto)}` : "";
}

/* ---------- planilha ---------- */
const limpaMotorista = (n) => String(n || "").replace(/^\s*F\s+[A-Z]{2,4}\s*-\s*/, "").trim() || "(sem motorista)";

async function lerArquivo(file) {
  const wb = /\.(csv|txt)$/i.test(file.name)
    ? XLSX.read(await file.text(), { type: "string" })
    : XLSX.read(await file.arrayBuffer(), { type: "array" });
  const linhas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });
  if (!linhas.length) throw new Error("A planilha está vazia.");
  const faltam = Object.values(COLUNAS).filter((c) => !(c in linhas[0]));
  if (faltam.length) throw new Error("Esta não parece a planilha do JMS. Colunas que faltam: " + faltam.join(", "));

  const totais = {}, pacotes = [];
  for (const l of linhas) {
    const mot = limpaMotorista(l[COLUNAS.motorista]);
    totais[mot] = (totais[mot] || 0) + 1;
    if (!norm(l[COLUNAS.motivo]).includes("endereco incorreto")) continue;
    const codigo = String(l[COLUNAS.codigo]).replace(/\D/g, "");
    if (!codigo) continue;
    pacotes.push({
      codigo, motorista: mot,
      cliente: String(l[COLUNAS.cliente]).trim(), endereco: String(l[COLUNAS.endereco]).trim(),
      bairro: String(l[COLUNAS.bairro]).trim(), cidade: String(l[COLUNAS.cidade]).trim(), cep: String(l[COLUNAS.cep]).trim(),
    });
  }
  const codigos = new Set(pacotes.map((p) => p.codigo));
  const manter = (o) => Object.fromEntries(Object.entries(o).filter(([k]) => codigos.has(k)));
  estado = { ...estado, pacotes, totais, res: manter(estado.res), feitos: manter(estado.feitos), arquivo: file.name, atual: null };
  salvar();
}

/* ---------- dados derivados ---------- */
function mensagem(p) {
  const pn = p.cliente.split(/\s+/)[0] || "";
  const v = { primeiro_nome: pn ? pn[0].toUpperCase() + pn.slice(1).toLowerCase() : "", codigo: p.codigo, endereco: p.endereco, bairro: p.bairro, cidade: p.cidade };
  return estado.modelo.replace(/\{(\w+)\}/g, (_, k) => v[k] ?? "");
}

function derivar() {
  const donos = {};
  for (const p of estado.pacotes) {
    const r = estado.res[p.codigo];
    if (r && r.s === "ok") (donos[telNacional(r.t)] = donos[telNacional(r.t)] || new Set()).add(norm(p.cliente));
  }
  return estado.pacotes.map((p) => {
    const r = estado.res[p.codigo];
    let status = "pend", aviso = "", motivoErro = "";
    const d = r && r.s === "ok" ? telNacional(r.t) : "";
    if (r && r.s === "ok") {
      if (d.length !== 10 && d.length !== 11) { status = "sem"; motivoErro = "o JMS não tem telefone cadastrado"; }
      else {
        status = "ok";
        if (d.length === 10) aviso = "Telefone com 10 dígitos (pode faltar o 9). O WhatsApp pode não abrir.";
        if (donos[d].size > 1) aviso += (aviso ? " " : "") + "Mesmo telefone aparece em clientes diferentes; confira antes de enviar.";
      }
    } else if (r) { status = "erro"; motivoErro = r.e || "erro"; }
    const msg = mensagem(p);
    return { ...p, status, aviso, motivoErro, tel: status === "ok" ? telBonito(r.t) : "", telCru: status === "ok" ? telNacional(r.t) : "",
      mensagem: msg, whats: status === "ok" ? linkWhats(r.t, msg) : "" };
  });
}

function rankingDe(linhas) {
  const m = {};
  for (const p of linhas) m[p.motorista] = (m[p.motorista] || 0) + 1;
  return Object.entries(m)
    .map(([motorista, qtd]) => ({ motorista, qtd, total: estado.totais[motorista] || qtd, pct: Math.round((qtd / (estado.totais[motorista] || qtd)) * 1000) / 10 }))
    .sort((a, b) => b.qtd - a.qtd || b.pct - a.pct)
    .map((r, i) => ({ pos: i + 1, ...r }));
}

function contexto() {
  const linhas = derivar();
  const rank = rankingDe(linhas);
  const pos = Object.fromEntries(rank.map((r) => [r.motorista, r.pos]));
  const fila = linhas.filter((l) => l.status === "ok" && !estado.feitos[l.codigo]).sort((a, b) => pos[a.motorista] - pos[b.motorista]);
  const paraBuscar = linhas.filter((l) => l.status === "pend" || l.status === "erro");
  return { linhas, rank, pos, fila, paraBuscar,
    ok: linhas.filter((l) => l.status === "ok"), sem: linhas.filter((l) => l.status === "sem") };
}

/* ---------- telas ---------- */
function render() {
  const c = contexto();
  document.querySelectorAll("#abas button").forEach((b) => b.classList.toggle("on", b.dataset.aba === aba));
  const bg = $("badgeFila");
  bg.hidden = !c.fila.length; bg.textContent = c.fila.length;
  $("tela").innerHTML = aba === "resumo" ? telaResumo(c) : aba === "atender" ? telaAtender(c) : telaMais(c);
}

function blocoUpload(titulo, sub) {
  return `<label class="upload"><b>${titulo}</b><span class="mut">${sub}</span>
    <input type="file" id="arquivo" accept=".xlsx,.xls,.csv" /></label><div id="erroArquivo"></div>`;
}

function telaResumo(c) {
  if (!estado.pacotes.length) {
    return `<div class="card"><h2>Comece por aqui</h2>
      <p class="mut">Escolha a planilha que você exporta do JMS. Ela é lida só neste computador.</p>
      ${blocoUpload("📄 Escolher planilha", "clique para selecionar o arquivo .xlsx")}</div>
      <div class="card"><h3>Como funciona</h3>
      <p>1. Você escolhe a planilha.</p><p>2. O programa busca sozinho o telefone de cada cliente no JMS.</p>
      <p>3. Você atende um cliente por vez, com a mensagem pronta e o WhatsApp a um clique.</p></div>`;
  }
  const falta = c.paraBuscar.length;
  const min = Math.max(1, Math.round(falta * 4 / 60));
  let busca_ui;
  if (busca.rodando) {
    const pct = busca.total ? Math.round(busca.feitos / busca.total * 100) : 0;
    busca_ui = `<h2>Buscando telefones…</h2>
      <div class="prog"><i style="width:${pct}%"></i></div>
      <p>${busca.feitos} de ${busca.total} <span class="mut">· não troque de aba nem feche o JMS</span></p>
      <button class="btn sec" data-a="parar">Parar</button>`;
  } else if (falta) {
    busca_ui = `<h2>Buscar telefones no JMS</h2>
      <p class="mut">${falta} cliente${falta > 1 ? "s" : ""} sem telefone ainda. Leva cerca de ${min} min. Uso o JMS que você já tem aberto.</p>
      <button class="btn" data-a="buscar">🔎 Buscar telefones (${falta})</button>
      ${c.ok.length ? `<div class="gap"></div><button class="btn sec" data-a="ir-atender">Já atender os ${c.fila.length} que tenho</button>` : ""}`;
  } else {
    busca_ui = `<h2>✓ Telefones prontos</h2>
      <p class="mut">${c.ok.length} com telefone${c.sem.length ? `, ${c.sem.length} sem número` : ""}.</p>
      <button class="btn wa" data-a="ir-atender">Começar atendimento →</button>`;
  }
  const max = Math.max(1, ...c.rank.map((r) => r.qtd));
  const linhaRank = (r) => `<tr><td class="pos">${r.pos}</td><td>${esc(r.motorista)}<div class="bar"><i style="width:${r.qtd / max * 100}%"></i></div></td>
      <td style="text-align:right;white-space:nowrap"><b>${r.qtd}</b> <span class="mut">${r.pct}%</span></td></tr>`;
  return `${busca.erro ? `<div class="erro">${esc(busca.erro)}</div>` : ""}
    <div class="grid">
      <div class="stat"><b>${c.linhas.length}</b><span>endereço incorreto</span></div>
      <div class="stat"><b class="ok">${c.ok.length}</b><span>com telefone</span></div>
      <div class="stat"><b>${falta}</b><span>falta buscar</span></div>
      <div class="stat"><b class="${c.sem.length ? "er" : ""}">${c.sem.length}</b><span>sem número</span></div>
    </div>
    <div class="card">${busca_ui}</div>
    <div class="card"><h3>Maiores ofensores</h3>
      <table class="rank">${c.rank.slice(0, 3).map(linhaRank).join("")}</table>
      ${c.rank.length > 3 ? `<details><summary>Ver todos os ${c.rank.length} motoristas</summary><table class="rank">${c.rank.slice(3).map(linhaRank).join("")}</table></details>` : ""}
      <p class="mut" style="margin-top:8px;font-size:12px">% = parte dos pacotes do motorista que teve endereço incorreto.</p>
    </div>
    <div class="mut" style="font-size:12px;text-align:center">${esc(estado.arquivo)}</div>`;
}

function telaAtender(c) {
  if (!estado.pacotes.length) return `<div class="vazio"><div class="g">📄</div><p>Escolha a planilha primeiro.</p><button class="btn" data-a="ir-resumo">Ir para o Resumo</button></div>`;
  const feitos = c.ok.length - c.fila.length;
  if (!c.fila.length) {
    if (!c.ok.length) return `<div class="vazio"><div class="g">🔎</div><p><b>Ainda não tem telefone para atender.</b></p><p class="mut">Busque os telefones no Resumo.</p><button class="btn" data-a="ir-resumo">Ir para o Resumo</button></div>`;
    const resto = [];
    if (c.paraBuscar.length) resto.push(`${c.paraBuscar.length} ainda sem telefone buscado ou com erro`);
    if (c.sem.length) resto.push(`${c.sem.length} sem número no JMS`);
    return `<div class="vazio"><div class="g">🎉</div><p><b>Tudo atendido!</b></p>
      <p class="mut">${feitos} cliente${feitos !== 1 ? "s" : ""} marcado${feitos !== 1 ? "s" : ""} como enviado.</p>
      ${resto.length ? `<p class="mut">Sobraram: ${resto.join(" · ")}. Veja em "Mais".</p>` : ""}
      ${c.paraBuscar.length ? `<button class="btn sec" data-a="ir-resumo">Buscar de novo no Resumo</button>` : ""}</div>`;
  }
  let i = c.fila.findIndex((l) => l.codigo === estado.atual);
  if (i < 0) i = 0;
  const p = c.fila[i];
  estado.atual = p.codigo;
  const pct = c.ok.length ? Math.round(feitos / c.ok.length * 100) : 0;
  const r = c.rank.find((x) => x.motorista === p.motorista);
  return `<div class="prog"><i style="width:${pct}%"></i></div>
    <div class="nav"><button data-a="ant" ${c.fila.length < 2 ? "disabled" : ""}>← Anterior</button>
      <span>Cliente ${i + 1} de ${c.fila.length} · ${feitos} enviados</span>
      <button data-a="pular" ${c.fila.length < 2 ? "disabled" : ""}>Pular →</button></div>
    <div class="card">
      <span class="pill">Motorista #${r.pos}: ${esc(p.motorista)}</span>
      <div class="cli">${esc(p.cliente)}</div>
      <div class="tel">${esc(p.tel)}</div>
      <div class="end">${esc(p.endereco)}<br><span class="mut">${esc(p.bairro)} · ${esc(p.cidade)} · CEP ${esc(p.cep)}</span></div>
      ${p.aviso ? `<div class="aviso">⚠ ${esc(p.aviso)}</div>` : ""}
      <div class="msg">${esc(p.mensagem)}</div>
      <button class="btn wa" data-a="whats" ${p.whats ? "" : "disabled"}>💬 Abrir WhatsApp com a mensagem</button>
      <div class="gap"></div>
      <div class="linha">
        <button class="btn sec peq" data-a="cp-msg">Copiar mensagem</button>
        <button class="btn sec peq" data-a="cp-tel">Copiar telefone</button>
      </div>
    </div>
    <button class="btn ${abriuWhats ? "" : "sec"}" data-a="feito">✓ Enviado — próximo</button>`;
}

function telaMais(c) {
  const lista = [...c.sem, ...c.linhas.filter((l) => l.status === "erro")];
  return `<div class="card"><h3>Exportar</h3>
      <button class="btn" data-a="excel" ${estado.pacotes.length ? "" : "disabled"}>📥 Baixar Excel (ranking + lista por motorista)</button></div>
    <div class="card"><h3>Mensagem para o cliente</h3>
      <p class="mut" style="font-size:12px">Variáveis: <code>{primeiro_nome}</code> <code>{codigo}</code> <code>{endereco}</code> <code>{bairro}</code> <code>{cidade}</code></p>
      <textarea id="modelo">${esc(estado.modelo)}</textarea>
      <div class="gap"></div><button class="btn sec peq" data-a="modelo-padrao">Voltar ao texto padrão</button></div>
    <div class="card"><h3>Sem número / com erro (${lista.length})</h3>
      ${lista.length ? lista.map((l) => `<div class="item"><b>${esc(l.cliente)}</b> <span class="mut">· ${esc(l.motorista)}</span><br>
        <span class="mut">${esc(l.codigo)} — ${esc(l.motivoErro)}</span><br><button class="btn sec peq" style="display:inline-block;width:auto;margin-top:4px" data-a="cp-cod" data-c="${esc(l.codigo)}">Copiar código</button></div>`).join("") : '<p class="mut">Nenhum.</p>'}</div>
    <div class="card"><h3>Planilha</h3>
      ${estado.pacotes.length ? blocoUpload("Trocar planilha", "os telefones já buscados são mantidos") : "<p class=\"mut\">Nenhuma planilha carregada.</p>"}
      <div class="gap"></div><button class="btn perigo peq" data-a="limpar">Apagar tudo e recomeçar</button></div>
    <p class="mut" style="text-align:center;font-size:12px">Ativo JMS 1.0 · dados só neste computador</p>`;
}

/* ---------- WhatsApp ---------- */
async function abrirWhats(url) {
  try {
    if (waTab != null) { await chrome.tabs.update(waTab, { url, active: true }); return; }
  } catch (e) { waTab = null; }
  try { waTab = (await chrome.tabs.create({ url })).id; }
  catch (e) { window.open(url, "_blank"); }
}
if (globalThis.chrome?.tabs?.onRemoved) chrome.tabs.onRemoved.addListener((id) => { if (id === waTab) waTab = null; });

/* ---------- busca no JMS ---------- */
async function noJms(tabId, func, args = []) {
  const r = await chrome.scripting.executeScript({ target: { tabId }, world: "MAIN", func, args });
  return r && r[0] ? r[0].result : undefined;
}
async function acharAbaJms() {
  const tabs = await chrome.tabs.query({ lastFocusedWindow: true });
  const jms = tabs.filter((t) => /^https:\/\/[^/]*jtjms-br\.com\//.test(t.url || ""));
  return jms.find((t) => t.active) || jms[0] || null;
}
function esperaCarregar(id) {
  return new Promise((res) => {
    const f = (tid, info) => { if (tid === id && info.status === "complete") { chrome.tabs.onUpdated.removeListener(f); res(); } };
    chrome.tabs.onUpdated.addListener(f);
    setTimeout(() => { chrome.tabs.onUpdated.removeListener(f); res(); }, 20000);
  });
}

async function buscarTelefones() {
  busca = { rodando: false, parar: false, feitos: 0, total: 0, erro: "" };
  const c = contexto();
  const lista = c.paraBuscar.map((l) => l.codigo);
  if (!lista.length) return;
  try {
    const tab = await acharAbaJms();
    if (!tab) { busca.erro = "Não achei o JMS aberto nesta janela do Chrome. Abra o JMS, faça o login e clique em Buscar de novo."; return render(); }
    if (!/trackingExpress/.test(tab.url)) {
      await chrome.tabs.update(tab.id, { url: JMS_URL_RASTREIO, active: true });
      await esperaCarregar(tab.id);
    } else if (!tab.active) await chrome.tabs.update(tab.id, { active: true });
    let pronta = false;
    for (let i = 0; i < 40 && !pronta; i++) { pronta = await noJms(tab.id, jmsPaginaPronta).catch(() => false); if (!pronta) await sleep(500); }
    if (!pronta) { busca.erro = "Não consegui abrir o Rastreamento do pacote. Confira se você está logado no JMS e tente de novo."; return render(); }

    busca.rodando = true; busca.total = lista.length; render();
    for (const codigo of lista) {
      if (busca.parar) break;
      let r;
      for (let t = 0; t < 3; t++) {
        try { r = await noJms(tab.id, jmsConsultarCodigo, [codigo]); } catch (e) { r = { s: "erro", e: "perdi a conexão com o JMS" }; }
        if (!r) r = { s: "erro", e: "sem resposta do JMS" };
        if (r.s === "ok") break;
        if (r.e === "perdi a conexão com o JMS" || r.e === "tela de rastreamento não está aberta") break;
        if (r.e === "pedido não encontrado" && t >= 1) break;
        await sleep(1200 * (t + 1));
      }
      if (r.e === "perdi a conexão com o JMS" || r.e === "tela de rastreamento não está aberta") {
        busca.erro = "A busca parou: a aba do JMS mudou de tela ou foi fechada. Volte ao JMS e clique em Buscar de novo — o que já foi buscado está salvo.";
        break;
      }
      estado.res[codigo] = r; busca.feitos++; salvar(); render();
    }
  } catch (e) {
    busca.erro = "Algo deu errado ao falar com o JMS: " + (e && e.message ? e.message : e);
  } finally {
    busca.rodando = false; salvar(); render();
  }
}

/* ---------- Excel ---------- */
function exportarExcel() {
  const c = contexto();
  const cab = ["Motorista", "Código", "Cliente", "Telefone", "Endereço", "Bairro", "Cidade", "CEP", "Status", "Aviso", "Enviado", "Mensagem", "Link WhatsApp"];
  const linha = (l) => [l.motorista, l.codigo, l.cliente, l.telCru, l.endereco, l.bairro, l.cidade, l.cep,
    l.status === "ok" ? "encontrado" : l.status === "pend" ? "não buscado" : "erro: " + l.motivoErro, l.aviso, estado.feitos[l.codigo] ? "sim" : "", l.mensagem, l.whats];
  const larg = (ws, w) => { ws["!cols"] = w.map((x) => ({ wch: x })); return ws; };
  const W = [30, 18, 26, 14, 40, 22, 16, 10, 28, 30, 9, 60, 40];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([["#", "Motorista", "Endereço incorreto", "Pacotes do motorista", "% dos pacotes"],
    ...c.rank.map((r) => [r.pos, r.motorista, r.qtd, r.total, r.pct])]), [5, 36, 18, 20, 14]), "Ranking");
  XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([cab, ...c.linhas.map(linha)]), W), "Todos");
  const usados = new Set(["Ranking", "Todos"]);
  for (const r of c.rank) {
    let nome = `${r.pos}-${r.motorista}`.replace(/[\[\]:*?/\\]/g, "").slice(0, 31);
    while (usados.has(nome)) nome = nome.slice(0, 29) + "_";
    usados.add(nome);
    XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([cab, ...c.linhas.filter((l) => l.motorista === r.motorista).map(linha)]), W), nome);
  }
  XLSX.writeFile(wb, `endereco_incorreto_${new Date().toISOString().slice(0, 10).replaceAll("-", "")}.xlsx`);
  toast("Excel baixado");
}

/* ---------- eventos ---------- */
function mover(fila, i, d) { return fila[(i + d + fila.length) % fila.length].codigo; }

$("abas").onclick = (e) => {
  const b = e.target.closest("button[data-aba]");
  if (b) { aba = b.dataset.aba; abriuWhats = false; render(); }
};

$("tela").onclick = async (e) => {
  const b = e.target.closest("[data-a]");
  if (!b || b.disabled) return;
  const a = b.dataset.a, c = contexto();
  const atual = c.fila.find((l) => l.codigo === estado.atual) || c.fila[0];
  const i = atual ? c.fila.indexOf(atual) : 0;
  switch (a) {
    case "buscar": buscarTelefones(); return;
    case "parar": busca.parar = true; toast("Parando…"); return;
    case "ir-atender": aba = "atender"; abriuWhats = false; break;
    case "ir-resumo": aba = "resumo"; break;
    case "whats": if (atual && atual.whats) { abriuWhats = true; abrirWhats(atual.whats); } break;
    case "cp-msg": if (atual) copiar(atual.mensagem, "Mensagem copiada"); return;
    case "cp-tel": if (atual) copiar(atual.telCru, "Telefone copiado"); return;
    case "cp-cod": copiar(b.dataset.c, "Código copiado"); return;
    case "ant": if (c.fila.length > 1) { estado.atual = mover(c.fila, i, -1); abriuWhats = false; } break;
    case "pular": if (c.fila.length > 1) { estado.atual = mover(c.fila, i, 1); abriuWhats = false; } break;
    case "feito":
      if (!atual) break;
      estado.feitos[atual.codigo] = 1;
      estado.atual = c.fila.length > 1 ? c.fila[(i + 1) % c.fila.length].codigo : null;
      abriuWhats = false;
      break;
    case "excel": exportarExcel(); return;
    case "modelo-padrao": estado.modelo = MODELO_PADRAO; break;
    case "limpar":
      if (!confirm("Apagar a planilha, os telefones e as marcações de enviado?")) return;
      estado = { ...novoEstado(), modelo: estado.modelo }; aba = "resumo"; busca.erro = ""; break;
  }
  salvar(); render();
};

$("tela").onchange = async (e) => {
  if (e.target.id !== "arquivo") return;
  const f = e.target.files[0];
  if (!f) return;
  try { busca.erro = ""; await lerArquivo(f); aba = "resumo"; render(); toast("Planilha carregada"); }
  catch (err) { const el = $("erroArquivo"); if (el) el.innerHTML = `<div class="erro">${esc(err.message)}</div>`; }
};

$("tela").oninput = (e) => {
  if (e.target.id === "modelo") { estado.modelo = e.target.value || MODELO_PADRAO; salvar(); }
};

(async () => {
  const s = await store.ler();
  if (s && Array.isArray(s.pacotes)) estado = { ...novoEstado(), ...s };
  render();
})();
