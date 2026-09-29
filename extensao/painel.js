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
const COLUNA_ORIGEM = "Origem do Pedido"; // opcional: se faltar na planilha, a origem só não aparece

// Plataformas/lojas reconhecidas pelo texto da coluna "Origem do Pedido" (ex.: "MERCADO CBT", "sheinDIR", "TEMU D2D").
// [palavra-chave sem acento em minúsculas, nome, "de/da/do + nome"]. Origens que não estão aqui
// (Melhor Envio, intelipost, APIJMS…) são integradores, não a loja: aparecem só no cartão, não na mensagem.
const ORIGENS = [
  ["mercado", "Mercado Livre", "do Mercado Livre"],
  ["shein", "Shein", "da Shein"],
  ["temu", "Temu", "da Temu"],
  ["tiktok", "TikTok Shop", "do TikTok Shop"],
  ["kwai", "Kwai", "do Kwai"],
  ["shopee", "Shopee", "da Shopee"],
  ["jequiti", "Jequiti", "da Jequiti"],
  ["dafiti", "Dafiti", "da Dafiti"],
  ["wepink", "WePink", "da WePink"],
  ["leiturinha", "Leiturinha", "da Leiturinha"],
];
function origemInfo(raw) {
  const n = norm(raw);
  const o = n && ORIGENS.find((x) => n.includes(x[0]));
  return o ? { nome: o[1], de: o[2] } : null;
}

const RETORNOS = {
  enviado: "Enviado",
  respondeu: "Respondeu",
  "sem-resposta": "Não respondeu",
  "nao-ligou": "Diz que o motorista não ligou",
};
const retornoDe = (codigo) => { const v = estado.feitos[codigo]; return v ? (RETORNOS[v] ? v : "enviado") : ""; };

const MODELO_ANTIGO =
  "Olá {primeiro_nome}, tudo bem? Aqui é da entrega do seu pedido (código {codigo}). " +
  "Não conseguimos localizar o endereço informado ({endereco}, {bairro}, {cidade}). " +
  "Pode nos enviar o endereço correto, com rua, número e ponto de referência? " +
  "Assim reagendamos a entrega. Obrigado!";
const MODELO_PADRAO =
  "Olá {primeiro_nome}, tudo bem? Aqui é da entrega do seu pedido {de_origem} (código {codigo}). " +
  "Não conseguimos localizar o endereço informado ({endereco}, {bairro}, {cidade}). " +
  "Pode nos enviar o endereço correto, com rua, número e ponto de referência? " +
  "Assim reagendamos a entrega. Obrigado!";

const novoEstado = () => ({ pacotes: [], totais: {}, res: {}, feitos: {}, desmarcados: {}, filtroMot: "", modelo: MODELO_PADRAO, arquivo: "", atual: null });
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
      origem: String(l[COLUNA_ORIGEM] ?? "").trim(),
    });
  }
  const codigos = new Set(pacotes.map((p) => p.codigo));
  const manter = (o) => Object.fromEntries(Object.entries(o).filter(([k]) => codigos.has(k)));
  estado = { ...estado, pacotes, totais, res: manter(estado.res), feitos: manter(estado.feitos), arquivo: file.name, atual: null, desmarcados: {}, filtroMot: "" };
  salvar();
}

/* ---------- dados derivados ---------- */
function mensagem(p) {
  const pn = p.cliente.split(/\s+/)[0] || "";
  const o = origemInfo(p.origem);
  const v = { primeiro_nome: pn ? pn[0].toUpperCase() + pn.slice(1).toLowerCase() : "", codigo: p.codigo, endereco: p.endereco, bairro: p.bairro, cidade: p.cidade,
    origem: o ? o.nome : "", de_origem: o ? o.de : "" };
  // se a origem não for uma loja conhecida, some junto com o espaço que vinha antes dela
  return estado.modelo.replace(/(\s?)\{(\w+)\}/g, (_, sp, k) => ((k === "origem" || k === "de_origem") && !v[k] ? "" : sp + (v[k] ?? "")));
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
    const o = origemInfo(p.origem);
    return { ...p, status, aviso, motivoErro, tel: status === "ok" ? telBonito(r.t) : "", telCru: status === "ok" ? telNacional(r.t) : "",
      origemNome: o ? o.nome : "", retorno: retornoDe(p.codigo),
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
  const filaTodos = linhas.filter((l) => l.status === "ok" && !l.retorno).sort((a, b) => pos[a.motorista] - pos[b.motorista]);
  if (estado.filtroMot && !rank.some((r) => r.motorista === estado.filtroMot)) estado.filtroMot = "";
  const fila = estado.filtroMot ? filaTodos.filter((l) => l.motorista === estado.filtroMot) : filaTodos; // fila do atendimento (filtrada)
  const paraBuscar = linhas.filter((l) => l.status === "pend" || l.status === "erro");
  const buscaveis = paraBuscar.filter((l) => !estado.desmarcados[l.motorista]); // só os motoristas marcados
  const marcados = rank.filter((r) => !estado.desmarcados[r.motorista]).length;
  const ok = linhas.filter((l) => l.status === "ok");
  return { linhas, rank, pos, fila, filaTodos, paraBuscar, buscaveis, marcados, ok,
    okFiltro: estado.filtroMot ? ok.filter((l) => l.motorista === estado.filtroMot) : ok,
    sem: linhas.filter((l) => l.status === "sem") };
}

// Resumo de retorno por motorista: quem foi contatado, respondeu, não respondeu ou disse que o motorista não ligou.
function retornoPorMotorista(c) {
  return c.rank.map((r) => {
    const ls = c.linhas.filter((l) => l.motorista === r.motorista);
    const n = (k) => ls.filter((l) => l.retorno === k).length;
    const contatados = ls.filter((l) => l.retorno).length;
    return { motorista: r.motorista, pos: r.pos, total: ls.length, contatados, respondeu: n("respondeu"), semResposta: n("sem-resposta"),
      naoLigou: n("nao-ligou"), soEnviado: n("enviado"), naFila: ls.filter((l) => l.status === "ok" && !l.retorno).length,
      semTel: ls.filter((l) => l.status !== "ok").length, clientes: ls };
  }).sort((a, b) => b.naoLigou - a.naoLigou || b.semResposta - a.semResposta || b.total - a.total);
}

/* ---------- telas ---------- */
function render() {
  const c = contexto();
  document.querySelectorAll("#abas button").forEach((b) => b.classList.toggle("on", b.dataset.aba === aba));
  const bg = $("badgeFila");
  bg.hidden = !c.filaTodos.length; bg.textContent = c.filaTodos.length;
  $("tela").innerHTML = aba === "resumo" ? telaResumo(c) : aba === "atender" ? telaAtender(c) : aba === "retorno" ? telaRetorno(c) : telaMais(c);
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
  const alvo = c.buscaveis.length;
  const min = Math.max(1, Math.round(alvo * 4 / 60));
  const motoristasTxt = `${c.marcados} de ${c.rank.length} motorista${c.rank.length > 1 ? "s" : ""} marcado${c.marcados !== 1 ? "s" : ""}`;
  let busca_ui;
  if (busca.rodando) {
    const pct = busca.total ? Math.round(busca.feitos / busca.total * 100) : 0;
    busca_ui = `<h2>Buscando telefones…</h2>
      <div class="prog"><i style="width:${pct}%"></i></div>
      <p>${busca.feitos} de ${busca.total} <span class="mut">· não troque de aba nem feche o JMS</span></p>
      <button class="btn sec" data-a="parar">Parar</button>`;
  } else if (alvo) {
    busca_ui = `<h2>Buscar telefones no JMS</h2>
      <p class="mut">${motoristasTxt}: ${alvo} cliente${alvo > 1 ? "s" : ""} para buscar (cerca de ${min} min). Uso o JMS que você já tem aberto.</p>
      <button class="btn" data-a="buscar">🔎 Buscar telefones (${alvo})</button>
      ${c.ok.length ? `<div class="gap"></div><button class="btn sec" data-a="ir-atender">Já atender os ${c.fila.length} que tenho</button>` : ""}`;
  } else if (!c.marcados) {
    busca_ui = `<h2>Marque os motoristas</h2>
      <p class="mut">Marque abaixo os motoristas cujos clientes você quer consultar.</p>
      ${c.ok.length ? `<button class="btn sec" data-a="ir-atender">Atender os ${c.fila.length} que tenho</button>` : ""}`;
  } else {
    busca_ui = `<h2>✓ ${falta ? "Motoristas marcados prontos" : "Telefones prontos"}</h2>
      <p class="mut">${c.ok.length} com telefone${c.sem.length ? `, ${c.sem.length} sem número` : ""}.${falta ? ` Ainda faltam ${falta} de outros motoristas: marque-os abaixo para buscar.` : ""}</p>
      <button class="btn wa" data-a="ir-atender">Começar atendimento →</button>`;
  }
  const max = Math.max(1, ...c.rank.map((r) => r.qtd));
  const faltaDe = (m) => c.paraBuscar.filter((l) => l.motorista === m).length;
  const linhaRank = (r) => {
    const f = faltaDe(r.motorista);
    return `<label class="mrow"><input type="checkbox" data-m="${esc(r.motorista)}" ${estado.desmarcados[r.motorista] ? "" : "checked"} ${busca.rodando ? "disabled" : ""} />
      <span class="mnome">${r.pos}. ${esc(r.motorista)}<span class="bar"><i style="width:${r.qtd / max * 100}%"></i></span></span>
      <span class="mnum"><b>${r.qtd}</b> <span class="mut">${r.pct}%</span>${f ? `<br><span class="mut">${f} sem tel.</span>` : `<br><span class="ok">✓ prontos</span>`}</span></label>`;
  };
  return `${busca.erro ? `<div class="erro">${esc(busca.erro)}</div>` : ""}
    <div class="grid">
      <div class="stat"><b>${c.linhas.length}</b><span>endereço incorreto</span></div>
      <div class="stat"><b class="ok">${c.ok.length}</b><span>com telefone</span></div>
      <div class="stat"><b>${falta}</b><span>falta buscar</span></div>
      <div class="stat"><b class="${c.sem.length ? "er" : ""}">${c.sem.length}</b><span>sem número</span></div>
    </div>
    <div class="card busca">${busca_ui}</div>
    <div class="card"><div class="cab"><h3>Motoristas</h3>
        <span><button class="link" data-a="marcar-todos" ${busca.rodando ? "disabled" : ""}>Todos</button> · <button class="link" data-a="marcar-nenhum" ${busca.rodando ? "disabled" : ""}>Nenhum</button></span></div>
      ${c.rank.map(linhaRank).join("")}
      <p class="mut" style="margin-top:8px;font-size:12px">% = parte dos pacotes do motorista que teve endereço incorreto. Só os motoristas marcados são consultados.</p>
    </div>
    <div class="mut" style="font-size:12px;text-align:center">${esc(estado.arquivo)}</div>`;
}

function telaAtender(c) {
  if (!estado.pacotes.length) return `<div class="vazio"><div class="g">📄</div><p>Escolha a planilha primeiro.</p><button class="btn" data-a="ir-resumo">Ir para o Resumo</button></div>`;
  if (!c.ok.length) return `<div class="vazio"><div class="g">🔎</div><p><b>Ainda não tem telefone para atender.</b></p><p class="mut">Busque os telefones no Resumo.</p><button class="btn" data-a="ir-resumo">Ir para o Resumo</button></div>`;
  const feitos = c.okFiltro.length - c.fila.length;
  const seletor = seletorMotorista(c);
  if (!c.fila.length) {
    if (estado.filtroMot) {
      return `${seletor}<div class="vazio"><div class="g">✅</div><p><b>Terminei este motorista.</b></p>
        <p class="mut">${feitos} cliente${feitos !== 1 ? "s" : ""} contatado${feitos !== 1 ? "s" : ""}. Escolha outro motorista acima${c.filaTodos.length ? "" : " ou veja o retorno"}.</p>
        <button class="btn sec" data-a="ir-retorno">Ver retorno dos clientes</button></div>`;
    }
    const resto = [];
    if (c.paraBuscar.length) resto.push(`${c.paraBuscar.length} ainda sem telefone buscado ou com erro`);
    if (c.sem.length) resto.push(`${c.sem.length} sem número no JMS`);
    return `<div class="vazio"><div class="g">🎉</div><p><b>Tudo atendido!</b></p>
      <p class="mut">${feitos} cliente${feitos !== 1 ? "s" : ""} contatado${feitos !== 1 ? "s" : ""}.</p>
      ${resto.length ? `<p class="mut">Sobraram: ${resto.join(" · ")}. Veja em "Mais".</p>` : ""}
      <button class="btn" data-a="ir-retorno">Acompanhar respostas →</button><div class="gap"></div>
      ${c.paraBuscar.length ? `<button class="btn sec" data-a="ir-resumo">Buscar de novo no Resumo</button>` : ""}</div>`;
  }
  let i = c.fila.findIndex((l) => l.codigo === estado.atual);
  if (i < 0) i = 0;
  const p = c.fila[i];
  estado.atual = p.codigo;
  const pct = c.okFiltro.length ? Math.round(feitos / c.okFiltro.length * 100) : 0;
  const r = c.rank.find((x) => x.motorista === p.motorista);
  return `${seletor}<div class="prog"><i style="width:${pct}%"></i></div>
    <div class="nav"><button data-a="ant" ${c.fila.length < 2 ? "disabled" : ""}>← Anterior</button>
      <span>Cliente ${i + 1} de ${c.fila.length} · ${feitos} enviados</span>
      <button data-a="pular" ${c.fila.length < 2 ? "disabled" : ""}>Pular →</button></div>
    <div class="card">
      <span class="pill">Motorista #${r.pos}: ${esc(p.motorista)}</span>
      ${p.origem ? `<span class="pill ${p.origemNome ? "loja" : ""}">🛒 ${esc(p.origemNome || p.origem)}</span>` : ""}
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

function seletorMotorista(c) {
  const restam = {};
  for (const l of c.filaTodos) restam[l.motorista] = (restam[l.motorista] || 0) + 1;
  const comTel = new Set(c.ok.map((l) => l.motorista));
  const ops = c.rank.filter((r) => comTel.has(r.motorista));
  return `<select id="fmot" class="selmot"><option value="">Todos os motoristas (${c.filaTodos.length} na fila)</option>
    ${ops.map((r) => `<option value="${esc(r.motorista)}" ${estado.filtroMot === r.motorista ? "selected" : ""}>${r.pos}. ${esc(r.motorista)} (${restam[r.motorista] || 0} na fila)</option>`).join("")}</select>`;
}

const abertos = {}; // motoristas expandidos na aba Retorno
function telaRetorno(c) {
  if (!estado.pacotes.length) return `<div class="vazio"><div class="g">📄</div><p>Escolha a planilha primeiro.</p><button class="btn" data-a="ir-resumo">Ir para o Resumo</button></div>`;
  const rs = retornoPorMotorista(c);
  const cont = (k) => c.linhas.filter((l) => l.retorno === k).length;
  const contatados = c.linhas.filter((l) => l.retorno).length;
  if (!contatados) {
    return `<div class="card"><h2>Retorno dos clientes</h2>
      <p class="mut">Depois de enviar as mensagens, volte aqui e marque como cada cliente respondeu. Assim você enxerga quais motoristas têm clientes que não respondem ou que dizem que o motorista nunca ligou.</p>
      <button class="btn" data-a="ir-atender">Ir para o atendimento</button></div>`;
  }
  const chips = (r) => [
    r.respondeu ? `<span class="ok">✅ ${r.respondeu} respondeu</span>` : "",
    r.semResposta ? `<span class="warn">⏳ ${r.semResposta} sem resposta</span>` : "",
    r.naoLigou ? `<span class="er">🚫 ${r.naoLigou} diz que não ligou</span>` : "",
    r.soEnviado ? `<span class="mut">📤 ${r.soEnviado} aguardando</span>` : "",
    r.naFila ? `<span class="mut">🕐 ${r.naFila} na fila</span>` : "",
  ].filter(Boolean).join(" · ");
  const acao = (l, v, rot) => `<button class="mini ${l.retorno === v ? "sel" : ""}" data-a="ret" data-c="${esc(l.codigo)}" data-v="${v}">${rot}</button>`;
  return `<div class="grid">
      <div class="stat"><b>${contatados}</b><span>contatados</span></div>
      <div class="stat"><b class="ok">${cont("respondeu")}</b><span>responderam</span></div>
      <div class="stat"><b class="warn">${cont("sem-resposta")}</b><span>não responderam</span></div>
      <div class="stat"><b class="er">${cont("nao-ligou")}</b><span>dizem que o motorista não ligou</span></div>
    </div>
    <p class="mut" style="font-size:12px">Motoristas com mais clientes que "não ligou" e "sem resposta" aparecem primeiro: são os que merecem cobrança sobre o ativo.</p>
    ${rs.map((r) => {
      const feitos = r.clientes.filter((l) => l.retorno);
      return `<details class="drv" data-d="${esc(r.motorista)}" ${abertos[r.motorista] ? "open" : ""}>
        <summary><span>${r.naoLigou ? "🚩 " : ""}${esc(r.motorista)}</span><span class="mut">${r.contatados}/${r.total}</span></summary>
        <div class="drvbody"><div class="chips">${chips(r) || '<span class="mut">Ninguém contatado ainda</span>'}</div>
          ${r.naFila ? `<button class="btn sec peq" data-a="atender-mot" data-m="${esc(r.motorista)}">Atender os ${r.naFila} da fila deste motorista</button>` : ""}
          ${feitos.map((l) => `<div class="item"><b>${esc(l.cliente)}</b> <span class="mut">${esc(l.tel)}${l.origemNome ? " · " + esc(l.origemNome) : ""}</span><br>
            <span class="tag">${esc(RETORNOS[l.retorno])}</span>
            <div class="acoes">${acao(l, "respondeu", "✅ Respondeu")}${acao(l, "sem-resposta", "⏳ Sem resposta")}${acao(l, "nao-ligou", "🚫 Não ligou")}
              <button class="mini" data-a="rwhats" data-c="${esc(l.codigo)}">💬</button><button class="mini" data-a="ret" data-c="${esc(l.codigo)}" data-v="">↩ Fila</button></div></div>`).join("")}
        </div></details>`;
    }).join("")}`;
}

function telaMais(c) {
  const lista = [...c.sem, ...c.linhas.filter((l) => l.status === "erro")];
  return `<div class="card"><h3>Exportar</h3>
      <button class="btn" data-a="excel" ${estado.pacotes.length ? "" : "disabled"}>📥 Baixar Excel (ranking + lista por motorista)</button></div>
    <div class="card"><h3>Mensagem para o cliente</h3>
      <p class="mut" style="font-size:12px">Variáveis: <code>{primeiro_nome}</code> <code>{codigo}</code> <code>{endereco}</code> <code>{bairro}</code> <code>{cidade}</code>
        <code>{de_origem}</code> (ex.: "do Mercado Livre", "da Shein") <code>{origem}</code> (ex.: "Mercado Livre"). Se a origem não for uma loja conhecida, some sozinha.</p>
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
  const lista = c.buscaveis.map((l) => l.codigo);
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
  const cab = ["Motorista", "Código", "Cliente", "Origem", "Telefone", "Endereço", "Bairro", "Cidade", "CEP", "Status da busca", "Aviso", "Retorno", "Mensagem", "Link WhatsApp"];
  const linha = (l) => [l.motorista, l.codigo, l.cliente, l.origemNome || l.origem, l.telCru, l.endereco, l.bairro, l.cidade, l.cep,
    l.status === "ok" ? "encontrado" : l.status === "pend" ? "não buscado" : "erro: " + l.motivoErro, l.aviso, l.retorno ? RETORNOS[l.retorno] : "", l.mensagem, l.whats];
  const larg = (ws, w) => { ws["!cols"] = w.map((x) => ({ wch: x })); return ws; };
  const W = [30, 18, 26, 16, 14, 40, 22, 16, 10, 28, 30, 26, 60, 40];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([["#", "Motorista", "Endereço incorreto", "Pacotes do motorista", "% dos pacotes"],
    ...c.rank.map((r) => [r.pos, r.motorista, r.qtd, r.total, r.pct])]), [5, 36, 18, 20, 14]), "Ranking");
  XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([["Motorista", "Endereço incorreto", "Contatados", "Responderam", "Sem resposta", "Disseram que não ligou", "Aguardando", "Na fila", "Sem telefone"],
    ...retornoPorMotorista(c).map((r) => [r.motorista, r.total, r.contatados, r.respondeu, r.semResposta, r.naoLigou, r.soEnviado, r.naFila, r.semTel])]),
    [36, 18, 12, 13, 13, 22, 12, 9, 13]), "Retorno por motorista");
  XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([cab, ...c.linhas.map(linha)]), W), "Todos");
  const usados = new Set(["Ranking", "Retorno por motorista", "Todos"]);
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
    case "ir-retorno": aba = "retorno"; break;
    case "atender-mot": estado.filtroMot = b.dataset.m; estado.atual = null; aba = "atender"; abriuWhats = false; break;
    case "ret": { // marca/atualiza o retorno de um cliente ("" = volta para a fila)
      if (b.dataset.v) estado.feitos[b.dataset.c] = b.dataset.v; else delete estado.feitos[b.dataset.c];
      break;
    }
    case "rwhats": { const l = c.linhas.find((x) => x.codigo === b.dataset.c); if (l && l.whats) abrirWhats(l.whats); return; }
    case "whats": if (atual && atual.whats) { abriuWhats = true; abrirWhats(atual.whats); } break;
    case "cp-msg": if (atual) copiar(atual.mensagem, "Mensagem copiada"); return;
    case "cp-tel": if (atual) copiar(atual.telCru, "Telefone copiado"); return;
    case "cp-cod": copiar(b.dataset.c, "Código copiado"); return;
    case "ant": if (c.fila.length > 1) { estado.atual = mover(c.fila, i, -1); abriuWhats = false; } break;
    case "pular": if (c.fila.length > 1) { estado.atual = mover(c.fila, i, 1); abriuWhats = false; } break;
    case "feito":
      if (!atual) break;
      estado.feitos[atual.codigo] = "enviado";
      estado.atual = c.fila.length > 1 ? c.fila[(i + 1) % c.fila.length].codigo : null;
      abriuWhats = false;
      break;
    case "marcar-todos": estado.desmarcados = {}; break;
    case "marcar-nenhum": estado.desmarcados = Object.fromEntries(c.rank.map((r) => [r.motorista, true])); break;
    case "excel": exportarExcel(); return;
    case "modelo-padrao": estado.modelo = MODELO_PADRAO; break;
    case "limpar":
      if (!confirm("Apagar a planilha, os telefones e as marcações de enviado?")) return;
      estado = { ...novoEstado(), modelo: estado.modelo }; aba = "resumo"; busca.erro = ""; break;
  }
  salvar(); render();
};

$("tela").addEventListener("toggle", (e) => { // lembra quais motoristas estão abertos na aba Retorno
  const d = e.target.dataset && e.target.dataset.d;
  if (d !== undefined) abertos[d] = e.target.open;
}, true);

$("tela").onchange = async (e) => {
  if (e.target.id === "fmot") { estado.filtroMot = e.target.value; estado.atual = null; abriuWhats = false; salvar(); render(); return; }
  if (e.target.dataset.m !== undefined) {
    const m = e.target.dataset.m;
    if (e.target.checked) delete estado.desmarcados[m]; else estado.desmarcados[m] = true;
    salvar(); render();
    return;
  }
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
  if (estado.modelo === MODELO_ANTIGO) estado.modelo = MODELO_PADRAO; // quem nunca editou a mensagem ganha a versão com a origem
  render();
})();
