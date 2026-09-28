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
const CHAVE = "ativos_endereco_v1";

let estado = { pacotes: [], totais: {}, res: {}, feitos: {}, modelo: MODELO_PADRAO, arquivo: "" };

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const norm = (s) => String(s ?? "").normalize("NFKD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

function salvar() {
  try { localStorage.setItem(CHAVE, JSON.stringify(estado)); } catch (e) { /* cota ou modo privado */ }
}
function carregarSalvo() {
  try {
    const s = JSON.parse(localStorage.getItem(CHAVE) || "null");
    if (s && Array.isArray(s.pacotes)) estado = { ...estado, ...s };
  } catch (e) { /* ignora */ }
}

function toast(msg) {
  const t = $("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast.t);
  toast.t = setTimeout(() => t.classList.remove("show"), 1600);
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
function telefoneNacional(bruto) {
  let d = String(bruto ?? "").replace(/\D/g, "");
  if (d.startsWith("55") && d.length >= 12) d = d.slice(2);
  return d;
}
function telefoneBonito(bruto) {
  const d = telefoneNacional(bruto);
  if (d.length === 11) return `(${d.slice(0, 2)}) ${d.slice(2, 7)}-${d.slice(7)}`;
  if (d.length === 10) return `(${d.slice(0, 2)}) ${d.slice(2, 6)}-${d.slice(6)}`;
  return String(bruto ?? "");
}
function linkWhats(bruto, texto) {
  const d = telefoneNacional(bruto);
  return d.length === 10 || d.length === 11 ? `https://wa.me/55${d}?text=${encodeURIComponent(texto)}` : "";
}

/* ---------- leitura da planilha ---------- */
function limpaMotorista(n) {
  return String(n || "").replace(/^\s*F\s+[A-Z]{2,4}\s*-\s*/, "").trim() || "(sem motorista)";
}

async function lerArquivo(file) {
  const wb = /\.(csv|txt)$/i.test(file.name)
    ? XLSX.read(await file.text(), { type: "string" })
    : XLSX.read(await file.arrayBuffer(), { type: "array" });
  const linhas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });
  if (!linhas.length) throw new Error("Planilha vazia.");
  const faltam = Object.values(COLUNAS).filter((c) => !(c in linhas[0]));
  if (faltam.length) throw new Error("Colunas ausentes: " + faltam.join(", "));

  const totais = {};
  const pacotes = [];
  for (const l of linhas) {
    const mot = limpaMotorista(l[COLUNAS.motorista]);
    totais[mot] = (totais[mot] || 0) + 1;
    if (!norm(l[COLUNAS.motivo]).includes("endereco incorreto")) continue;
    const codigo = String(l[COLUNAS.codigo]).replace(/\D/g, "");
    if (!codigo) continue;
    pacotes.push({
      codigo, motorista: mot,
      cliente: String(l[COLUNAS.cliente]).trim(),
      endereco: String(l[COLUNAS.endereco]).trim(),
      bairro: String(l[COLUNAS.bairro]).trim(),
      cidade: String(l[COLUNAS.cidade]).trim(),
      cep: String(l[COLUNAS.cep]).trim(),
    });
  }
  const codigos = new Set(pacotes.map((p) => p.codigo));
  const manter = (o) => Object.fromEntries(Object.entries(o).filter(([k]) => codigos.has(k)));
  estado = { ...estado, pacotes, totais, res: manter(estado.res), feitos: manter(estado.feitos), arquivo: file.name };
  salvar();
}

/* ---------- dados derivados ---------- */
function mensagem(p) {
  const primeiro = (p.cliente.split(/\s+/)[0] || "");
  const vars = {
    primeiro_nome: primeiro ? primeiro[0].toUpperCase() + primeiro.slice(1).toLowerCase() : "",
    codigo: p.codigo, endereco: p.endereco, bairro: p.bairro, cidade: p.cidade,
  };
  return estado.modelo.replace(/\{(\w+)\}/g, (_, k) => vars[k] ?? "");
}

function derivar() {
  const donos = {};
  for (const p of estado.pacotes) {
    const r = estado.res[p.codigo];
    if (r && r.s === "ok") {
      const d = telefoneNacional(r.t);
      (donos[d] = donos[d] || new Set()).add(norm(p.cliente));
    }
  }
  return estado.pacotes.map((p) => {
    const r = estado.res[p.codigo];
    const d = r && r.s === "ok" ? telefoneNacional(r.t) : "";
    let status = "pend", aviso = "", motivoErro = "";
    if (r && r.s === "ok") {
      if (d.length !== 10 && d.length !== 11) { status = "sem"; motivoErro = "sem telefone cadastrado no JMS"; }
      else {
        status = "ok";
        if (d.length === 10) aviso = "10 dígitos (falta o 9?)";
        if (donos[d].size > 1) aviso = (aviso ? aviso + " · " : "") + "mesmo telefone em clientes diferentes";
      }
    } else if (r) { status = "sem"; motivoErro = r.e || "erro"; }
    const msg = mensagem(p);
    return { ...p, status, aviso, motivoErro, telefone: status === "ok" ? telefoneBonito(r.t) : "",
      telefoneCru: status === "ok" ? r.t : "", mensagem: msg, whats: status === "ok" ? linkWhats(r.t, msg) : "" };
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

/* ---------- render ---------- */
function render() {
  const tem = estado.pacotes.length > 0;
  $("painel").hidden = !tem;
  $("btnReset").hidden = !tem;
  if (document.activeElement !== $("modelo")) $("modelo").value = estado.modelo;
  $("fileInfo").textContent = tem ? `${estado.arquivo} · ${estado.pacotes.length} pacotes com endereço incorreto` : "";
  if (!tem) return;

  const linhas = derivar();
  const rank = rankingDe(linhas);
  const cont = (s) => linhas.filter((l) => l.status === s).length;
  const pend = linhas.filter((l) => l.status === "pend" || (estado.res[l.codigo] && estado.res[l.codigo].s === "erro"));
  $("qtdPend").textContent = pend.length;

  $("stats").innerHTML = [
    ["Endereço incorreto", linhas.length], ["Com telefone", cont("ok")], ["Sem telefone / erro", cont("sem")],
    ["Não consultados", cont("pend")], ["Motoristas", rank.length],
  ].map(([l, v]) => `<div class="stat"><b>${v}</b><span>${l}</span></div>`).join("");

  const max = Math.max(1, ...rank.map((r) => r.qtd));
  $("rank").innerHTML = "<tr><th>#</th><th>Motorista</th><th>Qtd</th><th>% dos pacotes dele</th><th style='width:30%'></th></tr>" +
    rank.map((r) => `<tr><td>${r.pos}</td><td>${esc(r.motorista)}</td><td><b>${r.qtd}</b></td><td>${r.pct}% <span class="mut">(${r.qtd}/${r.total})</span></td><td><div class="meter" style="width:${r.qtd / max * 100}%"></div></td></tr>`).join("");

  const q = $("q").value.toLowerCase(), f = $("fStatus").value;
  const passa = (l) => {
    if (q && !JSON.stringify(l).toLowerCase().includes(q)) return false;
    if (f === "aviso") return !!l.aviso;
    if (f === "feito") return !!estado.feitos[l.codigo];
    return !f || l.status === f;
  };
  const lista = $("lista");
  lista.innerHTML = "";
  for (const r of rank) {
    const ps = linhas.filter((l) => l.motorista === r.motorista && passa(l));
    if (!ps.length) continue;
    const det = document.createElement("details");
    det.className = "drv";
    det.open = !!(q || f) || r.pos <= 2;
    det.innerHTML = `<summary><span>${r.pos}. ${esc(r.motorista)}</span><span class="mut">${ps.length} pacote(s)</span></summary>
      <div class="wrap"><table><tr><th>Código</th><th>Cliente</th><th>Telefone</th><th>Endereço</th><th>Status</th><th>Ações</th></tr></table></div>`;
    const tb = det.querySelector("table");
    for (const p of ps) {
      const tr = document.createElement("tr");
      if (estado.feitos[p.codigo]) tr.className = "feito";
      const cls = p.status === "ok" ? "ok" : p.status === "sem" ? "er" : "mut";
      const txt = p.status === "ok" ? "encontrado" : p.status === "sem" ? "erro: " + p.motivoErro : "não consultado";
      tr.innerHTML = `<td>${esc(p.codigo)}</td><td>${esc(p.cliente)}</td>
        <td>${p.telefone ? esc(p.telefone) : '<span class="mut">—</span>'}${p.aviso ? `<br><span class="warn">⚠ ${esc(p.aviso)}</span>` : ""}</td>
        <td>${esc(p.endereco)}<br><span class="mut">${esc(p.bairro)} · ${esc(p.cidade)} · ${esc(p.cep)}</span></td>
        <td class="${cls}">${esc(txt)}</td><td></td>`;
      const ac = tr.lastElementChild;
      const add = (rot, fn) => { const b = document.createElement("button"); b.className = "mini"; b.textContent = rot; b.onclick = () => fn(b); ac.append(b); };
      add("Cód", () => copiar(p.codigo, "Código copiado"));
      if (p.telefoneCru) add("Tel", () => copiar(telefoneNacional(p.telefoneCru), "Telefone copiado"));
      add("End", () => copiar(`${p.endereco}, ${p.bairro}, ${p.cidade}`, "Endereço copiado"));
      add("Msg", () => copiar(p.mensagem, "Mensagem copiada"));
      if (p.whats) add("WhatsApp", () => window.open(p.whats, "_blank", "noopener"));
      add(estado.feitos[p.codigo] ? "Desfazer" : "Enviado", () => {
        if (estado.feitos[p.codigo]) delete estado.feitos[p.codigo]; else estado.feitos[p.codigo] = 1;
        salvar(); render();
      });
      tb.append(tr);
    }
    lista.append(det);
  }
  if (!lista.children.length) lista.innerHTML = '<p class="mut">Nenhum pacote com esse filtro.</p>';
}

/* ---------- ações ---------- */
function importarResultado() {
  let dados;
  try { dados = JSON.parse($("colar").value.trim()); } catch (e) { $("importInfo").textContent = "Texto inválido — cole exatamente o que o Ativo JMS copiou."; return; }
  const codigos = new Set(estado.pacotes.map((p) => p.codigo));
  let n = 0;
  for (const [c, r] of Object.entries(dados)) {
    if (!codigos.has(c) || !r || !r.s) continue;
    const antigo = estado.res[c];
    if (antigo && antigo.s === "ok" && r.s !== "ok") continue;
    estado.res[c] = r; n++;
  }
  salvar();
  $("colar").value = "";
  $("importInfo").textContent = `${n} resultado(s) importado(s).`;
  render();
}

function exportarExcel() {
  const linhas = derivar();
  const rank = rankingDe(linhas);
  const cab = ["Motorista", "Código", "Cliente", "Telefone", "Endereço", "Bairro", "Cidade", "CEP", "Status", "Aviso", "Mensagem", "Link WhatsApp"];
  const linha = (l) => [l.motorista, l.codigo, l.cliente, l.telefoneCru ? telefoneNacional(l.telefoneCru) : "", l.endereco, l.bairro, l.cidade, l.cep,
    l.status === "ok" ? "encontrado" : l.status === "sem" ? "erro: " + l.motivoErro : "não consultado", l.aviso, l.mensagem, l.whats];
  const larg = (ws, cols) => { ws["!cols"] = cols.map((w) => ({ wch: w })); return ws; };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([["#", "Motorista", "Endereço incorreto", "Pacotes do motorista", "% dos pacotes"],
    ...rank.map((r) => [r.pos, r.motorista, r.qtd, r.total, r.pct])]), [5, 36, 18, 20, 14]), "Ranking");
  const larguras = [30, 18, 26, 14, 40, 22, 16, 10, 28, 30, 60, 40];
  XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([cab, ...linhas.map(linha)]), larguras), "Todos");
  const usados = new Set(["Ranking", "Todos"]);
  for (const r of rank) {
    let nome = `${r.pos}-${r.motorista}`.replace(/[\[\]:*?/\\]/g, "").slice(0, 31);
    while (usados.has(nome)) nome = nome.slice(0, 29) + "_";
    usados.add(nome);
    XLSX.utils.book_append_sheet(wb, larg(XLSX.utils.aoa_to_sheet([cab, ...linhas.filter((l) => l.motorista === r.motorista).map(linha)]), larguras), nome);
  }
  const hoje = new Date().toISOString().slice(0, 10).replaceAll("-", "");
  XLSX.writeFile(wb, `endereco_incorreto_${hoje}.xlsx`);
}

/* ---------- init ---------- */
$("bookmarklet").href = "javascript:" + encodeURIComponent("(" + ativoJms.toString() + ")()");
$("bookmarklet").onclick = (e) => { e.preventDefault(); toast("Arraste este botão para a barra de favoritos"); };

$("file").onchange = async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  try { await lerArquivo(f); render(); toast("Planilha importada"); }
  catch (err) { $("fileInfo").innerHTML = `<span class="er">${esc(err.message)}</span>`; }
  e.target.value = "";
};
$("btnCopiar").onclick = () => {
  const cods = derivar().filter((l) => l.status === "pend" || (estado.res[l.codigo] && estado.res[l.codigo].s === "erro")).map((l) => l.codigo);
  if (!cods.length) return toast("Nada pendente");
  copiar(cods.join("\n"), `${cods.length} códigos copiados`);
};
$("btnImportar").onclick = importarResultado;
$("btnExcel").onclick = exportarExcel;
$("q").oninput = render;
$("fStatus").onchange = render;
$("modelo").oninput = (e) => { estado.modelo = e.target.value || MODELO_PADRAO; salvar(); clearTimeout(render.t); render.t = setTimeout(render, 300); };
$("btnReset").onclick = () => {
  if (!confirm("Apagar a planilha, os telefones e as marcações deste computador?")) return;
  estado = { pacotes: [], totais: {}, res: {}, feitos: {}, modelo: estado.modelo, arquivo: "" };
  salvar(); render();
};

carregarSalvo();
render();
