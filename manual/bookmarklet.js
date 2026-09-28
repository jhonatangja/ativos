/* Roda DENTRO da tela "Rastreamento do pacote" do JMS.
   Este arquivo é embutido no index.html pelo app.js (função ativoJms) e transformado em favorito. */
function ativoJms() {
  if (document.getElementById("__ativo")) { document.getElementById("__ativo").remove(); }
  var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var ate = async function (fn, ms) {
    var t = Date.now();
    while (Date.now() - t < ms) { if (fn()) return true; await sleep(150); }
    return false;
  };
  var btn = function (t) {
    return Array.from(document.querySelectorAll("button")).find(function (b) { return b.textContent.trim() === t; });
  };
  var blk = function () {
    return Array.from(document.querySelectorAll(".waybill-msg-item")).find(function (b) {
      var t = b.querySelector(".item-title");
      return t && t.textContent.trim() === "Destinatário";
    });
  };
  var campo = function (b, l) {
    return Array.from(b.querySelectorAll(".item-msg")).find(function (i) {
      var t = i.querySelector(".tt");
      return t && t.textContent.trim() === l;
    });
  };
  var valor = function (b, l) {
    var c = campo(b, l);
    return c ? c.querySelector(".des").textContent.trim() : "";
  };

  async function consultarUm(code) {
    var limpar = btn("Limpar");
    if (!limpar || !document.querySelector("input.new-tag")) {
      return { s: "erro", e: "abra a tela Rastreamento do pacote" };
    }
    limpar.click();
    await ate(function () { return !blk(); }, 5000);
    var inp = document.querySelector("input.new-tag");
    inp.focus();
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value").set.call(inp, code);
    inp.dispatchEvent(new Event("input", { bubbles: true }));
    ["keydown", "keyup"].forEach(function (t) {
      inp.dispatchEvent(new KeyboardEvent(t, { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true }));
    });
    var lado = document.querySelector(".sidebar-input-tag");
    await ate(function () { return lado && lado.innerText.indexOf(code) >= 0; }, 3000);
    await sleep(200);
    btn("Consulta").click();
    var cab = new RegExp("Número do pedido JMS\\s*[:：]\\s*" + code);
    var achou = await ate(function () { return blk() || cab.test(document.body.innerText); }, 15000);
    if (!achou) return { s: "erro", e: "pedido não encontrado" };
    if (!blk()) {
      var h = btn("Informação básica");
      if (h) h.click();
      await ate(blk, 4000);
    }
    var b = blk();
    if (!b) return { s: "erro", e: "bloco do destinatário não apareceu" };
    var tel = campo(b, "Telefone");
    if (!tel) return { s: "erro", e: "sem campo telefone" };
    var des = tel.querySelector(".des");
    if (des.textContent.indexOf("*") >= 0) {
      var ic = tel.querySelector("i.iconfont");
      if (ic) ic.click();
      await ate(function () { return des.textContent.indexOf("*") < 0; }, 4000);
    }
    var n = des.textContent.trim();
    if (!n || n.indexOf("*") >= 0) return { s: "erro", e: "telefone não expandiu" };
    return { s: "ok", t: n, n: valor(b, "Nome"), a: valor(b, "Endereço") };
  }

  var css = "#__ativo{position:fixed;top:12px;right:12px;z-index:2147483647;width:360px;background:#fff;color:#1c2430;" +
    "border:2px solid #2563eb;border-radius:10px;padding:12px;font:13px system-ui,sans-serif;box-shadow:0 8px 30px rgba(0,0,0,.3)}" +
    "#__ativo textarea{width:100%;height:90px;box-sizing:border-box;font:12px monospace;margin:6px 0}" +
    "#__ativo button{padding:6px 12px;margin-right:6px;border:1px solid #2563eb;border-radius:6px;background:#2563eb;color:#fff;cursor:pointer}" +
    "#__ativo button.sec{background:#fff;color:#2563eb}#__ativo b{font-size:14px}#__ativo .x{float:right;cursor:pointer}";
  var box = document.createElement("div");
  box.id = "__ativo";
  box.innerHTML = "<style>" + css + "</style><span class='x' id='__ax'>✕</span><b>Ativo · consulta de telefones</b>" +
    "<div>1) Cole os códigos copiados da página Ativo:</div><textarea id='__ain'></textarea>" +
    "<button id='__ago'>Iniciar</button><button class='sec' id='__astop'>Parar</button>" +
    "<div id='__ast' style='margin:8px 0'></div>" +
    "<div>2) Ao terminar, copie o resultado e cole na página Ativo:</div><textarea id='__aout' readonly></textarea>" +
    "<button id='__acp'>Copiar resultado</button>";
  document.body.appendChild(box);
  var $ = function (id) { return document.getElementById(id); };
  var res = {}, parar = false, rodando = false;
  $("__ax").onclick = function () { parar = true; box.remove(); };
  $("__astop").onclick = function () { parar = true; };
  $("__acp").onclick = function () {
    var o = $("__aout"); o.select();
    try { navigator.clipboard.writeText(o.value); } catch (e) { document.execCommand("copy"); }
    $("__ast").textContent = "Resultado copiado. Volte à página Ativo e cole.";
  };
  $("__ago").onclick = async function () {
    if (rodando) return;
    var codes = Array.from(new Set(($("__ain").value.match(/\d{10,}/g) || [])));
    if (!codes.length) { $("__ast").textContent = "Nenhum código encontrado."; return; }
    rodando = true; parar = false;
    for (var i = 0; i < codes.length && !parar; i++) {
      var c = codes[i], r;
      for (var t = 0; t < 3; t++) {
        try { r = await consultarUm(c); } catch (e) { r = { s: "erro", e: String(e).slice(0, 80) }; }
        if (r.s === "ok" || r.e === "abra a tela Rastreamento do pacote") break;
        if (r.e === "pedido não encontrado" && t >= 1) break;
        await sleep(1200 * (t + 1));
      }
      res[c] = r;
      var ok = Object.keys(res).filter(function (k) { return res[k].s === "ok"; }).length;
      $("__ast").textContent = "Consultados " + (i + 1) + "/" + codes.length + " · ok: " + ok + " · erro: " + (i + 1 - ok);
      $("__aout").value = JSON.stringify(res);
      if (r.e === "abra a tela Rastreamento do pacote") break;
      await sleep(250);
    }
    rodando = false;
    $("__ast").textContent += parar ? " · parado" : " · CONCLUÍDO";
  };
}
