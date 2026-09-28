/* Funções executadas DENTRO da página do JMS (Rastreamento do pacote), no contexto da própria página.
   O chrome.scripting.executeScript serializa cada função: ela precisa ser 100% autossuficiente
   (nada de variáveis ou funções de fora). Se o JMS mudar de layout, ajuste os seletores aqui. */

var JMS_URL_RASTREIO = "https://jmsbr.jtjms-br.com/app/operatingPlatformIndex/trackingExpress?title=";

// Diz se a tela de rastreamento já está pronta para receber um código.
function jmsPaginaPronta() {
  var limpar = Array.from(document.querySelectorAll("button")).find(function (b) { return b.textContent.trim() === "Limpar"; });
  var consulta = Array.from(document.querySelectorAll("button")).find(function (b) { return b.textContent.trim() === "Consulta"; });
  return !!(limpar && consulta && document.querySelector("input.new-tag"));
}

// Consulta um código e devolve { s: "ok", t: telefone, n: nome, a: endereço } ou { s: "erro", e: motivo }.
async function jmsConsultarCodigo(code) {
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
    var d = c && c.querySelector(".des");
    return d ? d.textContent.trim() : "";
  };

  var limpar = btn("Limpar");
  var inp = document.querySelector("input.new-tag");
  if (!limpar || !inp || !btn("Consulta")) return { s: "erro", e: "tela de rastreamento não está aberta" };

  limpar.click();
  await ate(function () { return !blk(); }, 5000);
  inp = document.querySelector("input.new-tag");
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
  if (!blk()) { // seção "Informação básica" recolhida
    var h = btn("Informação básica");
    if (h) h.click();
    await ate(blk, 4000);
  }
  var b = blk();
  if (!b) return { s: "erro", e: "pedido não encontrado" };
  var tel = campo(b, "Telefone");
  if (!tel) return { s: "erro", e: "campo telefone não encontrado" };
  var des = tel.querySelector(".des");
  if (des.textContent.indexOf("*") >= 0) { // telefone mascarado: clica no olho para revelar
    var ic = tel.querySelector("i.iconfont");
    if (ic) ic.click();
    await ate(function () { return des.textContent.indexOf("*") < 0; }, 4000);
  }
  var n = des.textContent.trim();
  if (!n || n.indexOf("*") >= 0) return { s: "erro", e: "telefone não abriu" };
  return { s: "ok", t: n, n: valor(b, "Nome"), a: valor(b, "Endereço") };
}
