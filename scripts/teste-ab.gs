/**
 * Teste A/B de preço da página Símbolos da Santa Igreja (Google Apps Script, ligado a uma planilha).
 *
 * GET  ?acao=escolher        -> sorteia a variante do visitante, conta a visita e devolve {id, preco, checkout}
 * GET  ?acao=escolher&v=b    -> devolve a variante b sem contar visita (para conferir a página)
 * POST ?token=...            -> webhook da Hubla: registra a venda na variante correspondente
 *
 * A escolha usa amostragem de Thompson sobre o faturamento por visita: cada variante recebe tráfego
 * na proporção da chance de ser a que mais fatura. EXPLORACAO garante um mínimo de visitas para todas.
 * Para mudar preços ou links, edite a aba "variantes" da planilha.
 */

const EXPLORACAO = 0.1;
// conversão esperada antes de haver dados (~3%): evita que o preço mais alto leve todo o tráfego no começo
const PRIORI = { vendas: 1, visitas: 30 };
const VARIANTES_INICIAIS = [
  ["a", 10, "https://pay.hub.la/pYU0ErgIZMwYODv0GImh"],
  ["b", 17, "https://pay.hub.la/wSmMtUqAaNRpDM7Ckvey"],
  ["c", 27, "https://pay.hub.la/Qw2ugjX7kSskdSbDiJkr"],
];

/** Rode uma vez pelo editor: cria as abas e o token do webhook. */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let v = ss.getSheetByName("variantes");
  if (!v) {
    v = ss.insertSheet("variantes");
    v.appendRow(["variante", "preco", "checkout", "visitas", "vendas", "conversao", "faturamento por visita"]);
    VARIANTES_INICIAIS.forEach((r, i) => {
      const n = i + 2;
      v.appendRow([...r, 0, 0, `=IFERROR(E${n}/D${n},0)`, `=IFERROR(B${n}*E${n}/D${n},0)`]);
    });
    v.getRange("F2:F").setNumberFormat("0.0%");
    v.getRange("G2:G").setNumberFormat("R$ 0.00");
    v.setFrozenRows(1);
  }
  if (!ss.getSheetByName("vendas")) {
    const s = ss.insertSheet("vendas");
    s.appendRow(["data", "variante", "hash", "payload"]);
    s.setFrozenRows(1);
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("TOKEN")) props.setProperty("TOKEN", Utilities.getUuid().replace(/-/g, ""));
  Logger.log("Token do webhook: " + props.getProperty("TOKEN"));
}

function lerVariantes(sheet) {
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .map((r, i) => ({ linha: i + 2, id: String(r[0]), preco: Number(r[1]), checkout: String(r[2]), visitas: Number(r[3]) || 0, vendas: Number(r[4]) || 0 }))
    .filter(v => v.id && v.checkout);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variantes");
  const forcada = e.parameter.v;
  if (forcada) {
    const v = lerVariantes(sheet).find(x => x.id === forcada);
    return v ? json({ id: v.id, preco: v.preco, checkout: v.checkout }) : json({ erro: "variante desconhecida" });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const vs = lerVariantes(sheet);
    const v = Math.random() < EXPLORACAO ? vs[Math.floor(Math.random() * vs.length)] : thompson(vs);
    sheet.getRange(v.linha, 4).setValue(v.visitas + 1);
    return json({ id: v.id, preco: v.preco, checkout: v.checkout });
  } finally {
    lock.releaseLock();
  }
}

function thompson(vs) {
  let melhor = null, maior = -1;
  vs.forEach(v => {
    const a = PRIORI.vendas + v.vendas;
    const b = PRIORI.visitas - PRIORI.vendas + Math.max(v.visitas - v.vendas, 0);
    const score = beta(a, b) * v.preco;
    if (score > maior) { maior = score; melhor = v; }
  });
  return melhor;
}

function beta(a, b) {
  const x = gamma(a);
  return x / (x + gamma(b));
}

// Marsaglia e Tsang
function gamma(k) {
  if (k < 1) return gamma(k + 1) * Math.pow(Math.random(), 1 / k);
  const d = k - 1 / 3, c = 1 / Math.sqrt(9 * d);
  for (;;) {
    let x, v;
    do { x = normal(); v = 1 + c * x; } while (v <= 0);
    v = v * v * v;
    const u = Math.random();
    if (u < 1 - 0.0331 * x * x * x * x || Math.log(u) < 0.5 * x * x + d * (1 - v + Math.log(v))) return d * v;
  }
}

function normal() {
  return Math.sqrt(-2 * Math.log(1 - Math.random())) * Math.cos(2 * Math.PI * Math.random());
}

function doPost(e) {
  if (e.parameter.token !== PropertiesService.getScriptProperties().getProperty("TOKEN")) return json({ ok: false });
  const corpo = (e.postData && e.postData.contents) || "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("variantes");
  const log = ss.getSheetByName("vendas");
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, corpo));

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // a Hubla pode reenviar o mesmo evento; cada payload conta uma vez só
    if (log.getLastRow() > 1 && log.getRange(2, 3, log.getLastRow() - 1, 1).getValues().some(r => r[0] === hash)) return json({ ok: true, repetido: true });
    const v = identificar(lerVariantes(sheet), corpo);
    if (v) sheet.getRange(v.linha, 5).setValue(v.vendas + 1);
    log.appendRow([new Date(), v ? v.id : "?", hash, corpo.slice(0, 45000)]);
    return json({ ok: true, variante: v ? v.id : null });
  } finally {
    lock.releaseLock();
  }
}

// Acha a variante pelo código da oferta no payload (o final do link do checkout);
// se não houver, pelo valor pago.
function identificar(vs, corpo) {
  const porCodigo = vs.find(v => corpo.includes(v.checkout.split("/").pop()));
  if (porCodigo) return porCodigo;
  let dados;
  try { dados = JSON.parse(corpo); } catch (err) { return null; }
  const valores = [];
  (function andar(o) {
    if (!o || typeof o !== "object") return;
    Object.keys(o).forEach(k => {
      if (/amount|price|total|value|valor|preco/i.test(k) && !isNaN(Number(o[k]))) valores.push(Number(o[k]));
      andar(o[k]);
    });
  })(dados);
  const achadas = vs.filter(v => valores.some(n => n === v.preco || n === v.preco * 100));
  return achadas.length === 1 ? achadas[0] : null;
}
