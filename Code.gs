/**
 * PERFORMANCE FINANCIAL — BACKEND v2
 *
 * Principais diferenças face à v1:
 *  - Todas as funções públicas exigem um token de sessão (login deixa de ser só visual).
 *  - Passwords guardadas com hash (as antigas em texto simples são convertidas no 1.º login).
 *  - Cada folha é lida UMA vez por pedido (camada de dados em memória).
 *  - Faturas de cartão passam a ter ano (sem misturar out/2026 com out/2027).
 *  - Dívidas: editar já não perde prestações pagas; datas 29-31 corrigidas; arredondamentos corrigidos.
 *  - Gravações protegidas por LockService.
 *  - Alertas de metas fora do carregamento do dashboard (trigger diário + após novas despesas).
 *
 * A estrutura das folhas NÃO foi alterada.
 */

var SS_ID_PADRAO = "1r4c2q9vtOAhELtxADxmuh4Ytso3KpBk8uxPy-n23uCM";
var SESSAO_SEGUNDOS_ = 6 * 3600;   // máximo permitido pelo CacheService
var MAX_TENTATIVAS_ = 5;

var ABA_ = {
  DESP: "DESPESAS", REC: "RECEITA", CART_REG: "REGISTOS CARTÃO DE CRÉDITO",
  CART: "CARTÃO DE CRÉDITO", FAT: "FATURA CARTÃO DE CRÉDITO", CONTAS: "CONTAS",
  CATS: "CATEGORIAS", METAS: "METAS", USR: "USUÁRIO", DIV: "DÍVIDA", DET: "DÍVIDA - DETALHES",
  TRF: "TRANSFERÊNCIAS", AJU: "AJUSTES", RFIX: "RECORRENTES", LOG: "LOG", ORC: "ORÇAMENTOS"
};
var CAB_TRF_ = ["DATA", "ORIGEM", "DESTINO", "VALOR", "DESCRIÇÃO", "ID"];
var CAB_AJU_ = ["DATA", "CONTA", "VALOR", "MOTIVO"];
var CAB_RFIX_ = ["ID", "TIPO", "DESCRIÇÃO", "CATEGORIA", "CONTA/CARTÃO", "VALOR", "DIA", "INÍCIO", "FIM", "ÚLTIMO GERADO", "ATIVO"];
var TOTAL_ = "*TOTAL*";                        // categoria especial que guarda o orçamento total do mês
var MODALIDADES_ = ["Conta corrente", "Poupança", "Cartão refeição", "Dinheiro", "Investimento"];
var PASTA_BACKUP_ = "Performance Financial - Cópias de segurança";
var BACKUPS_MANTER_ = 8;
var USER_ATUAL_ = "";
var MESES_ = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
var ABREV_ = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];

/* ========================================================================
 * ENTRADA / MENU
 * ====================================================================== */

function menuAbrirAcessos() {
  SpreadsheetApp.getUi();                       // só funciona dentro da planilha (nunca a partir da web app)
  var ss = getSS_(), s = garantirAcessos_(ss);
  ss.setActiveSheet(s);
}

function menuNovoUtilizador() {
  var ui = SpreadsheetApp.getUi();
  var n = ui.prompt("Novo utilizador", "Nome de utilizador (ex.: Kenneth):", ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  var p = ui.prompt("Senha provisória", "Pelo menos 8 caracteres. A pessoa pode mudá-la depois na app:", ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  var q = ui.prompt("Pessoa nas despesas partilhadas", "Nome exatamente como na folha PESSOAS (ex.: Kenneth):", ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;
  var nome = n.getResponseText().trim(), senha = p.getResponseText().trim();
  if (nome.length < 2 || senha.length < 8) { ui.alert("Nome ou senha inválidos (a senha precisa de 8 caracteres)."); return; }
  var ss = getSS_(), usr = sheetObrig_(ss, ABA_.USR), s = garantirAcessos_(ss);
  var existe = usr.getLastRow() >= 2 && usr.getRange(2, 1, usr.getLastRow() - 1, 1).getValues().some(function (r) { return norm_(r[0]) === norm_(nome); });
  if (existe) { ui.alert("Esse utilizador já existe. Use a folha ACESSOS para mudar permissões."); return; }
  usr.getRange(usr.getLastRow() + 1, 1, 1, 3).setValues([[nome, hashSenha_(senha), ""]]);
  var pp = {}; ["PARTILHADAS_VER", "PARTILHADAS_LANCAR", "ACERTOS"].forEach(function (k) { pp[k] = true; });
  escreverAcesso_(ss, { UTIL: nome, PERFIL: "Membro", ATIVO: true, PESSOA: q.getResponseText().trim(), EMAIL: "", perms: pp, notif: {} });
  registar_("Utilizador criado (menu)", nome);
  ui.alert("Utilizador criado. Pode ver partilhadas, lançar compras e registar acertos. Para libertar mais páginas ou e-mails, mude as colunas para Sim na folha ACESSOS.");
}

function menuRedefinirSenha() {
  var ui = SpreadsheetApp.getUi();
  var n = ui.prompt("Redefinir senha", "Nome de utilizador:", ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  var p = ui.prompt("Nova senha provisória", "Pelo menos 8 caracteres:", ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  var usr = sheetObrig_(getSS_(), ABA_.USR), nome = n.getResponseText().trim(), senha = p.getResponseText().trim();
  if (senha.length < 8) { ui.alert("A senha precisa de pelo menos 8 caracteres."); return; }
  var v = usr.getLastRow() >= 2 ? usr.getRange(2, 1, usr.getLastRow() - 1, 1).getValues() : [];
  for (var i = 0; i < v.length; i++) if (norm_(v[i][0]) === norm_(nome)) { usr.getRange(i + 2, 2).setValue(hashSenha_(senha)); registar_("Senha redefinida (menu)", nome); ui.alert("Senha redefinida."); return; }
  ui.alert("Utilizador não encontrado.");
}

/* ========================================================================
 * API para a versão instalável (PWA no iPhone): o ecrã fica alojado fora do Apps Script e fala com este endereço.
 * Só passam as funções da lista; todas exigem sessão/permissão, exceto o login.
 * ====================================================================== */
var API_PUBLICA_ = [
  "ajustarSaldo",
  "alterarSenha",
  "aplicarCorrecaoParcelas",
  "arquivarCategoria",
  "atualizarCartao",
  "atualizarConta",
  "atualizarDivida",
  "atualizarPartilhada",
  "atualizarPrestacao",
  "atualizarRecorrente",
  "atualizarStatusPrestacao",
  "atualizarTransacao",
  "copiarOrcamentoAnterior",
  "desfazerPagamentoPrestacao",
  "enviarResumoTeste",
  "enviarTesteNotificacao",
  "excluirAcerto",
  "excluirPartilhada",
  "excluirRecorrente",
  "excluirTransacao",
  "executarRecorrentes",
  "fazerBackupAgora",
  "getAcessos",
  "getCartoes",
  "getCategorias",
  "getConsolidadoDividas",
  "getContas",
  "getDadosDashboard",
  "getDefinicoes",
  "getDespesasPorIntervalo",
  "getDividasDoMes",
  "getFaturasCartao",
  "getLog",
  "getMediaHistorica",
  "getMetasGastos",
  "getMeuAcesso",
  "getOpcoesFormulario",
  "getOrcadoVsRealizado",
  "getPartilhaMes",
  "getPartilhaOpcoes",
  "getQualidadeDados",
  "getRecorrentes",
  "getRelatorioDividas",
  "getResumoOrcamento",
  "getSeriesParcelas",
  "getTransacoesCartao",
  "getTransacoesContas",
  "getVisaoAnual",
  "pagarFaturaCartao",
  "pagarPrestacao",
  "pesquisarMovimentos",
  "reabrirFaturaCartao",
  "renomearCategoria",
  "salvarAcerto",
  "salvarAcesso",
  "salvarCartao",
  "salvarCategoria",
  "salvarConta",
  "salvarDefinicoes",
  "salvarMetaApp",
  "salvarPartilhada",
  "salvarRecorrente",
  "salvarTransacao",
  "salvarTransferencia",
  "terminarSessao",
  "validarLogin",
  "validarSessao"
];

function doPost(e) {
  var saida = function (o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); };
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = String(req.fn || ""), args = req.args instanceof Array ? req.args.slice(0, 6) : [];
    if (API_PUBLICA_.indexOf(fn) === -1) return saida({ err: "Pedido inválido." });
    var r = globalThis[fn].apply(null, args);
    return saida({ ok: r === undefined ? null : r });
  } catch (err) {
    return saida({ err: String((err && err.message) || err).replace(/^Error:\s*/, "") });
  }
}

function doGet() {
  var out = HtmlService.createHtmlOutputFromFile("Dashboard")
    .setTitle("Performance Financial")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
  try {
    out.setFaviconUrl("https://drive.google.com/uc?id=1rVv8WVC_c5TLpvfPQNJ-cDJx4TMG4L2B&export=download&format=png");
  } catch (e) {}
  return out;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Gestão Financeira")
    .addItem("🌐 Abrir Painel Web App", "mostrarLinkWeb")
    .addSeparator()
    .addItem("🛠 Preparar / atualizar estrutura (1.ª vez)", "migrarEstrutura")
    .addItem("⏰ Ativar rotinas automáticas", "ativarRotinasMenu")
    .addItem("💾 Fazer cópia de segurança agora", "backupMenu")
    .addSeparator()
    .addItem("👥 Acessos: abrir folha de permissões", "menuAbrirAcessos")
    .addItem("👥 Acessos: criar utilizador", "menuNovoUtilizador")
    .addItem("🔑 Acessos: redefinir senha de um utilizador", "menuRedefinirSenha")
    .addToUi();
}

function mostrarLinkWeb() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Segoe UI,sans-serif;padding:20px;text-align:center;">' +
    '<p style="margin-bottom:15px;font-size:14px;color:#333;">Clique no botão abaixo para abrir o seu painel financeiro:</p>' +
    '<a href="' + url + '" target="_blank" style="background:#7b1fa2;color:#fff;padding:10px 20px;text-decoration:none;border-radius:8px;font-weight:bold;display:inline-block;">Abrir Painel Web App</a>' +
    "</div>").setWidth(350).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, "Acesso ao Painel");
}

/* ========================================================================
 * UTILITÁRIOS
 * ====================================================================== */

function getSS_() {
  var id = PropertiesService.getScriptProperties().getProperty("SS_ID") || SS_ID_PADRAO;
  return SpreadsheetApp.openById(id);
}

function tz_() { return Session.getScriptTimeZone(); }

function norm_(s) {
  return String(s == null ? "" : s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim();
}

function num_(v) {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  var n = parseFloat(String(v == null ? "" : v).replace(/\s/g, "").replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function arred_(x) { return Math.round((x + 1e-9) * 100) / 100; }

function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (v === "" || v == null || typeof v === "number") return null;
  var s = String(v).trim(), m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function fmtData_(d) { return d ? Utilities.formatDate(d, tz_(), "dd/MM/yyyy") : ""; }
function fmtISO_(d) { return d ? Utilities.formatDate(d, tz_(), "yyyy-MM-dd") : ""; }
function inicioMes_(mes, ano) { return new Date(ano, mes - 1, 1); }
function fimMes_(mes, ano) { return new Date(ano, mes, 0, 23, 59, 59, 999); }
function mesAnterior_(mes, ano) { var d = new Date(ano, mes - 2, 1); return { mes: d.getMonth() + 1, ano: d.getFullYear() }; }

function mesPorTexto_(texto) {
  var letras = norm_(texto).replace(/[^a-z]/g, "");
  if (!letras) return 0;
  var melhor = 0, pos = 999;
  for (var i = 0; i < ABREV_.length; i++) {
    var p = letras.indexOf(ABREV_[i]);
    if (p !== -1 && p < pos) { pos = p; melhor = i + 1; }
  }
  return melhor;
}

/** Interpreta "12out.", "12out/2026", Date ou "10/2026". Sem ano explícito, infere-o pela data da compra. */
function parseFatura_(v, dataCompra) {
  if (v === "" || v == null) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : { mes: v.getMonth() + 1, ano: v.getFullYear() };
  var mes = mesPorTexto_(v);
  var ma = String(v).match(/(20\d{2})/);
  var ano = ma ? +ma[1] : 0;
  if (!mes) {
    var d = toDate_(v);
    return d ? { mes: d.getMonth() + 1, ano: d.getFullYear() } : null;
  }
  if (!ano) {
    if (dataCompra) { ano = dataCompra.getFullYear(); if (mes < dataCompra.getMonth() + 1) ano++; }
    else ano = new Date().getFullYear();
  }
  return { mes: mes, ano: ano };
}

function parseRef_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return { mes: v.getMonth() + 1, ano: v.getFullYear() };
  var m = String(v == null ? "" : v).trim().match(/^(\d{1,2})\/(\d{4})$/);
  return m ? { mes: +m[1], ano: +m[2] } : null;
}

function mesmoNome_(a, b) {
  var x = norm_(a), y = norm_(b);
  return x !== "" && x === y;
}

function comLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheetObrig_(ss, nome) {
  var s = ss.getSheetByName(nome);
  if (!s) throw new Error('Aba "' + nome + '" não encontrada.');
  return s;
}

/** Acrescenta várias linhas de uma vez (uma só escrita). Devolve a 1.ª linha escrita. */
function anexarLinhas_(sheet, linhas, formatos) {
  if (!linhas.length) return 0;
  var ini = Math.max(sheet.getLastRow() + 1, 2);
  if (formatos) Object.keys(formatos).forEach(function (col) { sheet.getRange(ini, +col, linhas.length, 1).setNumberFormat(formatos[col]); });
  sheet.getRange(ini, 1, linhas.length, linhas[0].length).setValues(linhas);
  return ini;
}

function anexarLinha_(sheet, valores, formatos) { return anexarLinhas_(sheet, [valores], formatos); }

function abaOuCria_(ss, nome, cab) {
  var s = ss.getSheetByName(nome);
  if (s) return s;
  s = ss.insertSheet(nome);
  s.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold");
  s.setFrozenRows(1);
  return s;
}

function novoId_(prefixo) { return prefixo + Utilities.getUuid().replace(/-/g, "").slice(0, 8); }
function pad2_(n) { return (n < 10 ? "0" : "") + n; }
function addMeses_(d, k) { var r = new Date(d.getFullYear(), d.getMonth() + k, 1); return dataSegura_(r.getFullYear(), r.getMonth() + 1, d.getDate()); }
function addMesRef_(ref, k) { var d = new Date(ref.ano, ref.mes - 1 + k, 1); return { mes: d.getMonth() + 1, ano: d.getFullYear() }; }

/** Registo de alterações (folha LOG). Nunca interrompe a operação principal. */
function registar_(acao, detalhe) {
  try {
    var s = abaOuCria_(getSS_(), ABA_.LOG, ["DATA/HORA", "UTILIZADOR", "AÇÃO", "DETALHE"]);
    anexarLinha_(s, [new Date(), USER_ATUAL_ || "sistema", acao, String(detalhe || "").slice(0, 480)], { 1: "dd/mm/yyyy hh:mm:ss" });
  } catch (e) { console.error("LOG: " + e); }
}

/* ========================================================================
 * AUTENTICAÇÃO
 * ====================================================================== */

function obterSal_() {
  var p = PropertiesService.getScriptProperties();
  var s = p.getProperty("SAL_SENHA");
  if (!s) { s = Utilities.getUuid(); p.setProperty("SAL_SENHA", s); }
  return s;
}

function hashSenha_(senha) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, obterSal_() + ":" + senha, Utilities.Charset.UTF_8);
  return "h1$" + bytes.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? "0" + v : v; }).join("");
}

/* Sessões longas ("manter sessão iniciada"): guardadas nas propriedades do script, validam-se a cada pedido (o acesso ativo/desativado é sempre confirmado). */
var SESSAO_LONGA_DIAS_ = 30;
function sessoesLongas_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("SESS_LONGAS") || "{}"); } catch (e) { return {}; } }
function gravarSessoesLongas_(o) { PropertiesService.getScriptProperties().setProperty("SESS_LONGAS", JSON.stringify(o)); }
function criarSessaoLonga_(token, nome) {
  var o = sessoesLongas_(), agora = Date.now();
  Object.keys(o).forEach(function (k) { if (o[k][1] < agora) delete o[k]; });
  o[token] = [nome, agora + SESSAO_LONGA_DIAS_ * 86400000];
  gravarSessoesLongas_(o);
}
function sessaoLonga_(token) { var x = sessoesLongas_()[token]; return x && x[1] > Date.now() ? x[0] : null; }
/** Termina as sessões longas de um utilizador (todas, ou todas menos "exceto"). */
function limparSessoesLongas_(nome, exceto) {
  var o = sessoesLongas_(), mudou = false;
  Object.keys(o).forEach(function (k) { if (k !== exceto && norm_(o[k][0]) === norm_(nome)) { delete o[k]; mudou = true; } });
  if (mudou) gravarSessoesLongas_(o);
}

function validarLogin(usuario, senha, lembrar) {
  usuario = String(usuario || "").trim();
  senha = String(senha || "").trim();
  var cache = CacheService.getScriptCache();
  var chaveTent = "tent_" + norm_(usuario);
  var tent = parseInt(cache.get(chaveTent) || "0", 10);
  if (tent >= MAX_TENTATIVAS_) return { sucesso: false, mensagem: "Demasiadas tentativas. Aguarde 10 minutos." };

  var sheet = getSS_().getSheetByName(ABA_.USR);
  if (!sheet || sheet.getLastRow() < 2) return { sucesso: false, mensagem: "Nenhum utilizador encontrado." };

  var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (norm_(dados[i][0]) !== norm_(usuario)) continue;
    var guardada = String(dados[i][1]).trim();
    var hashado = guardada.indexOf("h1$") === 0;
    var ok = hashado ? guardada === hashSenha_(senha) : guardada === senha;
    if (!ok) break;
    if (!hashado) sheet.getRange(i + 2, 2).setValue(hashSenha_(senha)); // migração automática
    var nome = String(dados[i][0]).trim();
    ACESSO_MEMO_ = {};
    var ac = acessoDe_(nome);
    if (!ac.ativo) return { sucesso: false, mensagem: "O seu acesso está desativado. Fale com o administrador." };
    cache.remove(chaveTent);
    var token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
    cache.put("sess_" + token, nome, SESSAO_SEGUNDOS_);
    if (lembrar === true) criarSessaoLonga_(token, nome);
    var pub = acessoPublico_(nome);
    return { sucesso: true, usuario: nome, token: token, perfil: pub.perfil, pessoa: pub.pessoa, perms: pub.perms };
  }
  cache.put(chaveTent, String(tent + 1), 600);
  return { sucesso: false, mensagem: "Utilizador ou senha incorretos." };
}

function validarSessao(token) {
  try { var u = exigirSessao_(token, "PROPRIO"); var r = acessoPublico_(u); return { sucesso: true, usuario: u, perfil: r.perfil, pessoa: r.pessoa, perms: r.perms }; }
  catch (e) { return { sucesso: false }; }
}

function terminarSessao(token) {
  if (token) {
    CacheService.getScriptCache().remove("sess_" + token);
    try { var o = sessoesLongas_(); if (o[token]) { delete o[token]; gravarSessoesLongas_(o); } } catch (e) {}
  }
  return true;
}

/* ========================================================================
 * ACESSOS (perfis e permissões)
 * O administrador é o 1.º utilizador da folha USUÁRIO. Os restantes só entram se estiverem
 * na folha ACESSOS com ATIVO = Sim, e só podem usar o que lá estiver marcado.
 * ====================================================================== */

var ABA_ACESSOS_ = "ACESSOS";

/** Catálogo de permissões: uma coluna por item na folha ACESSOS. Para o administrador está tudo sempre ligado. */
var PERM_DEF_ = [
  { k: "DASHBOARD_VER", t: "VER DASHBOARD", g: "Páginas" },
  { k: "CONTAS_VER", t: "VER CONTAS", g: "Páginas" },
  { k: "MOVIMENTOS_VER", t: "VER MOVIMENTOS", g: "Páginas" },
  { k: "CARTOES_VER", t: "VER CARTÕES", g: "Páginas" },
  { k: "ANUAL_VER", t: "VER VISÃO ANUAL", g: "Páginas" },
  { k: "DIVIDAS_VER", t: "VER RELATÓRIO DÍVIDAS", g: "Relatórios" },
  { k: "REL_INTERVALO", t: "VER REL. DESPESAS POR INTERVALO", g: "Relatórios" },
  { k: "REL_METAS", t: "VER REL. METAS E ALERTAS", g: "Relatórios" },
  { k: "REL_MEDIA", t: "VER REL. MÉDIA HISTÓRICA", g: "Relatórios" },
  { k: "REL_ORCADO", t: "VER REL. ORÇADO VS REALIZADO", g: "Relatórios" },
  { k: "LANCAR", t: "LANÇAR / EDITAR / APAGAR LANÇAMENTOS", g: "Alterar dados" },
  { k: "CONTAS_GERIR", t: "GERIR CONTAS E CARTÕES (criar, editar, pagar faturas, acertar saldo)", g: "Alterar dados" },
  { k: "DIVIDAS_GERIR", t: "GERIR DÍVIDAS (registar, pagar, editar)", g: "Alterar dados" },
  { k: "METAS_EDITAR", t: "EDITAR METAS E ORÇAMENTOS", g: "Alterar dados" },
  { k: "DEFINICOES", t: "DEFINIÇÕES (categorias, recorrentes, notificações, cópias, registo)", g: "Alterar dados" },
  { k: "PARTILHADAS_VER", t: "VER PARTILHADAS", g: "Despesas partilhadas" },
  { k: "PARTILHADAS_LANCAR", t: "LANÇAR PARTILHADAS", g: "Despesas partilhadas" },
  { k: "ACERTOS", t: "REGISTAR ACERTOS", g: "Despesas partilhadas" }
];
var PERMS_ = PERM_DEF_.map(function (p) { return p.k; });
/** Quem recebe cada tipo de e-mail (o dono recebe sempre tudo, nos e-mails das Definições). */
var NOTIF_DEF_ = [
  { k: "METAS", t: "RECEBE ALERTAS DE METAS", n: "Alertas de metas / orçamento" },
  { k: "VENCIMENTOS", t: "RECEBE LEMBRETES DE VENCIMENTO", n: "Lembretes de vencimentos" },
  { k: "SEMANAL", t: "RECEBE RESUMO SEMANAL", n: "Resumo semanal" },
  { k: "MENSAL", t: "RECEBE RESUMO MENSAL", n: "Resumo mensal" }
];
var CAB_ACESSOS_ = [{ k: "UTIL", t: "UTILIZADOR" }, { k: "PERFIL", t: "PERFIL" }, { k: "ATIVO", t: "ATIVO" }, { k: "PESSOA", t: "PESSOA" }, { k: "EMAIL", t: "E-MAIL" }]
  .concat(PERM_DEF_.map(function (p) { return { k: p.k, t: p.t }; }), NOTIF_DEF_.map(function (p) { return { k: "N_" + p.k, t: p.t }; }));
var ACESSO_MEMO_ = {};

function simNao_(v) { return norm_(v) === "sim"; }

/** Folha ACESSOS: cria-a ou completa os cabeçalhos em falta. Devolve {s, col} (col: chave → coluna 1-based). */
function garantirAcessosCols_(ss) {
  var s = ss.getSheetByName(ABA_ACESSOS_), nova = !s;
  if (!s) s = ss.insertSheet(ABA_ACESSOS_);
  var ncol = Math.max(s.getLastColumn(), 1);
  var cab = s.getRange(1, 1, 1, ncol).getValues()[0].map(function (x) { return norm_(x); });
  var col = {}, falta = [];
  CAB_ACESSOS_.forEach(function (c) { var i = cab.indexOf(norm_(c.t)); if (i === -1) falta.push(c); else col[c.k] = i + 1; });
  if (falta.length) {
    var ini = cab.some(function (x) { return x; }) ? s.getLastColumn() + 1 : 1;
    if (s.getMaxColumns() < ini + falta.length - 1) s.insertColumnsAfter(s.getMaxColumns(), ini + falta.length - 1 - s.getMaxColumns());
    s.getRange(1, ini, 1, falta.length).setValues([falta.map(function (c) { return c.t; })]).setFontWeight("bold").setWrap(true).setVerticalAlignment("top");
    falta.forEach(function (c, i) { col[c.k] = ini + i; });
    try {
      s.setFrozenRows(1); s.setFrozenColumns(1);
      var lista = function (v) { return SpreadsheetApp.newDataValidation().requireValueInList(v, true).setAllowInvalid(false).build(); };
      falta.forEach(function (c) {
        var r = s.getRange(2, col[c.k], 60, 1);
        if (c.k === "PERFIL") r.setDataValidation(lista(["Administrador", "Membro"]));
        else if (c.k !== "UTIL" && c.k !== "PESSOA" && c.k !== "EMAIL") r.setDataValidation(lista(["Sim", "Não"]));
      });
      s.setColumnWidth(col.UTIL, 130); s.setColumnWidth(col.EMAIL, 200);
    } catch (e) {}
  }
  if (nova || s.getLastRow() < 2) {
    var dono = "";
    try { dono = String(ss.getSheetByName(ABA_.USR).getRange(2, 1).getValue() || "").trim(); } catch (e) {}
    var o = { UTIL: dono || "admin", PERFIL: "Administrador", ATIVO: "Sim", PESSOA: "Fellipe", EMAIL: "" };
    PERMS_.forEach(function (k) { o[k] = "Sim"; }); NOTIF_DEF_.forEach(function (n) { o["N_" + n.k] = "Sim"; });
    var l = Math.max(s.getLastRow() + 1, 2);
    Object.keys(o).forEach(function (k) { if (col[k]) s.getRange(l, col[k]).setValue(o[k]); });
  }
  return { s: s, col: col };
}
function garantirAcessos_(ss) { return garantirAcessosCols_(ss).s; }

/** Lê a folha ACESSOS como lista de objetos {linha, chave → valor}. */
function lerAcessos_(ss) {
  var g = garantirAcessosCols_(ss), s = g.s, n = s.getLastRow() - 1, out = [];
  if (n < 1) return { s: s, col: g.col, linhas: out };
  var v = s.getRange(2, 1, n, s.getLastColumn()).getValues();
  v.forEach(function (r, i) {
    var o = { _linha: i + 2 };
    Object.keys(g.col).forEach(function (k) { o[k] = r[g.col[k] - 1]; });
    if (String(o.UTIL || "").trim()) out.push(o);
  });
  return { s: s, col: g.col, linhas: out };
}

function acessoDe_(user) {
  var k = norm_(user);
  if (ACESSO_MEMO_[k]) return ACESSO_MEMO_[k];
  var ss = getSS_(), usr = ss.getSheetByName(ABA_.USR);
  var primeiro = usr && usr.getLastRow() >= 2 ? norm_(usr.getRange(2, 1).getValue()) : "";
  var lst = lerAcessos_(ss).linhas;
  var res = { admin: false, ativo: false, pessoa: "", email: "", perms: {} };
  PERMS_.forEach(function (p) { res.perms[p] = false; });
  for (var i = 0; i < lst.length; i++) {
    if (norm_(lst[i].UTIL) !== k) continue;
    res.admin = norm_(lst[i].PERFIL) === "administrador"; res.ativo = simNao_(lst[i].ATIVO); res.pessoa = String(lst[i].PESSOA || "").trim(); res.email = String(lst[i].EMAIL || "").trim();
    PERMS_.forEach(function (p) { res.perms[p] = simNao_(lst[i][p]); });
    break;
  }
  if (k && k === primeiro) { res.admin = true; res.ativo = true; if (!res.pessoa) res.pessoa = "Fellipe"; }   // o dono nunca fica trancado de fora
  if (res.admin) PERMS_.forEach(function (p) { res.perms[p] = true; });
  return (ACESSO_MEMO_[k] = res);
}

function acessoPublico_(user) {
  var a = acessoDe_(user);
  return { perfil: a.admin ? "Administrador" : "Membro", pessoa: a.pessoa, perms: a.perms };
}

/** Todas as funções públicas passam por aqui. Sem "perm", só o administrador entra; "PROPRIO" = qualquer utilizador ativo;
 *  "perm" pode ser uma chave ou uma lista de chaves (basta ter uma). */
function exigirSessao_(token, perm) {
  var msg = "Sessão expirada. Inicie sessão novamente.";
  if (!token) throw new Error(msg);
  var cache = CacheService.getScriptCache();
  var u = cache.get("sess_" + token) || sessaoLonga_(token);
  if (!u) throw new Error(msg);
  cache.put("sess_" + token, u, SESSAO_SEGUNDOS_);
  USER_ATUAL_ = u;
  ACESSO_MEMO_ = {};
  var a = acessoDe_(u);
  if (!a.ativo) throw new Error("O seu acesso está desativado. Fale com o administrador.");
  if (a.admin || perm === "PROPRIO") return u;
  var lista = perm ? (perm instanceof Array ? perm : [perm]) : [];
  for (var i = 0; i < lista.length; i++) if (a.perms[lista[i]]) return u;
  throw new Error("Não tem permissão para esta ação.");
}

function getMeuAcesso(token) {
  var u = exigirSessao_(token, "PROPRIO");
  return acessoPublico_(u);
}

function getAcessos(token) {
  exigirSessao_(token);                                   // só administrador
  var ss = getSS_(), usr = ss.getSheetByName(ABA_.USR), lst = lerAcessos_(ss).linhas;
  var nomes = usr && usr.getLastRow() >= 2 ? usr.getRange(2, 1, usr.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }).filter(Boolean) : [];
  var porNome = {}; lst.forEach(function (r) { porNome[norm_(r.UTIL)] = r; });
  var usuarios = nomes.map(function (n, i) {
    var r = porNome[norm_(n)], dono = i === 0, perms = {}, notif = {};
    PERMS_.forEach(function (p) { perms[p] = dono || (r ? simNao_(r[p]) : false); });
    NOTIF_DEF_.forEach(function (x) { notif[x.k] = r ? simNao_(r["N_" + x.k]) : false; });
    return { utilizador: n, dono: dono, perfil: dono ? "Administrador" : (r ? String(r.PERFIL || "Membro") : "Membro"), ativo: dono ? true : (r ? simNao_(r.ATIVO) : false),
      pessoa: r ? String(r.PESSOA || "").trim() : "", email: r ? String(r.EMAIL || "").trim() : "", perms: perms, notif: notif };
  });
  return { usuarios: usuarios, catalogo: PERM_DEF_.map(function (p) { return { k: p.k, t: p.t.charAt(0) + p.t.slice(1).toLowerCase(), g: p.g }; }), notif: NOTIF_DEF_ };
}

/** Escreve (cria ou atualiza) a linha de um utilizador na folha ACESSOS. o = {UTIL, PERFIL, ATIVO, PESSOA, EMAIL, perms:{}, notif:{}} */
function escreverAcesso_(ss, o) {
  var g = lerAcessos_(ss), s = g.s, col = g.col, linha = 0;
  g.linhas.forEach(function (r) { if (norm_(r.UTIL) === norm_(o.UTIL)) linha = r._linha; });
  if (!linha) linha = Math.max(s.getLastRow() + 1, 2);
  var sn = function (b) { return b ? "Sim" : "Não"; }, p = o.perms || {}, nt = o.notif || {};
  var vals = { UTIL: o.UTIL, PERFIL: o.PERFIL, ATIVO: sn(o.ATIVO), PESSOA: o.PESSOA || "", EMAIL: o.EMAIL || "" };
  PERMS_.forEach(function (k) { vals[k] = sn(p[k]); });
  NOTIF_DEF_.forEach(function (x) { vals["N_" + x.k] = sn(nt[x.k]); });
  Object.keys(vals).forEach(function (k) { s.getRange(linha, col[k]).setValue(vals[k]); });
  return linha;
}

/** Cria ou atualiza um utilizador, as suas permissões e os e-mails que recebe. A senha só é pedida ao criar (ou ao redefinir). */
function salvarAcesso(token, d) {
  var quem = exigirSessao_(token);
  var nome = String(d.utilizador || "").trim(), senha = String(d.senha || "").trim(), email = String(d.email || "").trim();
  if (nome.length < 2 || nome.length > 30 || /[,;]/.test(nome)) throw new Error("Nome de utilizador inválido (2 a 30 caracteres, sem vírgulas).");
  if (email && !/^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$/.test(email)) throw new Error("E-mail inválido (use um só endereço).");
  return comLock_(function () {
    var ss = getSS_(), usr = sheetObrig_(ss, ABA_.USR);
    var nUsr = usr.getLastRow() >= 2 ? usr.getRange(2, 1, usr.getLastRow() - 1, 1).getValues() : [], lUsr = -1;
    nUsr.forEach(function (r, i) { if (norm_(r[0]) === norm_(nome)) lUsr = i + 2; });
    var dono = lUsr === 2;
    if (lUsr === -1) {
      if (senha.length < 8) throw new Error("Defina uma senha provisória com pelo menos 8 caracteres.");
      usr.getRange(usr.getLastRow() + 1, 1, 1, 3).setValues([[nome, hashSenha_(senha), ""]]);
    } else if (senha) {
      if (senha.length < 8) throw new Error("A senha deve ter pelo menos 8 caracteres.");
      usr.getRange(lUsr, 2).setValue(hashSenha_(senha));
      limparSessoesLongas_(nome);
    }
    var perfil = dono ? "Administrador" : (d.perfil === "Administrador" ? "Administrador" : "Membro");
    var ativo = dono ? true : d.ativo !== false;
    escreverAcesso_(ss, { UTIL: nome, PERFIL: perfil, ATIVO: ativo, PESSOA: String(d.pessoa || "").trim(), EMAIL: email, perms: d.perms || {}, notif: d.notif || {} });
    registar_(lUsr === -1 ? "Utilizador criado" : "Acessos atualizados", nome + " · " + perfil + (ativo ? "" : " · desativado") + (senha && lUsr !== -1 ? " · senha redefinida" : ""));
    return true;
  });
}

/** Endereços que recebem um tipo de e-mail: o dono (Definições) recebe tudo; os restantes só o que marcarem na folha ACESSOS. */
function destinatarios_(tipo) {
  var out = [], vistos = {}, add = function (e) { e = String(e || "").trim(); var k = e.toLowerCase(); if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e) && !vistos[k]) { vistos[k] = true; out.push(e); } };
  (obterEmailUsuario_() || "").split(",").forEach(add);
  if (tipo && ["METAS", "VENCIMENTOS", "SEMANAL", "MENSAL"].indexOf(tipo) !== -1) {
    try {
      lerAcessos_(getSS_()).linhas.forEach(function (r) { if (simNao_(r.ATIVO) && simNao_(r["N_" + tipo])) add(r.EMAIL); });
    } catch (e) { console.error("Destinatários: " + e); }
  }
  return out;
}

/* ========================================================================
 * DESPESAS PARTILHADAS
 * Dados na planilha partilhada (propriedade SS_PART; por omissão a "DESPESAS CONJUNTAS 2026"):
 * PARTILHADAS (lançamentos), ACERTOS (transferências), PESSOAS, CONTAS_CARTOES, CATEGORIAS_P.
 * Regras: Casa = (casa − receitas fixas) ÷ 2 para cada titular; Compra = quem paga é o titular do cartão e o outro
 * deve a sua parte (Igual = 50%, "100% Nome" = esse nome deve tudo).
 * ====================================================================== */

var SS_PART_PADRAO = "17-ut2AisGbgkFjxgGP9CPijrLLvxkIoi8pqT8NRByvY";
var ABA_PART_ = "PARTILHADAS", ABA_PACERTOS_ = "ACERTOS", ABA_PPESSOAS_ = "PESSOAS", ABA_PCONTAS_ = "CONTAS_CARTOES", ABA_PCATS_ = "CATEGORIAS_P";
var TIPOS_PART_ = ["Casa", "Compra", "Receita"];

function getSSP_() {
  var id = PropertiesService.getScriptProperties().getProperty("SS_PART") || SS_PART_PADRAO;
  return SpreadsheetApp.openById(id);
}

function ultimaLinhaCol_(sh, col) {
  var n = sh.getMaxRows(), v = sh.getRange(1, col, n, 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim() !== "") return i + 1;
  return 0;
}

function refTxt_(v) {
  if (v instanceof Date) return v.getFullYear() + "-" + (v.getMonth() < 9 ? "0" : "") + (v.getMonth() + 1);
  var s = String(v == null ? "" : v).trim(), m = s.match(/^(\d{4})-(\d{1,2})/);
  return m ? m[1] + "-" + (m[2].length === 1 ? "0" : "") + m[2] : s;
}
function rotuloRef_(ref) { var p = ref.split("-"); return MESES_[+p[1] - 1] + "/" + p[0]; }
function addRef_(ref, k) { var p = ref.split("-"), d = new Date(+p[0], +p[1] - 1 + k, 1); return d.getFullYear() + "-" + (d.getMonth() < 9 ? "0" : "") + (d.getMonth() + 1); }

function partPessoas_(ssp) {
  var s = ssp.getSheetByName(ABA_PPESSOAS_), tit = [], outros = [];
  (s && s.getLastRow() >= 2 ? s.getRange(2, 1, s.getLastRow() - 1, 6).getValues() : []).forEach(function (r) {
    var nome = String(r[1] || "").trim(); if (!nome || !simNao_(r[5])) return;
    (norm_(r[2]).indexOf("titular") === 0 ? tit : outros).push(nome);
  });
  if (tit.length !== 2) tit = ["Fellipe", "Kenneth"];
  return { tit: tit, outros: outros };
}

function partCartoes_(ssp) {
  var s = ssp.getSheetByName(ABA_PCONTAS_), out = [];
  (s && s.getLastRow() >= 2 ? s.getRange(2, 1, s.getLastRow() - 1, 7).getValues() : []).forEach(function (r) {
    var nome = String(r[0] || "").trim(); if (!nome || !simNao_(r[6])) return;
    out.push({ nome: nome, tipo: String(r[1] || "").trim(), titular: String(r[2] || "").trim(), fech: +r[3] || 30, venc: +r[4] || 12, proc: +r[5] || 0 });
  });
  return out;
}

function partLinhas_(ssp) {
  var sh = ssp.getSheetByName(ABA_PART_); if (!sh) throw new Error('Aba "' + ABA_PART_ + '" não encontrada na planilha partilhada.');
  var n = ultimaLinhaCol_(sh, 1); if (n < 2) return [];
  var out = [];
  sh.getRange(2, 1, n - 1, 16).getValues().forEach(function (r, i) {
    var id = String(r[0] || "").trim(); if (!id) return;
    out.push({ linha: i + 2, id: id, ref: refTxt_(r[1]), data: toDate_(r[2]), tipo: String(r[3] || "").trim(), desc: String(r[4] || "").trim(), cat: String(r[5] || "").trim(),
      parcela: String(r[6] || "").trim(), valor: num_(r[7]), cartao: String(r[8] || "").trim(), pagou: String(r[9] || "").trim(), div: String(r[10] || "").trim(),
      obs: String(r[13] || "").trim(), origem: String(r[14] || "").trim(), criadoPor: String(r[15] || "").trim() });
  });
  return out;
}

function partAcertos_(ssp) {
  var sh = ssp.getSheetByName(ABA_PACERTOS_); if (!sh) return [];
  var n = ultimaLinhaCol_(sh, 1); if (n < 2) return [];
  var out = [];
  sh.getRange(2, 1, n - 1, 9).getValues().forEach(function (r, i) {
    var id = String(r[0] || "").trim(); if (!id) return;
    out.push({ linha: i + 2, id: id, data: toDate_(r[1]), ref: refTxt_(r[2]), tipo: String(r[3] || "").trim(), de: String(r[4] || "").trim(), para: String(r[5] || "").trim(), valor: num_(r[6]), obs: String(r[7] || "").trim(), criadoPor: String(r[8] || "").trim() });
  });
  return out;
}

/** Parte de cada titular numa linha. */
function partes_(x, A, B) {
  var v = x.valor, p = {};
  if (x.tipo === "Receita") { p[A] = 0; p[B] = 0; return p; }
  var d = norm_(x.div);
  if (d === norm_("100% " + A)) { p[A] = v; p[B] = 0; }
  else if (d === norm_("100% " + B)) { p[A] = 0; p[B] = v; }
  else { p[A] = v / 2; p[B] = v / 2; }
  return p;
}

function refPartilha_(ct, data) {
  var r = faturaAbertaRef_(data, ct.venc, ct.fech, ct.proc);
  return r.ano + "-" + (r.mes < 10 ? "0" : "") + r.mes;
}

/** Cálculo de um mês de fatura. */
function resumoPartilha_(rows, acertos, ref, tit) {
  var A = tit[0], B = tit[1], R = { receitas: 0, casa: 0, compras: 0, pedro: 0, devBA: 0, devAB: 0 };
  var casa = [], rec = [], compras = [], porCartao = {};
  rows.forEach(function (x) {
    if (x.ref !== ref) return;
    var p = partes_(x, A, B);
    if (x.tipo === "Receita") { R.receitas += x.valor; if (norm_(x.pagou) === "pedro") R.pedro += x.valor; rec.push(x); }
    else if (x.tipo === "Casa") { R.casa += x.valor; casa.push(x); }
    else if (x.tipo === "Compra") {
      R.compras += x.valor; compras.push(x);
      porCartao[x.cartao] = (porCartao[x.cartao] || 0) + x.valor;
      if (norm_(x.pagou) === norm_(A)) R.devBA += p[B];
      else if (norm_(x.pagou) === norm_(B)) R.devAB += p[A];
    }
  });
  var BA = 0, AB = 0, cc = {}; cc[A] = 0; cc[B] = 0;
  var ac = acertos.filter(function (a) { return a.ref === ref; });
  ac.forEach(function (a) {
    if (a.tipo === "Entre nós") { if (norm_(a.de) === norm_(B) && norm_(a.para) === norm_(A)) BA += a.valor; else if (norm_(a.de) === norm_(A) && norm_(a.para) === norm_(B)) AB += a.valor; }
    else if (a.tipo === "Conta conjunta") { if (cc[a.de] !== undefined) cc[a.de] += a.valor; }
  });
  var aPart = R.casa - R.receitas, cadaUm = aPart / 2, liquido = R.devBA - R.devAB, acertado = BA - AB;
  return {
    ref: ref, rotulo: rotuloRef_(ref), receitas: arred_(R.receitas), casa: arred_(R.casa), compras: arred_(R.compras), totalDespesas: arred_(R.casa + R.compras),
    aPartilhar: arred_(aPart), cadaUm: arred_(cadaUm), pedro: arred_(R.pedro),
    devBA: arred_(R.devBA), devAB: arred_(R.devAB), liquido: arred_(liquido), acertado: arred_(acertado), emAberto: arred_(liquido - acertado),
    contaConjuntaPago: { a: arred_(cc[A]), b: arred_(cc[B]) },
    porCartao: Object.keys(porCartao).map(function (k) { return { cartao: k, total: arred_(porCartao[k]) }; }),
    _casa: casa, _rec: rec, _compras: compras, _acertos: ac
  };
}

function linhaPart_(x, A, B, a) {
  var p = partes_(x, A, B);
  var edit = a.admin || (x.tipo === "Compra" && a.perms.PARTILHADAS_LANCAR && norm_(x.pagou) === norm_(a.pessoa));
  return { id: x.id, ref: x.ref, data: fmtData_(x.data), iso: fmtISO_(x.data), tipo: x.tipo, desc: x.desc, cat: x.cat, parcela: x.parcela, valor: arred_(x.valor), cartao: x.cartao, pagou: x.pagou, div: x.div,
    parteA: arred_(p[A]), parteB: arred_(p[B]), obs: x.obs, origem: x.origem, editavel: !!edit };
}

function partPessoaAtual_(a, tit) {
  if (a.pessoa) return a.pessoa;
  return tit[0];
}

function getPartilhaOpcoes(token) {
  var user = exigirSessao_(token, "PARTILHADAS_VER"), a = acessoDe_(user);
  var ssp = getSSP_(), pe = partPessoas_(ssp), cats = [], sc = ssp.getSheetByName(ABA_PCATS_);
  (sc && sc.getLastRow() >= 2 ? sc.getRange(2, 1, sc.getLastRow() - 1, 2).getValues() : []).forEach(function (r) { if (String(r[0]).trim()) cats.push({ nome: String(r[0]).trim(), tipo: String(r[1]).trim() }); });
  return { titulares: pe.tit, outros: pe.outros, cartoes: partCartoes_(ssp), categorias: cats, eu: { utilizador: user, pessoa: partPessoaAtual_(a, pe.tit), admin: a.admin, perms: a.perms } };
}

function getPartilhaMes(token, ref) {
  var user = exigirSessao_(token, "PARTILHADAS_VER"), a = acessoDe_(user);
  var ssp = getSSP_(), pe = partPessoas_(ssp), tit = pe.tit, rows = partLinhas_(ssp), acertos = partAcertos_(ssp);
  var refs = {};
  rows.forEach(function (x) { if (x.ref) refs[x.ref] = true; });
  acertos.forEach(function (x) { if (x.ref) refs[x.ref] = true; });
  var hoje = new Date(); refs[hoje.getFullYear() + "-" + (hoje.getMonth() < 9 ? "0" : "") + (hoje.getMonth() + 1)] = true;
  var lista = Object.keys(refs).sort();
  if (!ref || !refs[ref]) ref = lista.filter(function (r) { return rows.some(function (x) { return x.ref === r; }); }).pop() || lista[lista.length - 1];
  var acum = 0;
  lista.forEach(function (r) { var m = resumoPartilha_(rows, acertos, r, tit); acum += m.emAberto; });
  var m = resumoPartilha_(rows, acertos, ref, tit), ord = function (l) { return l.sort(function (x, y) { return (fmtISO_(y.data) > fmtISO_(x.data)) ? 1 : -1; }); };
  var conv = function (l) { return ord(l).map(function (x) { return linhaPart_(x, tit[0], tit[1], a); }); };
  var out = {
    tit: tit, meses: lista.map(function (r) { return { ref: r, rotulo: rotuloRef_(r) }; }), ref: ref,
    resumo: {
      rotulo: m.rotulo, receitas: m.receitas, casa: m.casa, compras: m.compras, totalDespesas: m.totalDespesas, aPartilhar: m.aPartilhar, cadaUm: m.cadaUm, pedro: m.pedro,
      devBA: m.devBA, devAB: m.devAB, liquido: m.liquido, acertado: m.acertado, emAberto: m.emAberto, contaConjuntaPago: m.contaConjuntaPago, porCartao: m.porCartao
    },
    acumulado: arred_(acum),
    casa: conv(m._casa), receitas: conv(m._rec), compras: conv(m._compras),
    acertos: m._acertos.sort(function (x, y) { return fmtISO_(y.data) > fmtISO_(x.data) ? 1 : -1; }).map(function (x) {
      return { id: x.id, data: fmtData_(x.data), iso: fmtISO_(x.data), tipo: x.tipo, de: x.de, para: x.para, valor: arred_(x.valor), obs: x.obs, editavel: a.admin || norm_(x.criadoPor) === norm_(user) };
    })
  };
  // O que cada titular transfere neste mês (conta conjunta + diferença das compras)
  var A = tit[0], B = tit[1], liq = m.emAberto;
  out.transferir = [
    { pessoa: A, contaConjunta: m.cadaUm, paraOutro: liq < 0 ? arred_(-liq) : 0, receberOutro: liq > 0 ? arred_(liq) : 0 },
    { pessoa: B, contaConjunta: m.cadaUm, paraOutro: liq > 0 ? arred_(liq) : 0, receberOutro: liq < 0 ? arred_(-liq) : 0 }
  ];
  return out;
}

function validarLinhaPart_(d, a, pe, cartoes) {
  var A = pe.tit[0], B = pe.tit[1];
  var tipo = String(d.tipo || ""); if (TIPOS_PART_.indexOf(tipo) === -1) throw new Error("Tipo inválido.");
  var data = toDate_(d.data); if (!data) throw new Error("Data inválida.");
  var valor = arred_(num_(d.valor)); if (!(Math.abs(valor) > 0)) throw new Error("Indique um valor diferente de zero.");
  if (tipo !== "Compra" && valor < 0) throw new Error("Só compras podem ter valor negativo (estorno).");
  var desc = String(d.desc || "").trim(); if (!desc) throw new Error("Indique a descrição.");
  var ref = refTxt_(d.ref); if (!/^\d{4}-\d{2}$/.test(ref)) throw new Error("Mês de referência inválido.");
  var cartao = String(d.cartao || "").trim(), pagou = String(d.pagou || "").trim(), div = String(d.div || "").trim();
  if (!a.admin && tipo !== "Compra") throw new Error("Só pode lançar compras. Despesas da casa e receitas fixas são do administrador.");
  if (tipo === "Compra") {
    var ct = cartoes.filter(function (x) { return norm_(x.nome) === norm_(cartao); })[0];
    if (!ct) throw new Error("Escolha um cartão válido.");
    if (!a.admin) pagou = a.pessoa;
    if (pe.tit.map(norm_).indexOf(norm_(pagou)) === -1) throw new Error("Quem pagou tem de ser " + A + " ou " + B + ".");
    pagou = pe.tit.filter(function (t) { return norm_(t) === norm_(pagou); })[0];
    if (["igual", norm_("100% " + A), norm_("100% " + B)].indexOf(norm_(div)) === -1) throw new Error("Divisão inválida.");
    div = norm_(div) === "igual" ? "Igual" : (norm_(div) === norm_("100% " + A) ? "100% " + A : "100% " + B);
    cartao = ct.nome;
  } else if (tipo === "Casa") { cartao = "Conta conjunta"; pagou = "Conta conjunta"; div = "Igual"; }
  else { cartao = "Conta conjunta"; if (!pagou) pagou = "Conta conjunta"; div = "—"; }
  return { tipo: tipo, data: data, valor: valor, desc: desc, ref: ref, cartao: cartao, pagou: pagou, div: div, cat: String(d.cat || "").trim() || (tipo === "Compra" ? "Outros" : tipo === "Receita" ? "Receita fixa" : "Casa"), parcela: String(d.parcela || "").trim() };
}

function proximoIdPart_(sh, prefixo, col) {
  var n = ultimaLinhaCol_(sh, col), max = 0;
  if (n >= 2) sh.getRange(2, col, n - 1, 1).getValues().forEach(function (r) { var m = String(r[0]).match(/(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  var s = String(max + 1); while (s.length < 4) s = "0" + s;
  return prefixo + "-" + s;
}

function escreverLinhaPart_(sh, linha, v, id, criadoPor, origem, obs) {
  sh.getRange(linha, 2).setNumberFormat("@");
  sh.getRange(linha, 3).setNumberFormat("dd/mm/yyyy");
  sh.getRange(linha, 8).setNumberFormat('#,##0.00 "€"');
  sh.getRange(linha, 1, 1, 11).setValues([[id, v.ref, v.data, v.tipo, v.desc, v.cat, v.parcela, v.valor, v.cartao, v.pagou, v.div]]);
  sh.getRange(linha, 14, 1, 3).setValues([[obs || "", origem || "App", criadoPor || ""]]);
}

function salvarPartilhada(token, d) {
  var user = exigirSessao_(token, "PARTILHADAS_LANCAR"), a = acessoDe_(user);
  return comLock_(function () {
    var ssp = getSSP_(), pe = partPessoas_(ssp), v = validarLinhaPart_(d, a, pe, partCartoes_(ssp));
    var sh = sheetObrig_(ssp, ABA_PART_), id = proximoIdPart_(sh, "P", 1), linha = Math.max(ultimaLinhaCol_(sh, 1) + 1, 2);
    escreverLinhaPart_(sh, linha, v, id, user, "App", "");
    registar_("Partilhada lançada", v.tipo + " · " + v.desc + " · " + v.valor + " · " + v.ref);
    return id;
  });
}

function acharLinhaPart_(sh, id) {
  var n = ultimaLinhaCol_(sh, 1); if (n < 2) return 0;
  var v = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === id) return i + 2;
  return 0;
}

function podeMexerPart_(a, x) {
  return a.admin || (x.tipo === "Compra" && a.perms.PARTILHADAS_LANCAR && norm_(x.pagou) === norm_(a.pessoa));
}

function atualizarPartilhada(token, d) {
  var user = exigirSessao_(token, "PARTILHADAS_LANCAR"), a = acessoDe_(user);
  return comLock_(function () {
    var ssp = getSSP_(), sh = sheetObrig_(ssp, ABA_PART_), l = acharLinhaPart_(sh, String(d.id || ""));
    if (!l) throw new Error("Lançamento não encontrado. Atualize a página.");
    var atual = partLinhas_(ssp).filter(function (x) { return x.linha === l; })[0];
    if (!podeMexerPart_(a, atual)) throw new Error("Só pode alterar as suas próprias compras.");
    var pe = partPessoas_(ssp), dd = {}; Object.keys(d).forEach(function (k) { dd[k] = d[k]; });
    dd.tipo = a.admin ? d.tipo : "Compra";
    var v = validarLinhaPart_(dd, a, pe, partCartoes_(ssp));
    escreverLinhaPart_(sh, l, v, atual.id, atual.criadoPor, atual.origem, a.admin && d.obs !== undefined ? String(d.obs) : atual.obs);
    registar_("Partilhada editada", atual.id + " · " + v.desc + " · " + v.valor);
    return true;
  });
}

function excluirPartilhada(token, id) {
  var user = exigirSessao_(token, "PARTILHADAS_LANCAR"), a = acessoDe_(user);
  return comLock_(function () {
    var ssp = getSSP_(), sh = sheetObrig_(ssp, ABA_PART_), l = acharLinhaPart_(sh, String(id || ""));
    if (!l) throw new Error("Lançamento não encontrado. Atualize a página.");
    var atual = partLinhas_(ssp).filter(function (x) { return x.linha === l; })[0];
    if (!podeMexerPart_(a, atual)) throw new Error("Só pode apagar as suas próprias compras.");
    sh.deleteRow(l);
    registar_("Partilhada apagada", atual.id + " · " + atual.desc + " · " + atual.valor);
    return true;
  });
}

function salvarAcerto(token, d) {
  var user = exigirSessao_(token, "ACERTOS"), a = acessoDe_(user);
  return comLock_(function () {
    var ssp = getSSP_(), pe = partPessoas_(ssp), pessoas = pe.tit.concat(pe.outros, ["Conta conjunta"]);
    var tipo = String(d.tipo || ""); if (["Conta conjunta", "Entre nós"].indexOf(tipo) === -1) throw new Error("Tipo de acerto inválido.");
    var data = toDate_(d.data); if (!data) throw new Error("Data inválida.");
    var valor = arred_(num_(d.valor)); if (!(valor > 0)) throw new Error("Indique um valor maior que zero.");
    var ref = refTxt_(d.ref); if (!/^\d{4}-\d{2}$/.test(ref)) throw new Error("Mês de referência inválido.");
    var de = pessoas.filter(function (p) { return norm_(p) === norm_(d.de); })[0], para = pessoas.filter(function (p) { return norm_(p) === norm_(d.para); })[0];
    if (!de || !para || norm_(de) === norm_(para)) throw new Error("Indique quem paga e quem recebe.");
    if (tipo === "Entre nós" && (pe.tit.indexOf(de) === -1 || pe.tit.indexOf(para) === -1)) throw new Error("Acertos entre nós são só entre os dois titulares.");
    if (tipo === "Conta conjunta" && para !== "Conta conjunta") throw new Error("Pagamentos à conta conjunta têm como destino 'Conta conjunta'.");
    if (!a.admin && norm_(de) !== norm_(a.pessoa) && norm_(para) !== norm_(a.pessoa)) throw new Error("Só pode registar acertos em que participa.");
    var sh = sheetObrig_(ssp, ABA_PACERTOS_), id = proximoIdPart_(sh, "A", 1), linha = Math.max(ultimaLinhaCol_(sh, 1) + 1, 2);
    sh.getRange(linha, 2).setNumberFormat("dd/mm/yyyy"); sh.getRange(linha, 3).setNumberFormat("@"); sh.getRange(linha, 7).setNumberFormat('#,##0.00 "€"');
    sh.getRange(linha, 1, 1, 9).setValues([[id, data, ref, tipo, de, para, valor, String(d.obs || "").trim(), user]]);
    registar_("Acerto registado", id + " · " + de + " → " + para + " · " + valor);
    return id;
  });
}

function excluirAcerto(token, id) {
  var user = exigirSessao_(token, "ACERTOS"), a = acessoDe_(user);
  return comLock_(function () {
    var ssp = getSSP_(), sh = sheetObrig_(ssp, ABA_PACERTOS_), l = acharLinhaPart_(sh, String(id || ""));
    if (!l) throw new Error("Acerto não encontrado. Atualize a página.");
    var x = partAcertos_(ssp).filter(function (r) { return r.linha === l; })[0];
    if (!a.admin && norm_(x.criadoPor) !== norm_(user)) throw new Error("Só pode apagar os acertos que registou.");
    sh.deleteRow(l);
    registar_("Acerto apagado", id);
    return true;
  });
}

/* ---------- ligação com os lançamentos de cartão da app ---------- */

/** Prepara a partilha de uma compra de cartão (antes de gravar), para falhar cedo se algo não estiver pronto. */
function prepararPartilha_(user, modo, deConta) {
  var a = acessoDe_(user);
  if (!a.admin && !a.perms.PARTILHADAS_LANCAR) throw new Error("Não tem permissão para partilhar compras.");
  var ssp = getSSP_(), pe = partPessoas_(ssp), pessoa = partPessoaAtual_(a, pe.tit);
  if (pe.tit.indexOf(pessoa) === -1) throw new Error("A sua pessoa (" + pessoa + ") não está definida nas Despesas partilhadas. Veja a folha ACESSOS.");
  var outro = pe.tit.filter(function (t) { return t !== pessoa; })[0];
  var todos = partCartoes_(ssp);
  var cartao = todos.filter(function (c) { return norm_(c.tipo) === "cartao" && norm_(c.titular) === norm_(pessoa); })[0];
  if (deConta) cartao = todos.filter(function (c) { return norm_(c.nome) === "outro"; })[0] || cartao;      // despesa paga a partir da conta (débito, MB Way…)
  if (!cartao) throw new Error("Não há cartão de " + pessoa + " na folha CONTAS_CARTOES das Despesas partilhadas.");
  return { ssp: ssp, sh: sheetObrig_(ssp, ABA_PART_), pessoa: pessoa, div: modo === "outro" ? "100% " + outro : "Igual", cartao: cartao.nome, user: user, origem: deConta ? "App · conta pessoal" : "App · cartão pessoal" };
}

/** Cria a linha partilhada de uma prestação e devolve o ID. */
function criarLinhaPartilhada_(p, data, refMes, desc, cat, parcela, valor) {
  var id = proximoIdPart_(p.sh, "P", 1), linha = Math.max(ultimaLinhaCol_(p.sh, 1) + 1, 2);
  escreverLinhaPart_(p.sh, linha, { ref: refMes, data: data, tipo: "Compra", desc: desc, cat: cat || "Outros", parcela: parcela, valor: valor, cartao: p.cartao, pagou: p.pessoa, div: p.div }, id, p.user, p.origem || "App · cartão pessoal", "");
  return id;
}

function sincronizarPartilha_(id, v) {
  var ssp = getSSP_(), sh = sheetObrig_(ssp, ABA_PART_), l = acharLinhaPart_(sh, id);
  if (!l) return false;
  var r = sh.getRange(l, 1, 1, 16).getValues()[0];
  sh.getRange(l, 2).setNumberFormat("@");
  if (v.ref) sh.getRange(l, 2).setValue(v.ref);
  if (v.data) { sh.getRange(l, 3).setNumberFormat("dd/mm/yyyy"); sh.getRange(l, 3).setValue(v.data); }
  sh.getRange(l, 5).setValue(v.desc); sh.getRange(l, 6).setValue(v.cat); sh.getRange(l, 8).setValue(v.valor);
  return true;
}

function removerLinhaPartilhada_(id) {
  var ssp = getSSP_(), sh = sheetObrig_(ssp, ABA_PART_), l = acharLinhaPart_(sh, id);
  if (l) sh.deleteRow(l);
  return !!l;
}

function obterEmailUsuario_() {
  var sheet = getSS_().getSheetByName(ABA_.USR);
  if (!sheet || sheet.getLastRow() < 2) return null;
  // aceita vários endereços na mesma célula, separados por vírgula, ponto e vírgula ou espaço
  var lista = String(sheet.getRange(2, 3).getValue() || "").split(/[;,\s]+/).filter(function (e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); });
  return lista.length ? lista.join(",") : null;
}

/* ========================================================================
 * CAMADA DE DADOS (cada folha é lida uma vez por pedido)
 * ====================================================================== */

function criarCtx_() { return { ss: getSS_(), memo: {} }; }

function linhas_(c, nome, nCols) {
  var k = nome + "|" + nCols;
  if (c.memo[k]) return c.memo[k];
  var s = c.ss.getSheetByName(nome);
  var n = s ? Math.min(nCols, s.getMaxColumns()) : nCols;
  var v = (s && s.getLastRow() >= 2) ? s.getRange(2, 1, s.getLastRow() - 1, n).getValues() : [];
  c.memo[k] = v;
  return v;
}

function memo_(c, chave, fn) {
  if (!c.memo[chave]) c.memo[chave] = fn();
  return c.memo[chave];
}

function movimentos_(c, aba) {
  return memo_(c, "mov|" + aba, function () {
    var out = [];
    linhas_(c, aba, 6).forEach(function (r, i) {
      var d = toDate_(r[0]);
      if (!d) return;
      var cod = String(r[1]).trim(), desc = String(r[2] || ""), valor = num_(r[5]);
      out.push({
        linha: i + 2, data: d, cod: cod, desc: desc,
        cat: String(r[3] || "").trim(), conta: String(r[4] || "").trim(), valor: valor,
        interno: cod === "PAG_FAT" || norm_(desc).indexOf("fatura cartao") === 0
      });
    });
    return out;
  });
}

function registosCartao_(c) {
  return memo_(c, "regcart", function () {
    var out = [];
    linhas_(c, ABA_.CART_REG, 8).forEach(function (r, i) {
      var d = toDate_(r[0]);
      out.push({
        linha: i + 2, data: d, cod: String(r[1]).trim(), desc: String(r[2] || ""),
        cat: String(r[3] || "").trim(), cartao: String(r[4] || "").trim(), valor: num_(r[5]),
        fatura: parseFatura_(r[6], d), part: String(r[7] || "").trim()
      });
    });
    return out;
  });
}

function pagamentos_(c) {
  return memo_(c, "pags", function () {
    var out = [];
    linhas_(c, ABA_.FAT, 9).forEach(function (r, i) {
      var d = toDate_(r[0]);
      var ref = parseRef_(r[7]);
      if (!ref && d) ref = { mes: d.getMonth() + 1, ano: d.getFullYear() };
      out.push({
        linha: i + 2, data: d, conta: String(r[4] || "").trim(), cartao: String(r[5] || "").trim(),
        valor: num_(r[6]), refMes: ref ? ref.mes : 0, refAno: ref ? ref.ano : 0,
        parcial: norm_(r[8]) === "parcial"
      });
    });
    return out;
  });
}

function contas_(c) {
  return memo_(c, "contas", function () {
    var out = [];
    linhas_(c, ABA_.CONTAS, 6).forEach(function (r, i) {
      var nome = String(r[1] || "").trim();
      if (!nome) return;
      out.push({
        linha: i + 2, id: String(r[0] || "").trim(), nome: nome, instituicao: String(r[2] || "Banco").trim() || "Banco",
        modalidade: String(r[3] || "Conta corrente").trim() || "Conta corrente", saldoInicial: num_(r[4]),
        arquivada: norm_(r[5]) === "arquivada"
      });
    });
    return out;
  });
}

function cartoes_(c) {
  return memo_(c, "cartoes", function () {
    var out = [];
    linhas_(c, ABA_.CART, 7).forEach(function (r, i) {
      var nome = String(r[1] || "").trim();
      if (!nome) return;
      out.push({
        linha: i + 2, nome: nome, limite: num_(r[0]), bandeira: String(r[2] || "Visa"),
        conta: String(r[3] || ""), venc: parseInt(r[4], 10) || 12, fech: parseInt(r[5], 10) || 30, proc: Math.max(0, parseInt(r[6], 10) || 0)
      });
    });
    return out;
  });
}

/** Categorias das colunas A (receita) e B (despesa) da folha CATEGORIAS; C e D guardam as arquivadas. */
function categorias_(c) {
  return memo_(c, "cats2", function () {
    var rows = linhas_(c, ABA_.CATS, 4), arq = { receita: {}, despesa: {} }, res = { receita: [], despesa: [] }, vistos = { receita: {}, despesa: {} };
    rows.forEach(function (r) {
      var ar = String(r[2] || "").trim(), ad = String(r[3] || "").trim();
      if (ar) arq.receita[norm_(ar)] = true;
      if (ad) arq.despesa[norm_(ad)] = true;
    });
    rows.forEach(function (r) {
      [[0, "receita"], [1, "despesa"]].forEach(function (p) {
        var nome = String(r[p[0]] || "").trim(), k = norm_(nome);
        if (nome && !vistos[p[1]][k]) { vistos[p[1]][k] = true; res[p[1]].push({ nome: nome, arquivada: !!arq[p[1]][k] }); }
      });
    });
    return res;
  });
}

function catsDespesa_(c, comArquivadas) { return categorias_(c).despesa.filter(function (x) { return comArquivadas || !x.arquivada; }).map(function (x) { return x.nome; }); }
function catsReceita_(c, comArquivadas) { return categorias_(c).receita.filter(function (x) { return comArquivadas || !x.arquivada; }).map(function (x) { return x.nome; }); }
function listaCategorias_(c) { return catsDespesa_(c, false); }

/* ========================================================================
 * CÁLCULOS
 * ====================================================================== */

function somaMov_(c, aba, ini, fim) {
  var t = 0;
  movimentos_(c, aba).forEach(function (m) {
    if (!m.interno && m.data >= ini && m.data <= fim) t += m.valor;
  });
  return t;
}

function somaPag_(c, ini, fim, conta) {
  var t = 0;
  pagamentos_(c).forEach(function (p) {
    if (!p.data || p.data < ini || p.data > fim) return;
    if (conta && !mesmoNome_(p.conta, conta)) return;
    t += p.valor;
  });
  return t;
}

/**
 * Resumo do mês por COMPETÊNCIA (o que realmente gastou e recebeu):
 *  - despesa = despesas das contas + compras no cartão, pela DATA DA COMPRA. É o mesmo total do gráfico por categoria e das metas.
 *  - o pagamento da fatura NÃO é despesa (é uma transferência da conta para o cartão): vem à parte em pagFaturas.
 */
function resumoMes_(c, mes, ano) {
  var ini = inicioMes_(mes, ano), fim = fimMes_(mes, ano);
  var despesa = despesasIntervalo_(c, ini, fim).reduce(function (s, x) { return s + x.total; }, 0);
  return { receita: somaMov_(c, ABA_.REC, ini, fim), despesa: despesa, pagFaturas: somaPag_(c, ini, fim) };
}

function saldoAte_(c, fim) {
  var ini = new Date(1900, 0, 1);
  var inicial = 0;
  contas_(c).forEach(function (ct) { inicial += ct.saldoInicial; });
  var ajustes = 0;
  ajustes_(c).forEach(function (a) { if (a.data <= fim) ajustes += a.valor; });
  return inicial + somaMov_(c, ABA_.REC, ini, fim) - somaMov_(c, ABA_.DESP, ini, fim) - somaPag_(c, ini, fim) + ajustes;
}

function saldoConta_(c, ct, fim) {
  var t = ct.saldoInicial;
  movimentos_(c, ABA_.REC).forEach(function (m) { if (!m.interno && m.data <= fim && mesmoNome_(m.conta, ct.nome)) t += m.valor; });
  movimentos_(c, ABA_.DESP).forEach(function (m) { if (!m.interno && m.data <= fim && mesmoNome_(m.conta, ct.nome)) t -= m.valor; });
  pagamentos_(c).forEach(function (p) { if (p.data && p.data <= fim && mesmoNome_(p.conta, ct.nome)) t -= p.valor; });
  transferencias_(c).forEach(function (x) {
    if (x.data > fim) return;
    if (mesmoNome_(x.origem, ct.nome)) t -= x.valor;
    if (mesmoNome_(x.destino, ct.nome)) t += x.valor;
  });
  ajustes_(c).forEach(function (a) { if (a.data <= fim && mesmoNome_(a.conta, ct.nome)) t += a.valor; });
  return t;
}

function totalFatura_(c, cartao, mes, ano) {
  var t = 0;
  registosCartao_(c).forEach(function (r) {
    if (r.fatura && r.fatura.mes === mes && r.fatura.ano === ano && mesmoNome_(r.cartao, cartao)) t += r.valor;
  });
  return t;
}

function totalFaturasMes_(c, mes, ano) {
  var t = 0;
  registosCartao_(c).forEach(function (r) { if (r.fatura && r.fatura.mes === mes && r.fatura.ano === ano) t += r.valor; });
  return t;
}

/**
 * Estado de uma fatura. Pagamentos antigos (sem "TIPO") contam como pagamento total.
 * Pagamentos parciais abatem ao valor em dívida; a fatura só fica paga quando o restante chega a zero.
 */
function estadoFatura_(c, cartao, mes, ano) {
  var total = arred_(totalFatura_(c, cartao, mes, ano)), pago = 0, n = 0, completa = false;
  pagamentos_(c).forEach(function (p) {
    if (!mesmoNome_(p.cartao, cartao) || p.refMes !== mes || p.refAno !== ano) return;
    n++; pago += p.valor;
    if (!p.parcial) completa = true;
  });
  var restante = completa ? 0 : Math.max(0, arred_(total - pago));
  return { total: total, pago: arred_(pago), restante: restante, paga: n > 0 && (completa || restante < 0.005), temPagamentos: n > 0 };
}

function faturaPaga_(c, cartao, mes, ano) { return estadoFatura_(c, cartao, mes, ano).paga; }

function acumular_(mapa, cat, val) {
  var nome = String(cat || "").trim() || "Outros";
  var k = norm_(nome);
  if (!mapa[k]) mapa[k] = { nome: nome, total: 0 };
  mapa[k].total += val;
}

function mapaParaLista_(mapa) {
  return Object.keys(mapa).map(function (k) { return { nome: mapa[k].nome, total: arred_(mapa[k].total) }; })
    .sort(function (a, b) { return b.total - a.total; });
}

/** Gastos do mês por categoria: contas + cartões, sempre pela data em que o gasto aconteceu (nunca pelo mês da fatura). */
function gastosPorCategoria_(c, mes, ano) {
  return despesasIntervalo_(c, inicioMes_(mes, ano), fimMes_(mes, ano));
}

/** Faturas de um cartão: a em vigor e as fechadas por pagar (pelo valor ainda em dívida), independentemente do mês escolhido no ecrã. */
function faturasCartao_(c, ct, hoje) {
  var porMes = {};
  registosCartao_(c).forEach(function (r) {
    if (r.fatura && mesmoNome_(r.cartao, ct.nome)) {
      var k = r.fatura.ano + "-" + r.fatura.mes;
      porMes[k] = porMes[k] || { mes: r.fatura.mes, ano: r.fatura.ano, total: 0 };
      porMes[k].total += r.valor;
    }
  });
  var ab = faturaAbertaRef_(hoje, ct.venc, ct.fech, ct.proc), chaveAb = ab.ano * 12 + ab.mes, aberta = 0, futura = 0;
  var fechadas = [];
  Object.keys(porMes).forEach(function (k) {
    var x = porMes[k], ch = x.ano * 12 + x.mes;
    if (ch === chaveAb) { aberta += x.total; return; }
    if (ch > chaveAb) { futura += x.total; return; }
    var est = estadoFatura_(c, ct.nome, x.mes, x.ano);
    if (est.paga || !(x.total > 0)) return;
    var v = dataSegura_(x.ano, x.mes, ct.venc);
    fechadas.push({
      mes: x.mes, ano: x.ano, rotulo: ABREV_[x.mes - 1] + "/" + x.ano, total: arred_(est.restante), bruto: arred_(x.total), pago: est.pago,
      vencimentoTxt: fmtData_(v), dias: diasEntre_(v, hoje)
    });
  });
  fechadas.sort(function (a, b) { return (a.ano * 12 + a.mes) - (b.ano * 12 + b.mes); });
  var fechadaTotal = fechadas.reduce(function (s, f) { return s + f.total; }, 0);
  return { ab: ab, vigente: { mes: ab.mes, ano: ab.ano, rotulo: ABREV_[ab.mes - 1] + "/" + ab.ano, total: arred_(aberta) }, fechadas: fechadas, fechadaTotal: arred_(fechadaTotal), futuras: arred_(futura) };
}

function historico_(c, mes, ano, qtd) {
  var out = [];
  for (var i = qtd - 1; i >= 0; i--) {
    var d = new Date(ano, mes - 1 - i, 1);
    var m = d.getMonth() + 1, a = d.getFullYear();
    var r = resumoMes_(c, m, a);
    var ab = ABREV_[m - 1];
    out.push({
      mesAnoStr: MESES_[m - 1] + " " + a,
      mesCurto: ab.charAt(0).toUpperCase() + ab.slice(1) + "/" + String(a).slice(2),
      receita: arred_(r.receita), despesa: arred_(r.despesa)
    });
  }
  return out;
}

/* ========================================================================
 * DASHBOARD
 * ====================================================================== */

function getOpcoesFormulario(token) {
  exigirSessao_(token, ["DASHBOARD_VER","CONTAS_VER","MOVIMENTOS_VER","CARTOES_VER","LANCAR","CONTAS_GERIR","DIVIDAS_GERIR","METAS_EDITAR","DEFINICOES","REL_INTERVALO","REL_METAS","REL_MEDIA","REL_ORCADO","DIVIDAS_VER","ANUAL_VER"]);
  var c = criarCtx_();
  return {
    categorias: catsDespesa_(c, false),
    categoriasDespesa: catsDespesa_(c, false),
    categoriasReceita: catsReceita_(c, false),
    contas: contas_(c).filter(function (x) { return !x.arquivada; }).map(function (x) { return x.nome; }),
    cartoes: cartoes_(c).map(function (x) { return x.nome; }),
    cartoesInfo: cartoes_(c).map(function (x) { return { nome: x.nome, venc: x.venc, fech: x.fech, proc: x.proc }; }),
    modalidades: MODALIDADES_
  };
}

function getDadosDashboard(token, mes, ano) {
  exigirSessao_(token, "DASHBOARD_VER");
  mes = +mes; ano = +ano;
  var c = criarCtx_(), hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var atual = resumoMes_(c, mes, ano);
  var ant = mesAnterior_(mes, ano);
  var antR = resumoMes_(c, ant.mes, ant.ano);
  var cats = gastosPorCategoria_(c, mes, ano).filter(function (x) { return x.total > 0.004; });
  var saldo = saldoAte_(c, fimMes_(mes, ano)), fechadas = 0, vigentes = 0;
  cartoes_(c).forEach(function (ct) { var f = faturasCartao_(c, ct, hoje); fechadas += f.fechadaTotal; vigentes += f.vigente.total; });
  var orcTot = orcamentoDe_(c, TOTAL_, mes, ano);
  return {
    kpis: {
      saldo: arred_(saldo), saldoLivre: arred_(saldo - fechadas),
      receita: arred_(atual.receita), despesa: arred_(atual.despesa),
      receitaAnt: arred_(antR.receita), despesaAnt: arred_(antR.despesa),
      faturasFechadas: arred_(fechadas), faturasVigentes: arred_(vigentes)
    },
    criterioCartoes: criterioCartoes_(),
    categorias: { labels: cats.map(function (x) { return x.nome; }), valores: cats.map(function (x) { return x.total; }) },
    balanco: { receitas: arred_(atual.receita), despesas: arred_(atual.despesa), balanco: arred_(atual.receita - atual.despesa), pagFaturas: arred_(atual.pagFaturas) },
    historico: historico_(c, mes, ano, 6),
    dividasMes: dividasDoMes_(c, mes, ano),
    orcamentoTotal: { orcamento: orcTot, gasto: arred_(atual.despesa), progresso: orcTot > 0 ? Math.round(atual.despesa / orcTot * 100) : 0 },
    previsao: previsaoMes_(c, mes, ano, atual.despesa, orcTot)
  };
}

/** Projeção simples até ao fim do mês (só para o mês corrente): ritmo médio diário até hoje × dias que faltam. */
function previsaoMes_(c, mes, ano, gasto, orcamento) {
  var hoje = new Date();
  if (hoje.getMonth() + 1 !== mes || hoje.getFullYear() !== ano) return { ativa: false };
  var dia = hoje.getDate(), dias = new Date(ano, mes, 0).getDate();
  // Com o critério "fatura", as compras de cartão do mês já são conhecidas (fatura fechada ou em formação): ficam fixas e só o resto segue o ritmo diário.
  var fixo = 0;
  if (criterioCartoes_() === "fatura") {
    var ini = inicioMes_(mes, ano), fim = fimMes_(mes, ano);
    registosCartao_(c).forEach(function (r) { var d = dataContagemCartao_(c, r, "fatura"); if (d && d >= ini && d <= fim) fixo += r.valor; });
  }
  var ritmo = (gasto - fixo) / dia, prev = gasto + ritmo * (dias - dia);
  return {
    ativa: true, gastoAteHoje: arred_(gasto), previsaoFim: arred_(prev), ritmoDia: arred_(ritmo), diasRestantes: dias - dia,
    orcamento: orcamento, excede: orcamento > 0 && prev > orcamento, diferenca: orcamento > 0 ? arred_(orcamento - prev) : null
  };
}

/* ========================================================================
 * CONTAS E MOVIMENTOS
 * ====================================================================== */

function getContas(token, mes, ano) {
  exigirSessao_(token, "CONTAS_VER");
  var c = criarCtx_(), fim = fimMes_(+mes, +ano);
  return contas_(c).map(function (ct) {
    return { linha: ct.linha, id: ct.id, nome: ct.nome, instituicao: ct.instituicao, modalidade: ct.modalidade, saldoInicial: ct.saldoInicial, arquivada: ct.arquivada, saldo: arred_(saldoConta_(c, ct, fim)) };
  });
}

function salvarConta(token, dados) {
  exigirSessao_(token, "CONTAS_GERIR");
  var nome = String(dados.nome || "").trim();
  if (!nome) throw new Error("Indique o nome da conta.");
  return comLock_(function () {
    var c = criarCtx_();
    if (contas_(c).some(function (x) { return mesmoNome_(x.nome, nome); })) throw new Error("Já existe uma conta com esse nome.");
    var sheet = sheetObrig_(c.ss, ABA_.CONTAS);
    var linha = Math.max(sheet.getLastRow() + 1, 2);
    anexarLinha_(sheet, ["C" + linha, nome, String(dados.instituicao || "").trim(), dados.modalidade || "Conta corrente", num_(dados.saldoInicial), ""]);
    registar_("Conta criada", nome);
    return true;
  });
}

/** Edita uma conta. Se o nome mudar, atualiza todos os registos ligados. Só arquiva contas com saldo zero. */
function atualizarConta(token, dados) {
  exigirSessao_(token, "CONTAS_GERIR");
  var nome = String(dados.nome || "").trim();
  if (!nome) throw new Error("Indique o nome da conta.");
  return comLock_(function () {
    var c = criarCtx_();
    var atual = contas_(c).filter(function (x) { return x.linha === +dados.linha; })[0];
    if (!atual) throw new Error("Conta não encontrada. Atualize a página.");
    if (contas_(c).some(function (x) { return x.linha !== atual.linha && mesmoNome_(x.nome, nome); })) throw new Error("Já existe outra conta com esse nome.");
    var arquivar = dados.arquivada === true || dados.arquivada === "true";
    if (arquivar && !atual.arquivada && Math.abs(saldoConta_(c, atual, new Date(2999, 0, 1))) >= 0.005) throw new Error("Só pode arquivar contas com saldo zero. Transfira ou ajuste o saldo primeiro.");
    var modal = String(dados.modalidade || atual.modalidade).trim();
    var s = sheetObrig_(c.ss, ABA_.CONTAS);
    s.getRange(atual.linha, 2, 1, 4).setValues([[nome, String(dados.instituicao || "").trim() || atual.instituicao, modal, arred_(num_(dados.saldoInicial))]]);
    s.getRange(atual.linha, 6).setValue(arquivar ? "Arquivada" : "");
    if (!atual.id) s.getRange(atual.linha, 1).setValue("C" + atual.linha);
    if (nome !== atual.nome) {
      var de = atual.nome, ss = c.ss;
      renomearEm_(ss, ABA_.DESP, 5, de, nome); renomearEm_(ss, ABA_.REC, 5, de, nome); renomearEm_(ss, ABA_.FAT, 5, de, nome);
      renomearEm_(ss, ABA_.CART, 4, de, nome); renomearEm_(ss, ABA_.TRF, 2, de, nome); renomearEm_(ss, ABA_.TRF, 3, de, nome);
      renomearEm_(ss, ABA_.AJU, 2, de, nome);
      var sr = ss.getSheetByName(ABA_.RFIX);
      if (sr && sr.getLastRow() >= 2) {
        var n = sr.getLastRow() - 1, rg = sr.getRange(2, 2, n, 4), v = rg.getValues(), ch = false;
        v.forEach(function (r) { if (norm_(r[0]) !== "cartao" && mesmoNome_(r[3], de)) { r[3] = nome; ch = true; } });
        if (ch) rg.setValues(v);
      }
    }
    registar_("Conta editada", atual.nome + (nome !== atual.nome ? " → " + nome : "") + (arquivar ? " (arquivada)" : ""));
    return true;
  });
}

/** Substitui o nome em toda uma coluna de texto de uma folha. Devolve quantas células mudaram. */
function renomearEm_(ss, aba, col, de, para) {
  var s = ss.getSheetByName(aba);
  if (!s || s.getLastRow() < 2) return 0;
  var n = s.getLastRow() - 1, rg = s.getRange(2, col, n, 1), v = rg.getValues(), k = 0;
  v.forEach(function (r) { if (mesmoNome_(r[0], de)) { r[0] = para; k++; } });
  if (k) rg.setValues(v);
  return k;
}

function getTransacoesContas(token, mes, ano) {
  exigirSessao_(token, "CONTAS_VER");
  mes = +mes; ano = +ano;
  var c = criarCtx_(), ini = inicioMes_(mes, ano), fim = fimMes_(mes, ano), lista = [];
  [[ABA_.REC, "receita"], [ABA_.DESP, "despesa"]].forEach(function (par) {
    movimentos_(c, par[0]).forEach(function (m) {
      if (m.interno || m.data < ini || m.data > fim) return;
      lista.push({
        idLinha: m.linha, tipoAba: par[0], tipo: par[1], data: fmtData_(m.data), iso: fmtISO_(m.data),
        descricao: m.desc || "Sem descrição", categoria: m.cat || "Geral", conta: m.conta || "Conta", valor: m.valor,
        reembolso: par[1] === "despesa" && m.valor < 0
      });
    });
  });
  transferencias_(c).forEach(function (x) {
    if (x.data < ini || x.data > fim) return;
    lista.push({
      idLinha: x.linha, tipoAba: ABA_.TRF, tipo: "transferencia", data: fmtData_(x.data), iso: fmtISO_(x.data),
      descricao: x.desc || "Transferência", categoria: "Transferência", conta: x.origem + " → " + x.destino, valor: x.valor
    });
  });
  ajustes_(c).forEach(function (x) {
    if (x.data < ini || x.data > fim) return;
    lista.push({
      idLinha: x.linha, tipoAba: ABA_.AJU, tipo: "ajuste", data: fmtData_(x.data), iso: fmtISO_(x.data),
      descricao: x.desc || "Acerto de saldo", categoria: "Acerto de saldo", conta: x.conta, valor: x.valor
    });
  });
  lista.sort(function (a, b) { return a.iso < b.iso ? 1 : a.iso > b.iso ? -1 : 0; });
  return lista;
}

var ABAS_TRANSACAO_ = [ABA_.DESP, ABA_.REC, ABA_.CART_REG];

/** Fatura (data de vencimento) em que uma compra feita em "data" cai, segundo o ciclo do cartão. */
function faturaDeCompra_(ct, data) {
  var ref = faturaAbertaRef_(data, ct.venc, ct.fech, ct.proc);
  return { ref: ref, data: dataSegura_(ref.ano, ref.mes, ct.venc) };
}

/**
 * Cria um lançamento. Campos: tipo (aba), data, valor, conta (ou cartão), descricao, categoria, codigo,
 * reembolso (true = estorno: guarda valor negativo, que abate à categoria e ao total), fatura ("m/aaaa", só cartão),
 * parcelas (só cartão: divide a compra por N meses, cada parcela na sua data e fatura).
 */
function salvarTransacao(token, dados) {
  var user = exigirSessao_(token, "LANCAR");
  if (ABAS_TRANSACAO_.indexOf(dados.tipo) === -1) throw new Error("Tipo de lançamento inválido.");
  var data = toDate_(dados.data), valor = arred_(Math.abs(num_(dados.valor)));
  var ehCartao = dados.tipo === ABA_.CART_REG, ehReceita = dados.tipo === ABA_.REC;
  var reembolso = (dados.reembolso === true || dados.reembolso === "true") && !ehReceita;
  if (!data) throw new Error("Data inválida.");
  if (!(valor > 0)) throw new Error("Indique um valor maior que zero.");
  if (!String(dados.conta || "").trim()) throw new Error("Escolha a conta ou o cartão.");
  var n = ehCartao && !reembolso ? Math.max(1, Math.min(60, parseInt(dados.parcelas, 10) || 1)) : 1;

  var ehDesp = dados.tipo === ABA_.DESP;
  var modoPart = (ehCartao || ehDesp) && (dados.partilha === "igual" || dados.partilha === "outro") ? dados.partilha : "";
  var res = comLock_(function () {
    var c = criarCtx_(), sheet = sheetObrig_(c.ss, dados.tipo);
    var part = modoPart ? prepararPartilha_(user, modoPart, !ehCartao) : null;      // falha antes de gravar se a partilha não estiver pronta
    var refsPart = [];
    var desc = String(dados.descricao || "Sem descrição"), cat = String(dados.categoria || "Geral"), cod = String(dados.codigo || "").trim();
    var linhas = [], fmt = { 1: "dd/mm/yyyy" };
    if (ehCartao) {
      var ct = cartoes_(c).filter(function (x) { return mesmoNome_(x.nome, dados.conta); })[0];
      if (!ct) throw new Error("Cartão não encontrado.");
      var ref = dados.fatura ? (parseRef_(dados.fatura) || parseFatura_(dados.fatura, data)) : faturaDeCompra_(ct, data).ref;
      if (!ref) throw new Error("Escolha a fatura de referência.");
      var base = arred_(valor / n), acum = 0;
      for (var k = 0; k < n; k++) {
        var vk = k === n - 1 ? arred_(valor - acum) : base; acum += vk;
        var rk = addMesRef_(ref, k);
        linhas.push([k === 0 ? data : addMeses_(data, k), cod, n > 1 ? desc + " (" + (k + 1) + "/" + n + ")" : desc, cat, String(dados.conta), reembolso ? -vk : vk, dataSegura_(rk.ano, rk.mes, ct.venc)]);
        refsPart.push({ ref: rk.ano + "-" + (rk.mes < 10 ? "0" : "") + rk.mes, valor: reembolso ? -vk : vk });
      }
      fmt[7] = "dmmm";
    } else {
      linhas.push([data, cod, desc, cat, String(dados.conta), reembolso ? -valor : valor]);
    }
    if (part && !ehCartao && sheet.getMaxColumns() < 8) sheet.insertColumnsAfter(sheet.getMaxColumns(), 8 - sheet.getMaxColumns());
    var iniLinha = anexarLinhas_(sheet, linhas, fmt);
    if (part && !ehCartao) {                                               // despesa de conta: conta no mês da data
      var mm = linhas[0][0], rf0 = mm.getFullYear() + "-" + (mm.getMonth() < 9 ? "0" : "") + (mm.getMonth() + 1);
      sheet.getRange(iniLinha, 8).setValue(criarLinhaPartilhada_(part, mm, rf0, desc, cat, "", reembolso ? -valor : valor));
    } else if (part) {                                                            // compra partilhada: cria a linha na planilha das despesas partilhadas e liga as duas
      linhas.forEach(function (ln, k) {
        var id = criarLinhaPartilhada_(part, ln[0], refsPart[k].ref, desc, cat, n > 1 ? (k + 1) + "/" + n : "", refsPart[k].valor);
        sheet.getRange(iniLinha + k, 8).setValue(id);
      });
    }
    registar_(reembolso ? "Reembolso" : "Lançamento", dados.tipo + " · " + desc + " · " + valor + (n > 1 ? " em " + n + "x" : "") + (part ? " · partilhada (" + part.div + ")" : ""));
    return linhas.length;
  });

  if (dados.tipo !== ABA_.REC) {
    try { verificarAlertasMetas_(criarCtx_(), data.getMonth() + 1, data.getFullYear()); } catch (e) { console.error(e); }
  }
  return res;
}

function confereLinha_(sheet, idLinha, descEsperada) {
  if (!sheet || idLinha < 2 || idLinha > sheet.getLastRow()) throw new Error("Registo não encontrado. Atualize a página.");
  var atual = String(sheet.getRange(idLinha, 3).getValue() || "").trim();
  if (descEsperada != null && atual !== String(descEsperada).trim()) throw new Error("Os dados mudaram entretanto. Atualize a página e tente de novo.");
}

function atualizarTransacao(token, dados) {
  exigirSessao_(token, "LANCAR");
  if (ABAS_TRANSACAO_.indexOf(dados.tipoAba) === -1) throw new Error("Tipo inválido.");
  var data = toDate_(dados.data), valor = arred_(Math.abs(num_(dados.valor)));
  if (!data) throw new Error("Data inválida.");
  if (!(valor > 0)) throw new Error("Indique um valor maior que zero.");
  var reembolso = (dados.reembolso === true || dados.reembolso === "true") && dados.tipoAba !== ABA_.REC;
  return comLock_(function () {
    var sheet = getSS_().getSheetByName(dados.tipoAba);
    confereLinha_(sheet, +dados.idLinha, dados.descOriginal);
    var l = +dados.idLinha;
    sheet.getRange(l, 1).setNumberFormat("dd/mm/yyyy").setValue(data);
    sheet.getRange(l, 3, 1, 2).setValues([[String(dados.descricao || "Sem descrição"), String(dados.categoria || "Geral")]]);
    if (dados.tipoAba !== ABA_.CART_REG) sheet.getRange(l, 5).setValue(String(dados.conta || ""));
    sheet.getRange(l, 6).setValue(reembolso ? -valor : valor);
    if (dados.tipoAba === ABA_.CART_REG && dados.fatura) {
      var ref = parseRef_(dados.fatura), cc = criarCtx_();
      var ct = cartoes_(cc).filter(function (x) { return mesmoNome_(x.nome, sheet.getRange(l, 5).getValue()); })[0];
      if (!ref || !ct) throw new Error("Fatura ou cartão inválido.");
      sheet.getRange(l, 7).setNumberFormat("dmmm").setValue(dataSegura_(ref.ano, ref.mes, ct.venc));
    }
    if (dados.tipoAba === ABA_.CART_REG || (dados.tipoAba === ABA_.DESP && sheet.getMaxColumns() >= 8)) {   // compra partilhada: mantém a linha partilhada igual
      var pid = String(sheet.getRange(l, 8).getValue() || "").trim();
      if (pid) {
        var rf = dados.tipoAba === ABA_.DESP ? { ano: data.getFullYear(), mes: data.getMonth() + 1 } : parseFatura_(sheet.getRange(l, 7).getValue(), data);
        try { sincronizarPartilha_(pid, { ref: rf ? rf.ano + "-" + (rf.mes < 10 ? "0" : "") + rf.mes : "", data: data, desc: String(dados.descricao || "Sem descrição"), cat: String(dados.categoria || "Geral"), valor: reembolso ? -valor : valor }); }
        catch (e) { registar_("⚠ Partilha não sincronizada", pid + " · " + e.message); }
      }
    }
    registar_("Lançamento editado", dados.tipoAba + " linha " + l + " · " + String(dados.descricao || "") + " · " + (reembolso ? -valor : valor));
    return true;
  });
}

function excluirTransacao(token, tipoAba, idLinha, descEsperada) {
  exigirSessao_(token, "LANCAR");
  if (ABAS_TRANSACAO_.concat([ABA_.TRF, ABA_.AJU]).indexOf(tipoAba) === -1) throw new Error("Tipo inválido.");
  return comLock_(function () {
    var sheet = getSS_().getSheetByName(tipoAba);
    var colDesc = tipoAba === ABA_.TRF ? 5 : tipoAba === ABA_.AJU ? 4 : 3;
    if (!sheet || +idLinha < 2 || +idLinha > sheet.getLastRow()) throw new Error("Registo não encontrado. Atualize a página.");
    var atual = String(sheet.getRange(+idLinha, colDesc).getValue() || "").trim();
    var esperado = String(descEsperada == null ? "" : descEsperada).trim();
    if (descEsperada != null && atual !== esperado && atual !== "") throw new Error("Os dados mudaram entretanto. Atualize a página e tente de novo.");
    if (tipoAba === ABA_.CART_REG || (tipoAba === ABA_.DESP && sheet.getMaxColumns() >= 8)) {
      var pid = String(sheet.getRange(+idLinha, 8).getValue() || "").trim();
      if (pid) { try { removerLinhaPartilhada_(pid); } catch (e) { registar_("⚠ Partilha não removida", pid + " · " + e.message); } }
    }
    sheet.deleteRow(+idLinha);
    registar_("Lançamento apagado", tipoAba + " · " + atual);
    return true;
  });
}

/* ========================================================================
 * CARTÕES
 * ====================================================================== */

function getCartoes(token, mes, ano) {
  exigirSessao_(token, "CARTOES_VER");
  mes = +mes; ano = +ano;
  var c = criarCtx_(), hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return cartoes_(c).map(function (ct) {
    var f = faturasCartao_(c, ct, hoje);
    var est = estadoFatura_(c, ct.nome, mes, ano);              // fatura do período escolhido (para o detalhe e histórico)
    // Limite consumido = faturas fechadas por pagar (valor ainda em dívida) + fatura em vigor. Ao pagar, o limite é libertado.
    var usado = f.fechadaTotal + f.vigente.total;
    var vencData = new Date(ano, mes - 1, ct.venc);
    return {
      linha: ct.linha, descricao: ct.nome, limite: ct.limite, bandeira: ct.bandeira, contaVinculada: ct.conta,
      vigente: f.vigente, fechadas: f.fechadas, futuras: f.futuras,
      fatura: est.total, paga: est.paga, faturaPago: est.pago, faturaRestante: est.paga ? 0 : (est.temPagamentos ? est.restante : est.total),
      faturaFechada: f.fechadaTotal, faturaAberta: f.vigente.total, abertaRotulo: f.vigente.rotulo,
      emAberto: arred_(usado), disponivel: arred_(ct.limite - usado),
      vencimento: ct.venc, fechamento: ct.fech, processamento: ct.proc, diasParaVencimento: est.paga ? null : Math.round((vencData - hoje) / 86400000)
    };
  });
}

/** Fatura (mês/ano de vencimento) onde as compras de HOJE vão cair. */
function faturaAbertaRef_(hoje, venc, fech, proc) {
  // "proc" = dias que o banco leva a processar: compras dos últimos dias antes do fecho já entram na fatura seguinte
  var mc = hoje.getDate() <= Math.max(0, fech - (proc || 0)) ? hoje.getMonth() + 1 : hoje.getMonth() + 2; // mês em que a fatura fecha
  if (venc <= fech) mc += 1;                                                   // vence no mês seguinte ao fecho
  var d = new Date(hoje.getFullYear(), mc - 1, 1);
  return { mes: d.getMonth() + 1, ano: d.getFullYear() };
}

/** Todas as faturas de um cartão (as que têm compras, mais a em vigor), da mais antiga para a mais recente. */
function getFaturasCartao(token, nomeCartao) {
  exigirSessao_(token, "CARTOES_VER");
  var c = criarCtx_(), hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var ct = cartoes_(c).filter(function (x) { return mesmoNome_(x.nome, nomeCartao); })[0];
  if (!ct) throw new Error("Cartão não encontrado.");
  var ab = faturaAbertaRef_(hoje, ct.venc, ct.fech, ct.proc), chaveAb = ab.ano * 12 + ab.mes, mapa = {};
  registosCartao_(c).forEach(function (r) {
    if (!r.fatura || !mesmoNome_(r.cartao, ct.nome)) return;
    var k = r.fatura.ano * 12 + r.fatura.mes;
    mapa[k] = mapa[k] || { mes: r.fatura.mes, ano: r.fatura.ano, total: 0, n: 0 };
    mapa[k].total += r.valor; mapa[k].n++;
  });
  if (!mapa[chaveAb]) mapa[chaveAb] = { mes: ab.mes, ano: ab.ano, total: 0, n: 0 };
  return Object.keys(mapa).map(Number).sort(function (a, b) { return a - b; }).map(function (k) {
    var x = mapa[k], est = estadoFatura_(c, ct.nome, x.mes, x.ano);
    var estado = k === chaveAb ? "vigente" : est.paga ? "paga" : k > chaveAb ? "futura" : x.total > 0 ? "porPagar" : "vazia";
    return { mes: x.mes, ano: x.ano, rotulo: MESES_[x.mes - 1] + "/" + x.ano, abrev: ABREV_[x.mes - 1] + "/" + x.ano, total: arred_(x.total), pago: est.pago, n: x.n, estado: estado,
      vencimentoTxt: fmtData_(dataSegura_(x.ano, x.mes, ct.venc)) };
  });
}

function getTransacoesCartao(token, nomeCartao, mes, ano) {
  exigirSessao_(token, "CARTOES_VER");
  mes = +mes; ano = +ano;
  var c = criarCtx_(), lista = [];
  registosCartao_(c).forEach(function (r) {
    if (!mesmoNome_(r.cartao, nomeCartao) || !r.fatura || r.fatura.mes !== mes || r.fatura.ano !== ano) return;
    lista.push({
      idLinha: r.linha, tipoAba: ABA_.CART_REG, data: fmtData_(r.data), iso: fmtISO_(r.data),
      descricao: r.desc || "Sem descrição", categoria: r.cat || "Geral", valor: r.valor, reembolso: r.valor < 0, faturaRef: r.fatura.mes + "/" + r.fatura.ano, partilha: r.part
    });
  });
  lista.sort(function (a, b) { return a.iso < b.iso ? 1 : a.iso > b.iso ? -1 : 0; });
  return lista;
}

function salvarCartao(token, dados) {
  exigirSessao_(token, "CONTAS_GERIR");
  var nome = String(dados.descricao || "").trim();
  if (!nome) throw new Error("Indique o nome do cartão.");
  var venc = parseInt(dados.vencimento, 10) || 12, fech = parseInt(dados.fechamento, 10) || 30;
  if (venc < 1 || venc > 31 || fech < 1 || fech > 31) throw new Error("Dias de vencimento/fecho devem estar entre 1 e 31.");
  return comLock_(function () {
    var c = criarCtx_();
    if (cartoes_(c).some(function (x) { return mesmoNome_(x.nome, nome); })) throw new Error("Já existe um cartão com esse nome.");
    anexarLinha_(sheetObrig_(c.ss, ABA_.CART), [num_(dados.limite), nome, dados.bandeira || "Visa", String(dados.contaVinculada || ""), venc, fech, Math.max(0, parseInt(dados.processamento, 10) || 0)]);
    registar_("Cartão criado", nome);
    return true;
  });
}

/** Edita limite, bandeira, conta, vencimento e fecho. Se o nome mudar, atualiza também os registos ligados. */
function atualizarCartao(token, dados) {
  exigirSessao_(token, "CONTAS_GERIR");
  var nome = String(dados.descricao || "").trim();
  if (!nome) throw new Error("Indique o nome do cartão.");
  var venc = parseInt(dados.vencimento, 10), fech = parseInt(dados.fechamento, 10);
  if (!(venc >= 1 && venc <= 31) || !(fech >= 1 && fech <= 31)) throw new Error("Dias de vencimento/fecho devem estar entre 1 e 31.");
  var limite = arred_(num_(dados.limite));
  if (limite < 0) throw new Error("O limite não pode ser negativo.");

  return comLock_(function () {
    var c = criarCtx_();
    var atual = cartoes_(c).filter(function (x) { return x.linha === +dados.linha; })[0];
    if (!atual) throw new Error("Cartão não encontrado. Atualize a página.");
    if (cartoes_(c).some(function (x) { return x.linha !== atual.linha && mesmoNome_(x.nome, nome); })) throw new Error("Já existe outro cartão com esse nome.");

    var sCart = sheetObrig_(c.ss, ABA_.CART);
    sCart.getRange(atual.linha, 1, 1, 7).setValues([[limite, nome, dados.bandeira || "Visa", String(dados.contaVinculada || ""), venc, fech, Math.max(0, parseInt(dados.processamento, 10) || 0)]]);

    if (nome !== atual.nome) {
      renomearEm_(c.ss, ABA_.CART_REG, 5, atual.nome, nome);
      renomearEm_(c.ss, ABA_.FAT, 6, atual.nome, nome);
      var ant = norm_("Fatura Cartão " + atual.nome), novo = "Fatura Cartão " + nome;
      [[ABA_.FAT, 3], [ABA_.DESP, 3]].forEach(function (par) {
        var s = c.ss.getSheetByName(par[0]);
        if (!s || s.getLastRow() < 2) return;
        var n = s.getLastRow() - 1, rg = s.getRange(2, par[1], n, 1), v = rg.getValues(), ch = false;
        v.forEach(function (r) { if (norm_(r[0]) === ant) { r[0] = novo; ch = true; } });
        if (ch) rg.setValues(v);
      });
      var sr = c.ss.getSheetByName(ABA_.RFIX);
      if (sr && sr.getLastRow() >= 2) {
        var nr = sr.getLastRow() - 1, rr = sr.getRange(2, 2, nr, 4), vr = rr.getValues(), cr = false;
        vr.forEach(function (r) { if (norm_(r[0]) === "cartao" && mesmoNome_(r[3], atual.nome)) { r[3] = nome; cr = true; } });
        if (cr) rr.setValues(vr);
      }
    }
    registar_("Cartão editado", atual.nome + (nome !== atual.nome ? " → " + nome : "") + " · limite " + limite);
    return true;
  });
}

/**
 * Paga uma fatura (toda ou em parte). "valor" vazio = paga o que falta. "encargos" = juros/comissões cobrados por não pagar tudo:
 * entram como despesa do cartão (categoria "Juros e encargos") na fatura em vigor.
 */
function pagarFaturaCartao(token, dados) {
  exigirSessao_(token, "CONTAS_GERIR");
  var mes = +dados.mes, ano = +dados.ano, data = toDate_(dados.data);
  if (!data) throw new Error("Data de pagamento inválida.");
  if (!String(dados.conta || "").trim()) throw new Error("Escolha a conta de débito.");
  var encargos = arred_(Math.max(0, num_(dados.encargos)));
  return comLock_(function () {
    var c = criarCtx_();
    var ct = cartoes_(c).filter(function (x) { return mesmoNome_(x.nome, dados.cartao); })[0];
    if (!ct) throw new Error("Cartão não encontrado.");
    var est = estadoFatura_(c, ct.nome, mes, ano);
    if (est.paga) throw new Error("Esta fatura já está paga.");
    var falta = est.temPagamentos ? est.restante : est.total;
    if (!(falta > 0)) throw new Error("A fatura está a zero, não há nada a pagar.");
    var valor = String(dados.valor == null ? "" : dados.valor).trim() === "" ? falta : arred_(num_(dados.valor));
    if (!(valor > 0)) throw new Error("Indique um valor maior que zero.");
    if (valor > falta + 0.005) throw new Error("O valor é superior ao que falta pagar (" + eurTxt_(falta) + ").");
    var parcial = valor < falta - 0.005;
    var desc = "Fatura Cartão " + ct.nome;
    anexarLinha_(sheetObrig_(c.ss, ABA_.FAT), [data, "PAG_FAT", desc, "Fatura Cartão", dados.conta, ct.nome, valor, mes + "/" + ano, parcial ? "parcial" : "total"], { 1: "dd/mm/yyyy", 8: "@" });
    var sDesp = c.ss.getSheetByName(ABA_.DESP);
    if (sDesp) anexarLinha_(sDesp, [data, "PAG_FAT", desc, "Fatura Cartão", dados.conta, valor], { 1: "dd/mm/yyyy" });
    if (encargos > 0) {
      var fb = faturaDeCompra_(ct, data);
      anexarLinha_(sheetObrig_(c.ss, ABA_.CART_REG), [data, "JUROS", "Juros/encargos fatura " + ABREV_[mes - 1] + "/" + ano, "Juros e encargos", ct.nome, encargos, fb.data], { 1: "dd/mm/yyyy", 7: "dmmm" });
    }
    registar_(parcial ? "Fatura paga (parcial)" : "Fatura paga", ct.nome + " " + mes + "/" + ano + " · " + valor + (encargos ? " + encargos " + encargos : ""));
    return { parcial: parcial, restante: arred_(Math.max(0, falta - valor)) };
  });
}

function reabrirFaturaCartao(token, nomeCartao, mes, ano) {
  exigirSessao_(token, "CONTAS_GERIR");
  mes = +mes; ano = +ano;
  return comLock_(function () {
    var c = criarCtx_();
    var remover = pagamentos_(c).filter(function (p) { return mesmoNome_(p.cartao, nomeCartao) && p.refMes === mes && p.refAno === ano; });
    if (!remover.length) return false;

    var chaves = {}, datasPag = {};
    remover.forEach(function (p) { var k = fmtISO_(p.data) + "|" + p.valor.toFixed(2); chaves[k] = (chaves[k] || 0) + 1; datasPag[fmtISO_(p.data)] = true; });
    var descAlvo = norm_("Fatura Cartão " + nomeCartao);
    var linhasDesp = [];
    movimentos_(c, ABA_.DESP).forEach(function (m) {
      if (m.cod !== "PAG_FAT" || norm_(m.desc) !== descAlvo) return;
      var k = fmtISO_(m.data) + "|" + m.valor.toFixed(2);
      if (chaves[k] > 0) { chaves[k]--; linhasDesp.push(m.linha); }
    });
    // juros/encargos lançados ao pagar parcialmente
    var linhasJuros = [], descJ = "Juros/encargos fatura " + ABREV_[mes - 1] + "/" + ano;
    registosCartao_(c).forEach(function (r) {
      if (r.cod === "JUROS" && r.desc === descJ && mesmoNome_(r.cartao, nomeCartao) && r.data && datasPag[fmtISO_(r.data)]) linhasJuros.push(r.linha);
    });

    var sFat = c.ss.getSheetByName(ABA_.FAT), sDesp = c.ss.getSheetByName(ABA_.DESP), sReg = c.ss.getSheetByName(ABA_.CART_REG);
    remover.map(function (p) { return p.linha; }).sort(function (a, b) { return b - a; }).forEach(function (l) { sFat.deleteRow(l); });
    if (sDesp) linhasDesp.sort(function (a, b) { return b - a; }).forEach(function (l) { sDesp.deleteRow(l); });
    if (sReg) linhasJuros.sort(function (a, b) { return b - a; }).forEach(function (l) { sReg.deleteRow(l); });
    registar_("Fatura reaberta", nomeCartao + " " + mes + "/" + ano);
    return true;
  });
}

/* ========================================================================
 * METAS / ORÇAMENTO
 * ====================================================================== */

/** Orçamentos: folha ORÇAMENTOS (criada pela migração). Antes disso, lê as linhas com mês/ano da folha METAS. */
function folhaOrcNome_(c) { return c.ss.getSheetByName(ABA_.ORC) ? ABA_.ORC : ABA_.METAS; }

function metasLinhas_(c) {
  return memo_(c, "metas", function () {
    var out = [], nome = folhaOrcNome_(c);
    linhas_(c, nome, 6).forEach(function (r, i) {
      var cat = String(r[0] || "").trim(), ref = parseRef_(r[5]);
      if (!cat) return;
      if (nome === ABA_.METAS && !ref) return;  // as linhas do FILTER de METAS não são orçamentos
      out.push({ linha: i + 2, cat: cat, orc: num_(r[1]), ref: ref });
    });
    return out;
  });
}

/** Garante a folha ORÇAMENTOS (copia para lá as linhas já datadas de METAS, que tem uma fórmula FILTER que não deve ser misturada). */
function garantirOrc_(c) {
  var s = c.ss.getSheetByName(ABA_.ORC);
  if (s) return s;
  var antigas = metasLinhas_(c).map(function (m) { return m; });
  var rawMetas = linhas_(c, ABA_.METAS, 6);
  s = abaOuCria_(c.ss, ABA_.ORC, ["Categoria", "Orçamento Mensal (€)", "Alerta 50% Enviado", "Alerta 80% Enviado", "Alerta 100% Enviado", "Mês/Ano Ref."]);
  var linhas = [];
  antigas.forEach(function (m) {
    var r = rawMetas[m.linha - 2] || [];
    if (m.ref) linhas.push([m.cat, m.orc, r[2] === true, r[3] === true, r[4] === true, m.ref.mes + "/" + m.ref.ano]);
  });
  if (linhas.length) anexarLinhas_(s, linhas, { 6: "@" });
  delete c.memo.metas; delete c.memo[ABA_.ORC + "|6"];
  return s;
}

/** Usa o orçamento do mês; se não existir, herda o mais recente de meses anteriores. */
function orcamentoDe_(c, cat, mes, ano) {
  var alvo = ano * 12 + mes, melhor = null, melhorChave = -1;
  metasLinhas_(c).forEach(function (m) {
    if (norm_(m.cat) !== norm_(cat)) return;
    var chave = m.ref ? m.ref.ano * 12 + m.ref.mes : 0;
    if (chave <= alvo && chave >= melhorChave) { melhor = m; melhorChave = chave; }
  });
  return melhor ? melhor.orc : 0;
}

function metasDoMes_(c, mes, ano) {
  var gastos = {};
  gastosPorCategoria_(c, mes, ano).forEach(function (g) { gastos[norm_(g.nome)] = g.total; });
  var hoje = new Date(), ehAtual = hoje.getMonth() + 1 === mes && hoje.getFullYear() === ano;
  var dia = hoje.getDate(), dias = new Date(ano, mes, 0).getDate();
  var nomes = listaCategorias_(c).slice(), conhecidos = {};
  nomes.forEach(function (n) { conhecidos[norm_(n)] = true; });
  // categorias com gasto (ou com orçamento) que já não estão na lista ativa também aparecem
  Object.keys(gastos).forEach(function (k) { if (!conhecidos[k] && gastos[k] > 0.004) { conhecidos[k] = true; nomes.push(k); } });
  var nomesPorChave = {};
  gastosPorCategoria_(c, mes, ano).forEach(function (g) { nomesPorChave[norm_(g.nome)] = g.nome; });
  return nomes.map(function (cat) {
    var nome = nomesPorChave[norm_(cat)] || cat;
    var orc = orcamentoDe_(c, nome, mes, ano), gasto = gastos[norm_(nome)] || 0;
    var prev = ehAtual && gasto > 0 ? gasto + gasto / dia * (dias - dia) : gasto;
    return { categoria: nome, orcamento: orc, gasto: arred_(gasto), progresso: orc > 0 ? Math.max(0, Math.round(gasto / orc * 100)) : 0, previsao: arred_(prev) };
  });
}

function getMetasGastos(token, mes, ano) {
  exigirSessao_(token, ["REL_METAS","DASHBOARD_VER","REL_ORCADO"]);
  return metasDoMes_(criarCtx_(), +mes, +ano);
}

function getResumoOrcamento(token, mes, ano) {
  exigirSessao_(token, ["REL_METAS","DASHBOARD_VER","REL_ORCADO"]);
  mes = +mes; ano = +ano;
  var c = criarCtx_(), gasto = resumoMes_(c, mes, ano).despesa, orc = orcamentoDe_(c, TOTAL_, mes, ano);
  var soma = 0;
  metasDoMes_(c, mes, ano).forEach(function (m) { soma += m.orcamento; });
  return {
    categoriaTotal: TOTAL_, orcamento: orc, gasto: arred_(gasto), progresso: orc > 0 ? Math.max(0, Math.round(gasto / orc * 100)) : 0,
    somaCategorias: arred_(soma), previsao: previsaoMes_(c, mes, ano, gasto, orc)
  };
}

function salvarMetaApp(token, categoria, orcamento, mes, ano) {
  exigirSessao_(token, "METAS_EDITAR");
  mes = +mes; ano = +ano;
  var valor = arred_(num_(orcamento));
  if (valor < 0) throw new Error("O orçamento não pode ser negativo.");
  return comLock_(function () {
    var c = criarCtx_(), sheet = garantirOrc_(c);
    var existente = metasLinhas_(c).filter(function (m) { return norm_(m.cat) === norm_(categoria) && m.ref && m.ref.mes === mes && m.ref.ano === ano; })[0];
    if (existente) { sheet.getRange(existente.linha, 2).setValue(valor); }
    else anexarLinha_(sheet, [categoria, valor, false, false, false, mes + "/" + ano], { 6: "@" });
    registar_("Orçamento", categoria + " " + mes + "/" + ano + " = " + valor);
    return true;
  });
}

/** Grava explicitamente, no mês indicado, os orçamentos que estão a valer (herdados do mês anterior) e ainda não têm linha própria. */
function copiarOrcamentoAnterior(token, mes, ano) {
  exigirSessao_(token, "METAS_EDITAR");
  mes = +mes; ano = +ano;
  return comLock_(function () {
    var c = criarCtx_(), sheet = garantirOrc_(c), ant = mesAnterior_(mes, ano), novas = [], vistos = {};
    var tem = {};
    metasLinhas_(c).forEach(function (m) { if (m.ref && m.ref.mes === mes && m.ref.ano === ano) tem[norm_(m.cat)] = true; });
    var cats = listaCategorias_(c).concat([TOTAL_]);
    cats.forEach(function (cat) {
      var k = norm_(cat);
      if (tem[k] || vistos[k]) return;
      vistos[k] = true;
      var orc = orcamentoDe_(c, cat, ant.mes, ant.ano);
      if (orc > 0) novas.push([cat, orc, false, false, false, mes + "/" + ano]);
    });
    if (novas.length) anexarLinhas_(sheet, novas, { 6: "@" });
    registar_("Orçamento copiado", ant.mes + "/" + ant.ano + " → " + mes + "/" + ano + " (" + novas.length + ")");
    return novas.length;
  });
}

/* ========================================================================
 * NOTIFICAÇÕES POR E-MAIL
 * Destinatário: folha USUÁRIO, linha 2, coluna C.
 * Os templates usam tabelas e estilos inline (compatível com Gmail, Outlook e telemóvel).
 * ====================================================================== */
var FONTE_ = "-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";
var LOGO_URL_ = "https://lh3.googleusercontent.com/d/1OIQAMN2cUuVMErMSU5N6uXTV5DXLkMiN";
var TEMAS_ = {
  marca:  { cor: "#7b1fa2", suave: "#ecd5f7", fundo: "#f7edfc" },
  info:   { cor: "#0b86c4", suave: "#d6f0fc", fundo: "#eaf6fc" },
  aviso:  { cor: "#d97706", suave: "#fdecc8", fundo: "#fef6e4" },
  perigo: { cor: "#d92d3f", suave: "#fdd9dd", fundo: "#fdeeef" }
};

function escHtml_(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function eurTxt_(v) {
  v = Number(v) || 0;
  var p = Math.abs(v).toFixed(2).split(".");
  p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return (v < 0 ? "-" : "") + p[0] + "," + p[1] + " €";
}

function linkPainel_() {
  try { return ScriptApp.getService().getUrl() || ""; } catch (e) { return ""; }
}

/** Molde comum a todos os e-mails. o: {preheader, tema, emoji, kicker, titulo, subtitulo, corpo, url} */
function emailBase_(o) {
  var t = o.tema || TEMAS_.marca;
  var url = o.url !== undefined ? o.url : linkPainel_();
  var botao = url
    ? '<tr><td style="padding:2px 28px 30px;"><table role="presentation" cellpadding="0" cellspacing="0"><tr><td style="background:' + TEMAS_.marca.cor + ';border-radius:10px;">' +
      '<a href="' + escHtml_(url) + '" style="display:inline-block;padding:13px 26px;font-family:' + FONTE_ + ';font-size:14px;font-weight:700;color:#ffffff;text-decoration:none;">Abrir o painel &rarr;</a></td></tr></table></td></tr>'
    : "";
  return '<!DOCTYPE html><html lang="pt-PT"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light"><title>' + escHtml_(o.titulo) + '</title></head>' +
    '<body style="margin:0;padding:0;background:#f1f2f8;">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:#f1f2f8;font-size:1px;line-height:1px;">' + escHtml_(o.preheader || "") + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f1f2f8;"><tr><td align="center" style="padding:26px 12px;">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;">' +
    '<tr><td style="padding:0 6px 16px;"><img src="' + LOGO_URL_ + '" alt="Performance Financial" width="150" style="display:block;border:0;height:auto;max-width:150px;font-family:' + FONTE_ + ';font-size:16px;font-weight:800;color:#15171e;"></td></tr>' +
    '<tr><td style="background:#ffffff;border:1px solid #e4e7f0;border-radius:16px;overflow:hidden;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
    '<tr><td style="background:' + t.cor + ';padding:26px 28px;border-radius:16px 16px 0 0;">' +
      '<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>' +
      '<td width="52" valign="middle"><table role="presentation" cellpadding="0" cellspacing="0"><tr><td width="44" height="44" align="center" style="background:#ffffff;border-radius:22px;font-size:22px;line-height:44px;">' + (o.emoji || "") + '</td></tr></table></td>' +
      '<td valign="middle" style="font-family:' + FONTE_ + ';">' +
        '<div style="font-size:11px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:' + t.suave + ';">' + escHtml_(o.kicker || "") + '</div>' +
        '<div style="font-size:22px;line-height:28px;font-weight:800;color:#ffffff;margin-top:2px;">' + escHtml_(o.titulo) + '</div>' +
      '</td></tr></table>' +
      (o.subtitulo ? '<div style="font-family:' + FONTE_ + ';font-size:14px;line-height:20px;color:' + t.suave + ';margin-top:14px;">' + o.subtitulo + '</div>' : "") +
    '</td></tr>' +
    '<tr><td style="padding:26px 28px 22px;font-family:' + FONTE_ + ';color:#15171e;">' + o.corpo + '</td></tr>' +
    botao +
    '</table></td></tr>' +
    '<tr><td align="center" style="padding:18px 10px 6px;font-family:' + FONTE_ + ';font-size:12px;line-height:18px;color:#8a90a3;">' +
    'Recebe este e-mail porque tem alertas ativos no Performance Financial.<br>Mensagem automática, não é necessário responder.</td></tr>' +
    '</table></td></tr></table></body></html>';
}

/* ---------- Alerta de meta ---------- */
/** d: {categoria, orcamento, gasto, progresso, nivel(50|80|100), mesRotulo, diasRestantes|null, url?} */
function emailMeta_(d) {
  var n = d.nivel;
  var t = n >= 100 ? TEMAS_.perigo : n >= 80 ? TEMAS_.aviso : TEMAS_.info;
  var titulo = n >= 100 ? "Meta ultrapassada" : n >= 80 ? "Está quase no limite" : "Metade da meta atingida";
  var emoji = n >= 100 ? "🚨" : n >= 80 ? "⚠️" : "💡";
  var restante = d.orcamento - d.gasto, pct = Math.max(1, Math.min(100, Math.round(d.progresso)));

  var barra = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;"><tr>' +
    '<td width="' + pct + '%" height="12" style="background:' + t.cor + ';border-radius:6px;font-size:0;line-height:0;">&nbsp;</td>' +
    (pct < 100 ? '<td width="' + (100 - pct) + '%" height="12" style="background:#eceef6;border-radius:6px;font-size:0;line-height:0;">&nbsp;</td>' : "") +
    '</tr></table>';

  var celula = function (rotulo, valor, cor, borda) {
    return '<td width="33%" valign="top" style="padding:14px 12px;' + (borda ? "border-left:1px solid #e4e7f0;" : "") + '">' +
      '<div style="font-size:11px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:#6a7183;">' + rotulo + '</div>' +
      '<div style="font-size:17px;font-weight:800;color:' + (cor || "#15171e") + ';margin-top:4px;white-space:nowrap;">' + valor + '</div></td>';
  };
  var stats = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e4e7f0;border-radius:12px;margin-top:22px;border-collapse:separate;"><tr>' +
    celula("Orçamento", eurTxt_(d.orcamento), null, false) + celula("Gasto", eurTxt_(d.gasto), t.cor, true) +
    celula(restante >= 0 ? "Disponível" : "Excedido", eurTxt_(Math.abs(restante)), restante >= 0 ? "#12a150" : "#d92d3f", true) + '</tr></table>';

  var dica;
  if (n >= 100) dica = "Ultrapassou o orçamento em <b>" + eurTxt_(Math.abs(restante)) + "</b>. Reveja os próximos gastos desta categoria ou ajuste a meta do mês.";
  else if (d.diasRestantes && d.diasRestantes > 0) dica = "Faltam <b>" + d.diasRestantes + (d.diasRestantes === 1 ? " dia" : " dias") + "</b> para o fim do mês. Para ficar dentro da meta, pode gastar cerca de <b>" + eurTxt_(restante / d.diasRestantes) + "</b> por dia.";
  else dica = "Ainda tem <b>" + eurTxt_(restante) + "</b> disponíveis nesta categoria.";

  var corpo =
    '<div style="font-size:12px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;color:#6a7183;">Categoria · ' + escHtml_(d.mesRotulo) + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:6px 0 18px;"><tr>' +
    '<td style="font-size:26px;line-height:32px;font-weight:800;color:#15171e;">' + escHtml_(d.categoria) + '</td>' +
    '<td align="right" style="font-size:30px;line-height:32px;font-weight:800;color:' + t.cor + ';white-space:nowrap;">' + Math.round(d.progresso) + '%</td></tr></table>' +
    barra + stats +
    '<div style="margin-top:20px;padding:14px 16px;background:' + t.fundo + ';border-left:4px solid ' + t.cor + ';border-radius:8px;font-size:14px;line-height:21px;color:#2b2f3a;">' + dica + '</div>';

  return {
    assunto: emoji + " " + titulo + ": " + d.categoria + " (" + Math.round(d.progresso) + "%)",
    texto: titulo + "\nCategoria: " + d.categoria + " (" + d.mesRotulo + ")\nOrçamento: " + eurTxt_(d.orcamento) + "\nGasto: " + eurTxt_(d.gasto) + " (" + Math.round(d.progresso) + "%)\n" + (restante >= 0 ? "Disponível: " : "Excedido: ") + eurTxt_(Math.abs(restante)),
    html: emailBase_({ preheader: d.categoria + ": " + Math.round(d.progresso) + "% do orçamento de " + d.mesRotulo + " já utilizado.", tema: t, emoji: emoji, kicker: "Alerta de meta", titulo: titulo, corpo: corpo, url: d.url })
  };
}

/* ---------- Lembretes de vencimento ---------- */
function selo_(dias, tipo) {
  var vermelho = { bg: "#fdd9dd", fg: "#b4232f" }, ambar = { bg: "#fdecc8", fg: "#8f5200" }, azul = { bg: "#d6f0fc", fg: "#075e85" };
  var x, txt;
  if (tipo === "fecho") { x = ambar; txt = dias <= 0 ? "FECHA HOJE" : "FECHA AMANHÃ"; }
  else if (dias < 0) { x = vermelho; txt = "VENCIDA"; }
  else if (dias === 0) { x = vermelho; txt = "VENCE HOJE"; }
  else if (dias === 1) { x = ambar; txt = "VENCE AMANHÃ"; }
  else { x = azul; txt = "VENCE EM " + dias + " DIAS"; }
  return '<span style="display:inline-block;margin-top:6px;padding:3px 9px;border-radius:999px;background:' + x.bg + ';color:' + x.fg + ';font-size:10.5px;font-weight:800;letter-spacing:.5px;white-space:nowrap;">' + txt + '</span>';
}

/** itens: [{tipo:'fatura'|'fecho'|'prestacao', titulo, detalhe, valor|null, dias, dataTxt}] */
function emailLembretes_(itens, dataHoje, url) {
  var secoes = [["fatura", "💳", "Faturas de cartão"], ["fecho", "🔒", "Fecho de fatura"], ["prestacao", "📌", "Prestações de dívidas"]];
  var total = 0, urgentes = 0;
  itens.forEach(function (i) { if (i.tipo !== "fecho") total += i.valor || 0; if (i.dias <= 0 && i.tipo !== "fecho") urgentes++; });

  var corpo = "";
  secoes.forEach(function (s) {
    var lista = itens.filter(function (i) { return i.tipo === s[0]; }).sort(function (a, b) { return a.dias - b.dias; });
    if (!lista.length) return;
    corpo += '<div style="font-size:12px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;color:#6a7183;margin:' + (corpo ? "26px" : "0") + ' 0 4px;">' + s[1] + "&nbsp; " + s[2] + '</div>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
      lista.map(function (i) {
        return '<tr><td style="padding:14px 0;border-top:1px solid #eceef6;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
          '<td valign="top"><div style="font-size:15px;line-height:20px;font-weight:700;color:#15171e;">' + escHtml_(i.titulo) + '</div>' +
          '<div style="font-size:13px;line-height:19px;color:#6a7183;margin-top:2px;">' + escHtml_(i.detalhe) + '</div></td>' +
          '<td valign="top" align="right" style="padding-left:12px;white-space:nowrap;">' +
          (i.valor != null ? '<div style="font-size:16px;font-weight:800;color:#15171e;">' + eurTxt_(i.valor) + '</div>' : "") + selo_(i.dias, i.tipo) + '</td></tr></table></td></tr>';
      }).join("") + '</table>';
  });

  var n = itens.length;
  var sub = n + (n === 1 ? " aviso" : " avisos") + " para hoje, " + escHtml_(dataHoje) + (total > 0 ? " &middot; <b style=\"color:#ffffff;\">" + eurTxt_(total) + "</b> a pagar" : "");
  var tema = urgentes > 0 ? TEMAS_.perigo : TEMAS_.marca;
  return {
    assunto: "🔔 " + (urgentes > 0 ? "Vencimentos a precisar de atenção" : "Lembretes de vencimento") + " — " + dataHoje,
    texto: itens.map(function (i) { return "- " + i.titulo + " (" + i.detalhe + ")" + (i.valor != null ? ": " + eurTxt_(i.valor) : ""); }).join("\n"),
    html: emailBase_({ preheader: n + (n === 1 ? " vencimento" : " vencimentos") + " a acompanhar, total de " + eurTxt_(total) + ".", tema: tema, emoji: "🔔", kicker: "Lembretes financeiros", titulo: "Vencimentos a acompanhar", subtitulo: sub, corpo: corpo, url: url })
  };
}

/* ---------- Envio ---------- */
function enviarEmail_(msg, tipo) {
  var email = destinatarios_(tipo).join(",");
  if (!email) return false;
  try {
    MailApp.sendEmail({ to: email, subject: msg.assunto, body: msg.texto, htmlBody: msg.html, name: "Performance Financial" });
    return true;
  } catch (e) { console.error("E-mail: " + e); return false; }
}

function enviarTesteNotificacao(token) {
  exigirSessao_(token, "DEFINICOES");
  var email = obterEmailUsuario_();
  if (!email) throw new Error("Preencha o seu e-mail na folha USUÁRIO (célula C2).");
  var hoje = new Date();
  var msg = emailMeta_({ categoria: "Supermercado (exemplo)", orcamento: 600, gasto: 498, progresso: 83, nivel: 80, mesRotulo: MESES_[hoje.getMonth()] + " " + hoje.getFullYear(), diasRestantes: 9 });
  msg.assunto = "🧪 Teste — " + msg.assunto;
  if (!enviarEmail_(msg)) throw new Error("O envio falhou. Veja o registo de execuções do Apps Script.");
  return email.split(",").map(function (e) { return e.replace(/^(.).*(@.*)$/, "$1***$2"); }).join(", ");
}

/* ---------- Verificações ---------- */
function verificarAlertasMetas_(c, mes, ano) {
  if (!destinatarios_("METAS").length) return;
  var ref = mes + "/" + ano, hoje = new Date();
  var ehMesAtual = hoje.getMonth() + 1 === mes && hoje.getFullYear() === ano;
  var diasRest = ehMesAtual ? new Date(ano, mes, 0).getDate() - hoje.getDate() + 1 : null;

  var itens = metasDoMes_(c, mes, ano).map(function (m) { return { chave: m.categoria, rotulo: m.categoria, orc: m.orcamento, gasto: m.gasto, progresso: m.progresso }; });
  var orcT = orcamentoDe_(c, TOTAL_, mes, ano);
  if (orcT > 0) {
    var gT = resumoMes_(c, mes, ano).despesa;
    itens.push({ chave: TOTAL_, rotulo: "Orçamento total do mês", orc: orcT, gasto: arred_(gT), progresso: Math.max(0, Math.round(gT / orcT * 100)) });
  }
  var sheet = null;
  itens.forEach(function (m) {
    if (m.orc <= 0 || m.progresso < 50) return;
    var idx = m.progresso >= 100 ? 2 : m.progresso >= 80 ? 1 : 0;
    if (!sheet) sheet = garantirOrc_(c);
    var linha = 0;
    metasLinhas_(c).forEach(function (x) { if (norm_(x.cat) === norm_(m.chave) && x.ref && x.ref.mes === mes && x.ref.ano === ano) linha = x.linha; });
    if (!linha) linha = anexarLinha_(sheet, [m.chave, m.orc, false, false, false, ref], { 6: "@" });

    var flags = sheet.getRange(linha, 3, 1, 3).getValues()[0].map(function (f) { return f === true || String(f).toUpperCase() === "TRUE"; });
    if (flags[idx]) return;

    var msg = emailMeta_({ categoria: m.rotulo, orcamento: m.orc, gasto: m.gasto, progresso: m.progresso, nivel: [50, 80, 100][idx], mesRotulo: MESES_[mes - 1] + " " + ano, diasRestantes: diasRest });
    if (!enviarEmail_(msg, "METAS")) return; // só marca como enviado se o envio funcionou
    for (var k = 0; k <= idx; k++) flags[k] = true;
    sheet.getRange(linha, 3, 1, 3).setValues([flags]);
  });
}

/* Avisa 3 dias antes, 1 dia antes, no próprio dia e 1 dia depois. Cada aviso é enviado uma só vez. */
var DIAS_AVISO_ = [3, 1, 0];

function diasEntre_(alvo, hoje) { return Math.round((alvo - hoje) / 86400000); }
function dataSegura_(ano, mes, dia) { return new Date(ano, mes - 1, Math.min(dia, new Date(ano, mes, 0).getDate())); }

function verificarLembretes_(c) {
  if (!destinatarios_("VENCIMENTOS").length) return;
  var hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  var props = PropertiesService.getScriptProperties(), enviados = {};
  try { enviados = JSON.parse(props.getProperty("LEMBRETES_ENVIADOS") || "{}"); } catch (e) { enviados = {}; }
  var limite = hoje.getTime() - 90 * 86400000;
  Object.keys(enviados).forEach(function (k) { if (enviados[k] < limite) delete enviados[k]; });

  var itens = [], chaves = [];
  var add = function (chave, item) { if (!enviados[chave]) { itens.push(item); chaves.push(chave); } };

  cartoes_(c).forEach(function (ct) {
    for (var off = -1; off <= 1; off++) {
      var ref = new Date(hoje.getFullYear(), hoje.getMonth() + off, 1), m = ref.getMonth() + 1, a = ref.getFullYear();
      var est = estadoFatura_(c, ct.nome, m, a), total = est.temPagamentos ? est.restante : est.total;
      var venc = dataSegura_(a, m, ct.venc), dias = diasEntre_(venc, hoje);
      if (total > 0 && !est.paga && (DIAS_AVISO_.indexOf(dias) !== -1 || dias === -1))
        add("v|" + ct.nome + "|" + m + "/" + a + "|" + dias, { tipo: "fatura", titulo: "Cartão " + ct.nome, detalhe: "Fatura de " + MESES_[m - 1] + " · vence a " + fmtData_(venc), valor: total, dias: dias });
      var mf = ct.venc <= ct.fech ? m - 1 : m, fecho = dataSegura_(a, mf, ct.fech);
      if (diasEntre_(fecho, hoje) === 1)
        add("f|" + ct.nome + "|" + m + "/" + a, { tipo: "fecho", titulo: "Cartão " + ct.nome, detalhe: "A fatura fecha a " + fmtData_(fecho) + ". Compras a partir daí vão para a fatura seguinte.", valor: null, dias: 1 });
    }
  });

  detalhesDivida_(c).forEach(function (p) {
    if (p.pago || !p.data) return;
    var dias = diasEntre_(new Date(p.data.getFullYear(), p.data.getMonth(), p.data.getDate()), hoje);
    var num = (typeof p.numRaw === "string" && /^\d+\s*\/\s*\d+$/.test(p.numRaw.trim())) ? "Prestação " + p.numRaw.trim() + " · " : "";
    if (DIAS_AVISO_.indexOf(dias) !== -1 || dias === -1)
      add("d|" + p.linha + "|" + fmtISO_(p.data) + "|" + dias, { tipo: "prestacao", titulo: p.desc, detalhe: num + "vence a " + fmtData_(p.data), valor: p.valor, dias: dias });
  });

  if (!itens.length) return;
  if (enviarEmail_(emailLembretes_(itens, fmtData_(hoje)), "VENCIMENTOS")) {
    chaves.forEach(function (k) { enviados[k] = hoje.getTime(); });
    props.setProperty("LEMBRETES_ENVIADOS", JSON.stringify(enviados));
  }
}

/** Evita que uma função de rotina, chamável de fora, seja repetida em rajada. */
function umaVezPor_(chave, segundos) {
  var c = CacheService.getScriptCache();
  if (c.get("lim_" + chave)) return false;
  c.put("lim_" + chave, "1", segundos);
  return true;
}

function verificarAlertasDiario() {
  if (!umaVezPor_("alertas", 600)) return;
  var hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  try { comLock_(function () { processarRecorrentes_(criarCtx_(), hoje); }); } catch (e) { console.error("Recorrentes: " + e); }
  var c = criarCtx_();
  try { verificarAlertasMetas_(c, hoje.getMonth() + 1, hoje.getFullYear()); } catch (e) { console.error(e); }
  try { verificarLembretes_(criarCtx_()); } catch (e) { console.error(e); }
  try { enviarResumosPeriodicos_(criarCtx_(), hoje); } catch (e) { console.error(e); }
}

function instalarTriggers() {
  if (Session.getActiveUser().getEmail() !== Session.getEffectiveUser().getEmail()) return false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (["verificarAlertasDiario", "backupSemanal"].indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("verificarAlertasDiario").timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger("backupSemanal").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  return true;
}

function getOrcadoVsRealizado(token, mes, ano) {
  exigirSessao_(token, "REL_ORCADO");
  return metasDoMes_(criarCtx_(), +mes, +ano).map(function (m) {
    var semOrc = m.orcamento <= 0;
    return {
      categoria: m.categoria, orcado: m.orcamento, realizado: m.gasto,
      diferenca: semOrc ? 0 : arred_(m.orcamento - m.gasto),
      status: semOrc ? "Sem Orçamento Definido" : (m.gasto <= m.orcamento ? "Dentro da Meta" : "Estourado")
    };
  });
}

/* ========================================================================
 * RELATÓRIOS POR INTERVALO
 * ====================================================================== */

/** Por data da COMPRA (contas + cartões). Valores negativos (reembolsos/estornos) abatem à categoria. */
/** Critério global das compras de cartão: "fatura" (mês da fatura atribuída pelo banco, por omissão) ou "compra" (data da compra). */
function criterioCartoes_() { return cfg_().criterioCartoes === "compra" ? "compra" : "fatura"; }

/** Data em que uma compra de cartão conta como despesa: o vencimento da fatura do banco ou a data da compra, conforme o critério. */
function dataContagemCartao_(c, r, crit) {
  if (crit !== "fatura" || !r.fatura) return r.data;
  var ct = memo_(c, "ctsPorNome", function () { var o = {}; cartoes_(c).forEach(function (x) { o[norm_(x.nome)] = x; }); return o; })[norm_(r.cartao)];
  return dataSegura_(r.fatura.ano, r.fatura.mes, ct ? ct.venc : 1);
}

function despesasIntervalo_(c, ini, fim) {
  var mapa = {}, crit = criterioCartoes_();
  movimentos_(c, ABA_.DESP).forEach(function (m) {
    if (!m.interno && m.valor !== 0 && m.data >= ini && m.data <= fim) acumular_(mapa, m.cat, m.valor);
  });
  registosCartao_(c).forEach(function (r) {
    var d = dataContagemCartao_(c, r, crit);
    if (d && r.valor !== 0 && d >= ini && d <= fim) acumular_(mapa, r.cat, r.valor);
  });
  return mapaParaLista_(mapa);
}

function intervalo_(iniStr, fimStr) {
  var ini = toDate_(iniStr), fim = toDate_(fimStr);
  if (!ini || !fim || ini > fim) throw new Error("Intervalo de datas inválido.");
  fim.setHours(23, 59, 59, 999);
  return { ini: ini, fim: fim };
}

function getDespesasPorIntervalo(token, iniStr, fimStr) {
  exigirSessao_(token, "REL_INTERVALO");
  var r = intervalo_(iniStr, fimStr);
  var l = despesasIntervalo_(criarCtx_(), r.ini, r.fim).filter(function (x) { return x.total > 0.004; });
  return { labels: l.map(function (x) { return x.nome; }), valores: l.map(function (x) { return x.total; }) };
}

function getMediaHistorica(token, iniStr, fimStr) {
  exigirSessao_(token, "REL_MEDIA");
  var r = intervalo_(iniStr, fimStr);
  var meses = Math.max(1, (r.fim.getFullYear() - r.ini.getFullYear()) * 12 + (r.fim.getMonth() - r.ini.getMonth()) + 1);
  var l = despesasIntervalo_(criarCtx_(), r.ini, r.fim).filter(function (x) { return x.total > 0.004; });
  return { labels: l.map(function (x) { return x.nome; }), medias: l.map(function (x) { return arred_(x.total / meses); }), meses: meses };
}

/* ========================================================================
 * DÍVIDAS
 * ====================================================================== */

function dividasMestre_(c) {
  return memo_(c, "divm", function () {
    var out = [];
    linhas_(c, ABA_.DIV, 10).forEach(function (r, i) {
      var desc = String(r[1] || "").trim(), cap = num_(r[3]);
      if (!desc && cap === 0) return;
      out.push({
        linha: i + 2, dataCompra: toDate_(r[0]), desc: desc, credor: String(r[2] || "").trim(),
        capital: cap, juros: num_(r[4]), n: parseInt(r[5], 10) || 1, dia: parseInt(r[6], 10) || 10,
        inicio: toDate_(r[7]), cat: String(r[8] || "").trim(), id: String(r[9] || "").trim()
      });
    });
    return out;
  });
}

function detalhesDivida_(c) {
  return memo_(c, "divd", function () {
    var out = [];
    linhas_(c, ABA_.DET, 9).forEach(function (r, i) {
      if (!r[2] && !r[3]) return;
      var st = String(r[7] || "Pendente").trim();
      out.push({
        linha: i + 2, numRaw: r[0], data: toDate_(r[1]), desc: String(r[2] || "").trim(),
        valor: num_(r[3]), amort: num_(r[4]), juros: num_(r[5]), saldoApos: num_(r[6]),
        status: st, pago: /^pag[oa]$/i.test(st), idDiv: String(r[8] || "").trim()
      });
    });
    return out;
  });
}

/** Prestações agrupadas pela dívida a que pertencem: pelo ID único (quando existe) ou, em folhas antigas, pelo nome. */
function agrupar_(detalhes) {
  var g = {};
  detalhes.forEach(function (p) { var k = p.idDiv || norm_(p.desc); (g[k] = g[k] || []).push(p); });
  return g;
}

function detalhesDe_(grupos, d) { return (d.id && grupos[d.id]) || grupos[norm_(d.desc)] || []; }

function etiquetaPrest_(raw, idx, total) {
  if (typeof raw === "string" && /^\d+\s*\/\s*\d+$/.test(raw.trim())) return raw.trim();
  return (idx + 1) + "/" + total; // evita o "1/12" convertido em data pelo Sheets
}

function getRelatorioDividas(token, incluirQuitadas) {
  exigirSessao_(token, "DIVIDAS_VER");
  var c = criarCtx_(), grupos = agrupar_(detalhesDivida_(c)), out = [];
  dividasMestre_(c).forEach(function (d) {
    var det = detalhesDe_(grupos, d);
    var pagas = det.filter(function (p) { return p.pago; }).length;
    var total = det.length || d.n;
    var quitada = det.length > 0 && pagas >= det.length;
    if (quitada && !incluirQuitadas) return;
    var saldo = det.length ? det.reduce(function (s, p) { return s + (p.pago ? 0 : p.valor); }, 0) : d.capital + d.juros;
    out.push({
      idMestre: d.linha, id: d.id, dataCompra: fmtData_(d.dataCompra), dataCompraISO: fmtISO_(d.dataCompra),
      inicioISO: fmtISO_(d.inicio), diaVenc: d.dia, categoria: d.cat,
      descricao: d.desc, credor: d.credor, valorCapital: d.capital, totalJuros: d.juros, valorTotal: arred_(d.capital + d.juros),
      totalParcelas: total, parcelasPagas: pagas, parcelasRestantes: Math.max(0, total - pagas),
      saldoDevedor: arred_(saldo), progresso: total > 0 ? Math.round(pagas / total * 100) : 0, quitada: quitada,
      prestacoes: det.map(function (p, idx) {
        return {
          linha: p.linha, num: etiquetaPrest_(p.numRaw, idx, total), data: fmtData_(p.data), amort: p.amort, juros: p.juros,
          total: p.valor, saldoApos: Math.max(0, p.saldoApos), status: p.status, pago: p.pago
        };
      })
    });
  });
  return out;
}

function getConsolidadoDividas(token) {
  exigirSessao_(token, "DIVIDAS_VER");
  var c = criarCtx_(), hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var chaveHoje = hoje.getFullYear() * 12 + hoje.getMonth();
  var totalGeral = 0, saldo = 0, faltam = 0, pagas = 0, ultima = null, meses = {};
  dividasMestre_(c).forEach(function (d) { totalGeral += d.capital + d.juros; });
  detalhesDivida_(c).forEach(function (p) {
    if (p.pago) pagas++; else { faltam++; saldo += p.valor; }
    if (!p.data) return;
    if (!ultima || p.data > ultima) ultima = p.data;
    // O total de cada mês junta TODAS as prestações desse mês (pagas e por pagar): pagar uma não o altera.
    var chave = p.data.getFullYear() + "-" + (p.data.getMonth() < 9 ? "0" : "") + (p.data.getMonth() + 1);
    var m = meses[chave] = meses[chave] || { chave: chave, rotulo: MESES_[p.data.getMonth()] + "/" + p.data.getFullYear(), valorTotalMes: 0, pagoMes: 0, faltaMes: 0, quantidade: 0, quantidadePagas: 0, atrasado: false, atual: p.data.getFullYear() * 12 + p.data.getMonth() === chaveHoje };
    m.valorTotalMes += p.valor; m.quantidade++;
    if (p.pago) { m.pagoMes += p.valor; m.quantidadePagas++; }
    else { m.faltaMes += p.valor; if (p.data.getFullYear() * 12 + p.data.getMonth() < chaveHoje) m.atrasado = true; }
  });
  return {
    totalGeral: arred_(totalGeral), saldoDevedorGeral: arred_(saldo), totalPagas: pagas, totalFaltam: faltam,
    ultimaData: ultima ? fmtData_(ultima) : "---",
    meses: Object.keys(meses).sort().map(function (k) { var m = meses[k]; m.valorTotalMes = arred_(m.valorTotalMes); m.pagoMes = arred_(m.pagoMes); m.faltaMes = arred_(m.faltaMes); return m; })
  };
}

/** Prestações a vencer no mês. No mês atual inclui também as pendentes de meses anteriores (em atraso). */
function dividasDoMes_(c, mes, ano) {
  var out = [], hoje = new Date(), iniMes = inicioMes_(mes, ano);
  hoje.setHours(0, 0, 0, 0);
  var ehAtual = hoje.getMonth() + 1 === mes && hoje.getFullYear() === ano;
  var grupos = agrupar_(detalhesDivida_(c)), credores = {};
  dividasMestre_(c).forEach(function (m) { credores[m.id || norm_(m.desc)] = m.credor; });
  Object.keys(grupos).forEach(function (k) {
    var g = grupos[k];
    g.forEach(function (p, idx) {
      if (!p.data) return;
      var noMes = p.data.getMonth() + 1 === mes && p.data.getFullYear() === ano;
      var atrasada = ehAtual && !p.pago && p.data < iniMes;
      if (!noMes && !atrasada) return;
      out.push({
        linha: p.linha, prestacao: etiquetaPrest_(p.numRaw, idx, g.length), dataVencimento: fmtData_(p.data), iso: fmtISO_(p.data),
        descricao: p.desc, credor: credores[k] || credores[norm_(p.desc)] || "", valorTotal: p.valor, status: p.status, pago: p.pago, atrasada: atrasada,
        dias: p.pago ? null : diasEntre_(new Date(p.data.getFullYear(), p.data.getMonth(), p.data.getDate()), hoje)
      });
    });
  });
  return out.sort(function (a, b) { return a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : 0; });
}

function achaPrestacao_(c, linha) {
  var grupos = agrupar_(detalhesDivida_(c)), achada = null;
  Object.keys(grupos).forEach(function (k) {
    grupos[k].forEach(function (p, idx) { if (p.linha === linha) achada = { p: p, label: etiquetaPrest_(p.numRaw, idx, grupos[k].length) }; });
  });
  if (!achada) throw new Error("Prestação não encontrada. Atualize a página.");
  return achada;
}

/**
 * Marca a prestação como paga. Se "lancar" for verdadeiro, regista também a saída na conta escolhida (folha DESPESAS, código PAG_DIV),
 * para o saldo, as categorias e o orçamento ficarem certos. Use lancar=false se esta prestação já foi registada de outra forma (ex.: paga com o cartão).
 */
function pagarPrestacao(token, dados) {
  exigirSessao_(token, "DIVIDAS_GERIR");
  var linha = +dados.linhaDet, lancar = dados.lancar === true || dados.lancar === "sim";
  return comLock_(function () {
    var c = criarCtx_(), a = achaPrestacao_(c, linha);
    if (a.p.pago) throw new Error("Esta prestação já está paga.");
    var data = toDate_(dados.data), conta = String(dados.conta || "").trim();
    if (lancar) {
      if (!data) throw new Error("Data de pagamento inválida.");
      if (!conta) throw new Error("Escolha a conta de onde saiu o pagamento.");
    }
    sheetObrig_(c.ss, ABA_.DET).getRange(linha, 8).setValue("Pago");
    if (lancar) {
      var mestre = dividasMestre_(c).filter(function (m) { return a.p.idDiv ? m.id === a.p.idDiv : norm_(m.desc) === norm_(a.p.desc); })[0];
      var cat = (mestre && mestre.cat) || "Dívidas";
      anexarLinha_(sheetObrig_(c.ss, ABA_.DESP), [data, "PAG_DIV", "Prestação " + a.label + " · " + a.p.desc, cat, conta, a.p.valor], { 1: "dd/mm/yyyy" });
    }
    registar_("Prestação paga", a.p.desc + " " + a.label + (lancar ? " · lançada em " + conta : " · só marcada"));
    return true;
  });
}

/** Volta a pôr a prestação como pendente e remove a despesa que tinha sido lançada ao pagá-la (se existir). */
function desfazerPagamentoPrestacao(token, linhaDet) {
  exigirSessao_(token, "DIVIDAS_GERIR");
  var linha = +linhaDet;
  return comLock_(function () {
    var c = criarCtx_(), a = achaPrestacao_(c, linha);
    sheetObrig_(c.ss, ABA_.DET).getRange(linha, 8).setValue("Pendente");
    var alvo = "Prestação " + a.label + " · " + a.p.desc, achada = 0;
    movimentos_(c, ABA_.DESP).forEach(function (m) {
      if (m.cod === "PAG_DIV" && m.desc === alvo && Math.abs(m.valor - a.p.valor) < 0.005) achada = m.linha;
    });
    if (achada) c.ss.getSheetByName(ABA_.DESP).deleteRow(achada);
    registar_("Pagamento de prestação desfeito", a.p.desc + " " + a.label);
    return { despesaRemovida: !!achada };
  });
}

function getDividasDoMes(token, mes, ano) {
  exigirSessao_(token, ["DASHBOARD_VER","DIVIDAS_VER"]);
  return dividasDoMes_(criarCtx_(), +mes, +ano);
}

/** Prestações com cêntimos certos (a última ajusta o total) e vencimento seguro em meses curtos. */
function gerarCronograma_(principal, jurosTotais, n, diaVenc, dataInicio) {
  var amortBase = arred_(principal / n), jurBase = arred_(jurosTotais / n);
  var saldo = arred_(principal + jurosTotais), amortAc = 0, jurAc = 0, out = [];
  for (var p = 1; p <= n; p++) {
    var amort = p === n ? arred_(principal - amortAc) : amortBase;
    var jur = p === n ? arred_(jurosTotais - jurAc) : jurBase;
    amortAc += amort; jurAc += jur;
    var total = arred_(amort + jur);
    saldo = Math.max(0, arred_(saldo - total));
    var ref = new Date(dataInicio.getFullYear(), dataInicio.getMonth() + (p - 1), 1);
    var ultimo = new Date(ref.getFullYear(), ref.getMonth() + 1, 0).getDate();
    out.push({ num: p + "/" + n, data: new Date(ref.getFullYear(), ref.getMonth(), Math.min(diaVenc, ultimo)), amort: amort, juros: jur, total: total, saldo: saldo });
  }
  return out;
}

function escreverCronograma_(ss, v, statusAnteriores) {
  var sDet = sheetObrig_(ss, ABA_.DET);
  var cron = gerarCronograma_(v.principal, v.juros, v.n, v.dia, v.inicio);
  var linha = Math.max(sDet.getLastRow() + 1, 2);
  sDet.getRange(linha, 1, cron.length, 1).setNumberFormat("@");
  sDet.getRange(linha, 2, cron.length, 1).setNumberFormat("dd/mm/yyyy");
  var vals = cron.map(function (p, i) { return [p.num, p.data, v.descricao, p.total, p.amort, p.juros, p.saldo, (statusAnteriores && statusAnteriores[i]) || "Pendente", v.id || ""]; });
  sDet.getRange(linha, 1, vals.length, 9).setValues(vals);
}

function normalizarDivida_(d) {
  var r = {
    dataCompra: toDate_(d.dataCompra) || new Date(), descricao: String(d.descricao || "").trim(), credor: String(d.credor || "").trim(),
    principal: arred_(num_(d.valorPrincipal)), juros: arred_(num_(d.totalJuros)), n: parseInt(d.numParcelas, 10),
    dia: parseInt(d.diaVencimento, 10), inicio: toDate_(d.dataIniciacao), categoria: String(d.categoria || "").trim()
  };
  if (!r.descricao) throw new Error("Indique a descrição da dívida.");
  if (!(r.principal > 0)) throw new Error("O valor principal deve ser maior que zero.");
  if (r.juros < 0) throw new Error("Os juros não podem ser negativos.");
  if (!(r.n >= 1 && r.n <= 600)) throw new Error("Número de parcelas inválido.");
  if (!(r.dia >= 1 && r.dia <= 31)) throw new Error("Dia de vencimento inválido.");
  if (!r.inicio) throw new Error("Indique a data da primeira prestação.");
  return r;
}

function valoresMestre_(v) { return [v.dataCompra, v.descricao, v.credor, v.principal, v.juros, v.n, v.dia, v.inicio, v.categoria, v.id || ""]; }

/** Recalcula as prestações mas MANTÉM o estado (pago/pendente) de cada uma, por posição. */
function atualizarDivida(token, dados) {
  exigirSessao_(token, "DIVIDAS_GERIR");
  var v = normalizarDivida_(dados), linhaMestre = +dados.idMestre;
  return comLock_(function () {
    var c = criarCtx_();
    var masters = dividasMestre_(c);
    var antiga = masters.filter(function (m) { return m.linha === linhaMestre; })[0];
    if (!antiga) throw new Error("Dívida não encontrada. Atualize a página.");
    if (masters.some(function (m) { return m.linha !== linhaMestre && norm_(m.desc) === norm_(v.descricao); })) throw new Error("Já existe outra dívida com essa descrição.");
    v.id = antiga.id || novoId_("D");

    var antigas = detalhesDe_(agrupar_(detalhesDivida_(c)), antiga);
    var status = antigas.map(function (p) { return p.status; });
    var sDet = sheetObrig_(c.ss, ABA_.DET);
    antigas.map(function (p) { return p.linha; }).sort(function (a, b) { return b - a; }).forEach(function (l) { sDet.deleteRow(l); });

    var sDiv = sheetObrig_(c.ss, ABA_.DIV);
    sDiv.getRange(linhaMestre, 1).setNumberFormat("dd/mm/yyyy");
    sDiv.getRange(linhaMestre, 8).setNumberFormat("dd/mm/yyyy");
    sDiv.getRange(linhaMestre, 1, 1, 10).setValues([valoresMestre_(v)]);
    escreverCronograma_(c.ss, v, status);
    registar_("Dívida editada", antiga.desc + (antiga.desc !== v.descricao ? " → " + v.descricao : ""));
    return true;
  });
}

function atualizarStatusPrestacao(token, linhaDet, novoStatus) {
  exigirSessao_(token, "DIVIDAS_GERIR");
  if (["Pago", "Pendente"].indexOf(novoStatus) === -1) throw new Error("Estado inválido.");
  return comLock_(function () {
    var sDet = sheetObrig_(getSS_(), ABA_.DET);
    if (linhaDet < 2 || linhaDet > sDet.getLastRow()) throw new Error("Prestação não encontrada.");
    sDet.getRange(linhaDet, 8).setValue(novoStatus);
    return true;
  });
}

function atualizarPrestacao(token, dados) {
  exigirSessao_(token, "DIVIDAS_GERIR");
  var amort = arred_(num_(dados.amortizacao)), juros = arred_(num_(dados.juros)), linha = +dados.linhaDet;
  if (amort < 0 || juros < 0) throw new Error("Os valores não podem ser negativos.");
  return comLock_(function () {
    var sDet = sheetObrig_(getSS_(), ABA_.DET);
    if (linha < 2 || linha > sDet.getLastRow()) throw new Error("Prestação não encontrada.");
    var n = Math.min(9, sDet.getMaxColumns());
    var todas = sDet.getRange(2, 1, sDet.getLastRow() - 1, n).getValues();
    var alvo = todas[linha - 2], idAlvo = String(alvo[8] || "").trim(), descAlvo = norm_(alvo[2]);
    sDet.getRange(linha, 4, 1, 3).setValues([[arred_(amort + juros), amort, juros]]);

    // recalcula o saldo devedor após cada prestação da mesma dívida
    var idx = [];
    todas.forEach(function (r, i) {
      var id = String(r[8] || "").trim();
      if (idAlvo ? id === idAlvo : norm_(r[2]) === descAlvo) idx.push(i);
    });
    var acumulado = 0;
    for (var k = idx.length - 1; k >= 0; k--) {
      sDet.getRange(idx[k] + 2, 7).setValue(arred_(acumulado));
      acumulado += num_(todas[idx[k]][3]) + (idx[k] + 2 === linha ? (arred_(amort + juros) - num_(todas[idx[k]][3])) : 0);
    }
    registar_("Prestação editada", String(alvo[2]) + " linha " + linha);
    return true;
  });
}

/* ========================================================================
 * TRANSFERÊNCIAS E ACERTOS DE SALDO
 * Uma transferência move dinheiro entre contas SEM ser receita nem despesa.
 * Um acerto (reconciliação) corrige a diferença entre o saldo da app e o do banco.
 * ====================================================================== */

function transferencias_(c) {
  return memo_(c, "trf", function () {
    var out = [];
    linhas_(c, ABA_.TRF, 6).forEach(function (r, i) {
      var d = toDate_(r[0]), v = num_(r[3]);
      if (!d || !v) return;
      out.push({ linha: i + 2, data: d, origem: String(r[1] || "").trim(), destino: String(r[2] || "").trim(), valor: v, desc: String(r[4] || "").trim(), id: String(r[5] || "") });
    });
    return out;
  });
}

function ajustes_(c) {
  return memo_(c, "aju", function () {
    var out = [];
    linhas_(c, ABA_.AJU, 4).forEach(function (r, i) {
      var d = toDate_(r[0]), v = num_(r[2]);
      if (!d || !v) return;
      out.push({ linha: i + 2, data: d, conta: String(r[1] || "").trim(), valor: v, desc: String(r[3] || "").trim() });
    });
    return out;
  });
}

function salvarTransferencia(token, dados) {
  exigirSessao_(token, "LANCAR");
  var data = toDate_(dados.data), valor = arred_(num_(dados.valor));
  if (!data) throw new Error("Data inválida.");
  if (!(valor > 0)) throw new Error("Indique um valor maior que zero.");
  if (mesmoNome_(dados.origem, dados.destino)) throw new Error("A origem e o destino têm de ser contas diferentes.");
  return comLock_(function () {
    var c = criarCtx_(), nomes = contas_(c).map(function (x) { return x.nome; });
    var o = nomes.filter(function (n) { return mesmoNome_(n, dados.origem); })[0], d = nomes.filter(function (n) { return mesmoNome_(n, dados.destino); })[0];
    if (!o || !d) throw new Error("Escolha duas contas existentes.");
    var desc = String(dados.descricao || "").trim() || "Transferência";
    var s = abaOuCria_(c.ss, ABA_.TRF, CAB_TRF_);
    anexarLinha_(s, [data, o, d, valor, desc, novoId_("T")], { 1: "dd/mm/yyyy" });
    registar_("Transferência", o + " → " + d + " · " + valor);
    return true;
  });
}

/** Põe o saldo da conta igual ao do banco: grava a diferença como acerto (não conta como receita nem despesa). */
function ajustarSaldo(token, dados) {
  exigirSessao_(token, "CONTAS_GERIR");
  var data = toDate_(dados.data) || new Date(), real = arred_(num_(dados.saldoReal));
  if (String(dados.saldoReal == null ? "" : dados.saldoReal).trim() === "") throw new Error("Indique o saldo que o banco mostra.");
  return comLock_(function () {
    var c = criarCtx_();
    var ct = contas_(c).filter(function (x) { return mesmoNome_(x.nome, dados.conta); })[0];
    if (!ct) throw new Error("Conta não encontrada.");
    var fim = new Date(data.getFullYear(), data.getMonth(), data.getDate(), 23, 59, 59, 999);
    var atual = saldoConta_(c, ct, fim), delta = arred_(real - atual);
    if (Math.abs(delta) < 0.005) throw new Error("O saldo já coincide com o do banco. Nada a acertar.");
    var s = abaOuCria_(c.ss, ABA_.AJU, CAB_AJU_);
    anexarLinha_(s, [data, ct.nome, delta, String(dados.motivo || "").trim() || "Acerto de saldo"], { 1: "dd/mm/yyyy" });
    registar_("Acerto de saldo", ct.nome + " · " + (delta > 0 ? "+" : "") + delta);
    return { diferenca: delta, saldoApp: arred_(atual) };
  });
}

/* ========================================================================
 * LANÇAMENTOS RECORRENTES (salário, renda, subscrições)
 * Gerados automaticamente todos os dias pelas 8h (ver instalarTriggers) ou pelo botão "Gerar agora".
 * ====================================================================== */

function recorrentes_(c) {
  return memo_(c, "rfix", function () {
    var out = [];
    linhas_(c, ABA_.RFIX, 11).forEach(function (r, i) {
      var id = String(r[0] || "").trim();
      if (!id) return;
      out.push({
        linha: i + 2, id: id, tipo: norm_(r[1]).replace(/[^a-z]/g, ""), desc: String(r[2] || "").trim(), cat: String(r[3] || "").trim(),
        conta: String(r[4] || "").trim(), valor: num_(r[5]), dia: parseInt(r[6], 10) || 1, inicio: toDate_(r[7]), fim: toDate_(r[8]),
        ultimo: String(r[9] || "").trim(), ativo: norm_(r[10]) !== "nao"
      });
    });
    return out;
  });
}

function getRecorrentes(token) {
  exigirSessao_(token, "DEFINICOES");
  var c = criarCtx_(), hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return recorrentes_(c).map(function (r) {
    var cur = hoje.getFullYear() * 12 + hoje.getMonth(), prox = null;
    for (var i = 0; i < 3; i++) {
      var d = dataSegura_(Math.floor((cur + i) / 12), (cur + i) % 12 + 1, r.dia);
      if (d >= hoje && (!r.ultimo || fmtISO_(d).slice(0, 7) > r.ultimo) && (!r.fim || d <= r.fim)) { prox = d; break; }
    }
    return { id: r.id, tipo: r.tipo, descricao: r.desc, categoria: r.cat, conta: r.conta, valor: r.valor, dia: r.dia, inicio: fmtISO_(r.inicio), fim: fmtISO_(r.fim), ativo: r.ativo, proxima: r.ativo && prox ? fmtData_(prox) : "" };
  });
}

function normalizarRecorrente_(d) {
  var tipo = norm_(d.tipo).replace(/[^a-z]/g, "");
  if (["despesa", "receita", "cartao"].indexOf(tipo) === -1) throw new Error("Tipo inválido.");
  var r = { tipo: tipo, desc: String(d.descricao || "").trim(), cat: String(d.categoria || "").trim(), conta: String(d.conta || "").trim(), valor: arred_(Math.abs(num_(d.valor))), dia: parseInt(d.dia, 10), inicio: toDate_(d.inicio), fim: toDate_(d.fim) };
  if (!r.desc) throw new Error("Indique a descrição.");
  if (!(r.valor > 0)) throw new Error("Indique um valor maior que zero.");
  if (!r.conta) throw new Error("Escolha a conta ou o cartão.");
  if (!(r.dia >= 1 && r.dia <= 31)) throw new Error("Dia do mês inválido.");
  if (!r.inicio) throw new Error("Indique a data de início.");
  if (r.fim && r.fim < r.inicio) throw new Error("A data de fim é anterior à de início.");
  return r;
}

function linhaRecorrente_(r, id, ultimo, ativo) {
  return [id, r.tipo, r.desc, r.cat || "Geral", r.conta, r.valor, r.dia, r.inicio, r.fim || "", ultimo || "", ativo === false ? "Não" : "Sim"];
}

/** Uma regra nova só gera a partir da PRÓXIMA ocorrência, para não duplicar o que já lançou à mão (use "retroativo" para gerar desde o início). */
function salvarRecorrente(token, dados) {
  exigirSessao_(token, "DEFINICOES");
  var r = normalizarRecorrente_(dados);
  return comLock_(function () {
    var c = criarCtx_(), hoje = new Date();
    var s = abaOuCria_(c.ss, ABA_.RFIX, CAB_RFIX_);
    s.getRange("J:J").setNumberFormat("@");
    var ultimo = "";
    if (!(dados.retroativo === true || dados.retroativo === "true")) {
      var dEste = dataSegura_(hoje.getFullYear(), hoje.getMonth() + 1, r.dia);
      var ref = dEste <= hoje ? new Date(hoje.getFullYear(), hoje.getMonth(), 1) : new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
      if (r.inicio > ref) ref = null;                                    // começa no futuro: nada a saltar
      if (ref) ultimo = ref.getFullYear() + "-" + pad2_(ref.getMonth() + 1);
    }
    anexarLinha_(s, linhaRecorrente_(r, novoId_("R"), ultimo, true), { 8: "dd/mm/yyyy", 9: "dd/mm/yyyy", 10: "@" });
    registar_("Recorrente criado", r.desc + " · " + r.valor + " dia " + r.dia);
    return true;
  });
}

function atualizarRecorrente(token, dados) {
  exigirSessao_(token, "DEFINICOES");
  var r = normalizarRecorrente_(dados);
  return comLock_(function () {
    var c = criarCtx_();
    var a = recorrentes_(c).filter(function (x) { return x.id === dados.id; })[0];
    if (!a) throw new Error("Regra não encontrada. Atualize a página.");
    var ativo = dados.ativo === undefined ? a.ativo : (dados.ativo === true || dados.ativo === "true");
    var s = sheetObrig_(c.ss, ABA_.RFIX);
    s.getRange(a.linha, 8, 1, 2).setNumberFormat("dd/mm/yyyy");
    s.getRange(a.linha, 1, 1, 11).setValues([linhaRecorrente_(r, a.id, a.ultimo, ativo)]);
    registar_("Recorrente editado", r.desc);
    return true;
  });
}

function excluirRecorrente(token, id) {
  exigirSessao_(token, "DEFINICOES");
  return comLock_(function () {
    var c = criarCtx_();
    var a = recorrentes_(c).filter(function (x) { return x.id === id; })[0];
    if (!a) throw new Error("Regra não encontrada. Atualize a página.");
    sheetObrig_(c.ss, ABA_.RFIX).deleteRow(a.linha);
    registar_("Recorrente apagado", a.desc);
    return true;
  });
}

/** Gera os lançamentos em falta de todas as regras ativas (até hoje). Nunca repete um mês já gerado. */
function processarRecorrentes_(c, hoje) {
  var sR = c.ss.getSheetByName(ABA_.RFIX);
  if (!sR) return 0;
  var criadas = 0;
  recorrentes_(c).forEach(function (r) {
    if (!r.ativo || !r.inicio || !r.valor || !r.conta) return;
    var cur = r.inicio.getFullYear() * 12 + r.inicio.getMonth(), m;
    if (r.ultimo && (m = r.ultimo.match(/^(\d{4})-(\d{2})$/))) cur = Math.max(cur, (+m[1]) * 12 + (+m[2] - 1) + 1);
    var limite = hoje.getFullYear() * 12 + hoje.getMonth(), linhas = [], ultimoGer = null, guard = 0;
    var ct = r.tipo === "cartao" ? cartoes_(c).filter(function (x) { return mesmoNome_(x.nome, r.conta); })[0] : null;
    if (r.tipo === "cartao" && !ct) return;
    for (; cur <= limite && guard < 36; cur++, guard++) {
      var ano = Math.floor(cur / 12), mes = cur % 12 + 1, data = dataSegura_(ano, mes, r.dia);
      if (data < r.inicio) continue;
      if (r.fim && data > r.fim) break;
      if (data > hoje) break;
      if (ct) linhas.push([data, "RECOR", r.desc, r.cat, ct.nome, r.valor, faturaDeCompra_(ct, data).data]);
      else linhas.push([data, "RECOR", r.desc, r.cat, r.conta, r.valor]);
      ultimoGer = ano + "-" + pad2_(mes);
    }
    if (!linhas.length) return;
    var aba = r.tipo === "receita" ? ABA_.REC : r.tipo === "cartao" ? ABA_.CART_REG : ABA_.DESP;
    anexarLinhas_(sheetObrig_(c.ss, aba), linhas, r.tipo === "cartao" ? { 1: "dd/mm/yyyy", 7: "dmmm" } : { 1: "dd/mm/yyyy" });
    sR.getRange(r.linha, 10).setNumberFormat("@").setValue(ultimoGer);
    criadas += linhas.length;
    registar_("Recorrente gerado", r.desc + " · " + linhas.length + " lançamento(s)");
  });
  return criadas;
}

function executarRecorrentes(token) {
  exigirSessao_(token, "DEFINICOES");
  return comLock_(function () {
    var hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    return processarRecorrentes_(criarCtx_(), hoje);
  });
}

/* ========================================================================
 * CATEGORIAS (gestão na app)
 * ====================================================================== */

function getCategorias(token) {
  exigirSessao_(token, "DEFINICOES");
  var c = criarCtx_(), usos = { despesa: {}, receita: {} };
  movimentos_(c, ABA_.DESP).forEach(function (m) { if (!m.interno) usos.despesa[norm_(m.cat)] = (usos.despesa[norm_(m.cat)] || 0) + 1; });
  registosCartao_(c).forEach(function (r) { usos.despesa[norm_(r.cat)] = (usos.despesa[norm_(r.cat)] || 0) + 1; });
  movimentos_(c, ABA_.REC).forEach(function (m) { usos.receita[norm_(m.cat)] = (usos.receita[norm_(m.cat)] || 0) + 1; });
  var cats = categorias_(c);
  var f = function (tipo) { return cats[tipo].map(function (x) { return { nome: x.nome, arquivada: x.arquivada, usos: usos[tipo][norm_(x.nome)] || 0 }; }); };
  return { despesa: f("despesa"), receita: f("receita") };
}

function colCategoria_(tipo) {
  if (tipo !== "despesa" && tipo !== "receita") throw new Error("Tipo de categoria inválido.");
  return tipo === "receita" ? 1 : 2;
}

function linhaDaCategoria_(s, col, nome) {
  var n = Math.max(0, s.getLastRow() - 1);
  if (!n) return 0;
  var v = s.getRange(2, col, n, 1).getValues();
  for (var i = 0; i < v.length; i++) if (mesmoNome_(v[i][0], nome)) return i + 2;
  return 0;
}

function proximaLinhaLivre_(s, col) {
  var n = Math.max(0, s.getLastRow() - 1), ultimo = 1;
  if (n) { var v = s.getRange(2, col, n, 1).getValues(); v.forEach(function (r, i) { if (String(r[0]).trim() !== "") ultimo = i + 2; }); }
  return ultimo + 1;
}

function salvarCategoria(token, dados) {
  exigirSessao_(token, "DEFINICOES");
  var col = colCategoria_(dados.tipo), nome = String(dados.nome || "").trim();
  if (!nome) throw new Error("Indique o nome da categoria.");
  return comLock_(function () {
    var c = criarCtx_(), s = sheetObrig_(c.ss, ABA_.CATS);
    if (linhaDaCategoria_(s, col, nome)) throw new Error("Já existe uma categoria com esse nome.");
    s.getRange(proximaLinhaLivre_(s, col), col).setValue(nome);
    registar_("Categoria criada", dados.tipo + " · " + nome);
    return true;
  });
}

function renomearCategoria(token, dados) {
  exigirSessao_(token, "DEFINICOES");
  var col = colCategoria_(dados.tipo), de = String(dados.de || "").trim(), para = String(dados.para || "").trim();
  if (!de || !para) throw new Error("Indique o nome novo.");
  return comLock_(function () {
    var c = criarCtx_(), s = sheetObrig_(c.ss, ABA_.CATS);
    var linha = linhaDaCategoria_(s, col, de);
    if (!linha) throw new Error("Categoria não encontrada. Atualize a página.");
    if (!mesmoNome_(de, para) && linhaDaCategoria_(s, col, para)) throw new Error("Já existe uma categoria com esse nome.");
    s.getRange(linha, col).setValue(para);
    var arq = linhaDaCategoria_(s, col + 2, de);
    if (arq) s.getRange(arq, col + 2).setValue(para);
    var n = 0;
    if (dados.tipo === "despesa") {
      n += renomearEm_(c.ss, ABA_.DESP, 4, de, para) + renomearEm_(c.ss, ABA_.CART_REG, 4, de, para) + renomearEm_(c.ss, ABA_.DIV, 9, de, para);
    } else n += renomearEm_(c.ss, ABA_.REC, 4, de, para);
    renomearEm_(c.ss, ABA_.ORC, 1, de, para);
    var sr = c.ss.getSheetByName(ABA_.RFIX);
    if (sr && sr.getLastRow() >= 2) {
      var nr = sr.getLastRow() - 1, rg = sr.getRange(2, 2, nr, 3), v = rg.getValues(), ch = false;
      v.forEach(function (r) { if ((norm_(r[0]) === "receita") === (dados.tipo === "receita") && mesmoNome_(r[2], de)) { r[2] = para; ch = true; } });
      if (ch) rg.setValues(v);
    }
    registar_("Categoria renomeada", de + " → " + para + " (" + n + " lançamentos)");
    return n;
  });
}

function arquivarCategoria(token, dados) {
  exigirSessao_(token, "DEFINICOES");
  var col = colCategoria_(dados.tipo), nome = String(dados.nome || "").trim(), arquivar = dados.arquivar === true || dados.arquivar === "true";
  return comLock_(function () {
    var c = criarCtx_(), s = sheetObrig_(c.ss, ABA_.CATS), colArq = col + 2;
    var ja = linhaDaCategoria_(s, colArq, nome);
    if (arquivar && !ja) s.getRange(proximaLinhaLivre_(s, colArq), colArq).setValue(nome);
    if (!arquivar && ja) s.getRange(ja, colArq).clearContent();
    registar_(arquivar ? "Categoria arquivada" : "Categoria reativada", nome);
    return true;
  });
}

/* ========================================================================
 * CÓPIAS DE SEGURANÇA (Google Drive)
 * ====================================================================== */

function fazerBackup_() {
  var ss = getSS_(), agora = Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HHmm");
  var nome = "Cópia · " + ss.getName() + " · " + agora;
  var it = DriveApp.getFoldersByName(PASTA_BACKUP_);
  var pasta = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_BACKUP_);
  DriveApp.getFileById(ss.getId()).makeCopy(nome, pasta);
  var ficheiros = [], fi = pasta.getFiles();
  while (fi.hasNext()) ficheiros.push(fi.next());
  ficheiros.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  ficheiros.slice(BACKUPS_MANTER_).forEach(function (f) { f.setTrashed(true); });
  return nome;
}

function fazerBackupAgora(token) {
  exigirSessao_(token, "DEFINICOES");
  var nome = fazerBackup_();
  registar_("Cópia de segurança", nome);
  return nome;
}

/** Trigger semanal (domingo, 3h). */
function backupSemanal() {
  if (!umaVezPor_("backup", 1800)) return;
  try { registar_("Cópia de segurança automática", fazerBackup_()); } catch (e) { console.error("Backup: " + e); }
  try { limparLog_(); } catch (e) { console.error(e); }
}

function limparLog_() {
  var s = getSS_().getSheetByName(ABA_.LOG);
  if (s && s.getLastRow() > 3000) s.deleteRows(2, s.getLastRow() - 2000);
}

/* ========================================================================
 * SEGURANÇA, DEFINIÇÕES E LOG
 * ====================================================================== */

function alterarSenha(token, atual, nova) {
  var user = exigirSessao_(token, "PROPRIO");
  atual = String(atual || "").trim(); nova = String(nova || "").trim();
  if (nova.length < 8) throw new Error("A nova senha deve ter pelo menos 8 caracteres.");
  if (nova === atual) throw new Error("A nova senha tem de ser diferente da atual.");
  return comLock_(function () {
    var sheet = getSS_().getSheetByName(ABA_.USR);
    var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < dados.length; i++) {
      if (norm_(dados[i][0]) !== norm_(user)) continue;
      var g = String(dados[i][1]).trim(), ok = g.indexOf("h1$") === 0 ? g === hashSenha_(atual) : g === atual;
      if (!ok) throw new Error("A senha atual não está correta.");
      sheet.getRange(i + 2, 2).setValue(hashSenha_(nova));
      limparSessoesLongas_(user, token);                      // os outros dispositivos têm de entrar de novo
      registar_("Senha alterada", user);
      return true;
    }
    throw new Error("Utilizador não encontrado.");
  });
}

function cfg_() {
  var base = { resumoSemanal: true, resumoMensal: true };
  try { var o = JSON.parse(PropertiesService.getScriptProperties().getProperty("CFG") || "{}"); Object.keys(o).forEach(function (k) { base[k] = o[k]; }); } catch (e) {}
  return base;
}

function getDefinicoes(token) {
  exigirSessao_(token, "DEFINICOES");
  var cfg = cfg_(), emails = "";
  try { emails = String(getSS_().getSheetByName(ABA_.USR).getRange(2, 3).getValue() || ""); } catch (e) {}
  var triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  return {
    emails: emails, resumoSemanal: !!cfg.resumoSemanal, resumoMensal: !!cfg.resumoMensal, criterioCartoes: cfg.criterioCartoes === "compra" ? "compra" : "fatura",
    rotinasAtivas: triggers.indexOf("verificarAlertasDiario") !== -1, backupAtivo: triggers.indexOf("backupSemanal") !== -1
  };
}

function salvarDefinicoes(token, d) {
  exigirSessao_(token, "DEFINICOES");
  var lista = String(d.emails == null ? "" : d.emails).split(/[;,\s]+/).filter(Boolean);
  var invalidos = lista.filter(function (e) { return !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); });
  if (invalidos.length) throw new Error("E-mail inválido: " + invalidos[0]);
  PropertiesService.getScriptProperties().setProperty("CFG", JSON.stringify({ resumoSemanal: !!d.resumoSemanal, resumoMensal: !!d.resumoMensal, criterioCartoes: d.criterioCartoes === "compra" ? "compra" : "fatura" }));
  getSS_().getSheetByName(ABA_.USR).getRange(2, 3).setValue(lista.join(", "));
  registar_("Definições guardadas", "e-mails: " + lista.length);
  return true;
}

function getLog(token, limite) {
  exigirSessao_(token, "DEFINICOES");
  var c = criarCtx_(), v = linhas_(c, ABA_.LOG, 4), n = Math.min(+limite || 200, 500), out = [];
  for (var i = v.length - 1; i >= 0 && out.length < n; i--) {
    var d = v[i][0] instanceof Date ? v[i][0] : null;
    if (!d) continue;
    out.push({ quando: Utilities.formatDate(d, tz_(), "dd/MM/yyyy HH:mm"), utilizador: String(v[i][1] || ""), acao: String(v[i][2] || ""), detalhe: String(v[i][3] || "") });
  }
  return out;
}

/* ========================================================================
 * PESQUISA, VISÃO ANUAL E QUALIDADE DOS DADOS
 * ====================================================================== */

function pesquisarMovimentos(token, f) {
  exigirSessao_(token, "MOVIMENTOS_VER");
  f = f || {};
  var c = criarCtx_(), texto = norm_(f.texto), tipo = f.tipo || "todos", cat = norm_(f.categoria), conta = norm_(f.conta);
  var ini = f.ini ? toDate_(f.ini) : null, fim = f.fim ? toDate_(f.fim) : null;
  if (fim) fim.setHours(23, 59, 59, 999);
  var vazio = function (x) { return x === "" || x == null; };
  var min = vazio(f.min) ? null : num_(f.min), max = vazio(f.max) ? null : num_(f.max);
  var out = [], somaRec = 0, somaDesp = 0;
  var passa = function (data, desc, categoria, ct, valor) {
    if (ini && data < ini) return false;
    if (fim && data > fim) return false;
    if (texto && norm_(desc + " " + categoria + " " + ct).indexOf(texto) === -1) return false;
    if (cat && norm_(categoria) !== cat) return false;
    if (conta && norm_(ct).indexOf(conta) === -1) return false;
    var a = Math.abs(valor);
    if (min !== null && a < min) return false;
    if (max !== null && a > max) return false;
    return true;
  };
  var add = function (o) { out.push(o); };
  if (tipo === "todos" || tipo === "receita") movimentos_(c, ABA_.REC).forEach(function (m) {
    if (m.interno || !passa(m.data, m.desc, m.cat, m.conta, m.valor)) return;
    somaRec += m.valor;
    add({ idLinha: m.linha, tipoAba: ABA_.REC, tipo: "receita", data: fmtData_(m.data), iso: fmtISO_(m.data), descricao: m.desc, categoria: m.cat, conta: m.conta, valor: m.valor });
  });
  if (tipo === "todos" || tipo === "despesa") movimentos_(c, ABA_.DESP).forEach(function (m) {
    if (m.interno || !passa(m.data, m.desc, m.cat, m.conta, m.valor)) return;
    somaDesp += m.valor;
    add({ idLinha: m.linha, tipoAba: ABA_.DESP, tipo: "despesa", data: fmtData_(m.data), iso: fmtISO_(m.data), descricao: m.desc, categoria: m.cat, conta: m.conta, valor: m.valor, reembolso: m.valor < 0 });
  });
  if (tipo === "todos" || tipo === "cartao") registosCartao_(c).forEach(function (r) {
    if (!r.data || !passa(r.data, r.desc, r.cat, r.cartao, r.valor)) return;
    somaDesp += r.valor;
    add({ idLinha: r.linha, tipoAba: ABA_.CART_REG, tipo: "cartao", data: fmtData_(r.data), iso: fmtISO_(r.data), descricao: r.desc, categoria: r.cat, conta: r.cartao, valor: r.valor, reembolso: r.valor < 0 });
  });
  if (tipo === "todos" || tipo === "transferencia") transferencias_(c).forEach(function (x) {
    if (!passa(x.data, x.desc, "Transferência", x.origem + " " + x.destino, x.valor)) return;
    add({ idLinha: x.linha, tipoAba: ABA_.TRF, tipo: "transferencia", data: fmtData_(x.data), iso: fmtISO_(x.data), descricao: x.desc || "Transferência", categoria: "Transferência", conta: x.origem + " → " + x.destino, valor: x.valor });
  });
  if (tipo === "todos" || tipo === "ajuste") ajustes_(c).forEach(function (x) {
    if (!passa(x.data, x.desc, "Acerto de saldo", x.conta, x.valor)) return;
    add({ idLinha: x.linha, tipoAba: ABA_.AJU, tipo: "ajuste", data: fmtData_(x.data), iso: fmtISO_(x.data), descricao: x.desc || "Acerto de saldo", categoria: "Acerto de saldo", conta: x.conta, valor: x.valor });
  });
  out.sort(function (a, b) { return a.iso < b.iso ? 1 : a.iso > b.iso ? -1 : 0; });
  var limite = 600;
  return { itens: out.slice(0, limite), total: out.length, truncado: out.length > limite, receitas: arred_(somaRec), despesas: arred_(somaDesp) };
}

function getVisaoAnual(token, ano, criterio) {
  exigirSessao_(token, "ANUAL_VER");
  ano = +ano;
  // criterio "fatura": as compras de cartão contam no mês da fatura em que o banco as lançou; por omissão (e por escolha do utilizador) contam na fatura; "compra" usa a data da compra
  var porFatura = (criterio || criterioCartoes_()) !== "compra";
  var c = criarCtx_(), rec = [], desp = [], cd = {}, cr = {};
  for (var i = 0; i < 12; i++) { rec.push(0); desp.push(0); }
  var somar = function (mapa, nome, mes, v) {
    var n = String(nome || "").trim() || "Outros", k = norm_(n);
    if (!mapa[k]) { mapa[k] = { nome: n, meses: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], total: 0 }; }
    mapa[k].meses[mes] += v; mapa[k].total += v;
  };
  movimentos_(c, ABA_.REC).forEach(function (m) { if (!m.interno && m.data.getFullYear() === ano) { rec[m.data.getMonth()] += m.valor; somar(cr, m.cat, m.data.getMonth(), m.valor); } });
  movimentos_(c, ABA_.DESP).forEach(function (m) { if (!m.interno && m.data.getFullYear() === ano) { desp[m.data.getMonth()] += m.valor; somar(cd, m.cat, m.data.getMonth(), m.valor); } });
  var cartFat = {}, pagFat = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  registosCartao_(c).forEach(function (r) {
    if (r.fatura && r.fatura.ano === ano) somar(cartFat, r.cartao, r.fatura.mes - 1, r.valor);
    var mi = porFatura ? (r.fatura ? r.fatura.mes - 1 : -1) : (r.data ? r.data.getMonth() : -1);
    var ai = porFatura ? (r.fatura ? r.fatura.ano : 0) : (r.data ? r.data.getFullYear() : 0);
    if (mi >= 0 && ai === ano) { desp[mi] += r.valor; somar(cd, r.cat, mi, r.valor); }
  });
  pagamentos_(c).forEach(function (p) { if (p.data && p.data.getFullYear() === ano) pagFat[p.data.getMonth()] += p.valor; });
  var meses = [], tr = 0, td = 0;
  for (var m = 0; m < 12; m++) {
    tr += rec[m]; td += desp[m];
    meses.push({ mes: m + 1, rotulo: MESES_[m].slice(0, 3), receita: arred_(rec[m]), despesa: arred_(desp[m]), balanco: arred_(rec[m] - desp[m]) });
  }
  var lista = function (mapa) {
    return Object.keys(mapa).map(function (k) { var x = mapa[k]; return { nome: x.nome, meses: x.meses.map(arred_), total: arred_(x.total) }; })
      .filter(function (x) { return Math.abs(x.total) > 0.004; }).sort(function (a, b) { return b.total - a.total; });
  };
  return {
    ano: ano, criterio: porFatura ? "fatura" : "compra", meses: meses,
    cartoesFatura: lista(cartFat), pagamentosFaturas: pagFat.map(arred_),
    totais: { receita: arred_(tr), despesa: arred_(td), balanco: arred_(tr - td), poupanca: tr > 0 ? Math.round((tr - td) / tr * 100) : 0 },
    categoriasDespesa: lista(cd), categoriasReceita: lista(cr)
  };
}

/** Linhas com datas ou valores que a app não consegue ler (ficam fora das estatísticas até serem corrigidas). */
function getQualidadeDados(token) {
  exigirSessao_(token, "DEFINICOES");
  var c = criarCtx_(), problemas = [], faturasTexto = 0;
  [[ABA_.DESP, 6], [ABA_.REC, 6], [ABA_.CART_REG, 7]].forEach(function (par) {
    linhas_(c, par[0], par[1]).forEach(function (r, i) {
      var vazia = r.every(function (x) { return x === "" || x == null; });
      if (vazia) return;
      var tem = String(r[2] || "").trim() || String(r[5] == null ? "" : r[5]).trim();
      if (!tem) return;
      if (!toDate_(r[0])) problemas.push({ aba: par[0], linha: i + 2, texto: "Data inválida: \"" + String(r[0] instanceof Date ? "" : r[0]) + "\" · " + String(r[2] || "") });
      var v = r[5];
      if (v !== "" && v != null && typeof v !== "number" && isNaN(parseFloat(String(v).replace(",", ".")))) problemas.push({ aba: par[0], linha: i + 2, texto: "Valor inválido: \"" + v + "\" · " + String(r[2] || "") });
      if (par[0] === ABA_.CART_REG) {
        if (typeof r[6] === "string" && r[6].trim()) faturasTexto++;
        if (r[6] !== "" && r[6] != null && !parseFatura_(r[6], toDate_(r[0]))) problemas.push({ aba: par[0], linha: i + 2, texto: "Fatura ilegível: \"" + r[6] + "\" · " + String(r[2] || "") });
      }
    });
  });
  return { problemas: problemas.slice(0, 100), total: problemas.length, faturasTexto: faturasTexto };
}

/* ========================================================================
 * RESUMO PERIÓDICO POR E-MAIL (semanal e mensal, por categoria)
 * ====================================================================== */

function resumoIntervalo_(c, ini, fim) {
  var cats = despesasIntervalo_(c, ini, fim);
  var receita = somaMov_(c, ABA_.REC, ini, fim), despesa = cats.reduce(function (s, x) { return s + x.total; }, 0);
  return { receita: receita, despesa: despesa, cats: cats };
}

/** d: {tipo, rotulo, ant:{receita,despesa,cats}, atual:{...}, estouradas:[{categoria,orcamento,gasto}], url?} */
function emailResumo_(d) {
  var a = d.atual, p = d.ant, bal = a.receita - a.despesa;
  var var_ = function (x, y) { return y > 0.004 ? Math.round((x - y) / y * 100) : null; };
  var dv = var_(a.despesa, p.despesa);
  var celula = function (rotulo, valor, cor, borda) {
    return '<td width="33%" valign="top" style="padding:14px 12px;' + (borda ? "border-left:1px solid #e4e7f0;" : "") + '"><div style="font-size:11px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:#6a7183;">' + rotulo + '</div><div style="font-size:17px;font-weight:800;color:' + (cor || "#15171e") + ';margin-top:4px;white-space:nowrap;">' + valor + '</div></td>';
  };
  var stats = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e4e7f0;border-radius:12px;border-collapse:separate;"><tr>' +
    celula("Receitas", eurTxt_(a.receita), "#12a150", false) + celula("Despesas", eurTxt_(a.despesa), "#d92d3f", true) +
    celula("Balanço", eurTxt_(bal), bal >= 0 ? "#12a150" : "#d92d3f", true) + '</tr></table>';
  var comp = dv === null ? "" : '<div style="margin-top:14px;font-size:14px;line-height:21px;color:#2b2f3a;">Gastou <b>' + Math.abs(dv) + '% ' + (dv > 0 ? "mais" : "menos") + '</b> do que no período anterior (' + eurTxt_(p.despesa) + ').</div>';

  var topo = a.cats.filter(function (x) { return x.total > 0.004; }).slice(0, 6), maximo = topo.length ? topo[0].total : 1;
  var antMap = {};
  p.cats.forEach(function (x) { antMap[norm_(x.nome)] = x.total; });
  var linhas = topo.map(function (x) {
    var pct = Math.max(2, Math.round(x.total / maximo * 100)), ant = antMap[norm_(x.nome)] || 0, v = var_(x.total, ant);
    var seta = v === null ? "" : '<span style="font-size:12px;color:' + (v > 0 ? "#d92d3f" : "#12a150") + ';font-weight:700;"> ' + (v > 0 ? "▲" : "▼") + " " + Math.abs(v) + "%</span>";
    return '<tr><td style="padding:10px 0;border-top:1px solid #eceef6;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
      '<td style="font-size:14px;font-weight:700;color:#15171e;">' + escHtml_(x.nome) + '</td><td align="right" style="font-size:14px;font-weight:800;color:#15171e;white-space:nowrap;">' + eurTxt_(x.total) + seta + '</td></tr>' +
      '<tr><td colspan="2" style="padding-top:6px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td width="' + pct + '%" height="6" style="background:' + TEMAS_.marca.cor + ';border-radius:3px;font-size:0;line-height:0;">&nbsp;</td>' + (pct < 100 ? '<td width="' + (100 - pct) + '%" height="6" style="font-size:0;line-height:0;">&nbsp;</td>' : "") + '</tr></table></td></tr></table></td></tr>';
  }).join("");
  var cats = topo.length ? '<div style="font-size:12px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;color:#6a7183;margin:26px 0 4px;">Onde gastou mais</div><table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + linhas + '</table>' : "";

  var est = (d.estouradas || []).length
    ? '<div style="margin-top:22px;padding:14px 16px;background:' + TEMAS_.perigo.fundo + ';border-left:4px solid ' + TEMAS_.perigo.cor + ';border-radius:8px;font-size:14px;line-height:21px;color:#2b2f3a;"><b>Metas ultrapassadas:</b><br>' +
      d.estouradas.map(function (m) { return escHtml_(m.categoria) + " — " + eurTxt_(m.gasto) + " de " + eurTxt_(m.orcamento); }).join("<br>") + '</div>' : "";

  var semanal = d.tipo === "semanal";
  var titulo = semanal ? "O seu resumo da semana" : "O seu resumo de " + d.rotulo;
  return {
    assunto: (semanal ? "📊 Resumo semanal" : "📈 Resumo de " + d.rotulo) + " — gastou " + eurTxt_(a.despesa),
    texto: titulo + "\nReceitas: " + eurTxt_(a.receita) + "\nDespesas: " + eurTxt_(a.despesa) + "\nBalanço: " + eurTxt_(bal),
    html: emailBase_({ preheader: "Receitas " + eurTxt_(a.receita) + " · Despesas " + eurTxt_(a.despesa) + " · Balanço " + eurTxt_(bal), tema: TEMAS_.marca, emoji: semanal ? "📊" : "📈", kicker: semanal ? "Resumo semanal" : "Resumo mensal", titulo: titulo, subtitulo: escHtml_(d.periodoTxt), corpo: stats + comp + cats + est, url: d.url })
  };
}

function montarResumo_(c, tipo, hoje) {
  var ini, fim, iniA, fimA, rotulo, periodoTxt;
  if (tipo === "semanal") {
    ini = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 7); fim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 1, 23, 59, 59, 999);
    iniA = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate() - 7); fimA = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate() - 1, 23, 59, 59, 999);
    rotulo = "semana"; periodoTxt = fmtData_(ini) + " a " + fmtData_(fim);
  } else {
    var m = mesAnterior_(hoje.getMonth() + 1, hoje.getFullYear()), mm = mesAnterior_(m.mes, m.ano);
    ini = inicioMes_(m.mes, m.ano); fim = fimMes_(m.mes, m.ano); iniA = inicioMes_(mm.mes, mm.ano); fimA = fimMes_(mm.mes, mm.ano);
    rotulo = MESES_[m.mes - 1] + " " + m.ano; periodoTxt = rotulo;
    var estouradas = metasDoMes_(c, m.mes, m.ano).filter(function (x) { return x.orcamento > 0 && x.gasto > x.orcamento; });
  }
  return emailResumo_({ tipo: tipo, rotulo: rotulo, periodoTxt: periodoTxt, atual: resumoIntervalo_(c, ini, fim), ant: resumoIntervalo_(c, iniA, fimA), estouradas: estouradas || [] });
}

function enviarResumosPeriodicos_(c, hoje) {
  var cfg = cfg_(), props = PropertiesService.getScriptProperties(), chave = fmtISO_(hoje);
  [["semanal", hoje.getDay() === 1, cfg.resumoSemanal], ["mensal", hoje.getDate() === 1, cfg.resumoMensal]].forEach(function (x) {
    if (!x[1] || !x[2] || props.getProperty("RESUMO_" + x[0]) === chave) return;
    if (enviarEmail_(montarResumo_(c, x[0], hoje), x[0] === "semanal" ? "SEMANAL" : "MENSAL")) props.setProperty("RESUMO_" + x[0], chave);
  });
}

function enviarResumoTeste(token, tipo) {
  exigirSessao_(token, "DEFINICOES");
  if (!obterEmailUsuario_()) throw new Error("Preencha primeiro o seu e-mail nas Definições.");
  var msg = montarResumo_(criarCtx_(), tipo === "semanal" ? "semanal" : "mensal", new Date());
  msg.assunto = "🧪 Teste — " + msg.assunto;
  if (!enviarEmail_(msg)) throw new Error("O envio falhou. Veja o registo de execuções do Apps Script.");
  return true;
}

/* ========================================================================
 * MIGRAÇÃO / ESTRUTURA (executar UMA vez; é seguro repetir)
 * Faz primeiro uma cópia de segurança. Não apaga nem altera lançamentos existentes.
 * ====================================================================== */

function migrarEstrutura_() {
  var notas = [], ss = getSS_();
  try { notas.push("Cópia de segurança: " + fazerBackup_()); } catch (e) { notas.push("⚠ Não foi possível fazer a cópia de segurança (" + e.message + "). Faça uma cópia manual antes de continuar."); }
  var c = criarCtx_();

  abaOuCria_(ss, ABA_.TRF, CAB_TRF_); abaOuCria_(ss, ABA_.AJU, CAB_AJU_); abaOuCria_(ss, ABA_.RFIX, CAB_RFIX_); abaOuCria_(ss, ABA_.LOG, ["DATA/HORA", "UTILIZADOR", "AÇÃO", "DETALHE"]);
  var novoOrc = !ss.getSheetByName(ABA_.ORC);
  garantirOrc_(c);
  if (novoOrc) notas.push("Folha ORÇAMENTOS criada (os orçamentos já datados de METAS foram copiados).");
  notas.push("Folhas TRANSFERÊNCIAS, AJUSTES, RECORRENTES e LOG prontas.");

  var cab = function (aba, col, texto) { var s = ss.getSheetByName(aba); if (s && s.getMaxColumns() < col) s.insertColumnsAfter(s.getMaxColumns(), col - s.getMaxColumns()); if (s && !String(s.getRange(1, col).getValue() || "").trim()) s.getRange(1, col).setValue(texto).setFontWeight("bold"); };
  cab(ABA_.CONTAS, 6, "ESTADO"); cab(ABA_.CATS, 3, "ARQUIVADAS (RECEITA)"); cab(ABA_.CATS, 4, "ARQUIVADAS (DESPESA)");
  cab(ABA_.FAT, 9, "TIPO"); cab(ABA_.DIV, 10, "ID"); cab(ABA_.DET, 9, "ID DÍVIDA"); cab(ABA_.CART_REG, 8, "ID PARTILHA"); cab(ABA_.DESP, 8, "ID PARTILHA");
  garantirAcessos_(ss); notas.push("Folha ACESSOS pronta (perfis e permissões).");

  // IDs únicos nas dívidas
  var mestres = dividasMestre_(c), sDiv = ss.getSheetByName(ABA_.DIV), sDet = ss.getSheetByName(ABA_.DET), idPorNome = {}, novosIds = 0;
  mestres.forEach(function (m) {
    if (!m.id) { m.id = novoId_("D"); sDiv.getRange(m.linha, 10).setValue(m.id); novosIds++; }
    if (!idPorNome[norm_(m.desc)]) idPorNome[norm_(m.desc)] = m.id;
  });
  var ligadas = 0;
  detalhesDivida_(c).forEach(function (p) {
    if (p.idDiv || !idPorNome[norm_(p.desc)]) return;
    sDet.getRange(p.linha, 9).setValue(idPorNome[norm_(p.desc)]); ligadas++;
  });
  notas.push("Dívidas: " + novosIds + " IDs novos, " + ligadas + " prestações ligadas pelo ID.");

  // faturas escritas como texto ("12dez/2026") passam a datas reais
  var sReg = ss.getSheetByName(ABA_.CART_REG), convertidas = 0;
  if (sReg && sReg.getLastRow() >= 2) {
    var n = sReg.getLastRow() - 1, rgF = sReg.getRange(2, 7, n, 1), vf = rgF.getValues(), datas = sReg.getRange(2, 1, n, 1).getValues(), nomes = sReg.getRange(2, 5, n, 1).getValues();
    var cartoes = cartoes_(c);
    vf.forEach(function (r, i) {
      if (typeof r[0] !== "string" || !r[0].trim()) return;
      var ref = parseFatura_(r[0], toDate_(datas[i][0])), ct = cartoes.filter(function (x) { return mesmoNome_(x.nome, nomes[i][0]); })[0];
      if (!ref) return;
      r[0] = dataSegura_(ref.ano, ref.mes, ct ? ct.venc : 12); convertidas++;
    });
    if (convertidas) { rgF.setNumberFormat("dmmm"); rgF.setValues(vf); }
  }
  notas.push("Faturas em texto convertidas em datas: " + convertidas + ".");

  aplicarValidacoes_(ss);
  notas.push("Validação de dados e avisos de proteção aplicados.");
  registar_("Estrutura preparada", notas.join(" | "));
  return notas;
}

function aplicarValidacoes_(ss) {
  var lista = function (aba, col, rangeStr) {
    var s = ss.getSheetByName(aba), r = ss.getRange(rangeStr);
    if (!s) return;
    s.getRange(2, col, Math.max(1, s.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(r, true).setAllowInvalid(true).build());
  };
  lista(ABA_.DESP, 4, ABA_.CATS + "!B2:B500"); lista(ABA_.REC, 4, ABA_.CATS + "!A2:A500"); lista(ABA_.CART_REG, 4, ABA_.CATS + "!B2:B500");
  lista(ABA_.DESP, 5, ABA_.CONTAS + "!B2:B500"); lista(ABA_.REC, 5, ABA_.CONTAS + "!B2:B500"); lista(ABA_.CART_REG, 5, ABA_.CART + "!B2:B200");
  [ABA_.DESP, ABA_.REC, ABA_.CART_REG].forEach(function (aba) {
    var s = ss.getSheetByName(aba);
    if (!s) return;
    var n = Math.max(1, s.getMaxRows() - 1);
    s.getRange(2, 1, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText("Data no formato dd/mm/aaaa").build());
    s.getRange(2, 6, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(-1000000000, 1000000000).setAllowInvalid(true).setHelpText("Número. Valores negativos = reembolso/estorno.").build());
  });
  // cabeçalhos protegidos (só aviso) — evita estragar sem querer a linha 1 de cada folha
  Object.keys(ABA_).forEach(function (k) {
    var s = ss.getSheetByName(ABA_[k]);
    if (!s || s.getProtections(SpreadsheetApp.ProtectionType.RANGE).some(function (p) { return p.getDescription() === "Cabeçalho (app)"; })) return;
    s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).protect().setDescription("Cabeçalho (app)").setWarningOnly(true);
  });
}

/** Menu da folha: prepara a estrutura e mostra o resultado. */
function migrarEstrutura() {
  SpreadsheetApp.getUi();                       // só funciona dentro da planilha (nunca a partir da web app)
  var notas = migrarEstrutura_();
  SpreadsheetApp.getUi().alert("Estrutura preparada", notas.join("\n\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}

function backupMenu() {
  SpreadsheetApp.getUi();                       // só funciona dentro da planilha (nunca a partir da web app)
  SpreadsheetApp.getUi().alert("Cópia criada", fazerBackup_(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function ativarRotinasMenu() {
  SpreadsheetApp.getUi();                       // só funciona dentro da planilha (nunca a partir da web app)
  var ok = instalarTriggers();
  SpreadsheetApp.getUi().alert(ok ? "Rotinas ativadas: alertas, lançamentos recorrentes e resumos todos os dias às 8h; cópia de segurança ao domingo às 3h." : "Execute esta opção com a conta dona da folha.");
}

/* ========================================================================
 * CORRIGIR PRESTAÇÕES ANTIGAS
 * Prestações "x/n" lançadas todas com a data da compra original passam a ter a data do mês em que caem (mesmo dia).
 * ====================================================================== */

function getSeriesParcelas(token) {
  exigirSessao_(token, "DEFINICOES");
  var c = criarCtx_(), series = {};
  var cts = {};
  cartoes_(c).forEach(function (ct) { cts[norm_(ct.nome)] = ct; });
  registosCartao_(c).forEach(function (r) {
    var m = r.desc.trim().match(/^(.*?)\s*(\d+)\s*\/\s*(\d+)\s*$/);
    if (!m || !r.data || !r.fatura) return;
    var k = norm_(m[1]) + "|" + norm_(r.cartao) + "|" + m[3];
    (series[k] = series[k] || { descricao: m[1].trim() || r.desc, cartao: r.cartao, total: +m[3], itens: [] }).itens.push({ linha: r.linha, num: +m[2], data: r.data, fatura: r.fatura, valor: r.valor });
  });
  var out = [];
  Object.keys(series).forEach(function (k) {
    var sr = series[k], ct = cts[norm_(sr.cartao)], desloc = ct && ct.venc <= ct.fech ? 1 : 0;
    var base = sr.itens.slice().sort(function (a, b) { return a.num - b.num || a.data - b.data; })[0], dia = base.data.getDate();
    var itens = [], vistos = {}, repetidas = false;
    sr.itens.forEach(function (it) { if (vistos[it.num]) repetidas = true; vistos[it.num] = true; });
    if (repetidas) { out.push({ descricao: sr.descricao, cartao: sr.cartao, revisar: true, itens: [] }); return; }   // números repetidos: ambíguo, corrigir à mão
    sr.itens.forEach(function (it) {
      if (it.num === 1) return;                // a 1.ª prestação mantém a data real da compra
      var ref = addMesRef_(it.fatura, -desloc), alvo = dataSegura_(ref.ano, ref.mes, dia);
      if (fmtISO_(alvo) === fmtISO_(it.data)) return;
      itens.push({ linha: it.linha, num: it.num + "/" + sr.total, valor: it.valor, fatura: ABREV_[it.fatura.mes - 1] + "/" + it.fatura.ano, dataAtual: fmtData_(it.data), dataProposta: fmtData_(alvo), isoProposta: fmtISO_(alvo) });
    });
    if (itens.length) out.push({ descricao: sr.descricao, cartao: sr.cartao, itens: itens.sort(function (a, b) { return parseInt(a.num, 10) - parseInt(b.num, 10); }) });
  });
  return out;
}

function aplicarCorrecaoParcelas(token, alteracoes) {
  exigirSessao_(token, "DEFINICOES");
  return comLock_(function () {
    var s = sheetObrig_(getSS_(), ABA_.CART_REG), n = 0;
    (alteracoes || []).forEach(function (a) {
      var l = +a.linha, d = toDate_(a.iso);
      if (!d || l < 2 || l > s.getLastRow()) return;
      if (!/\d+\s*\/\s*\d+\s*$/.test(String(s.getRange(l, 3).getValue()).trim())) return;
      s.getRange(l, 1).setNumberFormat("dd/mm/yyyy").setValue(d); n++;
    });
    registar_("Datas de prestações corrigidas", n + " linha(s)");
    return n;
  });
}
