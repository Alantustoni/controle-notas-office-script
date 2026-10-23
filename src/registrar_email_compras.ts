/**
 * Registrar e-mails de Compras — Office Script (Excel Online), v1.3
 *
 * Chamado pelo Power Automate a cada e-mail recebido no Outlook. Lê o corpo
 * HTML, extrai as notas fiscais informadas (NF, empresa, fornecedor, AF,
 * vencimento, valores) e grava nas tabelas tbNotas, tbEmails e tbRevisao.
 *
 * - Idempotente: o mesmo e-mail (ID da mensagem) nunca é registrado duas vezes.
 * - Não adivinha: campo ausente fica em branco e vira pendência na aba Revisao.
 * - Alerta de "possível duplicidade" quando outro e-mail traz os mesmos dados,
 *   sem bloquear o registro.
 * - Protege a planilha contra fórmulas injetadas pelo texto do e-mail.
 *
 * Veja docs/power-automate.md para montar o fluxo.
 */

// ── Configuração ──────────────────────────────────────────────────────────
// Empresas destinatárias que o script reconhece no e-mail (em MAIÚSCULAS).
const EMPRESAS: string[] = ["ALFAMETAL", "BETAVIARIA"];
// Células da aba "Notas recebidas" com o contador de cada empresa (mesma ordem).
const CONTADORES_PAINEL: string[] = ["B6", "F6"];
// Fuso local em horas em relação ao UTC (Brasília = -3).
const FUSO_HORAS = -3;
// ──────────────────────────────────────────────────────────────────────────

type Celula = string | number | boolean;
type Quantia = number | "";
interface Nota {
  nf: string; empresa: string; fornecedor: string; codigo: string; af: string;
  vencimento: string | number; tipo: string; observacao: string;
  valorNF: Quantia; valorAF: Quantia; diferenca: Quantia; liquido: Quantia;
  origem: string;
}
interface Pendencia { nf: string; tipo: string; detalhe: string; }
interface Leitura { notas: Nota[]; pendencias: Pendencia[]; texto: string; }
interface Resultado {
  status: string; idEmail: string; referencias: number; novasReferencias: number;
  pendencias: number; mensagem: string;
}
interface CelulaHtml { texto: string; colspan: number; rowspan: number; }
interface TabelaHtml { linhas: string[][]; posicao: number; fim: number; indice: number; }
interface EstadoTabela {
  linhas: CelulaHtml[][]; linha: CelulaHtml[]; celula: CelulaHtml | null;
  linhaAberta: boolean; aninhada: boolean; linhaAninhada: boolean; posicao: number;
}
interface Mapa {
  empresa: number; codigo: number; fornecedor: number; nf: number; af: number; vencimento: number;
  valorNF: number; valorAF: number; diferenca: number; liquido: number;
}
interface Bloco { notas: Nota[]; pendencias: Pendencia[]; tabela: number; posicao: number; }

function main(
  workbook: ExcelScript.Workbook,
  idMensagem: string,
  recebidoEmUTC: string,
  remetente: string,
  assunto: string,
  corpoHtml: string,
  temAnexos: boolean = false,
  internetMessageId: string = "",
  linkOriginal: string = ""
): Resultado {
  const idEmail = (internetMessageId || idMensagem || "").trim();
  if (!idEmail || idEmail.length > 4000) throw new Error("Mapeie ID da mensagem ou Internet Message Id. Não use ID da conversa.");
  const instante = Date.parse(recebidoEmUTC);
  if (!Number.isFinite(instante) || !/(Z|[+-]\d{2}:\d{2})$/i.test(recebidoEmUTC)) {
    throw new Error("Data inválida. Mapeie a hora de recebimento original do Outlook, com fuso.");
  }
  const recebidoLocal = instante / 86400000 + 25569 + FUSO_HORAS / 24;
  const tbNotas = tabela(workbook, "tbNotas", cabecalhosNotas());
  const tbEmails = tabela(workbook, "tbEmails", cabecalhosEmails());
  const tbRevisao = tabela(workbook, "tbRevisao", cabecalhosRevisao());
  const emails = linhas(tbEmails), anteriores = linhas(tbNotas), revisoes = linhas(tbRevisao);
  let indiceEmail = emails.findIndex((r) => String(r[0]) === idEmail);
  const anteriorEmail: Celula[] = indiceEmail >= 0 ? emails[indiceEmail] : [];
  const completo = indiceEmail >= 0 && ["Concluído", "Com revisão"].indexOf(String(anteriorEmail[9])) >= 0;
  const versaoAtual = indiceEmail >= 0 && String(anteriorEmail[12]).indexOf("[LEITOR_V13]") === 0;
  if (completo && versaoAtual) {
    atualizarPainel(tbNotas);
    return { status: "Já registrado", idEmail, referencias: Number(anteriorEmail[7]), novasReferencias: 0,
      pendencias: Number(anteriorEmail[8]), mensagem: "v1.3: esta mensagem já foi processada. Nenhuma linha acrescentada." };
  }
  // Novas versões podem complementar uma leitura antiga, nunca apagar notas ou conferências.
  const leitura = interpretar(assunto, corpoHtml);
  if (!remetente.trim()) leitura.pendencias.push({ nf: "", tipo: "Metadados", detalhe: "Remetente não informado. Conferir o fluxo." });
  const registroEmail: Celula[] = [idEmail, seguro(idMensagem), seguro(internetMessageId), recebidoLocal,
    seguro(recebidoEmUTC), seguro(remetente), seguro(assunto), 0, 0, "Processando",
    temAnexos ? "Sim — arquivos não lidos" : "Não informado pelo conector", linkSeguro(linkOriginal), "[LEITOR_V13] Processando"];
  if (indiceEmail < 0) { tbEmails.addRows(-1, [registroEmail]); indiceEmail = tbEmails.getRowCount() - 1; }
  else tbEmails.getRangeBetweenHeaderAndTotal().getRow(indiceEmail).setValues([registroEmail]);

  const ids = new Set<string>(anteriores.map((r) => String(r[16])));
  const usados = new Set<number>();
  const novas: Celula[][] = [];
  for (let i = 0; i < leitura.notas.length; i++) {
    const n = leitura.notas[i];
    const idRegistro = idEmail + "|V13|" + n.origem;
    if (ids.has(idRegistro)) continue;
    const assinatura = assinaturaNota(n);
    // Compatibilidade com IDs e registros gravados pela v1.1/v1.2; comparação por conteúdo.
    const igualMesmoEmail = anteriores.findIndex((r, j) => !usados.has(j) && String(r[15]) === idEmail && assinaturaLinha(r) === assinatura);
    if (igualMesmoEmail >= 0) { usados.add(igualMesmoEmail); continue; }
    const nfs = numerosNF(n.nf);
    const chave = normalizar(n.empresa) + "|" + (n.codigo ? "FOR:" + normalizar(n.codigo) : "NOME:" + normalizar(n.fornecedor)) + "|NF:" + n.nf;
    let vinculo = "Não localizado nesta base";
    const similares = anteriores.filter((r) => nfs.length > 0 && normalizar(String(r[1])) === normalizar(n.empresa) &&
      ((n.codigo && String(r[3]) === n.codigo) || normalizar(String(r[2])) === normalizar(n.fornecedor)) &&
      numerosNF(String(r[0])).some((v) => nfs.indexOf(v) >= 0));
    if (similares.length) vinculo = "Possível vínculo — conferir AF/vencimento/CNPJ/série";
    // Um novo e-mail com os mesmos dados continua sendo registrado e recebe um alerta para conferência.
    // A repetição da mesma mensagem permanece bloqueada pelo ID do e-mail no início do script.
    const reenvioExato = n.nf !== "" && n.empresa !== "" && n.fornecedor !== "" && n.af !== "" && n.vencimento !== "" &&
      anteriores.some((r) => String(r[15]) !== idEmail && assinaturaLinha(r) === assinatura);
    if (reenvioExato) {
      vinculo = "Possível duplicidade - conferir";
      leitura.pendencias.push({ nf: n.nf, tipo: "Possível duplicidade - conferir",
        detalhe: "Mesmos empresa, fornecedor, NF, AF, vencimento, tipo e valores em outra mensagem. Nova linha criada para conferência; não é prova de duplicidade fiscal. Dados recebidos: " + JSON.stringify(n) });
    }
    if (similares.some((r) => String(r[15]) === idEmail) && !versaoAtual && completo) {
      leitura.pendencias.push({ nf: n.nf, tipo: "Conferir migração",
        detalhe: "Há referência anterior da mesma mensagem com dados diferentes. Não foi sobrescrita nem duplicada automaticamente. Dados lidos na v1.3: " + JSON.stringify(n) });
      continue;
    }
    if (similares.length && !reenvioExato) leitura.pendencias.push({ nf: n.nf, tipo: "Possível vínculo entre registros",
      detalhe: "NF e fornecedor coincidem com outra referência, mas há diferenças ou dados incompletos. Linha preservada; conferir AF, vencimento, CNPJ e série. " + n.origem });
    novas.push([seguro(n.nf), seguro(n.empresa), seguro(n.fornecedor), seguro(n.codigo), seguro(n.af),
      typeof n.vencimento === "number" ? n.vencimento : seguro(n.vencimento), seguro(n.tipo), recebidoLocal,
      "Aguardando conferência", vinculo, seguro(n.observacao + " | " + n.origem + " | Referência de e-mail; documento fiscal não verificado."),
      n.valorNF, n.valorAF, n.diferenca, n.liquido, idEmail, idRegistro, seguro(chave)]);
    ids.add(idRegistro);
  }
  if (novas.length) {
    // O formato de texto é aplicado antes da entrada de identificadores com barras/zeros iniciais.
    for (const col of [0, 1, 2, 3, 4, 6, 8, 9, 10, 15, 16, 17]) tbNotas.getRange().getColumn(col).setNumberFormat("@");
    tbNotas.addRows(0, novas);
  }
  const idsRevisao = new Set<string>(revisoes.map((r) => String(r[0])));
  const novasPendencias: Celula[][] = [];
  for (let i = 0; i < leitura.pendencias.length; i++) {
    const p = leitura.pendencias[i], id = idEmail + "|V13P" + (i + 1);
    if (!idsRevisao.has(id)) novasPendencias.push([id, idEmail, seguro(p.nf), seguro(p.tipo), seguro(p.detalhe), "Pendente"]);
  }
  if (novasPendencias.length) tbRevisao.addRows(-1, novasPendencias);
  // Somente bloqueios técnicos antigos ficam resolvidos ao serem substituídos por nova análise.
  const tecnicos = ["Resposta/encaminhamento", "Formato não reconhecido", "Linha não interpretada", "Linha incompleta", "Empresa ausente/ambígua"];
  for (let i = 0; i < revisoes.length; i++) {
    const r = revisoes[i];
    if (String(r[1]) === idEmail && String(r[0]).indexOf("|V13P") < 0 && tecnicos.indexOf(String(r[3])) >= 0 && String(r[5]) === "Pendente") {
      const atualizado = r.slice(); atualizado[5] = "Resolvida";
      atualizado[4] = seguro(String(r[4]) + " | v1.3: revisão técnica substituída por nova leitura. Ver pendências V13. Isso não confirma documentos nem pagamentos.");
      tbRevisao.getRangeBetweenHeaderAndTotal().getRow(i).setValues([atualizado]);
    }
  }
  formatar(tbNotas, tbEmails);
  destacarDuplicidades(tbNotas);
  atualizarPainel(tbNotas);
  const pendenciasAbertas = linhas(tbRevisao).filter((r) => String(r[1]) === idEmail && String(r[5]) !== "Resolvida").length;
  const status = pendenciasAbertas ? "Com revisão" : "Concluído";
  registroEmail[7] = leitura.notas.length; registroEmail[8] = pendenciasAbertas; registroEmail[9] = status;
  registroEmail[12] = seguro("[LEITOR_V13]\n" + leitura.texto.slice(0, 29000) + (leitura.texto.length > 29000 ? "\n[TRECHO; consultar o e-mail original]" : ""));
  tbEmails.getRangeBetweenHeaderAndTotal().getRow(indiceEmail).setValues([registroEmail]);
  return { status, idEmail, referencias: leitura.notas.length, novasReferencias: novas.length, pendencias: pendenciasAbertas,
    mensagem: "v1.3: linhas interpretadas, incluindo incompletas. Barras preservadas; empresa ausente fica em branco. Não confirma recebimento de documentos ou pagamento." };
}

function cabecalhosNotas(): string[] {
  return ["NF", "Empresa", "Fornecedor", "FOR", "AF / Pedido", "Vencimento", "Comunicação", "Recebido em", "Conferencia", "Possivel_Vinculo", "Observacoes", "Valor_NF_Informado", "Valor_AF_Informado", "Diferenca_Informada", "Liquido_Informado", "ID_Email", "ID_Registro", "Chave_Provisoria"];
}
function cabecalhosEmails(): string[] {
  return ["ID_Email", "Message_ID", "Internet_Message_ID", "Recebido_UTC_3", "Recebido_UTC_Original", "Remetente", "Assunto", "Referencias", "Pendencias", "Status_Captura", "Tem_Anexos", "Link_Original", "Trecho_Corpo"];
}
function cabecalhosRevisao(): string[] { return ["ID_Revisao", "ID_Email", "NF", "Tipo", "Detalhe", "Situacao"]; }
function tabela(w: ExcelScript.Workbook, nome: string, cab: string[]): ExcelScript.Table {
  const t = w.getTable(nome);
  if (!t) throw new Error("Tabela " + nome + " não encontrada. Use a planilha-modelo (modelo/Controle_Notas_modelo.xlsx).");
  if (t.getHeaderRowRange().getValues()[0].map((v) => String(v)).join("|") !== cab.join("|")) throw new Error("Cabeçalhos de " + nome + " alterados. Não renomeie as colunas.");
  if (t.getRowCount() === 1 && t.getRangeBetweenHeaderAndTotal().getValues()[0].every((v) => v === "")) t.deleteRowsAt(0, 1);
  return t;
}
function linhas(t: ExcelScript.Table): Celula[][] { return t.getRowCount() ? t.getRangeBetweenHeaderAndTotal().getValues() : []; }
function formatar(n: ExcelScript.Table, e: ExcelScript.Table): void {
  if (n.getRowCount()) {
    const r = n.getRangeBetweenHeaderAndTotal();
    for (const c of [0, 1, 2, 3, 4, 6, 8, 9, 10, 15, 16, 17]) r.getColumn(c).setNumberFormat("@");
    r.getColumn(5).setNumberFormat("dd/mm/yyyy"); r.getColumn(7).setNumberFormat("dd/mm/yyyy hh:mm");
    for (const c of [11, 12, 13, 14]) r.getColumn(c).setNumberFormat('#,##0.0000;[Red]-#,##0.0000');
  }
  if (e.getRowCount()) e.getRangeBetweenHeaderAndTotal().getColumn(3).setNumberFormat("dd/mm/yyyy hh:mm");
}
function destacarDuplicidades(n: ExcelScript.Table): void {
  if (!n.getRowCount()) return;
  const dados = linhas(n);
  const grupos = new Map<string, number[]>();
  for (let i = 0; i < dados.length; i++) {
    const nf = normalizar(String(dados[i][0]));
    const fornecedor = normalizar(String(dados[i][2]));
    if (!nf || !fornecedor) continue;
    const chave = nf + "|" + fornecedor;
    const indices = grupos.get(chave) || [];
    indices.push(i);
    grupos.set(chave, indices);
  }
  const corpo = n.getRangeBetweenHeaderAndTotal();
  grupos.forEach((indices) => {
    if (indices.length < 2) return;
    for (const indice of indices) {
      const formato = corpo.getRow(indice).getFormat();
      formato.getFill().setColor("#FFF2CC");
      formato.getFont().setColor("#7F6000");
      formato.getFont().setBold(true);
    }
  });
}
function atualizarPainel(n: ExcelScript.Table): void {
  const ws = n.getWorksheet();
  if (ws.getName() !== "Notas recebidas") return;
  // Mantém as 18 colunas. Só altera os dois contadores e legendas do modelo aprovado.
  for (let i = 0; i < EMPRESAS.length && i < CONTADORES_PAINEL.length; i++) {
    ws.getRange(CONTADORES_PAINEL[i]).setFormula('=IFERROR(SUMPRODUCT(--(tbNotas[Empresa]="' + EMPRESAS[i] + '"),--(tbNotas[NF]<>""),LEN(tbNotas[NF])-LEN(SUBSTITUTE(tbNotas[NF],"/",""))+1),0)');
  }
  ws.getRange("B3").setValue("Leitor v1.3 • Campos ausentes ficam em branco. Pendências na aba Revisao.");
  ws.getRange("B10").setValue("Totais de referências (não NFs únicas). Barras somam cada NF; sem empresa/NF não entra nos totais.");
}
function seguro(s: string): string {
  const v = s.replace(/\u0000/g, "").slice(0, 30000);
  return /^[\s]*[=+@-]/.test(v) ? "'" + v : v;
}
function linkSeguro(s: string): string { return /^https:\/\/(outlook\.office\.com|outlook\.office365\.com)\//i.test(s.trim()) ? s.trim() : ""; }
function padraoEmpresas(): string {
  return EMPRESAS.map((e) => normalizar(e).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
}
function normalizar(s: string): string { return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[“”"']/g, "").replace(/\s+/g, " ").trim(); }
function limpar(s: string): string { return s.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim(); }
function entidades(s: string): string {
  const mapa: { [chave: string]: string } = { nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", ndash: "–", mdash: "—", ordm: "º", deg: "°", aacute: "á", eacute: "é", iacute: "í", oacute: "ó", uacute: "ú", agrave: "à", acirc: "â", ecirc: "ê", ocirc: "ô", atilde: "ã", otilde: "õ", ccedil: "ç", Aacute: "Á", Eacute: "É", Iacute: "Í", Oacute: "Ó", Uacute: "Ú", Agrave: "À", Acirc: "Â", Ecirc: "Ê", Ocirc: "Ô", Atilde: "Ã", Otilde: "Õ", Ccedil: "Ç" };
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (inteiro: string, chave: string) => {
    if (chave.charAt(0) === "#") { const n = chave.charAt(1).toLowerCase() === "x" ? parseInt(chave.slice(2), 16) : parseInt(chave.slice(1), 10); return Number.isFinite(n) && n >= 0 && n <= 0x10FFFF ? String.fromCodePoint(n) : inteiro; }
    return mapa[chave] !== undefined ? mapa[chave] : inteiro;
  });
}
function semEstilo(html: string): string { return html.replace(/<!--[\s\S]*?-->/g, "").replace(/<(style|script)\b[^>]*>[\s\S]*?<\/\1>/gi, ""); }
function paraTexto(html: string): string {
  return entidades(semEstilo(html).replace(/<br\b[^>]*>|<\/(?:td|th|tr|p|div|li|table|h[1-6])\s*>/gi, "\n").replace(/<[^>]*>/g, "")).replace(/\r/g, "").replace(/\u00a0/g, " ").trim();
}
function dataBR(s: string): number | "" {
  const m = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/.exec(limpar(s));
  if (!m) return "";
  const d = Number(m[1]), mo = Number(m[2]), a = Number(m[3]);
  const t = new Date(Date.UTC(a, mo - 1, d));
  if (a < 1900 || a > 2200 || t.getUTCFullYear() !== a || t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== d) return "";
  return t.getTime() / 86400000 + 25569;
}
function numerosNF(s: string): string[] { return /^\d+(?:\s*\/\s*\d+)*$/.test(limpar(s)) ? limpar(s).split("/").map((v) => v.trim()) : []; }
function dinheiro(s: string): Quantia {
  const v = s.replace(/R\$/gi, "").replace(/\s/g, "");
  if (!/^-?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d{1,6})?$/.test(v)) return "";
  const n = Number(v.replace(/\./g, "").replace(",", ".")); return Number.isFinite(n) ? n : "";
}
function assinaturaNota(n: Nota): string { return JSON.stringify([normalizar(n.nf), normalizar(n.empresa), normalizar(n.fornecedor), normalizar(n.codigo), normalizar(n.af), n.vencimento, n.tipo, n.valorNF, n.valorAF, n.diferenca, n.liquido]); }
function assinaturaLinha(r: Celula[]): string { return JSON.stringify([normalizar(String(r[0])), normalizar(String(r[1])), normalizar(String(r[2])), normalizar(String(r[3])), normalizar(String(r[4])), r[5], r[6], r[11], r[12], r[13], r[14]]); }
function assuntoSemPrefixos(s: string): string {
  let t = s.trim(); const re = /^(?:ENC|RES|RE|FW|FWD)(?:\s*\[\d+\])?\s*[:：]\s*/i;
  while (re.test(t)) t = t.replace(re, "").trim(); return t;
}

function atributoNumero(tag: string, nome: string): number {
  const re = new RegExp("\\b" + nome + "\\s*=\\s*[\"']?(\\d+)", "i");
  const m = re.exec(tag); return m ? Math.min(100, Math.max(1, Number(m[1]))) : 1;
}
function fecharCelula(e: EstadoTabela): void { if (e.celula) { e.linha.push(e.celula); e.celula = null; } }
function fecharLinha(e: EstadoTabela): void {
  fecharCelula(e);
  if (e.linhaAberta && e.linha.length && !e.linhaAninhada) e.linhas.push(e.linha);
  e.linha = []; e.linhaAberta = false; e.linhaAninhada = false;
}
function grade(linhasOrigem: CelulaHtml[][]): string[][] {
  const saida: string[][] = []; const futuras: { [linha: string]: { [coluna: string]: string } } = {};
  for (let i = 0; i < linhasOrigem.length; i++) {
    const r: string[] = [], ocupado: { [coluna: string]: boolean } = {};
    if (futuras[String(i)]) for (const k of Object.keys(futuras[String(i)])) { r[Number(k)] = futuras[String(i)][k]; ocupado[k] = true; }
    let c = 0;
    for (const cel of linhasOrigem[i]) {
      while (ocupado[String(c)]) c++;
      const valor = limpar(entidades(cel.texto));
      for (let x = 0; x < cel.colspan; x++) {
        r[c + x] = x === 0 ? valor : ""; ocupado[String(c + x)] = true;
        for (let y = 1; y < cel.rowspan && i + y < linhasOrigem.length; y++) {
          if (!futuras[String(i + y)]) futuras[String(i + y)] = {};
          futuras[String(i + y)][String(c + x)] = x === 0 ? valor : "";
        }
      }
      c += cel.colspan;
      if (c > 100) throw new Error("Tabela acima de 100 colunas. Exige revisão do formato.");
    }
    for (let j = 0; j < r.length; j++) if (r[j] === undefined) r[j] = "";
    saida.push(r);
  }
  return saida;
}
function tabelasHTML(html: string): TabelaHtml[] {
  const texto = semEstilo(html), resultado: TabelaHtml[] = [], pilha: EstadoTabela[] = [];
  const re = /<[^>]*>|[^<]+|</g; let m: RegExpExecArray | null;
  while ((m = re.exec(texto)) !== null) {
    const token = m[0]; const atual = pilha.length ? pilha[pilha.length - 1] : null;
    if (/^<table\b/i.test(token)) {
      if (atual) atual.linhaAninhada = true;
      pilha.push({ linhas: [], linha: [], celula: null, linhaAberta: false, linhaAninhada: false, aninhada: pilha.length > 0, posicao: m.index });
    } else if (/^<\/table\s*>/i.test(token)) {
      if (!atual) continue; fecharLinha(atual); pilha.pop();
      resultado.push({ linhas: grade(atual.linhas), posicao: atual.posicao, fim: re.lastIndex, indice: 0 });
    } else if (atual) {
      if (/^<tr\b/i.test(token)) { fecharLinha(atual); atual.linhaAberta = true; }
      else if (/^<\/tr\s*>/i.test(token)) fecharLinha(atual);
      else if (/^<(?:td|th)\b/i.test(token)) {
        fecharCelula(atual); atual.linhaAberta = true;
        atual.celula = { texto: "", colspan: atributoNumero(token, "colspan"), rowspan: atributoNumero(token, "rowspan") };
      } else if (/^<\/(?:td|th)\s*>/i.test(token)) fecharCelula(atual);
      else if (atual.celula) {
        if (token.charAt(0) !== "<") atual.celula.texto += token;
        else if (/^<br\b|^<\/(?:p|div|li)\b/i.test(token)) atual.celula.texto += " ";
      }
    }
  }
  if (pilha.length) throw new Error("HTML de tabela incompleto. Consultar o corpo integral do e-mail.");
  resultado.sort((a, b) => a.posicao - b.posicao);
  for (let i = 0; i < resultado.length; i++) resultado[i].indice = i + 1;
  return resultado;
}
function mapaVazio(): Mapa { return { empresa: -1, codigo: -1, fornecedor: -1, nf: -1, af: -1, vencimento: -1, valorNF: -1, valorAF: -1, diferenca: -1, liquido: -1 }; }
function campoCabecalho(s: string): string {
  let t = normalizar(s).replace(/[.:º°]/g, "").replace(/\s+/g, " ").trim();
  // Alguns e-mails trazem o cabeçalho colado do Excel com a referência da célula.
  if (/^EMPRESA(?:\s+\+AA5:F1972)?$/.test(normalizar(s)) || /^EMPRESA\s*\+AA5F1972$/.test(t)) return "empresa";
  const aliases: { [campo: string]: string[] } = {
    empresa: ["EMPRESA", "EMPRESA DESTINATARIA", "DESTINATARIO", "EMPRESA TOMADORA"],
    codigo: ["FOR", "COD FOR", "CODIGO FORNECEDOR", "CODIGO DO FORNECEDOR", "COD FORNECEDOR"],
    fornecedor: ["NOME", "FORNECEDOR", "NOME FORNECEDOR", "NOME DO FORNECEDOR", "RAZAO SOCIAL", "RAZAO SOCIAL DO FORNECEDOR"],
    nf: ["NF", "NFE", "NF-E", "NFS-E", "NFS", "NOTA FISCAL", "N NF", "NUM NF", "NUMERO NF", "NUMERO DA NF", "NUMERO DA NOTA", "NUMERO DA NOTA FISCAL", "N NOTA FISCAL"],
    af: ["AF", "PEDIDO", "N PEDIDO", "NUMERO DO PEDIDO", "AF / PEDIDO", "AF/PEDIDO", "AUTORIZACAO DE FORNECIMENTO"],
    vencimento: ["VENCIMENTO", "VENC", "VCTO", "VENCTO", "DATA VENCIMENTO", "DATA DE VENCIMENTO", "DT VENCIMENTO", "DT VENC"],
    valorNF: ["VLR NF", "VALOR NF", "VALOR DA NF", "VALOR NOTA FISCAL"],
    valorAF: ["VLR AF", "VALOR AF", "VALOR PEDIDO", "VALOR DO PEDIDO"],
    diferenca: ["DIF", "DIFERENCA"], liquido: ["VLR LIQ", "VALOR LIQUIDO", "LIQUIDO"]
  };
  for (const k of Object.keys(aliases)) if (aliases[k].indexOf(t) >= 0) return k;
  return "";
}
function mapear(linha: string[]): Mapa | null {
  const m = mapaVazio(); const vistos = new Set<string>(); let duplicado = false;
  for (let i = 0; i < linha.length; i++) {
    const k = campoCabecalho(linha[i]); if (!k) continue;
    if (vistos.has(k)) duplicado = true; vistos.add(k);
    switch (k) {
      case "empresa": m.empresa = i; break; case "codigo": m.codigo = i; break; case "fornecedor": m.fornecedor = i; break;
      case "nf": m.nf = i; break; case "af": m.af = i; break; case "vencimento": m.vencimento = i; break;
      case "valorNF": m.valorNF = i; break; case "valorAF": m.valorAF = i; break; case "diferenca": m.diferenca = i; break; case "liquido": m.liquido = i; break;
    }
  }
  return !duplicado && m.fornecedor >= 0 && (m.nf >= 0 || m.af >= 0 || m.vencimento >= 0) ? m : null;
}
function valorEm(r: string[], i: number): string { return i >= 0 && i < r.length ? limpar(r[i]) : ""; }
function vazio(s: string): boolean { return /^(?:|#N\/D|#N\/A|N\/A|N\/D|—|-)$/i.test(limpar(s)); }
function empresaExplicita(contexto: string): string {
  const ls = paraTexto(contexto).split(/\n/).map((s) => normalizar(s));
  const nomes: string[] = [];
  for (const l of ls) {
    if (/^(ATENCIOSAMENTE|ATT|CORDIALMENTE)\b/.test(l)) break;
    const m = new RegExp("^(?:(?:EMPRESA|DESTINATARIO)\\s*[:\\-]\\s*)?(" + padraoEmpresas() + ")$").exec(l);
    if (m && nomes.indexOf(m[1]) < 0) nomes.push(m[1]);
  }
  return nomes.length === 1 ? nomes[0] : "";
}
function nfAssunto(s: string): string {
  const m = /\b(?:NF|NOTA FISCAL)\s*[:#-]?\s*(\d+(?:\s*\/\s*\d+)*)(?![\dA-Za-z])/i.exec(assuntoSemPrefixos(s));
  return m ? limpar(m[1]) : "";
}
function notaDaLinha(r: string[], m: Mapa, tipo: string, empresa: string, origem: string): { nota: Nota | null; pendencias: Pendencia[] } {
  const p: Pendencia[] = [];
  const brutoNF = valorEm(r, m.nf), nome = valorEm(r, m.fornecedor), af = valorEm(r, m.af), venc = valorEm(r, m.vencimento);
  if ([brutoNF, nome, af, venc].every((s) => vazio(s))) return { nota: null, pendencias: [] };
  if (/^(TOTAL|TOTAIS|SUBTOTAL)$/i.test(nome) || r.some((s) => /^(TOTAL|TOTAIS|SUBTOTAL)$/i.test(s))) return { nota: null, pendencias: [] };
  const nf = numerosNF(brutoNF).length ? brutoNF : "";
  if (!nf && vazio(nome) && vazio(af)) return { nota: null, pendencias: [] };
  const data = dataBR(venc);
  const vencimento: string | number = data !== "" ? data : (vazio(venc) ? "" : venc);
  const n: Nota = { nf, empresa: m.empresa >= 0 ? valorEm(r, m.empresa) : empresa, fornecedor: vazio(nome) ? "" : nome,
    codigo: vazio(valorEm(r, m.codigo)) ? "" : valorEm(r, m.codigo), af: vazio(af) ? "" : af,
    vencimento, tipo, observacao: "Dados de origem: " + JSON.stringify(r), valorNF: dinheiro(valorEm(r, m.valorNF)),
    valorAF: dinheiro(valorEm(r, m.valorAF)), diferenca: dinheiro(valorEm(r, m.diferenca)), liquido: dinheiro(valorEm(r, m.liquido)), origem };
  if (vazio(n.empresa)) n.empresa = "";
  if (new RegExp("^(" + padraoEmpresas() + ")$", "i").test(n.empresa)) n.empresa = n.empresa.toUpperCase();
  const avisos: string[] = [];
  if (!nf) avisos.push(brutoNF && !vazio(brutoNF) ? "NF não reconhecida: " + brutoNF : "NF não informada");
  if (!n.fornecedor) avisos.push("Fornecedor não informado");
  if (vencimento === "") avisos.push("Vencimento não informado");
  else if (typeof vencimento === "string" && !/^(?:A VISTA|À VISTA)$/i.test(vencimento)) avisos.push("Vencimento textual preservado: " + vencimento);
  if (avisos.length) {
    n.observacao += " | " + avisos.join("; ");
    if (!nf) n.tipo = "Registro incompleto — NF não informada";
    p.push({ nf, tipo: "Campos incompletos / conferir", detalhe: origem + ": " + avisos.join("; ") + ". Linha preservada. " + JSON.stringify(r) });
  }
  if (nf && n.af && normalizar(nf) === normalizar(n.af)) p.push({ nf, tipo: "NF igual à AF", detalhe: origem + ": números iguais preservados; não se concluiu que há erro." });
  return { nota: n, pendencias: p };
}
function cabecalhoSemRotulos(r: string[]): boolean {
  return r.length === 5 && (vazio(r[0]) || /^\d+$/.test(r[0])) && /[A-Za-zÀ-ÿ]/.test(r[1]) &&
    (vazio(r[2]) || dataBR(r[2]) !== "" || normalizar(r[2]) === "A VISTA") &&
    (vazio(r[3]) || numerosNF(r[3]).length > 0) && (vazio(r[4]) || numerosNF(r[4]).length > 0);
}
function interpretarTabela(t: TabelaHtml, assunto: string, corpo: string): Bloco {
  const b: Bloco = { notas: [], pendencias: [], tabela: t.indice, posicao: t.posicao };
  const sub = normalizar(assuntoSemPrefixos(assunto));
  const contexto = corpo.slice(0, t.posicao);
  const empresa = empresaExplicita(contexto);
  let mapa: Mapa | null = null; let modo = ""; let achou = false; let semRotulos = false; let larguraCabecalho = 0;
  for (let i = 0; i < t.linhas.length; i++) {
    const r = t.linhas[i]; if (r.every((s) => vazio(s))) continue;
    const novo = mapear(r);
    if (novo) {
      mapa = novo; achou = true; semRotulos = false; larguraCabecalho = r.length;
      modo = /^DESCONTO\b/.test(sub) && novo.nf >= 0 ? "Desconto informado" :
        (novo.valorAF >= 0 || /^PAGAMENTO\b/.test(sub)) ? "Ajuste solicitado" : "Envio informado";
      continue;
    }
    if (mapa && modo === "Desconto informado" && r.some((s) => /^VALOR\s*-?\s*(PEDIDO|NOTA FISCAL)/i.test(s))) { mapa = null; continue; }
    if (!mapa && !achou && /NOTA|PAGAMENTO/.test(sub) && cabecalhoSemRotulos(r)) {
      mapa = { empresa: -1, codigo: 0, fornecedor: 1, vencimento: 2, nf: 3, af: 4, valorNF: -1, valorAF: -1, diferenca: -1, liquido: -1 };
      modo = "Envio informado"; semRotulos = true;
    }
    if (!mapa) continue;
    // Cabeçalhos de itens, blocos de instrução e rodapés não são relações de notas.
    if (r.filter((s) => campoCabecalho(s) !== "").length >= 2 || r.some((s) => /^MOTIVO\b|^SOLICITAMOS\b/i.test(s))) { mapa = null; continue; }
    if ((semRotulos && !cabecalhoSemRotulos(r)) || (!semRotulos && r.length < larguraCabecalho)) {
      b.pendencias.push({ nf: "", tipo: "Estrutura de linha diferente", detalhe: "T" + t.indice + "L" + (i + 1) + ": quantidade/estrutura de células incompatível; linha preservada aqui, sem deslocar campos. " + JSON.stringify(r) });
      continue;
    }
    const x = notaDaLinha(r, mapa, modo, empresa, "T" + t.indice + "L" + (i + 1));
    if (x.nota) {
      if (modo === "Desconto informado" && !x.nota.empresa) {
        const em = new RegExp("\\bX\\s+(" + padraoEmpresas() + ")\\b", "i").exec(assuntoSemPrefixos(assunto));
        if (em) x.nota.empresa = em[1].toUpperCase();
      }
      b.notas.push(x.nota);
    }
    b.pendencias.push(...x.pendencias);
  }
  if (b.notas.some((n) => n.tipo === "Desconto informado")) {
    const total = t.linhas.find((r) => r.some((v) => normalizar(v) === "TOTAL"));
    if (total) {
      const k = total.findIndex((v) => normalizar(v) === "TOTAL");
      const nums = total.slice(k + 1).filter((v) => !vazio(v));
      if (nums.length === 3) for (const n of b.notas) if (n.tipo === "Desconto informado") { n.valorNF = dinheiro(nums[0]); n.diferenca = dinheiro(nums[1]); n.liquido = dinheiro(nums[2]); }
    }
  }
  for (const n of b.notas) {
    if (/Desconto|Ajuste/.test(n.tipo)) {
      n.observacao += " | Instruções comunicadas: " + limpar(paraTexto(corpo).slice(0, 1800));
      b.pendencias.push({ nf: n.nf, tipo: "Validar desconto/ajuste", detalhe: n.origem + ": valores e instruções transcritos, não autorizam pagamento. Conferir mensagem original." });
    }
  }
  if (!b.notas.length && !achou) {
    const rotulos: string[] = []; const valores: string[] = [];
    for (const r of t.linhas) if (r.length === 2 && campoCabecalho(r[0])) { rotulos.push(r[0]); valores.push(r[1]); }
    const vertical = mapear(rotulos);
    if (vertical) {
      const x = notaDaLinha(valores, vertical, "Envio informado", empresa, "T" + t.indice + "VERTICAL");
      if (x.nota) b.notas.push(x.nota); b.pendencias.push(...x.pendencias);
    }
  }
  if (!b.notas.length && !achou && /NOTA|PAGAMENTO/.test(sub) && t.linhas.some((r) => r.some((s) => /\d/.test(s)))) {
    b.pendencias.push({ nf: "", tipo: "Layout sem identificação", detalhe: "Tabela " + t.indice + ": sem cabeçalhos reconhecidos e fora da sequência de cinco colunas confirmada. Não adivinhou NF/data. " + JSON.stringify(t.linhas).slice(0, 23000) });
  }
  return b;
}
function lerRotulos(texto: string): { [campo: string]: string[] } {
  const saida: { [campo: string]: string[] } = {};
  const partes = texto.split(/\n|\s*[;|]\s*/).map((v) => limpar(v)).filter((v) => v !== "");
  for (let i = 0; i < partes.length; i++) {
    const m = /^([^:]{1,45}):\s*(.*)$/.exec(partes[i]);
    if (!m) continue;
    const k = campoCabecalho(m[1]); if (!k) continue;
    let v = m[2];
    if (!v && i + 1 < partes.length && !/^[^:]{1,45}:/.test(partes[i + 1])) v = partes[i + 1];
    if (!saida[k]) saida[k] = [];
    if (v && saida[k].indexOf(v) < 0) saida[k].push(v);
  }
  return saida;
}
function interpretarTexto(assunto: string, texto: string): Bloco {
  const b: Bloco = { notas: [], pendencias: [], tabela: 0, posicao: 0 };
  const linhasTexto = texto.split(/\n/); const assinatura = linhasTexto.findIndex((s) => /^(ATENCIOSAMENTE|ATT|CORDIALMENTE)\b/.test(normalizar(s)));
  const textoLimpo = (assinatura >= 0 ? linhasTexto.slice(0, assinatura) : linhasTexto).join("\n");
  const campos = lerRotulos(textoLimpo); const sub = assuntoSemPrefixos(assunto);
  if (/ADIANTAMENTO/i.test(textoLimpo) && /^PAGAMENTO\s*[-:]/i.test(sub) && !campos.nf) {
    const fornecedor = sub.replace(/^PAGAMENTO\s*[-:]\s*/i, "").trim();
    const af = /(?:^|\n)\s*AF\s*[:#-]?\s*(\d+(?:\s*\/\s*\d+)*)\s*(?:\n|$)/i.exec(textoLimpo);
    const n: Nota = { nf: "", empresa: empresaExplicita(textoLimpo), fornecedor, codigo: "", af: af ? limpar(af[1]) : "",
      vencimento: "", tipo: "Adiantamento — sem NF", observacao: "Solicitação de adiantamento, não nota fiscal. " + textoLimpo.slice(0, 6000),
      valorNF: "", valorAF: "", diferenca: "", liquido: "", origem: "TEXTO1" };
    b.notas.push(n); b.pendencias.push({ nf: "", tipo: "Adiantamento sem NF", detalhe: "Comunicação preservada fora dos totais de NF. Conferir o pedido e documentos; não confirma pagamento." }); return b;
  }
  if (Object.keys(campos).some((k) => campos[k].length > 1)) {
    b.pendencias.push({ nf: "", tipo: "Texto com várias referências", detalhe: "Há mais de um valor para rótulos no texto; não misturou fornecedores, NFs e datas. " + textoLimpo.slice(0, 14000) }); return b;
  }
  if (!campos.fornecedor || (!campos.nf && !campos.af && !campos.vencimento)) return b;
  const get = (k: string): string => campos[k] && campos[k].length ? campos[k][0] : "";
  const r = [get("empresa"), get("codigo"), get("fornecedor"), get("af"), get("vencimento"), get("nf")];
  const m: Mapa = { empresa: 0, codigo: 1, fornecedor: 2, af: 3, vencimento: 4, nf: 5, valorNF: -1, valorAF: -1, diferenca: -1, liquido: -1 };
  const x = notaDaLinha(r, m, "Envio informado", "", "TEXTO1");
  if (x.nota) b.notas.push(x.nota); b.pendencias.push(...x.pendencias); return b;
}
function interpretar(assunto: string, corpo: string): Leitura {
  const saida: Leitura = { notas: [], pendencias: [], texto: "" };
  if (corpo.length > 500000) {
    saida.texto = paraTexto(corpo.slice(0, 29000)); saida.pendencias.push({ nf: "", tipo: "Corpo extenso", detalhe: "Corpo acima de 500 mil caracteres; não interpretado. Consultar o original." }); return saida;
  }
  saida.texto = paraTexto(corpo);
  const corpoLimpo = semEstilo(corpo); let tabs: TabelaHtml[] = [];
  try { tabs = tabelasHTML(corpoLimpo); }
  catch {
    saida.pendencias.push({ nf: "", tipo: "Tabela HTML incompleta", detalhe: "O HTML não permitiu preservar a estrutura das células. Conferir o corpo completo antes de registrar." }); return saida;
  }
  const blocos = tabs.map((t) => interpretarTabela(t, assunto, corpoLimpo));
  const uteis = blocos.filter((b) => b.notas.length > 0);
  const historico = assuntoSemPrefixos(assunto) !== assunto.trim() || /divRplyFwdMsg|<blockquote\b|gmail_quote|-----\s*(?:Original Message|Mensagem original)|(?:^|\n)De:[^\n]+\nEnviado:/i.test(corpo);
  if (historico && uteis.length > 1) {
    const assinaturas = uteis.map((b) => JSON.stringify(b.notas.map((n) => assinaturaNota(n))));
    if (assinaturas.some((s) => s !== assinaturas[0])) {
      saida.pendencias.push({ nf: "", tipo: "Histórico com relações diferentes", detalhe: "Mais de uma relação diferente em resposta/encaminhamento. Nenhuma foi selecionada silenciosamente. Conferir as tabelas na mensagem." }); return saida;
    }
    saida.notas = uteis[0].notas; saida.pendencias = uteis[0].pendencias;
    for (const n of saida.notas) n.observacao += " | Relação repetida no histórico considerada uma vez.";
  } else {
    for (const b of blocos) { saida.notas.push(...b.notas); saida.pendencias.push(...b.pendencias); }
  }
  if (!saida.notas.length && !tabs.length) {
    const b = interpretarTexto(assunto, saida.texto); saida.notas.push(...b.notas); saida.pendencias.push(...b.pendencias);
  }
  const citada = nfAssunto(assunto);
  if (citada && saida.notas.length === 1 && saida.notas[0].nf && normalizar(citada) !== normalizar(saida.notas[0].nf)) {
    const n = saida.notas[0];
    saida.pendencias.push({ nf: "", tipo: "Conflito assunto/corpo", detalhe: "Assunto NF " + citada + "; tabela NF " + n.nf + ". Empresa/fornecedor/AF preservados, NF deixada em branco até conferência." });
    n.observacao += " | Conflito de NF: assunto=" + citada + "; corpo=" + n.nf; n.nf = ""; n.tipo = "Conflito de NF — conferir";
  }
  if (historico) for (const n of saida.notas) n.observacao += " | ENC/RES aceito; data de recebimento é a desta mensagem, não do envio original.";
  if (!saida.notas.length && !saida.pendencias.length) saida.pendencias.push({ nf: "", tipo: "Formato não reconhecido", detalhe: "Mensagem preservada, sem extração segura. Não lê anexos/PDF/XML/imagens nem adivinha números e datas sem rótulos." });
  return saida;
}