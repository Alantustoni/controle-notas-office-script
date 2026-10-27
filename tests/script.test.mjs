// Testes do Office Script. Todos os e-mails abaixo são fictícios.
import { test } from "node:test";
import assert from "node:assert/strict";
import { S, planilhaFalsa } from "./carregar.mjs";

const TABELA = `
<p>Boa tarde, seguem as notas recebidas.</p>
<table>
  <tr><th>EMPRESA</th><th>FOR</th><th>FORNECEDOR</th><th>NF</th><th>AF</th><th>VENCIMENTO</th></tr>
  <tr><td>alfametal</td><td>1020</td><td>Fornecedor Exemplo Ltda</td><td>1001</td><td>5501</td><td>10/11/2026</td></tr>
  <tr><td>BETAVIARIA</td><td>3040</td><td>Outro Fornecedor SA</td><td>2002/2003</td><td>5502</td><td>15/11/2026</td></tr>
</table>
<p>Atenciosamente,<br>Compras</p>`;

const excel = (d, m, a) => Date.UTC(a, m - 1, d) / 86400000 + 25569;

// ── parser ──────────────────────────────────────────────────────────────
test("lê a tabela do e-mail e padroniza a empresa", () => {
  const r = S.interpretar("Notas recebidas", TABELA);
  assert.equal(r.notas.length, 2);
  const [a, b] = r.notas;
  assert.equal(a.empresa, "ALFAMETAL");
  assert.equal(a.fornecedor, "Fornecedor Exemplo Ltda");
  assert.equal(a.codigo, "1020");
  assert.equal(a.nf, "1001");
  assert.equal(a.af, "5501");
  assert.equal(a.vencimento, excel(10, 11, 2026));
  assert.equal(b.nf, "2002/2003");           // várias NFs na mesma linha
  assert.equal(r.pendencias.length, 0);
});

test("empresa escrita antes da tabela vale para as linhas sem coluna empresa", () => {
  const html = `<p>Empresa: BETAVIARIA</p><table>
    <tr><td>FORNECEDOR</td><td>NF</td><td>VENCIMENTO</td></tr>
    <tr><td>Fornecedor Exemplo Ltda</td><td>777</td><td>01/12/2026</td></tr></table>`;
  const r = S.interpretar("NF 777", html);
  assert.equal(r.notas[0].empresa, "BETAVIARIA");
});

test("campo ausente fica em branco e vira pendência — não adivinha", () => {
  const html = `<table><tr><td>FORNECEDOR</td><td>NF</td><td>VENCIMENTO</td></tr>
    <tr><td>Fornecedor Exemplo Ltda</td><td>123</td><td></td></tr></table>`;
  const r = S.interpretar("Notas", html);
  assert.equal(r.notas[0].vencimento, "");
  assert.ok(r.pendencias.some((p) => p.tipo === "Campos incompletos / conferir"));
});

test("NF do assunto diferente da tabela: deixa a NF em branco e avisa", () => {
  const html = `<table><tr><td>FORNECEDOR</td><td>NF</td><td>VENCIMENTO</td></tr>
    <tr><td>Fornecedor Exemplo Ltda</td><td>456</td><td>01/12/2026</td></tr></table>`;
  const r = S.interpretar("NF 123", html);
  assert.equal(r.notas[0].nf, "");
  assert.ok(r.pendencias.some((p) => p.tipo === "Conflito assunto/corpo"));
});

test("e-mail sem tabela, com rótulos no texto", () => {
  const html = "<p>Empresa: ALFAMETAL</p><p>Fornecedor: Fornecedor Exemplo Ltda</p><p>NF: 9090</p><p>Vencimento: 05/12/2026</p>";
  const r = S.interpretar("Nota fiscal", html);
  assert.equal(r.notas.length, 1);
  assert.equal(r.notas[0].nf, "9090");
  assert.equal(r.notas[0].empresa, "ALFAMETAL");
});

test("tabela com rowspan/colspan é expandida sem deslocar colunas", () => {
  const html = `<table><tr><td rowspan="2">A</td><td colspan="2">B</td></tr><tr><td>C</td><td>D</td></tr></table>`;
  const [t] = S.tabelasHTML(html);
  assert.deepEqual(t.linhas, [["A", "B", ""], ["A", "C", "D"]]);
});

test("HTML de tabela quebrado não é interpretado pela metade", () => {
  const r = S.interpretar("Notas", "<table><tr><td>FORNECEDOR</td><td>NF</td>");
  assert.equal(r.notas.length, 0);
  assert.ok(r.pendencias.some((p) => p.tipo === "Tabela HTML incompleta"));
});

test("ENC/RES com a mesma relação repetida no histórico conta uma vez", () => {
  const r = S.interpretar("ENC: Notas recebidas", TABELA + "<blockquote>" + TABELA + "</blockquote>");
  assert.equal(r.notas.length, 2);
});

// ── utilitários ─────────────────────────────────────────────────────────
test("valores e datas no formato brasileiro", () => {
  assert.equal(S.dinheiro("R$ 1.234,56"), 1234.56);
  assert.equal(S.dinheiro("12,3,4"), "");
  assert.equal(S.dataBR("31/02/2026"), "");          // data impossível
  assert.equal(S.dataBR("01/01/2026"), excel(1, 1, 2026));
  assert.deepEqual(S.numerosNF("10 / 11/12"), ["10", "11", "12"]);
  assert.equal(S.assuntoSemPrefixos("ENC: RES: RE: Notas"), "Notas");
});

test("texto do e-mail nunca vira fórmula na planilha", () => {
  assert.equal(S.seguro('=HYPERLINK("http://x")'), `'=HYPERLINK("http://x")`);
  assert.equal(S.seguro("+55 11"), "'+55 11");
  assert.equal(S.seguro("Fornecedor"), "Fornecedor");
});

test("link original só é aceito se for do Outlook", () => {
  assert.equal(S.linkSeguro("https://outlook.office365.com/owa/?id=1"), "https://outlook.office365.com/owa/?id=1");
  assert.equal(S.linkSeguro("https://site-qualquer.com/phish"), "");
});

// ── fluxo completo (main) numa planilha falsa ───────────────────────────
const chamar = (wb, id, corpo = TABELA, assunto = "Notas recebidas") =>
  S.main(wb, id, "2026-11-03T13:00:00Z", "compras@exemplo.com", assunto, corpo, false, "", "");

test("registra as notas, o e-mail e os contadores do painel", () => {
  const { workbook, tabelas, celulas } = planilhaFalsa();
  const r = chamar(workbook, "msg-1");
  assert.equal(r.status, "Concluído");
  assert.equal(r.novasReferencias, 2);
  assert.equal(tabelas.tbNotas.dados.length, 2);
  assert.equal(tabelas.tbEmails.dados.length, 1);
  assert.match(String(celulas.B6), /tbNotas\[Empresa\]="ALFAMETAL"/);
  assert.match(String(celulas.F6), /tbNotas\[Empresa\]="BETAVIARIA"/);
  // horário gravado no fuso local (UTC-3): 13:00Z -> 10:00
  const recebido = tabelas.tbNotas.dados[0][7];
  assert.ok(Math.abs(recebido - (Date.UTC(2026, 10, 3, 10) / 86400000 + 25569)) < 1e-9);
});

test("o mesmo e-mail processado de novo não duplica nada", () => {
  const { workbook, tabelas } = planilhaFalsa();
  chamar(workbook, "msg-1");
  const r = chamar(workbook, "msg-1");
  assert.equal(r.status, "Já registrado");
  assert.equal(tabelas.tbNotas.dados.length, 2);
});

test("outro e-mail com os mesmos dados é registrado e marcado para conferência", () => {
  const { workbook, tabelas } = planilhaFalsa();
  chamar(workbook, "msg-1");
  const r = chamar(workbook, "msg-2");
  assert.equal(r.status, "Com revisão");
  assert.equal(tabelas.tbNotas.dados.length, 4);          // não bloqueia
  const novas = tabelas.tbNotas.dados.slice(0, 2);         // novas entram no topo
  assert.ok(novas.every((l) => l[9] === "Possível duplicidade - conferir"));
  assert.ok(tabelas.tbRevisao.dados.some((l) => l[3] === "Possível duplicidade - conferir"));
});

test("data sem fuso é recusada (evita gravar horário errado)", () => {
  const { workbook } = planilhaFalsa();
  assert.throws(() => S.main(workbook, "msg-1", "2026-11-03 13:00", "a@b.com", "x", TABELA), /Data inválida/);
});

// ── casos frequentes em produção (e-mails fictícios com o mesmo formato) ──
test("desconto: lê valores do TOTAL, tira a empresa do assunto e manda para revisão", () => {
  const corpo = `<table>
    <tr><th>FORNECEDOR</th><th>NF</th><th>VALOR NOTA FISCAL</th><th>DIFERENÇA</th><th>VALOR LÍQUIDO</th></tr>
    <tr><td>Gases Modelo Ltda</td><td>8120</td><td>1.500,00</td><td>150,00</td><td>1.350,00</td></tr>
    <tr><td>TOTAL</td><td></td><td>1.500,00</td><td>150,00</td><td>1.350,00</td></tr></table>`;
  const r = S.interpretar("DESCONTO: GASES MODELO X ALFAMETAL", corpo);
  assert.equal(r.notas.length, 1);
  const n = r.notas[0];
  assert.equal(n.tipo, "Desconto informado");
  assert.equal(n.empresa, "ALFAMETAL");
  assert.deepEqual([n.valorNF, n.diferenca, n.liquido], [1500, 150, 1350]);
  assert.ok(r.pendencias.some((p) => p.tipo === "Validar desconto/ajuste"));
});

test("NF igual à AF é preservada e sinalizada, sem concluir que é erro", () => {
  const corpo = `<p>Empresa: ALFAMETAL</p><table>
    <tr><th>FOR</th><th>FORNECEDOR</th><th>NF</th><th>AF</th><th>VENCIMENTO</th></tr>
    <tr><td>2210</td><td>Usinagem Fictícia Ltda</td><td>5530</td><td>5530</td><td>20/11/2026</td></tr></table>`;
  const r = S.interpretar("ENVIO DE NOTA FISCAL", corpo);
  assert.equal(r.notas[0].nf, "5530");
  assert.equal(r.notas[0].af, "5530");
  assert.ok(r.pendencias.some((p) => p.tipo === "NF igual à AF"));
});

test("pedido de adiantamento sem NF fica fora dos totais e vai para revisão", () => {
  const corpo = `<p>Empresa: BETAVIARIA</p><p>Solicitamos adiantamento conforme proposta.</p><p>AF: 5620</p><p>Atenciosamente,<br>Compras</p>`;
  const r = S.interpretar("PAGAMENTO - Montagens Exemplo Ltda", corpo);
  assert.equal(r.notas.length, 1);
  assert.equal(r.notas[0].nf, "");
  assert.equal(r.notas[0].af, "5620");
  assert.equal(r.notas[0].fornecedor, "Montagens Exemplo Ltda");
  assert.ok(r.pendencias.some((p) => p.tipo === "Adiantamento sem NF"));
});

test("e-mail sem tabela nem rótulos ('segue em anexo') não inventa nota", () => {
  const r = S.interpretar("NF Peças Exemplo", "<p>Boa tarde, segue nota em anexo.</p><p>Atenciosamente,<br>Compras</p>");
  assert.equal(r.notas.length, 0);
  assert.ok(r.pendencias.some((p) => p.tipo === "Formato não reconhecido"));
});
