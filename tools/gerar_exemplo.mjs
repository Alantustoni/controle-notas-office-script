// Roda o script sobre 7 e-mails FICTÍCIOS numa planilha simulada e salva o
// conteúdo das tabelas em modelo/exemplo.json. Esse JSON alimenta
// tools/gerar_planilhas.py, que monta o .xlsx de exemplo.
//
//   node tools/gerar_exemplo.mjs
import { writeFileSync } from "node:fs";
import { S, planilhaFalsa } from "../tests/carregar.mjs";

const { workbook, tabelas } = planilhaFalsa();

const email1 = `<p>Boa tarde, seguem as notas recebidas hoje.</p>
<table>
  <tr><th>EMPRESA</th><th>FOR</th><th>FORNECEDOR</th><th>NF</th><th>AF</th><th>VENCIMENTO</th></tr>
  <tr><td>ALFAMETAL</td><td>1020</td><td>Fornecedor Exemplo Ltda</td><td>1001</td><td>5501</td><td>10/11/2026</td></tr>
  <tr><td>ALFAMETAL</td><td>1187</td><td>Comercial Modelo ME</td><td>48213</td><td>5507</td><td>12/11/2026</td></tr>
  <tr><td>BETAVIARIA</td><td>3040</td><td>Transportes Fictícios SA</td><td>2002/2003</td><td>5502</td><td>15/11/2026</td></tr>
</table><p>Atenciosamente,<br>Compras</p>`;

const email2 = `<p>Empresa: BETAVIARIA</p>
<table>
  <tr><td>FORNECEDOR</td><td>NF</td><td>AF</td><td>VENCIMENTO</td></tr>
  <tr><td>Serviços Demonstração Ltda</td><td>777</td><td>5610</td><td>01/12/2026</td></tr>
  <tr><td>Peças Exemplo Eireli</td><td>3315</td><td>5611</td><td></td></tr>
</table>`;

// casos que aparecem com frequência em produção (números agregados no README)
const email4 = `<p>Bom dia, segue desconto negociado.</p>
<table>
  <tr><th>FORNECEDOR</th><th>NF</th><th>VALOR NOTA FISCAL</th><th>DIFERENÇA</th><th>VALOR LÍQUIDO</th></tr>
  <tr><td>Gases Modelo Ltda</td><td>8120</td><td>1.500,00</td><td>150,00</td><td>1.350,00</td></tr>
  <tr><td>TOTAL</td><td></td><td>1.500,00</td><td>150,00</td><td>1.350,00</td></tr>
</table><p>Favor pagar o valor líquido.</p>`;

const email5 = `<p>Empresa: ALFAMETAL</p>
<table>
  <tr><th>FOR</th><th>FORNECEDOR</th><th>NF</th><th>AF</th><th>VENCIMENTO</th></tr>
  <tr><td>2210</td><td>Usinagem Fictícia Ltda</td><td>5530</td><td>5530</td><td>20/11/2026</td></tr>
</table>`;

const email6 = `<p>Empresa: BETAVIARIA</p>
<p>Solicitamos adiantamento conforme proposta.</p>
<p>AF: 5620</p>
<p>Atenciosamente,<br>Compras</p>`;

const email7 = `<p>Boa tarde, segue nota em anexo.</p><p>Atenciosamente,<br>Compras</p>`;

// reenvio do primeiro e-mail, com outro ID: deve gerar alerta de duplicidade
const chamadas = [
  ["msg-exemplo-1", "2026-11-03T13:05:00Z", "Notas recebidas 03/11", email1],
  ["msg-exemplo-2", "2026-11-04T16:40:00Z", "NF 777", email2],
  ["msg-exemplo-3", "2026-11-05T12:15:00Z", "ENC: Notas recebidas 03/11", email1],
  ["msg-exemplo-4", "2026-11-05T14:20:00Z", "DESCONTO: GASES MODELO X ALFAMETAL", email4],
  ["msg-exemplo-5", "2026-11-06T09:10:00Z", "ENVIO DE NOTA FISCAL", email5],
  ["msg-exemplo-6", "2026-11-06T10:45:00Z", "PAGAMENTO - Montagens Exemplo Ltda", email6],
  ["msg-exemplo-7", "2026-11-06T15:30:00Z", "NF Peças Exemplo", email7],
];
for (const [id, quando, assunto, corpo] of chamadas) {
  const r = S.main(workbook, id, quando, "compras@exemplo.com.br", assunto, corpo, false, "", "");
  console.log(id, "->", r.status, `(${r.novasReferencias} nova(s), ${r.pendencias} pendência(s))`);
}

const saida = {
  tbNotas: { cab: tabelas.tbNotas.cab, dados: tabelas.tbNotas.dados },
  tbEmails: { cab: tabelas.tbEmails.cab, dados: tabelas.tbEmails.dados },
  tbRevisao: { cab: tabelas.tbRevisao.cab, dados: tabelas.tbRevisao.dados },
  empresas: S.EMPRESAS,
};
writeFileSync(new URL("../modelo/exemplo.json", import.meta.url), JSON.stringify(saida, null, 1));
console.log("modelo/exemplo.json gravado");
