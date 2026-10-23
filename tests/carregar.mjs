// Carrega o Office Script fora do Excel para os testes.
// O arquivo original não exporta nada (o Office Scripts não aceita export),
// então copiamos o fonte para uma pasta temporária e acrescentamos o export.
// Requer Node >= 22.18 (roda TypeScript direto, removendo os tipos).
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const raiz = join(dirname(fileURLToPath(import.meta.url)), "..");
const fonte = readFileSync(join(raiz, "src", "registrar_email_compras.ts"), "utf8");
const pasta = join(raiz, "tests", ".tmp");
mkdirSync(pasta, { recursive: true });
const alvo = join(pasta, "script.ts");
writeFileSync(alvo, fonte + `
export { main, interpretar, tabelasHTML, seguro, linkSeguro, dataBR, dinheiro,
  numerosNF, assuntoSemPrefixos, empresaExplicita, EMPRESAS };
`);

export const S = await import(pathToFileURL(alvo).href);

// ── Planilha falsa: só o pedaço da API do ExcelScript que o script usa ──
class Faixa {
  constructor(get, set) { this.get = get; this.set = set; }
  getValues() { return this.get(); }
  setValues(v) { this.set(v); }
  setNumberFormat() {}
  setFormula(f) { this.set([[f]]); }
  setValue(v) { this.set([[v]]); }
  getFormat() {
    const nada = { setColor() {}, setBold() {} };
    return { getFill: () => nada, getFont: () => nada };
  }
}

class Tabela {
  constructor(nome, cabecalhos, folha) {
    this.nome = nome; this.cab = cabecalhos; this.dados = []; this.folha = folha;
  }
  getHeaderRowRange() { return new Faixa(() => [this.cab.slice()]); }
  getRowCount() { return this.dados.length; }
  getRangeBetweenHeaderAndTotal() {
    const corpo = new Faixa(() => this.dados.map((r) => r.slice()));
    corpo.getRow = (i) => new Faixa(() => [this.dados[i].slice()], (v) => { this.dados[i] = v[0].slice(); });
    corpo.getColumn = () => new Faixa(() => []);
    return corpo;
  }
  getRange() { const f = new Faixa(() => []); f.getColumn = () => new Faixa(() => []); return f; }
  addRows(indice, linhas) {
    const copia = linhas.map((r) => r.slice());
    if (indice === -1) this.dados.push(...copia); else this.dados.splice(indice, 0, ...copia);
  }
  deleteRowsAt(i, n) { this.dados.splice(i, n); }
  getWorksheet() { return this.folha; }
}

export function planilhaFalsa() {
  const celulas = {};
  const folha = {
    getName: () => "Notas recebidas",
    getRange: (end) => new Faixa(() => [[celulas[end]]], (v) => { celulas[end] = v[0][0]; }),
  };
  const outra = { getName: () => "Outra", getRange: () => new Faixa(() => [[""]], () => {}) };
  const t = {
    tbNotas: new Tabela("tbNotas", ["NF", "Empresa", "Fornecedor", "FOR", "AF / Pedido", "Vencimento", "Comunicação", "Recebido em", "Conferencia", "Possivel_Vinculo", "Observacoes", "Valor_NF_Informado", "Valor_AF_Informado", "Diferenca_Informada", "Liquido_Informado", "ID_Email", "ID_Registro", "Chave_Provisoria"], folha),
    tbEmails: new Tabela("tbEmails", ["ID_Email", "Message_ID", "Internet_Message_ID", "Recebido_UTC_3", "Recebido_UTC_Original", "Remetente", "Assunto", "Referencias", "Pendencias", "Status_Captura", "Tem_Anexos", "Link_Original", "Trecho_Corpo"], outra),
    tbRevisao: new Tabela("tbRevisao", ["ID_Revisao", "ID_Email", "NF", "Tipo", "Detalhe", "Situacao"], outra),
  };
  return { workbook: { getTable: (n) => t[n] }, tabelas: t, celulas };
}
