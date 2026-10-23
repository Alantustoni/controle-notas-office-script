# Montando o fluxo no Power Automate

O Office Script não roda sozinho: um fluxo do Power Automate o chama a cada
e-mail que chega. Este guia descreve o fluxo mínimo.

> Os nomes dos campos abaixo são os do conector em português. Eles podem
> variar um pouco com o idioma da conta ou a versão do conector.

## 1. Preparar a planilha e o script

1. Suba `modelo/Controle_Notas_modelo.xlsx` para o OneDrive ou SharePoint.
2. Abra a planilha no Excel para a Web → **Automatizar** → **Novo script**.
3. Cole o conteúdo de `src/registrar_email_compras.ts` e salve como
   **Registrar_Email_Compras**. O Excel guarda o script em
   `OneDrive/Documentos/Scripts do Office/`.
4. No topo do script, ajuste a configuração:
   - `EMPRESAS`: os nomes das empresas destinatárias, em MAIÚSCULAS, como
     aparecem nos e-mails;
   - `CONTADORES_PAINEL`: as células do painel com o contador de cada
     empresa (mesma ordem de `EMPRESAS`);
   - `FUSO_HORAS`: o fuso local em relação ao UTC (Brasília = -3).
5. Se mudar os nomes das empresas, troque também os rótulos do painel
   (células B5 e F5).

## 2. Criar o fluxo

**Gatilho:** Office 365 Outlook → **Quando um novo email chegar (V3)**

- Pasta: a caixa onde chegam os avisos de Compras.
- Recomendado: filtrar por remetente ou assunto, para o script não ser
  chamado por e-mails de outros assuntos. O script não quebra com e-mails
  sem notas (registra o e-mail e abre uma pendência "Formato não
  reconhecido"), mas o filtro evita ruído.

**Ação:** Excel Online (Business) → **Executar script**

- Local / Biblioteca / Arquivo: a planilha do passo 1.
- Script: **Registrar_Email_Compras**.
- Parâmetros (conteúdo dinâmico do gatilho):

| Parâmetro do script | Campo do gatilho | Observação |
|---|---|---|
| `idMensagem` | ID da Mensagem | **Não** use o ID da Conversa: várias mensagens compartilham o mesmo |
| `recebidoEmUTC` | Hora de Recebimento | Precisa vir com fuso (ex.: `2026-11-03T13:05:00Z`); sem fuso o script recusa |
| `remetente` | De | Vazio vira pendência |
| `assunto` | Assunto | Usado para detectar ENC/RES e a NF citada no assunto |
| `corpoHtml` | Corpo | O corpo em **HTML**, não a prévia em texto |
| `temAnexos` | Tem Anexos | Só informativo: o script não lê anexos |
| `internetMessageId` | ID da Mensagem da Internet | Opcional; quando existe, é o identificador preferido |
| `linkOriginal` | (opcional) | Só é gravado se for um link `outlook.office.com` ou `outlook.office365.com` |

## 3. Testar

1. Peça para alguém mandar um e-mail de teste com uma tabela de notas.
2. Acompanhe em **Meus fluxos → Histórico de execuções**. A execução mostra
   o retorno do script, por exemplo
   `{"status": "Concluído", "novasReferencias": 2, ...}`.
3. Confira a planilha. As novas notas entram no topo de `tbNotas`.

**Atraso:** o gatilho do Outlook verifica a caixa periodicamente, então a
linha pode levar alguns minutos para aparecer. No uso real, um caso passou de
15 minutos.

**Reprocessar:** usar **Reenviar** numa execução antiga é seguro. O script
reconhece o ID do e-mail e responde `"Já registrado"` sem duplicar nada.

## 4. O que o script devolve

```json
{
  "status": "Concluído | Com revisão | Já registrado",
  "idEmail": "...",
  "referencias": 3,
  "novasReferencias": 3,
  "pendencias": 0,
  "mensagem": "..."
}
```

Use `status` numa condição do fluxo se quiser, por exemplo, mandar um aviso
no Teams quando vier `"Com revisão"`.
