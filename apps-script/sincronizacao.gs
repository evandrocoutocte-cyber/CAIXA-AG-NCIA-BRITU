// ============================================================
// BRITU / AGÊNCIA DE MARKETING
// Google Sheets -> Supabase
// CORRIGIDO: ignora TOTAL e detecta automaticamente o ano em C1
// ============================================================

const MESES = {
  JANEIRO: 1,
  FEVEREIRO: 2,
  "MARÇO": 3,
  ABRIL: 4,
  MAIO: 5,
  JUNHO: 6,
  JULHO: 7,
  AGOSTO: 8,
  SETEMBRO: 9,
  OUTUBRO: 10,
  NOVEMBRO: 11,
  DEZEMBRO: 12
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Supabase")
    .addItem("Sincronizar tudo", "sincronizarTudo")
    .addItem("Sincronizar receitas", "sincronizarReceitas")
    .addSeparator()
    .addItem("Criar atualização automática (5 min)", "criarAcionador5Minutos")
    .addItem("Remover atualização automática", "removerAcionadores")
    .addToUi();
}

function sincronizarTudo() {
  sincronizarReceitas();
  sincronizarClientes();
  sincronizarDespesas();
  sincronizarMetas();
  sincronizarDebitos();
  sincronizarPlanos();

  SpreadsheetApp.getActive().toast(
    "Sincronização concluída.",
    "Supabase",
    5
  );
}

function sincronizarReceitas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const anoPadrao = detectarAnoPlanilha_(ss);
  const registros = [];

  Object.entries(MESES).forEach(([nomeAba, numeroMes]) => {
    const sh = ss.getSheetByName(nomeAba);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 4) return;

    const values = sh.getRange(4, 1, lastRow - 3, 8).getValues();

    values.forEach((row, idx) => {
      let [id, data, banco, empresa, valor, tipo, status, observacoes] = row;

      if (!empresa && !valor && !data) return;
      if (String(data || "").trim().toUpperCase() === "TOTAL") return;
      if (String(empresa || "").trim().toUpperCase() === "TOTAL") return;

      if (!id) {
        id = "REC-" + Utilities.getUuid().toUpperCase();
        sh.getRange(idx + 4, 1).setValue(id);
      }

      registros.push({
        id_registro: String(id),
        data: parseData_(data, anoPadrao),
        mes: numeroMes,
        ano: anoPadrao,
        banco: textoOuNull_(banco),
        empresa: textoOuNull_(empresa),
        valor: parseMoeda_(valor) || 0,
        tipo: textoOuNull_(tipo),
        status: textoOuNull_(status),
        observacoes: textoOuNull_(observacoes),
        origem_aba: nomeAba
      });
    });
  });

  upsert_("agencia_receitas", registros);
}

function sincronizarClientes() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CLIENTES");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 4) return;

  const values = sh.getRange(4, 1, lastRow - 3, 8).getValues();
  const registros = [];

  values.forEach((row, idx) => {
    let [id, empresa, vencimento, valor, plano, status, dataInicio, observacoes] = row;

    if (!empresa && !valor) return;
    if (String(empresa || "").trim().toUpperCase() === "TOTAL") return;

    if (!id) {
      id = "CLI-" + Utilities.getUuid().toUpperCase();
      sh.getRange(idx + 4, 1).setValue(id);
    }

    registros.push({
      id_registro: String(id),
      empresa: textoOuNull_(empresa),
      vencimento: numeroOuNull_(vencimento),
      valor: parseMoeda_(valor) || 0,
      plano: textoOuNull_(plano),
      status: textoOuNull_(status),
      data_inicio: parseData_(dataInicio, new Date().getFullYear()),
      observacoes: textoOuNull_(observacoes)
    });
  });

  upsert_("agencia_clientes", registros);
}

function sincronizarDespesas() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DESPESAS");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 4) return;

  const values = sh.getRange(4, 2, lastRow - 3, 6).getValues();
  const registros = [];

  values.forEach((row, idx) => {
    let [idFixa, descFixa, valorFixa, idOp, descOp, valorOp] = row;
    const sheetRow = idx + 4;

    if (descFixa && String(descFixa).trim().toUpperCase() !== "TOTAL") {
      if (!idFixa) {
        idFixa = "DES-FIX-" + Utilities.getUuid().toUpperCase();
        sh.getRange(sheetRow, 2).setValue(idFixa);
      }

      registros.push({
        id_registro: String(idFixa),
        descricao: textoOuNull_(descFixa),
        tipo: "Fixa",
        valor: parseMoeda_(valorFixa) || 0,
        status: null,
        observacoes: null
      });
    }

    if (descOp && String(descOp).trim().toUpperCase() !== "TOTAL") {
      if (!idOp) {
        idOp = "DES-OP-" + Utilities.getUuid().toUpperCase();
        sh.getRange(sheetRow, 5).setValue(idOp);
      }

      registros.push({
        id_registro: String(idOp),
        descricao: textoOuNull_(descOp),
        tipo: "Custo Operacional",
        valor: parseMoeda_(valorOp) || 0,
        status: null,
        observacoes: null
      });
    }
  });

  upsert_("agencia_despesas", registros);
}

function sincronizarMetas() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("META");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 4) return;

  const values = sh.getRange(4, 1, lastRow - 3, 5).getValues();
  const registros = [];

  values.forEach((row, idx) => {
    let [id, , empresa, cumprido, valor] = row;

    if (!empresa && !valor) return;
    if (String(empresa || "").trim().toUpperCase() === "TOTAL") return;

    if (!id) {
      id = "META-" + Utilities.getUuid().toUpperCase();
      sh.getRange(idx + 4, 1).setValue(id);
    }

    registros.push({
      id_registro: String(id),
      empresa: textoOuNull_(empresa),
      cumprido: Boolean(cumprido),
      valor: parseMoeda_(valor) || 0
    });
  });

  upsert_("agencia_metas", registros);
}

function sincronizarDebitos() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DÉBITOS");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 4) return;

  const values = sh.getRange(4, 1, lastRow - 3, 6).getValues();
  const registros = [];

  values.forEach((row, idx) => {
    let [id, empresa, vencimento, valor, status, observacoes] = row;

    if (!empresa && !valor) return;
    if (String(empresa || "").trim().toUpperCase() === "TOTAL") return;

    if (!id) {
      id = "DEB-" + Utilities.getUuid().toUpperCase();
      sh.getRange(idx + 4, 1).setValue(id);
    }

    registros.push({
      id_registro: String(id),
      empresa: textoOuNull_(empresa),
      vencimento: parseData_(vencimento, new Date().getFullYear()),
      valor: parseMoeda_(valor) || 0,
      status: textoOuNull_(status),
      observacoes: textoOuNull_(observacoes)
    });
  });

  upsert_("agencia_debitos", registros);
}

function sincronizarPlanos() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANOS");
  if (!sh) return;

  const registros = [];
  const blocos = [
    { nome: "PRESENÇA", colNome: 1, colValor: 2 },
    { nome: "AUTORIDADE", colNome: 4, colValor: 5 },
    { nome: "DOMINANÇA", colNome: 7, colValor: 8 }
  ];

  blocos.forEach(bloco => {
    for (let row = 2; row <= 8; row++) {
      const descricao = sh.getRange(row, bloco.colNome).getValue();
      const valor = sh.getRange(row, bloco.colValor).getValue();

      if (!descricao) continue;
      if (String(descricao).trim().toUpperCase() === "TOTAL") continue;

      registros.push({
        id_registro: slugId_(`PLANO-${bloco.nome}-SERVICO-${row}`),
        plano: bloco.nome,
        categoria: "SERVICO",
        descricao: String(descricao).trim(),
        valor: parseMoeda_(valor),
        ordem: row
      });
    }

    for (let row = 16; row <= 20; row++) {
      const descricao = sh.getRange(row, bloco.colNome).getValue();
      const valor = sh.getRange(row, bloco.colValor).getValue();

      if (!descricao) continue;
      if (String(descricao).trim().toUpperCase() === "TOTAL") continue;

      registros.push({
        id_registro: slugId_(`PLANO-${bloco.nome}-CUSTO-${row}`),
        plano: bloco.nome,
        categoria: "CUSTO",
        descricao: String(descricao).trim(),
        valor: parseMoeda_(valor),
        ordem: row
      });
    }
  });

  upsert_("agencia_planos", registros);
}

function upsert_(table, registros) {
  if (!registros || registros.length === 0) return;

  const baseUrl = getConfig_("SUPABASE_URL");
  const secretKey = getConfig_("SUPABASE_SECRET_KEY");

  if (!baseUrl || !secretKey) {
    throw new Error(
      "Configure SUPABASE_URL e SUPABASE_SECRET_KEY nas Propriedades do script."
    );
  }

  const endpoint =
    baseUrl.replace(/\/$/, "") +
    "/rest/v1/" +
    table +
    "?on_conflict=id_registro";

  const response = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    headers: {
      apikey: secretKey,
      Prefer: "resolution=merge-duplicates,return=minimal"
    },
    payload: JSON.stringify(registros),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code < 200 || code >= 300) {
    throw new Error(
      `Erro ao sincronizar ${table}. HTTP ${code}: ${response.getContentText()}`
    );
  }
}

function criarAcionador5Minutos() {
  removerAcionadores();

  ScriptApp.newTrigger("sincronizarTudo")
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getActive().toast(
    "Atualização automática criada: a cada 5 minutos.",
    "Supabase",
    5
  );
}

function removerAcionadores() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "sincronizarTudo") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function detectarAnoPlanilha_(ss) {
  const anosEncontrados = [];

  Object.keys(MESES).forEach(nomeAba => {
    const sh = ss.getSheetByName(nomeAba);
    if (!sh) return;

    const valor = Number(sh.getRange("C1").getValue());

    if (Number.isInteger(valor) && valor >= 2025 && valor <= 2050) {
      anosEncontrados.push(valor);
    }
  });

  const anosUnicos = [...new Set(anosEncontrados)];

  if (anosUnicos.length === 1) {
    return anosUnicos[0];
  }

  if (anosUnicos.length > 1) {
    throw new Error(
      "Existem anos diferentes no topo das abas mensais: " +
      anosUnicos.join(", ") +
      ". Deixe o mesmo ano em C1 de todas as abas."
    );
  }

  const nomeArquivo = ss.getName();
  const match = nomeArquivo.match(/\b(20(?:2[5-9]|[3-4][0-9]|50))\b/);

  if (match) {
    return Number(match[1]);
  }

  throw new Error(
    'Não foi possível identificar o ano da planilha. ' +
    'Informe o ano em C1 das abas mensais (ex.: 2026).'
  );
}

function getConfig_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function textoOuNull_(value) {
  if (value === null || value === undefined || value === "") return null;
  const t = String(value).trim();
  return t ? t : null;
}

function numeroOuNull_(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function parseMoeda_(value) {
  if (value === null || value === undefined || value === "") return null;

  if (typeof value === "number") {
    return Number(value.toFixed(2));
  }

  const cleaned = String(value)
    .replace(/\s/g, "")
    .replace("R$", "")
    .replace(/\./g, "")
    .replace(",", ".");

  const n = Number(cleaned);
  return Number.isFinite(n) ? Number(n.toFixed(2)) : null;
}

function parseData_(value, anoPadrao) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(
      value,
      "America/Sao_Paulo",
      "yyyy-MM-dd"
    );
  }

  const t = String(value).trim();
  const parts = t.split("/");

  if (parts.length < 2) return null;

  const dia = Number(parts[0]);
  const mes = Number(parts[1]);
  let ano = parts.length >= 3 && parts[2] ? Number(parts[2]) : Number(anoPadrao);

  if (ano < 100) ano += 2000;
  if (!dia || !mes || !ano) return null;

  const d = new Date(ano, mes - 1, dia);

  return Utilities.formatDate(
    d,
    "America/Sao_Paulo",
    "yyyy-MM-dd"
  );
}

function slugId_(text) {
  return String(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toUpperCase();
}
