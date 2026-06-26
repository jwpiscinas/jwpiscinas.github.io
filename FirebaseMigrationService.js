var JW_FIRESTORE_MIGRATION_SHEETS = [
  { sheetName: "Usuarios", collection: "users", idFields: ["Email", "Usuario"], fallbackPrefix: "USR" },
  { sheetName: "TelefonesAutorizados", collection: "authorized_phones", idFields: ["Telefone"], fallbackPrefix: "PHONE" },
  { sheetName: "Produtos", collection: "products", idFields: ["ID", "Nome"], fallbackPrefix: "PROD" },
  { sheetName: "Manuais", collection: "manuals", idFields: ["ID", "Produto"], fallbackPrefix: "MAN" },
  { sheetName: "Compras", collection: "purchases", idFields: ["ID"], fallbackPrefix: "COMP" },
  { sheetName: "Pagamentos", collection: "payments", idFields: ["ID"], fallbackPrefix: "PAG" },
  { sheetName: "Orcamentos", collection: "quotes", idFields: ["ID"], fallbackPrefix: "ORC" },
  { sheetName: "Servicos", collection: "services", idFields: ["ID"], fallbackPrefix: "SERV" },
  { sheetName: "Notificacoes", collection: "notifications", idFields: ["ID"], fallbackPrefix: "NOT" },
  { sheetName: "Tecnicos", collection: "technicians", idFields: ["ID"], fallbackPrefix: "TEC" },
  { sheetName: "OrdensServico", collection: "work_orders", idFields: ["ID"], fallbackPrefix: "OS" },
  { sheetName: "Agenda", collection: "schedule", idFields: ["ID"], fallbackPrefix: "AGE" },
  { sheetName: "HistoricoGeral", collection: "history", idFields: ["ID"], fallbackPrefix: "HIST" },
  { sheetName: "LogsSistema", collection: "system_logs", idFields: ["DataHora", "Usuario", "Acao"], fallbackPrefix: "LOG" },
  { sheetName: "MensagensChat", collection: "chat_messages", idFields: ["ID", "ClienteEmail", "Email"], fallbackPrefix: "CHAT" },
  { sheetName: "ItensCobranca", collection: "billing_items", idFields: ["ID", "Nome"], fallbackPrefix: "ITEM" },
  { sheetName: "LancamentosFinanceiros", collection: "financial_entries", idFields: ["ID"], fallbackPrefix: "LCT" },
  { sheetName: "AssinaturasMensais", collection: "monthly_subscriptions", idFields: ["ID"], fallbackPrefix: "ASN" },
  { sheetName: "CobrancasClientes", collection: "customer_charges", idFields: ["ID"], fallbackPrefix: "COB" }
];

var JW_LEGACY_CHAT_REALTIME_DATABASE_URL = "https://jwpiscinas-default-rtdb.firebaseio.com";

function listarAbasJwParaMigracaoFirebase() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var summary = [];

  for (var i = 0; i < JW_FIRESTORE_MIGRATION_SHEETS.length; i++) {
    var config = JW_FIRESTORE_MIGRATION_SHEETS[i];
    var sheet = spreadsheet.getSheetByName(config.sheetName);
    var totalRows = sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;

    summary.push({
      sheetName: config.sheetName,
      collection: getFirebaseCollectionName_(config.collection),
      exists: !!sheet,
      totalRows: totalRows
    });
  }

  return {
    success: true,
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    sheets: summary
  };
}

function simularMigracaoBaseJwParaFirestore(opcoes) {
  return executarMigracaoBaseJwParaFirestore_(opcoes || {}, true);
}

function migrarBaseJwParaFirestore(opcoes) {
  return executarMigracaoBaseJwParaFirestore_(opcoes || {}, false);
}

function migrarAbaJwParaFirestore(sheetName, opcoes) {
  opcoes = opcoes || {};
  opcoes.sheets = [sheetName];
  return executarMigracaoBaseJwParaFirestore_(opcoes, false);
}

function executarMigracaoBaseJwParaFirestore_(opcoes, dryRun) {
  var validation = validarConfiguracaoFirebaseMigracao();
  if (!validation.success) return validation;

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    return {
      success: false,
      message: "Nao foi possivel acessar a planilha ativa do sistema JW."
    };
  }

  var selected = getSelectedMigrationSheets_(opcoes.sheets);
  var runId = "MIG-" + Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyyMMdd-HHmmss");
  var summary = {
    success: true,
    dryRun: dryRun === true,
    runId: runId,
    projectId: getFirebaseMigrationConfig_().projectId,
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    startedAt: new Date().toISOString(),
    totalDocumentsPrepared: 0,
    totalDocumentsWritten: 0,
    sheets: []
  };

  for (var i = 0; i < selected.length; i++) {
    var sheetResult = migrateSingleSheetToFirestore_(spreadsheet, selected[i], dryRun, opcoes);
    summary.sheets.push(sheetResult);
    summary.totalDocumentsPrepared += sheetResult.documentsPrepared || 0;
    summary.totalDocumentsWritten += sheetResult.documentsWritten || 0;

    if (sheetResult.success === false) {
      summary.success = false;
    }
  }

  summary.finishedAt = new Date().toISOString();

  if (!dryRun && summary.success) {
    firestoreSetDocument_("migration_runs", runId, {
      run_id: runId,
      source: "jw_piscinas",
      spreadsheet_id: summary.spreadsheetId,
      spreadsheet_name: summary.spreadsheetName,
      started_at: summary.startedAt,
      finished_at: summary.finishedAt,
      total_documents_prepared: summary.totalDocumentsPrepared,
      total_documents_written: summary.totalDocumentsWritten,
      sheets: summary.sheets
    });
  }

  summary.message = dryRun
    ? "Simulacao de migracao concluida."
    : (summary.success
      ? "Migracao da base JW para Firestore concluida."
      : "Migracao executada com falhas em uma ou mais abas.");

  return summary;
}

function migrateSingleSheetToFirestore_(spreadsheet, config, dryRun, options) {
  var sheet = spreadsheet.getSheetByName(config.sheetName);
  if (!sheet) {
    return {
      success: true,
      skipped: true,
      sheetName: config.sheetName,
      collection: getFirebaseCollectionName_(config.collection),
      documentsPrepared: 0,
      documentsWritten: 0,
      message: "Aba nao encontrada; pulando migracao."
    };
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return {
      success: true,
      skipped: false,
      sheetName: config.sheetName,
      collection: getFirebaseCollectionName_(config.collection),
      documentsPrepared: 0,
      documentsWritten: 0,
      message: "Aba sem registros para migrar."
    };
  }

  var headers = values[0];
  var maxRows = getMigrationMaxRows_(options);
  var writes = [];
  var sampleIds = [];
  var skippedRows = 0;
  var usedDocumentIds = {};

  for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
    if (maxRows && writes.length >= maxRows) break;

    var row = values[rowIndex];
    if (isEmptySheetRow_(row)) {
      skippedRows++;
      continue;
    }

    var record = convertSheetRowToObject_(headers, row);
    var documentId = ensureUniqueMigrationDocumentId_(
      buildMigrationDocumentId_(config, record, rowIndex + 1),
      usedDocumentIds,
      rowIndex + 1
    );
    var payload = buildMigrationDocumentPayload_(spreadsheet, config.sheetName, headers, row, rowIndex + 1, record, documentId);

    writes.push(buildFirestoreSetWrite_(config.collection, documentId, payload));

    if (sampleIds.length < 5) {
      sampleIds.push(documentId);
    }
  }

  var prepared = writes.length;
  if (!prepared) {
    return {
      success: true,
      skipped: false,
      sheetName: config.sheetName,
      collection: getFirebaseCollectionName_(config.collection),
      documentsPrepared: 0,
      documentsWritten: 0,
      skippedRows: skippedRows,
      message: "Nenhum documento util encontrado na aba."
    };
  }

  var written = 0;
  if (!dryRun) {
    var batches = splitFirestoreWrites_(writes, 200);
    for (var i = 0; i < batches.length; i++) {
      var response = firestoreCommitWrites_(batches[i]);
      written += (response.writeResults || []).length;
    }
  }

  return {
    success: true,
    skipped: false,
    sheetName: config.sheetName,
    collection: getFirebaseCollectionName_(config.collection),
    totalRows: values.length - 1,
    skippedRows: skippedRows,
    documentsPrepared: prepared,
    documentsWritten: dryRun ? 0 : written,
    sampleDocumentIds: sampleIds,
    message: dryRun
      ? "Simulacao pronta para a aba " + config.sheetName + "."
      : "Aba " + config.sheetName + " migrada com sucesso."
  };
}

function buildMigrationDocumentPayload_(spreadsheet, sheetName, headers, row, rowNumber, record, documentId) {
  return {
    document_id: documentId,
    source_sheet: sheetName,
    source_row_number: rowNumber,
    source_spreadsheet_id: spreadsheet.getId(),
    source_spreadsheet_name: spreadsheet.getName(),
    migrated_at: new Date().toISOString(),
    source_headers: normalizeSheetHeaderArray_(headers),
    source_values: normalizeSheetValueArray_(row),
    data: record
  };
}

function buildMigrationDocumentId_(config, record, rowNumber) {
  var candidates = config.idFields || [];

  for (var i = 0; i < candidates.length; i++) {
    var normalizedKey = normalizeSheetFieldName_(candidates[i]);
    var value = record[normalizedKey];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return sanitizeFirestoreDocumentId_(value, config.fallbackPrefix + "_" + rowNumber);
    }
  }

  return sanitizeFirestoreDocumentId_(config.fallbackPrefix + "_" + rowNumber, config.fallbackPrefix + "_" + rowNumber);
}

function convertSheetRowToObject_(headers, row) {
  var data = {};

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var key = normalizeSheetFieldName_(header);
    var value = normalizeSheetCellValue_(row[i]);

    if (!key) key = "campo_" + (i + 1);

    if (data.hasOwnProperty(key)) {
      key = key + "_" + (i + 1);
    }

    data[key] = value;
  }

  return data;
}

function normalizeSheetHeaderArray_(headers) {
  var result = [];

  for (var i = 0; i < headers.length; i++) {
    result.push(headers[i] === undefined || headers[i] === null ? "" : String(headers[i]));
  }

  return result;
}

function normalizeSheetValueArray_(row) {
  var result = [];

  for (var i = 0; i < row.length; i++) {
    result.push(normalizeSheetCellValue_(row[i]));
  }

  return result;
}

function normalizeSheetCellValue_(value) {
  if (value === undefined || value === null) return null;
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (Array.isArray(value)) return value;
  return String(value);
}

function normalizeSheetFieldName_(value) {
  var text = value === undefined || value === null ? "" : String(value).trim();
  text = removeAccentsFirestore_(text).toLowerCase();
  text = text.replace(/[^a-z0-9]+/g, "_");
  text = text.replace(/^_+|_+$/g, "");
  return text;
}

function isEmptySheetRow_(row) {
  for (var i = 0; i < row.length; i++) {
    var value = row[i];
    if (value === null || value === undefined) continue;
    if (typeof value === "string" && value.trim() === "") continue;
    if (typeof value === "boolean" && value === false) continue;
    return false;
  }
  return true;
}

function ensureUniqueMigrationDocumentId_(documentId, usedDocumentIds, rowNumber) {
  var baseId = sanitizeFirestoreDocumentId_(documentId, "ROW_" + rowNumber);

  if (!usedDocumentIds[baseId]) {
    usedDocumentIds[baseId] = 1;
    return baseId;
  }

  usedDocumentIds[baseId]++;
  return sanitizeFirestoreDocumentId_(baseId + "_row_" + rowNumber, "ROW_" + rowNumber);
}

function splitFirestoreWrites_(writes, size) {
  var groups = [];
  var chunkSize = size || 200;

  for (var i = 0; i < writes.length; i += chunkSize) {
    groups.push(writes.slice(i, i + chunkSize));
  }

  return groups;
}

function getSelectedMigrationSheets_(requestedSheets) {
  if (!requestedSheets || !requestedSheets.length) {
    return JW_FIRESTORE_MIGRATION_SHEETS.slice();
  }

  var lookup = {};
  for (var i = 0; i < requestedSheets.length; i++) {
    lookup[String(requestedSheets[i]).toLowerCase()] = true;
  }

  var result = [];
  for (var j = 0; j < JW_FIRESTORE_MIGRATION_SHEETS.length; j++) {
    var config = JW_FIRESTORE_MIGRATION_SHEETS[j];
    if (lookup[config.sheetName.toLowerCase()]) {
      result.push(config);
    }
  }

  return result;
}

function getMigrationMaxRows_(options) {
  if (!options) return 0;
  if (!options.maxRows) return 0;

  var parsed = parseInt(options.maxRows, 10);
  return isNaN(parsed) || parsed <= 0 ? 0 : parsed;
}

function migrarChatRealtimeParaFirestore(opcoes) {
  opcoes = opcoes || {};

  try {
    if (typeof shouldUseFirebaseAppData_ !== "function" || !shouldUseFirebaseAppData_()) {
      return {
        success: false,
        message: "Firestore ainda nao esta habilitado para a base principal."
      };
    }

    var mensagensLegadas = buscarChatRealtimeLegado_("messages");
    if (!mensagensLegadas || typeof mensagensLegadas !== "object") {
      return {
        success: true,
        imported: 0,
        skipped: 0,
        message: "Nenhuma mensagem legada encontrada no Realtime Database."
      };
    }

    var config = JW_FIRESTORE_APP_DATA_CONFIG.MensagensChat;
    var headers = config.headers.slice();
    var existentes = {};
    var docsAtuais = listFirestoreCollectionRecords_(config.collection);
    for (var i = 0; i < docsAtuais.length; i++) {
      var docIdAtual = extractFirestoreDocumentId_(docsAtuais[i] && docsAtuais[i].name);
      if (docIdAtual) existentes[docIdAtual] = true;
    }

    var writes = [];
    var imported = 0;
    var skipped = 0;

    for (var conversaId in mensagensLegadas) {
      if (!mensagensLegadas.hasOwnProperty(conversaId)) continue;
      var mensagensConversa = mensagensLegadas[conversaId] || {};

      for (var mensagemId in mensagensConversa) {
        if (!mensagensConversa.hasOwnProperty(mensagemId)) continue;
        if (existentes[mensagemId]) {
          skipped++;
          continue;
        }

        var item = mensagensConversa[mensagemId] || {};
        var remetenteTipo = valorTextoMigracaoChat_(item.remetenteTipo) || "Cliente";
        var clienteEmail = valorTextoMigracaoChat_(item.clienteEmail).toLowerCase();
        var remetenteEmail = valorTextoMigracaoChat_(item.remetenteEmail).toLowerCase();
        var timestamp = parseInt(item.timestamp, 10);
        if (isNaN(timestamp)) timestamp = Date.now();
        var dataHora = valorTextoMigracaoChat_(item.dataHora) || new Date(timestamp).toLocaleString("pt-BR");
        var tipoMensagem = valorTextoMigracaoChat_(item.tipo).toLowerCase() || "texto";
        var row = [
          mensagemId,
          dataHora,
          clienteEmail,
          valorTextoMigracaoChat_(item.clienteNome),
          remetenteEmail,
          valorTextoMigracaoChat_(item.remetenteNome),
          remetenteTipo,
          valorTextoMigracaoChat_(item.mensagem),
          "Sim",
          "Sim",
          tipoMensagem,
          valorTextoMigracaoChat_(item.audioDataUrl),
          valorTextoMigracaoChat_(item.imagemDataUrl),
          valorTextoMigracaoChat_(item.mimeType),
          valorTextoMigracaoChat_(item.duracao),
          timestamp
        ];
        var payload = buildAdapterDocumentPayload_(config, headers, row, imported + 2, mensagemId, null);
        writes.push(buildFirestoreSetWrite_(config.collection, mensagemId, payload));
        existentes[mensagemId] = true;
        imported++;
      }
    }

    if (!writes.length) {
      return {
        success: true,
        imported: 0,
        skipped: skipped,
        message: "Nenhuma nova mensagem do chat precisava ser migrada."
      };
    }

    var batches = splitFirestoreWrites_(writes, 200);
    var escritos = 0;
    for (var j = 0; j < batches.length; j++) {
      var response = firestoreCommitWrites_(batches[j]);
      escritos += (response.writeResults || []).length;
    }

    if (JW_FIRESTORE_SHEET_STATE_CACHE.MensagensChat) {
      JW_FIRESTORE_SHEET_STATE_CACHE.MensagensChat.loaded = false;
      JW_FIRESTORE_SHEET_STATE_CACHE.MensagensChat.records = [];
    }

    return {
      success: true,
      imported: imported,
      written: escritos,
      skipped: skipped,
      message: "Chat legado migrado do Realtime Database para o Firestore."
    };
  } catch (error) {
    return {
      success: false,
      message: "Erro ao migrar chat legado: " + error.toString()
    };
  }
}

function buscarChatRealtimeLegado_(path) {
  var url = JW_LEGACY_CHAT_REALTIME_DATABASE_URL.replace(/\/+$/, "") + "/jwChat/" + String(path || "").replace(/^\/+/, "") + ".json";
  var response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true
  });
  var statusCode = response.getResponseCode();
  if (statusCode === 404) return null;
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Realtime Database HTTP " + statusCode + ": " + response.getContentText());
  }
  var body = response.getContentText() || "null";
  return JSON.parse(body);
}

function valorTextoMigracaoChat_(valor) {
  if (valor === undefined || valor === null) return "";
  return String(valor);
}
