var JW_FIRESTORE_APP_DATA_CONFIG = {
  Usuarios: {
    collection: "users",
    idFields: ["Email", "Usuario"],
    fallbackPrefix: "USR",
    headers: ["Email", "Senha", "Usuario", "Telefone", "Tipo", "DataCadastro", "Status", "CodigoRecuperacao", "DataCodigo", "AvatarURL", "Admin", "Endereco", "Bairro"]
  },
  TelefonesAutorizados: {
    collection: "authorized_phones",
    idFields: ["Telefone"],
    fallbackPrefix: "PHONE",
    headers: ["Telefone", "Nome", "DataAutorizacao", "Status", "Usado", "DataUso"]
  },
  Produtos: {
    collection: "products",
    idFields: ["ID", "Nome"],
    fallbackPrefix: "PROD",
    headers: ["ID", "Nome", "Categoria", "Descricao", "Preco", "Estoque", "ImagemURL", "ManualURL", "ManualNome", "PrecoComDesconto", "ExibirManual", "ManualVideoURL"]
  },
  Compras: {
    collection: "purchases",
    idFields: ["ID"],
    fallbackPrefix: "COMP",
    headers: ["ID", "Cliente", "Email", "DataCompra", "Produtos", "QuantidadeTotal", "ValorTotal", "StatusPagamento", "StatusEntrega", "ObservacoesAdmin"]
  },
  Pagamentos: {
    collection: "payments",
    idFields: ["ID"],
    fallbackPrefix: "PAG",
    headers: ["ID", "Cliente", "Email", "Tipo", "Valor", "Vencimento", "Status", "DataPagamento", "Metodo", "Observacoes"]
  },
  Orcamentos: {
    collection: "quotes",
    idFields: ["ID"],
    fallbackPrefix: "ORC",
    headers: ["ID", "Cliente", "Email", "Telefone", "DataSolicitacao", "TipoServico", "Descricao", "ImagemURL", "ArquivoURL", "Valor", "Status", "DataValidade", "Observacoes", "ArquivoNome"]
  },
  Servicos: {
    collection: "services",
    idFields: ["ID"],
    fallbackPrefix: "SERV",
    headers: ["ID", "Cliente", "Email", "TipoServico", "DataInicio", "DataConclusao", "Valor", "Status", "Descricao", "OrcamentoID"]
  },
  Notificacoes: {
    collection: "notifications",
    idFields: ["ID"],
    fallbackPrefix: "NOT",
    headers: ["ID", "Titulo", "Mensagem", "Tipo", "DataEnvio", "Destinatarios", "Lida", "CriadoPor", "Acao", "Icone"]
  },
  Tecnicos: {
    collection: "technicians",
    idFields: ["ID"],
    fallbackPrefix: "TEC",
    headers: ["ID", "Nome", "Telefone", "Especialidade", "Status", "DataCadastro"]
  },
  OrdensServico: {
    collection: "work_orders",
    idFields: ["ID"],
    fallbackPrefix: "OS",
    headers: ["ID", "Cliente", "TecnicoID", "TipoServico", "DataAbertura", "DataAgendamento", "DataConclusao", "Valor", "Status", "Descricao", "OrcamentoID"]
  },
  Agenda: {
    collection: "schedule",
    idFields: ["ID"],
    fallbackPrefix: "AGE",
    headers: ["ID", "Data", "Hora", "TecnicoID", "Cliente", "Servico", "Status"]
  },
  HistoricoGeral: {
    collection: "history",
    idFields: ["ID"],
    fallbackPrefix: "HIST",
    headers: ["ID", "Tipo", "Cliente", "Email", "Telefone", "Titulo", "Descricao", "Valor", "Status", "DataRegistro", "LinkAcao", "Observacoes"]
  },
  LogsSistema: {
    collection: "system_logs",
    idFields: ["DataHora", "Usuario", "Acao"],
    fallbackPrefix: "LOG",
    headers: ["DataHora", "Tipo", "Usuario", "Acao", "Detalhes"]
  },
  MensagensChat: {
    collection: "chat_messages",
    idFields: ["ID", "ClienteEmail", "RemetenteEmail"],
    fallbackPrefix: "CHAT",
    headers: ["ID", "DataHora", "ClienteEmail", "ClienteNome", "RemetenteEmail", "RemetenteNome", "RemetenteTipo", "Mensagem", "LidaCliente", "LidaAdmin", "Tipo", "AudioDataUrl", "ImagemDataUrl", "MimeType", "Duracao", "Timestamp"]
  },
  Manuais: {
    collection: "manuals",
    idFields: ["ID", "Produto"],
    fallbackPrefix: "MAN",
    headers: ["ID", "Produto", "Categoria", "Tipo", "URL", "Descricao", "Icone"]
  },
  ItensCobranca: {
    collection: "billing_items",
    idFields: ["ID", "Nome"],
    fallbackPrefix: "ITEM",
    headers: ["ID", "Nome", "Categoria", "Unidade", "ValorPadrao", "TipoUso", "Ativo", "Descricao", "DataCadastro", "AtualizadoEm", "EstoqueInicial", "ImagemURL", "ProdutoID"]
  },
  LancamentosFinanceiros: {
    collection: "financial_entries",
    idFields: ["ID"],
    fallbackPrefix: "LCT",
    headers: ["ID", "ClienteEmail", "ClienteNome", "Telefone", "TipoLancamento", "ItemID", "ItemNome", "Descricao", "Quantidade", "Unidade", "ValorUnitario", "ValorTotal", "Competencia", "Referencia", "Status", "CobrancaID", "DataLancamento", "Observacoes", "CriadoPor", "AssinaturaID"]
  },
  AssinaturasMensais: {
    collection: "monthly_subscriptions",
    idFields: ["ID"],
    fallbackPrefix: "ASN",
    headers: ["ID", "ClienteEmail", "ClienteNome", "Telefone", "Descricao", "Valor", "DiaVencimento", "Status", "UltimaCompetenciaGerada", "Observacoes", "DataCadastro"]
  },
  CobrancasClientes: {
    collection: "customer_charges",
    idFields: ["ID"],
    fallbackPrefix: "COB",
    headers: ["ID", "ClienteEmail", "ClienteNome", "Telefone", "Competencia", "Referencia", "ItensJson", "QuantidadeItens", "ValorTotal", "Vencimento", "Status", "LinkWhatsApp", "MensagemWhatsApp", "DataGeracao", "DataEnvio", "Observacoes", "CriadoPor", "PdfUrl", "PdfFileId", "PdfGeradoEm"]
  }
};

var JW_FIRESTORE_APP_DATA_ENABLED_CACHE = null;
var JW_FIRESTORE_SHEET_STATE_CACHE = {};

function shouldUseFirebaseAppData_() {
  if (JW_FIRESTORE_APP_DATA_ENABLED_CACHE !== null) {
    return JW_FIRESTORE_APP_DATA_ENABLED_CACHE;
  }

  try {
    if (typeof validarConfiguracaoFirebaseMigracao !== "function") {
      JW_FIRESTORE_APP_DATA_ENABLED_CACHE = false;
      return false;
    }

    var status = validarConfiguracaoFirebaseMigracao();
    JW_FIRESTORE_APP_DATA_ENABLED_CACHE = !!(status && status.success);
    return JW_FIRESTORE_APP_DATA_ENABLED_CACHE;
  } catch (error) {
    JW_FIRESTORE_APP_DATA_ENABLED_CACHE = false;
    return false;
  }
}

function getFirestoreSheetAdapter_(sheetName, headers, overrides) {
  var baseConfig = JW_FIRESTORE_APP_DATA_CONFIG[sheetName] || {};
  var config = {};
  var key;

  for (key in baseConfig) {
    if (baseConfig.hasOwnProperty(key)) config[key] = baseConfig[key];
  }
  for (key in (overrides || {})) {
    if (overrides.hasOwnProperty(key)) config[key] = overrides[key];
  }

  config.sheetName = sheetName;
  config.headers = (headers && headers.length ? headers : config.headers || []).slice();

  if (!config.collection) {
    throw new Error("Colecao Firestore nao configurada para a aba " + sheetName + ".");
  }

  if (!config.headers.length) {
    throw new Error("Cabecalhos nao configurados para a aba " + sheetName + ".");
  }

  return createFirestoreSheetAdapter_(config);
}

function createFirestoreSheetAdapter_(config) {
  var cacheKey = config.sheetName;
  var state = JW_FIRESTORE_SHEET_STATE_CACHE[cacheKey];
  if (!state) {
    state = {
      config: config,
      headers: config.headers.slice(),
      records: [],
      loaded: false
    };
    JW_FIRESTORE_SHEET_STATE_CACHE[cacheKey] = state;
  } else {
    state.config = config;
    state.headers = config.headers.slice();
  }

  function ensureLoaded(force) {
    if (force) state.loaded = false;
    if (state.loaded) return;

    var docs = listFirestoreCollectionRecords_(state.config.collection);

    state.records = docs.map(function(doc) {
      return buildFirestoreSheetRecord_(state.config, doc);
    });

    state.records.sort(compareFirestoreSheetRecords_);
    state.loaded = true;
  }

  function getMatrix() {
    ensureLoaded(false);
    var values = [state.headers.slice()];
    for (var i = 0; i < state.records.length; i++) {
      values.push(state.records[i].row.slice());
    }
    return values;
  }

  function persistRecordAtBodyIndex(bodyIndex) {
    ensureLoaded(false);
    if (bodyIndex < 0 || bodyIndex >= state.records.length) return;

    var record = state.records[bodyIndex];
    var row = normalizeRowLength_(record.row, state.headers.length);

    if (isEmptySheetRow_(row)) {
      deleteFirestoreCollectionRecord_(state.config.collection, record.docId);
      state.records.splice(bodyIndex, 1);
      return;
    }

    var docId = buildAdapterDocumentId_(state.config, row, bodyIndex + 2, record.docId);
    var payload = buildAdapterDocumentPayload_(state.config, state.headers, row, bodyIndex + 2, docId, record.payload);

    upsertFirestoreCollectionRecord_(state.config.collection, docId, payload);
    record.docId = docId;
    record.payload = payload;
    record.row = row;
    record.sortKey = getFirestoreRecordSortKey_(payload, docId);
  }

  function rangeFactory(row, column, numRows, numColumns) {
    var startRow = Math.max(parseInt(row, 10) || 1, 1);
    var startColumn = Math.max(parseInt(column, 10) || 1, 1);
    var rowCount = Math.max(parseInt(numRows, 10) || 1, 1);
    var columnCount = Math.max(parseInt(numColumns, 10) || 1, 1);

    function sliceMatrix() {
      var matrix = getMatrix();
      var output = [];

      for (var r = 0; r < rowCount; r++) {
        var sourceRow = matrix[startRow - 1 + r] || [];
        var line = [];
        for (var c = 0; c < columnCount; c++) {
          line.push(sourceRow[startColumn - 1 + c]);
        }
        output.push(line);
      }

      return output;
    }

    function applyValues(values) {
      ensureLoaded(false);
      var touched = {};

      for (var r = 0; r < rowCount; r++) {
        var absoluteRow = startRow + r;
        if (absoluteRow === 1) continue;

        var bodyIndex = absoluteRow - 2;
        while (bodyIndex >= state.records.length) {
          state.records.push({
            docId: "",
            payload: null,
            row: createBlankRow_(state.headers.length),
            sortKey: String(new Date().getTime()) + "_" + state.records.length
          });
        }

        var record = state.records[bodyIndex];
        for (var c = 0; c < columnCount; c++) {
          record.row[startColumn - 1 + c] = values[r][c];
        }
        touched[bodyIndex] = true;
      }

      var indexes = Object.keys(touched).map(function(item) { return parseInt(item, 10); });
      indexes.sort(function(a, b) { return a - b; });

      for (var i = 0; i < indexes.length; i++) {
        persistRecordAtBodyIndex(indexes[i]);
      }

      state.records.sort(compareFirestoreSheetRecords_);
    }

    function clearCells() {
      ensureLoaded(false);
      var touched = {};

      for (var r = 0; r < rowCount; r++) {
        var absoluteRow = startRow + r;
        if (absoluteRow === 1) continue;
        var bodyIndex = absoluteRow - 2;
        if (bodyIndex < 0 || bodyIndex >= state.records.length) continue;

        var record = state.records[bodyIndex];
        for (var c = 0; c < columnCount; c++) {
          record.row[startColumn - 1 + c] = "";
        }
        touched[bodyIndex] = true;
      }

      var indexes = Object.keys(touched).map(function(item) { return parseInt(item, 10); });
      indexes.sort(function(a, b) { return b - a; });

      for (var i = 0; i < indexes.length; i++) {
        persistRecordAtBodyIndex(indexes[i]);
      }

      state.records.sort(compareFirestoreSheetRecords_);
    }

    return {
      getValues: function() {
        return sliceMatrix();
      },
      getValue: function() {
        var values = sliceMatrix();
        return values[0] && values[0].length ? values[0][0] : "";
      },
      setValues: function(values) {
        applyValues(values || []);
        return this;
      },
      setValue: function(value) {
        applyValues([[value]]);
        return this;
      },
      clearContent: function() {
        clearCells();
        return this;
      },
      setDataValidation: function() { return this; },
      setNumberFormat: function() { return this; },
      setFontWeight: function() { return this; },
      setBackground: function() { return this; },
      setFontColor: function() { return this; },
      setWrap: function() { return this; },
      setHorizontalAlignment: function() { return this; }
    };
  }

  return {
    __isFirestoreSheetAdapter: true,
    getName: function() {
      return state.config.sheetName;
    },
    getDataRange: function() {
      return {
        getValues: function() {
          return getMatrix();
        }
      };
    },
    getRange: function(row, column, numRows, numColumns) {
      return rangeFactory(row, column, numRows, numColumns);
    },
    appendRow: function(values) {
      ensureLoaded(false);
      var row = normalizeRowLength_(values || [], state.headers.length);
      if (isEmptySheetRow_(row)) return this;

      var docId = buildAdapterDocumentId_(state.config, row, state.records.length + 2, "");
      var payload = buildAdapterDocumentPayload_(state.config, state.headers, row, state.records.length + 2, docId, null);
      upsertFirestoreCollectionRecord_(state.config.collection, docId, payload);

      state.records.push({
        docId: docId,
        payload: payload,
        row: row,
        sortKey: getFirestoreRecordSortKey_(payload, docId)
      });
      state.records.sort(compareFirestoreSheetRecords_);
      return this;
    },
    deleteRow: function(rowIndex) {
      ensureLoaded(false);
      var absoluteRow = Math.max(parseInt(rowIndex, 10) || 1, 1);
      if (absoluteRow <= 1) return this;

      var bodyIndex = absoluteRow - 2;
      if (bodyIndex < 0 || bodyIndex >= state.records.length) return this;

      deleteFirestoreCollectionRecord_(state.config.collection, state.records[bodyIndex].docId);
      state.records.splice(bodyIndex, 1);
      return this;
    },
    setFrozenRows: function() { return this; },
    autoResizeColumns: function() { return this; },
    getMaxRows: function() {
      ensureLoaded(false);
      return Math.max(state.records.length + 25, 2);
    },
    getLastRow: function() {
      ensureLoaded(false);
      return state.records.length + 1;
    },
    getLastColumn: function() {
      return state.headers.length;
    }
  };
}

function listFirestoreCollectionRecords_(collectionBaseName) {
  var collectionName = getFirebaseCollectionName_(collectionBaseName);
  var path = "documents/" + collectionName;
  var pageToken = "";
  var docs = [];

  do {
    var response = firestoreRequest_("get", path, null, {
      pageSize: 500,
      pageToken: pageToken || null
    });

    var pageDocs = response.documents || [];
    for (var i = 0; i < pageDocs.length; i++) {
      docs.push(pageDocs[i]);
    }

    pageToken = response.nextPageToken || "";
  } while (pageToken);

  return docs;
}

function getFirestoreCollectionRecord_(collectionBaseName, documentId) {
  try {
    return firestoreRequest_(
      "get",
      "documents/" + getFirebaseCollectionName_(collectionBaseName) + "/" + encodeURIComponent(documentId)
    );
  } catch (error) {
    if (String(error).indexOf("HTTP 404") !== -1) return null;
    throw error;
  }
}

function upsertFirestoreCollectionRecord_(collectionBaseName, documentId, payload) {
  return firestoreSetDocument_(collectionBaseName, documentId, payload);
}

function deleteFirestoreCollectionRecord_(collectionBaseName, documentId) {
  if (!documentId) return { success: true };

  return firestoreRequest_(
    "delete",
    "documents/" + getFirebaseCollectionName_(collectionBaseName) + "/" + encodeURIComponent(documentId)
  );
}

function clearFirestoreSheetData_(sheetName) {
  var config = JW_FIRESTORE_APP_DATA_CONFIG[sheetName];
  if (!config) return { success: false, message: "Aba nao mapeada: " + sheetName };

  var docs = listFirestoreCollectionRecords_(config.collection);
  if (!docs.length) return { success: true, deleted: 0 };

  var writes = [];
  for (var i = 0; i < docs.length; i++) {
    writes.push({ delete: docs[i].name });
  }

  var batches = splitFirestoreWrites_(writes, 200);
  var deleted = 0;
  for (var j = 0; j < batches.length; j++) {
    firestoreCommitWrites_(batches[j]);
    deleted += batches[j].length;
  }

  if (JW_FIRESTORE_SHEET_STATE_CACHE[sheetName]) {
    JW_FIRESTORE_SHEET_STATE_CACHE[sheetName].loaded = false;
    JW_FIRESTORE_SHEET_STATE_CACHE[sheetName].records = [];
  }

  return { success: true, deleted: deleted };
}

function buildFirestoreSheetRecord_(config, document) {
  var payload = fromFirestoreFields_((document && document.fields) || {});
  var docId = extractFirestoreDocumentId_(document && document.name);
  var data = payload && payload.data ? payload.data : (payload || {});
  var headers = config.headers || [];
  var sourceValues = Array.isArray(payload.source_values) ? payload.source_values : [];
  var row = [];

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var key = normalizeSheetFieldName_(header);
    var value = data.hasOwnProperty(key) ? data[key] : sourceValues[i];
    if (value === undefined || value === null) value = "";
    row.push(value);
  }

  return {
    docId: payload.document_id || docId,
    payload: payload,
    row: normalizeRowLength_(row, headers.length),
    sortKey: getFirestoreRecordSortKey_(payload, docId)
  };
}

function buildAdapterDocumentPayload_(config, headers, row, rowNumber, documentId, existingPayload) {
  var nowIso = new Date().toISOString();
  var previous = existingPayload || {};

  return {
    document_id: documentId,
    source_sheet: config.sheetName,
    source_row_number: previous.source_row_number || rowNumber,
    source_headers: normalizeSheetHeaderArray_(headers),
    source_values: normalizeSheetValueArray_(row),
    data: convertSheetRowToObject_(headers, row),
    created_at: previous.created_at || nowIso,
    updated_at: nowIso,
    sort_order: previous.sort_order !== undefined && previous.sort_order !== null
      ? previous.sort_order
      : (previous.source_row_number || rowNumber)
  };
}

function buildAdapterDocumentId_(config, row, rowNumber, currentDocId, usedDocumentIds, headersOverride) {
  if (currentDocId) {
    return sanitizeFirestoreDocumentId_(currentDocId, config.fallbackPrefix + "_" + rowNumber);
  }

  var headers = headersOverride || config.headers || [];
  var record = convertSheetRowToObject_(headers, row);
  var idFields = config.idFields || [];
  var documentId = "";

  for (var i = 0; i < idFields.length; i++) {
    var key = normalizeSheetFieldName_(idFields[i]);
    var value = record[key];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      documentId = sanitizeFirestoreDocumentId_(value, config.fallbackPrefix + "_" + rowNumber);
      break;
    }
  }

  if (!documentId) {
    documentId = sanitizeFirestoreDocumentId_(config.fallbackPrefix + "_" + rowNumber, config.fallbackPrefix + "_" + rowNumber);
  }

  if (usedDocumentIds) {
    documentId = ensureUniqueMigrationDocumentId_(documentId, usedDocumentIds, rowNumber);
  }

  return documentId;
}

function getFirestoreRecordSortKey_(payload, docId) {
  if (payload && payload.sort_order !== undefined && payload.sort_order !== null) {
    return "1_" + String(payload.sort_order);
  }
  if (payload && payload.source_row_number !== undefined && payload.source_row_number !== null) {
    return "2_" + String(payload.source_row_number);
  }
  if (payload && payload.created_at) {
    return "3_" + String(payload.created_at);
  }
  return "9_" + String(docId || "");
}

function compareFirestoreSheetRecords_(a, b) {
  var keyA = a && a.sortKey ? String(a.sortKey) : "";
  var keyB = b && b.sortKey ? String(b.sortKey) : "";
  if (keyA < keyB) return -1;
  if (keyA > keyB) return 1;

  var docA = a && a.docId ? String(a.docId) : "";
  var docB = b && b.docId ? String(b.docId) : "";
  if (docA < docB) return -1;
  if (docA > docB) return 1;
  return 0;
}

function normalizeRowLength_(row, targetLength) {
  var output = [];
  var source = row || [];
  for (var i = 0; i < targetLength; i++) {
    output.push(source[i] === undefined ? "" : source[i]);
  }
  return output;
}

function createBlankRow_(length) {
  var output = [];
  for (var i = 0; i < length; i++) output.push("");
  return output;
}

function extractFirestoreDocumentId_(name) {
  var parts = String(name || "").split("/");
  return parts.length ? parts[parts.length - 1] : "";
}

function fromFirestoreFields_(fields) {
  var output = {};
  var source = fields || {};

  for (var key in source) {
    if (!source.hasOwnProperty(key)) continue;
    output[key] = fromFirestoreValue_(source[key]);
  }

  return output;
}

function fromFirestoreValue_(value) {
  if (!value || typeof value !== "object") return null;
  if (value.stringValue !== undefined) return value.stringValue;
  if (value.booleanValue !== undefined) return value.booleanValue;
  if (value.integerValue !== undefined) return parseInt(value.integerValue, 10);
  if (value.doubleValue !== undefined) return value.doubleValue;
  if (value.timestampValue !== undefined) return value.timestampValue;
  if (value.nullValue !== undefined) return null;

  if (value.arrayValue) {
    var items = value.arrayValue.values || [];
    var output = [];
    for (var i = 0; i < items.length; i++) {
      output.push(fromFirestoreValue_(items[i]));
    }
    return output;
  }

  if (value.mapValue) {
    return fromFirestoreFields_(value.mapValue.fields || {});
  }

  return null;
}

function warmFirestoreAppDataCollections_() {
  var details = {};

  for (var sheetName in JW_FIRESTORE_APP_DATA_CONFIG) {
    if (!JW_FIRESTORE_APP_DATA_CONFIG.hasOwnProperty(sheetName)) continue;

    var config = JW_FIRESTORE_APP_DATA_CONFIG[sheetName];
    var collectionName = config.collection;

    try {
      getFirestoreSheetAdapter_(sheetName, config.headers).getLastRow();
      details[sheetName] = {
        collection: getFirebaseCollectionName_(collectionName),
        documents: listFirestoreCollectionRecords_(collectionName).length
      };
    } catch (error) {
      details[sheetName] = {
        collection: getFirebaseCollectionName_(collectionName),
        error: error && error.message ? error.message : String(error)
      };
    }
  }

  return details;
}

function getFirebaseAppDataHealth_() {
  var summary = {};
  var enabled = shouldUseFirebaseAppData_();
  var warmup = {};
  var configStatus = typeof validarConfiguracaoFirebaseMigracao === "function"
    ? validarConfiguracaoFirebaseMigracao()
    : { success: false, message: "Configuracao Firebase indisponivel." };

  if (enabled) {
    warmup = warmFirestoreAppDataCollections_();
  }

  for (var sheetName in JW_FIRESTORE_APP_DATA_CONFIG) {
    if (!JW_FIRESTORE_APP_DATA_CONFIG.hasOwnProperty(sheetName)) continue;
    var collection = JW_FIRESTORE_APP_DATA_CONFIG[sheetName].collection;
    try {
      summary[collection] = listFirestoreCollectionRecords_(collection).length;
    } catch (error) {
      summary[collection] = "erro";
    }
  }

  return {
    success: !!(configStatus && configStatus.success),
    firebaseEnabled: enabled,
    config: configStatus,
    warmup: warmup,
    collections: summary,
    checkedAt: new Date().toISOString()
  };
}
