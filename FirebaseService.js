var FIREBASE_MIGRATION_DEFAULTS = {
  PROJECT_ID: "jwpiscinas",
  DATABASE_ID: "(default)",
  COLLECTION_PREFIX: "jw",
  TOKEN_SCOPE: "https://www.googleapis.com/auth/datastore",
  TOKEN_URL: "https://oauth2.googleapis.com/token",
  CACHE_KEY: "jw_firebase_access_token"
};

function getFirebaseMigrationConfig_() {
  var props = PropertiesService.getScriptProperties();

  return {
    projectId: props.getProperty("FIREBASE_PROJECT_ID") || FIREBASE_MIGRATION_DEFAULTS.PROJECT_ID,
    databaseId: props.getProperty("FIREBASE_DATABASE_ID") || FIREBASE_MIGRATION_DEFAULTS.DATABASE_ID,
    collectionPrefix: props.getProperty("FIREBASE_COLLECTION_PREFIX") || FIREBASE_MIGRATION_DEFAULTS.COLLECTION_PREFIX,
    clientEmail: props.getProperty("FIREBASE_SERVICE_ACCOUNT_EMAIL") || "",
    privateKey: normalizeFirebasePrivateKey_(props.getProperty("FIREBASE_SERVICE_ACCOUNT_PRIVATE_KEY") || "")
  };
}

function salvarConfiguracaoFirebaseMigracao(config) {
  config = config || {};

  var props = PropertiesService.getScriptProperties();
  var data = {};

  if (config.projectId) data.FIREBASE_PROJECT_ID = String(config.projectId).trim();
  if (config.databaseId) data.FIREBASE_DATABASE_ID = String(config.databaseId).trim();
  if (config.collectionPrefix) data.FIREBASE_COLLECTION_PREFIX = String(config.collectionPrefix).trim();
  if (config.clientEmail) data.FIREBASE_SERVICE_ACCOUNT_EMAIL = String(config.clientEmail).trim();
  if (config.privateKey) data.FIREBASE_SERVICE_ACCOUNT_PRIVATE_KEY = normalizeFirebasePrivateKey_(config.privateKey);

  if (!Object.keys(data).length) {
    return {
      success: false,
      message: "Nenhum valor foi enviado para salvar na configuracao do Firebase."
    };
  }

  props.setProperties(data, false);
  CacheService.getScriptCache().remove(FIREBASE_MIGRATION_DEFAULTS.CACHE_KEY);

  return validarConfiguracaoFirebaseMigracao();
}

function validarConfiguracaoFirebaseMigracao() {
  var config = getFirebaseMigrationConfig_();
  var missing = [];

  if (!config.projectId) missing.push("FIREBASE_PROJECT_ID");
  if (!config.clientEmail) missing.push("FIREBASE_SERVICE_ACCOUNT_EMAIL");
  if (!config.privateKey) missing.push("FIREBASE_SERVICE_ACCOUNT_PRIVATE_KEY");

  return {
    success: missing.length === 0,
    projectId: config.projectId,
    databaseId: config.databaseId,
    collectionPrefix: config.collectionPrefix,
    clientEmailConfigured: !!config.clientEmail,
    privateKeyConfigured: !!config.privateKey,
    missing: missing,
    message: missing.length === 0
      ? "Configuracao do Firebase pronta para uso."
      : "Configure as propriedades faltantes do Firebase: " + missing.join(", ")
  };
}

function testarConexaoFirebaseFirestore() {
  try {
    var status = validarConfiguracaoFirebaseMigracao();
    if (!status.success) return status;

    var payload = {
      status: "ok",
      source: "apps_script",
      updated_at: new Date().toISOString()
    };

    firestoreSetDocument_("meta_runtime", "connectivity", payload);

    return {
      success: true,
      message: "Conexao com Cloud Firestore validada com sucesso.",
      projectId: getFirebaseMigrationConfig_().projectId
    };
  } catch (error) {
    return {
      success: false,
      message: "Falha ao validar conexao com Cloud Firestore: " + error.toString()
    };
  }
}

function firestoreSetDocument_(collectionBaseName, documentId, data) {
  var write = buildFirestoreSetWrite_(collectionBaseName, documentId, data);
  return firestoreCommitWrites_([write]);
}

function firestoreCommitWrites_(writes) {
  if (!writes || !writes.length) {
    return {
      success: true,
      writeResults: []
    };
  }

  var response = firestoreRequest_("post", "documents:commit", {
    writes: writes
  });

  return {
    success: true,
    writeResults: response.writeResults || [],
    commitTime: response.commitTime || ""
  };
}

function buildFirestoreSetWrite_(collectionBaseName, documentId, data) {
  return {
    update: {
      name: buildFirestoreDocumentName_(collectionBaseName, documentId),
      fields: toFirestoreFields_(data || {})
    }
  };
}

function buildFirestoreDocumentName_(collectionBaseName, documentId) {
  var config = getFirebaseMigrationConfig_();
  var collectionName = getFirebaseCollectionName_(collectionBaseName);
  var safeDocumentId = sanitizeFirestoreDocumentId_(documentId, Utilities.getUuid());

  return [
    "projects",
    config.projectId,
    "databases",
    config.databaseId,
    "documents",
    collectionName,
    safeDocumentId
  ].join("/");
}

function getFirebaseCollectionName_(collectionBaseName) {
  var config = getFirebaseMigrationConfig_();
  var prefix = (config.collectionPrefix || "").trim();
  var base = sanitizeFirestoreSegment_(collectionBaseName);

  if (!prefix) return base;
  return sanitizeFirestoreSegment_(prefix + "_" + base);
}

function firestoreRequest_(method, path, payload, query) {
  var token = getFirebaseAccessToken_();
  var config = getFirebaseMigrationConfig_();
  var url = "https://firestore.googleapis.com/v1/projects/" +
    encodeURIComponent(config.projectId) +
    "/databases/" +
    encodeURIComponent(config.databaseId) +
    "/" +
    path;

  if (query) {
    var params = [];
    for (var key in query) {
      if (!query.hasOwnProperty(key) || query[key] === undefined || query[key] === null) continue;
      params.push(encodeURIComponent(key) + "=" + encodeURIComponent(query[key]));
    }
    if (params.length) url += "?" + params.join("&");
  }

  var options = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + token
    }
  };

  if (payload !== undefined && payload !== null) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();
  var body = response.getContentText() || "";

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Firestore HTTP " + statusCode + ": " + body);
  }

  return body ? JSON.parse(body) : {};
}

function getFirebaseAccessToken_() {
  var cache = CacheService.getScriptCache();
  var cachedToken = cache.get(FIREBASE_MIGRATION_DEFAULTS.CACHE_KEY);
  if (cachedToken) return cachedToken;

  var config = getFirebaseMigrationConfig_();
  if (!config.clientEmail || !config.privateKey) {
    throw new Error("Service account do Firebase nao configurada nas Script Properties.");
  }

  var nowSeconds = Math.floor(Date.now() / 1000);
  var header = {
    alg: "RS256",
    typ: "JWT"
  };
  var claimSet = {
    iss: config.clientEmail,
    scope: FIREBASE_MIGRATION_DEFAULTS.TOKEN_SCOPE,
    aud: FIREBASE_MIGRATION_DEFAULTS.TOKEN_URL,
    exp: nowSeconds + 3600,
    iat: nowSeconds
  };

  var unsignedToken = [
    base64WebSafeNoPad_(JSON.stringify(header)),
    base64WebSafeNoPad_(JSON.stringify(claimSet))
  ].join(".");

  var signatureBytes = Utilities.computeRsaSha256Signature(unsignedToken, config.privateKey);
  var signedToken = unsignedToken + "." + base64WebSafeNoPad_(signatureBytes);

  var response = UrlFetchApp.fetch(FIREBASE_MIGRATION_DEFAULTS.TOKEN_URL, {
    method: "post",
    muteHttpExceptions: true,
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: signedToken
    }
  });

  var statusCode = response.getResponseCode();
  var body = response.getContentText() || "";

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Nao foi possivel obter token OAuth do Firebase: HTTP " + statusCode + " - " + body);
  }

  var tokenData = JSON.parse(body);
  if (!tokenData.access_token) {
    throw new Error("Resposta de token invalida: " + body);
  }

  var ttlSeconds = Math.max((tokenData.expires_in || 3600) - 60, 300);
  cache.put(FIREBASE_MIGRATION_DEFAULTS.CACHE_KEY, tokenData.access_token, ttlSeconds);

  return tokenData.access_token;
}

function toFirestoreFields_(obj) {
  var fields = {};
  var source = obj || {};

  for (var key in source) {
    if (!source.hasOwnProperty(key)) continue;
    if (source[key] === undefined) continue;
    fields[key] = toFirestoreValue_(source[key]);
  }

  return fields;
}

function toFirestoreValue_(value) {
  if (value === null) return { nullValue: null };

  if (value instanceof Date) {
    return { timestampValue: value.toISOString() };
  }

  if (Array.isArray(value)) {
    var values = [];
    for (var i = 0; i < value.length; i++) {
      values.push(toFirestoreValue_(value[i]));
    }
    return { arrayValue: { values: values } };
  }

  var type = typeof value;

  if (type === "string") return { stringValue: value };
  if (type === "boolean") return { booleanValue: value };
  if (type === "number") {
    if (Math.floor(value) === value) return { integerValue: String(value) };
    return { doubleValue: value };
  }

  if (type === "object") {
    return { mapValue: { fields: toFirestoreFields_(value) } };
  }

  return { stringValue: String(value) };
}

function normalizeFirebasePrivateKey_(value) {
  if (!value) return "";
  return String(value).replace(/\\n/g, "\n").trim();
}

function sanitizeFirestoreDocumentId_(value, fallback) {
  var text = value === undefined || value === null ? "" : String(value).trim();
  text = removeAccentsFirestore_(text);
  text = text.replace(/[\/#?\[\]]/g, "_");
  text = text.replace(/\s+/g, "_");
  text = text.replace(/[^A-Za-z0-9._-]/g, "_");
  text = text.replace(/_+/g, "_");
  text = text.replace(/^_+|_+$/g, "");

  if (!text) {
    text = fallback || Utilities.getUuid();
  }

  return text.substring(0, 500);
}

function sanitizeFirestoreSegment_(value) {
  var text = value === undefined || value === null ? "" : String(value).trim().toLowerCase();
  text = removeAccentsFirestore_(text);
  text = text.replace(/\s+/g, "_");
  text = text.replace(/[^a-z0-9_-]/g, "_");
  text = text.replace(/_+/g, "_");
  text = text.replace(/^_+|_+$/g, "");
  return text || "default";
}

function removeAccentsFirestore_(value) {
  var text = String(value || "");

  if (typeof text.normalize === "function") {
    return text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  return text
    .replace(/[ÁÀÃÂÄ]/g, "A")
    .replace(/[áàãâä]/g, "a")
    .replace(/[ÉÈÊË]/g, "E")
    .replace(/[éèêë]/g, "e")
    .replace(/[ÍÌÎÏ]/g, "I")
    .replace(/[íìîï]/g, "i")
    .replace(/[ÓÒÕÔÖ]/g, "O")
    .replace(/[óòõôö]/g, "o")
    .replace(/[ÚÙÛÜ]/g, "U")
    .replace(/[úùûü]/g, "u")
    .replace(/[Ç]/g, "C")
    .replace(/[ç]/g, "c")
    .replace(/[Ñ]/g, "N")
    .replace(/[ñ]/g, "n");
}

function base64WebSafeNoPad_(value) {
  return Utilities.base64EncodeWebSafe(value).replace(/=+$/g, "");
}
