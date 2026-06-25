// CORREÇÃO COMPLETA - Main.gs
function doGet(e) {
  console.log("🚀 Servindo aplicação...");
  
  try {
    if (e && e.parameter && e.parameter.asset) {
      return servirAsset_(e.parameter.asset);
    }

    if (e && e.parameter && e.parameter.api === "health") {
      return ContentService
        .createTextOutput(JSON.stringify(getFirebaseAppDataHealth_(), null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Retornar o HTML processado CORRETAMENTE
    return HtmlService
      .createTemplateFromFile('AppIndex')
      .evaluate()
      .setTitle('JW Piscinas - Sistema Mobile')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
      
  } catch (error) {
    console.error("❌ Erro no doGet:", error);
    return HtmlService.createHtmlOutput("Erro ao carregar: " + error.toString());
  }
}

// FUNÇÃO include CORRIGIDA - DEVE ESTAR NO Main.gs
function servirAsset_(nome) {
  if (nome === "sw") {
    return ContentService
      .createTextOutput(getServiceWorkerContent_())
      .setMimeType(getJavascriptMimeType_());
  }

  var permitidos = {
    Script: true,
    LoginBootstrap: true
  };

  if (!permitidos[nome]) {
    return ContentService
      .createTextOutput("// Asset nao permitido: " + nome)
      .setMimeType(getJavascriptMimeType_());
  }

  var conteudo = HtmlService.createHtmlOutputFromFile(nome).getContent();
  conteudo = extrairJavascriptDeHtml_(conteudo);

  return ContentService
    .createTextOutput(conteudo)
    .setMimeType(getJavascriptMimeType_());
}

function getAssetUrl(nome) {
  var url = ScriptApp.getService().getUrl();
  url = url.replace(/\/dev(\?|$)/, "/exec$1");
  var separador = url.indexOf("?") === -1 ? "?" : "&";
  return url + separador + "asset=" + encodeURIComponent(nome) + "&v=" + new Date().getTime();
}

function getAssetPath(nome) {
  return "?asset=" + encodeURIComponent(nome) + "&v=" + new Date().getTime();
}

function extrairJavascriptDeHtml_(conteudo) {
  conteudo = conteudo || "";
  var inicioScript = conteudo.toLowerCase().indexOf("<script");
  if (inicioScript === -1) return conteudo;

  var fimTagInicio = conteudo.indexOf(">", inicioScript);
  var fimScript = conteudo.toLowerCase().lastIndexOf("</script>");

  if (fimTagInicio === -1) return conteudo;
  if (fimScript === -1 || fimScript <= fimTagInicio) return conteudo.substring(fimTagInicio + 1);

  return conteudo.substring(fimTagInicio + 1, fimScript);
}

function getJavascriptMimeType_() {
  return ContentService.MimeType.JAVASCRIPT || ContentService.MimeType.TEXT;
}

function getServiceWorkerContent_() {
  return [
    "const CACHE_NAME = 'jw-piscinas-pwa-v1';",
    "self.addEventListener('install', event => { self.skipWaiting(); });",
    "self.addEventListener('activate', event => { event.waitUntil(self.clients.claim()); });",
    "self.addEventListener('fetch', event => {",
    "  if (event.request.method !== 'GET') return;",
    "  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));",
    "});"
  ].join('\\n');
}

function include(filename) {
  console.log("📄 Incluindo arquivo:", filename);
  try {
    // Retorna o conteúdo do arquivo HTML
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error("❌ Erro ao incluir", filename, ":", error);
    var mensagem = "Erro ao incluir " + filename + ": " + error.toString();
    return "<script>window.__jwIncludeErrors = window.__jwIncludeErrors || []; window.__jwIncludeErrors.push(" +
      JSON.stringify(mensagem) +
      ");</script><div class=\"alert alert-error\" style=\"margin:12px;\">Arquivo nao carregado: " +
      filename +
      "<br><small>" +
      mensagem.replace(/</g, "&lt;").replace(/>/g, "&gt;") +
      "</small></div>";
  }
}

// Função de teste (mantida)
function testarConexao() {
  console.log("Testando conexão...");
  
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var nome = planilha.getName();
    
    return {
      success: true,
      message: "✅ Sistema JW Piscinas Mobile está funcionando!",
      timestamp: new Date().toLocaleString('pt-BR'),
      planilha: nome,
      versao: "2.0"
    };
    
  } catch (error) {
    console.error("❌ Erro no teste de conexão:", error);
    return {
      success: false,
      message: "❌ Erro na conexão: " + error.toString()
    };
  }
}

function diagnosticoArquivosHtml() {
  var arquivos = [
    "AppIndex",
    "Style",
    "Header",
    "LoginScreen",
    "MainApp",
    "Modals",
    "LoginBootstrap",
    "Script"
  ];
  var resultado = {};

  for (var i = 0; i < arquivos.length; i++) {
    var nome = arquivos[i];
    try {
      var conteudo = HtmlService.createHtmlOutputFromFile(nome).getContent();
      resultado[nome] = {
        ok: true,
        tamanho: conteudo.length,
        inicio: conteudo.substring(0, 120),
        temScriptStarted: conteudo.indexOf("window.__jwScriptStarted") !== -1,
        temScriptLoaded: conteudo.indexOf("window.__jwScriptLoaded") !== -1
      };
    } catch (error) {
      resultado[nome] = {
        ok: false,
        erro: error.toString()
      };
    }
  }

  try {
    resultado.assetScriptUrl = getAssetUrl("Script");
    resultado.assetScriptPath = getAssetPath("Script");
    var assetConteudo = extrairJavascriptDeHtml_(HtmlService.createHtmlOutputFromFile("Script").getContent());
    resultado.assetScript = {
      tamanho: assetConteudo.length,
      inicio: assetConteudo.substring(0, 120),
      terminaComLoaded: assetConteudo.indexOf("window.__jwScriptLoaded = true") !== -1
    };
  } catch (urlError) {
    resultado.assetScriptUrl = "Erro ao gerar URL: " + urlError.toString();
  }

  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
