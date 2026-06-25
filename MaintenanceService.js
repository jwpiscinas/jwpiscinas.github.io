function reorganizarSistemaPreservandoProdutos(confirmacao) {
  if (confirmacao !== "CONFIRMAR") {
    return {
      success: false,
      message: "Para limpar dados de teste, execute reorganizarSistemaPreservandoProdutos('CONFIRMAR')."
    };
  }

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abasParaLimpar = [
    "HistoricoGeral",
    "Compras",
    "Orcamentos",
    "Notificacoes",
    "Pagamentos",
    "Servicos",
    "OrdensServico",
    "Agenda",
    "LogsSistema"
  ];

  var limpas = [];

  if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
    for (var firebaseIndex = 0; firebaseIndex < abasParaLimpar.length; firebaseIndex++) {
      var sheetName = abasParaLimpar[firebaseIndex];
      try {
        clearFirestoreSheetData_(sheetName);
        limpas.push(sheetName);
      } catch (firebaseError) {}
    }

    getAbaUsuarios();
    getAbaProdutos();
    getAbaOrcamentos();
    getAbaCompras();
    getAbaNotificacoes();
    criarAbaHistoricoGeral();
    try { getAbaOrdensServico(); } catch (e1) {}
    try { getAbaAgenda(); } catch (e2) {}

    return {
      success: true,
      message: "Sistema reorganizado no Firestore. A colecao de Produtos foi preservada.",
      abasLimpas: limpas
    };
  }

  for (var i = 0; i < abasParaLimpar.length; i++) {
    var aba = planilha.getSheetByName(abasParaLimpar[i]);
    if (!aba) continue;

    var ultimaLinha = aba.getLastRow();
    var ultimaColuna = aba.getLastColumn();

    if (ultimaLinha > 1 && ultimaColuna > 0) {
      aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).clearContent();
    }

    limpas.push(abasParaLimpar[i]);
  }

  // Reaplica os cabecalhos e validacoes sem alterar dados de Produtos.
  getAbaUsuarios();
  getAbaProdutos();
  getAbaOrcamentos();
  getAbaCompras();
  getAbaNotificacoes();
  criarAbaHistoricoGeral();
  try { getAbaOrdensServico(); } catch (e) {}
  try { getAbaAgenda(); } catch (e) {}

  SpreadsheetApp.flush();

  return {
    success: true,
    message: "Sistema reorganizado. A aba Produtos foi preservada.",
    abasLimpas: limpas
  };
}
