// OrdemServicoService.gs - Gerenciamento de Ordens de Serviço

function getAbaOrdensServico() {
  try {
    if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
      return getFirestoreSheetAdapter_("OrdensServico");
    }

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("OrdensServico");
    
    if (!aba) {
      aba = planilha.insertSheet("OrdensServico");
      aba.getRange("A1:K1").setValues([[
        "ID", "Cliente", "TecnicoID", "TipoServico", "DataAbertura", 
        "DataAgendamento", "DataConclusao", "Valor", "Status", "Descricao", "OrcamentoID"
      ]]);
      aba.setFrozenRows(1);
      
      // Adicionar ordens de exemplo
      var exemplos = [
        ["OS-001", "Cliente Exemplo", "TEC-001", "Limpeza", new Date().toLocaleDateString('pt-BR'), 
         new Date(Date.now() + 86400000).toLocaleDateString('pt-BR'), "", 150.00, "Agendado", "Limpeza semanal", "ORC-001"]
      ];
      for (var i = 0; i < exemplos.length; i++) {
        aba.appendRow(exemplos[i]);
      }
    }
    return aba;
  } catch (error) {
    console.error("Erro em getAbaOrdensServico:", error);
    throw error;
  }
}

function getAllOrdensServico() {
  try {
    var aba = getAbaOrdensServico();
    var dados = aba.getDataRange().getValues();
    var ordens = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      ordens.push({
        id: linha[0] || "",
        cliente: linha[1] || "",
        tecnicoID: linha[2] || "",
        tipoServico: linha[3] || "",
        dataAbertura: linha[4] || "",
        dataAgendamento: linha[5] || "",
        dataConclusao: linha[6] || "",
        valor: linha[7] || 0,
        status: linha[8] || "Pendente",
        descricao: linha[9] || "",
        orcamentoID: linha[10] || ""
      });
    }
    return ordens;
  } catch (error) {
    console.error("Erro em getAllOrdensServico:", error);
    return [];
  }
}

function getOrdensAtivas() {
  try {
    var ordens = getAllOrdensServico();
    var ativas = [];
    var statusAtivos = ["Agendado", "Em Andamento", "Pendente"];
    
    for (var i = 0; i < ordens.length; i++) {
      if (statusAtivos.indexOf(ordens[i].status) !== -1) {
        ativas.push(ordens[i]);
      }
    }
    return ativas;
  } catch (error) {
    console.error("Erro em getOrdensAtivas:", error);
    return [];
  }
}
