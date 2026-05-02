// AgendaService.gs - Gerenciamento de Agenda

function getAbaAgenda() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("Agenda");
    
    if (!aba) {
      aba = planilha.insertSheet("Agenda");
      aba.getRange("A1:G1").setValues([[
        "ID", "Data", "Hora", "TecnicoID", "Cliente", "Servico", "Status"
      ]]);
      aba.setFrozenRows(1);
      
      // Adicionar agendamentos de exemplo
      var dataHoje = new Date();
      var dataAmanha = new Date(Date.now() + 86400000);
      var exemplos = [
        ["AGE-001", dataHoje.toLocaleDateString('pt-BR'), "09:00", "TEC-001", "Cliente Exemplo", "Limpeza", "Agendado"],
        ["AGE-002", dataAmanha.toLocaleDateString('pt-BR'), "14:00", "TEC-002", "Cliente Teste", "Manutenção", "Agendado"]
      ];
      for (var i = 0; i < exemplos.length; i++) {
        aba.appendRow(exemplos[i]);
      }
    }
    return aba;
  } catch (error) {
    console.error("Erro em getAbaAgenda:", error);
    throw error;
  }
}

function getProximosAgendamentos() {
  try {
    var aba = getAbaAgenda();
    var dados = aba.getDataRange().getValues();
    var agenda = [];
    var hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var dataStr = linha[1] || "";
      var partes = dataStr.split('/');
      var dataAgendamento = null;
      
      if (partes.length === 3) {
        dataAgendamento = new Date(partes[2], partes[1] - 1, partes[0]);
      }
      
      if (dataAgendamento && dataAgendamento >= hoje) {
        agenda.push({
          id: linha[0] || "",
          data: linha[1] || "",
          hora: linha[2] || "",
          tecnicoID: linha[3] || "",
          cliente: linha[4] || "",
          servico: linha[5] || "",
          status: linha[6] || "Agendado"
        });
      }
    }
    
    // Ordenar por data e hora
    agenda.sort(function(a, b) {
      var dateA = a.data.split('/').reverse().join('-') + ' ' + a.hora;
      var dateB = b.data.split('/').reverse().join('-') + ' ' + b.hora;
      return dateA.localeCompare(dateB);
    });
    
    return agenda.slice(0, 5);
  } catch (error) {
    console.error("Erro em getProximosAgendamentos:", error);
    return [];
  }
}

function getAgendamentosPorTecnico(tecnicoID) {
  try {
    var aba = getAbaAgenda();
    var dados = aba.getDataRange().getValues();
    var agenda = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      if (linha[3] === tecnicoID) {
        agenda.push({
          id: linha[0] || "",
          data: linha[1] || "",
          hora: linha[2] || "",
          cliente: linha[4] || "",
          servico: linha[5] || "",
          status: linha[6] || ""
        });
      }
    }
    return agenda;
  } catch (error) {
    console.error("Erro em getAgendamentosPorTecnico:", error);
    return [];
  }
}