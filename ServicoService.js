// Obter aba de serviços
function getAbaServicos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Servicos");
  if (!aba) {
    aba = planilha.insertSheet("Servicos");
    aba.getRange("A1:J1").setValues([[
      "ID", 
      "Cliente", 
      "Email", 
      "TipoServico", 
      "DataInicio", 
      "DataConclusao", 
      "Valor", 
      "Status", 
      "Descricao",
      "OrcamentoID"
    ]]);
  }
  return aba;
}

// Obter serviços do cliente - ATUALIZADO PARA HISTÓRICO
function getServicosCliente(email) {
  try {
    console.log("=== BUSCANDO SERVIÇOS PARA HISTÓRICO:", email);
    
    var aba = getAbaServicos();
    var dados = aba.getDataRange().getValues();
    var servicos = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailServico = linha[2];
      
      if (emailServico && 
          emailServico.toString().trim().toLowerCase() === email.toLowerCase()) {
        
        // Formatar datas
        var dataInicio = linha[4] ? new Date(linha[4]) : null;
        var dataConclusao = linha[5] ? new Date(linha[5]) : null;
        
        var dataInicioFormatada = dataInicio ? dataInicio.toLocaleDateString('pt-BR') : '';
        var dataConclusaoFormatada = dataConclusao ? dataConclusao.toLocaleDateString('pt-BR') : '';
        
        servicos.push({
          id: linha[0] || "",
          tipoServico: linha[3] || "",
          dataInicio: dataInicioFormatada,
          dataConclusao: dataConclusaoFormatada,
          valor: linha[6] ? "R$ " + parseFloat(linha[6]).toFixed(2) : "A definir",
          status: linha[7] || "Pendente",
          descricao: linha[8] || "",
          orcamentoID: linha[9] || ""
        });
      }
    }
    
    // Ordenar por data mais recente primeiro
    servicos.sort(function(a, b) {
      var dateA = a.dataInicio ? new Date(a.dataInicio.split('/').reverse().join('-')) : new Date(0);
      var dateB = b.dataInicio ? new Date(b.dataInicio.split('/').reverse().join('-')) : new Date(0);
      return dateB - dateA;
    });
    
    console.log("✅ Serviços encontrados para histórico:", servicos.length);
    return servicos;
    
  } catch (error) {
    console.error("Erro getServicosCliente:", error);
    return [];
  }
}

// Resto do código permanece igual...