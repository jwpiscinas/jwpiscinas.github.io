// TecnicoService.gs - Gerenciamento de Técnicos

function getAbaTecnicos() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("Tecnicos");
    
    if (!aba) {
      aba = planilha.insertSheet("Tecnicos");
      aba.getRange("A1:F1").setValues([[
        "ID", "Nome", "Telefone", "Especialidade", "Status", "DataCadastro"
      ]]);
      aba.setFrozenRows(1);
      
      // Adicionar técnicos de exemplo
      var exemplos = [
        ["TEC-001", "João Silva", "(11) 91234-5678", "Manutenção Geral", "Ativo", new Date().toLocaleDateString('pt-BR')],
        ["TEC-002", "Maria Santos", "(11) 98765-4321", "Tratamento Químico", "Ativo", new Date().toLocaleDateString('pt-BR')]
      ];
      for (var i = 0; i < exemplos.length; i++) {
        aba.appendRow(exemplos[i]);
      }
    }
    return aba;
  } catch (error) {
    console.error("Erro em getAbaTecnicos:", error);
    throw error;
  }
}

function getTecnicosAtivos() {
  try {
    var aba = getAbaTecnicos();
    var dados = aba.getDataRange().getValues();
    var tecnicos = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var status = linha[4] ? linha[4].toString() : "";
      if (status === "Ativo") {
        tecnicos.push({
          id: linha[0] || "",
          nome: linha[1] || "",
          telefone: linha[2] || "",
          especialidade: linha[3] || "",
          status: status,
          dataCadastro: linha[5] || ""
        });
      }
    }
    return tecnicos;
  } catch (error) {
    console.error("Erro em getTecnicosAtivos:", error);
    return [];
  }
}

function getAllTecnicos() {
  try {
    var aba = getAbaTecnicos();
    var dados = aba.getDataRange().getValues();
    var tecnicos = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      tecnicos.push({
        id: linha[0] || "",
        nome: linha[1] || "",
        telefone: linha[2] || "",
        especialidade: linha[3] || "",
        status: linha[4] || "",
        dataCadastro: linha[5] || ""
      });
    }
    return tecnicos;
  } catch (error) {
    console.error("Erro em getAllTecnicos:", error);
    return [];
  }
}

function adicionarTecnico(nome, telefone, especialidade) {
  try {
    var aba = getAbaTecnicos();
    var id = "TEC-" + new Date().getTime().toString().slice(-6);
    var novaLinha = [
      id,
      nome,
      telefone,
      especialidade,
      "Ativo",
      new Date().toLocaleDateString('pt-BR')
    ];
    aba.appendRow(novaLinha);
    SpreadsheetApp.flush();
    return { success: true, message: "✅ Técnico adicionado com sucesso!", id: id };
  } catch (error) {
    return { success: false, message: "Erro: " + error.toString() };
  }
}