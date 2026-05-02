// PagamentoService.gs - Adicionar função

function getAbaPagamentos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Pagamentos");

  if (!aba) {
    aba = planilha.insertSheet("Pagamentos");
    aba.getRange("A1:J1").setValues([[
      "ID", "Cliente", "Email", "Tipo", "Valor", "Vencimento", "Status", "DataPagamento", "Metodo", "Observacoes"
    ]]);
    aba.setFrozenRows(1);
  }

  return aba;
}

function getPagamentosPendentes() {
  try {
    var aba = getAbaPagamentos();
    var dados = aba.getDataRange().getValues();
    var pagamentos = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var status = linha[6] ? linha[6].toString() : "";
      if (status === "Pendente") {
        var valor = linha[4];
        if (typeof valor === 'string') {
          valor = parseFloat(valor.replace('R$', '').replace(',', '.')) || 0;
        }
        pagamentos.push({
          id: linha[0] || "",
          cliente: linha[1] || "",
          email: linha[2] || "",
          tipo: linha[3] || "",
          valor: valor,
          vencimento: linha[5] || "",
          status: status
        });
      }
    }
    return pagamentos;
  } catch (error) {
    console.error("Erro em getPagamentosPendentes:", error);
    return [];
  }
}
