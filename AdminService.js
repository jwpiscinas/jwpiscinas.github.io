function getCentralAdmin() {
  return {
    orcamentosPendentes: getOrcamentosPendentes(),
    comprasPendentes: getComprasPendentesAdmin(),
    ordensAtivas: getOrdensAtivas()
  };
}

function getComprasPendentesAdmin() {
  try {
    var aba = getAbaCompras();
    var dados = aba.getDataRange().getValues();
    var compras = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var statusEntrega = linha[8] ? linha[8].toString() : "Processando";

      if (statusEntrega === "Entregue" || statusEntrega === "Cancelado") continue;

      compras.push({
        id: linha[0] || "",
        cliente: linha[1] || "",
        email: linha[2] || "",
        dataCompra: linha[3] || "",
        produtos: linha[4] || "",
        quantidadeTotal: linha[5] || 0,
        valorTotal: linha[6] || 0,
        statusPagamento: linha[7] || "Pendente",
        statusEntrega: statusEntrega,
        linha: i + 1
      });
    }

    return compras;
  } catch (error) {
    console.error("Erro em getComprasPendentesAdmin:", error);
    return [];
  }
}

function atualizarStatusCompra(idCompra, statusEntrega) {
  try {
    var aba = getAbaCompras();
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === idCompra) {
        aba.getRange(i + 1, 9).setValue(statusEntrega);
        SpreadsheetApp.flush();
        return { success: true, message: "Compra atualizada." };
      }
    }

    return { success: false, message: "Compra nao encontrada." };
  } catch (error) {
    return { success: false, message: "Erro: " + error.toString() };
  }
}

function getRelatoriosAdmin() {
  try {
    var orcamentos = getAllOrcamentos();
    var compras = getTodasComprasAdmin_();
    var usuarios = getTodosUsuarios();

    var totalOrcamentos = orcamentos.length;
    var orcamentosPendentes = orcamentos.filter(function(o) {
      return o.status === "Solicitado" || o.status === "Em Negociacao" || o.status === "Em Negociação";
    }).length;

    var totalCompras = compras.length;
    var comprasPendentes = compras.filter(function(c) {
      return c.statusEntrega !== "Entregue" && c.statusEntrega !== "Cancelado";
    }).length;

    var faturamentoCompras = compras.reduce(function(total, compra) {
      return total + (parseFloat(compra.valorTotal) || 0);
    }, 0);

    return {
      totalClientes: usuarios.filter(function(u) { return u.tipo !== "Administrador"; }).length,
      totalAdmins: usuarios.filter(function(u) { return u.tipo === "Administrador"; }).length,
      totalOrcamentos: totalOrcamentos,
      orcamentosPendentes: orcamentosPendentes,
      totalCompras: totalCompras,
      comprasPendentes: comprasPendentes,
      faturamentoCompras: faturamentoCompras,
      orcamentosRecentes: orcamentos.slice(0, 5),
      comprasRecentes: compras.slice(0, 5)
    };
  } catch (error) {
    return { error: error.toString() };
  }
}

function getTodasComprasAdmin_() {
  var aba = getAbaCompras();
  var dados = aba.getDataRange().getValues();
  var compras = [];

  for (var i = 1; i < dados.length; i++) {
    compras.push({
      id: dados[i][0] || "",
      cliente: dados[i][1] || "",
      email: dados[i][2] || "",
      dataCompra: dados[i][3] || "",
      produtos: dados[i][4] || "",
      quantidadeTotal: dados[i][5] || 0,
      valorTotal: dados[i][6] || 0,
      statusPagamento: dados[i][7] || "Pendente",
      statusEntrega: dados[i][8] || "Processando"
    });
  }

  return compras;
}
