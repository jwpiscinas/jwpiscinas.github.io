// Sistema de logs
function formatarDataHoraSistema_(data, formato) {
  return Utilities.formatDate(data || new Date(), "America/Sao_Paulo", formato || "dd/MM/yyyy HH:mm:ss");
}

function formatarDataSistema_(data) {
  return Utilities.formatDate(data || new Date(), "America/Sao_Paulo", "dd/MM/yyyy");
}

function logEvent(tipo, usuario, acao, detalhes) {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("LogsSistema");

    if (!aba) {
      aba = planilha.insertSheet("LogsSistema");
      aba.getRange("A1:E1").setValues([[
        "DataHora", "Tipo", "Usuario", "Acao", "Detalhes"
      ]]);
      aba.setFrozenRows(1);
    }

    var detalhesStr = "";
    try {
      detalhesStr = JSON.stringify(detalhes);
    } catch (e) {
      detalhesStr = String(detalhes);
    }

    aba.appendRow([
      formatarDataHoraSistema_(),
      tipo,
      usuario,
      acao,
      detalhesStr
    ]);

    if (aba.getLastRow() > 1000) {
      aba.deleteRow(2);
    }
  } catch (error) {
    console.error("Erro ao registrar log:", error);
  }
}

// Upload de imagem para Google Drive
function uploadImagemParaDrive(base64Data, nomeArquivo) {
  try {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data.split(",")[1]),
      "image/jpeg",
      nomeArquivo
    );
    var driveFolder = DriveApp.getRootFolder();
    var file = driveFolder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (error) {
    console.error("Erro ao fazer upload:", error);
    return null;
  }
}

// Enviar notificacao generica
function enviarNotificacaoCliente(email, titulo, mensagem) {
  try {
    var assunto = "JW Piscinas - " + titulo;
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <p>${mensagem}</p>
        <hr>
        <p style="font-size: 12px; color: #666;">
          Esta e uma notificacao automatica do sistema JW Piscinas.
        </p>
      </div>
    `;

    MailApp.sendEmail({
      to: email,
      subject: assunto,
      htmlBody: corpo
    });

    console.log("Notificacao enviada para:", email);
    return true;
  } catch (error) {
    console.error("Erro ao enviar notificacao:", error);
    return false;
  }
}

var HISTORICO_GERAL_SHEET_NAME = "HistoricoGeral";
var HISTORICO_GERAL_HEADERS = [
  "ID",
  "Tipo",
  "Cliente",
  "Email",
  "Telefone",
  "Titulo",
  "Descricao",
  "Valor",
  "Status",
  "DataRegistro",
  "LinkAcao",
  "Observacoes"
];

function getAbaHistoricoGeral() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName(HISTORICO_GERAL_SHEET_NAME);

  if (!aba) {
    aba = planilha.insertSheet(HISTORICO_GERAL_SHEET_NAME);
    aba.getRange(1, 1, 1, HISTORICO_GERAL_HEADERS.length).setValues([HISTORICO_GERAL_HEADERS]);
    aba.setFrozenRows(1);
    return aba;
  }

  garantirCabecalhosHistoricoGeral_(aba);
  return aba;
}

function garantirCabecalhosHistoricoGeral_(aba) {
  var totalColunas = HISTORICO_GERAL_HEADERS.length;
  var primeiraLinha = aba.getRange(1, 1, 1, totalColunas).getValues()[0];
  var precisaCorrigir = false;

  for (var i = 0; i < totalColunas; i++) {
    if (primeiraLinha[i] !== HISTORICO_GERAL_HEADERS[i]) {
      precisaCorrigir = true;
      break;
    }
  }

  if (precisaCorrigir) {
    aba.getRange(1, 1, 1, totalColunas).setValues([HISTORICO_GERAL_HEADERS]);
    aba.setFrozenRows(1);
  }
}

function criarAbaHistoricoGeral() {
  try {
    getAbaHistoricoGeral();
    SpreadsheetApp.flush();
    return {
      success: true,
      message: "Aba HistoricoGeral pronta."
    };
  } catch (error) {
    console.error("Erro ao criar/verificar historico:", error);
    return {
      success: false,
      message: "Erro ao criar/verificar historico: " + error.toString()
    };
  }
}

function registrarHistorico(dados) {
  try {
    dados = dados || {};

    var aba = getAbaHistoricoGeral();
    var id = dados.id || "HIST-" + new Date().getTime().toString().slice(-8);
    var dataRegistro = dados.dataRegistro || formatarDataHoraSistema_();

    var novaLinha = [
      id,
      dados.tipo || "Geral",
      dados.cliente || "",
      dados.email || "",
      dados.telefone || "",
      dados.titulo || "",
      dados.descricao || "",
      dados.valor || "Aguardando",
      dados.status || "Ativo",
      dataRegistro,
      dados.linkAcao || "",
      dados.observacoes || ""
    ];

    aba.appendRow(novaLinha);
    SpreadsheetApp.flush();

    console.log("Registro historico adicionado:", id);
    return { success: true, id: id };
  } catch (error) {
    console.error("Erro ao registrar historico:", error);
    return { success: false, error: error.toString() };
  }
}

function getHistoricoUsuario(email) {
  try {
    if (!email) return [];

    var aba = getAbaHistoricoGeral();
    var dados = aba.getDataRange().getValues();

    var historico = [];
    var emailLower = email.toString().trim().toLowerCase();

    if (dados.length > 1) {
      for (var i = 1; i < dados.length; i++) {
        var linha = dados[i];
        var emailRegistro = linha[3] ? linha[3].toString().trim().toLowerCase() : "";

        if (emailRegistro === emailLower) {
          historico.push({
            id: linha[0] || "",
            tipo: linha[1] || "",
            cliente: linha[2] || "",
            email: linha[3] || "",
            telefone: linha[4] || "",
            titulo: linha[5] || "",
            descricao: linha[6] || "",
            valor: linha[7] || "Aguardando",
            status: linha[8] || "",
            dataRegistro: linha[9] || "",
            linkAcao: linha[10] || "",
            observacoes: linha[11] || ""
          });
        }
      }
    }

    historico = historico.concat(getHistoricoOrcamentosDoUsuario_(emailLower, historico));
    historico = historico.concat(getHistoricoComprasDoUsuario_(emailLower, historico));

    historico.sort(function(a, b) {
      return normalizarDataHistorico_(b.dataRegistro) - normalizarDataHistorico_(a.dataRegistro);
    });

    console.log("Historico encontrado:", historico.length, "registros");
    return historico;
  } catch (error) {
    console.error("Erro em getHistoricoUsuario:", error);
    return [];
  }
}

function historicoJaTemId_(historico, id) {
  if (!id) return false;

  for (var i = 0; i < historico.length; i++) {
    if (historico[i].id && historico[i].id.toString() === id.toString()) {
      return true;
    }
  }

  return false;
}

function getHistoricoOrcamentosDoUsuario_(emailLower, historicoExistente) {
  try {
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var itens = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailOrcamento = linha[2] ? linha[2].toString().trim().toLowerCase() : "";

      if (emailOrcamento !== emailLower) continue;

      var id = linha[0] || "";
      if (historicoJaTemId_(historicoExistente, id)) continue;

      var valorRaw = linha[9] || "Aguardando";
      var valor = valorRaw;
      if (typeof valorRaw === "number") {
        valor = "R$ " + valorRaw.toFixed(2);
      }

      itens.push({
        id: id,
        tipo: "Orcamento",
        cliente: linha[1] || "",
        email: linha[2] || emailLower,
        telefone: linha[3] || "",
        titulo: linha[5] || "Orcamento",
        tipoServico: linha[5] || "",
        descricao: linha[6] || "",
        imagemURL: linha[7] || "",
        arquivoURL: linha[8] || "",
        valor: valor || "Aguardando",
        status: linha[10] || "Solicitado",
        dataRegistro: linha[4] || "",
        dataValidade: linha[11] || "",
        linkAcao: "tab:orcamentos",
        observacoes: linha[12] || "",
        arquivoNome: linha[13] || ""
      });
    }

    return itens;
  } catch (error) {
    console.error("Erro ao montar historico de orcamentos:", error);
    return [];
  }
}

function getHistoricoComprasDoUsuario_(emailLower, historicoExistente) {
  try {
    var compras = getComprasCliente(emailLower);
    var itens = [];

    for (var i = 0; i < compras.length; i++) {
      var compra = compras[i];

      if (historicoJaTemId_(historicoExistente, compra.id)) continue;

      itens.push({
        id: compra.id || "",
        tipo: "Compra",
        cliente: "",
        email: emailLower,
        telefone: "",
        titulo: montarTituloCompra_(compra.produtos),
        descricao: "Compra de " + (compra.quantidadeTotal || 0) + " item(ns)",
        valor: compra.valorTotal || "",
        status: compra.statusEntrega || compra.statusPagamento || "Pendente",
        dataRegistro: compra.dataCompra || "",
        linkAcao: "tab:historico",
        observacoes: "Pagamento: " + (compra.statusPagamento || "Pendente")
      });
    }

    return itens;
  } catch (error) {
    console.error("Erro ao montar historico de compras:", error);
    return [];
  }
}

function montarTituloCompra_(produtosJson) {
  try {
    var produtos = JSON.parse(produtosJson || "[]");

    if (produtos.length === 1) {
      return produtos[0].nome || "Compra";
    }

    if (produtos.length > 1) {
      return produtos.length + " produtos";
    }
  } catch (error) {}

  return "Compra";
}

function normalizarDataHistorico_(valor) {
  if (!valor) return new Date(0);
  if (Object.prototype.toString.call(valor) === "[object Date]") return valor;

  var texto = valor.toString();
  var partes = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:,\s*(\d{2}):(\d{2})(?::(\d{2}))?)?/);

  if (partes) {
    return new Date(
      Number(partes[3]),
      Number(partes[2]) - 1,
      Number(partes[1]),
      Number(partes[4] || 0),
      Number(partes[5] || 0),
      Number(partes[6] || 0)
    );
  }

  var data = new Date(texto);
  return isNaN(data.getTime()) ? new Date(0) : data;
}

function testarHistoricoNoBrowser() {
  var resultado = getHistoricoUsuario("cawanfernandoanjok@hotmail.com");
  Logger.log("RESULTADO FINAL: " + JSON.stringify(resultado));
  return resultado;
}

function getHistoricoCompletoUsuarioV2(email) {
  try {
    if (!email) return [];

    var emailLower = email.toString().trim().toLowerCase();
    var historico = [];
    var ids = {};

    try {
      var abaHistorico = getAbaHistoricoGeral();
      var dadosHistorico = abaHistorico.getDataRange().getValues();

      for (var h = 1; h < dadosHistorico.length; h++) {
        var linhaHist = dadosHistorico[h];
        var emailHist = linhaHist[3] ? linhaHist[3].toString().trim().toLowerCase() : "";
        if (emailHist !== emailLower) continue;

        var itemHist = {
          id: linhaHist[0] || "",
          tipo: linhaHist[1] || "Geral",
          cliente: linhaHist[2] || "",
          email: linhaHist[3] || email,
          telefone: linhaHist[4] || "",
          titulo: linhaHist[5] || "",
          descricao: linhaHist[6] || "",
          valor: linhaHist[7] || "",
          status: linhaHist[8] || "",
          dataRegistro: linhaHist[9] || "",
          linkAcao: linhaHist[10] || "",
          observacoes: linhaHist[11] || ""
        };
        historico.push(itemHist);
        ids[itemHist.id] = true;
      }
    } catch (eHist) {
      console.error("HistoricoGeral ignorado:", eHist);
    }

    try {
      var abaOrc = getAbaOrcamentos();
      var dadosOrc = abaOrc.getDataRange().getValues();

      for (var o = 1; o < dadosOrc.length; o++) {
        var linhaOrc = dadosOrc[o];
        var emailOrc = linhaOrc[2] ? linhaOrc[2].toString().trim().toLowerCase() : "";
        if (emailOrc !== emailLower) continue;

        var idOrc = linhaOrc[0] || "";
        if (ids[idOrc]) continue;

        var valorOrc = linhaOrc[9] || "Aguardando";
        if (typeof valorOrc === "number") valorOrc = "R$ " + valorOrc.toFixed(2);

        historico.push({
          id: idOrc,
          tipo: "Orcamento",
          cliente: linhaOrc[1] || "",
          email: linhaOrc[2] || email,
          telefone: linhaOrc[3] || "",
          titulo: linhaOrc[5] || "Orcamento",
          tipoServico: linhaOrc[5] || "",
          descricao: linhaOrc[6] || "",
          imagemURL: linhaOrc[7] || "",
          arquivoURL: linhaOrc[8] || "",
          valor: valorOrc,
          status: linhaOrc[10] || "Solicitado",
          dataRegistro: linhaOrc[4] || "",
          dataValidade: linhaOrc[11] || "",
          linkAcao: "tab:historico",
          observacoes: linhaOrc[12] || "",
          arquivoNome: linhaOrc[13] || ""
        });
        ids[idOrc] = true;
      }
    } catch (eOrc) {
      console.error("Orcamentos ignorados:", eOrc);
    }

    try {
      var abaComp = getAbaCompras();
      var dadosComp = abaComp.getDataRange().getValues();

      for (var c = 1; c < dadosComp.length; c++) {
        var linhaComp = dadosComp[c];
        var emailComp = linhaComp[2] ? linhaComp[2].toString().trim().toLowerCase() : "";
        if (emailComp !== emailLower) continue;

        var idComp = linhaComp[0] || "";
        if (ids[idComp]) continue;

        historico.push({
          id: idComp,
          tipo: "Compra",
          cliente: linhaComp[1] || "",
          email: linhaComp[2] || email,
          telefone: "",
          titulo: montarTituloCompra_(linhaComp[4]),
          descricao: "Compra de " + (linhaComp[5] || 0) + " item(ns)",
          valor: linhaComp[6] ? "R$ " + (parseFloat(linhaComp[6]) || 0).toFixed(2) : "",
          status: linhaComp[8] || linhaComp[7] || "Processando",
          dataRegistro: linhaComp[3] || "",
          linkAcao: "tab:historico",
          observacoes: "Pagamento: " + (linhaComp[7] || "Pendente")
        });
        ids[idComp] = true;
      }
    } catch (eComp) {
      console.error("Compras ignoradas:", eComp);
    }

    historico.sort(function(a, b) {
      return normalizarDataHistorico_(b.dataRegistro) - normalizarDataHistorico_(a.dataRegistro);
    });

    return historico;
  } catch (error) {
    console.error("Erro em getHistoricoCompletoUsuarioV2:", error);
    return [];
  }
}

function getCentralAdminV2() {
  return {
    orcamentosPendentes: getOrcamentosPendentesV2_(),
    comprasPendentes: getComprasPendentesV2_(),
    ordensAtivas: []
  };
}

function getOrcamentosPendentesV2_() {
  try {
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var itens = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var status = linha[10] ? linha[10].toString() : "Solicitado";
      var statusNormalizado = status.toLowerCase();

      if (statusNormalizado === "aprovado" || statusNormalizado === "cancelado" || statusNormalizado === "recusado") continue;

      itens.push({
        id: linha[0] || "",
        cliente: linha[1] || "",
        email: linha[2] || "",
        telefone: linha[3] || "",
        dataSolicitacao: linha[4] || "",
        tipoServico: linha[5] || "",
        descricao: linha[6] || "",
        valor: linha[9] || "Aguardando",
        status: status,
        observacoes: linha[12] || ""
      });
    }

    return itens;
  } catch (error) {
    console.error("Erro em getOrcamentosPendentesV2_:", error);
    return [];
  }
}

function getComprasPendentesV2_() {
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
        statusEntrega: statusEntrega
      });
    }

    return compras;
  } catch (error) {
    console.error("Erro em getComprasPendentesV2_:", error);
    return [];
  }
}

function getRelatoriosAdminV2() {
  var orcamentos = getTodosOrcamentosV2_();
  var compras = getTodasComprasV2_();
  var usuarios = getTodosUsuarios();

  return {
    totalClientes: usuarios.filter(function(u) { return u.tipo !== "Administrador"; }).length,
    totalAdmins: usuarios.filter(function(u) { return u.tipo === "Administrador"; }).length,
    totalOrcamentos: orcamentos.length,
    orcamentosPendentes: getOrcamentosPendentesV2_().length,
    totalCompras: compras.length,
    comprasPendentes: getComprasPendentesV2_().length,
    faturamentoCompras: compras.reduce(function(total, compra) {
      return total + (parseFloat(compra.valorTotal) || 0);
    }, 0)
  };
}

function atualizarStatusCompraV2(idCompra, statusEntrega) {
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

function getTodosOrcamentosV2_() {
  try {
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var itens = [];
    for (var i = 1; i < dados.length; i++) {
      itens.push({ id: dados[i][0], status: dados[i][10] || "Solicitado" });
    }
    return itens;
  } catch (error) {
    return [];
  }
}

function getTodasComprasV2_() {
  try {
    var aba = getAbaCompras();
    var dados = aba.getDataRange().getValues();
    var itens = [];
    for (var i = 1; i < dados.length; i++) {
      itens.push({
        id: dados[i][0] || "",
        valorTotal: dados[i][6] || 0,
        statusEntrega: dados[i][8] || "Processando"
      });
    }
    return itens;
  } catch (error) {
    return [];
  }
}

function enviarNotificacaoV2(dados) {
  try {
    var aba = getAbaNotificacoes();
    var id = "NOT-" + new Date().getTime().toString().slice(-8);
    var dataEnvio = formatarDataHoraSistema_();

    aba.appendRow([
      id,
      dados.titulo || "",
      dados.mensagem || "",
      dados.tipo || "info",
      dataEnvio,
      dados.destinatarios || "todos",
      "",
      Session.getActiveUser().getEmail(),
      dados.acao || "",
      dados.icone || "fas fa-bell"
    ]);
    SpreadsheetApp.flush();

    if (dados.enviarEmail === true) {
      try {
        enviarNotificacaoPorEmail(dados);
      } catch (emailError) {
        console.error("Notificacao salva, mas email nao enviado:", emailError);
      }
    }

    return { success: true, id: id, message: "Notificacao enviada com sucesso." };
  } catch (error) {
    return { success: false, message: "Erro ao enviar notificacao: " + error.toString() };
  }
}

function getNotificacoesUsuarioV2(email) {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var emailLower = email ? email.toString().trim().toLowerCase() : "";
    var tipoUsuario = getTipoUsuarioPorEmailV2_(emailLower);
    var notificacoes = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var destinatarios = linha[5] ? linha[5].toString().toLowerCase() : "todos";
      var lidas = linha[6] ? linha[6].toString().toLowerCase() : "";
      var deveReceber = destinatarios === "todos" ||
        (destinatarios === "clientes" && tipoUsuario !== "Administrador") ||
        (destinatarios === "admin" && tipoUsuario === "Administrador") ||
        destinatarios.indexOf(emailLower) !== -1;

      if (!deveReceber) continue;

      notificacoes.push({
        id: linha[0] || "",
        titulo: linha[1] || "",
        mensagem: linha[2] || "",
        tipo: linha[3] || "info",
        dataEnvio: linha[4] || "",
        acao: linha[8] || "",
        icone: linha[9] || "fas fa-bell",
        lida: lidas.indexOf(emailLower) !== -1
      });
    }

    notificacoes.sort(function(a, b) {
      if (a.lida && !b.lida) return 1;
      if (!a.lida && b.lida) return -1;
      return normalizarDataHistorico_(b.dataEnvio) - normalizarDataHistorico_(a.dataEnvio);
    });

    return notificacoes;
  } catch (error) {
    console.error("Erro em getNotificacoesUsuarioV2:", error);
    return [];
  }
}

function getAllNotificacoesV2() {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var notificacoes = [];

    for (var i = 1; i < dados.length; i++) {
      notificacoes.push({
        id: dados[i][0] || "",
        titulo: dados[i][1] || "",
        mensagem: dados[i][2] || "",
        tipo: dados[i][3] || "",
        dataEnvio: dados[i][4] || "",
        destinatarios: dados[i][5] || "",
        criadoPor: dados[i][7] || "",
        acao: dados[i][8] || "",
        icone: dados[i][9] || "fas fa-bell"
      });
    }

    return notificacoes;
  } catch (error) {
    console.error("Erro em getAllNotificacoesV2:", error);
    return [];
  }
}

function marcarNotificacaoLidaV2(idNotificacao, email) {
  return atualizarLeituraNotificacaoV2_(idNotificacao, email);
}

function marcarTodasNotificacoesLidasV2(email) {
  try {
    var notificacoes = getNotificacoesUsuarioV2(email);
    for (var i = 0; i < notificacoes.length; i++) {
      atualizarLeituraNotificacaoV2_(notificacoes[i].id, email);
    }
    return { success: true };
  } catch (error) {
    return { success: false };
  }
}

function atualizarLeituraNotificacaoV2_(idNotificacao, email) {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var emailLower = email ? email.toString().trim().toLowerCase() : "";

    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === idNotificacao) {
        var lidas = dados[i][6] ? dados[i][6].toString() : "";
        if (lidas.toLowerCase().indexOf(emailLower) === -1) {
          aba.getRange(i + 1, 7).setValue(lidas ? lidas + "," + emailLower : emailLower);
          SpreadsheetApp.flush();
        }
        return { success: true };
      }
    }

    return { success: false };
  } catch (error) {
    return { success: false };
  }
}

function getTipoUsuarioPorEmailV2_(emailLower) {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      var emailUsuario = dados[i][0] ? dados[i][0].toString().trim().toLowerCase() : "";
      if (emailUsuario === emailLower) {
        return isUsuarioAdmin_(dados[i]) ? "Administrador" : "Cliente";
      }
    }
  } catch (error) {}
  return "Cliente";
}

function diagnosticoSistemaV2(email) {
  var emailLower = email ? email.toString().trim().toLowerCase() : "";
  var resultado = {
    email: emailLower,
    historicoQuantidade: 0,
    orcamentosTotal: 0,
    orcamentosDoUsuario: 0,
    comprasTotal: 0,
    comprasDoUsuario: 0,
    notificacoesTotal: 0,
    notificacoesDoUsuario: 0,
    erros: []
  };

  try {
    resultado.historicoQuantidade = getHistoricoCompletoUsuarioV2(emailLower).length;
  } catch (eHist) {
    resultado.erros.push("historico: " + eHist.toString());
  }

  try {
    var orc = getAbaOrcamentos().getDataRange().getValues();
    resultado.orcamentosTotal = Math.max(orc.length - 1, 0);
    for (var i = 1; i < orc.length; i++) {
      var emailOrc = orc[i][2] ? orc[i][2].toString().trim().toLowerCase() : "";
      if (emailOrc === emailLower) resultado.orcamentosDoUsuario++;
    }
  } catch (eOrc) {
    resultado.erros.push("orcamentos: " + eOrc.toString());
  }

  try {
    var comp = getAbaCompras().getDataRange().getValues();
    resultado.comprasTotal = Math.max(comp.length - 1, 0);
    for (var j = 1; j < comp.length; j++) {
      var emailComp = comp[j][2] ? comp[j][2].toString().trim().toLowerCase() : "";
      if (emailComp === emailLower) resultado.comprasDoUsuario++;
    }
  } catch (eComp) {
    resultado.erros.push("compras: " + eComp.toString());
  }

  try {
    resultado.notificacoesTotal = getAllNotificacoesV2().length;
    resultado.notificacoesDoUsuario = getNotificacoesUsuarioV2(emailLower).length;
  } catch (eNotif) {
    resultado.erros.push("notificacoes: " + eNotif.toString());
  }

  return resultado;
}

function normalizarTextoV3_(valor) {
  if (valor === null || valor === undefined) return "";
  return valor.toString().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function valorParaTextoV3_(valor) {
  if (valor === null || valor === undefined) return "";
  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return formatarDataHoraSistema_(valor);
  }
  return valor.toString();
}

function parseValorV3_(valor) {
  if (typeof valor === "number") return valor;
  if (!valor) return 0;
  var texto = valor.toString()
    .replace(/\s/g, "")
    .replace("R$", "");
  if (texto.indexOf(",") !== -1) {
    texto = texto.replace(/\./g, "").replace(",", ".");
  }
  var numero = parseFloat(texto);
  return isNaN(numero) ? 0 : numero;
}

function formatarValorV3_(valor) {
  var numero = parseValorV3_(valor);
  return "R$ " + numero.toFixed(2).replace(".", ",");
}

function statusPagoV3_(status) {
  var s = normalizarTextoV3_(status);
  return s === "pago" || s === "paga" || s === "quitado" || s === "quitada";
}

function statusCanceladoV3_(status) {
  var s = normalizarTextoV3_(status);
  return s === "cancelado" || s === "cancelada" || s === "recusado" || s === "recusada";
}

function statusOrcamentoPendenteV3_(status) {
  var s = normalizarTextoV3_(status || "Solicitado");
  return s !== "aprovado" && s !== "cancelado" && s !== "recusado" && s !== "concluido";
}

function normalizarItemHistoricoV3_(item) {
  item = item || {};
  return {
    id: valorParaTextoV3_(item.id),
    tipo: valorParaTextoV3_(item.tipo || "Geral"),
    cliente: valorParaTextoV3_(item.cliente),
    email: valorParaTextoV3_(item.email),
    telefone: valorParaTextoV3_(item.telefone),
    titulo: valorParaTextoV3_(item.titulo),
    descricao: valorParaTextoV3_(item.descricao),
    valor: valorParaTextoV3_(item.valor),
    status: valorParaTextoV3_(item.status),
    dataRegistro: valorParaTextoV3_(item.dataRegistro),
    linkAcao: valorParaTextoV3_(item.linkAcao),
    observacoes: valorParaTextoV3_(item.observacoes),
    tipoServico: valorParaTextoV3_(item.tipoServico),
    imagemURL: valorParaTextoV3_(item.imagemURL),
    arquivoURL: valorParaTextoV3_(item.arquivoURL),
    arquivoNome: valorParaTextoV3_(item.arquivoNome),
    dataValidade: valorParaTextoV3_(item.dataValidade)
  };
}

function getHistoricoServicosUsuarioV3_(emailLower, idsExistentes) {
  try {
    var aba = getAbaServicos();
    var dados = aba.getDataRange().getValues();
    var itens = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailServico = linha[2] ? linha[2].toString().trim().toLowerCase() : "";
      if (emailServico !== emailLower) continue;

      var id = valorParaTextoV3_(linha[0]);
      if (idsExistentes[id]) continue;

      itens.push({
        id: id,
        tipo: "Servico",
        cliente: valorParaTextoV3_(linha[1]),
        email: valorParaTextoV3_(linha[2]),
        telefone: "",
        titulo: valorParaTextoV3_(linha[3] || "Servico"),
        descricao: valorParaTextoV3_(linha[8]),
        valor: linha[6] ? formatarValorV3_(linha[6]) : "A definir",
        status: valorParaTextoV3_(linha[7] || "Pendente"),
        dataRegistro: valorParaTextoV3_(linha[4]),
        linkAcao: "tab:historico",
        observacoes: linha[9] ? "Orcamento: " + valorParaTextoV3_(linha[9]) : ""
      });
      idsExistentes[id] = true;
    }

    return itens;
  } catch (error) {
    console.error("Erro em getHistoricoServicosUsuarioV3_:", error);
    return [];
  }
}

function getHistoricoCompletoUsuarioV3(email) {
  var emailLower = email ? email.toString().trim().toLowerCase() : "";
  var erros = [];
  var historico = [];
  var ids = {};

  try {
    historico = getHistoricoCompletoUsuarioV2(emailLower) || [];
  } catch (errorV2) {
    erros.push("historicoV2: " + errorV2.toString());
    historico = [];
  }

  historico = historico.map(function(item) {
    var normalizado = normalizarItemHistoricoV3_(item);
    if (normalizado.id) ids[normalizado.id] = true;
    return normalizado;
  });

  historico = historico.concat(getHistoricoServicosUsuarioV3_(emailLower, ids).map(normalizarItemHistoricoV3_));
  historico.sort(function(a, b) {
    return normalizarDataHistorico_(b.dataRegistro) - normalizarDataHistorico_(a.dataRegistro);
  });

  return {
    success: true,
    email: emailLower,
    itens: historico,
    total: historico.length,
    erros: erros
  };
}

function getUsuariosAtivosV3_() {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    var usuarios = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var email = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      if (!email) continue;

      var status = linha[6] ? linha[6].toString() : "Ativo";
      var statusNorm = normalizarTextoV3_(status);
      if (statusNorm === "inativo" || statusNorm === "bloqueado" || statusNorm === "cancelado") continue;

      usuarios.push({
        email: email,
        usuario: valorParaTextoV3_(linha[2]),
        telefone: valorParaTextoV3_(linha[3]),
        tipo: isUsuarioAdmin_(linha) ? "Administrador" : "Cliente",
        dataCadastro: valorParaTextoV3_(linha[5]),
        status: status || "Ativo",
        endereco: valorParaTextoV3_(linha[11]),
        bairro: valorParaTextoV3_(linha[12])
      });
    }

    return usuarios;
  } catch (error) {
    console.error("Erro em getUsuariosAtivosV3_:", error);
    return [];
  }
}

function getTelefonesAutorizadosV3_() {
  try {
    var aba = getAbaTelefonesAutorizados();
    var dados = aba.getDataRange().getValues();
    var telefones = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      if (!linha[0]) continue;
      telefones.push({
        telefone: valorParaTextoV3_(linha[0]),
        nome: valorParaTextoV3_(linha[1]),
        dataAutorizacao: valorParaTextoV3_(linha[2]),
        status: valorParaTextoV3_(linha[3] || "Ativo"),
        usado: valorParaTextoV3_(linha[4] || "Nao"),
        dataUso: valorParaTextoV3_(linha[5])
      });
    }

    return telefones;
  } catch (error) {
    console.error("Erro em getTelefonesAutorizadosV3_:", error);
    return [];
  }
}

function getAdminCadastrosV3() {
  var usuarios = getUsuariosAtivosV3_();
  var clientes = usuarios.filter(function(u) { return u.tipo !== "Administrador"; });
  var admins = usuarios.filter(function(u) { return u.tipo === "Administrador"; });
  var telefones = getTelefonesAutorizadosV3_();

  return {
    success: true,
    clientes: clientes,
    admins: admins,
    telefones: telefones,
    totalClientes: clientes.length,
    totalAdmins: admins.length,
    totalTelefones: telefones.length
  };
}

function getTodosOrcamentosV3_() {
  try {
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var itens = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      if (!linha[0]) continue;
      itens.push({
        id: valorParaTextoV3_(linha[0]),
        cliente: valorParaTextoV3_(linha[1]),
        email: valorParaTextoV3_(linha[2]),
        telefone: valorParaTextoV3_(linha[3]),
        dataSolicitacao: valorParaTextoV3_(linha[4]),
        tipoServico: valorParaTextoV3_(linha[5]),
        descricao: valorParaTextoV3_(linha[6]),
        imagemURL: valorParaTextoV3_(linha[7]),
        arquivoURL: valorParaTextoV3_(linha[8]),
        valor: linha[9] && linha[9] !== "Aguardando" ? formatarValorV3_(linha[9]) : "Aguardando",
        valorNumero: parseValorV3_(linha[9]),
        status: valorParaTextoV3_(linha[10] || "Solicitado"),
        dataValidade: valorParaTextoV3_(linha[11]),
        observacoes: valorParaTextoV3_(linha[12]),
        arquivoNome: valorParaTextoV3_(linha[13])
      });
    }

    itens.sort(function(a, b) {
      return normalizarDataHistorico_(b.dataSolicitacao) - normalizarDataHistorico_(a.dataSolicitacao);
    });

    return itens;
  } catch (error) {
    console.error("Erro em getTodosOrcamentosV3_:", error);
    return [];
  }
}

function getTodasComprasV3_() {
  try {
    var aba = getAbaCompras();
    var dados = aba.getDataRange().getValues();
    var compras = [];
    var produtos = getProdutos();
    var produtosPorId = {};
    for (var p = 0; p < produtos.length; p++) {
      produtosPorId[produtos[p].id.toString()] = produtos[p];
    }

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      if (!linha[0]) continue;
      var produtosLista = parseProdutosCompraV3_(linha[4]);
      for (var j = 0; j < produtosLista.length; j++) {
        var produtoRef = produtosPorId[produtosLista[j].id ? produtosLista[j].id.toString() : ""];
        if (produtoRef) {
          produtosLista[j].imagemURL = produtosLista[j].imagemURL || produtoRef.imagemURL || "";
          produtosLista[j].categoria = produtosLista[j].categoria || produtoRef.categoria || "";
        }
      }
      compras.push({
        id: valorParaTextoV3_(linha[0]),
        cliente: valorParaTextoV3_(linha[1]),
        email: valorParaTextoV3_(linha[2]),
        dataCompra: valorParaTextoV3_(linha[3]),
        produtos: valorParaTextoV3_(linha[4]),
        produtosLista: produtosLista,
        quantidadeTotal: parseInt(linha[5], 10) || 0,
        valorTotal: parseValorV3_(linha[6]),
        valorTotalFormatado: formatarValorV3_(linha[6]),
        statusPagamento: valorParaTextoV3_(linha[7] || "Pendente"),
        statusEntrega: valorParaTextoV3_(linha[8] || "Processando"),
        observacoesAdmin: valorParaTextoV3_(linha[9])
      });
    }

    compras.sort(function(a, b) {
      return normalizarDataHistorico_(b.dataCompra) - normalizarDataHistorico_(a.dataCompra);
    });

    return compras;
  } catch (error) {
    console.error("Erro em getTodasComprasV3_:", error);
    return [];
  }
}

function parseProdutosCompraV3_(produtosJson) {
  try {
    var lista = JSON.parse(produtosJson || "[]");
    if (Object.prototype.toString.call(lista) === "[object Array]") return lista;
  } catch (error) {}
  return [];
}

function getPagamentosPendentesV3_() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("Pagamentos");
    if (!aba) return [];
    var dados = aba.getDataRange().getValues();
    var pagamentos = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var status = valorParaTextoV3_(linha[6] || "Pendente");
      if (statusPagoV3_(status) || statusCanceladoV3_(status)) continue;

      pagamentos.push({
        id: valorParaTextoV3_(linha[0]),
        cliente: valorParaTextoV3_(linha[1]),
        email: valorParaTextoV3_(linha[2]),
        tipo: valorParaTextoV3_(linha[3] || "Servico"),
        valor: parseValorV3_(linha[4]),
        valorFormatado: formatarValorV3_(linha[4]),
        vencimento: valorParaTextoV3_(linha[5]),
        status: status
      });
    }

    return pagamentos;
  } catch (error) {
    console.error("Erro em getPagamentosPendentesV3_:", error);
    return [];
  }
}

function getServicosNaoPagosV3_() {
  try {
    var aba = getAbaServicos();
    var dados = aba.getDataRange().getValues();
    var servicos = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var statusNorm = normalizarTextoV3_(linha[7] || "");
      var statusIndicaPagamento = statusNorm.indexOf("pagamento") !== -1 || statusNorm.indexOf("nao pago") !== -1 || statusNorm.indexOf("pendente financeiro") !== -1;
      if (!statusIndicaPagamento) continue;

      servicos.push({
        id: valorParaTextoV3_(linha[0]),
        cliente: valorParaTextoV3_(linha[1]),
        email: valorParaTextoV3_(linha[2]),
        tipo: valorParaTextoV3_(linha[3]),
        valor: parseValorV3_(linha[6]),
        valorFormatado: formatarValorV3_(linha[6]),
        status: valorParaTextoV3_(linha[7]),
        descricao: valorParaTextoV3_(linha[8])
      });
    }

    return servicos;
  } catch (error) {
    console.error("Erro em getServicosNaoPagosV3_:", error);
    return [];
  }
}

function calcularAReceberV3_() {
  var compras = getTodasComprasV3_().filter(function(compra) {
    return !statusPagoV3_(compra.statusPagamento) && !statusCanceladoV3_(compra.statusEntrega);
  });
  var pagamentos = getPagamentosPendentesV3_();
  var servicosSemPagamento = getServicosNaoPagosV3_();

  var totalCompras = compras.reduce(function(total, compra) {
    return total + (parseFloat(compra.valorTotal) || 0);
  }, 0);

  var totalPagamentos = pagamentos.reduce(function(total, pagamento) {
    return total + (parseFloat(pagamento.valor) || 0);
  }, 0);

  var totalServicosSemPagamento = servicosSemPagamento.reduce(function(total, servico) {
    return total + (parseFloat(servico.valor) || 0);
  }, 0);

  return {
    total: totalCompras + totalPagamentos + totalServicosSemPagamento,
    totalFormatado: formatarValorV3_(totalCompras + totalPagamentos + totalServicosSemPagamento),
    comprasNaoPagas: compras,
    pagamentosPendentes: pagamentos,
    servicosNaoPagos: servicosSemPagamento,
    totalComprasNaoPagas: totalCompras,
    totalServicosNaoPagos: totalPagamentos + totalServicosSemPagamento
  };
}

function getCentralAdminV3() {
  var erros = [];
  var orcamentos = [];
  var compras = [];

  try {
    orcamentos = getTodosOrcamentosV3_().filter(function(orcamento) {
      return statusOrcamentoPendenteV3_(orcamento.status);
    });
  } catch (eOrc) {
    erros.push("orcamentos: " + eOrc.toString());
  }

  try {
    compras = getTodasComprasV3_().filter(function(compra) {
      return !statusCanceladoV3_(compra.statusEntrega) && normalizarTextoV3_(compra.statusEntrega) !== "entregue";
    });
  } catch (eComp) {
    erros.push("compras: " + eComp.toString());
  }

  return {
    success: true,
    orcamentosPendentes: orcamentos,
    comprasPendentes: compras,
    ordensAtivas: [],
    erros: erros
  };
}

function getDashboardAdminV3() {
  var cadastros = getAdminCadastrosV3() || {};
  var central = getCentralAdminV3() || {};
  var aReceber = calcularAReceberV3_() || {};
  cadastros.clientes = cadastros.clientes || [];
  cadastros.admins = cadastros.admins || [];
  cadastros.telefones = cadastros.telefones || [];
  cadastros.totalClientes = cadastros.totalClientes || cadastros.clientes.length || 0;
  cadastros.totalAdmins = cadastros.totalAdmins || cadastros.admins.length || 0;
  cadastros.totalTelefones = cadastros.totalTelefones || cadastros.telefones.length || 0;
  central.orcamentosPendentes = central.orcamentosPendentes || [];
  central.comprasPendentes = central.comprasPendentes || [];
  var ordens = [];
  var tecnicos = [];
  var agenda = [];

  try {
    if (typeof getAllOrdensServico === "function") ordens = getAllOrdensServico() || [];
  } catch (errorOrdens) {
    ordens = [];
  }

  try {
    if (typeof getTecnicosAtivos === "function") tecnicos = getTecnicosAtivos() || [];
  } catch (errorTecnicos) {
    tecnicos = [];
  }

  try {
    if (typeof getProximosAgendamentos === "function") agenda = getProximosAgendamentos() || [];
  } catch (errorAgenda) {
    agenda = [];
  }

  var ordensAtivas = ordens.filter(function(ordem) {
    var status = normalizarTextoV3_(ordem.status);
    return status === "agendado" || status === "em andamento" || status === "pendente";
  });

  return {
    success: true,
    clientesAtivos: cadastros.totalClientes,
    telefonesAutorizados: cadastros.totalTelefones,
    ordensAtivas: ordensAtivas.length + central.orcamentosPendentes.length + central.comprasPendentes.length,
    tecnicosAtivos: tecnicos.length,
    aReceber: aReceber,
    ordensRecentes: ordens.slice(0, 5),
    tecnicos: tecnicos,
    agenda: agenda,
    cadastros: cadastros,
    central: central
  };
}

function getRelatoriosAdminV3() {
  var cadastros = getAdminCadastrosV3() || {};
  cadastros.clientes = cadastros.clientes || [];
  cadastros.admins = cadastros.admins || [];
  cadastros.telefones = cadastros.telefones || [];
  cadastros.totalClientes = cadastros.totalClientes || cadastros.clientes.length || 0;
  cadastros.totalAdmins = cadastros.totalAdmins || cadastros.admins.length || 0;
  var orcamentos = getTodosOrcamentosV3_() || [];
  var compras = getTodasComprasV3_() || [];
  var orcamentosPendentes = orcamentos.filter(function(orcamento) {
    return statusOrcamentoPendenteV3_(orcamento.status);
  });
  var comprasPendentes = compras.filter(function(compra) {
    return !statusCanceladoV3_(compra.statusEntrega) && normalizarTextoV3_(compra.statusEntrega) !== "entregue";
  });
  var aReceber = calcularAReceberV3_();
  var faturamentoCompras = compras.reduce(function(total, compra) {
    return total + (parseFloat(compra.valorTotal) || 0);
  }, 0);

  return {
    success: true,
    totalClientes: cadastros.totalClientes,
    totalAdmins: cadastros.totalAdmins,
    totalOrcamentos: orcamentos.length,
    orcamentosPendentes: orcamentosPendentes.length,
    totalCompras: compras.length,
    comprasPendentes: comprasPendentes.length,
    faturamentoCompras: faturamentoCompras,
    aReceber: aReceber,
    clientes: cadastros.clientes,
    admins: cadastros.admins,
    telefones: cadastros.telefones,
    orcamentos: orcamentos,
    orcamentosPendentesLista: orcamentosPendentes,
    compras: compras,
    comprasPendentesLista: comprasPendentes,
    vendas: compras
  };
}

function getNotificacoesUsuarioV3(email) {
  var itens = [];
  var erros = [];
  try {
    itens = getNotificacoesUsuarioV2(email) || [];
  } catch (error) {
    erros.push(error.toString());
    itens = [];
  }

  itens = itens.map(function(notif) {
    return {
      id: valorParaTextoV3_(notif.id),
      titulo: valorParaTextoV3_(notif.titulo),
      mensagem: valorParaTextoV3_(notif.mensagem),
      tipo: valorParaTextoV3_(notif.tipo || "info"),
      dataEnvio: valorParaTextoV3_(notif.dataEnvio),
      acao: valorParaTextoV3_(notif.acao),
      icone: valorParaTextoV3_(notif.icone || "fas fa-bell"),
      lida: notif.lida === true
    };
  });

  return {
    success: true,
    itens: itens,
    total: itens.length,
    erros: erros
  };
}

function getAllNotificacoesV3() {
  var itens = [];
  var erros = [];
  try {
    itens = getAllNotificacoesV2() || [];
  } catch (error) {
    erros.push(error.toString());
    itens = [];
  }

  itens = itens.map(function(notif) {
    return {
      id: valorParaTextoV3_(notif.id),
      titulo: valorParaTextoV3_(notif.titulo),
      mensagem: valorParaTextoV3_(notif.mensagem),
      tipo: valorParaTextoV3_(notif.tipo || "info"),
      dataEnvio: valorParaTextoV3_(notif.dataEnvio),
      destinatarios: valorParaTextoV3_(notif.destinatarios),
      criadoPor: valorParaTextoV3_(notif.criadoPor),
      acao: valorParaTextoV3_(notif.acao),
      icone: valorParaTextoV3_(notif.icone || "fas fa-bell")
    };
  });

  return {
    success: true,
    itens: itens,
    total: itens.length,
    erros: erros
  };
}

function excluirNotificacaoV3(idNotificacao) {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var id = idNotificacao ? idNotificacao.toString() : "";

    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] && dados[i][0].toString() === id) {
        aba.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: "Notificacao removida." };
      }
    }

    return { success: false, message: "Notificacao nao encontrada." };
  } catch (error) {
    return { success: false, message: "Erro ao remover notificacao: " + error.toString() };
  }
}

function enviarNotificacaoV3(dados) {
  var resultado = enviarNotificacaoV2(dados || {});
  if (!resultado) {
    return { success: false, message: "Falha ao enviar notificacao." };
  }
  resultado.notificacoes = getAllNotificacoesV3().itens;
  return resultado;
}

function marcarNotificacaoLidaV3(idNotificacao, email) {
  return marcarNotificacaoLidaV2(idNotificacao, email);
}

function marcarTodasNotificacoesLidasV3(email) {
  return marcarTodasNotificacoesLidasV2(email);
}

function getPerfilUsuarioV3(email) {
  try {
    var emailLower = email ? email.toString().trim().toLowerCase() : "";
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailUsuario = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      if (emailUsuario !== emailLower) continue;

      return {
        success: true,
        usuario: valorParaTextoV3_(linha[2]),
        email: emailUsuario,
        telefone: valorParaTextoV3_(linha[3]),
        tipo: isUsuarioAdmin_(linha) ? "Administrador" : valorParaTextoV3_(linha[4] || "Cliente"),
        dataCadastro: valorParaTextoV3_(linha[5]),
        status: valorParaTextoV3_(linha[6] || "Ativo"),
        avatarURL: valorParaTextoV3_(linha[9]),
        endereco: valorParaTextoV3_(linha[11]),
        bairro: valorParaTextoV3_(linha[12])
      };
    }

    return { success: false, message: "Usuario nao encontrado." };
  } catch (error) {
    return { success: false, message: "Erro ao carregar perfil: " + error.toString() };
  }
}

function atualizarEnderecoUsuarioV3(email, endereco, bairro) {
  try {
    var emailLower = email ? email.toString().trim().toLowerCase() : "";
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      var emailUsuario = dados[i][0] ? dados[i][0].toString().trim().toLowerCase() : "";
      if (emailUsuario !== emailLower) continue;

      aba.getRange(i + 1, 12).setValue(endereco || "");
      aba.getRange(i + 1, 13).setValue(bairro || "");
      SpreadsheetApp.flush();
      return { success: true, message: "Endereco atualizado." };
    }

    return { success: false, message: "Usuario nao encontrado." };
  } catch (error) {
    return { success: false, message: "Erro ao atualizar endereco: " + error.toString() };
  }
}

function getAbaChatV3_() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("MensagensChat");
  var cabecalhos = [
    "ID", "DataHora", "ClienteEmail", "ClienteNome", "RemetenteEmail",
    "RemetenteNome", "RemetenteTipo", "Mensagem", "LidaCliente", "LidaAdmin"
  ];

  if (!aba) {
    aba = planilha.insertSheet("MensagensChat");
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    aba.setFrozenRows(1);
    return aba;
  }

  var atuais = aba.getRange(1, 1, 1, cabecalhos.length).getValues()[0];
  var precisa = false;
  for (var i = 0; i < cabecalhos.length; i++) {
    if (!atuais[i]) {
      atuais[i] = cabecalhos[i];
      precisa = true;
    }
  }
  if (precisa) aba.getRange(1, 1, 1, cabecalhos.length).setValues([atuais]);
  return aba;
}

function enviarMensagemChatV3(dados) {
  try {
    dados = dados || {};
    var mensagem = dados.mensagem ? dados.mensagem.toString().trim() : "";
    if (!mensagem) return { success: false, message: "Digite uma mensagem." };

    var remetenteTipo = dados.remetenteTipo || "Cliente";
    var clienteEmail = remetenteTipo === "Administrador" ? dados.clienteEmail : dados.email;
    clienteEmail = clienteEmail ? clienteEmail.toString().trim().toLowerCase() : "";
    if (!clienteEmail) return { success: false, message: "Cliente nao identificado." };

    var clientePerfil = getPerfilUsuarioV3(clienteEmail);
    var clienteNome = clientePerfil && clientePerfil.success ? clientePerfil.usuario : (dados.clienteNome || clienteEmail);
    var remetenteEmail = dados.email ? dados.email.toString().trim().toLowerCase() : "";
    var remetenteNome = dados.usuario || remetenteEmail;
    var id = "MSG-" + new Date().getTime().toString().slice(-9);
    var aba = getAbaChatV3_();

    aba.appendRow([
      id,
      formatarDataHoraSistema_(),
      clienteEmail,
      clienteNome,
      remetenteEmail,
      remetenteNome,
      remetenteTipo,
      mensagem,
      remetenteTipo === "Cliente" ? "Sim" : "",
      remetenteTipo === "Administrador" ? "Sim" : ""
    ]);
    SpreadsheetApp.flush();

    if (remetenteTipo === "Administrador") {
      try {
        enviarNotificacaoV3({
          titulo: "Nova resposta da equipe",
          mensagem: "A equipe JW Piscinas respondeu sua mensagem no suporte.",
          tipo: "info",
          destinatarios: clienteEmail,
          acao: "chat:open",
          icone: "fas fa-comments"
        });
      } catch (notifError) {}
    }

    return { success: true, message: "Mensagem enviada.", id: id };
  } catch (error) {
    return { success: false, message: "Erro ao enviar mensagem: " + error.toString() };
  }
}

function getMensagensConversaV3(emailCliente) {
  try {
    var emailLower = emailCliente ? emailCliente.toString().trim().toLowerCase() : "";
    var aba = getAbaChatV3_();
    var dados = aba.getDataRange().getValues();
    var mensagens = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var clienteEmail = linha[2] ? linha[2].toString().trim().toLowerCase() : "";
      if (clienteEmail !== emailLower) continue;

      mensagens.push({
        id: valorParaTextoV3_(linha[0]),
        dataHora: valorParaTextoV3_(linha[1]),
        clienteEmail: clienteEmail,
        clienteNome: valorParaTextoV3_(linha[3]),
        remetenteEmail: valorParaTextoV3_(linha[4]),
        remetenteNome: valorParaTextoV3_(linha[5]),
        remetenteTipo: valorParaTextoV3_(linha[6]),
        mensagem: valorParaTextoV3_(linha[7])
      });
    }

    return { success: true, mensagens: mensagens, total: mensagens.length };
  } catch (error) {
    return { success: false, mensagens: [], message: "Erro ao carregar mensagens: " + error.toString() };
  }
}

function getConversasAdminV3() {
  try {
    var aba = getAbaChatV3_();
    var dados = aba.getDataRange().getValues();
    var mapa = {};

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var clienteEmail = linha[2] ? linha[2].toString().trim().toLowerCase() : "";
      if (!clienteEmail) continue;
      var item = mapa[clienteEmail] || {
        clienteEmail: clienteEmail,
        clienteNome: valorParaTextoV3_(linha[3]),
        ultimaMensagem: "",
        ultimaData: "",
        total: 0,
        naoLidasAdmin: 0
      };
      item.clienteNome = item.clienteNome || valorParaTextoV3_(linha[3]);
      item.ultimaMensagem = valorParaTextoV3_(linha[7]);
      item.ultimaData = valorParaTextoV3_(linha[1]);
      item.total++;
      if (valorParaTextoV3_(linha[6]) === "Cliente" && !valorParaTextoV3_(linha[9])) item.naoLidasAdmin++;
      mapa[clienteEmail] = item;
    }

    var conversas = [];
    for (var email in mapa) conversas.push(mapa[email]);
    conversas.sort(function(a, b) {
      return normalizarDataHistorico_(b.ultimaData) - normalizarDataHistorico_(a.ultimaData);
    });

    return { success: true, conversas: conversas, total: conversas.length };
  } catch (error) {
    return { success: false, conversas: [], message: "Erro ao carregar conversas: " + error.toString() };
  }
}

function marcarMensagensAdminLidasV3(emailCliente) {
  try {
    var emailLower = emailCliente ? emailCliente.toString().trim().toLowerCase() : "";
    if (!emailLower) return { success: false, message: "Cliente nao informado." };

    var aba = getAbaChatV3_();
    var dados = aba.getDataRange().getValues();
    var atualizadas = 0;

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var clienteEmail = linha[2] ? linha[2].toString().trim().toLowerCase() : "";
      var remetenteTipo = valorParaTextoV3_(linha[6]);
      var lidaAdmin = valorParaTextoV3_(linha[9]);
      if (clienteEmail === emailLower && remetenteTipo === "Cliente" && !lidaAdmin) {
        aba.getRange(i + 1, 10).setValue("Sim");
        atualizadas++;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, atualizadas: atualizadas };
  } catch (error) {
    return { success: false, message: "Erro ao marcar mensagens: " + error.toString() };
  }
}

function atualizarStatusCompraAdminV3(idCompra, statusEntrega, observacao) {
  try {
    var aba = getAbaCompras();
    if (!aba.getRange(1, 10).getValue()) aba.getRange(1, 10).setValue("ObservacoesAdmin");
    var dados = aba.getDataRange().getValues();
    var id = idCompra ? idCompra.toString() : "";

    for (var i = 1; i < dados.length; i++) {
      if (!dados[i][0] || dados[i][0].toString() !== id) continue;

      aba.getRange(i + 1, 9).setValue(statusEntrega);
      if (observacao) {
        var atual = aba.getRange(i + 1, 10).getValue();
        var novaObs = (atual ? atual + "\n\n" : "") + formatarDataHoraSistema_() + " - " + observacao;
        aba.getRange(i + 1, 10).setValue(novaObs);
      }
      SpreadsheetApp.flush();

      var emailCliente = dados[i][2] || "";
      var cliente = dados[i][1] || "";
      try {
        registrarHistorico({
          id: id + "-" + statusEntrega,
          tipo: "Compra",
          cliente: cliente,
          email: emailCliente,
          titulo: "Atualizacao da compra " + id,
          descricao: observacao || "Status atualizado para " + statusEntrega,
          valor: formatarValorV3_(dados[i][6]),
          status: statusEntrega,
          linkAcao: "tab:historico",
          observacoes: observacao || ""
        });
      } catch (histError) {}

      try {
        enviarNotificacaoV3({
          titulo: "Atualizacao da sua compra",
          mensagem: "Sua compra " + id + " foi atualizada para: " + statusEntrega + (observacao ? ". " + observacao : ""),
          tipo: statusEntrega === "Cancelado" ? "danger" : "info",
          destinatarios: emailCliente,
          acao: "tab:historico",
          icone: "fas fa-shopping-cart"
        });
      } catch (notifError) {}

      return { success: true, message: "Compra atualizada." };
    }

    return { success: false, message: "Compra nao encontrada." };
  } catch (error) {
    return { success: false, message: "Erro ao atualizar compra: " + error.toString() };
  }
}

function diagnosticoSistemaV3(email) {
  var historico = getHistoricoCompletoUsuarioV3(email);
  var relatorios = getRelatoriosAdminV3();
  var notificacoes = getNotificacoesUsuarioV3(email);

  return {
    email: email,
    historicoQuantidade: historico.total,
    notificacoesQuantidade: notificacoes.total,
    totalClientes: relatorios.totalClientes,
    totalOrcamentos: relatorios.totalOrcamentos,
    totalCompras: relatorios.totalCompras,
    orcamentosPendentes: relatorios.orcamentosPendentes,
    comprasPendentes: relatorios.comprasPendentes,
    aReceber: relatorios.aReceber.total,
    erros: [].concat(historico.erros || [], notificacoes.erros || [])
  };
}
