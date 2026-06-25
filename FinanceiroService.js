var FINANCEIRO_STATUS_LANCAMENTO_ABERTO_V3 = "Aberto";
var FINANCEIRO_STATUS_LANCAMENTO_FATURADO_V3 = "Faturado";
var FINANCEIRO_STATUS_LANCAMENTO_CANCELADO_V3 = "Cancelado";

var FINANCEIRO_STATUS_COBRANCA_PENDENTE_V3 = "Pendente";
var FINANCEIRO_STATUS_COBRANCA_ENVIADA_V3 = "Enviado";
var FINANCEIRO_STATUS_COBRANCA_PAGA_V3 = "Pago";
var FINANCEIRO_STATUS_COBRANCA_CANCELADA_V3 = "Cancelado";

var FINANCEIRO_TIPO_CONSUMO_V3 = "Consumo";
var FINANCEIRO_TIPO_SERVICO_V3 = "Servico";
var FINANCEIRO_TIPO_MENSALIDADE_V3 = "Mensalidade";
var FINANCEIRO_PASTA_PDF_V3 = "JW_Cobrancas_PDF";
var FINANCEIRO_EMPRESA_NOME_V3 = "J.W PISCINAS";
var FINANCEIRO_PIX_CHAVE_V3 = "16993433911";
var FINANCEIRO_PIX_FAVORECIDO_V3 = "Jones William Rodrigues Duarte";
var FINANCEIRO_PIX_BANCO_V3 = "Bradesco";
var FINANCEIRO_CONTATO_TELEFONE_V3 = "(16) 9 9343-3911";
var FINANCEIRO_CONTATO_EMAIL_V3 = "jones-william-duarte@hotmail.com";

var FRASES_BASE_MENSAL_FINANCEIRO_V3_ = {
  "01": "Janeiro é o começo ideal para organizar a rotina e avançar com mais clareza.",
  "02": "Fevereiro mostra que constância e dedicação fazem a diferença todos os dias.",
  "03": "Março reforça que disciplina e cuidado abrem caminho para resultados sólidos.",
  "04": "Abril convida você a manter a calma, ajustar o rumo e seguir firme no processo.",
  "05": "Maio é o mês de manter o foco, mesmo quando o resultado ainda não apareceu.",
  "06": "Junho lembra que quem persiste com inteligência constrói um caminho mais seguro.",
  "07": "Julho pede equilíbrio, ritmo e confiança para continuar fazendo o que precisa ser feito.",
  "08": "Agosto reforça que evolução constante nasce de atitudes simples bem feitas.",
  "09": "Setembro inspira recomeços conscientes e passos firmes em direção ao objetivo.",
  "10": "Outubro mostra que crescer exige consistência, visão e coragem para continuar.",
  "11": "Novembro recompensa quem segue atento aos detalhes e não desiste no meio do caminho.",
  "12": "Dezembro é tempo de reconhecer conquistas e preparar com sabedoria o próximo ciclo."
};

var FRASES_COMPLEMENTO_ANUAL_FINANCEIRO_V3_ = {
  "2026": "Cada cuidado de hoje fortalece o sucesso de amanhã. 🍀✨",
  "2027": "Persistência com propósito transforma esforço em conquista real. 🍀✨",
  "2028": "Pequenos avanços consistentes criam grandes resultados ao longo do tempo. 🍀✨",
  "2029": "Quem mantém a direção certa colhe segurança, crescimento e tranquilidade. 🍀✨",
  "2030": "Determinação bem aplicada faz a rotina virar realização. 🍀✨"
};

var NOMES_MESES_FINANCEIRO_V3_ = {
  "01": "JANEIRO",
  "02": "FEVEREIRO",
  "03": "MARÇO",
  "04": "ABRIL",
  "05": "MAIO",
  "06": "JUNHO",
  "07": "JULHO",
  "08": "AGOSTO",
  "09": "SETEMBRO",
  "10": "OUTUBRO",
  "11": "NOVEMBRO",
  "12": "DEZEMBRO"
};

function registrarHistoricoFinanceiroV3_(dados) {
  try {
    if (typeof registrarHistorico !== "function") return;
    registrarHistorico(dados || {});
  } catch (error) {
    console.error("Erro ao registrar historico financeiro:", error);
  }
}

function gerarIdFinanceiroV3_(prefixo) {
  return prefixo + "-" + new Date().getTime().toString().slice(-9);
}

function getEmailOperadorFinanceiroV3_() {
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (error) {
    return "";
  }
}

function getLogoDataUriFinanceiroV3_() {
  try {
    var conteudo = HtmlService.createHtmlOutputFromFile("LogoData").getContent();
    var match = conteudo.match(/window\.JW_LOGO_DATA_URI\s*=\s*"([^"]+)"/);
    return match && match[1] ? match[1] : "";
  } catch (error) {
    return "";
  }
}

function getPastaPdfFinanceiroV3_() {
  var pastas = DriveApp.getFoldersByName(FINANCEIRO_PASTA_PDF_V3);
  if (pastas.hasNext()) return pastas.next();
  return DriveApp.createFolder(FINANCEIRO_PASTA_PDF_V3);
}

function garantirCabecalhosFinanceiroV3_(aba, cabecalhos) {
  var atuais = aba.getRange(1, 1, 1, cabecalhos.length).getValues()[0];
  var precisaAtualizar = false;

  for (var i = 0; i < cabecalhos.length; i++) {
    if (atuais[i] !== cabecalhos[i]) {
      atuais[i] = cabecalhos[i];
      precisaAtualizar = true;
    }
  }

  if (precisaAtualizar) {
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([atuais]);
  }

  aba.setFrozenRows(1);
}

function getAbaItensCobrancaFinanceiroV3_() {
  if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
    return getFirestoreSheetAdapter_("ItensCobranca");
  }

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("ItensCobranca");
  var cabecalhos = [
    "ID",
    "Nome",
    "Categoria",
    "Unidade",
    "ValorPadrao",
    "TipoUso",
    "Ativo",
    "Descricao",
    "DataCadastro",
    "AtualizadoEm",
    "EstoqueInicial",
    "ImagemURL",
    "ProdutoID"
  ];

  if (!aba) {
    aba = planilha.insertSheet("ItensCobranca");
  }

  garantirCabecalhosFinanceiroV3_(aba, cabecalhos);

  try {
    var regra = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    aba.getRange(2, 7, Math.max(aba.getMaxRows() - 1, 1), 1).setDataValidation(regra);
  } catch (error) {}

  return aba;
}

function getAbaLancamentosFinanceirosV3_() {
  if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
    return getFirestoreSheetAdapter_("LancamentosFinanceiros");
  }

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("LancamentosFinanceiros");
  var cabecalhos = [
    "ID",
    "ClienteEmail",
    "ClienteNome",
    "Telefone",
    "TipoLancamento",
    "ItemID",
    "ItemNome",
    "Descricao",
    "Quantidade",
    "Unidade",
    "ValorUnitario",
    "ValorTotal",
    "Competencia",
    "Referencia",
    "Status",
    "CobrancaID",
    "DataLancamento",
    "Observacoes",
    "CriadoPor",
    "AssinaturaID"
  ];

  if (!aba) {
    aba = planilha.insertSheet("LancamentosFinanceiros");
  }

  garantirCabecalhosFinanceiroV3_(aba, cabecalhos);
  return aba;
}

function getAbaAssinaturasMensaisFinanceiroV3_() {
  if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
    return getFirestoreSheetAdapter_("AssinaturasMensais");
  }

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("AssinaturasMensais");
  var cabecalhos = [
    "ID",
    "ClienteEmail",
    "ClienteNome",
    "Telefone",
    "Descricao",
    "Valor",
    "DiaVencimento",
    "Status",
    "UltimaCompetenciaGerada",
    "Observacoes",
    "DataCadastro"
  ];

  if (!aba) {
    aba = planilha.insertSheet("AssinaturasMensais");
  }

  garantirCabecalhosFinanceiroV3_(aba, cabecalhos);
  return aba;
}

function getAbaCobrancasFinanceiroV3_() {
  if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
    return getFirestoreSheetAdapter_("CobrancasClientes");
  }

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("CobrancasClientes");
  var cabecalhos = [
    "ID",
    "ClienteEmail",
    "ClienteNome",
    "Telefone",
    "Competencia",
    "Referencia",
    "ItensJson",
    "QuantidadeItens",
    "ValorTotal",
    "Vencimento",
    "Status",
    "LinkWhatsApp",
    "MensagemWhatsApp",
    "DataGeracao",
    "DataEnvio",
    "Observacoes",
    "CriadoPor",
    "PdfUrl",
    "PdfFileId",
    "PdfGeradoEm"
  ];

  if (!aba) {
    aba = planilha.insertSheet("CobrancasClientes");
  }

  garantirCabecalhosFinanceiroV3_(aba, cabecalhos);
  return aba;
}

function getClientesFinanceiroV3_() {
  var usuarios = typeof getUsuariosAtivosV3_ === "function" ? getUsuariosAtivosV3_() : [];
  return usuarios.filter(function(usuario) {
    return usuario.tipo !== "Administrador";
  }).map(function(usuario) {
    return {
      email: valorParaTextoV3_(usuario.email).toLowerCase(),
      usuario: valorParaTextoV3_(usuario.usuario),
      telefone: valorParaTextoV3_(usuario.telefone),
      endereco: valorParaTextoV3_(usuario.endereco),
      bairro: valorParaTextoV3_(usuario.bairro),
      status: valorParaTextoV3_(usuario.status || "Ativo")
    };
  });
}

function buscarClienteFinanceiroV3_(email) {
  var emailLower = email ? email.toString().trim().toLowerCase() : "";
  if (!emailLower) return null;

  var clientes = getClientesFinanceiroV3_();
  for (var i = 0; i < clientes.length; i++) {
    if (clientes[i].email === emailLower) {
      return clientes[i];
    }
  }

  return null;
}

function normalizarTelefoneWhatsappFinanceiroV3_(telefone) {
  var numeros = valorParaTextoV3_(telefone).replace(/\D/g, "");
  if (!numeros) return "";
  if (numeros.length === 10 || numeros.length === 11) return "55" + numeros;
  return numeros;
}

function formatarCompetenciaFinanceiroV3_(data) {
  return Utilities.formatDate(data || new Date(), "America/Sao_Paulo", "MM/yyyy");
}

function normalizarCompetenciaFinanceiroV3_(competencia) {
  var texto = valorParaTextoV3_(competencia).trim();
  if (!texto) return formatarCompetenciaFinanceiroV3_(new Date());

  var partes = texto.split("/");
  if (partes.length === 2) {
    var mes = ("0" + parseInt(partes[0], 10)).slice(-2);
    var ano = partes[1];
    if (mes !== "NaN" && ano && ano.length === 4) {
      return mes + "/" + ano;
    }
  }

  return formatarCompetenciaFinanceiroV3_(new Date());
}

function parseDiaVencimentoFinanceiroV3_(valor) {
  var dia = parseInt(valor, 10);
  if (isNaN(dia) || dia < 1) return 10;
  if (dia > 28) return 28;
  return dia;
}

function montarDataVencimentoFinanceiroV3_(competencia, dia) {
  var comp = normalizarCompetenciaFinanceiroV3_(competencia);
  var partes = comp.split("/");
  var mes = parseInt(partes[0], 10) - 1;
  var ano = parseInt(partes[1], 10);
  var data = new Date(ano, mes, parseDiaVencimentoFinanceiroV3_(dia));
  return formatarDataSistema_(data);
}

function getItensCobrancaFinanceiroV3() {
  try {
    var aba = getAbaItensCobrancaFinanceiroV3_();
    var dados = aba.getDataRange().getValues();
    var itens = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      if (!linha[0]) continue;
      itens.push({
        id: valorParaTextoV3_(linha[0]),
        nome: valorParaTextoV3_(linha[1]),
        categoria: valorParaTextoV3_(linha[2]),
        unidade: valorParaTextoV3_(linha[3] || "un"),
        valorPadrao: parseValorV3_(linha[4]),
        valorPadraoFormatado: formatarValorV3_(linha[4]),
        tipoUso: valorParaTextoV3_(linha[5] || "Financeiro"),
        ativo: linha[6] === true || normalizarTextoV3_(linha[6]) === "true" || normalizarTextoV3_(linha[6]) === "sim",
        descricao: valorParaTextoV3_(linha[7]),
        dataCadastro: valorParaTextoV3_(linha[8]),
        atualizadoEm: valorParaTextoV3_(linha[9]),
        estoqueInicial: parseInt(linha[10], 10) || 0,
        imagemURL: valorParaTextoV3_(linha[11]),
        produtoId: valorParaTextoV3_(linha[12] || linha[0])
      });
    }

    return {
      success: true,
      itens: itens,
      total: itens.length
    };
  } catch (error) {
    return {
      success: false,
      itens: [],
      message: "Erro ao carregar itens: " + error.toString()
    };
  }
}

function getItemCobrancaFinanceiroV3_(itemId) {
  var itens = getItensCobrancaFinanceiroV3();
  var lista = itens && itens.itens ? itens.itens : [];
  var itemIdStr = valorParaTextoV3_(itemId);

  for (var i = 0; i < lista.length; i++) {
    if (lista[i].id === itemIdStr) return lista[i];
  }

  return null;
}

function getAbaProdutosFinanceiroV3_() {
  if (typeof getAbaProdutos === "function") return getAbaProdutos();

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Produtos");
  var cabecalhos = ["ID", "Nome", "Categoria", "Descricao", "Preco", "Estoque", "ImagemURL", "ManualURL", "ManualNome", "PrecoComDesconto", "ExibirManual", "ManualVideoURL"];

  if (!aba) aba = planilha.insertSheet("Produtos");
  garantirCabecalhosFinanceiroV3_(aba, cabecalhos);
  return aba;
}

function sincronizarProdutoLojaFinanceiroV3_(itemId, dados) {
  try {
    var aba = getAbaProdutosFinanceiroV3_();
    var valores = aba.getDataRange().getValues();
    var produtoId = valorParaTextoV3_(dados.produtoId || itemId || gerarIdFinanceiroV3_("PRD"));
    var linhaDestino = -1;

    for (var i = 1; i < valores.length; i++) {
      var idAtual = valorParaTextoV3_(valores[i][0]);
      if (idAtual === produtoId || idAtual === valorParaTextoV3_(itemId)) {
        linhaDestino = i + 1;
        produtoId = idAtual || produtoId;
        break;
      }
    }

    var existente = linhaDestino > 0 ? aba.getRange(linhaDestino, 1, 1, 12).getValues()[0] : [];
    var linha = [
      produtoId,
      valorParaTextoV3_(dados.nome),
      valorParaTextoV3_(dados.categoria || "Geral"),
      valorParaTextoV3_(dados.descricao),
      parseValorV3_(dados.valorPadrao),
      Math.max(parseInt(dados.estoque, 10) || 0, 0),
      valorParaTextoV3_(dados.imagemURL),
      existente[7] || "",
      existente[8] || "",
      existente[9] || "",
      existente[10] === true,
      existente[11] || ""
    ];

    if (linhaDestino > 0) aba.getRange(linhaDestino, 1, 1, linha.length).setValues([linha]);
    else aba.appendRow(linha);

    return { success: true, produtoId: produtoId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function salvarItemCobrancaFinanceiroV3(dados) {
  try {
    dados = dados || {};

    var nome = valorParaTextoV3_(dados.nome).trim();
    if (!nome) {
      return { success: false, message: "Informe o nome do item." };
    }

    var valorPadrao = parseValorV3_(dados.valorPadrao);
    if (valorPadrao < 0) {
      return { success: false, message: "Valor padrao invalido." };
    }

    var estoqueInicial = Math.max(parseInt(dados.estoque, 10) || 0, 0);
    var imagemURL = valorParaTextoV3_(dados.imagemDataUrl || dados.imagemURL);

    var aba = getAbaItensCobrancaFinanceiroV3_();
    var dadosPlanilha = aba.getDataRange().getValues();
    var agora = formatarDataHoraSistema_();
    var itemId = valorParaTextoV3_(dados.id) || gerarIdFinanceiroV3_("ITEM");
    var linhaDestino = -1;
    var produtoIdAtual = "";

    for (var i = 1; i < dadosPlanilha.length; i++) {
      if (valorParaTextoV3_(dadosPlanilha[i][0]) === itemId) {
        linhaDestino = i + 1;
        produtoIdAtual = valorParaTextoV3_(dadosPlanilha[i][12] || dadosPlanilha[i][0]);
        break;
      }
    }

    if (!produtoIdAtual) produtoIdAtual = valorParaTextoV3_(dados.produtoId || itemId);

    var linha = [
      itemId,
      nome,
      valorParaTextoV3_(dados.categoria || "Geral"),
      valorParaTextoV3_(dados.unidade || "un"),
      valorPadrao,
      valorParaTextoV3_(dados.tipoUso || "Financeiro"),
      dados.ativo === false ? false : true,
      valorParaTextoV3_(dados.descricao),
      linhaDestino > 0 ? valorParaTextoV3_(aba.getRange(linhaDestino, 9).getValue()) || agora : agora,
      agora,
      estoqueInicial,
      imagemURL,
      produtoIdAtual
    ];

    if (linhaDestino > 0) {
      aba.getRange(linhaDestino, 1, 1, linha.length).setValues([linha]);
    } else {
      aba.appendRow(linha);
    }

    var syncProduto = sincronizarProdutoLojaFinanceiroV3_(itemId, {
      produtoId: produtoIdAtual,
      nome: nome,
      categoria: valorParaTextoV3_(dados.categoria || "Geral"),
      descricao: valorParaTextoV3_(dados.descricao),
      valorPadrao: valorPadrao,
      estoque: estoqueInicial,
      imagemURL: imagemURL
    });

    SpreadsheetApp.flush();
    if (!syncProduto.success) {
      return {
        success: false,
        message: "Item salvo, mas houve erro ao sincronizar com a Loja: " + syncProduto.message
      };
    }
    return {
      success: true,
      id: itemId,
      produtoId: syncProduto.produtoId || produtoIdAtual,
      message: linhaDestino > 0 ? "Item atualizado com sucesso e sincronizado com a Loja." : "Item cadastrado com sucesso e disponivel na Loja."
    };
  } catch (error) {
    return { success: false, message: "Erro ao salvar item: " + error.toString() };
  }
}

function registrarLancamentoFinanceiroV3(dados) {
  try {
    dados = dados || {};
    var clienteEmail = valorParaTextoV3_(dados.clienteEmail).trim().toLowerCase();
    var tipoLancamento = valorParaTextoV3_(dados.tipoLancamento || FINANCEIRO_TIPO_CONSUMO_V3).trim();
    if (!clienteEmail) return { success: false, message: "Selecione um cliente." };

    var cliente = buscarClienteFinanceiroV3_(clienteEmail);
    if (!cliente) return { success: false, message: "Cliente nao encontrado." };

    var item = dados.itemId ? getItemCobrancaFinanceiroV3_(dados.itemId) : null;
    var quantidade = parseFloat(dados.quantidade);
    if (isNaN(quantidade) || quantidade <= 0) quantidade = 1;

    var valorUnitario = parseValorV3_(dados.valorUnitario);
    if (!valorUnitario && item) valorUnitario = item.valorPadrao;
    if (valorUnitario < 0) return { success: false, message: "Valor unitario invalido." };

    var descricao = valorParaTextoV3_(dados.descricao).trim();
    if (!descricao) descricao = item ? item.nome : tipoLancamento;

    var competencia = normalizarCompetenciaFinanceiroV3_(dados.competencia);
    var referencia = valorParaTextoV3_(dados.referencia || competencia).trim();
    var itemNome = item ? item.nome : valorParaTextoV3_(dados.itemNome);
    var unidade = item ? item.unidade : valorParaTextoV3_(dados.unidade || "un");
    var total = quantidade * valorUnitario;
    var lancamentoId = gerarIdFinanceiroV3_("LCT");

    getAbaLancamentosFinanceirosV3_().appendRow([
      lancamentoId,
      cliente.email,
      cliente.usuario,
      cliente.telefone,
      tipoLancamento,
      item ? item.id : valorParaTextoV3_(dados.itemId),
      itemNome,
      descricao,
      quantidade,
      unidade,
      valorUnitario,
      total,
      competencia,
      referencia,
      FINANCEIRO_STATUS_LANCAMENTO_ABERTO_V3,
      "",
      formatarDataHoraSistema_(),
      valorParaTextoV3_(dados.observacoes),
      getEmailOperadorFinanceiroV3_(),
      valorParaTextoV3_(dados.assinaturaId)
    ]);

    registrarHistoricoFinanceiroV3_({
      id: lancamentoId,
      tipo: tipoLancamento,
      cliente: cliente.usuario,
      email: cliente.email,
      telefone: cliente.telefone,
      titulo: descricao,
      descricao: descricao + " | Competencia: " + competencia,
      valor: formatarValorV3_(total),
      status: FINANCEIRO_STATUS_LANCAMENTO_ABERTO_V3,
      dataRegistro: formatarDataHoraSistema_(),
      linkAcao: "tab:historico",
      observacoes: referencia
    });

    SpreadsheetApp.flush();
    return {
      success: true,
      id: lancamentoId,
      totalFormatado: formatarValorV3_(total),
      message: "Lancamento registrado com sucesso."
    };
  } catch (error) {
    return { success: false, message: "Erro ao registrar lancamento: " + error.toString() };
  }
}

function getLancamentosFinanceirosV3_() {
  var aba = getAbaLancamentosFinanceirosV3_();
  var dados = aba.getDataRange().getValues();
  var itens = [];

  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    if (!linha[0]) continue;
    itens.push({
      id: valorParaTextoV3_(linha[0]),
      clienteEmail: valorParaTextoV3_(linha[1]).toLowerCase(),
      clienteNome: valorParaTextoV3_(linha[2]),
      telefone: valorParaTextoV3_(linha[3]),
      tipoLancamento: valorParaTextoV3_(linha[4]),
      itemId: valorParaTextoV3_(linha[5]),
      itemNome: valorParaTextoV3_(linha[6]),
      descricao: valorParaTextoV3_(linha[7]),
      quantidade: parseFloat(linha[8]) || 0,
      unidade: valorParaTextoV3_(linha[9]),
      valorUnitario: parseValorV3_(linha[10]),
      valorUnitarioFormatado: formatarValorV3_(linha[10]),
      valorTotal: parseValorV3_(linha[11]),
      valorTotalFormatado: formatarValorV3_(linha[11]),
      competencia: valorParaTextoV3_(linha[12]),
      referencia: valorParaTextoV3_(linha[13]),
      status: valorParaTextoV3_(linha[14] || FINANCEIRO_STATUS_LANCAMENTO_ABERTO_V3),
      cobrancaId: valorParaTextoV3_(linha[15]),
      dataLancamento: valorParaTextoV3_(linha[16]),
      observacoes: valorParaTextoV3_(linha[17]),
      criadoPor: valorParaTextoV3_(linha[18]),
      assinaturaId: valorParaTextoV3_(linha[19])
    });
  }

  return itens;
}

function getLancamentosFinanceirosAbertosAgrupadosV3_() {
  var lancamentos = getLancamentosFinanceirosV3_().filter(function(lancamento) {
    return normalizarTextoV3_(lancamento.status) === normalizarTextoV3_(FINANCEIRO_STATUS_LANCAMENTO_ABERTO_V3);
  });
  var mapa = {};

  for (var i = 0; i < lancamentos.length; i++) {
    var lancamento = lancamentos[i];
    var chave = lancamento.clienteEmail;
    if (!mapa[chave]) {
      mapa[chave] = {
        clienteEmail: lancamento.clienteEmail,
        clienteNome: lancamento.clienteNome,
        telefone: lancamento.telefone,
        totalAberto: 0,
        totalAbertoFormatado: "R$ 0,00",
        itens: [],
        competencias: {}
      };
    }

    mapa[chave].itens.push(lancamento);
    mapa[chave].totalAberto += lancamento.valorTotal;
    if (lancamento.competencia) mapa[chave].competencias[lancamento.competencia] = true;
  }

  var grupos = [];
  for (var email in mapa) {
    mapa[email].totalAbertoFormatado = formatarValorV3_(mapa[email].totalAberto);
    mapa[email].competenciasLista = Object.keys(mapa[email].competencias).sort();
    grupos.push(mapa[email]);
  }

  grupos.sort(function(a, b) {
    return a.clienteNome.localeCompare(b.clienteNome);
  });
  return grupos;
}

function getAssinaturasMensaisFinanceiroV3() {
  try {
    var aba = getAbaAssinaturasMensaisFinanceiroV3_();
    var dados = aba.getDataRange().getValues();
    var itens = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      if (!linha[0]) continue;
      itens.push({
        id: valorParaTextoV3_(linha[0]),
        clienteEmail: valorParaTextoV3_(linha[1]).toLowerCase(),
        clienteNome: valorParaTextoV3_(linha[2]),
        telefone: valorParaTextoV3_(linha[3]),
        descricao: valorParaTextoV3_(linha[4]),
        valor: parseValorV3_(linha[5]),
        valorFormatado: formatarValorV3_(linha[5]),
        diaVencimento: parseDiaVencimentoFinanceiroV3_(linha[6]),
        status: valorParaTextoV3_(linha[7] || "Ativo"),
        ultimaCompetenciaGerada: valorParaTextoV3_(linha[8]),
        observacoes: valorParaTextoV3_(linha[9]),
        dataCadastro: valorParaTextoV3_(linha[10])
      });
    }

    return { success: true, itens: itens, total: itens.length };
  } catch (error) {
    return { success: false, itens: [], message: "Erro ao carregar mensalidades: " + error.toString() };
  }
}

function salvarAssinaturaMensalFinanceiroV3(dados) {
  try {
    dados = dados || {};
    var clienteEmail = valorParaTextoV3_(dados.clienteEmail).trim().toLowerCase();
    if (!clienteEmail) return { success: false, message: "Selecione um cliente." };

    var cliente = buscarClienteFinanceiroV3_(clienteEmail);
    if (!cliente) return { success: false, message: "Cliente nao encontrado." };

    var descricao = valorParaTextoV3_(dados.descricao).trim();
    if (!descricao) return { success: false, message: "Informe a descricao da mensalidade." };

    var valor = parseValorV3_(dados.valor);
    if (valor <= 0) return { success: false, message: "Informe um valor valido." };

    var aba = getAbaAssinaturasMensaisFinanceiroV3_();
    var dadosPlanilha = aba.getDataRange().getValues();
    var assinaturaId = valorParaTextoV3_(dados.id) || gerarIdFinanceiroV3_("MEN");
    var linhaDestino = -1;

    for (var i = 1; i < dadosPlanilha.length; i++) {
      if (valorParaTextoV3_(dadosPlanilha[i][0]) === assinaturaId) {
        linhaDestino = i + 1;
        break;
      }
    }

    var linha = [
      assinaturaId,
      cliente.email,
      cliente.usuario,
      cliente.telefone,
      descricao,
      valor,
      parseDiaVencimentoFinanceiroV3_(dados.diaVencimento),
      valorParaTextoV3_(dados.status || "Ativo"),
      linhaDestino > 0 ? valorParaTextoV3_(aba.getRange(linhaDestino, 9).getValue()) : "",
      valorParaTextoV3_(dados.observacoes),
      linhaDestino > 0 ? valorParaTextoV3_(aba.getRange(linhaDestino, 11).getValue()) : formatarDataHoraSistema_()
    ];

    if (linhaDestino > 0) {
      aba.getRange(linhaDestino, 1, 1, linha.length).setValues([linha]);
    } else {
      aba.appendRow(linha);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      id: assinaturaId,
      message: linhaDestino > 0 ? "Mensalidade atualizada com sucesso." : "Mensalidade cadastrada com sucesso."
    };
  } catch (error) {
    return { success: false, message: "Erro ao salvar mensalidade: " + error.toString() };
  }
}

function gerarLancamentosMensaisFinanceiroV3(competencia) {
  try {
    var competenciaNormalizada = normalizarCompetenciaFinanceiroV3_(competencia);
    var assinaturasResp = getAssinaturasMensaisFinanceiroV3();
    var assinaturas = assinaturasResp && assinaturasResp.itens ? assinaturasResp.itens : [];
    var aba = getAbaAssinaturasMensaisFinanceiroV3_();
    var criadas = 0;

    for (var i = 0; i < assinaturas.length; i++) {
      var assinatura = assinaturas[i];
      if (normalizarTextoV3_(assinatura.status) !== "ativo") continue;
      if (assinatura.ultimaCompetenciaGerada === competenciaNormalizada) continue;

      var resultado = registrarLancamentoFinanceiroV3({
        clienteEmail: assinatura.clienteEmail,
        tipoLancamento: FINANCEIRO_TIPO_MENSALIDADE_V3,
        descricao: assinatura.descricao,
        quantidade: 1,
        unidade: "mensal",
        valorUnitario: assinatura.valor,
        competencia: competenciaNormalizada,
        referencia: "Mensalidade " + competenciaNormalizada,
        observacoes: assinatura.observacoes,
        assinaturaId: assinatura.id
      });

      if (!resultado.success) continue;

      for (var linha = 2; linha <= aba.getLastRow(); linha++) {
        if (valorParaTextoV3_(aba.getRange(linha, 1).getValue()) === assinatura.id) {
          aba.getRange(linha, 9).setValue(competenciaNormalizada);
          criadas++;
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      competencia: competenciaNormalizada,
      totalCriado: criadas,
      message: criadas ? "Mensalidades do periodo geradas com sucesso." : "Nao havia novas mensalidades para gerar."
    };
  } catch (error) {
    return { success: false, message: "Erro ao gerar mensalidades: " + error.toString() };
  }
}

function montarMensagemWhatsAppCobrancaFinanceiroV3_(cobranca) {
  var linhas = [];
  var itens = cobranca && cobranca.itens ? cobranca.itens : [];
  var nome = cobranca && cobranca.clienteNome ? cobranca.clienteNome : "";
  var mesAno = obterMesAnoCobrancaFinanceiroV3_(cobranca);
  var nomeMes = obterNomeMesFinanceiroV3_(mesAno.mes);
  var totalMensalidade = 0;
  var servicos = [];
  var produtos = [];

  linhas.push("💧 J.W PISCINAS");
  linhas.push("");
  linhas.push("Prezado(a) " + nome + ", tudo bem?");
  linhas.push("Segue a cobrança referente à manutenção da piscina deste mês, incluindo produtos de tratamento.");
  linhas.push("");

  for (var i = 0; i < itens.length; i++) {
    var item = itens[i];
    var descricao = valorParaTextoV3_(item.descricao || item.itemNome || item.tipoLancamento || "Lançamento");
    var valorTotal = formatarValorV3_(item.valorTotal);
    var quantidade = parseFloat(item.quantidade) || 0;
    var sufixoQtd = quantidade > 1 ? " x" + quantidade : "";
    var tipoItem = classificarTipoItemCobrancaFinanceiroV3_(item);
    if (tipoItem === "mensalidade") {
      totalMensalidade += parseValorV3_(item.valorTotal);
      continue;
    }
    if (tipoItem === "servico") {
      servicos.push("- " + descricao + sufixoQtd + " " + valorTotal);
      continue;
    }
    produtos.push("- " + descricao + sufixoQtd + " " + valorTotal);
  }

  linhas.push("");
  if (totalMensalidade > 0) {
    linhas.push("Mensalidade de " + nomeMes + ": " + formatarValorV3_(totalMensalidade));
  }
  if (servicos.length) {
    linhas.push("Serviços:");
    linhas = linhas.concat(servicos);
  }
  if (produtos.length) {
    linhas.push("Produtos:");
    linhas = linhas.concat(produtos);
  }
  if (!totalMensalidade && !servicos.length && !produtos.length) {
    linhas.push("Itens cobrados:");
    linhas.push("- " + formatarValorV3_(cobranca.valorTotal));
  }
  linhas.push("");
  linhas.push("Valor total: " + formatarValorV3_(cobranca.valorTotal));
  linhas.push("");
  linhas.push("📅 Vencimento: " + valorParaTextoV3_(cobranca.vencimento || "-"));
  linhas.push("💲 Dados para pagamento (Pix):");
  linhas.push("🔑 Chave: 16993433911");
  linhas.push("👤 Nome: Jones William Rodrigues Duarte");
  linhas.push("🏦 Banco: Bradesco");
  linhas.push("");
  linhas.push("✨ Frase do mês:");
  linhas.push(obterFraseMesFinanceiroV3_(cobranca));
  if (cobranca.observacoes) {
    linhas.push("");
    linhas.push("Observações:");
    linhas.push(valorParaTextoV3_(cobranca.observacoes));
  }
  linhas.push("");
  linhas.push("Agradecemos pela confiança!");
  linhas.push("📞 (16) 9 9343-3911");
  linhas.push("✉️ jones-william-duarte@hotmail.com");

  return linhas.join("\n");
}

function construirLinkWhatsAppFinanceiroV3_(telefone, mensagem) {
  var numero = normalizarTelefoneWhatsappFinanceiroV3_(telefone);
  if (!numero) return "";
  return "https://wa.me/" + numero + "?text=" + encodeURIComponent(mensagem || "");
}

function getConfiguracaoWhatsAppFinanceiroV3_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var apiUrl = valorParaTextoV3_(props.getProperty("WHATSAPP_API_URL")).trim();
    var token = valorParaTextoV3_(props.getProperty("WHATSAPP_ACCESS_TOKEN")).trim();
    var templateName = valorParaTextoV3_(props.getProperty("WHATSAPP_TEMPLATE_NAME")).trim();
    var templateLanguage = valorParaTextoV3_(props.getProperty("WHATSAPP_TEMPLATE_LANGUAGE")).trim() || "pt_BR";
    var modoEnvio = valorParaTextoV3_(props.getProperty("WHATSAPP_SEND_MODE")).trim().toLowerCase();
    if (!modoEnvio) modoEnvio = templateName ? "template" : "text";

    return {
      ativo: !!(apiUrl && token && (templateName || modoEnvio === "text")),
      apiUrl: apiUrl,
      token: token,
      templateName: templateName,
      templateLanguage: templateLanguage,
      modoEnvio: modoEnvio
    };
  } catch (error) {
    return {
      ativo: false,
      apiUrl: "",
      token: "",
      templateName: "",
      templateLanguage: "pt_BR",
      modoEnvio: "text",
      erro: error.toString()
    };
  }
}

function getCobrancasFinanceiroV3_() {
  var aba = getAbaCobrancasFinanceiroV3_();
  var dados = aba.getDataRange().getValues();
  var cobrancas = [];

  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    if (!linha[0]) continue;

    var itens = [];
    try {
      itens = JSON.parse(linha[6] || "[]");
    } catch (error) {
      itens = [];
    }

    var cobranca = {
      id: valorParaTextoV3_(linha[0]),
      clienteEmail: valorParaTextoV3_(linha[1]).toLowerCase(),
      clienteNome: valorParaTextoV3_(linha[2]),
      telefone: valorParaTextoV3_(linha[3]),
      competencia: valorParaTextoV3_(linha[4]),
      referencia: valorParaTextoV3_(linha[5]),
      itens: itens,
      quantidadeItens: parseInt(linha[7], 10) || itens.length,
      valorTotal: parseValorV3_(linha[8]),
      valorTotalFormatado: formatarValorV3_(linha[8]),
      vencimento: valorParaTextoV3_(linha[9]),
      status: valorParaTextoV3_(linha[10] || FINANCEIRO_STATUS_COBRANCA_PENDENTE_V3),
      linkWhatsApp: "",
      mensagemWhatsApp: "",
      dataGeracao: valorParaTextoV3_(linha[13]),
      dataEnvio: valorParaTextoV3_(linha[14]),
      observacoes: valorParaTextoV3_(linha[15]),
      criadoPor: valorParaTextoV3_(linha[16]),
      pdfUrl: valorParaTextoV3_(linha[17]),
      pdfFileId: valorParaTextoV3_(linha[18]),
      pdfGeradoEm: valorParaTextoV3_(linha[19])
    };
    cobranca.mensagemWhatsApp = montarMensagemWhatsAppCobrancaFinanceiroV3_(cobranca);
    cobranca.linkWhatsApp = construirLinkWhatsAppFinanceiroV3_(cobranca.telefone, cobranca.mensagemWhatsApp);
    cobrancas.push(cobranca);
  }

  cobrancas.sort(function(a, b) {
    return normalizarDataHistorico_(b.dataGeracao) - normalizarDataHistorico_(a.dataGeracao);
  });
  return cobrancas;
}

function gerarCobrancaFinanceiroV3(dados) {
  try {
    dados = dados || {};
    var clienteEmail = valorParaTextoV3_(dados.clienteEmail).trim().toLowerCase();
    if (!clienteEmail) return { success: false, message: "Selecione um cliente para cobrar." };

    var lancamentos = getLancamentosFinanceirosV3_().filter(function(lancamento) {
      return lancamento.clienteEmail === clienteEmail &&
        normalizarTextoV3_(lancamento.status) === normalizarTextoV3_(FINANCEIRO_STATUS_LANCAMENTO_ABERTO_V3);
    });

    if (!lancamentos.length) {
      return { success: false, message: "Esse cliente nao possui lancamentos em aberto." };
    }

    var cliente = buscarClienteFinanceiroV3_(clienteEmail);
    if (!cliente) return { success: false, message: "Cliente nao encontrado." };

    var cobrancaId = gerarIdFinanceiroV3_("COB");
    var valorTotal = 0;
    var competenciasMapa = {};
    var referencias = [];
    var itensResumo = [];

    for (var i = 0; i < lancamentos.length; i++) {
      var lancamento = lancamentos[i];
      valorTotal += lancamento.valorTotal;
      if (lancamento.competencia) competenciasMapa[lancamento.competencia] = true;
      if (lancamento.referencia) referencias.push(lancamento.referencia);

      itensResumo.push({
        id: lancamento.id,
        tipoLancamento: lancamento.tipoLancamento,
        itemId: lancamento.itemId,
        itemNome: lancamento.itemNome,
        descricao: lancamento.descricao,
        quantidade: lancamento.quantidade,
        unidade: lancamento.unidade,
        valorUnitario: lancamento.valorUnitario,
        valorTotal: lancamento.valorTotal,
        competencia: lancamento.competencia,
        referencia: lancamento.referencia
      });
    }

    var competencias = Object.keys(competenciasMapa).sort();
    var competenciaTexto = competencias.join(", ");
    var referenciaTexto = referencias.slice(0, 5).join(" | ");
    var vencimento = valorParaTextoV3_(dados.vencimento).trim();
    if (!vencimento) vencimento = formatarDataSistema_(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000));

    var cobranca = {
      id: cobrancaId,
      clienteEmail: cliente.email,
      clienteNome: cliente.usuario,
      telefone: cliente.telefone,
      competencia: competenciaTexto,
      referencia: referenciaTexto,
      itens: itensResumo,
      valorTotal: valorTotal,
      vencimento: vencimento,
      observacoes: valorParaTextoV3_(dados.observacoes)
    };

    var pdfInfo = gerarArquivoPdfCobrancaFinanceiroV3_(cobranca);
    cobranca.pdfUrl = pdfInfo.pdfUrl;
    cobranca.pdfFileId = pdfInfo.pdfFileId;
    cobranca.pdfGeradoEm = pdfInfo.pdfGeradoEm;
    var mensagem = montarMensagemWhatsAppCobrancaFinanceiroV3_(cobranca);
    var link = construirLinkWhatsAppFinanceiroV3_(cliente.telefone, mensagem);

    getAbaCobrancasFinanceiroV3_().appendRow([
      cobrancaId,
      cliente.email,
      cliente.usuario,
      cliente.telefone,
      competenciaTexto,
      referenciaTexto,
      JSON.stringify(itensResumo),
      itensResumo.length,
      valorTotal,
      vencimento,
      FINANCEIRO_STATUS_COBRANCA_PENDENTE_V3,
      link,
      mensagem,
      formatarDataHoraSistema_(),
      "",
      valorParaTextoV3_(dados.observacoes),
      getEmailOperadorFinanceiroV3_(),
      cobranca.pdfUrl,
      cobranca.pdfFileId,
      cobranca.pdfGeradoEm
    ]);

    registrarHistoricoFinanceiroV3_({
      id: cobrancaId,
      tipo: "Cobranca",
      cliente: cliente.usuario,
      email: cliente.email,
      telefone: cliente.telefone,
      titulo: referenciaTexto || ("Cobranca " + cobrancaId),
      descricao: "Cobranca gerada com " + itensResumo.length + " item(ns).",
      valor: formatarValorV3_(valorTotal),
      status: FINANCEIRO_STATUS_COBRANCA_PENDENTE_V3,
      dataRegistro: formatarDataHoraSistema_(),
      linkAcao: "tab:historico",
      observacoes: competenciaTexto ? "Competencia: " + competenciaTexto : ""
    });

    var abaLancamentos = getAbaLancamentosFinanceirosV3_();
    var dadosLancamentos = abaLancamentos.getDataRange().getValues();
    for (var linha = 1; linha < dadosLancamentos.length; linha++) {
      var lancamentoId = valorParaTextoV3_(dadosLancamentos[linha][0]);
      for (var indice = 0; indice < itensResumo.length; indice++) {
        if (itensResumo[indice].id !== lancamentoId) continue;
        abaLancamentos.getRange(linha + 1, 15).setValue(FINANCEIRO_STATUS_LANCAMENTO_FATURADO_V3);
        abaLancamentos.getRange(linha + 1, 16).setValue(cobrancaId);
        break;
      }
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      id: cobrancaId,
      linkWhatsApp: link,
      mensagemWhatsApp: mensagem,
      pdfUrl: cobranca.pdfUrl,
      valorTotalFormatado: formatarValorV3_(valorTotal),
      message: "Cobranca gerada com sucesso."
    };
  } catch (error) {
    return { success: false, message: "Erro ao gerar cobranca: " + error.toString() };
  }
}

function obterMesAnoCobrancaFinanceiroV3_(cobranca) {
  var competencia = valorParaTextoV3_(cobranca && cobranca.competencia).trim();
  if (competencia) {
    var primeira = competencia.split(",")[0].trim();
    var partesComp = primeira.split("/");
    if (partesComp.length === 2) {
      return {
        mes: ("0" + parseInt(partesComp[0], 10)).slice(-2),
        ano: valorParaTextoV3_(partesComp[1])
      };
    }
  }

  var vencimento = valorParaTextoV3_(cobranca && cobranca.vencimento).trim();
  var partesData = vencimento.split("/");
  if (partesData.length === 3) {
    return {
      mes: ("0" + parseInt(partesData[1], 10)).slice(-2),
      ano: valorParaTextoV3_(partesData[2])
    };
  }

  return {
    mes: Utilities.formatDate(new Date(), "America/Sao_Paulo", "MM"),
    ano: Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyy")
  };
}

function obterNomeMesFinanceiroV3_(mes) {
  return NOMES_MESES_FINANCEIRO_V3_[valorParaTextoV3_(mes)] || "MÊS";
}

function obterFraseMesFinanceiroV3_(cobranca) {
  var mesAno = obterMesAnoCobrancaFinanceiroV3_(cobranca);
  var base = FRASES_BASE_MENSAL_FINANCEIRO_V3_[mesAno.mes] || "Todo mês traz uma nova oportunidade de seguir em frente com confiança.";
  var complemento = FRASES_COMPLEMENTO_ANUAL_FINANCEIRO_V3_[mesAno.ano] || "Persistência hoje é sucesso amanhã. 🍀✨";
  return base + " " + complemento;
}

function classificarTipoItemCobrancaFinanceiroV3_(item) {
  var tipo = normalizarTextoV3_(item && item.tipoLancamento);
  if (tipo.indexOf("mensal") !== -1) return "mensalidade";
  if (tipo.indexOf("servic") !== -1) return "servico";
  return "produto";
}

function resumirItensCobrancaFinanceiroV3_(itens) {
  var resumo = {
    mensalidades: [],
    servicos: [],
    produtos: [],
    totalMensalidade: 0,
    totalServicos: 0,
    totalProdutos: 0
  };

  itens = itens || [];
  for (var i = 0; i < itens.length; i++) {
    var item = itens[i] || {};
    var tipoItem = classificarTipoItemCobrancaFinanceiroV3_(item);
    var valor = parseValorV3_(item.valorTotal);
    var linha = {
      descricao: valorParaTextoV3_(item.descricao || item.itemNome || item.tipoLancamento || "Lancamento"),
      quantidade: parseFloat(item.quantidade) || 0,
      unidade: valorParaTextoV3_(item.unidade || "un"),
      valorUnitario: parseValorV3_(item.valorUnitario),
      valorUnitarioFormatado: formatarValorV3_(item.valorUnitario),
      valorTotal: valor,
      valorTotalFormatado: formatarValorV3_(item.valorTotal),
      competencia: valorParaTextoV3_(item.competencia),
      referencia: valorParaTextoV3_(item.referencia)
    };

    if (tipoItem === "mensalidade") {
      resumo.mensalidades.push(linha);
      resumo.totalMensalidade += valor;
    } else if (tipoItem === "servico") {
      resumo.servicos.push(linha);
      resumo.totalServicos += valor;
    } else {
      resumo.produtos.push(linha);
      resumo.totalProdutos += valor;
    }
  }

  return resumo;
}

function montarMensagemDetalhadaWhatsAppCobrancaFinanceiroV3_(cobranca) {
  var linhas = [];
  var nome = cobranca && cobranca.clienteNome ? cobranca.clienteNome : "";
  var mesAno = obterMesAnoCobrancaFinanceiroV3_(cobranca);
  var nomeMes = obterNomeMesFinanceiroV3_(mesAno.mes);
  var resumo = resumirItensCobrancaFinanceiroV3_(cobranca && cobranca.itens ? cobranca.itens : []);

  linhas.push("💧 " + FINANCEIRO_EMPRESA_NOME_V3);
  linhas.push("");
  linhas.push("Prezado(a) " + nome + ", tudo bem?");
  linhas.push("Segue a cobranca referente a manutencao da piscina deste mes, incluindo produtos de tratamento.");
  linhas.push("");

  if (resumo.totalMensalidade > 0) {
    linhas.push("Mensalidade de " + nomeMes + ": " + formatarValorV3_(resumo.totalMensalidade));
  }
  if (resumo.servicos.length) {
    linhas.push("Servicos:");
    for (var i = 0; i < resumo.servicos.length; i++) {
      linhas.push("- " + resumo.servicos[i].descricao + (resumo.servicos[i].quantidade > 1 ? " x" + resumo.servicos[i].quantidade : "") + " " + resumo.servicos[i].valorTotalFormatado);
    }
  }
  if (resumo.produtos.length) {
    linhas.push("Produtos:");
    for (var j = 0; j < resumo.produtos.length; j++) {
      linhas.push("- " + resumo.produtos[j].descricao + (resumo.produtos[j].quantidade > 1 ? " x" + resumo.produtos[j].quantidade : "") + " " + resumo.produtos[j].valorTotalFormatado);
    }
  }
  if (!resumo.totalMensalidade && !resumo.servicos.length && !resumo.produtos.length) {
    linhas.push("Itens cobrados:");
    linhas.push("- " + formatarValorV3_(cobranca.valorTotal));
  }
  linhas.push("");
  linhas.push("Valor total: " + formatarValorV3_(cobranca.valorTotal));
  linhas.push("");
  linhas.push("📅 Vencimento: " + valorParaTextoV3_(cobranca.vencimento || "-"));
  linhas.push("💲 Dados para pagamento (Pix):");
  linhas.push("🔑 Chave: " + FINANCEIRO_PIX_CHAVE_V3);
  linhas.push("👤 Nome: " + FINANCEIRO_PIX_FAVORECIDO_V3);
  linhas.push("🏦 Banco: " + FINANCEIRO_PIX_BANCO_V3);
  linhas.push("");
  linhas.push("✨ Frase do mes:");
  linhas.push(obterFraseMesFinanceiroV3_(cobranca));
  if (cobranca.observacoes) {
    linhas.push("");
    linhas.push("Observacoes:");
    linhas.push(valorParaTextoV3_(cobranca.observacoes));
  }
  linhas.push("");
  linhas.push("Agradecemos pela confianca!");
  linhas.push("📞 " + FINANCEIRO_CONTATO_TELEFONE_V3);
  linhas.push("✉️ " + FINANCEIRO_CONTATO_EMAIL_V3);

  return linhas.join("\n");
}

function montarMensagemResumoWhatsAppCobrancaFinanceiroV3_(cobranca) {
  var nome = valorParaTextoV3_(cobranca && cobranca.clienteNome).trim() || "cliente";
  var mesAno = obterMesAnoCobrancaFinanceiroV3_(cobranca);
  var nomeMes = obterNomeMesFinanceiroV3_(mesAno.mes);
  var linhas = [
    "💧 " + FINANCEIRO_EMPRESA_NOME_V3,
    "",
    "Ola, " + nome + ".",
    "Segue sua cobranca de " + nomeMes + ".",
    "Valor total: " + formatarValorV3_(cobranca && cobranca.valorTotal),
    "Vencimento: " + valorParaTextoV3_(cobranca && cobranca.vencimento || "-")
  ];

  if (cobranca && cobranca.pdfUrl) {
    linhas.push("");
    linhas.push("PDF da cobranca:");
    linhas.push(cobranca.pdfUrl);
  }

  linhas.push("");
  linhas.push("Pix para pagamento:");
  linhas.push("Chave: " + FINANCEIRO_PIX_CHAVE_V3);
  linhas.push("");
  linhas.push("Qualquer duvida, ficamos a disposicao.");
  return linhas.join("\n");
}

function montarMensagemWhatsAppCobrancaFinanceiroV3_(cobranca) {
  if (cobranca && cobranca.pdfUrl) return montarMensagemResumoWhatsAppCobrancaFinanceiroV3_(cobranca);
  return montarMensagemDetalhadaWhatsAppCobrancaFinanceiroV3_(cobranca);
}

function construirLinkWhatsAppFinanceiroV3_(telefone, mensagem) {
  var numero = normalizarTelefoneWhatsappFinanceiroV3_(telefone);
  if (!numero) return "";
  return "https://wa.me/" + numero + "?text=" + encodeURIComponent(mensagem || "");
}

function montarDadosPdfCobrancaFinanceiroV3_(cobranca) {
  var mesAno = obterMesAnoCobrancaFinanceiroV3_(cobranca);
  var resumo = resumirItensCobrancaFinanceiroV3_(cobranca && cobranca.itens ? cobranca.itens : []);

  return {
    logoDataUri: getLogoDataUriFinanceiroV3_(),
    empresaNome: FINANCEIRO_EMPRESA_NOME_V3,
    empresaTelefone: FINANCEIRO_CONTATO_TELEFONE_V3,
    empresaEmail: FINANCEIRO_CONTATO_EMAIL_V3,
    pixChave: FINANCEIRO_PIX_CHAVE_V3,
    pixFavorecido: FINANCEIRO_PIX_FAVORECIDO_V3,
    pixBanco: FINANCEIRO_PIX_BANCO_V3,
    fraseMes: obterFraseMesFinanceiroV3_(cobranca),
    tituloCompetencia: obterNomeMesFinanceiroV3_(mesAno.mes) + "/" + mesAno.ano,
    cobranca: {
      id: valorParaTextoV3_(cobranca && cobranca.id),
      clienteNome: valorParaTextoV3_(cobranca && cobranca.clienteNome),
      clienteEmail: valorParaTextoV3_(cobranca && cobranca.clienteEmail),
      telefone: valorParaTextoV3_(cobranca && cobranca.telefone),
      competencia: valorParaTextoV3_(cobranca && cobranca.competencia),
      referencia: valorParaTextoV3_(cobranca && cobranca.referencia),
      vencimento: valorParaTextoV3_(cobranca && cobranca.vencimento),
      valorTotal: formatarValorV3_(cobranca && cobranca.valorTotal),
      observacoes: valorParaTextoV3_(cobranca && cobranca.observacoes)
    },
    resumo: {
      mensalidades: resumo.mensalidades,
      servicos: resumo.servicos,
      produtos: resumo.produtos,
      totalMensalidade: formatarValorV3_(resumo.totalMensalidade),
      totalServicos: formatarValorV3_(resumo.totalServicos),
      totalProdutos: formatarValorV3_(resumo.totalProdutos)
    }
  };
}

function gerarArquivoPdfCobrancaFinanceiroV3_(cobranca) {
  var template = HtmlService.createTemplateFromFile("CobrancaPdf");
  template.dados = montarDadosPdfCobrancaFinanceiroV3_(cobranca);

  var nomeArquivo = "Cobranca-" + valorParaTextoV3_(cobranca && cobranca.id) + ".pdf";
  var blob = template.evaluate().getBlob().getAs(MimeType.PDF).setName(nomeArquivo);

  try {
    if (cobranca && cobranca.pdfFileId) {
      var arquivoAntigo = DriveApp.getFileById(cobranca.pdfFileId);
      if (arquivoAntigo) arquivoAntigo.setTrashed(true);
    }
  } catch (errorArquivoAntigo) {}

  var arquivo = getPastaPdfFinanceiroV3_().createFile(blob);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    pdfUrl: arquivo.getUrl(),
    pdfFileId: arquivo.getId(),
    pdfGeradoEm: formatarDataHoraSistema_()
  };
}

function atualizarComunicacaoCobrancaFinanceiroV3_(cobrancaId, dados) {
  var aba = getAbaCobrancasFinanceiroV3_();
  var ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;

  var valores = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
  for (var i = 0; i < valores.length; i++) {
    if (valorParaTextoV3_(valores[i][0]) !== valorParaTextoV3_(cobrancaId)) continue;
    if (dados.linkWhatsApp !== undefined) aba.getRange(i + 2, 12).setValue(valorParaTextoV3_(dados.linkWhatsApp));
    if (dados.mensagemWhatsApp !== undefined) aba.getRange(i + 2, 13).setValue(valorParaTextoV3_(dados.mensagemWhatsApp));
    if (dados.pdfUrl !== undefined) aba.getRange(i + 2, 18).setValue(valorParaTextoV3_(dados.pdfUrl));
    if (dados.pdfFileId !== undefined) aba.getRange(i + 2, 19).setValue(valorParaTextoV3_(dados.pdfFileId));
    if (dados.pdfGeradoEm !== undefined) aba.getRange(i + 2, 20).setValue(valorParaTextoV3_(dados.pdfGeradoEm));
    return;
  }
}

function garantirPdfCobrancaFinanceiroV3_(cobranca) {
  if (!cobranca || !cobranca.id) return cobranca;
  if (!cobranca.pdfUrl) {
    var pdfInfo = gerarArquivoPdfCobrancaFinanceiroV3_(cobranca);
    cobranca.pdfUrl = pdfInfo.pdfUrl;
    cobranca.pdfFileId = pdfInfo.pdfFileId;
    cobranca.pdfGeradoEm = pdfInfo.pdfGeradoEm;
  }

  cobranca.mensagemWhatsApp = montarMensagemWhatsAppCobrancaFinanceiroV3_(cobranca);
  cobranca.linkWhatsApp = construirLinkWhatsAppFinanceiroV3_(cobranca.telefone, cobranca.mensagemWhatsApp);
  atualizarComunicacaoCobrancaFinanceiroV3_(cobranca.id, {
    linkWhatsApp: cobranca.linkWhatsApp,
    mensagemWhatsApp: cobranca.mensagemWhatsApp,
    pdfUrl: cobranca.pdfUrl,
    pdfFileId: cobranca.pdfFileId,
    pdfGeradoEm: cobranca.pdfGeradoEm
  });
  return cobranca;
}

function getPdfCobrancaFinanceiroV3(cobrancaId) {
  try {
    var cobrancas = getCobrancasFinanceiroV3_();
    var id = valorParaTextoV3_(cobrancaId);
    for (var i = 0; i < cobrancas.length; i++) {
      if (cobrancas[i].id !== id) continue;
      var cobranca = garantirPdfCobrancaFinanceiroV3_(cobrancas[i]);
      return {
        success: true,
        id: cobranca.id,
        pdfUrl: cobranca.pdfUrl,
        pdfFileId: cobranca.pdfFileId,
        pdfGeradoEm: cobranca.pdfGeradoEm
      };
    }
    return { success: false, message: "Cobranca nao encontrada." };
  } catch (error) {
    return { success: false, message: "Erro ao gerar PDF da cobranca: " + error.toString() };
  }
}

function getLinkWhatsAppCobrancaFinanceiroV3(cobrancaId) {
  try {
    var cobrancas = getCobrancasFinanceiroV3_();
    var id = valorParaTextoV3_(cobrancaId);

    for (var i = 0; i < cobrancas.length; i++) {
      if (cobrancas[i].id !== id) continue;
      var cobranca = garantirPdfCobrancaFinanceiroV3_(cobrancas[i]);
      return {
        success: true,
        id: cobranca.id,
        linkWhatsApp: cobranca.linkWhatsApp,
        mensagemWhatsApp: cobranca.mensagemWhatsApp,
        pdfUrl: cobranca.pdfUrl
      };
    }

    return { success: false, message: "Cobranca nao encontrada." };
  } catch (error) {
    return { success: false, message: "Erro ao obter link do WhatsApp: " + error.toString() };
  }
}

function marcarCobrancaFinanceiroEnviadaV3(cobrancaId) {
  try {
    var aba = getAbaCobrancasFinanceiroV3_();
    var dados = aba.getDataRange().getValues();
    var id = valorParaTextoV3_(cobrancaId);

    for (var i = 1; i < dados.length; i++) {
      if (valorParaTextoV3_(dados[i][0]) !== id) continue;
      if (statusPagoV3_(dados[i][10]) || statusCanceladoV3_(dados[i][10])) {
        return { success: false, message: "Nao foi possivel marcar o envio dessa cobranca." };
      }

      aba.getRange(i + 1, 11).setValue(FINANCEIRO_STATUS_COBRANCA_ENVIADA_V3);
      aba.getRange(i + 1, 15).setValue(formatarDataHoraSistema_());
      registrarHistoricoFinanceiroV3_({
        id: id + "-ENVIO",
        tipo: "Cobranca",
        cliente: valorParaTextoV3_(dados[i][2]),
        email: valorParaTextoV3_(dados[i][1]),
        telefone: valorParaTextoV3_(dados[i][3]),
        titulo: "Cobranca enviada",
        descricao: "Cobranca " + id + " marcada como enviada.",
        valor: formatarValorV3_(dados[i][8]),
        status: FINANCEIRO_STATUS_COBRANCA_ENVIADA_V3,
        dataRegistro: formatarDataHoraSistema_(),
        linkAcao: "tab:historico",
        observacoes: valorParaTextoV3_(dados[i][5])
      });
      SpreadsheetApp.flush();
      return { success: true, message: "Cobranca marcada como enviada." };
    }

    return { success: false, message: "Cobranca nao encontrada." };
  } catch (error) {
    return { success: false, message: "Erro ao atualizar envio: " + error.toString() };
  }
}

function atualizarStatusCobrancaFinanceiroV3(cobrancaId, status, observacoes) {
  try {
    var aba = getAbaCobrancasFinanceiroV3_();
    var dados = aba.getDataRange().getValues();
    var id = valorParaTextoV3_(cobrancaId);
    var statusDestino = valorParaTextoV3_(status || FINANCEIRO_STATUS_COBRANCA_PENDENTE_V3);

    for (var i = 1; i < dados.length; i++) {
      if (valorParaTextoV3_(dados[i][0]) !== id) continue;

      aba.getRange(i + 1, 11).setValue(statusDestino);
      if (observacoes) {
        var textoAtual = valorParaTextoV3_(aba.getRange(i + 1, 16).getValue());
        var novoTexto = (textoAtual ? textoAtual + "\n" : "") + formatarDataHoraSistema_() + " - " + observacoes;
        aba.getRange(i + 1, 16).setValue(novoTexto);
      }
      if (normalizarTextoV3_(statusDestino) === normalizarTextoV3_(FINANCEIRO_STATUS_COBRANCA_ENVIADA_V3) ||
          normalizarTextoV3_(statusDestino) === normalizarTextoV3_(FINANCEIRO_STATUS_COBRANCA_PAGA_V3)) {
        aba.getRange(i + 1, 15).setValue(formatarDataHoraSistema_());
      }

      if (normalizarTextoV3_(statusDestino) === normalizarTextoV3_(FINANCEIRO_STATUS_COBRANCA_CANCELADA_V3)) {
        var itens = [];
        try {
          itens = JSON.parse(dados[i][6] || "[]");
        } catch (errorJson) {
          itens = [];
        }

        var abaLancamentos = getAbaLancamentosFinanceirosV3_();
        var dadosLancamentos = abaLancamentos.getDataRange().getValues();
        for (var linha = 1; linha < dadosLancamentos.length; linha++) {
          var lancamentoId = valorParaTextoV3_(dadosLancamentos[linha][0]);
          for (var indice = 0; indice < itens.length; indice++) {
            if (valorParaTextoV3_(itens[indice].id) !== lancamentoId) continue;
            abaLancamentos.getRange(linha + 1, 15).setValue(FINANCEIRO_STATUS_LANCAMENTO_ABERTO_V3);
            abaLancamentos.getRange(linha + 1, 16).setValue("");
            break;
          }
        }
      }

      registrarHistoricoFinanceiroV3_({
        id: id + "-" + normalizarTextoV3_(statusDestino).toUpperCase(),
        tipo: "Cobranca",
        cliente: valorParaTextoV3_(dados[i][2]),
        email: valorParaTextoV3_(dados[i][1]),
        telefone: valorParaTextoV3_(dados[i][3]),
        titulo: "Status da cobranca atualizado",
        descricao: "Cobranca " + id + " atualizada para " + statusDestino + ".",
        valor: formatarValorV3_(dados[i][8]),
        status: statusDestino,
        dataRegistro: formatarDataHoraSistema_(),
        linkAcao: "tab:historico",
        observacoes: observacoes || valorParaTextoV3_(dados[i][5])
      });

      SpreadsheetApp.flush();
      return { success: true, message: "Status da cobranca atualizado." };
    }

    return { success: false, message: "Cobranca nao encontrada." };
  } catch (error) {
    return { success: false, message: "Erro ao atualizar cobranca: " + error.toString() };
  }
}

function getCobrancasPendentesFinanceiroV3_() {
  return getCobrancasFinanceiroV3_().filter(function(cobranca) {
    return !statusPagoV3_(cobranca.status) && !statusCanceladoV3_(cobranca.status);
  });
}

function enviarCobrancaWhatsAppFinanceiroV3(cobrancaId) {
  try {
    var id = valorParaTextoV3_(cobrancaId);
    var cobrancas = getCobrancasFinanceiroV3_();
    var cobranca = null;

    for (var i = 0; i < cobrancas.length; i++) {
      if (cobrancas[i].id === id) {
        cobranca = cobrancas[i];
        break;
      }
    }

    if (!cobranca) return { success: false, message: "Cobranca nao encontrada." };
    cobranca = garantirPdfCobrancaFinanceiroV3_(cobranca);

    var config = getConfiguracaoWhatsAppFinanceiroV3_();
    if (!config.ativo) {
      return {
        success: false,
        manual: true,
        linkWhatsApp: cobranca.linkWhatsApp,
        message: "Automacao do WhatsApp nao configurada. Use o envio manual."
      };
    }

    var telefone = normalizarTelefoneWhatsappFinanceiroV3_(cobranca.telefone);
    if (!telefone) return { success: false, message: "Cliente sem telefone valido para WhatsApp." };

    var payload;
    if (config.modoEnvio === "text" || !config.templateName) {
      payload = {
        messaging_product: "whatsapp",
        to: telefone,
        type: "text",
        text: {
          preview_url: false,
          body: cobranca.mensagemWhatsApp || montarMensagemWhatsAppCobrancaFinanceiroV3_(cobranca)
        }
      };
    } else {
      payload = {
        messaging_product: "whatsapp",
        to: telefone,
        type: "template",
        template: {
          name: config.templateName,
          language: { code: config.templateLanguage },
          components: [{
            type: "body",
            parameters: [
              { type: "text", text: cobranca.clienteNome || "Cliente" },
              { type: "text", text: cobranca.referencia || cobranca.id },
              { type: "text", text: formatarValorV3_(cobranca.valorTotal) },
              { type: "text", text: cobranca.vencimento || "-" }
            ]
          }]
        }
      };
    }

    var resposta = UrlFetchApp.fetch(config.apiUrl, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + config.token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var statusHttp = resposta.getResponseCode();
    var corpo = resposta.getContentText() || "";
    var json = {};

    try {
      json = JSON.parse(corpo || "{}");
    } catch (errorJson) {
      json = { raw: corpo };
    }

    if (statusHttp < 200 || statusHttp >= 300) {
      return {
        success: false,
        message: "Falha ao enviar cobranca pelo WhatsApp.",
        statusCode: statusHttp,
        response: json
      };
    }

    var aba = getAbaCobrancasFinanceiroV3_();
    var dados = aba.getDataRange().getValues();
    for (var linha = 1; linha < dados.length; linha++) {
      if (valorParaTextoV3_(dados[linha][0]) !== id) continue;
      aba.getRange(linha + 1, 11).setValue(FINANCEIRO_STATUS_COBRANCA_ENVIADA_V3);
      aba.getRange(linha + 1, 15).setValue(formatarDataHoraSistema_());
      break;
    }

    registrarHistoricoFinanceiroV3_({
      id: id + "-WHATSAPP",
      tipo: "Cobranca",
      cliente: cobranca.clienteNome,
      email: cobranca.clienteEmail,
      telefone: cobranca.telefone,
      titulo: "Cobranca enviada por WhatsApp",
      descricao: "Cobranca " + id + " enviada automaticamente por WhatsApp.",
      valor: cobranca.valorTotalFormatado,
      status: FINANCEIRO_STATUS_COBRANCA_ENVIADA_V3,
      dataRegistro: formatarDataHoraSistema_(),
      linkAcao: "tab:historico",
      observacoes: cobranca.referencia
    });

    SpreadsheetApp.flush();
    return {
      success: true,
      message: "Cobranca enviada automaticamente por WhatsApp.",
      response: json
    };
  } catch (error) {
    return { success: false, message: "Erro ao enviar cobranca por WhatsApp: " + error.toString() };
  }
}

function getResumoFinanceiroAdminV3_() {
  var clientes = getClientesFinanceiroV3_();
  var itensResp = getItensCobrancaFinanceiroV3();
  var assinaturasResp = getAssinaturasMensaisFinanceiroV3();
  var lancamentosAbertos = getLancamentosFinanceirosAbertosAgrupadosV3_();
  var cobrancasPendentes = getCobrancasPendentesFinanceiroV3_();
  var totalAberto = 0;

  for (var i = 0; i < cobrancasPendentes.length; i++) {
    totalAberto += cobrancasPendentes[i].valorTotal;
  }

  return {
    totalClientes: clientes.length,
    totalItens: itensResp && itensResp.itens ? itensResp.itens.length : 0,
    totalMensalidades: assinaturasResp && assinaturasResp.itens ? assinaturasResp.itens.length : 0,
    totalClientesComLancamento: lancamentosAbertos.length,
    totalCobrancasPendentes: cobrancasPendentes.length,
    valorAberto: totalAberto,
    valorAbertoFormatado: formatarValorV3_(totalAberto)
  };
}

function getFinanceiroAdminV3() {
  try {
    var itensResp = getItensCobrancaFinanceiroV3();
    var assinaturasResp = getAssinaturasMensaisFinanceiroV3();
    return {
      success: true,
      resumo: getResumoFinanceiroAdminV3_(),
      whatsappAutomacao: getConfiguracaoWhatsAppFinanceiroV3_(),
      clientes: getClientesFinanceiroV3_(),
      itens: itensResp && itensResp.itens ? itensResp.itens : [],
      mensalidades: assinaturasResp && assinaturasResp.itens ? assinaturasResp.itens : [],
      lancamentosAgrupados: getLancamentosFinanceirosAbertosAgrupadosV3_(),
      cobrancas: getCobrancasFinanceiroV3_()
    };
  } catch (error) {
    return {
      success: false,
      resumo: {},
      whatsappAutomacao: { ativo: false },
      clientes: [],
      itens: [],
      mensalidades: [],
      lancamentosAgrupados: [],
      cobrancas: [],
      message: "Erro ao carregar financeiro: " + error.toString()
    };
  }
}

function getHistoricoFinanceiroUsuarioV3_(emailLower, idsExistentes) {
  var itens = [];
  var ids = idsExistentes || {};
  var lancamentos = getLancamentosFinanceirosV3_();
  var cobrancas = getCobrancasFinanceiroV3_();

  for (var i = 0; i < lancamentos.length; i++) {
    var lancamento = lancamentos[i];
    if (lancamento.clienteEmail !== emailLower) continue;
    if (ids[lancamento.id]) continue;
    if (statusCanceladoV3_(lancamento.status)) continue;

    itens.push({
      id: lancamento.id,
      tipo: lancamento.tipoLancamento || "Consumo",
      cliente: lancamento.clienteNome,
      email: lancamento.clienteEmail,
      telefone: lancamento.telefone,
      titulo: lancamento.itemNome || lancamento.descricao || lancamento.tipoLancamento,
      descricao: lancamento.descricao + (lancamento.quantidade ? " | Quantidade: " + lancamento.quantidade : ""),
      valor: lancamento.valorTotalFormatado,
      status: lancamento.status,
      dataRegistro: lancamento.dataLancamento,
      linkAcao: "tab:historico",
      observacoes: lancamento.competencia ? "Competencia: " + lancamento.competencia : ""
    });
    ids[lancamento.id] = true;
  }

  for (var j = 0; j < cobrancas.length; j++) {
    var cobranca = cobrancas[j];
    if (cobranca.clienteEmail !== emailLower) continue;
    if (ids[cobranca.id]) continue;

    itens.push({
      id: cobranca.id,
      tipo: "Cobranca",
      cliente: cobranca.clienteNome,
      email: cobranca.clienteEmail,
      telefone: cobranca.telefone,
      titulo: cobranca.referencia || ("Cobranca " + cobranca.id),
      descricao: "Cobranca gerada com " + cobranca.quantidadeItens + " item(ns).",
      valor: cobranca.valorTotalFormatado,
      status: cobranca.status,
      dataRegistro: cobranca.dataGeracao,
      linkAcao: "tab:historico",
      observacoes: cobranca.competencia ? "Competencia: " + cobranca.competencia : cobranca.observacoes
    });
    ids[cobranca.id] = true;
  }

  return itens;
}
