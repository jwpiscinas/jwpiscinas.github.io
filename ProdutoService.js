// ProdutoService.gs - Gerenciamento de Produtos (IDs numéricos)

// Obter aba de produtos
function getAbaProdutos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Produtos");
  if (!aba) {
    aba = planilha.insertSheet("Produtos");
    aba.getRange("A1:L1").setValues([[
      "ID", "Nome", "Categoria", "Descricao", "Preco", "Estoque", "ImagemURL", "ManualURL", "ManualNome", "PrecoComDesconto", "ExibirManual", "ManualVideoURL"
    ]]);
    
    var produtosExemplo = [
      [1, "Cloro Granulado 1kg", "Químicos", 
       "Cloro para tratamento de piscinas", 45.90, 50, "https://images.unsplash.com/photo-1583485764546-bb9aa2d8d0d2?w=400", "", "Manual do Cloro", null, false],
      [2, "Aspirador de Piscina", "Equipamentos", 
       "Aspirador manual para limpeza", 189.90, 15, "https://images.unsplash.com/photo-1584697964403-4f5e6e3c4c0a?w=400", "", "Manual do Aspirador", null, false],
      [3, "Kit Teste de Ph", "Químicos",
       "Kit completo para teste de pH e cloro", 89.90, 30, "https://images.unsplash.com/photo-1581093458791-9f3c3b5b6c7d?w=400", "", "Manual do Kit", null, false]
    ];
    
    for (var i = 0; i < produtosExemplo.length; i++) {
      aba.appendRow(produtosExemplo[i]);
    }
  }
  garantirCabecalhosProdutos_(aba);
  return aba;
}

function garantirCabecalhosProdutos_(aba) {
  var cabecalhos = ["ID", "Nome", "Categoria", "Descricao", "Preco", "Estoque", "ImagemURL", "ManualURL", "ManualNome", "PrecoComDesconto", "ExibirManual", "ManualVideoURL"];
  var atuais = aba.getRange(1, 1, 1, cabecalhos.length).getValues()[0];
  var precisaAtualizar = false;

  for (var i = 0; i < cabecalhos.length; i++) {
    if (!atuais[i]) {
      atuais[i] = cabecalhos[i];
      precisaAtualizar = true;
    }
  }

  if (precisaAtualizar) {
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([atuais]);
  }

  try {
    var regra = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    aba.getRange(2, 11, Math.max(aba.getMaxRows() - 1, 1), 1).setDataValidation(regra);
  } catch (error) {}
}

// Obter produtos
function parseNumeroProduto_(valor) {
  if (typeof valor === "number") return valor;
  if (!valor) return 0;
  var texto = valor.toString().replace(/\s/g, "").replace("R$", "");
  if (texto.indexOf(",") !== -1) {
    texto = texto.replace(/\./g, "").replace(",", ".");
  }
  var numero = parseFloat(texto);
  return isNaN(numero) ? 0 : numero;
}

function getProdutos() {
  try {
    var aba = getAbaProdutos();
    var dados = aba.getDataRange().getValues();
    var produtos = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      
      if (linha[0]) {
        produtos.push({
          id: linha[0] ? linha[0].toString() : "",  // ID como string
          nome: linha[1] || "",
          categoria: linha[2] || "",
          descricao: linha[3] || "",
          preco: parseNumeroProduto_(linha[4]),
          estoque: typeof linha[5] === 'number' ? linha[5] : parseInt(linha[5]) || 0,
          imagemURL: linha[6] || "",
          manualURL: linha[7] || "",
          manualNome: linha[8] || "",
          precoComDesconto: linha[9] ? parseNumeroProduto_(linha[9]) : null,
          exibirManual: linha[10] === true || (linha[10] && linha[10].toString().toLowerCase() === "true"),
          manualVideoURL: linha[11] || ""
        });
      }
    }
    
    console.log("✅ Produtos carregados:", produtos);
    return produtos;
    
  } catch (error) {
    console.error("Erro getProdutos:", error);
    return [];
  }
}

// Atualizar preço do produto
function atualizarPrecoProduto(produtoId, novoPreco) {
  try {
    console.log("🔧 Atualizando preço do produto ID:", produtoId, "para:", novoPreco);
    
    var aba = getAbaProdutos();
    var dados = aba.getDataRange().getValues();
    var linhaEncontrada = -1;
    
    // Comparar como string para funcionar com "1" e 1
    var produtoIdStr = produtoId.toString();
    
    for (var i = 1; i < dados.length; i++) {
      var idProduto = dados[i][0] ? dados[i][0].toString() : "";
      if (idProduto === produtoIdStr) {
        linhaEncontrada = i + 1;
        break;
      }
    }
    
    if (linhaEncontrada > 0) {
      aba.getRange(linhaEncontrada, 5).setValue(parseFloat(novoPreco));
      SpreadsheetApp.flush();
      
      console.log("✅ Preço atualizado na linha:", linhaEncontrada);
      return {
        success: true,
        message: "Preço atualizado com sucesso!"
      };
    } else {
      console.log("❌ Produto não encontrado ID:", produtoId);
      return {
        success: false,
        message: "Produto não encontrado."
      };
    }
    
  } catch (error) {
    console.error("❌ Erro em atualizarPrecoProduto:", error);
    return {
      success: false,
      message: "Erro ao atualizar preço: " + error.toString()
    };
  }
}

function atualizarPrecoProdutoCampoV3(produtoId, novoPreco, campo) {
  try {
    var campoNormalizado = campo ? campo.toString().trim().toLowerCase() : "preco";
    var coluna = campoNormalizado === "desconto" || campoNormalizado === "precocomdesconto" ? 10 : 5;

    var aba = getAbaProdutos();
    var dados = aba.getDataRange().getValues();
    var produtoIdStr = produtoId ? produtoId.toString() : "";

    for (var i = 1; i < dados.length; i++) {
      var idProduto = dados[i][0] ? dados[i][0].toString() : "";
      if (idProduto !== produtoIdStr) continue;

      var valorNumerico = parseFloat(novoPreco);
      if (isNaN(valorNumerico) || valorNumerico < 0) {
        return { success: false, message: "Preco invalido." };
      }

      aba.getRange(i + 1, coluna).setValue(valorNumerico);
      SpreadsheetApp.flush();

      return {
        success: true,
        message: campoNormalizado === "desconto" || campoNormalizado === "precocomdesconto"
          ? "Preco com desconto atualizado com sucesso!"
          : "Preco principal atualizado com sucesso!"
      };
    }

    return { success: false, message: "Produto nao encontrado." };
  } catch (error) {
    return { success: false, message: "Erro ao atualizar preco: " + error.toString() };
  }
}

// Registrar compra
function registrarCompra(dados) {
  try {
    console.log("🛒 Registrando compra:", dados);
    
    var abaCompras = getAbaCompras();
    var idCompra = "COMP-" + new Date().getTime().toString().slice(-8);
    
    var produtosArray;
    try {
      produtosArray = JSON.parse(dados.produtos);
    } catch (e) {
      console.error("Erro ao parsear produtos:", e);
      return { success: false, message: "Erro ao processar produtos" };
    }
    
    var valorTotal = 0;
    var quantidadeTotal = 0;
    
    for (var i = 0; i < produtosArray.length; i++) {
      var item = produtosArray[i];
      var preco = typeof item.preco === 'number' ? item.preco : parseFloat(item.preco) || 0;
      var quantidade = typeof item.quantidade === 'number' ? item.quantidade : parseInt(item.quantidade) || 0;
      var disponibilidade = verificarEstoqueProduto_(item.id, quantidade);

      if (!disponibilidade.success) {
        return disponibilidade;
      }

      valorTotal += (preco * quantidade);
      quantidadeTotal += quantidade;
    }
    
    var novaLinha = [
      idCompra,
      dados.cliente,
      dados.email,
      formatarDataHoraSistema_(),
      JSON.stringify(produtosArray),
      quantidadeTotal,
      valorTotal,
      "Pendente",
      "Processando"
    ];
    
    abaCompras.appendRow(novaLinha);
    SpreadsheetApp.flush();
    
    atualizarEstoque(produtosArray);

    try {
      registrarHistorico({
        id: idCompra,
        tipo: "Compra",
        cliente: dados.cliente,
        email: dados.email,
        titulo: produtosArray.length === 1 ? produtosArray[0].nome : produtosArray.length + " produtos",
        descricao: "Compra de " + quantidadeTotal + " item(ns)",
        valor: "R$ " + valorTotal.toFixed(2),
        status: "Pendente",
        linkAcao: "tab:historico",
        observacoes: "Status da entrega: Processando"
      });
    } catch (histError) {
      console.log("Erro ao registrar compra no historico:", histError);
    }
    
    return { 
      success: true, 
      message: "✅ Compra realizada com sucesso!",
      id: idCompra,
      valorTotal: valorTotal.toFixed(2)
    };
    
  } catch (error) {
    console.error("Erro registrarCompra:", error);
    return { success: false, message: "Erro ao processar compra: " + error.toString() };
  }
}

function verificarEstoqueProduto_(produtoId, quantidade) {
  var aba = getAbaProdutos();
  var dados = aba.getDataRange().getValues();
  var produtoIdStr = produtoId ? produtoId.toString() : "";

  for (var i = 1; i < dados.length; i++) {
    var idProduto = dados[i][0] ? dados[i][0].toString() : "";

    if (idProduto === produtoIdStr) {
      var estoqueAtual = typeof dados[i][5] === 'number' ? dados[i][5] : parseInt(dados[i][5]) || 0;

      if (quantidade < 1) {
        return { success: false, message: "Quantidade invalida." };
      }

      if (estoqueAtual <= 0) {
        return { success: false, message: "Produto sem estoque disponivel." };
      }

      if (quantidade > estoqueAtual) {
        return { success: false, message: "Quantidade maior que o estoque disponivel." };
      }

      return { success: true };
    }
  }

  return { success: false, message: "Produto nao encontrado." };
}

// Atualizar estoque
function atualizarEstoque(produtos) {
  try {
    var aba = getAbaProdutos();
    var dados = aba.getDataRange().getValues();
    
    for (var i = 0; i < produtos.length; i++) {
      var itemCompra = produtos[i];
      var idCompraStr = itemCompra.id.toString();
      
      for (var j = 1; j < dados.length; j++) {
        var linha = dados[j];
        var idProduto = linha[0] ? linha[0].toString() : "";
        
        if (idProduto === idCompraStr) {
          var estoqueAtual = linha[5];
          var estoqueNumero = typeof estoqueAtual === 'number' ? estoqueAtual : parseInt(estoqueAtual) || 0;
          var quantidadeCompra = typeof itemCompra.quantidade === 'number' ? itemCompra.quantidade : parseInt(itemCompra.quantidade) || 0;
          
          var novoEstoque = estoqueNumero - quantidadeCompra;
          if (novoEstoque < 0) novoEstoque = 0;
          
          aba.getRange(j + 1, 6).setValue(novoEstoque);
          console.log("✅ Estoque atualizado. Produto:", idProduto, "Novo estoque:", novoEstoque);
          break;
        }
      }
    }
    
    SpreadsheetApp.flush();
    
  } catch (error) {
    console.error("Erro atualizarEstoque:", error);
  }
}

// Obter aba de compras
function getAbaCompras() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Compras");
  if (!aba) {
    aba = planilha.insertSheet("Compras");
    aba.getRange("A1:J1").setValues([[
      "ID", "Cliente", "Email", "DataCompra", "Produtos", "QuantidadeTotal", "ValorTotal", "StatusPagamento", "StatusEntrega", "ObservacoesAdmin"
    ]]);
  } else if (!aba.getRange(1, 10).getValue()) {
    aba.getRange(1, 10).setValue("ObservacoesAdmin");
  }
  return aba;
}

// Obter compras do cliente
function getComprasCliente(email) {
  try {
    var aba = getAbaCompras();
    var dados = aba.getDataRange().getValues();
    var compras = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailCompra = linha[2];
      
      if (emailCompra && emailCompra.toString().trim().toLowerCase() === email.toLowerCase()) {
        compras.push({
          id: linha[0] || "",
          dataCompra: linha[3] || "",
          produtos: linha[4] || "",
          quantidadeTotal: linha[5] || 0,
          valorTotal: linha[6] ? "R$ " + parseFloat(linha[6]).toFixed(2) : "0.00",
          statusPagamento: linha[7] || "Pendente",
          statusEntrega: linha[8] || "Processando"
        });
      }
    }
    
    return compras;
    
  } catch (error) {
    console.error("Erro getComprasCliente:", error);
    return [];
  }
}

// Função para recriar a aba Produtos com IDs numéricos (executar uma vez)
function recriarAbaProdutosNumericos(confirmacao) {
  if (confirmacao !== "CONFIRMAR_PRODUTOS") {
    return {
      success: false,
      message: "Protecao ativa: a aba Produtos nao sera recriada sem confirmacao explicita."
    };
  }

  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var abaExistente = planilha.getSheetByName("Produtos");
    
    if (abaExistente) {
      planilha.deleteSheet(abaExistente);
    }
    
    var aba = planilha.insertSheet("Produtos");
    aba.getRange("A1:L1").setValues([[
      "ID", "Nome", "Categoria", "Descricao", "Preco", "Estoque", "ImagemURL", "ManualURL", "ManualNome", "PrecoComDesconto", "ExibirManual", "ManualVideoURL"
    ]]);
    
    var produtos = [
      [1, "Cloro Granulado 1kg", "Químicos", 
       "Cloro para tratamento de piscinas", 45.90, 50, 
       "https://images.unsplash.com/photo-1583485764546-bb9aa2d8d0d2?w=400", "", "Manual do Cloro", null],
      [2, "Aspirador de Piscina", "Equipamentos", 
       "Aspirador manual para limpeza", 189.90, 15, 
       "https://images.unsplash.com/photo-1584697964403-4f5e6e3c4c0a?w=400", "", "Manual do Aspirador", null],
      [3, "Kit Teste de Ph", "Químicos",
       "Kit completo para teste de pH e cloro", 89.90, 30, 
       "https://images.unsplash.com/photo-1581093458791-9f3c3b5b6c7d?w=400", "", "Manual do Kit", null],
      [4, "Escova para Piscina", "Equipamentos",
       "Escova para limpeza de paredes e fundo", 35.90, 20,
       "https://images.unsplash.com/photo-1583485764546-bb9aa2d8d0d2?w=400", "", "Manual da Escova", null]
    ];
    
    for (var i = 0; i < produtos.length; i++) {
      aba.appendRow(produtos[i]);
    }
    
    aba.setFrozenRows(1);
    SpreadsheetApp.flush();
    
    console.log("✅ Aba Produtos recriada com IDs numéricos!");
    return { success: true, message: "Aba Produtos recriada com IDs 1,2,3,4!" };
    
  } catch (error) {
    console.error("Erro:", error);
    return { success: false, message: error.toString() };
  }
}
