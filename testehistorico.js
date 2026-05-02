// DebugHistorico.gs - Código completo para diagnosticar problema do histórico

function testarDiagnosticoHistorico() {
  console.log("=== INICIANDO DIAGNÓSTICO DO HISTÓRICO ===");
  
  var resultado = {
    orcamentosEncontrados: 0,
    orcamentosComEmailCorrespondente: 0,
    orcamentosDetalhes: [],
    problemaIdentificado: null,
    solucaoSugerida: null
  };
  
  try {
    // 1. Verificar aba Orcamentos
    var abaOrcamentos = getAbaOrcamentos();
    var dadosOrcamentos = abaOrcamentos.getDataRange().getValues();
    var cabecalhos = dadosOrcamentos[0];
    
    console.log("📊 Aba Orcamentos encontrada com", (dadosOrcamentos.length - 1), "registros");
    console.log("📋 Cabeçalhos:", cabecalhos);
    
    // 2. Listar todos os emails únicos na aba Orcamentos
    var emailsUnicos = {};
    for (var i = 1; i < dadosOrcamentos.length; i++) {
      var emailOrc = dadosOrcamentos[i][2] ? dadosOrcamentos[i][2].toString().trim().toLowerCase() : "";
      if (emailOrc) {
        emailsUnicos[emailOrc] = (emailsUnicos[emailOrc] || 0) + 1;
        resultado.orcamentosDetalhes.push({
          linha: i + 1,
          id: dadosOrcamentos[i][0],
          cliente: dadosOrcamentos[i][1],
          email: emailOrc,
          status: dadosOrcamentos[i][10],
          data: dadosOrcamentos[i][4]
        });
        resultado.orcamentosEncontrados++;
      }
    }
    
    console.log("📧 Emails únicos encontrados nos orçamentos:", Object.keys(emailsUnicos));
    console.log("📝 Detalhes dos orçamentos:", resultado.orcamentosDetalhes);
    
    // 3. Verificar usuários cadastrados
    var abaUsuarios = getAbaUsuarios();
    var dadosUsuarios = abaUsuarios.getDataRange().getValues();
    
    console.log("👥 Usuários cadastrados:", dadosUsuarios.length - 1);
    
    for (var i = 1; i < dadosUsuarios.length; i++) {
      var emailUser = dadosUsuarios[i][0] ? dadosUsuarios[i][0].toString().trim().toLowerCase() : "";
      var nomeUser = dadosUsuarios[i][2] || "";
      console.log("  -", emailUser, "(" + nomeUser + ")");
      
      // Verificar se este usuário tem orçamentos
      var orcamentosUser = [];
      for (var j = 0; j < resultado.orcamentosDetalhes.length; j++) {
        if (resultado.orcamentosDetalhes[j].email === emailUser) {
          orcamentosUser.push(resultado.orcamentosDetalhes[j]);
        }
      }
      
      if (orcamentosUser.length > 0) {
        resultado.orcamentosComEmailCorrespondente += orcamentosUser.length;
        console.log("  ✅ Usuário", emailUser, "tem", orcamentosUser.length, "orçamento(s)");
      } else {
        console.log("  ⚠️ Usuário", emailUser, "NÃO tem orçamentos na aba Orcamentos");
      }
    }
    
    // 4. Verificar a função getOrcamentosCliente diretamente
    console.log("\n=== TESTANDO getOrcamentosCliente PARA CADA USUÁRIO ===");
    
    for (var i = 1; i < dadosUsuarios.length; i++) {
      var emailTeste = dadosUsuarios[i][0] ? dadosUsuarios[i][0].toString().trim() : "";
      var nomeTeste = dadosUsuarios[i][2] || "";
      
      if (emailTeste) {
        var orcamentosCliente = getOrcamentosCliente(emailTeste);
        console.log("📌 Usuário:", nomeTeste, "(" + emailTeste + ")");
        console.log("   Retornou", orcamentosCliente.length, "orçamento(s)");
        
        if (orcamentosCliente.length === 0 && resultado.orcamentosDetalhes.some(function(o) { return o.email === emailTeste.toLowerCase(); })) {
          console.log("   ❌ PROBLEMA DETECTADO: Usuário tem orçamento na planilha mas getOrcamentosCliente retornou vazio!");
          resultado.problemaIdentificado = "getOrcamentosCliente não está encontrando orçamentos para o email: " + emailTeste;
          resultado.solucaoSugerida = "Verificar se o campo de email na planilha está sendo lido corretamente (pode ter espaços extras ou formatação diferente)";
        }
      }
    }
    
    // 5. Verificar formatação dos emails (possíveis espaços ou caracteres invisíveis)
    console.log("\n=== VERIFICANDO FORMATAÇÃO DOS EMAILS ===");
    for (var i = 1; i < dadosOrcamentos.length; i++) {
      var emailRaw = dadosOrcamentos[i][2];
      var emailTrim = emailRaw ? emailRaw.toString().trim() : "";
      var emailLower = emailTrim.toLowerCase();
      
      if (emailRaw && emailRaw.toString() !== emailTrim) {
        console.log("⚠️ Email com espaços extras na linha", i+1, ":", "'" + emailRaw + "'");
        // Corrigir automaticamente
        abaOrcamentos.getRange(i+1, 3).setValue(emailTrim);
        console.log("   ✅ Corrigido para:", "'" + emailTrim + "'");
      }
    }
    
    // 6. Resumo do diagnóstico
    console.log("\n=== RESUMO DO DIAGNÓSTICO ===");
    console.log("Total de orçamentos na planilha:", resultado.orcamentosEncontrados);
    console.log("Orçamentos com email correspondente a usuários:", resultado.orcamentosComEmailCorrespondente);
    
    if (resultado.orcamentosEncontrados === 0) {
      resultado.problemaIdentificado = "Nenhum orçamento encontrado na planilha";
      resultado.solucaoSugerida = "Criar orçamentos de exemplo ou verificar se a aba 'Orcamentos' foi criada corretamente";
    } else if (resultado.orcamentosComEmailCorrespondente === 0) {
      resultado.problemaIdentificado = "Nenhum orçamento tem email correspondente a um usuário cadastrado";
      resultado.solucaoSugerida = "Verificar se os emails nos orçamentos estão iguais aos emails cadastrados na aba 'Usuarios'";
    } else {
      resultado.problemaIdentificado = "Sistema aparentemente normal. Verificar front-end (JavaScript)";
      resultado.solucaoSugerida = "Executar o código de diagnóstico no front-end (DebugHistoricoFrontend)";
    }
    
    return resultado;
    
  } catch (error) {
    console.error("❌ Erro no diagnóstico:", error);
    return {
      success: false,
      problemaIdentificado: "Erro ao executar diagnóstico: " + error.toString(),
      solucaoSugerida: "Verificar se todas as funções dependentes existem (getAbaOrcamentos, getAbaUsuarios, getOrcamentosCliente)"
    };
  }
}

// Função para testar no front-end (adicione no Script.html temporariamente)
function getDebugInfoFrontend() {
  try {
    var abaOrcamentos = getAbaOrcamentos();
    var dados = abaOrcamentos.getDataRange().getValues();
    
    var usuariosAba = getAbaUsuarios();
    var dadosUsuarios = usuariosAba.getDataRange().getValues();
    
    var usuariosLogados = [];
    for (var i = 1; i < dadosUsuarios.length; i++) {
      usuariosLogados.push({
        email: dadosUsuarios[i][0] || "",
        nome: dadosUsuarios[i][2] || ""
      });
    }
    
    var orcamentosRaw = [];
    for (var i = 1; i < dados.length; i++) {
      orcamentosRaw.push({
        linha: i,
        id: dados[i][0],
        cliente: dados[i][1],
        email: dados[i][2] ? dados[i][2].toString().trim() : "",
        status: dados[i][10]
      });
    }
    
    return {
      totalOrcamentos: dados.length - 1,
      usuarios: usuariosLogados,
      orcamentos: orcamentosRaw,
      mensagem: "Execute testarDiagnosticoHistorico() no Apps Script para diagnóstico completo"
    };
    
  } catch (error) {
    return { error: error.toString() };
  }
}

// Função para corrigir emails problemáticos
function corrigirEmailsOrcamentos() {
  try {
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var corrigidos = 0;
    
    for (var i = 1; i < dados.length; i++) {
      var emailAtual = dados[i][2];
      if (emailAtual && typeof emailAtual === 'string') {
        var emailCorrigido = emailAtual.trim().toLowerCase();
        if (emailAtual !== emailCorrigido) {
          aba.getRange(i + 1, 3).setValue(emailCorrigido);
          corrigidos++;
          console.log("✅ Corrigido email na linha", i+1, ":", "'" + emailAtual + "' -> '" + emailCorrigido + "'");
        }
      }
    }
    
    SpreadsheetApp.flush();
    return {
      success: true,
      corrigidos: corrigidos,
      message: corrigidos + " email(s) corrigido(s)"
    };
    
  } catch (error) {
    return {
      success: false,
      message: "Erro: " + error.toString()
    };
  }
}

function corrigirAbaOrcamentos() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("Orcamentos");
    
    if (!aba) {
      console.log("❌ Aba Orcamentos não encontrada!");
      return;
    }
    
    // Obter dados atuais
    var dados = aba.getDataRange().getValues();
    var cabecalhosAtuais = dados[0];
    
    console.log("📋 Cabeçalhos atuais:", cabecalhosAtuais);
    
    // Definir cabeçalhos corretos
    var cabecalhosCorretos = [
      "ID", "Cliente", "Email", "Telefone", "DataSolicitacao", 
      "TipoServico", "Descricao", "ImagemURL", "ArquivoURL", 
      "Valor", "Status", "DataValidade", "Observacoes", "ArquivoNome"
    ];
    
    // Verificar se já está correto
    var precisaCorrigir = false;
    for (var i = 0; i < cabecalhosCorretos.length; i++) {
      if (cabecalhosAtuais[i] !== cabecalhosCorretos[i]) {
        precisaCorrigir = true;
        break;
      }
    }
    
    if (!precisaCorrigir) {
      console.log("✅ Aba Orcamentos já está com estrutura correta!");
      return;
    }
    
    // Backup dos dados existentes
    var dadosExistentes = [];
    for (var i = 1; i < dados.length; i++) {
      dadosExistentes.push(dados[i]);
    }
    
    // Limpar a aba e recriar com cabeçalhos corretos
    aba.clear();
    aba.getRange("A1:N1").setValues([cabecalhosCorretos]);
    aba.setFrozenRows(1);
    
    // Restaurar dados com mapeamento correto
    for (var i = 0; i < dadosExistentes.length; i++) {
      var linha = dadosExistentes[i];
      var novaLinha = [];
      
      // Mapear colunas antigas para novas baseado na posição
      // Assumindo que os dados estão em ordem aproximada
      novaLinha[0] = linha[0] || "";  // ID
      novaLinha[1] = linha[1] || "";  // Cliente
      novaLinha[2] = linha[2] || "";  // Email
      novaLinha[3] = linha[3] || "";  // Telefone
      novaLinha[4] = linha[4] || "";  // DataSolicitacao
      novaLinha[5] = linha[5] || "";  // TipoServico
      novaLinha[6] = linha[6] || "";  // Descricao
      novaLinha[7] = linha[7] || "";  // ImagemURL
      novaLinha[8] = linha[8] || "";  // ArquivoURL
      novaLinha[9] = linha[11] || "Aguardando";  // Valor (antes coluna J, agora K)
      novaLinha[10] = linha[9] || "Solicitado";  // Status (antes coluna I, agora K)
      novaLinha[11] = linha[12] || "";  // DataValidade (antes coluna J, agora L)
      novaLinha[12] = linha[10] || "";  // Observacoes (antes coluna K, agora M)
      novaLinha[13] = linha[13] || "";  // ArquivoNome
      
      aba.appendRow(novaLinha);
    }
    
    SpreadsheetApp.flush();
    
    console.log("✅ Aba Orcamentos corrigida com sucesso!");
    console.log("📊", dadosExistentes.length, "registros migrados");
    
    return {
      success: true,
      message: "Aba Orcamentos corrigida!",
      registros: dadosExistentes.length
    };
    
  } catch (error) {
    console.error("❌ Erro:", error);
    return { success: false, message: error.toString() };
  }
}