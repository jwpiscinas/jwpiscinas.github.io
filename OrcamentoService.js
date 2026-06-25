// === ORCAMENTO SERVICE - COMPLETO ===

// OBTER ABA DE ORÇAMENTOS
function getAbaOrcamentos() {
  try {
    console.log("📋 Obtendo aba de orçamentos...");
    
    if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
      return getFirestoreSheetAdapter_("Orcamentos");
    }

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("Orcamentos");
    
    if (!aba) {
      console.log("📁 Criando aba Orcamentos...");
      aba = planilha.insertSheet("Orcamentos");
      
      // Definir cabeçalhos
      var cabecalhos = [[
        "ID",                 // A
        "Cliente",            // B
        "Email",              // C
        "Telefone",           // D
        "DataSolicitacao",    // E
        "TipoServico",        // F
        "Descricao",          // G
        "ImagemURL",          // H
        "ArquivoURL",         // I
        "Valor",              // J
        "Status",             // K
        "DataValidade",       // L
        "Observacoes",        // M
        "ArquivoNome"         // N
      ]];
      
      aba.getRange("A1:N1").setValues(cabecalhos);
      aba.setFrozenRows(1);
      
      console.log("✅ Aba Orcamentos criada com sucesso!");
    }
    
    return aba;
    
  } catch (error) {
    console.error("❌ Erro em getAbaOrcamentos:", error);
    throw error;
  }
}

// GERAR ID ÚNICO PARA ORÇAMENTO
function gerarIdOrcamento() {
  var timestamp = new Date().getTime();
  var random = Math.floor(Math.random() * 1000);
  return "ORC-" + timestamp.toString().slice(-6) + "-" + random;
}

function solicitarOrcamento(dados) {
  console.log("📝 === SOLICITANDO ORÇAMENTO ===");
  console.log("Dados recebidos:", dados);
  
  try {
    // VALIDAÇÕES
    if (!dados || !dados.cliente || !dados.email || !dados.tipoServico || !dados.descricao) {
      console.log("❌ Dados incompletos");
      return {
        success: false,
        message: "Preencha todos os campos obrigatórios!"
      };
    }
    
    if (dados.descricao.length < 10) {
      return {
        success: false,
        message: "A descrição deve ter pelo menos 10 caracteres!"
      };
    }
    
    var aba = getAbaOrcamentos();
    var id = gerarIdOrcamento();
    var dataAtual = new Date();
    var dataSolicitacao = formatarDataHoraSistema_(dataAtual);
    
    var dataValidade = new Date();
    dataValidade.setDate(dataValidade.getDate() + 30);
    var dataValidadeStr = formatarDataSistema_(dataValidade);
    
    // COLUNAS CORRETAS:
    // 0:ID, 1:Cliente, 2:Email, 3:Telefone, 4:DataSolicitacao
    // 5:TipoServico, 6:Descricao, 7:ImagemURL, 8:ArquivoURL
    // 9:Valor, 10:Status, 11:DataValidade, 12:Observacoes, 13:ArquivoNome
    
    var novaLinha = [
      id,                                   // 0: ID
      dados.cliente,                        // 1: Cliente
      dados.email,                          // 2: Email
      dados.telefone || "",                  // 3: Telefone
      dataSolicitacao,                       // 4: DataSolicitacao
      dados.tipoServico,                     // 5: TipoServico
      dados.descricao,                       // 6: Descricao
      dados.imagemURL || "",                  // 7: ImagemURL
      dados.arquivoURL || "",                 // 8: ArquivoURL
      "Aguardando",                          // 9: Valor
      "Solicitado",                          // 10: Status
      dataValidadeStr,                       // 11: DataValidade
      dados.observacoes || "",                // 12: Observacoes
      dados.arquivoNome || ""                 // 13: ArquivoNome
    ];
    
    console.log("📌 Inserindo orçamento ID:", id);
    aba.appendRow(novaLinha);
    SpreadsheetApp.flush();
    
    // =============================================
    // REGISTRAR NO HISTÓRICO GERAL (NOVO)
    // =============================================
    try {
      registrarHistorico({
        id: id,
        tipo: "Orçamento",
        cliente: dados.cliente,
        email: dados.email,
        telefone: dados.telefone || "",
        titulo: dados.tipoServico,
        descricao: dados.descricao,
        valor: "Aguardando",
        status: "Solicitado",
        linkAcao: "tab:orcamentos",
        observacoes: dados.observacoes || ""
      });
      console.log("✅ Registrado no histórico geral");
    } catch (histError) {
      console.log("⚠️ Erro ao registrar no histórico:", histError);
    }
    
    // Enviar email de confirmação
    try {
      enviarEmailConfirmacaoOrcamento(dados.email, dados.cliente, id, dados.tipoServico);
      console.log("✅ Email de confirmação enviado");
    } catch (emailError) {
      console.log("⚠️ Erro ao enviar email:", emailError);
    }
    
    return {
      success: true,
      message: "✅ Orçamento solicitado com sucesso!\n\nID: " + id + "\n\nAcompanhe o status no seu histórico.",
      id: id
    };
    
  } catch (error) {
    console.error("❌ Erro em solicitarOrcamento:", error);
    return {
      success: false,
      message: "Erro ao solicitar orçamento: " + error.toString()
    };
  }
}
// ENVIAR EMAIL DE CONFIRMAÇÃO DE ORÇAMENTO
function enviarEmailConfirmacaoOrcamento(email, nome, idOrcamento, tipoServico) {
  try {
    var assunto = "Orçamento Solicitado - JW Piscinas";
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <h3>Olá, ${nome}!</h3>
        <p>Sua solicitação de orçamento foi recebida com sucesso.</p>
        
        <p><strong>Detalhes da Solicitação:</strong></p>
        <ul>
          <li><strong>ID do Orçamento:</strong> ${idOrcamento}</li>
          <li><strong>Tipo de Serviço:</strong> ${tipoServico}</li>
          <li><strong>Data da Solicitação:</strong> ${new Date().toLocaleString('pt-BR')}</li>
          <li><strong>Validade do Orçamento:</strong> 30 dias</li>
        </ul>
        
        <p><strong>Próximos passos:</strong></p>
        <ol>
          <li>Nossa equipe analisará sua solicitação</li>
          <li>Em até 48 horas úteis retornaremos com o orçamento</li>
          <li>Você receberá um e-mail com todas as informações</li>
          <li>Acompanhe o status no aplicativo JW Piscinas</li>
        </ol>
        
        <p>Ficamos à disposição para qualquer dúvida.</p>
        
        <p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>
        <hr>
        <p style="font-size: 12px; color: #666;">
          ID do Orçamento: ${idOrcamento}
        </p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      htmlBody: corpo
    });
    
    console.log("✅ Email de confirmação enviado para:", email);
    
  } catch (error) {
    console.error("❌ Erro ao enviar email de confirmação:", error);
  }
}

function getOrcamentosCliente(email) {
  console.log("🔍 [DEBUG] getOrcamentosCliente chamado com email:", email);
  console.log("🔍 [DEBUG] Tipo do email:", typeof email);
  
  try {
    // VALIDAÇÃO ROBUSTA
    if (!email) {
      console.log("❌ [DEBUG] Email é null/undefined");
      return [];
    }
    
    var emailStr = email.toString();
    if (emailStr === "" || emailStr === "undefined" || emailStr === "null") {
      console.log("❌ [DEBUG] Email inválido:", emailStr);
      return [];
    }
    
    var emailLower = emailStr.trim().toLowerCase();
    console.log("🔍 [DEBUG] Buscando por email normalizado:", emailLower);
    
    var aba = getAbaOrcamentos();
    if (!aba) {
      console.log("❌ [DEBUG] Aba Orcamentos não encontrada");
      return [];
    }
    
    var dados = aba.getDataRange().getValues();
    console.log("📊 [DEBUG] Total de linhas na planilha:", dados.length);
    
    if (dados.length <= 1) {
      console.log("📊 [DEBUG] Nenhum dado na planilha");
      return [];
    }
    
    var orcamentos = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailOrcamento = linha[2] ? linha[2].toString().trim().toLowerCase() : "";
      
      console.log(`📌 [DEBUG] Linha ${i}: Email='${emailOrcamento}'`);
      
      if (emailOrcamento && emailOrcamento === emailLower) {
        console.log(`✅ [DEBUG] Match encontrado na linha ${i}`);
        
        // Formatar valor
        var valorRaw = linha[9] || "Aguardando";
        var valorFormatado = "Aguardando";
        
        if (typeof valorRaw === 'number') {
          valorFormatado = "R$ " + valorRaw.toFixed(2);
        } else if (valorRaw !== "Aguardando" && valorRaw !== "" && !isNaN(parseFloat(valorRaw))) {
          valorFormatado = "R$ " + parseFloat(valorRaw).toFixed(2);
        }
        
        var orcamento = {
          id: linha[0] || "",
          cliente: linha[1] || "",
          email: linha[2] || "",
          telefone: linha[3] || "",
          dataSolicitacao: linha[4] || "",
          tipoServico: linha[5] || "",
          descricao: linha[6] || "",
          imagemURL: linha[7] || "",
          arquivoURL: linha[8] || "",
          valor: valorFormatado,
          status: linha[10] || "Solicitado",
          dataValidade: linha[11] || "",
          observacoes: linha[12] || "",
          arquivoNome: linha[13] || ""
        };
        
        orcamentos.push(orcamento);
      }
    }
    
    console.log("✅ [DEBUG] Total de orçamentos encontrados:", orcamentos.length);
    return orcamentos;
    
  } catch (error) {
    console.error("❌ [DEBUG] Erro em getOrcamentosCliente:", error);
    console.error("❌ [DEBUG] Stack:", error.stack);
    return []; // SEMPRE retornar array vazio em caso de erro
  }
}

function testarGetOrcamentosDireto() {
  console.log("=== TESTE DIRETO ===");
  
  var emailTeste = "cawanfernandoanjok@hotmail.com";
  console.log("Testando com email:", emailTeste);
  
  var resultado = getOrcamentosCliente(emailTeste);
  console.log("Resultado:", resultado);
  console.log("Quantidade:", resultado ? resultado.length : "null");
  
  if (resultado && resultado.length > 0) {
    console.log("✅ SUCESSO! Primeiro orçamento:", resultado[0]);
  } else {
    console.log("❌ FALHA! Nenhum orçamento encontrado");
    
    // Verificar diretamente a planilha
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    console.log("Verificando planilha manualmente:");
    
    for (var i = 1; i < dados.length; i++) {
      console.log("Linha", i, "- Email:", dados[i][2]);
    }
  }
  
  return resultado;
}

// OBTER TODOS OS ORÇAMENTOS (para admin)
function getAllOrcamentos() {
  try {
    console.log("📊 Buscando todos os orçamentos...");
    
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var orcamentos = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      
      // Formatar valor
      var valor = linha[9] || "Aguardando";
      if (typeof valor === 'number') {
        valor = "R$ " + valor.toFixed(2);
      } else if (valor !== "Aguardando" && valor !== "") {
        valor = "R$ " + parseFloat(valor).toFixed(2);
      }
      
      orcamentos.push({
        id: linha[0] || "",
        cliente: linha[1] || "",
        email: linha[2] || "",
        telefone: linha[3] || "",
        dataSolicitacao: linha[4] || "",
        tipoServico: linha[5] || "",
        descricao: linha[6] || "",
        imagemURL: linha[7] || "",
        arquivoURL: linha[8] || "",
        valor: valor,
        status: linha[10] || "Solicitado",
        dataValidade: linha[11] || "",
        observacoes: linha[12] || "",
        arquivoNome: linha[13] || "",
        linha: i + 1
      });
    }
    
    // Ordenar por data mais recente primeiro
    orcamentos.sort(function(a, b) {
      var dateA = a.dataSolicitacao ? new Date(a.dataSolicitacao.split('/').reverse().join('-')) : new Date(0);
      var dateB = b.dataSolicitacao ? new Date(b.dataSolicitacao.split('/').reverse().join('-')) : new Date(0);
      return dateB - dateA;
    });
    
    console.log("✅ Total de orçamentos:", orcamentos.length);
    return orcamentos;
    
  } catch (error) {
    console.error("❌ Erro em getAllOrcamentos:", error);
    return [];
  }
}

// DiagnosticoCompleto.gs - EXECUTE ESTA FUNÇÃO ABAIXO
function DIAGNOSTICO_E_CORRECAO() {
  console.log("=== INICIANDO DIAGNÓSTICO COMPLETO ===");
  
  // 1. Verificar se o orçamento existe
  var aba = getAbaOrcamentos();
  var dados = aba.getDataRange().getValues();
  var orcamentoEncontrado = null;
  var linhaOrcamento = -1;
  
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][1] === "Cawan Oliveira") {
      orcamentoEncontrado = dados[i];
      linhaOrcamento = i + 1;
      break;
    }
  }
  
  if (!orcamentoEncontrado) {
    console.log("❌ Orçamento não encontrado");
    return;
  }
  
  console.log("✅ Orçamento encontrado na linha", linhaOrcamento);
  console.log("Status atual:", orcamentoEncontrado[10]);
  
  // 2. Garantir que o status está correto
  if (orcamentoEncontrado[10] !== "Solicitado") {
    aba.getRange(linhaOrcamento, 11).setValue("Solicitado");
    console.log("✅ Status corrigido para 'Solicitado'");
  }
  
  // 3. Garantir que o email está correto
  var emailCorreto = "cawanfernandoanjok@hotmail.com";
  if (orcamentoEncontrado[2] !== emailCorreto) {
    aba.getRange(linhaOrcamento, 3).setValue(emailCorreto);
    console.log("✅ Email corrigido");
  }
  
  // 4. Testar a função diretamente
  var resultado = getOrcamentosCliente(emailCorreto);
  console.log("Teste getOrcamentosCliente retornou:", resultado.length, "orçamento(s)");
  
  if (resultado.length > 0) {
    console.log("✅ SISTEMA ESTÁ FUNCIONANDO!");
    console.log("Dados do orçamento:", JSON.stringify(resultado[0]));
  } else {
    console.log("❌ AINDA COM PROBLEMA");
  }
  
  SpreadsheetApp.flush();
  
  // 5. GERAR URL CORRETA PARA ACESSO
  var url = ScriptApp.getService().getUrl();
  console.log("\n🔗 ACESSE ESTA URL PARA TESTAR:");
  console.log(url);
  console.log("\n👤 FAÇA LOGIN COM:");
  console.log("Email:", emailCorreto);
  console.log("Senha: admin123 (ou a senha que você cadastrou)");
  
  return {
    status: resultado.length > 0 ? "✅ FUNCIONANDO" : "❌ COM PROBLEMA",
    url: url,
    email: emailCorreto,
    orcamentoEncontrado: resultado.length > 0
  };
}

// Função para TESTAR DIRETO no navegador (adicione no final do Script.html)
// Cole o código abaixo no CONSOLE DO NAVEGADOR (F12) após fazer login
/*
(async function testeHistorico() {
  console.log("=== TESTE RÁPIDO ===");
  var email = "cawanfernandoanjok@hotmail.com";
  var result = await google.script.run.withSuccessHandler(r => r).withFailureHandler(e => e).getOrcamentosCliente(email);
  console.log("Resultado:", result);
  if (result && result.length > 0) {
    alert("✅ Orçamento encontrado! " + result.length + " orçamento(s)");
  } else {
    alert("❌ Nenhum orçamento encontrado para: " + email);
  }
})();
*/

function verOrcamentoManual() {
  var aba = getAbaOrcamentos();
  var dados = aba.getDataRange().getValues();
  
  console.log("=== VERIFICANDO ORÇAMENTOS MANUALMENTE ===");
  console.log("Cabeçalhos:", dados[0]);
  
  for (var i = 1; i < dados.length; i++) {
    console.log("\n--- Linha", i, "---");
    console.log("ID:", dados[i][0]);
    console.log("Cliente:", dados[i][1]);
    console.log("Email:", "'" + dados[i][2] + "'");
    console.log("Status:", dados[i][10]);
    console.log("Valor:", dados[i][9]);
  }
}

// OBTER ORÇAMENTOS PENDENTES
function getOrcamentosPendentes() {
  try {
    console.log("⏳ Buscando orçamentos pendentes...");
    
    var todos = getAllOrcamentos();
    var pendentes = todos.filter(function(orc) {
      return orc.status === "Solicitado" || orc.status === "Aguardando" || orc.status === "Em Negociação";
    });
    
    console.log("✅ Orçamentos pendentes:", pendentes.length);
    return pendentes;
    
  } catch (error) {
    console.error("❌ Erro em getOrcamentosPendentes:", error);
    return [];
  }
}

// APROVAR/RECUSAR ORÇAMENTO COM CONTRA-PROPOSTA
function aprovarOrcamento(idOrcamento, valor, status, observacoesAdmin) {
  try {
    console.log("=== APROVANDO ORÇAMENTO ===");
    console.log("ID:", idOrcamento, "Valor:", valor, "Status:", status);
    
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var linhaEncontrada = -1;
    
    // Buscar orçamento pelo ID
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === idOrcamento) {
        linhaEncontrada = i + 1;
        break;
      }
    }
    
    if (linhaEncontrada > 0) {
      // Atualizar valor e status
      if (valor && valor > 0) {
        aba.getRange(linhaEncontrada, 10).setValue(parseFloat(valor)); // Coluna J: Valor
      }
      aba.getRange(linhaEncontrada, 11).setValue(status); // Coluna K: Status
      
      // Adicionar observações do admin
      var observacoesAtuais = aba.getRange(linhaEncontrada, 13).getValue(); // Coluna M: Observacoes
      var timestamp = formatarDataHoraSistema_();
      
      var novasObservacoes = observacoesAtuais ? 
          observacoesAtuais + "\n\n--- ADMIN (" + timestamp + ") ---\n" +
          "Status: " + status + "\n" +
          (valor ? "Valor: R$ " + parseFloat(valor).toFixed(2) + "\n" : "") +
          (observacoesAdmin ? "Observações: " + observacoesAdmin : "") :
          "--- ADMIN (" + timestamp + ") ---\n" +
          "Status: " + status + "\n" +
          (valor ? "Valor: R$ " + parseFloat(valor).toFixed(2) + "\n" : "") +
          (observacoesAdmin ? "Observações: " + observacoesAdmin : "");
      
      aba.getRange(linhaEncontrada, 13).setValue(novasObservacoes);
      SpreadsheetApp.flush();
      
      // Obter dados do cliente para email
      var emailCliente = aba.getRange(linhaEncontrada, 3).getValue();
      var cliente = aba.getRange(linhaEncontrada, 2).getValue();
      var tipoServico = aba.getRange(linhaEncontrada, 6).getValue();
      
      // Enviar email apropriado
      if (status === "Em Negociação" || status === "Contraproposta") {
        enviarEmailContraproposta(emailCliente, cliente, idOrcamento, tipoServico, valor, observacoesAdmin);
      } else if (status === "Aprovado") {
        enviarEmailOrcamentoAprovado(emailCliente, cliente, idOrcamento, tipoServico, valor, observacoesAdmin);
        
        // Criar pagamento se necessário
        try {
          criarPagamentoParaOrcamento(idOrcamento, {
            cliente: cliente,
            email: emailCliente,
            tipoServico: tipoServico,
            valor: valor
          });
        } catch (pagError) {
          console.log("⚠️ Erro ao criar pagamento:", pagError);
        }
        
      } else if (status === "Recusado") {
        enviarEmailOrcamentoRecusado(emailCliente, cliente, idOrcamento, tipoServico, observacoesAdmin);
      }
      
      // Registrar log
      try {
        logEvent("ORCAMENTO", emailCliente, "APROVACAO", { 
          id: idOrcamento, 
          status: status, 
          valor: valor 
        });
      } catch (logError) {
        console.log("⚠️ Erro ao registrar log:", logError);
      }
      
      return {
        success: true,
        message: "✅ Orçamento atualizado com sucesso! Cliente notificado."
      };
    } else {
      return {
        success: false,
        message: "Orçamento não encontrado."
      };
    }
    
  } catch (error) {
    console.error("❌ Erro em aprovarOrcamento:", error);
    return {
      success: false,
      message: "Erro ao aprovar orçamento: " + error.toString()
    };
  }
}

// ENVIAR EMAIL DE CONTRA-PROPOSTA
function enviarEmailContraproposta(email, cliente, idOrcamento, tipoServico, valor, observacoes) {
  try {
    var assunto = "Contraproposta para seu Orçamento - JW Piscinas";
    var valorText = valor ? "R$ " + parseFloat(valor).toFixed(2) : "A definir";
    
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <h3>Olá, ${cliente}!</h3>
        <p>Recebemos seu orçamento e temos uma contraproposta para você.</p>
        
        <p><strong>Detalhes da Contraproposta:</strong></p>
        <ul>
          <li><strong>ID do Orçamento:</strong> ${idOrcamento}</li>
          <li><strong>Tipo de Serviço:</strong> ${tipoServico}</li>
          <li><strong>Valor Proposto:</strong> ${valorText}</li>
          <li><strong>Data da Proposta:</strong> ${new Date().toLocaleString('pt-BR')}</li>
        </ul>
        
        ${observacoes ? `<p><strong>Observações da nossa equipe:</strong><br>${observacoes.replace(/\n/g, '<br>')}</p>` : ''}
        
        <p><strong>Próximos passos:</strong></p>
        <ol>
          <li>Acesse o sistema JW Piscinas</li>
          <li>Vá até a aba "Histórico"</li>
          <li>Veja os detalhes da contraproposta</li>
          <li>Entre em contato para aceitar ou negociar</li>
        </ol>
        
        <p>Ficamos à disposição para esclarecer qualquer dúvida.</p>
        
        <p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>
        <hr>
        <p style="font-size: 12px; color: #666;">
          ID do Orçamento: ${idOrcamento}
        </p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      htmlBody: corpo
    });
    
    console.log("✅ Email de contraproposta enviado para:", email);
    
  } catch (error) {
    console.error("❌ Erro ao enviar email de contraproposta:", error);
  }
}

// ENVIAR EMAIL DE ORÇAMENTO APROVADO
function enviarEmailOrcamentoAprovado(email, cliente, idOrcamento, tipoServico, valor, observacoes) {
  try {
    var assunto = "Orçamento Aprovado - JW Piscinas";
    var valorText = valor ? "R$ " + parseFloat(valor).toFixed(2) : "A definir";
    
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <h3>Olá, ${cliente}!</h3>
        <p>Seu orçamento foi aprovado pela nossa equipe!</p>
        
        <p><strong>Detalhes do Orçamento:</strong></p>
        <ul>
          <li><strong>ID do Orçamento:</strong> ${idOrcamento}</li>
          <li><strong>Tipo de Serviço:</strong> ${tipoServico}</li>
          <li><strong>Valor Aprovado:</strong> ${valorText}</li>
        </ul>
        
        ${observacoes ? `<p><strong>Observações:</strong><br>${observacoes.replace(/\n/g, '<br>')}</p>` : ''}
        
        <p><strong>Próximos passos:</strong></p>
        <ol>
          <li>Em breve entraremos em contato para agendar o serviço</li>
          <li>Você receberá informações sobre pagamento</li>
          <li>Acompanhe o status no aplicativo JW Piscinas</li>
        </ol>
        
        <p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>
        <hr>
        <p style="font-size: 12px; color: #666;">
          ID do Orçamento: ${idOrcamento}
        </p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      htmlBody: corpo
    });
    
    console.log("✅ Email de aprovação enviado para:", email);
    
  } catch (error) {
    console.error("❌ Erro ao enviar email de aprovação:", error);
  }
}

// ENVIAR EMAIL DE ORÇAMENTO RECUSADO
function enviarEmailOrcamentoRecusado(email, cliente, idOrcamento, tipoServico, observacoes) {
  try {
    var assunto = "Atualização do seu Orçamento - JW Piscinas";
    
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <h3>Olá, ${cliente}!</h3>
        <p>Informamos que seu orçamento foi revisado pela nossa equipe.</p>
        
        <p><strong>Detalhes do Orçamento:</strong></p>
        <ul>
          <li><strong>ID do Orçamento:</strong> ${idOrcamento}</li>
          <li><strong>Tipo de Serviço:</strong> ${tipoServico}</li>
          <li><strong>Status:</strong> Não aprovado</li>
        </ul>
        
        ${observacoes ? `<p><strong>Observações da nossa equipe:</strong><br>${observacoes.replace(/\n/g, '<br>')}</p>` : ''}
        
        <p><strong>O que fazer agora?</strong></p>
        <ul>
          <li>Entre em contato conosco para esclarecer dúvidas</li>
          <li>Solicite um novo orçamento com informações adicionais</li>
          <li>Converse conosco para entender melhor os requisitos</li>
        </ul>
        
        <p>Estamos à disposição para ajudar!</p>
        
        <p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>
        <hr>
        <p style="font-size: 12px; color: #666;">
          ID do Orçamento: ${idOrcamento}
        </p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      htmlBody: corpo
    });
    
    console.log("✅ Email de recusa enviado para:", email);
    
  } catch (error) {
    console.error("❌ Erro ao enviar email de recusa:", error);
  }
}

// CRIAR PAGAMENTO PARA ORÇAMENTO APROVADO
function criarPagamentoParaOrcamento(idOrcamento, dados) {
  try {
    var abaPagamentos = getAbaPagamentos();
    
    if (!abaPagamentos) {
      console.log("📁 Aba Pagamentos não encontrada, criando...");
      abaPagamentos = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Pagamentos");
      abaPagamentos.getRange("A1:J1").setValues([[
        "ID", "Cliente", "Email", "Tipo", "Valor", "Vencimento", "Status", "DataPagamento", "Metodo", "Observacoes"
      ]]);
    }
    
    var idPagamento = "PAG-" + new Date().getTime().toString().slice(-6);
    var vencimento = new Date();
    vencimento.setDate(vencimento.getDate() + 7); // 7 dias para pagamento
    
    var novaLinha = [
      idPagamento,
      dados.cliente,
      dados.email,
      "Orçamento - " + dados.tipoServico,
      dados.valor,
      formatarDataSistema_(vencimento),
      "Pendente",
      "",
      "",
      "Referente ao orçamento " + idOrcamento
    ];
    
    abaPagamentos.appendRow(novaLinha);
    SpreadsheetApp.flush();
    
    console.log("✅ Pagamento criado:", idPagamento);
    return idPagamento;
    
  } catch (error) {
    console.error("❌ Erro ao criar pagamento:", error);
    return null;
  }
}

// CANCELAR ORÇAMENTO
function cancelarOrcamento(idOrcamento, motivo) {
  try {
    console.log("🚫 Cancelando orçamento:", idOrcamento);
    
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var linhaEncontrada = -1;
    
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === idOrcamento) {
        linhaEncontrada = i + 1;
        break;
      }
    }
    
    if (linhaEncontrada > 0) {
      // Atualizar status
      aba.getRange(linhaEncontrada, 11).setValue("Cancelado"); // Coluna K: Status
      
      // Adicionar observações
      var observacoesAtuais = aba.getRange(linhaEncontrada, 13).getValue();
      var timestamp = formatarDataHoraSistema_();
      
      var novasObservacoes = observacoesAtuais ? 
          observacoesAtuais + "\n\n--- CANCELADO (" + timestamp + ") ---\nMotivo: " + motivo :
          "--- CANCELADO (" + timestamp + ") ---\nMotivo: " + motivo;
      
      aba.getRange(linhaEncontrada, 13).setValue(novasObservacoes);
      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: "✅ Orçamento cancelado com sucesso!"
      };
    } else {
      return {
        success: false,
        message: "Orçamento não encontrado."
      };
    }
    
  } catch (error) {
    console.error("❌ Erro em cancelarOrcamento:", error);
    return {
      success: false,
      message: "Erro ao cancelar orçamento: " + error.toString()
    };
  }
}

// ATUALIZAR ORÇAMENTO
function atualizarOrcamento(idOrcamento, dadosAtualizados) {
  try {
    console.log("✏️ Atualizando orçamento:", idOrcamento);
    
    var aba = getAbaOrcamentos();
    var dados = aba.getDataRange().getValues();
    var linhaEncontrada = -1;
    
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === idOrcamento) {
        linhaEncontrada = i + 1;
        break;
      }
    }
    
    if (linhaEncontrada > 0) {
      // Atualizar campos permitidos
      if (dadosAtualizados.telefone) {
        aba.getRange(linhaEncontrada, 4).setValue(dadosAtualizados.telefone);
      }
      
      if (dadosAtualizados.descricao) {
        aba.getRange(linhaEncontrada, 7).setValue(dadosAtualizados.descricao);
      }
      
      if (dadosAtualizados.observacoes) {
        var observacoesAtuais = aba.getRange(linhaEncontrada, 13).getValue();
        var timestamp = formatarDataHoraSistema_();
        
        var novasObservacoes = observacoesAtuais ? 
            observacoesAtuais + "\n\n--- ATUALIZADO (" + timestamp + ") ---\n" + dadosAtualizados.observacoes :
            "--- ATUALIZADO (" + timestamp + ") ---\n" + dadosAtualizados.observacoes;
        
        aba.getRange(linhaEncontrada, 13).setValue(novasObservacoes);
      }
      
      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: "✅ Orçamento atualizado com sucesso!"
      };
    } else {
      return {
        success: false,
        message: "Orçamento não encontrado."
      };
    }
    
  } catch (error) {
    console.error("❌ Erro em atualizarOrcamento:", error);
    return {
      success: false,
      message: "Erro ao atualizar orçamento: " + error.toString()
    };
  }
}

// EXPORTAR ORÇAMENTOS PARA CSV
function exportarOrcamentosCSV() {
  try {
    var orcamentos = getAllOrcamentos();
    
    if (!orcamentos || orcamentos.length === 0) {
      return "Nenhum orçamento encontrado.";
    }
    
    var csv = "ID;Cliente;Email;Telefone;Data;Tipo;Status;Valor;Observações\n";
    
    orcamentos.forEach(function(orc) {
      var linha = [
        orc.id,
        orc.cliente,
        orc.email,
        orc.telefone,
        orc.dataSolicitacao,
        orc.tipoServico,
        orc.status,
        orc.valor,
        (orc.observacoes || "").replace(/;/g, ',').replace(/\n/g, ' ')
      ].join(';');
      
      csv += linha + "\n";
    });
    
    return csv;
    
  } catch (error) {
    console.error("❌ Erro em exportarOrcamentosCSV:", error);
    return "Erro ao exportar: " + error.toString();
  }
}

// ESTATÍSTICAS DE ORÇAMENTOS
function getEstatisticasOrcamentos() {
  try {
    var orcamentos = getAllOrcamentos();
    
    var stats = {
      total: orcamentos.length,
      solicitados: 0,
      aprovados: 0,
      recusados: 0,
      cancelados: 0,
      negociacao: 0,
      valorTotalAprovado: 0,
      mediaValor: 0
    };
    
    orcamentos.forEach(function(orc) {
      switch(orc.status) {
        case "Solicitado":
        case "Aguardando":
          stats.solicitados++;
          break;
        case "Aprovado":
          stats.aprovados++;
          var valorNum = parseFloat(orc.valor.replace('R$ ', '').replace(',', '.'));
          if (!isNaN(valorNum)) {
            stats.valorTotalAprovado += valorNum;
          }
          break;
        case "Recusado":
          stats.recusados++;
          break;
        case "Cancelado":
          stats.cancelados++;
          break;
        case "Em Negociação":
        case "Contraproposta":
          stats.negociacao++;
          break;
      }
    });
    
    if (stats.aprovados > 0) {
      stats.mediaValor = stats.valorTotalAprovado / stats.aprovados;
    }
    
    return stats;
    
  } catch (error) {
    console.error("❌ Erro em getEstatisticasOrcamentos:", error);
    return null;
  }
}
