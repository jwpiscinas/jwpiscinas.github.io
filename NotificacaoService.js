// NotificacaoService.gs - Sistema de Notificações

// Obter aba de notificações
function getAbaNotificacoes() {
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("Notificacoes");
    
    if (!aba) {
      aba = planilha.insertSheet("Notificacoes");
      
      var cabecalhos = [[
        "ID", "Titulo", "Mensagem", "Tipo", "DataEnvio", 
        "Destinatarios", "Lida", "CriadoPor", "Acao", "Icone"
      ]];
      
      aba.getRange("A1:J1").setValues(cabecalhos);
      aba.setFrozenRows(1);
    }
    
    return aba;
    
  } catch (error) {
    console.error("Erro em getAbaNotificacoes:", error);
    throw error;
  }
}

// Obter notificações do usuário
function getNotificacoesUsuario(email) {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var notificacoes = [];
    var emailLower = email ? email.toString().trim().toLowerCase() : "";
    var tipoUsuario = getTipoUsuarioPorEmail_(emailLower);
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var destinatarios = linha[5] ? linha[5].toString().toLowerCase() : "";
      var lidas = linha[6] ? linha[6].toString().toLowerCase() : "";
      
      // Verificar se é para este usuário
      var deveReceber = false;
      if (destinatarios === "todos" ||
          (destinatarios === "clientes" && tipoUsuario !== "Administrador") ||
          (destinatarios === "admin" && tipoUsuario === "Administrador") ||
          destinatarios.includes(emailLower)) {
        deveReceber = true;
      }
      
      if (!deveReceber) continue;
      
      notificacoes.push({
        id: linha[0] || "",
        titulo: linha[1] || "",
        mensagem: linha[2] || "",
        tipo: linha[3] || "info",
        dataEnvio: linha[4] || "",
        acao: linha[8] || "",
        icone: linha[9] || "fas fa-bell",
        lida: lidas.includes(emailLower)
      });
    }
    
    // Ordenar: não lidas primeiro
    notificacoes.sort(function(a, b) {
      if (a.lida && !b.lida) return 1;
      if (!a.lida && b.lida) return -1;
      return 0;
    });
    
    return notificacoes;
    
  } catch (error) {
    console.error("Erro em getNotificacoesUsuario:", error);
    return [];
  }
}

function getTipoUsuarioPorEmail_(email) {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      var emailUsuario = dados[i][0] ? dados[i][0].toString().trim().toLowerCase() : "";

      if (emailUsuario === email) {
        return isUsuarioAdmin_(dados[i]) ? "Administrador" : "Cliente";
      }
    }
  } catch (error) {
    console.error("Erro ao identificar tipo do usuario:", error);
  }

  return "Cliente";
}

// Marcar notificação como lida
function marcarNotificacaoLida(idNotificacao, email) {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var emailLower = email ? email.toString().trim().toLowerCase() : "";
    
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === idNotificacao) {
        var lidas = dados[i][6] ? dados[i][6].toString() : "";
        var lidasLower = lidas.toLowerCase();
        
        if (!lidasLower.includes(emailLower)) {
          var novasLidas = lidas ? lidas + "," + emailLower : emailLower;
          aba.getRange(i + 1, 7).setValue(novasLidas);
          SpreadsheetApp.flush();
        }
        
        return { success: true };
      }
    }
    
    return { success: false };
    
  } catch (error) {
    console.error("Erro em marcarNotificacaoLida:", error);
    return { success: false };
  }
}

// Marcar todas como lidas
function marcarTodasNotificacoesLidas(email) {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var emailLower = email ? email.toString().trim().toLowerCase() : "";
    var tipoUsuario = getTipoUsuarioPorEmail_(emailLower);
    
    for (var i = 1; i < dados.length; i++) {
      var destinatarios = dados[i][5] ? dados[i][5].toString().toLowerCase() : "";
      var lidas = dados[i][6] ? dados[i][6].toString() : "";
      var lidasLower = lidas.toLowerCase();
      
      var deveReceber = (destinatarios === "todos" ||
                        (destinatarios === "clientes" && tipoUsuario !== "Administrador") ||
                        (destinatarios === "admin" && tipoUsuario === "Administrador") ||
                        destinatarios.includes(emailLower));
      
      if (deveReceber && !lidasLower.includes(emailLower)) {
        var novasLidas = lidas ? lidas + "," + emailLower : emailLower;
        aba.getRange(i + 1, 7).setValue(novasLidas);
      }
    }
    
    SpreadsheetApp.flush();
    return { success: true };
    
  } catch (error) {
    console.error("Erro em marcarTodasNotificacoesLidas:", error);
    return { success: false };
  }
}

// Enviar notificação (admin)
function enviarNotificacao(dados) {
  try {
    var aba = getAbaNotificacoes();
    
    var id = "NOT-" + new Date().getTime().toString().slice(-8);
    var dataEnvio = typeof formatarDataHoraSistema_ === "function"
      ? formatarDataHoraSistema_()
      : Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
    
    var novaLinha = [
      id,
      dados.titulo,
      dados.mensagem,
      dados.tipo || "info",
      dataEnvio,
      dados.destinatarios || "todos",
      "", // Lidas
      Session.getActiveUser().getEmail(),
      dados.acao || "",
      dados.icone || "fas fa-bell"
    ];
    
    aba.appendRow(novaLinha);
    SpreadsheetApp.flush();
    
    // Enviar email se necessário
    if (dados.enviarEmail === true) {
      enviarNotificacaoPorEmail(dados);
    }
    
    return {
      success: true,
      id: id,
      message: "✅ Notificação enviada com sucesso!"
    };
    
  } catch (error) {
    console.error("Erro em enviarNotificacao:", error);
    return {
      success: false,
      message: "Erro ao enviar notificação: " + error.toString()
    };
  }
}

// Enviar notificação por email (COM CONFIGURAÇÕES ANTI-SPAM)
function enviarNotificacaoPorEmail(dados) {
  try {
    var destinatarios = [];
    
    if (dados.destinatarios === "todos") {
      var usuarios = getTodosUsuarios();
      destinatarios = usuarios.map(u => u.email);
    } else if (dados.destinatarios === "clientes") {
      var clientes = listarClientesAtivos();
      destinatarios = clientes.clientes.map(c => c.email);
    } else if (dados.destinatarios === "admin") {
      var admins = getTodosUsuarios().filter(function(u) { return u.tipo === "Administrador"; });
      destinatarios = admins.map(function(u) { return u.email; });
    }
    
    if (destinatarios.length === 0) return;
    
    // Configurações para evitar spam
    var assunto = "🔔 JW Piscinas - " + dados.titulo;
    
    // Corpo HTML com domínio próprio e links válidos
    var corpo = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="font-family: Arial, sans-serif; margin: 0; padding: 0; background: #f4f4f4;">
        <div style="max-width: 600px; margin: 20px auto; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
          
          <!-- Header -->
          <div style="background: linear-gradient(135deg, #0077B6, #00B4D8); padding: 30px; text-align: center;">
            <h1 style="color: white; margin: 0; font-size: 28px;">JW Piscinas</h1>
          </div>
          
          <!-- Conteúdo -->
          <div style="padding: 30px;">
            <h2 style="color: #0077B6; margin-top: 0;">${dados.titulo}</h2>
            <p style="color: #333; line-height: 1.6; font-size: 16px;">${dados.mensagem.replace(/\n/g, '<br>')}</p>
            
            ${dados.acao ? `
              <div style="text-align: center; margin: 30px 0;">
                <a href="${dados.acao}" style="background: #0077B6; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; display: inline-block; font-weight: bold;">Ver Detalhes</a>
              </div>
            ` : ''}
            
            <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 30px 0;">
            
            <p style="color: #666; font-size: 14px;">Esta é uma comunicação oficial do sistema JW Piscinas. Você recebeu este e-mail porque é um cliente cadastrado.</p>
            <p style="color: #666; font-size: 12px;">Para não receber mais e-mails, acesse suas configurações de notificação no sistema.</p>
          </div>
          
          <!-- Footer -->
          <div style="background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #e0e0e0;">
            <p style="color: #999; font-size: 12px; margin: 0;">© ${new Date().getFullYear()} JW Piscinas. Todos os direitos reservados.</p>
            <p style="color: #999; font-size: 12px; margin: 5px 0 0;">contato@jwpiscinas.com.br | (16) 99629-4795</p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    // Configurações avançadas para melhor deliverabilidade
    var opcoes = {
      name: 'JW Piscinas',
      replyTo: 'naoresponder@jwpiscinas.com.br'
    };
    
    // Enviar em lotes com intervalo
    for (var i = 0; i < destinatarios.length; i += 50) {
      var lote = destinatarios.slice(i, i + 50);
      
      MailApp.sendEmail({
        to: lote.join(','),
        subject: assunto,
        htmlBody: corpo,
        name: 'JW Piscinas',
        replyTo: 'naoresponder@jwpiscinas.com.br'
      });
      
      // Pequena pausa entre lotes para não parecer spam
      if (i + 50 < destinatarios.length) {
        Utilities.sleep(1000);
      }
    }
    
    console.log("✅ Emails enviados para", destinatarios.length, "destinatários");
    
  } catch (error) {
    console.error("Erro ao enviar emails:", error);
  }
}

// Obter contagem de não lidas
function getContagemNaoLidas(email) {
  try {
    var notificacoes = getNotificacoesUsuario(email);
    return notificacoes.filter(n => !n.lida).length;
  } catch (error) {
    return 0;
  }
}

// Obter todas as notificações (admin)
function getAllNotificacoes() {
  try {
    var aba = getAbaNotificacoes();
    var dados = aba.getDataRange().getValues();
    var notificacoes = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      notificacoes.push({
        id: linha[0],
        titulo: linha[1],
        mensagem: linha[2],
        tipo: linha[3],
        dataEnvio: linha[4],
        destinatarios: linha[5],
        criadoPor: linha[7],
        acao: linha[8],
        icone: linha[9]
      });
    }
    
    return notificacoes;
    
  } catch (error) {
    console.error("Erro em getAllNotificacoes:", error);
    return [];
  }
}
