// === AUTH SERVICE - COM CONTROLE DE TELEFONES AUTORIZADOS ===

// 1. ABA DE USUÁRIOS
function getAbaUsuarios() {
  try {
    if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
      return getFirestoreSheetAdapter_("Usuarios");
    }

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("Usuarios");
    
    if (!aba) {
      console.log("📁 Criando aba Usuarios...");
      aba = planilha.insertSheet("Usuarios");
      
      var cabecalhos = [
        ["Email", "Senha", "Usuario", "Telefone", "Tipo", "DataCadastro", "Status", "CodigoRecuperacao", "DataCodigo", "AvatarURL", "Admin", "Endereco", "Bairro"]
      ];
      
      aba.getRange("A1:M1").setValues(cabecalhos);
      aba.setFrozenRows(1);
      aplicarCheckboxAdmin_(aba);
    } else {
      garantirCabecalhosUsuarios_(aba);
    }
    
    return aba;
    
  } catch (error) {
    console.error("❌ Erro em getAbaUsuarios:", error);
    throw error;
  }
}

function garantirCabecalhosUsuarios_(aba) {
  var cabecalhos = ["Email", "Senha", "Usuario", "Telefone", "Tipo", "DataCadastro", "Status", "CodigoRecuperacao", "DataCodigo", "AvatarURL", "Admin", "Endereco", "Bairro"];
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

  aplicarCheckboxAdmin_(aba);
}

function aplicarCheckboxAdmin_(aba) {
  try {
    var regra = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    aba.getRange(2, 11, Math.max(aba.getMaxRows() - 1, 1), 1).setDataValidation(regra);
  } catch (error) {
    console.error("Erro ao aplicar checkbox de admin:", error);
  }
}

function isUsuarioAdmin_(linha) {
  var tipoUsuario = linha[4] ? linha[4].toString().trim().toLowerCase() : "";
  var adminCheckbox = linha[10];

  if (tipoUsuario === "administrador" || tipoUsuario === "admin") return true;
  if (adminCheckbox === true) return true;
  if (adminCheckbox && adminCheckbox.toString().trim().toLowerCase() === "true") return true;
  if (adminCheckbox && adminCheckbox.toString().trim().toLowerCase() === "sim") return true;

  return false;
}

// 2. ABA DE TELEFONES AUTORIZADOS
function getAbaTelefonesAutorizados() {
  try {
    if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
      return getFirestoreSheetAdapter_("TelefonesAutorizados");
    }

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName("TelefonesAutorizados");
    
    if (!aba) {
      aba = planilha.insertSheet("TelefonesAutorizados");
      
      var cabecalhos = [
        ["Telefone", "Nome", "DataAutorizacao", "Status", "Usado", "DataUso"]
      ];
      
      aba.getRange("A1:F1").setValues(cabecalhos);
      aba.setFrozenRows(1);
      
      var exemplos = [
        ["(16) 99629-4795", "Cawan Oliveira", new Date().toLocaleDateString('pt-BR'), "Ativo", "Não", ""],
        ["(16) 99123-4567", "Cliente Exemplo", new Date().toLocaleDateString('pt-BR'), "Ativo", "Não", ""]
      ];
      
      for (var i = 0; i < exemplos.length; i++) {
        aba.appendRow(exemplos[i]);
      }
    }
    
    return aba;
    
  } catch (error) {
    console.error("❌ Erro em getAbaTelefonesAutorizados:", error);
    throw error;
  }
}

// 3. FUNÇÃO DE VALIDAÇÃO DE EMAIL
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

// 4. FUNÇÃO DE VALIDAÇÃO DE TELEFONE
function isValidTelefone(telefone) {
  if (!telefone || typeof telefone !== 'string') return false;
  var numeros = telefone.replace(/\D/g, '');
  return numeros.length >= 10 && numeros.length <= 11;
}

// 5. FUNÇÃO PARA NORMALIZAR TELEFONE
function normalizarTelefone(telefone) {
  if (!telefone) return "";
  var numeros = telefone.replace(/\D/g, '');
  
  if (numeros.length === 11) {
    return numeros.replace(/^(\d{2})(\d{5})(\d{4})$/, '($1) $2-$3');
  } else if (numeros.length === 10) {
    return numeros.replace(/^(\d{2})(\d{4})(\d{4})$/, '($1) $2-$3');
  }
  
  return telefone;
}

// 6. VERIFICAR SE TELEFONE ESTÁ AUTORIZADO
function verificarTelefoneAutorizado(telefone) {
  try {
    var aba = getAbaTelefonesAutorizados();
    var dados = aba.getDataRange().getValues();
    var telefoneBuscado = normalizarTelefone(telefone);
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var telefoneAutorizado = linha[0] ? normalizarTelefone(linha[0].toString()) : "";
      var status = linha[3] ? linha[3].toString() : "";
      var usado = linha[4] ? linha[4].toString() : "Não";
      
      if (telefoneAutorizado === telefoneBuscado) {
        if (status !== "Ativo") {
          return { autorizado: false, motivo: "Telefone não está ativo para cadastro." };
        }
        
        if (usado === "Sim") {
          return { autorizado: false, motivo: "Este telefone já foi cadastrado anteriormente." };
        }
        
        return { 
          autorizado: true, 
          motivo: "Telefone autorizado para cadastro.",
          linha: i + 1,
          telefone: telefoneAutorizado
        };
      }
    }
    
    return { autorizado: false, motivo: "Número de telefone não autorizado para cadastro." };
    
  } catch (error) {
    return { autorizado: false, motivo: "Erro ao verificar telefone: " + error.toString() };
  }
}

// 7. MARCAR TELEFONE COMO UTILIZADO
function marcarTelefoneComoUtilizado(telefone, emailUsuario) {
  try {
    var aba = getAbaTelefonesAutorizados();
    var dados = aba.getDataRange().getValues();
    var telefoneBuscado = normalizarTelefone(telefone);
    
    for (var i = 1; i < dados.length; i++) {
      var telefoneAutorizado = dados[i][0] ? normalizarTelefone(dados[i][0].toString()) : "";
      
      if (telefoneAutorizado === telefoneBuscado) {
        aba.getRange(i + 1, 5).setValue("Sim");
        aba.getRange(i + 1, 6).setValue(formatarDataHoraSistema_());
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    return false;
  }
}

// 8. FUNÇÃO PRINCIPAL DE LOGIN
function verificarLogin(dados) {
  try {
    if (!dados || !dados.usuario || !dados.senha) {
      return { success: false, message: "Preencha usuario e senha!" };
    }
    
    var usuarioDigitado = dados.usuario.toString().trim();
    var senhaDigitada = dados.senha.toString();
    
    var aba = getAbaUsuarios();
    var dadosPlanilha = aba.getDataRange().getValues();
    
    if (dadosPlanilha.length <= 1) {
      return { success: false, message: "Nenhum usuário cadastrado no sistema!" };
    }
    
    for (var i = 1; i < dadosPlanilha.length; i++) {
      var linha = dadosPlanilha[i];
      var emailUsuario = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      var senhaUsuario = linha[1] ? linha[1].toString() : "";
      var nomeUsuario = linha[2] ? linha[2].toString().trim().toLowerCase() : "";
      var telefoneUsuario = linha[3] ? normalizarTelefone(linha[3].toString()).toLowerCase() : "";
      var tipoUsuario = isUsuarioAdmin_(linha) ? "Administrador" : (linha[4] ? linha[4].toString() : "Cliente");
      var statusUsuario = linha[6] ? linha[6].toString() : "";
      
      var usuarioDigitadoLower = usuarioDigitado.toLowerCase();
      var usuarioDigitadoNormalizado = normalizarTelefone(usuarioDigitado).toLowerCase();
      
      var usuarioEncontrado = false;
      
      if (nomeUsuario === usuarioDigitadoLower || 
          emailUsuario === usuarioDigitadoLower || 
          telefoneUsuario === usuarioDigitadoLower ||
          telefoneUsuario === usuarioDigitadoNormalizado) {
        usuarioEncontrado = true;
      }
      
      if (usuarioEncontrado) {
        if (statusUsuario !== "Ativo") {
          return { success: false, message: "Usuário inativo!" };
        }
        
        if (senhaUsuario === senhaDigitada) {
          return {
            success: true,
            message: "Login realizado com sucesso!",
            usuario: linha[2].toString().trim(),
            email: emailUsuario,
            telefone: linha[3] ? linha[3].toString() : "",
            tipo: tipoUsuario,
            endereco: linha[11] ? linha[11].toString() : "",
            bairro: linha[12] ? linha[12].toString() : ""
          };
        } else {
          return { success: false, message: "Senha incorreta!" };
        }
      }
    }
    
    return { success: false, message: "Usuário não cadastrado!" };
    
  } catch (error) {
    return { success: false, message: "Erro no servidor: " + error.toString() };
  }
}

// 9. FUNÇÃO PRINCIPAL DE CADASTRO
function cadastrarUsuario(dados) {
  try {
    if (!dados || !dados.email || !dados.senha || !dados.confirmacaoSenha || !dados.usuario || !dados.telefone || !dados.endereco || !dados.bairro) {
      return { success: false, message: "Preencha todos os campos, incluindo bairro e endereco!" };
    }
    
    var email = dados.email.toString().trim().toLowerCase();
    var senha = dados.senha.toString();
    var confirmacaoSenha = dados.confirmacaoSenha.toString();
    var nomeUsuario = dados.usuario.toString().trim();
    var telefone = normalizarTelefone(dados.telefone.toString().trim());
    var endereco = dados.endereco.toString().trim();
    var bairro = dados.bairro.toString().trim();
    
    if (!isValidEmail(email)) {
      return { success: false, message: "Email inválido!" };
    }
    
    if (nomeUsuario.length < 3) {
      return { success: false, message: "Nome de usuário deve ter pelo menos 3 caracteres!" };
    }
    
    if (senha.length < 4) {
      return { success: false, message: "Senha deve ter pelo menos 4 caracteres!" };
    }
    
    if (senha.startsWith('0')) {
      return { success: false, message: "A senha não pode começar com zero!" };
    }
    
    if (senha !== confirmacaoSenha) {
      return { success: false, message: "As senhas não coincidem!" };
    }
    
    if (!isValidTelefone(telefone)) {
      return { success: false, message: "Telefone inválido! Use DDD + número (10 ou 11 dígitos)" };
    }
    
    var abaUsuarios = getAbaUsuarios();
    var dadosUsuarios = abaUsuarios.getDataRange().getValues();
    
    for (var i = 1; i < dadosUsuarios.length; i++) {
      var usuarioExistente = dadosUsuarios[i][2] ? dadosUsuarios[i][2].toString().trim().toLowerCase() : "";
      if (usuarioExistente === nomeUsuario.toLowerCase()) {
        return { success: false, message: "Este nome de usuário já está em uso!" };
      }
    }
    
    for (var i = 1; i < dadosUsuarios.length; i++) {
      var emailExistente = dadosUsuarios[i][0] ? dadosUsuarios[i][0].toString().trim().toLowerCase() : "";
      if (emailExistente === email) {
        return { success: false, message: "Este email já está cadastrado!" };
      }
    }
    
    var verificacaoTelefone = verificarTelefoneAutorizado(telefone);
    
    if (!verificacaoTelefone.autorizado) {
      return { success: false, message: verificacaoTelefone.motivo };
    }
    
    var novaLinha = [
      email,
      senha,
      nomeUsuario,
      telefone,
      "Cliente",
      formatarDataHoraSistema_(),
      "Ativo",
      "",
      "",
      "",
      false,
      endereco,
      bairro
    ];
    
    abaUsuarios.appendRow(novaLinha);
    SpreadsheetApp.flush();
    marcarTelefoneComoUtilizado(telefone, email);
    
    try {
      enviarEmailConfirmacaoCadastroV3_(email, nomeUsuario, telefone, endereco, bairro);
    } catch (emailError) {}
    
    return {
      success: true,
      message: "✅ Cadastro realizado com sucesso!\nFaça login para continuar."
    };
    
  } catch (error) {
    return { success: false, message: "Erro ao cadastrar: " + error.toString() };
  }
}

// 10. ENVIAR EMAIL DE CONFIRMAÇÃO
function enviarEmailConfirmacaoCadastro(email, nome, telefone, endereco) {
  try {
    var assunto = "Bem-vindo ao JW Piscinas!";
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <h3>Olá, ${nome}!</h3>
        <p>Seu cadastro no sistema JW Piscinas foi realizado com sucesso.</p>
        <p><strong>Seus dados:</strong></p>
        <ul>
          <li><strong>Usuário para login:</strong> ${nome}</li>
          <li><strong>Email:</strong> ${email}</li>
          <li><strong>Telefone:</strong> ${telefone}</li>
          <li><strong>Endereco:</strong> ${endereco || ""}</li>
        </ul>
        <p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      body: "Olá, " + nome + ". Seu código de recuperação de senha JW Piscinas é: " + codigo + ". Ele é válido por 1 hora. Se você não solicitou esta recuperação, ignore este email.",
      htmlBody: corpo
    });
  } catch (error) {}
}

// 11. SISTEMA DE RECUPERAÇÃO DE SENHA
function solicitarRecuperacaoSenha(emailOuTelefone) {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    var usuarioEncontrado = null;
    var linhaEncontrada = -1;
    
    var busca = emailOuTelefone.toString().trim().toLowerCase();
    var buscaNormalizada = normalizarTelefone(emailOuTelefone).toLowerCase();
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailUsuario = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      var nomeUsuario = linha[2] ? linha[2].toString().trim().toLowerCase() : "";
      var telefoneUsuario = linha[3] ? normalizarTelefone(linha[3].toString()).toLowerCase() : "";
      
      if (emailUsuario === busca || nomeUsuario === busca || telefoneUsuario === busca || telefoneUsuario === buscaNormalizada) {
        usuarioEncontrado = {
          email: linha[0].toString().trim(),
          nome: linha[2].toString().trim(),
          linha: i
        };
        linhaEncontrada = i;
        break;
      }
    }
    
    if (!usuarioEncontrado) {
      return { success: false, message: "Usuário não encontrado! Verifique o email, telefone ou nome de usuário." };
    }
    
    var codigo = Math.floor(100000 + Math.random() * 900000).toString();
    var dataExpiracao = new Date();
    dataExpiracao.setHours(dataExpiracao.getHours() + 1);
    
    aba.getRange(linhaEncontrada + 1, 8).setValue(codigo);
    aba.getRange(linhaEncontrada + 1, 9).setValue(dataExpiracao.toISOString());
    SpreadsheetApp.flush();
    
    enviarEmailRecuperacaoSenha(usuarioEncontrado.email, usuarioEncontrado.nome, codigo);
    
    return {
      success: true,
      message: "Código de recuperação enviado para o email cadastrado!",
      email: usuarioEncontrado.email
    };
    
  } catch (error) {
    return { success: false, message: "Erro ao processar recuperação: " + error.toString() };
  }
}

function verificarCodigoRecuperacao(email, codigo) {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailUsuario = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      
      if (emailUsuario === email.toLowerCase()) {
        var codigoArmazenado = linha[7] ? linha[7].toString().trim() : "";
        var dataExpiracaoStr = linha[8] ? linha[8].toString() : "";
        
        if (!codigoArmazenado) {
          return { success: false, message: "Nenhum código de recuperação solicitado." };
        }
        
        if (dataExpiracaoStr) {
          var dataExpiracao = new Date(dataExpiracaoStr);
          if (new Date() > dataExpiracao) {
            aba.getRange(i + 1, 8).setValue("");
            aba.getRange(i + 1, 9).setValue("");
            return { success: false, message: "Código expirado. Solicite um novo." };
          }
        }
        
        if (codigoArmazenado === codigo) {
          return { success: true, message: "Código válido!", email: emailUsuario };
        } else {
          return { success: false, message: "Código incorreto!" };
        }
      }
    }
    
    return { success: false, message: "Email não encontrado." };
    
  } catch (error) {
    return { success: false, message: "Erro ao verificar código." };
  }
}

function redefinirSenha(email, codigo, novaSenha, confirmacaoSenha) {
  try {
    if (!novaSenha || novaSenha.length < 4) {
      return { success: false, message: "A senha deve ter pelo menos 4 caracteres!" };
    }
    
    if (novaSenha.startsWith('0')) {
      return { success: false, message: "A senha não pode começar com zero!" };
    }
    
    if (novaSenha !== confirmacaoSenha) {
      return { success: false, message: "As senhas não coincidem!" };
    }
    
    var verificacao = verificarCodigoRecuperacao(email, codigo);
    if (!verificacao.success) {
      return verificacao;
    }
    
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailUsuario = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      
      if (emailUsuario === email.toLowerCase()) {
        aba.getRange(i + 1, 2).setValue(novaSenha);
        aba.getRange(i + 1, 8).setValue("");
        aba.getRange(i + 1, 9).setValue("");
        SpreadsheetApp.flush();
        
        enviarEmailSenhaRedefinida(email, linha[2].toString().trim());
        
        return {
          success: true,
          message: "✅ Senha redefinida com sucesso! Faça login com a nova senha."
        };
      }
    }
    
    return { success: false, message: "Usuário não encontrado." };
    
  } catch (error) {
    return { success: false, message: "Erro ao redefinir senha: " + error.toString() };
  }
}

function enviarEmailConfirmacaoCadastroV3_(email, nome, telefone, endereco, bairro) {
  try {
    MailApp.sendEmail({
      to: email,
      subject: "Bem-vindo ao JW Piscinas!",
      body: "Ola, " + nome + ". Seu cadastro no sistema JW Piscinas foi realizado com sucesso.",
      htmlBody:
        '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">' +
        '<h2 style="color: #0077B6;">JW Piscinas</h2>' +
        '<h3>Ola, ' + nome + '!</h3>' +
        '<p>Seu cadastro no sistema JW Piscinas foi realizado com sucesso.</p>' +
        '<p><strong>Seus dados:</strong></p>' +
        '<ul>' +
        '<li><strong>Usuario para login:</strong> ' + nome + '</li>' +
        '<li><strong>Email:</strong> ' + email + '</li>' +
        '<li><strong>Telefone:</strong> ' + telefone + '</li>' +
        '<li><strong>Bairro:</strong> ' + (bairro || "") + '</li>' +
        '<li><strong>Endereco:</strong> ' + (endereco || "") + '</li>' +
        '</ul>' +
        '<p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>' +
        '</div>'
    });
  } catch (error) {}
}

function enviarEmailRecuperacaoSenha(email, nome, codigo) {
  try {
    var assunto = "Recuperação de Senha - JW Piscinas";
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <h3>Olá, ${nome}!</h3>
        <p>Recebemos uma solicitação para redefinir sua senha.</p>
        <p><strong>Seu código de recuperação:</strong></p>
        <div style="background: #f8f9fa; padding: 20px; text-align: center; border-radius: 10px; margin: 20px 0;">
          <h1 style="color: #0077B6; letter-spacing: 10px; font-size: 2.5rem;">${codigo}</h1>
        </div>
        <p><strong>Instruções:</strong></p>
        <ol>
          <li>Insira este código no aplicativo JW Piscinas</li>
          <li>O código é válido por 1 hora</li>
          <li>Se você não solicitou esta recuperação, ignore este email</li>
        </ol>
        <p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      htmlBody: corpo
    });
  } catch (error) {}
}

function enviarEmailSenhaRedefinida(email, nome) {
  try {
    var assunto = "Senha Redefinida - JW Piscinas";
    var corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #0077B6;">JW Piscinas</h2>
        <h3>Olá, ${nome}!</h3>
        <p>Sua senha foi redefinida com sucesso.</p>
        <p>Se você não realizou esta alteração, entre em contato imediatamente.</p>
        <p>Atenciosamente,<br><strong>Equipe JW Piscinas</strong></p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      htmlBody: corpo
    });
  } catch (error) {}
}

// 12. FUNÇÕES ADMIN
function adicionarTelefoneAutorizado(telefone, nome) {
  try {
    var aba = getAbaTelefonesAutorizados();
    var telefoneNormalizado = normalizarTelefone(telefone);
    var dados = aba.getDataRange().getValues();
    
    for (var i = 1; i < dados.length; i++) {
      var telExistente = normalizarTelefone(dados[i][0].toString());
      if (telExistente === telefoneNormalizado) {
        return { success: false, message: "Este telefone já está na lista de autorizados!" };
      }
    }
    
    var novaLinha = [
      telefoneNormalizado,
      nome || "",
      formatarDataSistema_(),
      "Ativo",
      "Não",
      ""
    ];
    
    aba.appendRow(novaLinha);
    SpreadsheetApp.flush();
    
    return { success: true, message: "✅ Telefone adicionado à lista de autorizados!" };
    
  } catch (error) {
    return { success: false, message: "Erro: " + error.toString() };
  }
}

function removerTelefoneAutorizado(telefone) {
  try {
    var aba = getAbaTelefonesAutorizados();
    var dados = aba.getDataRange().getValues();
    var telefoneNormalizado = normalizarTelefone(telefone);
    var linhaEncontrada = -1;
    
    for (var i = 1; i < dados.length; i++) {
      var telExistente = normalizarTelefone(dados[i][0].toString());
      if (telExistente === telefoneNormalizado) {
        linhaEncontrada = i + 1;
        break;
      }
    }
    
    if (linhaEncontrada > 0) {
      aba.deleteRow(linhaEncontrada);
      SpreadsheetApp.flush();
      return { success: true, message: "✅ Telefone removido da lista de autorizados!" };
    } else {
      return { success: false, message: "Telefone não encontrado na lista de autorizados." };
    }
    
  } catch (error) {
    return { success: false, message: "Erro: " + error.toString() };
  }
}

function listarTelefonesAutorizados() {
  try {
    var aba = getAbaTelefonesAutorizados();
    var dados = aba.getDataRange().getValues();
    var telefones = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      telefones.push({
        telefone: linha[0] || "",
        nome: linha[1] || "",
        dataAutorizacao: linha[2] || "",
        status: linha[3] || "",
        usado: linha[4] || "Não",
        dataUso: linha[5] || "",
        linha: i + 1
      });
    }
    
    return { success: true, telefones: telefones, total: telefones.length };
    
  } catch (error) {
    return { success: false, message: "Erro: " + error.toString() };
  }
}

function listarClientesAtivos() {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    var clientes = [];
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var status = linha[6] ? linha[6].toString() : "";
      var statusNormalizado = status.toString().trim().toLowerCase();
      var tipo = isUsuarioAdmin_(linha) ? "Administrador" : (linha[4] ? linha[4].toString() : "Cliente");
      
      if ((status === "" || statusNormalizado === "ativo") && tipo !== "Administrador") {
        clientes.push({
          email: linha[0] || "",
          usuario: linha[2] || "",
          telefone: linha[3] || "",
          tipo: tipo,
          dataCadastro: linha[5] || "",
          status: status,
          endereco: linha[11] || "",
          bairro: linha[12] || "",
          linha: i + 1
        });
      }
    }
    
    return { success: true, clientes: clientes, total: clientes.length };
    
  } catch (error) {
    return { success: false, message: "Erro: " + error.toString() };
  }
}

function getTodosUsuarios() {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    var usuarios = [];

    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var status = linha[6] ? linha[6].toString() : "";

      if (status && status !== "Ativo") continue;

      usuarios.push({
        email: linha[0] || "",
        usuario: linha[2] || "",
        telefone: linha[3] || "",
        tipo: isUsuarioAdmin_(linha) ? "Administrador" : "Cliente",
        endereco: linha[11] || "",
        bairro: linha[12] || ""
      });
    }

    return usuarios;
  } catch (error) {
    console.error("Erro em getTodosUsuarios:", error);
    return [];
  }
}
