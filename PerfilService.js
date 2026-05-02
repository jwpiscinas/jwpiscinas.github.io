// PerfilService.gs - Gerenciamento de Perfil de Usuário

// Atualizar perfil do usuário
function atualizarPerfilUsuario(emailAntigo, dados) {
  try {
    var aba = getAbaUsuarios();
    var dadosPlanilha = aba.getDataRange().getValues();
    
    for (var i = 1; i < dadosPlanilha.length; i++) {
      var linha = dadosPlanilha[i];
      var emailUsuario = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      
      if (emailUsuario === emailAntigo.toLowerCase()) {
        // Verificar se novo email já existe
        if (dados.email && dados.email.toLowerCase() !== emailAntigo.toLowerCase()) {
          for (var j = 1; j < dadosPlanilha.length; j++) {
            if (j !== i && dadosPlanilha[j][0] && 
                dadosPlanilha[j][0].toString().toLowerCase() === dados.email.toLowerCase()) {
              return {
                success: false,
                message: "Este e-mail já está em uso!"
              };
            }
          }
        }
        
        // Verificar se novo nome de usuário já existe
        if (dados.usuario) {
          for (var k = 1; k < dadosPlanilha.length; k++) {
            if (k !== i && dadosPlanilha[k][2] && 
                dadosPlanilha[k][2].toString().toLowerCase() === dados.usuario.toLowerCase()) {
              return {
                success: false,
                message: "Este nome de usuário já está em uso!"
              };
            }
          }
        }
        
        // Atualizar dados
        if (dados.usuario) aba.getRange(i + 1, 3).setValue(dados.usuario);
        if (dados.email) aba.getRange(i + 1, 1).setValue(dados.email);
        if (dados.telefone) aba.getRange(i + 1, 4).setValue(dados.telefone);
        
        SpreadsheetApp.flush();
        
        return {
          success: true,
          message: "✅ Perfil atualizado com sucesso!"
        };
      }
    }
    
    return {
      success: false,
      message: "Usuário não encontrado."
    };
    
  } catch (error) {
    console.error("Erro em atualizarPerfilUsuario:", error);
    return {
      success: false,
      message: "Erro ao atualizar perfil: " + error.toString()
    };
  }
}

// Alterar senha do usuário
function alterarSenhaUsuario(email, senhaAtual, novaSenha) {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var emailUsuario = linha[0] ? linha[0].toString().trim().toLowerCase() : "";
      var senhaUsuario = linha[1] ? linha[1].toString() : "";
      
      if (emailUsuario === email.toLowerCase()) {
        // Verificar senha atual
        if (senhaUsuario !== senhaAtual) {
          return {
            success: false,
            message: "Senha atual incorreta!"
          };
        }
        
        // Validar nova senha
        if (novaSenha.length < 4) {
          return {
            success: false,
            message: "A nova senha deve ter pelo menos 4 caracteres!"
          };
        }
        
        if (novaSenha.startsWith('0')) {
          return {
            success: false,
            message: "A senha não pode começar com zero!"
          };
        }
        
        // Atualizar senha
        aba.getRange(i + 1, 2).setValue(novaSenha);
        SpreadsheetApp.flush();
        
        return {
          success: true,
          message: "✅ Senha alterada com sucesso!"
        };
      }
    }
    
    return {
      success: false,
      message: "Usuário não encontrado."
    };
    
  } catch (error) {
    console.error("Erro em alterarSenhaUsuario:", error);
    return {
      success: false,
      message: "Erro ao alterar senha: " + error.toString()
    };
  }
}

// Salvar avatar do usuário
function salvarAvatarUsuario(email, avatarBase64) {
  try {
    // Extrair dados da imagem
    var matches = avatarBase64.match(/^data:image\/([a-zA-Z]+);base64,(.+)$/);
    if (!matches) {
      return { success: false, message: "Formato de imagem inválido" };
    }
    
    var extensao = matches[1];
    var base64Data = matches[2];
    
    // Criar blob
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/' + extensao, 'avatar_' + email + '.' + extensao);
    
    // Verificar/criar pasta no Drive
    var folders = DriveApp.getFoldersByName("JW_Avatares");
    var folder;
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder("JW_Avatares");
    }
    
    // Remover avatar antigo se existir
    var arquivosAntigos = folder.getFilesByName('avatar_' + email + '.' + extensao);
    while (arquivosAntigos.hasNext()) {
      var arquivoAntigo = arquivosAntigos.next();
      arquivoAntigo.setTrashed(true);
    }
    
    // Salvar novo avatar
    var arquivo = folder.createFile(blob);
    
    // Tornar público
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Salvar URL na planilha (coluna J)
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === email) {
        aba.getRange(i + 1, 10).setValue(arquivo.getUrl());
        SpreadsheetApp.flush();
        break;
      }
    }
    
    return {
      success: true,
      url: arquivo.getUrl(),
      message: "✅ Avatar salvo com sucesso!"
    };
    
  } catch (error) {
    console.error("Erro em salvarAvatarUsuario:", error);
    return {
      success: false,
      message: "Erro ao salvar avatar: " + error.toString()
    };
  }
}

// Obter avatar do usuário
function getAvatarUsuario(email) {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] === email) {
        var avatarUrl = dados[i][9]; // Coluna J
        return avatarUrl || "";
      }
    }
    
    return "";
    
  } catch (error) {
    console.error("Erro em getAvatarUsuario:", error);
    return "";
  }
}