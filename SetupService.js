// SetupService.gs - ATUALIZADO

function configurarSistemaInicial() {
  console.log("⚙️ === CONFIGURANDO SISTEMA INICIAL ===");
  
  try {
    var abasCriadas = [];
    
    // Abas existentes
    try { getAbaUsuarios(); abasCriadas.push("Usuarios"); } catch(e) {}
    try { getAbaTelefonesAutorizados(); abasCriadas.push("TelefonesAutorizados"); } catch(e) {}
    try { getAbaOrcamentos(); abasCriadas.push("Orcamentos"); } catch(e) {}
    try { getAbaProdutos(); abasCriadas.push("Produtos"); } catch(e) {}
    try { getAbaManuais(); abasCriadas.push("Manuais"); } catch(e) {}
    try { getAbaCompras(); abasCriadas.push("Compras"); } catch(e) {}
    try { getAbaPagamentos(); abasCriadas.push("Pagamentos"); } catch(e) {}
    try { getAbaServicos(); abasCriadas.push("Servicos"); } catch(e) {}
    try { getAbaNotificacoes(); abasCriadas.push("Notificacoes"); } catch(e) {}
    
    // NOVAS ABAS
    try { getAbaTecnicos(); abasCriadas.push("Tecnicos"); } catch(e) {}
    try { getAbaOrdensServico(); abasCriadas.push("OrdensServico"); } catch(e) {}
    try { getAbaAgenda(); abasCriadas.push("Agenda"); } catch(e) {}
    
    // Adicionar telefones de exemplo
    try {
      adicionarTelefoneAutorizado("16996294795", "Cawan Oliveira");
      adicionarTelefoneAutorizado("11999999999", "Cliente Teste");
    } catch(e) {}
    
    // Criar usuário admin
    var adminCriado = criarUsuarioAdmin();
    
    SpreadsheetApp.flush();
    
    return {
      success: true,
      message: "✅ Sistema configurado com sucesso!",
      abasCriadas: abasCriadas,
      adminCriado: adminCriado.success || false
    };
    
  } catch (error) {
    return { success: false, message: "❌ Erro: " + error.toString() };
  }
}

function criarUsuarioAdmin() {
  try {
    var aba = getAbaUsuarios();
    var dados = aba.getDataRange().getValues();
    var adminExiste = false;
    
    for (var i = 1; i < dados.length; i++) {
      var usuario = dados[i][2] ? dados[i][2].toString().trim().toLowerCase() : "";
      if (usuario === "admin") {
        adminExiste = true;
        break;
      }
    }
    
    if (!adminExiste) {
      var novaLinha = [
        "admin@jwpiscinas.com.br",
        "admin123",
        "admin",
        "(16) 99629-4795",
        "Administrador",
        new Date().toLocaleString('pt-BR'),
        "Ativo",
        "",
        "",
        "",
        true
      ];
      aba.appendRow(novaLinha);
      return { success: true, message: "Usuário admin criado" };
    }
    return { success: true, message: "Admin já existe" };
  } catch (error) {
    return { success: false, message: "Erro: " + error.toString() };
  }
}
