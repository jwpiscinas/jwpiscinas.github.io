// Diagnosticos antigos de historico.
// As funcoes principais ficam em UtilityService.js. Os nomes abaixo usam prefixo
// para nao sobrescrever funcoes globais do Apps Script.

function debugCriarAbaHistoricoGeralComDadosExemplo() {
  var resultado = criarAbaHistoricoGeral();

  if (!resultado || !resultado.success) {
    return resultado;
  }

  registrarHistorico({
    id: "ORC-425973-581",
    tipo: "Orcamento",
    cliente: "Cawan Oliveira",
    email: "cawanfernandoanjok@hotmail.com",
    telefone: "(16) 99629-4795",
    titulo: "Limpeza de Piscina",
    descricao: "Descricao detalhada do servico",
    valor: "Aguardando",
    status: "Solicitado",
    dataRegistro: "10/04/2026, 19:23:45",
    linkAcao: "tab:orcamentos",
    observacoes: "Registro de exemplo para diagnostico"
  });

  registrarHistorico({
    id: "TESTE-001",
    tipo: "Orcamento",
    cliente: "Cawan Oliveira",
    email: "cawanfernandoanjok@hotmail.com",
    telefone: "(16) 99629-4795",
    titulo: "Manutencao Preventiva",
    descricao: "Manutencao completa da piscina",
    valor: "R$ 350,00",
    status: "Aprovado",
    dataRegistro: new Date().toLocaleString("pt-BR"),
    linkAcao: "tab:orcamentos",
    observacoes: "Registro de exemplo para diagnostico"
  });

  return { success: true, message: "Registros de diagnostico adicionados." };
}

function debugTestarHistorico() {
  var email = "cawanfernandoanjok@hotmail.com";
  var resultado = getHistoricoUsuario(email);
  console.log("RESULTADO DO TESTE:", resultado);
  console.log("QUANTIDADE:", resultado.length);
  return resultado;
}
