// Obter aba de manuais
function getAbaManuais() {
  if (typeof shouldUseFirebaseAppData_ === "function" && shouldUseFirebaseAppData_()) {
    return getFirestoreSheetAdapter_("Manuais");
  }

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Manuais");
  if (!aba) {
    aba = planilha.insertSheet("Manuais");
    aba.getRange("A1:G1").setValues([[
      "ID", 
      "Produto", 
      "Categoria", 
      "Tipo", 
      "URL", 
      "Descricao",
      "Icone"
    ]]);
    
    // Manuais de exemplo
    var manuaisExemplo = [
      ["MAN-001", "Cloro Granulado", "Químicos", "PDF", 
       "https://drive.google.com/file/d/1abc123", "Como usar cloro em piscinas", "fas fa-file-pdf"],
      ["MAN-002", "Aspirador de Piscina", "Equipamentos", "Video", 
       "https://www.youtube.com/watch?v=exemplo", "Instruções de uso", "fas fa-video"]
    ];
    
    for (var i = 0; i < manuaisExemplo.length; i++) {
      aba.appendRow(manuaisExemplo[i]);
    }
  }
  return aba;
}

// Obter manuais
function getManuais() {
  try {
    var produtos = getProdutos();
    var manuais = [];
    
    for (var i = 0; i < produtos.length; i++) {
      var produto = produtos[i];

      if (!produto.exibirManual || (!produto.manualURL && !produto.manualVideoURL)) continue;
      
      // Determinar cor baseado no tipo
      var tipo = produto.manualURL ? "PDF" : "Video";
      var corIcone = tipo === "PDF" ? "#dc3545" : 
                    tipo === "Video" ? "#0077B6" : "#6c757d";
      
      // Determinar cor da categoria
      var corCategoria = "";
      var categoria = (produto.categoria || "").toLowerCase();
      if (categoria.includes("químicos") || categoria.includes("quimicos")) {
        corCategoria = "#ffc107"; // Amarelo
      } else if (categoria.includes("equipamentos")) {
        corCategoria = "#28a745"; // Verde
      } else {
        corCategoria = "#6f42c1"; // Roxo
      }
      
      manuais.push({
        id: "PROD-" + produto.id,
        produto: produto.nome || "",
        categoria: produto.categoria || "",
        tipo: tipo,
        url: produto.manualURL || "",
        videoURL: produto.manualVideoURL || "",
        imagemURL: produto.imagemURL || "",
        descricao: produto.manualNome || produto.descricao || "",
        icone: tipo === "Video" ? "fas fa-video" : "fas fa-file-pdf",
        corIcone: corIcone,
        corCategoria: corCategoria
      });
    }
    
    return manuais;
    
  } catch (error) {
    console.error("Erro getManuais:", error);
    return [];
  }
}
