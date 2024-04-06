function listarArquivos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet(); 
  var guiaPlan = planilha.getSheetByName("Exemplo"); 
  guiaPlan.getRange("A4:C").clearContent(); 

  guiaPlan.getRange("A3:C3").setValues([["Nome", "Link", "Pasta"]]); 
  
  var planID = guiaPlan.getRange('B1').getValue(); 
  var pastaPrincipal = DriveApp.getFolderById(planID); 
  
  var lista = [];
  
  function listarRecursivamente(pasta, nomes) {
    var arquivos = pasta.getFiles();
    var pastas = pasta.getFolders();

    while (arquivos.hasNext()) {
      if (nomes === ""){
        nomes = pasta; 
      }
      var arquivo = arquivos.next();
      lista.push([arquivo.getName(), arquivo.getUrl(), nomes]); 
    }
    
    while (pastas.hasNext()) {
      if (nomes === ""){
        nomes = pasta; 
      }
      var pasta = pastas.next();
      lista.push([pasta.getName(), pasta.getUrl(), nomes]);
      listarRecursivamente(pasta, pasta.getName()); 
    }
  }
  
  listarRecursivamente(pastaPrincipal, "");
 
  if(lista.length == 0){
    Browser.msgBox("Pasta Vazia"); 
    return;
  }
  
  guiaPlan.getRange(4, 1, lista.length, lista[0].length).setValues(lista);

  Browser.msgBox("Lista Atualizada"); 
}
