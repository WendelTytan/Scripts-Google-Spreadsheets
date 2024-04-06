function pesquisarArquivos() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet(); 
    var guiaPlan = planilha.getSheetByName("Exemplo"); 
  
    var Pesquisa = guiaPlan.getRange("B2").getValue(); 
  
    if (typeof Pesquisa == "string"){ 
      var Pesquisa = Pesquisa.toLowerCase();
    }
  
    guiaPlan.getRange("A4:C").clearContent(); 
    guiaPlan.getRange("A3:C3").setValues([["Nome", "Link", "Pasta"]]); 
  
    SpreadsheetApp.flush();
  
    var planID = guiaPlan.getRange('B1').getValue(); 
    var pasta = DriveApp.getFolderById(planID); 
  
    var lista = [];
  
    function listarRecursivamente(pasta, nomes) {
      var arquivos = pasta.getFiles();
      var pastas = pasta.getFolders();
  
      while (arquivos.hasNext()) { 
            var arquivo = arquivos.next();
            var Nome = arquivo.getName();


            if (nomes === ""){
            nomes = pasta; 
            }
    
            if (typeof Nome == "string"){
                var Nome = Nome.toLowerCase();
            }

            if(Nome.indexOf(Pesquisa) > -1) { 
            lista.push([arquivo.getName(), arquivo.getUrl(), nomes]); 
            }
        }
  
        while (pastas.hasNext()) {
          if (nomes === ""){
            nomes = pasta; 
          }
          var pasta = pastas.next();
          if (Pesquisa === "") {
            lista.push([pasta.getName(), pasta.getUrl(), nomes]);
          }
          listarRecursivamente(pasta, pasta.getName()); 
        }
      }
      
      
      listarRecursivamente(pasta, "");
  
  if (lista.length == 0){
    Browser.msgBox("ERRO"); 
    return;
  }
  
  guiaPlan.getRange(4,1, lista.length, lista[0].length).setValues(lista);
  
  SpreadsheetApp.flush();
  lista.length = 0;

}