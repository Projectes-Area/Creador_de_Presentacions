function crear_fitxa_apadrina_2222(id_element_apadrinat) {
  // Obtenim l'usuari 
  
  /*
  var userEmail = Session.getActiveUser().getEmail();
  
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var Pestanya_control = cache.get('Pestanya_control');
  var spread = SpreadsheetApp.openById(id);
  var first = spread.getSheetByName(Pestanya_control);
  var data = first.getDataRange().getValues();
  
  for (c=0;c<data[0].length; c++) {
    if (data[1][c].toLowerCase()=='id'){
      var columna_id=c;
      break;
    }
  }
  
  var ss_url = "";
  var row_gest=new Array(); 
  var row_reg=new Array();
  //Determinem la forma d'actuar en funció del tipus de distribucció que ha fet l'administrador llista oberta o tancada
  for (var i = 0; i < data.length; i++) {
     if ( id_element_apadrinat == data[i][columna_id]) {
       p=0;
       row_reg = data[i];       
       var fila_registre=i;
       for (c=0;c<data[i].length;c++) {
         if (data[0][c].toLowerCase()=='Fitxa'){
             var columna_fitxa=c;
         }         
         if ((data[1][c].toLowerCase()).indexOf('permis')>-1){
            row_gest[p]=data[i][c];
            p +=1;
         }             
       }
       break;
     }
   }
   // Carpeta "Centres" 
   var folderId="1Ni-jpvApg7p7_2lb-Iz-NQmZi8D1HimY";
   var masterListId="1gPKnj0H_9y5wjFJ0xcr7UnDrxZLSerO2xTpVwnkLwt0";
   var folder = DriveApp.getFolderById(folderId);
   var masterList = DriveApp.getFileById(masterListId);
      
   // Copia el full de càlcul i dona permissos d'edició a l'usuari
   var SS = masterList.makeCopy(userEmail,folder);
   //SS.addEditor(userEmail);
   SS.addEditors(row_gest);
   var ss_url = SS.getUrl();
   
  //Salvar la url de la fitxa
   //first.getRange(fila_registre+1,columna_fitxa+1).setValue(ss_url);    
   var row_resposta=id_element_apadrinat+"##"+ss_url;
  //*/
  var row_resposta="HOLA";

   return row_resposta;
}


/*
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var Pestanya_control = cache.get('Pestanya_control');
  var spread=SpreadsheetApp.openById(id);
  var first = spread.getSheetByName(Pestanya_control);
  var ss = spread.getActiveSheet();   
  
  var data = first.getDataRange().getValues();  

  for (c=0;c<data[0].length; c++) {
    if (data[1][c].toLowerCase()=='id'){
      var Columna_id=c;
      break;
    }
  }
  
  var missatge=new Array();
  var canvis=0;
  for (r=0;r<retorn.length;r++ ){
    var str=retorn[r]["name"];
    var pos=str.indexOf("#$");
    var id=str.substr(0,pos); 
    for (var i = 3; i < data.length; i++) {
        if (data[i][Columna_id]==id) {
          for (c=0; c<data[0].length;c++){
            var nom=retorn[r]["name"];            
            nom=nom.substr(pos+2,nom.length);
            if (data[0][c]==nom){
              var valor_ant=ss.getRange(i+1,c+1).getValue();                                
              var valor_nou=retorn[r]["value"];                  
              if (valor_ant!=valor_nou){
                ss.getRange(i+1,c+1).setValue(valor_nou);    
                canvis +=1;
                missatge[canvis]="<tr><td style='border: 1px solid #cacaca !important; padding: 3px !important;'>"+data[i][Columna_id]+"</td><td style='border: 1px solid #cacaca !important; padding: 3px !important;'>"+nom+"</td><td style='border: 1px solid #cacaca !important; padding: 3px !important;'>"+valor_ant+"</td><td style='border: 1px solid #cacaca !important; padding: 3px !important;'>"+valor_nou+"</td></tr>";
              }
              break;
            }          
          }
          break;             
        }
    }
  }
  missatge[0]="Canvis efectuats: "+canvis;  
  return missatge;
  
//*/