<script src="https://code.jquery.com/jquery-1.10.2.js"></script>
var id;
var config;
var colconfig;
var tipus;
var Pestanya_configuracio;
var fltr; 
var fltrPol; 

var idFullModel="1NpyfKlvJ3_SSowkg0QFQfZZOA_kbuFufda21NXCfEUY";
var urlModel="https://docs.google.com/spreadsheets/d/1NpyfKlvJ3_SSowkg0QFQfZZOA_kbuFufda21NXCfEUY/edit?usp=sharing";
//var urlAplicacio="https://script.google.com/a/macros/xtec.cat/s/AKfycbyaz-KrhOESpoUOrGHncEnSSWixoxXJ6Zvh4Xm7tgO32Az4TIpd/exec";
var urlAplicacio="https://script.google.com/a/macros/xtec.cat/s/AKfycbxHRlIiK-j2CfC9jqGE0cG84_7TG8OATAxdmI3pzJ2fdYnoTDE/exec";


function doGet(e) {
  
  	if (e.queryString.localeCompare("")==0) {
    	var template0 = HtmlService.createTemplateFromFile('index');
      	template0.idFullModel = idFullModel;      
      	template0.urlModel = urlModel;         
        template0.urlAplicacio = urlAplicacio
        return template0.evaluate().setTitle("Eina de creació de presentacions")
                                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);                           
  	} else { 

      	if (e.parameter.ID==undefined || e.parameter.config==undefined){
          
          var missatge="<ul>";
          if (e.parameter.ID==undefined){
             missatge +="<li>No s'ha trobat l'enllaç al full de càlcul on es troben les dades Parametre: ID=</li>";
          }
          if (e.parameter.config==undefined){
            missatge +="<li>No s'ha trobat el nom de la columna de configuració. Parametre: config=</li>";            
          }
          if (e.parameter.config==""){
            missatge +="<li>No s'ha indicat el nom de la columna de configuració. Parametre: config=</li>";            
          }          
          missatge +="</ul>";                      
          var template0 = HtmlService.createTemplateFromFile('pantalla-info');
              template0.urlAplicacio = "https://script.google.com/macros/s/AKfycbyaz-KrhOESpoUOrGHncEnSSWixoxXJ6Zvh4Xm7tgO32Az4TIpd/exec";     

              template0.missatge= missatge;               
          return template0.evaluate().setTitle("Error")
                                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);            
       
      	} else {
          
	        //var cache = CacheService.getPublicCache();
	        // Obtenim l'usuari 
	        var userEmail = Session.getActiveUser().getEmail();

	        //determinar la ID del full de calcul a partir de l'entrada si és una URL complerta 
	        
            try {
                if (e.parameter.ID.indexOf('https')>-1){
                  var ss = SpreadsheetApp.openByUrl(e.parameter.ID);
                  id=ss.getId();                          
                } else {
                  id=e.parameter.ID;          
                }
              }
            catch(e) {
              var template=HtmlService.createTemplateFromFile('pantalla-error');
                    template.Titol_aplicacio = "Error en la URL del full de càlcul";                    
	            	template.Icona = "https://drive.google.com/uc?id=16_To8pkSC-KZPTYve7P49_vHDFwZG800";                        
	            	template.Contacte = "";                            
	            	template.Capçalera = "S";      
	            	template.Tipus_missatge = "id_full_incorrecta";
	            	template.userEmail = Session.getActiveUser().getEmail();          
	            	return template.evaluate().setTitle("Error").setSandboxMode(HtmlService.SandboxMode.IFRAME)
			                 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);    
              }

            var ss = SpreadsheetApp.openById(id);
            var config="Config";
            
            var first = ss.getSheetByName(config);
            var data = first.getDataRange().getValues();
            var colconfig=5;
            for (c=0;c<data[0].length; c++) {
              if (data[0][c]==e.parameter.config){              
                colconfig=c;
              }
            }


            var Pestanya_configuracio=config; 
			var config=config;    
			var tipus=e.parameter.tipus;  
            var colconfig=colconfig;

            fltr=""; 
			if (e.parameter.fltr!=undefined){
				fltr=e.parameter.fltr;            
	        }
                    
            fltrPol=""; 
			if (e.parameter.fltrPol!=undefined){
				fltrPol=e.parameter.fltrPol;            
	        }          
           
          var spread = SpreadsheetApp.openById(id);
           try {
              var cnf_fll = spread.getSheetByName(config); 
           } catch (e) {
              var template=HtmlService.createTemplateFromFile('pantalla-error');
                    template.Titol_aplicacio = "Error en la ID del full de càlcul";                    
	            	template.Icona = "https://drive.google.com/uc?id=16_To8pkSC-KZPTYve7P49_vHDFwZG800";                          
	            	template.Contacte = "";                            
	            	template.Capçalera = "S";      
	            	template.Tipus_missatge = "id_full_incorrecta";
	            	template.userEmail = Session.getActiveUser().getEmail();          
	            	return template.evaluate().setTitle("Error").setSandboxMode(HtmlService.SandboxMode.IFRAME)
			                 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
           }


          //organitzar l'entorn
          //full de càlcul de la configuració
			var cnf_fll = spread.getSheetByName(config); 

            try {
              var cnf_rows = cnf_fll.getDataRange();
            }
            catch (e) {
              var template=HtmlService.createTemplateFromFile('pantalla-error');
                    template.Titol_aplicacio = "Error en la configuració";                    
	            	template.Icona = "https://drive.google.com/uc?id=16_To8pkSC-KZPTYve7P49_vHDFwZG800";                         
	            	template.Contacte = "";                            
	            	template.Capçalera = "S";      
	            	template.Tipus_missatge = "config_incorrecta";
	            	template.userEmail = Session.getActiveUser().getEmail();          
	            	return template.evaluate().setTitle("Error").setSandboxMode(HtmlService.SandboxMode.IFRAME)
			                 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
            }
			var cnf_numRows = cnf_rows.getNumRows();
			var cnf_values = cnf_rows.getValues();

			//recollir les variables de configuracio
			//variables a cercar
			var x=0;
			var nc=0;
			for (var v = 1; v < cnf_values.length; v++) {
	          eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
			}

	      
	        //Comprovar i registrar la ID del full de càlcul amb que es demana l'execució
			var spreadControl = SpreadsheetApp.openById("1Nwb59IZhwEvB6BJsdeZcuJrKJD2tj66QH6bY3qZLjoE");      
	        var ssControl = spreadControl.getSheetByName("Dades"); 
			var ssRows = ssControl.getDataRange();
			var ssValues = ssRows.getValues();        
	        var pas=false;
	        for (f=1;f<ssValues.length;f++){
	          if (ssValues[f][2]==id){
	            var avui=new Date();
	            var numPet=parseInt(ssValues[f][4]+1);            
	            ssControl.getRange(f+1,2).setValue(avui);
	            ssControl.getRange(f+1,4).setValue(e.parameter.config);               
	            ssControl.getRange(f+1,5).setValue(numPet);
	            ssControl.getRange(f+1,10).setValue(Titol_aplicacio);
	            ssControl.getRange(f+1,11).setValue(Correu_gestors.replace(",",", "));    
                ssControl.getRange(f+1,12).setValue("Presentacions");                      
	            pas=true;
	            break;
	          }
	        }
	        if (pas==false){
	          var avui=new Date();
              //eliminar el registre dels diaris curriculars 
              if (Titol_aplicacio.toLowerCase().indexOf("diari curricular")<0){
                ssControl.appendRow([avui,avui,id,config,1,'','','','',Titol_aplicacio,Correu_gestors,"Presentacions"])          
              }
	        }
	      

	      //Comprovem els permisos 
	        if (Acces_restringit=='S') {
	          	if (!userEmail){
	            	var template = HtmlService.createTemplateFromFile('pantalla-error');    
	            	template.Titol_aplicacio = Titol_aplicacio;                    
	            	template.Icona = Icona;                        
	            	template.Contacte = Contacte;                            
	            	template.Capçalera = Capçalera;      
	            	template.Tipus_missatge = "no_usuari";
	            	template.userEmail = "";          
	            	return template.evaluate().setTitle("Error").setSandboxMode(HtmlService.SandboxMode.IFRAME)
			                 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);     
			  
				}

	        	//inicialitzem la variable d'usuari existent a la plataforma
	        	var existeixUsuari=false;
	        	var permisEdit=false;
	      
	        	//Comprovar si hi ha o no un compte d'administrador      
	        	cGestors=Correu_gestors.split(",");
	        	if (cGestors.indexOf(userEmail)>-1){
			  		var existeixUsuari=true;        
	          		permisEdit="S"; 
	        	}
	      
	        	//comprovar si l'usuari està o no autoritzat 
	        	if (existeixUsuari==false) {      
	          		// Obrim el registre de les dades per comprovar quina és la columna que compte els permisos i també si l'usuri actiu està o no autoritzat
	          		var sheet = spread.getSheetByName(Pestanya_dades); 
	          		var data = sheet.getDataRange().getValues();      
	      
	          		var row_colPermis=new Array();
	          		var p=0;
	          		for (c=0;c<data[0].length; c++) {
	            		if ((data[1][c].toLowerCase()).indexOf('permis')>-1){
	              			row_colPermis[p]=c;
	              			p +=1;
	            		}
	          		}
	      
	          		if (tipus=='mapa'){
	            		var existeixUsuari=true;        
	            		permisEdit="S"; 
	          		} else {
	            		for (var i = 3; i < data.length; i++) {
	              			for (p=0;p<row_colPermis.length; p++){
	                			if (data[i][row_colPermis[p]].length>0) {    
	                  				if (data[i][row_colPermis[p]].indexOf(userEmail)>-1 ) {
	                    				existeixUsuari=true;
	                    				if (data[1][row_colPermis[p]].indexOf(":edit")>-1){
	                      					permisEdit="S";
	                    				} else {
	                      					permisEdit="N";              
	                    				}
	                    				break;
	                  				}
	                			}
	              			}
	            		}
	          		}
	        	}
	      
	      	} else {
	        	existeixUsuari=true;
	        	permisEdit='N';
			}

	      	if (existeixUsuari){
	        	Logger.log("existeix");
	        	//usuariCorrecte(userEmail);
	        	//var cache = CacheService.getPublicCache();
	        	//var id = cache.get('id');
	        	//var colconfig = cache.get('colconfig');              
	        	//var Titol_aplicacio = cache.get('Titol_aplicacio');
	        	//var tipus = cache.get('tipus');    

	        	var template = HtmlService.createTemplateFromFile('pantalla');          
	          
	        	template.userEmail = userEmail;      
	        	template.prmEdit = permisEdit;   
	        	template.Permetre_edicio_taula=Permetre_edicio_taula;
	        	template.Titol_aplicacio = Titol_aplicacio;                    
	        	template.Icona = Icona;                        
	        	template.Contacte = Contacte;  

              /*
                if (Correu_gestors.indexOf(userEmail)>-1 && userEmail!=""){
	        	   template.id=id;
                } else {
	        	   template.id='';
                }
              //*/
	        	template.id=id;              
              	template.colconfig=colconfig;
              	template.fltr=fltr;              
              	template.fltrPol=fltrPol;                            
	        	template.Capçalera = Capçalera;
                template.Clau_API = Clau_API;
                
	        	return template.evaluate().setTitle(Titol_aplicacio).setSandboxMode(HtmlService.SandboxMode.IFRAME)
			                 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);       
			  
	      	} else {
	        	var template = HtmlService.createTemplateFromFile('pantalla-error');
	        	template.userEmail = userEmail;                
	        	template.Titol_aplicacio = Titol_aplicacio;                    
	        	template.Icona = Icona;                        
	        	template.Contacte = Contacte;  
	        	template.Capçalera = Capçalera;      
	        	template.Tipus_missatge = "usuari_no_autoritzat";
	        	return template.evaluate().setTitle("Error").setSandboxMode(HtmlService.SandboxMode.IFRAME)
			                 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);     
			  
	      	}
  		}
  	}
}




/*
function usuariCorrecte(userNom){
    userEmail=userNom;
    //var template = HtmlService.createTemplateFromFile('pantalla-taules');    
  
    var template = HtmlService.createTemplateFromFile('pantalla');      
    template.userEmail = userEmail;  
    return template.evaluate().setTitle(Titol_aplicacio).setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);     
}
//*/



//funció de busqueda de dades d'usuaris no xtec
function entradaNoXTEC(nom,contrasenya,id,colconfig){
  /*
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var Pestanya_dades = cache.get('Pestanya_dades');
  //*/
  
  var spread = SpreadsheetApp.openById(id);
  var first = spread.getSheetByName(Pestanya_dades);
  var data = first.getDataRange().getValues();

  var row_colPermis=new Array();
  var p=0;
  for (c=0;c<data[0].length; c++) {
    if ((data[1][c].toLowerCase()).indexOf('permis')>-1){
      row_colPermis[p]=c;
      p +=1;
    }
    if ((data[1][c].toLowerCase()).indexOf('contrasenya')>-1){
      var colContrasenya=c;
    }
  }
   
  var existeixUsuari=false;  
  var missatge="nom: "+nom+"----Contrasenya: "+contrasenya;
  for (var i = 3; i < data.length; i++) {
    for (p=0;p<row_colPermis.length; p++){
       if (data[i][row_colPermis[p]].length>0) {    
          if (data[i][row_colPermis[p]].indexOf('Josep')>-1 && data[i][colContrasenya]=='12345') {         
            existeixUsuari=true;
            missatge +="\nUsuari ("+nom+") i contrasenya ("+contrasenya+") correctes";
            break;
          }
       }
     }
    if (existeixUsuari==true){
      break;
    }
  }
  return missatge;
}


function obrir_model_index(url,id,colconfig){
  if (url.indexOf('https')==-1){
     return "error";
  }
  //url="https://docs.google.com/spreadsheets/d/1BTFM2CIU7_51pfFdBHA5vJUWuqmirl23Gy9gLu1UdO0/edit#gid=730869239";
  
  var spread = SpreadsheetApp.openByUrl(url);
  var ss = spread.getActiveSheet();   
  var rows = ss.getDataRange();
  var dades = rows.getValues();
  return dades;
}
//final de la recollida de dades del full de càlcul de model de l'usuari




//recollida de les dades des de la pantalla 
function recollidaDades(id,colconfig,fltr){
  /*
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var config = cache.get('config');
  var numColConfig=cache.get('colconfig');
  //*/


  //var Pestanya_dades = cache.get('Pestanya_dades');
 
  var userEmail = Session.getActiveUser().getEmail();
  var spread=SpreadsheetApp.openById(id);

   //organitzar l'entorn
  //full de càlcul de la configuració
  var cnf_fll = spread.getSheetByName('config'); 
  var cnf_rows = cnf_fll.getDataRange();
  var cnf_values = cnf_rows.getValues();
  

  for (var v = 1; v < cnf_values.length; v++) {
      eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
  }

  
  
  var first = spread.getSheetByName(Pestanya_dades);
  //var data = first.getDataRange().getValues();
  var data = first.getDataRange().getDisplayValues();  
  


  var row_colPermis=new Array();
  var p=0;
  for (c=0;c<data[0].length; c++) {
    if (data[1][c].toLowerCase().indexOf('permis')>-1){
      row_colPermis[p]=c;
      p +=1;
    }
  }

  //determinar si hi ha o no filtres a aplicar  
  var row_colFiltre=new Array();
  var row_flt=new Array();
  var cmp=new Array();

  /*
  var sheet = SpreadsheetApp.getActiveSheet();
  var criteria = sheet.getFilter().getColumnFilterCriteria(1);
  var newCriteria = SpreadsheetApp.newFilterCriteria()
     .withCriteria(criteria.getCriteriaType(), criteria.getCriteriaValues())
     .setHiddenValues(['c'])
     .build();
 
  //*/
 
  
  //var fltr = cache.get('fltr');
  if (fltr!=""){
    var Filtre=fltr;
  } else if (Filtre_inicial!=""){
    var Filtre=Filtre_inicial;
  } else {
    var Filtre="";  
  }

  
  if (Filtre!=""){
    Filtre=Filtre.substr(1,Filtre.length-2)
    if (Filtre.indexOf("**and**")>-1){
      var tipoFiltre="and";
      row_flt=Filtre.split("**and**");      
    } else if (Filtre.indexOf("**or**")>-1){
      var tipoFiltre="or";
      row_flt=Filtre.split("**or**");      
    } else {
      var tipoFiltre="and";
      row_flt[0]=Filtre;      
    }      

 
    
    for (var f=0;f<row_flt.length;f++){
      cmp[f]=row_flt[f].split("=");
      for (var c=0;c<data[0].length; c++) {
        if (data[0][c].toLowerCase()==cmp[f][0].toLowerCase()){
          row_colFiltre[f]=c;
        }
      }
    }    
  }
  
  
  //selecionar els registres que compleixen el filtre inicial si n'hi ha
  var row_data=new Array();
  //if (typeof(Filtre_inicial) != 'undefined' && Filtre_inicial!=""){  
  if (Filtre!=""){      
     row_data[0]=data[0];
     row_data[1]=data[1];  
     row_data[2]=data[2];  
     var z=3;  
     for (var i = 3; i < data.length; i++) {
       var pas=0;
       for (f=0;f<row_colFiltre.length; f++){
         if (data[i][row_colFiltre[f]]==cmp[f][1] ) {
           pas +=1;           
         }
       }
       if (tipoFiltre=="and"){
         if (pas==row_colFiltre.length){
           row_data[z]=data[i];
           z +=1;    
         }
       } else if (tipoFiltre=="or"){
         if (pas>0){
           row_data[z]=data[i];
           z +=1;    
         }
       }         
     }    
  } else {
    row_data=data;
  }

   //return row_data; 
  //ordenar la matriu abans d'enviar-la a l'usuari
  var row=new Array();
  var row_ord=new Array();    
  if (typeof Columna_ordre!= 'undefined'){  
     if (Columna_ordre!=""){  

       //determinar la quantitat de columnes a ordenar
       var matOrd=Columna_ordre.split(",");
       var colOrd=[];
       for (p=0;p<matOrd.length;p++){       
         //determinar el numero de la columna i eliminar el sentit de l'ordre
         var n=matOrd[p].substr(0,matOrd[p].length-1);
         var o=matOrd[p].substr(matOrd[p].length - 1);
         for (c=0;c<row_data[0].length;c++){
            if (row_data[0][c].toLowerCase()==n.toLowerCase()){
               colOrd[p]=[c,o];
               break;
            }
         }
       }
       
      
       //trespassar dades a una matriu per ordenar
       for (var i = 3; i < row_data.length; i++) {
          row_ord[i-3]=row_data[i];
       }  
       for (p=0;p<matOrd.length;p++){                     
         row_ord.sort(function(a, b) {
           o=colOrd[p][0];
           if (a[o] === b[o]) {
             return 0;
           } else if (colOrd[p][1]=="+"){
             return (a[o] < b[o]) ? -1 : 1;
           } else if (colOrd[p][1]=="-"){
             return (a[o] > b[o]) ? -1 : 1;
           }
         });
       }
       row[0]=row_data[0];
       row[1]=row_data[1];  
       row[2]=row_data[2];
       for (var i = 0; i < row_ord.length; i++) {
         row.push(row_ord[i]);
       }
       //revertir l'array després de la ordenacio
       row_data=row;       
     } 
  }

  
 //retorn de totes les dades si l'usuari és l'administrador 
  cGestors=Correu_gestors.split(",");
  if (cGestors.indexOf(userEmail)>-1){
      return row_data;       
  }
       
  
  //selecionar els registres que relacionats amb l'usuari
  var row=new Array();
  var row_dades=new Array();  
  if (Acces_restringit.toLowerCase()=="n"){
     //row_dades=row_data;
     var z=0;      
     for (var i = 3; i < row_data.length; i++) {
        row_dades[z]=row_data[i];
        z +=1;            
     }    
  } else {
     //row[0]=row_data[0];
     //row[1]=row_data[1];  
     //row[2]=row_data[2];  
     var z=0;  
     for (var i = 3; i < row_data.length; i++) {
       for (p=0;p<row_colPermis.length; p++){
         if (row_data[i][row_colPermis[p]].length>0) {    
           if (row_data[i][row_colPermis[p]].indexOf(userEmail)>-1 ) {
             row_dades[z]=row_data[i];
             z +=1;            
             break;
           }
         }
       }
     }
  }
  
  row[0]=row_data[0];
  row[1]=row_data[1];  
  row[2]=row_data[2];
  for (var i = 0; i < row_dades.length; i++) {
     row.push(row_dades[i]);
  }
  return (row);
}
//Final del la recollida de dades 




function ordenaArray(a, b) {
  //*
    //Columna_ordre="B";
    colOrd=Columna_ordre.charCodeAt(0)-65;
    colOrd=parseInt(colOrd);
    //colOrd=4;
    //*/
    if (a[colOrd] === b[colOrd]) {
        return 0;
    } else {
        return (a[colOrd] < b[colOrd]) ? -1 : 1;
    }
}





//recollida de les dades des de la pantalla 
function recollidaDadesSub(pestanya,colCerca,valCerca,id,colconfig){
 //return (pestanya+"\n"+colCerca+"\n"+valCerca+"\n"+id+"\n"+colconfig);
  
  var spread=SpreadsheetApp.openById(id);

   //organitzar l'entorn
  //full de càlcul de la configuració
  var cnf_fll = spread.getSheetByName('config'); 
  var cnf_rows = cnf_fll.getDataRange();
  var cnf_values = cnf_rows.getValues();
  

  for (var v = 1; v < cnf_values.length; v++) {
      eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
  }


  var ss = spread.getSheetByName(pestanya);
  var data = ss.getDataRange().getDisplayValues(); 

  //return data;
  //var userEmail = Session.getActiveUser().getEmail();  

  
/*
  //determinar si hi ha columnes amb permisos
  var row_colPermis=new Array();
  var p=0;
  for (c=0;c<data[0].length; c++) {
    if (data[1][c].toLowerCase().indexOf('permis')>-1){
      row_colPermis[p]=c;
      p +=1;
    }
  }
  

  //determinar si hi ha o no filtres a aplicar  
  var row_colFiltre=new Array();
  var row_flt=new Array();
  var cmp=new Array();
  if (typeof(Filtre_inicial) !== 'undefined' && Filtre_inicial!=""){
  //if (Filtre_inicial!=""){  
    var row_flt=Filtre_inicial.split("&");
    for (var f=0;f<row_flt.length;f++){
      cmp[f]=row_flt[f].split("=");
      for (var c=0;c<data[0].length; c++) {
        if (data[0][c].toLowerCase()==cmp[f][0].toLowerCase()){
          row_colFiltre[f]=c;
        }
      }
    }    
  }
 
  
  //selecionar els registres que compleixen el filtre si n'hi ha
  var row_data=new Array();
  //if (typeof(Filtre_inicial) !== 'undefined'){
  //if (Filtre_inicial!=""){    
  if (typeof(Filtre_inicial) !== 'undefined' && Filtre_inicial!=""){  
     row_data[0]=data[0];
     row_data[1]=data[1];  
     row_data[2]=data[2];  
     var z=3;  
     for (var i = 3; i < data.length; i++) {
       for (f=0;f<row_colFiltre.length; f++){
         if (data[i][row_colFiltre[f]]==cmp[f][1] ) {
           row_data[z]=data[i];
           z +=1;            
           break;
         }
       }
     }    
  } else {
    row_data=data;
  }

  
  //selecionar els registres que relacionats amb l'usuari
  var row=new Array();
  if (Acces_restringit.toLowerCase()=="n"){
     row2=row;
  } else {
     row[0]=row_data[0];
     row[1]=row_data[1];  
     row[2]=row_data[2];  
     var z=3;  
     for (var i = 3; i < row_data.length; i++) {
       for (p=0;p<row_colPermis.length; p++){
         if (row_data[i][row_colPermis[p]].length>0) {    
           if (row_data[i][row_colPermis[p]].indexOf(userEmail)>-1 ) {
             row[z]=row_data[i];
             z +=1;            
             break;
           }
         }
       }
     }

  }
  
//*/
  
  
  
  var row=data;
 //selecionar els registres buscats
  var row2=new Array();
  //valCerca=valCerca.replace(/[']+/g, '&apos;');
  //valCerca=valCerca.replace(/[&prime;]+/g, '&apos;');
  if (isNaN(valCerca)){
    valCerca=valCerca.replace(new RegExp("′", 'g'),"'");   
  }
  if (!colCerca || !valCerca){
     //row2=row;
  } else {
     row2[0]=row[0];
     row2[1]=row[1]; 
     row2[2]=row[2];  
     var z=3;  
     for (var f = 0; f < row.length; f++) {
       //return (colCerca+"---"+row[f][c]+"-----"+valCerca)                               
       for (var c=0; c<row[0].length;c++){
         if (row[0][c]==colCerca) {    
           if (row[f][c]==valCerca ) {
             row2[z]=row[f];
             z +=1;            
             break;
           }
         }
       }
     }
  }
  
  return (row2);
}






//Recollida de dades de les presentacions permeses i de les plantilles
function recollidaPoligons(id,colconfig,fltrPol){
//*
  var spread=SpreadsheetApp.openById(id);
  var cnf_fll = spread.getSheetByName('config'); 
  var cnf_rows = cnf_fll.getDataRange();
  var cnf_values = cnf_rows.getValues();
  var pestanya='';

  for (var v = 1; v < cnf_values.length; v++) {
    if (cnf_values[v][1]=="Pestanya_poligons"){
      pestanya=cnf_values[v][colconfig];
    }    
    if (cnf_values[v][1]=="Filtre_inicial_poligons"){
      var Filtre_inicial_poligons=cnf_values[v][colconfig];
    }        
  }
  
  if (pestanya==''){
    return;
  }
  
  //per si es pasa un filtre en la url
  if (fltrPol=="" &&  Filtre_inicial_poligons.length>0 ){
     fltrPol=Filtre_inicial_poligons;
  }
  

  
  //existeis la pestanya de poligons
  var ss = spread.getSheetByName(pestanya);  
  try {

  } catch (e) {
    //codi per revisar
    /*
      var template=HtmlService.createTemplateFromFile('pantalla-error');
        template.Titol_aplicacio = "Error en el nom de la pestanya que conté els poligons";                    
	  	template.Icona = "https://drive.google.com/uc?id=16_To8pkSC-KZPTYve7P49_vHDFwZG800";                          
       	template.Contacte = "";                            
       	template.Capçalera = "S";      
       	template.Tipus_missatge = "id_full_incorrecta";
       	template.userEmail = Session.getActiveUser().getEmail();          
       	return template.evaluate().setTitle("Error").setSandboxMode(HtmlService.SandboxMode.IFRAME)
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
     //*/
  }
  
  
  
  var data = ss.getDataRange().getValues();  
  for (c=0; c<data[1].length;c++){
    switch (data[1][c].toLowerCase()){
    case "id":
       var colId=c;
       break;
    case "nompoligon":
       var colNom=c;
       break;    
    case "poligon":
       var colGeo=c;
       break;        
    }
  }
  

//des d'aquí
//* 
  //var fltr = cache.get('fltr');
  if (fltrPol!=""){
    var Filtre=fltrPol;
  } else if (Filtre_inicial_poligons!=""){
    var Filtre=Filtre_inicial_poligons;
  } else {
    var Filtre="";  
  }

  var  row_flt=[];
  if (Filtre!=""){
    Filtre=Filtre.substr(1,Filtre.length-2)
    if (Filtre.indexOf("**and**")>-1){
      var tipoFiltre="and";
      row_flt=Filtre.split("**and**");      
    } else if (Filtre.indexOf("**or**")>-1){
      var tipoFiltre="or";
      row_flt=Filtre.split("**or**");      
    } else {
      var tipoFiltre="and";
      row_flt[0]=Filtre;      
    }      

    var cmp=[];
    var row_colFiltre=[];
    for (var f=0;f<row_flt.length;f++){
      cmp[f]=row_flt[f].split("=");
      for (var c=0;c<data[0].length; c++) {
        if (data[0][c].toLowerCase()==cmp[f][0].toLowerCase()){
          row_colFiltre[f]=c;
        }
      }
    }    
  }
  
  
  //selecionar els registres que compleixen el filtre inicial si n'hi ha
  var row_data=new Array();
  //if (typeof(Filtre_inicial) != 'undefined' && Filtre_inicial!=""){  
  if (Filtre!=""){      
     row_data[0]=data[0];
     row_data[1]=data[1];  
     row_data[2]=data[2];  
     var z=3;  
     for (var i = 3; i < data.length; i++) {
       var pas=0;
       for (f=0;f<row_colFiltre.length; f++){
         if (data[i][row_colFiltre[f]]==cmp[f][1] ) {
           pas +=1;           
         }
       }
       if (tipoFiltre=="and"){
         if (pas==row_colFiltre.length){
           row_data[z]=data[i];
           z +=1;    
         }
       } else if (tipoFiltre=="or"){
         if (pas>0){
           row_data[z]=data[i];
           z +=1;    
         }
       }         
     }    
  } else {
    row_data=data;
  }
//fins aquí

//*/

 //selecionar els registres buscats
  var row=new Array();
  row[0]=row_data[0];
  row[1]=row_data[1];  
  row[2]=row_data[2]; 
  var z=3;
  for (r=3; r<row_data.length;r++){
    if (row_data[r][colGeo].length>0){
      row[z]=row_data[r];
      z +=1;          
    }
  }

  return (row);    
}




//Recollida de dades de les presentacions permeses i de les plantilles
function recollidaPlantilles(id,colconfig){
    /*
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var config = cache.get('config');
  var numColConfig=cache.get('colconfig');  
  //*/
  
  
  var spread=SpreadsheetApp.openById(id);
  var cnf_fll = spread.getSheetByName('config'); 
  var cnf_rows = cnf_fll.getDataRange();
  var cnf_values = cnf_rows.getValues();

  for (var v = 1; v < cnf_values.length; v++) {
    eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
    if (cnf_values[v][1].indexOf("Plantilla_")>-1){
      var str=cnf_values[v][1]+"_titol";
      var valVar=cnf_values[v+1][colconfig].replace(new RegExp("'", 'g'),"&apos;"); 
      eval(str+'="'+valVar+'"');
    }    
  }
 
  var row_tip=Presentacions_permeses.split(",");
  var plantilles=[];
  var p=0; 

  for (var x=0;x<row_tip.length;x++){
     var tipus=row_tip[x];
     for (var s=0;s<spread.getNumSheets();s++){
       var sheet = spread.getSheets()[s];
       var sheetName=sheet.getSheetName();
       if (sheetName===eval("Plantilla_"+tipus)){
         plantilles[p]=[];     
         plantilles[p][0]=row_tip[x].toLowerCase();        
         plantilles[p][1]=eval("Plantilla_"+tipus)         
         plantilles[p][2]=eval("Plantilla_"+tipus+"_titol")
         var plantilla=eval("Plantilla_"+tipus);
         var ss = spread.getSheetByName(plantilla); 
         var ss_rows = ss.getDataRange();
         var ss_numCols = ss_rows.getNumColumns();
         plantilles[p][3] = ss_rows.getValues();
         p +=1;
       }
     }
  }

  return (plantilles);    
}


function recollidaVariables(id,colconfig){
  /*
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var config = cache.get('config');
  var numColConfig=cache.get('colconfig'); 
  //*/
      
  var spread=SpreadsheetApp.openById(id);
  var cnf_fll = spread.getSheetByName('config'); 
  var cnf_rows = cnf_fll.getDataRange();
  var cnf_values = cnf_rows.getValues();
  var variables=new Array();
  for (var i = 0; i < cnf_values.length; i++) {
    //var v=cnf_values[i][1]+"{#}"+cnf_values[i][4]+"{#}"+cnf_values[i][5];     
    var v=cnf_values[i][1]+"{#}"+cnf_values[i][colconfig]+"{#}"+"";  
    variables[i]=v.split("{#}");
  }
  /*
  var nper=permisEdit.length;
  permisEdit[nper][0]='userEmail';
  permisEdit[nper][1]='';
  permisEdit[nper][2]=userEmail; 
  //*/
  var userEmail = Session.getActiveUser().getEmail();   
  variables.push(['userEmail',userEmail,'']);
  
  return (variables);
}




//Comprovar si l'usuari està o no autoritzat a editar les dades
function comprovacioPermis(id,colconfig){
  // Obrim el registre i comprovem si l'usuari/centre ja ha entrat prèviament
  /*
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var config = cache.get('config');
  var numColConfig=cache.get('colconfig');   
  var Pestanya_dades = cache.get('Pestanya_dades');
  //*/

  
  var spread = SpreadsheetApp.openById(id);

    //recollir les variables de configuracio
  //variables a cercar
  var cnf_fll = spread.getSheetByName('config'); 
  var cnf_rows = cnf_fll.getDataRange();
  var cnf_values = cnf_rows.getValues();  
  for (var v = 1; v < cnf_values.length; v++) {
      eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
  }
 
 
  //selecionar els registres relacionats amb l'usuari

  var permisEdit=[];
  var userEmail = Session.getActiveUser().getEmail();  
  permisEdit[0]=userEmail;

  
  if (Acces_restringit.toLowerCase()=="n"){
      permisEdit[1]="N";    
  } else {
    permisEdit[1]="S";     
    permisEdit[2]="N";           
    //cGestors=Correu_gestors.split(",");
    //if (cGestors.indexOf(userEmail)>-1){
    if (Correu_gestors.indexOf(userEmail)>-1){    
      permisEdit[2]="S";       
    } else {
      var first = spread.getSheetByName(Pestanya_dades);
      var data = first.getDataRange().getValues();
      var row_colPermis=new Array();
      var p=0;
      for (c=0;c<data[0].length; c++) {
        if ((data[1][c].toLowerCase()).indexOf('permis')>-1){
          row_colPermis[p]=c;
          p +=1;
        }
      }
      //var permisEdit='N';
      for (var i = 3; i < data.length; i++) {
        for (p=0;p<row_colPermis.length; p++){
          if (data[i][row_colPermis[p]].length>0) {    
            if (data[i][row_colPermis[p]].indexOf(userEmail)>-1 ) {
              if (data[1][row_colPermis[p]].toLowerCase().indexOf(":edit")>-1){
                permisEdit[2]="S";    
              } else {
                permisEdit[2]="N";    
              }
              break;
            }
          }
        }
      }
    }
  }
  return permisEdit; 
}




function salvarDades(id,colconfig,row_dad){
  var permis= comprovacioPermis(id,colconfig);
  var missatge=new Array();  
  if (permis[2]!='S'){
    var userEmail = Session.getActiveUser().getEmail();    
    missatge[0]="L'usuari actual ("+userEmail+") no està autoritzat a realitzar canvis<br><br>No s'ha efectuat cap canvi en les dades";  
    return missatge;    
  } else {
    /*
    var cache = CacheService.getPublicCache();
    var id = cache.get('id');
    var config = cache.get('config');
    var numColConfig=cache.get('colconfig');  
    //*/

    var spread=SpreadsheetApp.openById(id);
    
    //recollir les variables de configuracio
   //variables a cercar
    var cnf_fll = spread.getSheetByName('config'); 
    var cnf_rows = cnf_fll.getDataRange();
    var cnf_values = cnf_rows.getValues();  

    for (var v = 1; v < cnf_values.length; v++) {
        eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
    }
    
    var ss = spread.getSheetByName(Pestanya_desti);  
    var ss = spread.getActiveSheet();   
    var data = ss.getDataRange().getValues();  

   
    for (c=0;c<data[0].length; c++) {
      if (data[1][c].toLowerCase()=='id'){
        var Columna_id=c;
        break;
      }
    }
    
    var missatge=new Array();
    var canvis=0;
    var pasCamps=false;
    for (r=0;r<row_dad.length;r++ ){
      var str=row_dad[r]["name"];
      //var str=row_dad[r][0];      
      //if (str!="undefined"){
        //return row_dad[r];
        var pos=str.indexOf("#$");
        var idReg=parseInt(str.substr(0,pos)); 
        for (var i = 0; i < data.length; i++) {
          if (data[i][Columna_id]==idReg) {
            for (c=0; c<data[0].length;c++){
              var nom=row_dad[r]["name"];            
              //var nom=row_dad[r][0];                        
              nom=nom.substr(pos+2,nom.length);
              if (data[0][c]==nom){
                pasCamps=true;
                var valor_ant=ss.getRange(i+1,c+1).getValue();                                
                var valor_nou=row_dad[r]["value"];                  
                //var valor_nou=row_dad[r][1];                                
                if (valor_ant!=valor_nou){
                  ss.getRange(i+1,c+1).setValue(valor_nou);                    
                  canvis +=1;
                  str=data[i][Columna_id]+"(##)"+parseInt(Columna_id+1)+"(##)"+parseInt(c+1)+"(##)"+nom+"(##)"+valor_ant+"(##)"+valor_nou;
                  missatge[canvis]=str.split("(##)");
                }
                break;
              }          
            }
            break;             
          }
        }
      //}
    }
    if (pasCamps==false) {
      missatge[0]="No hi ha coincidencia entre títol de les columnes i els noms dels camps del formulari d'edició de dades.\nReviseu les etiquetes";      
    } else {
      missatge[0]="Canvis efectuats: "+canvis;  
    }
    return missatge;    
  }
}



//*
function crear_fitxer(idElement,nomColFitxa,idModel,idCarpeta,nomFitxer,id,colconfig,url) {
  //return id_element_apadrinat;
  // Obtenim l'usuari 
  var userEmail = Session.getActiveUser().getEmail();

  var spread=SpreadsheetApp.openById(id);
  //recollir les variables de configuracio 

  var cnf_values = spread.getSheetByName('config').getDataRange().getDisplayValues();  
  for (var v = 1; v < cnf_values.length; v++) {
      eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
  }
   
  var data = spread.getSheetByName(Pestanya_dades).getDataRange().getDisplayValues();
  for (c=0;c<data[0].length; c++) {
    if (data[1][c].toLowerCase()=='id'){
      var columna_id=c;
      break;
    }
  }
  
  var ss_url = "";
  var row_gest=new Array(); 
  var row_reg=new Array();
  Correu_gestors=Correu_gestors.replace(new RegExp(";", 'g'),","); 
  row_gest=Correu_gestors.split(",");
  
  //Determinem la forma d'actuar en funció del tipus de distribucció que ha fet l'administrador llista oberta o tancada
  for (var i = 0; i < data.length; i++) {
     if ( idElement == data[i][columna_id]) {
       row_reg = data[i];       
       var fila_registre=i;
       for (c=0;c<data[i].length;c++) {
         if (data[0][c].toLowerCase()==nomColFitxa){
             var columna_fitxa=c;
             var ss_url=data[i][c];
         }         
         if ((data[1][c].toLowerCase()).indexOf('permis')>-1){
           if (data[i][c].slice(-9)=="@xtec.cat"){
                row_gest.push(data[i][c]);
           }
         }             
       }
       break;
     }
   }
   // Carpeta desti
  /*
   var idCarpeta=ID_carpeta_desti_fitxers;
   var idModel=ID_plantilla_model;
   //*/
  
   var carpeta = DriveApp.getFolderById(idCarpeta);
   var model = DriveApp.getFileById(idModel);

   
   // Copia el full de càlcul i dona permissos d'edició a l'usuari si encara no esta fet
  if (ss_url==""){
    var SS = model.makeCopy(nomFitxer,carpeta);
    SS.addEditors(row_gest);
    var ss_url = SS.getUrl();
    var idFtxCreat=SS.getId();
  }
 
  //return (Pestanya_dades,fila_registre,columna_fitxa,ss_url);  
  //Salvar la url de la fitxa
   ss_url="https://docs.google.com/spreadsheets/d/16LsLj5CxQ30Gg3SJAWEzJt-YoNqlZbWXRYM11M9b70U/edit#gid=1787670486";
   var fila=parseInt(fila_registre+1);
   var columna=parseInt(columna_fitxa+1);
   //spread.getSheetByName(Pestanya_dades).getRange(fila,columna).setValue(ss_url);    
   spread.getSheetByName("Respostes").getRange(5,16).setValue(ss_url);      

   var resposta=new Array();
   resposta[0]=fila_registre;  
   resposta[1]=columna_fitxa;  
   resposta[2]=ss_url;    
   resposta[3]=nomFitxer;
   resposta[4]=idElement;
   resposta[5]=idFtxCreat;  
   resposta[6]=row_gest;  
   return resposta;
}
//*/




//*
function crear_fitxer_apadrina(id_element_apadrinat,id,colconfig) {
  //return id_element_apadrinat;
  //id_element_apadrinat=id_element_apadrinat+1;
  // Obtenim l'usuari 
  var userEmail = Session.getActiveUser().getEmail();
  /*
  var cache = CacheService.getPublicCache();
  var id = cache.get('id');
  var config = cache.get('config');  
  var numColConfig=cache.get('colconfig'); 
  //*/

  var spread=SpreadsheetApp.openById(id);
  //recollir les variables de configuracio
 //variables a cercar
  var cnf_fll = spread.getSheetByName('config'); 
  var cnf_rows = cnf_fll.getDataRange();
  var cnf_values = cnf_rows.getValues();  
  for (var v = 1; v < cnf_values.length; v++) {
      eval("var "+cnf_values[v][1]+"=\""+cnf_values[v][colconfig]+"\"");
  }
   

  var first = spread.getSheetByName(Pestanya_dades);
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
  Correu_gestors=Correu_gestors.replace(new RegExp(";", 'g'),","); 
  row_gest=Correu_gestors.split(",");
  
  //Determinem la forma d'actuar en funció del tipus de distribucció que ha fet l'administrador llista oberta o tancada
  for (var i = 0; i < data.length; i++) {
     if ( id_element_apadrinat == data[i][columna_id]) {
       row_reg = data[i];       
       var fila_registre=i;
       for (c=0;c<data[i].length;c++) {
         if (data[0][c].toLowerCase()=='fitxa'){
             var columna_fitxa=c;
             var ss_url=data[i][c];
         }         
         if (data[0][c].toLowerCase()=='servei territorial'){
             var nom_st=data[i][c];
         }   
         if (data[0][c].toLowerCase()=='nom_crp'){
             var nom_crp=data[i][c];
         }         
         if (data[0][c].toLowerCase()=='nom_centre'){
             var nom_centre=data[i][c];
         }         
         if (data[0][c].toLowerCase()=='localitat_centre'){
             var localitat_centre=data[i][c];
         }         
         if (data[0][c].toLowerCase()=='web_centre'){
             var web_centre=data[i][c];
         }                  
         if (data[0][c].toLowerCase()=='element_apadrinat'){
             var element_apadrinat=data[i][c];
         }           
         if ((data[1][c].toLowerCase()).indexOf('permis')>-1){
           if (data[i][c].slice(-9)=="@xtec.cat"){
                row_gest.push(data[i][c]);
           }
         }             
       }
       break;
     }
   }
   // Carpeta "Centres" 
   var folderId=ID_carpeta_desti_fitxers;
   var masterListId=ID_plantilla_model;
  
   var folder = DriveApp.getFolderById(folderId);
   var masterList = DriveApp.getFileById(masterListId);

   
   // Copia el full de càlcul i dona permissos d'edició a l'usuari si encara no esta fet
  if (ss_url==""){
    var nomFull=nom_crp+"-"+nom_centre+"-"+element_apadrinat;
    var SS = masterList.makeCopy(nomFull,folder);
    SS.addEditors(row_gest);
    var ss_url = SS.getUrl();
  }
  
  //obrir el nou full de càlcul per afegir dades d'identificacio   
   var ss = SpreadsheetApp.openByUrl(ss_url);
   var fullIdent = ss.getSheetByName("1. Dades del centre i del projecte");
   fullIdent.getRange(1,4).setValue(element_apadrinat);    
   fullIdent.getRange(2,4).setValue(nom_centre);      
   fullIdent.getRange(5,2).setValue(nom_centre);        
   fullIdent.getRange(6,2).setValue(localitat_centre);        
   fullIdent.getRange(7,2).setValue(web_centre);          
   fullIdent.getRange(5,4).setValue(nom_crp);          
   fullIdent.getRange(7,4).setValue(nom_st);          
  
  //Salvar la url de la fitxa
   first.getRange(fila_registre+1,columna_fitxa+1).setValue(ss_url);    


   var resposta=new Array();
   resposta[0]=fila_registre;  
   resposta[1]=columna_fitxa;  
   resposta[2]=ss_url;    
   resposta[3]=element_apadrinat;
   resposta[4]=id_element_apadrinat;

   return resposta;
}
//*/

