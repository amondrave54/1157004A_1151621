/**
  Angel Yesid Mondragón 
  1151621
  angelyesidmoro@ufps.edu.co
*/

function abrirArchivo() {
   crearMenu();
  limpiarDocumento();
}



function crearMenu(){
  const menu = SpreadsheetApp.getUi().createMenu("Acciones");
  menu.addItem("enviar email", 'mandarCorreo').addToUi();
  menu.addItem("enviar a todos", 'enviarCorreoTodos').addToUi();
  menu.addItem("cargar plantilla", 'traerPlantilla').addToUi();
  menu.addItem("Crear pdf", 'crearPDF').addToUi();
}

/**
Funcion que me permite el envio de correo a un contacto individual
*/
function mandarCorreo(){
  // configuramos variables generales de la hoja de cálculo
  const nombreHoja = "base"
  const archivo = SpreadsheetApp.getActiveSpreadsheet();
  const encabezados = 1
  var hoja = archivo.getActiveSheet();
  
  // definimos las constantes de las columnas que conforman el encabezado
  var colId = 1;
  var colCod = 2;
  var colNombre = 3;
  var colApellido = 4;
  var colCalificacion = 5;
  var colCorreo = 6;
  
  if(hoja.getName() == nombreHoja){
    var celdaActiva = hoja.getActiveCell();
    var filaActiva = celdaActiva.getRow();
    var asunto = "Prueba de google Scripts"
    if(filaActiva > encabezados){
      var id = hoja.getRange(filaActiva, colId).getValue();
      var codigo = hoja.getRange(filaActiva, colCod).getValue();
      var nombre = hoja.getRange(filaActiva, colNombre).getValue();
      var apellido = hoja.getRange(filaActiva, colApellido).getValue();
      var calificacion = hoja.getRange(filaActiva, colCalificacion).getValue();
      var correo = hoja.getRange(filaActiva, colCorreo).getValue();
      var fin = '\n cordial saludo.  \n Angel Mondragón';
      var mensaje = 'saludos '+ nombre + ' ' + apellido + '\n'
      + ' su calificación en la materia de nube es: '+ calificacion +'\n'
      +fin;
      Logger.log(mensaje);
      GmailApp.sendEmail(correo, asunto, mensaje);
      hoja.getRange(filaActiva, 7).setBackground("green").setValue("Enviado");
      SpreadsheetApp.getUi().alert("El mensaje fue enviado");
    }
  }
}

/**
* Función que permite enviarle correo electronico a todos los contactos de la lista
* para esta función tendremos que manejar un arreglo que contenga todas las variables
* de las celdas en el determinado rango
*/
function enviarCorreoTodos(){
  // configuración global
  const nombreHoja = "Base";
  
  // Obtenemos el archivo , hojas , e.t.c
  
  var libro = SpreadsheetApp.getActive();
  var hoja = libro.getSheetByName(nombreHoja);
  var contacto = hoja.getRange(2, 1, 6, 6).getValues();
  
  
  contacto.forEach(function(fila){
      var asunto = "Prueba envio de correo masivo Google Scripts";
      var mensaje = "Hola "+fila[2]+ " este es un mensaje enviado desde un script de google";
      GmailApp.sendEmail(fila[5], asunto, mensaje);
 
  })
  SpreadsheetApp.getUi().alert("Mensajes enviados");
}

function traerPlantilla(){
 limpiarDocumento();
 var archivo = SpreadsheetApp.getActive();
 var hojas = archivo.getSheets();
  // obtengo los datos de la hoja de cálculo
 var datos = hojas[0].getDataRange().getValues();
 const ultimaFila = hojas[0].getLastRow()-1;
     
 // Modificación del documento
 var documento = DocumentApp.openById('14BLm23-AujEm7Kiz16f1zKmt0bBX9ED3azRFWNVIAOM');
 documento.getBody().appendParagraph("Calificaciones finales de la materia");
 var columnas = [
     ['id','codigo','nombre','apellido','calificacion','correo']
   ];
 // diseño la tabla
  for(var i = 1; i<ultimaFila; i++){
     const fila = datos[i];
    
     const id = fila[0];
     const codigo = fila[1];
     const nombre = fila[2];
     const apellido = fila[3];
     const calificacion = fila[4];
     const correo = fila[5];
    
    var alumno = [id,codigo,nombre,apellido,calificacion,correo]
    columnas.push(alumno);
  }
  documento.getBody().appendTable(columnas);
}

function limpiarDocumento(){
   var documento = DocumentApp.openById('14BLm23-AujEm7Kiz16f1zKmt0bBX9ED3azRFWNVIAOM');
   documento.getBody().clear();
}

function crearPDF(){
 // Creo las constantes del id del documento y de mi carpeta del drive
  const documentoId = '14BLm23-AujEm7Kiz16f1zKmt0bBX9ED3azRFWNVIAOM';
  const idCarpetaDrive = '1N1IRyKw4ZokOaH-qacwjPacM8qmhj86F';
  
 // obtengo los archivos y carpeta
  var documento = DriveApp.getFileById(documentoId);
  var carpetaDrive = DriveApp.getFolderById(idCarpetaDrive);
 
 // Hago la copia del archivo docs
  const archivoTemporal =  documento.makeCopy(carpetaDrive);
  var documentoPdf = DocumentApp.openById(archivoTemporal.getId());
 // armo el documento del pdf
  const pdf = archivoTemporal.getAs(MimeType.PDF);
 // Creo el PDF
  carpetaDrive.createFile(pdf).setName("1151007A_1151621");
  
  SpreadsheetApp.getUi().alert("El archivo PDF fue creado con éxito");
}
