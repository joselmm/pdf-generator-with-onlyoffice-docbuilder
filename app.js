const fs = require('fs');
const childProcess = require('child_process');
const util = require('util');
const Papa = require('papaparse');
const async = require('async');


const docbuilderBinaryPath = "C:\\Program Files\\ONLYOFFICE\\DocumentBuilder\\docbuilder.exe"
const createCvsTemplatePath = __dirname + "\\docbuilder-template-scripts\\crearcsv.docbuilder"
const createPdfTemplatePath = __dirname + "\\docbuilder-template-scripts\\remplazar.docbuilder"
//const docxTemplatePath = __dirname + "\\databasesandtemplate\\" + "Carta ejemplo.docx"
const docxTemplatePath = "C:\\Users\\SORACA\\Desktop\\Plantilla Diploma - Pt 3.4 Guia 14 - para combinar.docx"
//const dataBasePath = __dirname + "\\databasesandtemplate\\basedatos.xlsx";
const dataBasePath = "C:\\Users\\SORACA\\Desktop\\base.xlsx";
const csvPathName = __dirname + "\\databasesandtemplate\\" + "basedatos.csv"
const pdfsPath = __dirname + "\\generated-pdfs\\";
const docbuilderScriptsPath = __dirname + "\\doc-builder-scripts\\"
const extractDataScriptName = "createCsv.docbuilder";
const replaceScriptName = "temporal.docbuilder";
const uniqueProperty = "Cédula";
const concurrencia = 3;
const fileExtension = 'pdf';
(async function () {
    try {
        console.log("Hora inicio: "+new Date().toString())
      // LEER ARCHIVO DE TEMPLATE DE CREAR CSV Y CREAR SCRIPT
      let createCsvScriptContent = fs.readFileSync(createCvsTemplatePath, 'utf8');
      createCsvScriptContent = createCsvScriptContent.replaceAll("{{dataBasePath}}", dataBasePath);
      createCsvScriptContent = createCsvScriptContent.replaceAll("{{csvPathName}}", csvPathName);
      fs.writeFileSync(docbuilderScriptsPath + extractDataScriptName, createCsvScriptContent);
  
      // EJECUTAR SCRIPT PARA CREAR CSV
      const commandToCreateCsv = `"${docbuilderBinaryPath}" "${docbuilderScriptsPath + extractDataScriptName}"`;
      const execPromise = util.promisify(childProcess.exec);
      await execPromise(commandToCreateCsv);
  
      // LEER CONTENIDO DE CREADOR DE PDF
      let createPdfScriptContent = fs.readFileSync(createPdfTemplatePath, 'utf8');
      createPdfScriptContent = createPdfScriptContent.replaceAll("{{pathToDocxTemplate}}", docxTemplatePath);
      var textToBeReplaced = "****";
  
      // ITERAR SOBRE CSV
      let csvContent = fs.readFileSync(csvPathName, 'utf8').trim();
      //console.log(JSON.stringify(csvContent))
      const csvData = Papa.parse(csvContent, {
        header: true,
        dynamicTyping: true
      });
      
      //console.log(csvData.data);
      var eee='oDocument.SearchAndReplace({searchString:"@ENCABEZADO@",replaceString:"@REMPLAZO@"})'
     /*  var filas = csvData.data;
      console.log(filas) */
      /* filas.forEach(async (rowObject) =>{
        //var textoTemporal = createPdfScriptContent.replaceAll("{{pdfName}}", pdfsPath + rowObject.Nombre +".pdf");
        var textoTemporal = createPdfScriptContent;
        for ( key in rowObject){
            var lineEncabezado = (""+eee).replace("@ENCABEZADO@","@"+key+"@")+"\n"+textToBeReplaced;
            lineEncabezado=lineEncabezado.replace("@REMPLAZO@", rowObject[key])
            textoTemporal=textoTemporal.replace(textToBeReplaced,lineEncabezado)
        }
        textoTemporal=textoTemporal.replace(textToBeReplaced,"");
        textoTemporal=textoTemporal.replace("{{pathToPdfs}}",pdfsPath+rowObject.Nombre+".pdf")
        fs.writeFileSync(docbuilderScriptsPath + replaceScriptName, textoTemporal);
        //console.log(docbuilderScriptsPath + replaceScriptName)
        const commandToPdf = `"${docbuilderBinaryPath}" "${docbuilderScriptsPath + replaceScriptName}"`;
        const execPromise = util.promisify(childProcess.exec);
        await execPromise(commandToPdf);
        console.log(textoTemporal)
    }) */
        var startTime =  new Date();
        console.log("Hora inicio creaciones pdf: "+startTime.toString())
        var ordenes = [];
        for (const rowObject of csvData.data) {
            var object = {comando:"",scriptPath:""};
            var textoTemporal = createPdfScriptContent;
            for (key in rowObject) {
              var lineEncabezado = ("" + eee).replace("@ENCABEZADO@", "@" + key + "@") + "\n" + textToBeReplaced;
              lineEncabezado = lineEncabezado.replace("@REMPLAZO@", rowObject[key]);
              textoTemporal = textoTemporal.replace(textToBeReplaced, lineEncabezado);
            }

            const scriptPath = docbuilderScriptsPath + rowObject[uniqueProperty] +".docbuilder";
            const pdfPath = pdfsPath + rowObject[uniqueProperty]+ "."+fileExtension;
            textoTemporal = textoTemporal.replace(textToBeReplaced, "");
            //{{fileExtension}}
            textoTemporal = textoTemporal.replace('{{fileExtension}}',fileExtension);
            textoTemporal = textoTemporal.replace("{{pathToPdfs}}", pdfPath);
            fs.writeFileSync(scriptPath, textoTemporal);
            console.log(textoTemporal)
            const commandToPdf = `"${docbuilderBinaryPath}" "${scriptPath}"`;
            /* const execPromise = util.promisify(childProcess.exec);
            await execPromise(commandToPdf); */
            object.comando=commandToPdf;
            object.scriptPath=scriptPath;
            object[uniqueProperty]=rowObject[uniqueProperty];
           // console.log(textoTemporal);
           ordenes.push(object);
          }
          //console.log(ordenes)

          /* for (let i = 0; i < ordenes.length; i++) {
            const object = ordenes[i];
            const execPromise = util.promisify(childProcess.exec);
            await execPromise(object.comando);
            console.log(new Date().toString()+" - finalizo "+object.scriptPath)
            
          } */
          
        const queue = async.queue(function(task, callback) {
          console.log("se esta ejecutando el: "+task[uniqueProperty])
          //console.log(task)
          const execPromise = util.promisify(childProcess.exec);
          execPromise(task.comando).then(() => {
            //console.log(new Date().toString() + " - finalizo " + task.scriptPath);
            
            callback();
          }).catch((err) => {
            console.error(err);
            callback();
          });
        }, concurrencia);
        /* fs.unlink(task.scriptPath, (err) => {
              if (err) {
                console.error(err);
              } else {
                console.log(`Script eliminado: ${task.scriptPath}`);
              }
            }); */
        // Agregamos las tareas a la cola
        ordenes.forEach((object) => {
          queue.push(object);
        });
        
        // Cuando la cola esté vacía, se llama al callback
        queue.drain = function() {
          console.log("Hora finalizacion: " + new Date().toString());
          console.log("Proceso finalizado");
        };
        //console.log("Hora finalizacion: "+new Date().toString())
    } catch (err) {
      console.error(err);
    }
  })();