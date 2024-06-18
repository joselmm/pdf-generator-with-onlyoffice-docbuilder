const fs = require('fs');
const childProcess = require('child_process');
const util = require('util');
const Papa = require('papaparse');


const docbuilderBinaryPath = "C:\\Program Files\\ONLYOFFICE\\DocumentBuilder\\docbuilder.exe"
const createCvsTemplatePath = __dirname + "\\docbuilder-template-scripts\\crearcsv.docbuilder"
const createPdfTemplatePath = __dirname + "\\docbuilder-template-scripts\\remplazar.docbuilder"
const pathToDocxTemplate = __dirname + "\\databasesandtemplate\\" + "Carta ejemplo.docx"
const dataBasePath = __dirname + "\\databasesandtemplate\\basedatos.xlsx";
const csvPathName = __dirname + "\\databasesandtemplate\\" + "basedatos.csv"
const pdfsPath = __dirname + "\\generated-pdfs\\";
const docbuilderScriptsPath = __dirname + "\\doc-builder-scripts\\"
const extractDataScriptName = "createCsv.docbuilder";
const replaceScriptName = "temporal.docbuilder";

(async function () {
    try {
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
      createPdfScriptContent = createPdfScriptContent.replaceAll("{{pathToDocxTemplate}}", pathToDocxTemplate);
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
      var filas = csvData.data;
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
        for (const rowObject of csvData.data) {
            var textoTemporal = createPdfScriptContent;
            for (key in rowObject) {
              var lineEncabezado = ("" + eee).replace("@ENCABEZADO@", "@" + key + "@") + "\n" + textToBeReplaced;
              lineEncabezado = lineEncabezado.replace("@REMPLAZO@", rowObject[key]);
              textoTemporal = textoTemporal.replace(textToBeReplaced, lineEncabezado);
            }
            textoTemporal = textoTemporal.replace(textToBeReplaced, "");
            textoTemporal = textoTemporal.replace("{{pathToPdfs}}", pdfsPath + rowObject.Nombre + ".pdf");
            fs.writeFileSync(docbuilderScriptsPath + replaceScriptName, textoTemporal);
            const commandToPdf = `"${docbuilderBinaryPath}" "${docbuilderScriptsPath + replaceScriptName}"`;
            const execPromise = util.promisify(childProcess.exec);
            await execPromise(commandToPdf);
           // console.log(textoTemporal);
          }
    } catch (err) {
      console.error(err);
    }
  })();