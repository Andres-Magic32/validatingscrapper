import { readFile } from 'fs/promises'
import xl from 'excel4node'
import minimist from 'minimist';
import fs from 'fs';


const validaciondeaccount = async (name, wb) => {
  console.log('procesando archivo',name)

  const ws = wb.addWorksheet(`${name}`);
  var style = wb.createStyle({
    font: {
      color: '#323136',
      size: 12,
    },
  });


  const file = await readFile(`./trxs/${name}`, 'utf-8');
  const json = JSON.parse(file);

  let idcelda = 1;
  for (let key in json) {

    /**Asignando medidasa */
    ws.column(idcelda).setWidth(50);

    const objet = typeof (json[key]);
    if (objet === 'object') {
      for (let items in json[key]) {
        const objet = typeof (json[key][items]);

        if (objet === 'object') { 

          let titulo = `${key} : ${items}`
          ws.cell(9, idcelda).string(titulo).style(style)
          const valor = json[key][items] === null ? null : json[key][items];
          if (Array.isArray(valor)) {

            let celda = 1
            for (let propiedad in valor) {
              let trx_id = 1
              for (let ite in valor[propiedad]) {                
                ws.cell(15, trx_id).string(ite).style(style)
                ws.cell(15 + celda, trx_id).string(JSON.stringify(valor[propiedad][ite])).style(style)
                trx_id = trx_id + 1;
              }
              celda = celda + 1;
            }

          } else {
            ws.cell(10, idcelda).string(JSON.stringify(valor)).style(style)
          }


        } else {
          let titulo = `${key} : ${items}`
          ws.cell(9, idcelda).string(titulo).style(style)
          let valor = json[key][items] === null ? null : json[key][items];
          ws.cell(10, idcelda).string(JSON.stringify(valor)).style(style)


        }
        idcelda = idcelda + 1;
      }
    }
  }//** Fin del for para reccorrer los json*/



}

const proceso = async () => {
  const args = minimist(process.argv.slice(2))

  const testFolder = './trxs/';
  const wb = new xl.Workbook();
  const nombre = args._[0];  
  
  fs.readdir(testFolder, (err, files) => {
    console.log('Se detectan', files.length, ' documentos')
    files.forEach(element => {
      if(element.slice(element.length - 5) === '.json'){
        validaciondeaccount(element, wb);
        }
    });
  })
  setTimeout(() => {
    const archivoname = nombre === undefined ? 'default' : nombre
    console.log('generando documento : ', archivoname)
    wb.write(`${archivoname}.xlsx`)
  }, 4000);




}

proceso() // Metodo void solo para ejecutar
