import { readFile } from 'fs/promises'
import xl from 'excel4node'
import minimist from 'minimist';


const validaciondeaccount = async (name, wb) => {

  const ws = wb.addWorksheet(`${name}`);
  var style = wb.createStyle({
    font: {
      color: '#323136',
      size: 12,
    },
  });


  const file = await readFile(`./account/${name}`, 'utf-8');
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
          ws.cell(10, idcelda).string(JSON.stringify(valor)).style(style)

        } else {
          let titulo = `${key} : ${items}`
          ws.cell(9, idcelda).string(titulo).style(style)
          let valor = json[key][items] === null ? null : json[key][items];
          ws.cell(10, idcelda).string(String.parse(valor)).style(style)



        }
        idcelda = idcelda + 1;
      }
    }
  }//** Fin del for para reccorrer los json*/



}

const proceso = async () => {
  const args = minimist(process.argv.slice(2))

  const wb = new xl.Workbook();
  for (let item in args._) {
    await validaciondeaccount(args._[item], wb);
  }

  wb.write('cuentas.xlsx')




}

proceso() // Metodo void solo para ejecutar