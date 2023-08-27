const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const logic = require('./routes/logic');
const { format } = require('date-fns'); // Importa la función format de date-fns

const app = express();
const port = 3000; // Puedes cambiar el puerto si es necesario

const viewsDir = path.join(__dirname, 'views');
app.set('view engine', 'ejs');
app.set('views', viewsDir);

// Cargar el archivo Excel
const workbook = XLSX.readFile('C:/Users/NetCell/Downloads/RESERVAS 2023.xlsx');
const worksheet = workbook.Sheets['Meses'];

app.use((req, res, next) => {
    req.worksheet = worksheet; // Pasar worksheet como propiedad de la solicitud
    next();
});

app.get('/', (req, res) => {
    const maxCellsToShow = 75; // Cambia esto al número de celdas que deseas mostrar
    const columnToRead = 'AZC';/* tu columna */
    const palabraEspecial = 'twin';
    const tableRows = logic.obtenerTablaRows(req.worksheet, maxCellsToShow, columnToRead);
    const huespedesFinal = tableRows[tableRows.length - 1].totalHuespedes
    res.render('resultados', { tableRows, huespedesFinal, palabraEspecial: palabraEspecial.toLowerCase() })
});

app.get('/fecha', (req, res) => {
  const workbook = XLSX.readFile('C:/Users/NetCell/Downloads/RESERVAS 2023.xlsx');
  const sheetName = workbook.SheetNames[0]; // Suponiendo que quieres buscar en la primera hoja
  const sheet = workbook.Sheets[sheetName];
  
  const cellValue = getCellValue(sheet, 45169); // Cambia 'A1' por la dirección de la celda que deseas buscar0
  console.log("cellvalue" + cellValue)

  function formatFecha(fecha) {
    const partes = fecha.split('-'); // Dividir la fecha en partes (año, mes, día)
    const dia = parseInt(partes[2], 10);
    const mes = parseInt(partes[1], 10);
    const anio = partes[0];
    const mesFormateado = mes < 10 ? `${mes}` : mes.toString(); // Eliminar 0 si el mes es mayor o igual a 10
    const fechaFormateada = `${dia}/${mesFormateado}/${anio}`;
    return fechaFormateada;
  }

  const fechaInput = req.query.fecha; // Obtener la fecha desde req.query.fecha
  console.log("fecha input"+fechaInput)
  const jsDate =  new Date(fechaInput);
  console.log("jsadate:"+jsDate)
 // const fechaFormateada = formatFecha(jsDate);
  const fechaExcel = convertToExcelDate(jsDate);
  
  console.log(fechaExcel); // Imprimir la fecha formateada
});

// el numero que me devuelve es el titulo de reserva. cellValue.
// las fechas en excel se parsean a ese formato de numero.

function getCellValue(sheet, cellAddress) {
  const cell = sheet[cellAddress];
  if (cell && cell.v !== undefined) {
    return cell.v;
  }
  return undefined;
}

function convertToExcelDate(jsDate) {
  const excelStartDate = new Date('1899-12-30'); // Fecha de inicio de Excel (30 de diciembre de 1899)
  const timeDiff = jsDate.getTime() - excelStartDate.getTime(); // Diferencia de tiempo en milisegundos
  const excelDateValue = timeDiff / (86400 * 1000); // Convertir a valor Excel (en días) y sumar 1
  
  return excelDateValue;
}


app.get('/procesar', (req, res) => {
    const fecha = req.query.fecha;
    res.send(`Fecha seleccionada: ${fecha}`);
  });

  app.get('/findCell', (req, res) => {
   // const searchString = 45170; // The string to search for
    const workbook = XLSX.readFile('C:/Users/NetCell/Downloads/RESERVAS 2023.xlsx');
    const sheetName = workbook.SheetNames[0]; // Supposing you want to search in the first sheet
    const worksheet = workbook.Sheets[sheetName];
  
    const dateToFind = "44562"; // replace with the date you want to find

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cell = worksheet[XLSX.utils.encode_cell({ c: col, r: row })];
        if (cell && cell.t === 'n' && new Date(parseInt(cell.v)) === dateToFind) {
          res.send(`Found date ${dateToFind} at row ${row} and column ${col}`);
          return;
        }
      }
    }
  
    res.send(`Date ${dateToFind} not found in the file.`);
  
  });



app.listen(port, () => {
    console.log(`Servidor web activo en http://localhost:${port}`);
});