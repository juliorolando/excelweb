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
const workbook = XLSX.readFile('C:/Users/NetCell/Downloads/RESERVAS 2023.xlsx', { cellDates: true });
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

/* app.get('/fecha', (req, res) => {
  const fecha = req.query.fecha;
  const buscarFecha = new Date(fecha).iso
  res.send(`Fecha seleccionada: ${fecha}`);
}); */

app.get('/buscar', (req, res) => {



 // ESTE CÓDIGO FUNCIONA PERFECTO PARA BUSCAR FECHAS CON EL BUSCADOR. 
 // ME FALTA IMPLEMETARLO AUN, PERO EL CODIGO ANDA BIENNNNN



  // const searchString = 45170; // The string to search for
  const workbook = XLSX.readFile('C:/Users/NetCell/Downloads/RESERVAS 2023.xlsx', { cellDates: true });
  const sheetName = workbook.SheetNames[0]; // Supposing you want to search in the first sheet
  const worksheet = workbook.Sheets[sheetName];

  function convertirAISO8601Completo(fecha) {
    // Dividir la fecha en sus componentes (año, mes, día)
    const partesFecha = fecha.split('-');
    // Crear un nuevo objeto Date con las partes de la fecha
    const fechaISO = new Date(partesFecha[0], partesFecha[1] - 1, partesFecha[2]);
    // Obtener las partes de la fecha en formato ISO8601
    const año = fechaISO.getFullYear();
    const mes = (fechaISO.getMonth() + 1).toString().padStart(2, '0'); // Añadir ceros iniciales si es necesario
    const día = fechaISO.getDate().toString().padStart(2, '0'); // Añadir ceros iniciales si es necesario
    // Construir la fecha en formato ISO8601 completo (con horas, minutos, segundos y milisegundos)
    const fechaCompleta = `${año}-${mes}-${día}T00:00:00.000Z`;
    return fechaCompleta;
  }


  const fechaOriginal = req.query.fecha;
  const fechaBuscada = new Date(convertirAISO8601Completo(fechaOriginal))

  // Recorre todas las celdas de la hoja de cálculo
    for (const cellAddress in worksheet) {
     if (!worksheet.hasOwnProperty(cellAddress) || cellAddress[0] === '!') continue; // Ignora celdas especiales
     const cell = worksheet[cellAddress];
     if (cell.t === 'd' && cell.v.getTime() === fechaBuscada.getTime()) {
       // Si la celda contiene una fecha y coincide con la fecha buscada
       console.log(`Fecha encontrada en la celda ${cellAddress}`);
     }
   } 
});

app.listen(port, () => {
  console.log(`Servidor web activo en http://localhost:${port}`);
});