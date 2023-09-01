const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const logic = require('./routes/logic');

const app = express();
const port = 3000;

const viewsDir = path.join(__dirname, 'views');
app.set('view engine', 'ejs');
app.set('views', viewsDir);


const workbook = XLSX.readFile('C:/Users/NetCell/Downloads/RESERVAS 2023.xlsx', { cellDates: true });
const worksheet = workbook.Sheets['Meses'];

app.use((req, res, next) => {
  req.worksheet = worksheet;
  next();
});

app.get('/', (req, res) => {
  const maxCellsToShow = 75;
  const columnToRead = 'AZC';
  const palabraEspecial = 'twin';
  const tableRows = logic.obtenerTablaRows(req.worksheet, maxCellsToShow, columnToRead);
  const huespedesFinal = tableRows[tableRows.length - 1].totalHuespedes
  res.render('resultados', { tableRows, huespedesFinal, palabraEspecial: palabraEspecial.toLowerCase() })
});

app.get('/buscar', (req, res) => {
  // PENDIENTE DE RENDERIZAR RESULTADOS EN EL INDEX
  const workbook = XLSX.readFile('C:/Users/NetCell/Downloads/RESERVAS 2023.xlsx', { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  function convertirAISO8601Completo(fecha) {
    const partesFecha = fecha.split('-');
    const fechaISO = new Date(partesFecha[0], partesFecha[1] - 1, partesFecha[2]);
    const año = fechaISO.getFullYear();
    const mes = (fechaISO.getMonth() + 1).toString().padStart(2, '0');
    const día = fechaISO.getDate().toString().padStart(2, '0');
    const fechaCompleta = `${año}-${mes}-${día}T00:00:00.000Z`;
    return fechaCompleta;
  }

  const fechaOriginal = req.query.fecha;
  const fechaISOBuscada = new Date(convertirAISO8601Completo(fechaOriginal))

  for (const cellAddress in worksheet) {
    if (!worksheet.hasOwnProperty(cellAddress) || cellAddress[0] === '!') continue;
    const cell = worksheet[cellAddress];
    if (cell.t === 'd' && cell.v.getTime() === fechaISOBuscada.getTime()) {
      console.log(`Fecha encontrada en la celda ${cellAddress}`);
    }
  }
});

app.listen(port, () => {
  console.log(`Servidor web activo en http://localhost:${port}`);
});