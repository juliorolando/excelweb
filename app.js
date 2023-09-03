const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const busquedaPorColumna = require('./routes/busquedaPorColumna');
const busquedaPorFecha = require('./routes/busquedaPorFecha');
const columnaDeHoy = require('./routes/columnaFechaActual');

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
  const fechaActual = new Date();
  const fechaFormateada = busquedaPorFecha.formatearFecha(fechaActual);
  const fechaISOBuscada = new Date(busquedaPorFecha.convertirAISO8601Completo(fechaFormateada));
  const columnaHoy = columnaDeHoy.buscarValorCeldaHoy(req.worksheet, fechaISOBuscada) 
  const limiteDeFilas = 75;
  const columna = busquedaPorColumna.obtenerColumna(columnaHoy)
  const columnToRead = columna.posterior
  const palabraEspecial = 'twin'; // pinto en negrita las reservas que contengan TWIN (y/o agregar las que sean necesarias)
  const tableRows = busquedaPorColumna.obtenerTablaRows(req.worksheet, limiteDeFilas, columnToRead);
  const huespedesFinal = tableRows[tableRows.length - 1].totalHuespedes
  res.render('resultados', { tableRows, huespedesFinal, palabraEspecial: palabraEspecial.toLowerCase() })
});

app.get('/buscar', (req, res) => {  
  const fechaBuscada = req.query.fecha;
  const fechaISOBuscada = new Date(busquedaPorFecha.convertirAISO8601Completo(fechaBuscada));
  const columnaHoy = columnaDeHoy.buscarValorCeldaHoy(req.worksheet, fechaISOBuscada) 
  const limiteDeFilas = 75;
  const columna = busquedaPorColumna.obtenerColumna(columnaHoy)
  const columnToRead = columna.posterior
  const palabraEspecial = 'twin'; // pinto en negrita las reservas que contengan TWIN (y/o agregar las que sean necesarias)
  const tableRows = busquedaPorColumna.obtenerTablaRows(req.worksheet, limiteDeFilas, columnToRead);
  const huespedesFinal = tableRows[tableRows.length - 1].totalHuespedes
  res.render('resultados_busqueda', { tableRows, huespedesFinal, palabraEspecial: palabraEspecial.toLowerCase() })
});

app.listen(port, () => {
  console.log(`Servidor web activo en http://localhost:${port}`);
});