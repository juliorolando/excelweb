const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const logic = require('./routes/logic');

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
    console.log(req.query)
  });
  
  app.get('/procesar', (req, res) => {
    const fecha = req.query.fecha;
    res.send(`Fecha seleccionada: ${fecha}`);
  });

app.listen(port, () => {
    console.log(`Servidor web activo en http://localhost:${port}`);
});