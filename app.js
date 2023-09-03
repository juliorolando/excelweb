const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const busquedaPorColumna = require('./routes/busquedaPorColumna');
const busquedaPorFecha = require('./routes/busquedaPorFecha');

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
  const year = fechaActual.getFullYear();
  const month = String(fechaActual.getMonth() + 1).padStart(2, '0'); // +1 porque los meses se cuentan desde 0
  const day = String(fechaActual.getDate()).padStart(2, '0');
  const fechaFormateada = `${year}-${month}-${day}`;
  const fechaISOBuscada = new Date(busquedaPorFecha.convertirAISO8601Completo(fechaFormateada))
  let valorCelda = [];
  for (const cellAddress in worksheet) {
    if (!worksheet.hasOwnProperty(cellAddress) || cellAddress[0] === '!') continue;
    const cell = worksheet[cellAddress];
    if (cell.t === 'd' && cell.v.getTime() === fechaISOBuscada.getTime()) {
      testeo.push(cellAddress);
    }
  }
  const maxCellsToShow = 75;
  const data = valorCelda[0].replace(/\d+$/, '');
  const dataFinal = busquedaPorColumna.obtenerColumnaPosterior(data)
  const columnToRead = dataFinal.posterior
  const palabraEspecial = 'twin'; // pinto en negrita las reservas que contengan TWIN (y/o agregar las que sean necesarias)
  const tableRows = busquedaPorColumna.obtenerTablaRows(req.worksheet, maxCellsToShow, columnToRead);
  const huespedesFinal = tableRows[tableRows.length - 1].totalHuespedes
  res.render('resultados', { tableRows, huespedesFinal, palabraEspecial: palabraEspecial.toLowerCase() })
});

app.get('/buscar', (req, res) => {

  //LA COLUMNA QUE CONSIGO CON REQ.FECHA, PUEDO APLICARLO PARA LISTAR LAS SALIDAS.
  // PARA LISTAR LAS ENTRADAS, TENGO QUE MOSTRAR EL RESULTADO DE LA COLUMNA SIGUIENTE.

  // PENDIENTE - SEGÚN FECHA Y CELDA SELECCIONADA, PODER DIFERENCIAR ENTRE LAS SALIDAS Y LAS ENTRADAS.

  //## ESTO funciona ok, pero ahora estoy testeando lo otro
  const fechaOriginal = req.query.fecha;
  const fechaISOBuscada = new Date(busquedaPorFecha.convertirAISO8601Completo(fechaOriginal))
  let testeo = [];

  for (const cellAddress in worksheet) {
    if (!worksheet.hasOwnProperty(cellAddress) || cellAddress[0] === '!') continue;
    const cell = worksheet[cellAddress];
    if (cell.t === 'd' && cell.v.getTime() === fechaISOBuscada.getTime()) {
      console.log(`Fecha encontrada en la celda ${cellAddress}`);
      testeo.push(cellAddress);
    }
  }


  function obtenerColumnasPosterior(columna) {
    // Eliminar cualquier número que aparezca al final de la columna
    columna = columna.replace(/\d+$/, '');

    // Validar que la columna resultante sea válida (por ejemplo, que esté en formato "A", "AA", "AAA", etc.)
    const columnaRegex = /^[A-Z]+$/;
    if (!columnaRegex.test(columna)) {
      throw new Error('Formato de columna no válido');
    }

    // Convertir la columna a un número de índice
    let indice = 0;
    for (let i = columna.length - 1, j = 0; i >= 0; i--, j++) {
      const letra = columna[i];
      const valor = letra.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
      indice += valor * Math.pow(26, j);
    }

    // Calcular la columna posterior
    const columnaPosterior = obtenerColumnaDesdeIndice(indice + 1);

    return {
      posterior: columnaPosterior,
    };
  }

  function obtenerColumnaDesdeIndice(indice) {
    let columna = '';
    while (indice > 0) {
      const modulo = (indice - 1) % 26;
      columna = String.fromCharCode('A'.charCodeAt(0) + modulo) + columna;
      indice = Math.floor((indice - 1) / 26);
    }
    return columna;
  }

  // Ejemplo de uso
  const columnaActual = testeo[0].replace(/\d+$/, '');
  const resultado = obtenerColumnasPosterior(columnaActual);
  console.log(`Columna posterior a ${columnaActual}: ${resultado.posterior}`);



});

app.listen(port, () => {
  console.log(`Servidor web activo en http://localhost:${port}`);
});