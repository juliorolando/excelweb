
function buscarValorCeldaHoy(worksheet, fechaISOBuscada) {

let valorCelda = [];
for (const cellAddress in worksheet) {
  if (!worksheet.hasOwnProperty(cellAddress) || cellAddress[0] === '!') continue;
  const cell = worksheet[cellAddress];
  if (cell.t === 'd' && cell.v.getTime() === fechaISOBuscada.getTime()) {
    valorCelda.push(cellAddress);
  }
}

 return valorCelda[0].replace(/\d+$/, '');

}

module.exports = {
    buscarValorCeldaHoy
}