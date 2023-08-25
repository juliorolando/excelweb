const XLSX = require('xlsx');


function obtenerTablaRows(worksheet, maxCellsToShow, columnToRead, palabraEspecial) {
    const tableRows = [];
for (let rowNum = 1; rowNum <= maxCellsToShow; rowNum++) {
    const cellAddress = columnToRead + rowNum;
    const cellValue = worksheet[cellAddress]?.v;
  
    if (cellValue !== undefined && cellValue !== 'L') {
      const match = cellAddress.match(/[A-Za-z]+(\d+)/);
      if (match) {
        habitacionNum = parseInt(match[1]);
        let numeroDePiso;
  
        if (habitacionNum >= 3 && habitacionNum <= 5) {
          numeroDePiso = 101;
        } else if (habitacionNum >= 6 && habitacionNum <= 8) {
          numeroDePiso = 102;
        } else if (habitacionNum >= 72 && habitacionNum <= 74) {
          numeroDePiso = "Depto";
        } else if (habitacionNum >= 9 && habitacionNum <= 11) {
          numeroDePiso = 111;
        } else if (habitacionNum >= 12 && habitacionNum <= 14) {
          numeroDePiso = 112;
        }else if (habitacionNum >= 15 && habitacionNum <= 17) {
          numeroDePiso = 113;
        }else if (habitacionNum >= 18 && habitacionNum <= 20) {
          numeroDePiso = 114;
        }else if (habitacionNum >= 21 && habitacionNum <= 23) {
          numeroDePiso = 115;
        }else if (habitacionNum >= 24 && habitacionNum <= 26) {
          numeroDePiso = 116;
        }else if (habitacionNum >= 27 && habitacionNum <= 29) {
          numeroDePiso = 117;
        }else if (habitacionNum >= 30 && habitacionNum <= 32) {
          numeroDePiso = 118;
        }else if (habitacionNum >= 33 && habitacionNum <= 35) {
          numeroDePiso = 119;
        }else if (habitacionNum >= 36 && habitacionNum <= 38) {
          numeroDePiso = 120;
        }else if (habitacionNum >= 39 && habitacionNum <= 41) {
          numeroDePiso = 201;
        }else if (habitacionNum >= 42 && habitacionNum <= 44) {
          numeroDePiso = 202;
        }else if (habitacionNum >= 45 && habitacionNum <= 47) {
          numeroDePiso = 203;
        }else if (habitacionNum >= 48 && habitacionNum <= 50) {
          numeroDePiso = 204;
        }else if (habitacionNum >= 51 && habitacionNum <= 53) {
          numeroDePiso = 205;
        }else if (habitacionNum >= 54 && habitacionNum <= 56) {
          numeroDePiso = 206;
        }else if (habitacionNum >= 57 && habitacionNum <= 59) {
          numeroDePiso = 207;
        }else if (habitacionNum >= 60 && habitacionNum <= 62) {
          numeroDePiso = 208;
        }else if (habitacionNum >= 63 && habitacionNum <= 65) {
          numeroDePiso = 301;
        }else if (habitacionNum >= 66 && habitacionNum <= 68) {
          numeroDePiso = 302;
        }else if (habitacionNum >= 69 && habitacionNum <= 71) {
          numeroDePiso = 303;
        } 
        let tituloReserva = `${cellValue.toLowerCase()}`;
        const habitacionAsign = `${numeroDePiso}`;
/*         if (tituloReserva.includes(palabraEspecial)) {
            tituloReserva = `<strong>${tituloReserva}</strong>`;
            console.log(tituloReserva)
          } */

        tableRows.push({habitacionAsign, tituloReserva});
      }
    }
}
return tableRows;
}
  
module.exports = {
    obtenerTablaRows
  };