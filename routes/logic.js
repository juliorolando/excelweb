const XLSX = require('xlsx');

function obtenerTablaRows(worksheet, maxCellsToShow, columnToRead) {
    const tableRows = [];
   let totalHuespedes = 0;

    for (let rowNum = 1; rowNum <= maxCellsToShow; rowNum++) {
        const cellAddress = columnToRead + rowNum;
        const cellValue = worksheet[cellAddress]?.v;

        if (cellValue !== undefined && cellValue !== 'L') {
            const match = cellAddress.match(/[A-Za-z]+(\d+)/);
            if (match) {
                habitacionNum = parseInt(match[1]);
                let numeroDePiso;

                if (habitacionNum === 4) {
                    numeroDePiso = 101;
                } else if (habitacionNum === 7) {
                    numeroDePiso = 102;
                } else if (habitacionNum === 73) {
                    numeroDePiso = "Depto";
                } else if (habitacionNum === 10) {
                    numeroDePiso = 111;
                } else if (habitacionNum === 13) {
                    numeroDePiso = 112;
                } else if (habitacionNum === 16) {
                    numeroDePiso = 113;
                } else if (habitacionNum === 19) {
                    numeroDePiso = 114;
                } else if (habitacionNum === 22) {
                    numeroDePiso = 115;
                } else if (habitacionNum === 25) {
                    numeroDePiso = 116;
                } else if (habitacionNum === 28) {
                    numeroDePiso = 117;
                } else if (habitacionNum === 31) {
                    numeroDePiso = 118;
                } else if (habitacionNum === 34) {
                    numeroDePiso = 119;
                } else if (habitacionNum === 37) {
                    numeroDePiso = 120;
                } else if (habitacionNum === 40) {
                    numeroDePiso = 201;
                } else if (habitacionNum === 43) {
                    numeroDePiso = 202;
                } else if (habitacionNum === 46) {
                    numeroDePiso = 203;
                } else if (habitacionNum === 49) {
                    numeroDePiso = 204;
                } else if (habitacionNum === 52) {
                    numeroDePiso = 205;
                } else if (habitacionNum === 55) {
                    numeroDePiso = 206;
                } else if (habitacionNum === 58) {
                    numeroDePiso = 207;
                } else if (habitacionNum === 61) {
                    numeroDePiso = 208;
                } else if (habitacionNum === 64) {
                    numeroDePiso = 301;
                } else if (habitacionNum === 67) {
                    numeroDePiso = 302;
                } else if (habitacionNum === 70) {
                    numeroDePiso = 303;
                }

                const habitacionAsign = `${numeroDePiso}`;
                let tituloReserva = `${cellValue.toLowerCase()}`;

                let huespedesPorReserva = tituloReserva.match(/x(\d+)/i);
                let cantidadHuespedes = 1; // Valor predeterminado si no se encuentra la cantidad de huéspedes
                if (huespedesPorReserva) {
                    cantidadHuespedes = parseInt(huespedesPorReserva[1]);
                    totalHuespedes += cantidadHuespedes; // Sumar la cantidad de huéspedes al total
                }

                tableRows.push({ habitacionAsign, tituloReserva, totalHuespedes });
            }
        }
    }


    return tableRows;
}

module.exports = {
    obtenerTablaRows
};