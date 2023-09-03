function convertirAISO8601Completo(fecha) {
    const partesFecha = fecha.split('-');
    const fechaISO = new Date(partesFecha[0], partesFecha[1] - 1, partesFecha[2]);
    const año = fechaISO.getFullYear();
    const mes = (fechaISO.getMonth() + 1).toString().padStart(2, '0');
    const día = fechaISO.getDate().toString().padStart(2, '0');
    const fechaCompleta = `${año}-${mes}-${día}T00:00:00.000Z`;
    return fechaCompleta;
}

function formatearFecha(fechaActual) {

    year = fechaActual.getFullYear();
    month = String(fechaActual.getMonth() + 1).padStart(2, '0'); // +1 porque los meses se cuentan desde 0
    day = String(fechaActual.getDate()).padStart(2, '0');
    const fechaFormateada = `${year}-${month}-${day}`;
    return fechaFormateada

}
module.exports = {
    convertirAISO8601Completo,
    formatearFecha
}