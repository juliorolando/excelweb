function convertirAISO8601Completo(fecha) {
    const partesFecha = fecha.split('-');
    const fechaISO = new Date(partesFecha[0], partesFecha[1] - 1, partesFecha[2]);
    const año = fechaISO.getFullYear();
    const mes = (fechaISO.getMonth() + 1).toString().padStart(2, '0');
    const día = fechaISO.getDate().toString().padStart(2, '0');
    const fechaCompleta = `${año}-${mes}-${día}T00:00:00.000Z`;
    return fechaCompleta;
  }

  module.exports = {
    convertirAISO8601Completo
  }