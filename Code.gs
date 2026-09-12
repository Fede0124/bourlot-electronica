function doGet(e) {
  const tipo = String((e && e.parameter && e.parameter.tipo) || "").trim().toLowerCase();
  const respuesta = {
    ok: true,
    tipo: tipo,
    datos: leerRegistros(tipo)
  };

  return ContentService
    .createTextOutput(JSON.stringify(respuesta))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {

  try {

    const datos = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hoja = obtenerHoja(ss, datos.tipo);

    hoja.appendRow([
      new Date(),
      datos.accion || "",
      datos.id || "",
      datos.datos || ""
    ]);

    return ContentService
      .createTextOutput(
        JSON.stringify({
          ok: true,
          mensaje: "Guardado correctamente"
        })
      )
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {

    return ContentService
      .createTextOutput(
        JSON.stringify({
          ok: false,
          error: error.toString()
        })
      )
      .setMimeType(ContentService.MimeType.JSON);

  }

}

function leerRegistros(tipo) {
  const nombres = {
    clientes: "Clientes",
    vehiculos: "Vehículos",
    ordenes: "Órdenes",
    turnos: "Turnos",
    presupuestos: "Presupuestos",
    stock: "Stock",
    caja: "Caja"
  };

  const nombre = nombres[tipo];
  if (!nombre) {
    return [];
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja || hoja.getLastRow() < 2) {
    return [];
  }

  const filas = hoja.getDataRange().getValues();
  const registros = new Map();

  filas.slice(1).forEach(fila => {
    const accion = String(fila[1] || "").trim().toLowerCase();
    const id = String(fila[2] || "").trim();
    const datosTexto = String(fila[3] || "").trim();

    if (!id || !datosTexto) {
      return;
    }

    let datos;
    try {
      datos = JSON.parse(datosTexto);
    } catch (error) {
      return;
    }

    if (accion === "eliminar") {
      registros.delete(id);
      return;
    }

    if (accion === "crear" || accion === "modificar" || accion === "editar") {
      registros.set(id, datos);
    }
  });

  return Array.from(registros.values());
}

function obtenerHoja(ss, tipo) {

  const nombres = {

    clientes: "Clientes",
    vehiculos: "Vehículos",
    ordenes: "Órdenes",
    turnos: "Turnos",
    presupuestos: "Presupuestos",
    stock: "Stock",
    caja: "Caja"

  };

  const nombre = nombres[tipo] || "Otros";

  let hoja = ss.getSheetByName(nombre);

  if (!hoja) {

    hoja = ss.insertSheet(nombre);

    hoja.appendRow([
      "Fecha",
      "Acción",
      "ID",
      "Datos"
    ]);

  }

  return hoja;

}
