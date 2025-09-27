//inventariomedicamentos (nombre del appscript)
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("inventario") || ss.insertSheet("inventario");

    const body = e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(body);

    // Cabeceras esperadas:
    // Medicamento | Fecha | Presentación | Acción | Cantidad | Restante | Dinamización | Tipo | Usuario
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Medicamento","Fecha","Presentación","Acción","Cantidad","Restante","Dinamización","Tipo","Usuario"]);
    }

    // Stock restante (último valor de la col 6 = "Restante")
    let stock = 0;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      stock = Number(sheet.getRange(lastRow, 6).getValue()) || 0;
    }

    const cant = Number(data.cantidad) || 0;
    const accion = String(data.accion).toLowerCase();
    if (accion === "ingreso") stock += cant;
    if (accion === "despacho") stock -= cant;

    // Nueva fila
    sheet.appendRow([
      data.medicamento || "",
      new Date(),
      data.presentacion || "",
      data.accion || "",
      cant,
      stock,
      data.nivel || "",
      data.tipo || "",
      data.usuario || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", restante: stock }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("inventario");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const headers = data[0];
  const rows = data.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });

  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}
