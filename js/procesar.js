/************************************
 *  LIBRERÍAS USADAS:
 *  XLSX.js     → leer XLSX
 *  ExcelJS     → generar XLSX nuevo
 ************************************/


/************************************
 * CONFIG / CONSTANTES
 ************************************/
const ORDER_RESUMEN = [
  'BONIFICACIÓN/BONIFIC.',
  'GASTOS',
  'RECHAZOS',
  'GASTOS DEL COLEGIO',
  'DESCUENTO DE ANTICIPOS',
  'ANTICIPO',
  'TARJETAS'
];

function toNumber(val) {
  if (!val) return 0;
  if (typeof val === "number") return val;

  let s = val.toString().trim();
  s = s.replace(/\s/g, "");

  // caso 1.234,56
  if (/,\d{2}$/.test(s)) {
    s = s.replace(/\./g, "").replace(",", ".");
  }
  return parseFloat(s) || 0;
}

function extractPeriod(fileName) {
  const s = fileName.replace(/[^\d\-_/]/g, ' ');

  let m = /(\d{2})[-_/](\d{2})[-_/](\d{4})/.exec(s);
  if (m) return `${m[3]}-${m[2]}`;

  m = /(\d{4})[-_/](\d{2})[-_/](\d{2})/.exec(s);
  if (m) return `${m[1]}-${m[2]}`;

  m = /(\d{4})(\d{2})(\d{2})/.exec(s);
  if (m) return `${m[1]}-${m[2]}`;

  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}


/************************************
 *  PROCESAR DESDE EL BOTÓN
 ************************************/
async function procesarArchivo() {
  const input = document.getElementById("archivoXLSX");
  if (!input.files.length) {
    alert("Seleccioná un archivo XLSX primero.");
    return;
  }

  const file = input.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  const sheet = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const headers = data.shift();
  const detalle = [];
  const noMatch = [];

  // buscar índices
  const idxConcepto = headers.findIndex(h => /CONCEPTO/i.test(h));
  const idxBonif = headers.findIndex(h => /BONIFICAC/i.test(h));
  const idxDbCr = headers.findIndex(h => /DB.?CR|DEBITO/i.test(h));
  const idxPagos = headers.findIndex(h => /PAGOS?/i.test(h));
  const idxPercibir = headers.findIndex(h => /A[_\s]*PERCIBIR/i.test(h));

  const period = extractPeriod(file.name);

  data.forEach(row => {
    const concepto = (row[idxConcepto] || '').toString().toUpperCase().replace(/\s+/g, ' ').trim();
    const bonif = toNumber(row[idxBonif]);
    const dbcr = toNumber(row[idxDbCr]);
    const pagos = toNumber(row[idxPagos]);
    const percibir = toNumber(row[idxPercibir]);

    // REGLAS (las mismas que tu Apps Script)

    if (/BONIFICAC|BONIFICACION/.test(concepto)) {
      detalle.push([file.name, "BONIFICACIÓN/BONIFIC.", concepto, pagos, period]);
    }
    else if (bonif !== 0 && !/^TOTALES/i.test(concepto)) {
      detalle.push([file.name, "BONIFICACIÓN/BONIFIC.", concepto, bonif, period]);
    }
    else if (/D[ÉE]BITOS|DB\.?\/?CR/.test(concepto)) {
      detalle.push([file.name, "RECHAZOS", concepto, dbcr, period]);
    }
    else if (/CLIPER|MASTERCARD|MATERCARD|AMERICAN|FAVACARD|FAVA*|CABAL|VISA/.test(concepto)) {
      detalle.push([file.name, "TARJETAS", concepto, percibir, period]);
    }
    else if (/RETENCI[ÓO]N\s+TARJETAS/.test(concepto)) {
      detalle.push([file.name, "TARJETAS (RETENCION)", concepto, percibir, period]);
    }
    else if (/RET\s*FDO\s*RESERVA\s*CC/.test(concepto)) {
      detalle.push([file.name, "GASTOS DEL COLEGIO", concepto, dbcr || bonif, period]);
    }
    else if (/RETENCION\s*FDO/.test(concepto)) {
      detalle.push([file.name, "GASTOS DEL COLEGIO", concepto, percibir, period]);
    }
    else if (/RETENCI[ÓO]N\s*COLEGIO/.test(concepto)) {
      detalle.push([file.name, "GASTOS DEL COLEGIO", concepto, percibir, period]);
    }
    else if (/(RET\s*COFA|RET\.(?!\s*IMPOSITIVA)|\bRETENCION(?!\s+IMPOSITIVA)|DEV|REINT\s*FDO|REINT.*RES|PERMANENTE|SEGURO\s*MALA\s*PRAXIS|AREA\s*PROTEGIDA)/i.test(concepto)) {
      let importe;
      if (/PERMANENTE|MALA\s*PRAXIS|AREA\s*PROTEGIDA/i.test(concepto)) {
        importe = percibir || pagos;
      } else {
        importe = bonif + dbcr + pagos;
      }

      detalle.push([file.name, "GASTOS", concepto, importe, period]);
    }
    else if (/VDESC|DESCUENTO\s*ANT|DESC\.?\s*ANT|RECUP\s*ADEL/i.test(concepto)) {
      detalle.push([file.name, "DESCUENTO DE ANTICIPOS", concepto, pagos || percibir, period]);
    }
    else if (/^ANTICIPO\b/.test(concepto)) {
      detalle.push([file.name, "ANTICIPO", concepto, pagos || bonif, period]);
    }
    else {
      noMatch.push(concepto);
    }
  });

  generarExcel(detalle);

  if (noMatch.length) console.log("NO MATCH:", noMatch.slice(0, 50));

  alert("Archivo procesado. Se descargará el Excel generado.");
}


/************************************
 *   GENERAR EXCEL FINAL (2 HOJAS)
 ************************************/
async function generarExcel(detalleRows) {
  const workbook = new ExcelJS.Workbook();

  /*********** HOJA DETALLE ***********/
  const sh1 = workbook.addWorksheet("Detalle");
  sh1.addRow(["Archivo", "Categoría", "Concepto", "Importe", "Periodo"]);

  detalleRows.forEach(r => sh1.addRow(r));
  sh1.columns = [
    { width: 30 },
    { width: 30 },
    { width: 60 },
    { width: 15 },
    { width: 12 }
  ];
  sh1.getColumn(4).numFmt = "#,##0.00";


  /*********** GENERAR RESUMEN ***********/
  let totals = {
    'BONIFICACIÓN/BONIFIC.': 0,
    'GASTOS': 0,
    'RECHAZOS': 0,
    'GASTOS DEL COLEGIO': 0,
    'DESCUENTO DE ANTICIPOS': 0,
    'ANTICIPO': 0,
    'TARJETAS': 0
  };

  detalleRows.forEach(row => {
    const cat = row[1];
    const imp = Number(row[3]) || 0;

    if (cat === "TARJETAS" || cat === "TARJETAS (RETENCION)") {
      totals["TARJETAS"] += imp;
    } else if (totals.hasOwnProperty(cat)) {
      totals[cat] += imp;
    }
  });

  /*********** HOJA RESUMEN ***********/
  const sh2 = workbook.addWorksheet("Resúmen");
  sh2.addRow(["Categoría", "Total"]);
  sh2.getRow(1).font = { bold: true };

  ORDER_RESUMEN.forEach(cat => {
    sh2.addRow([cat, totals[cat]]);
  });

  sh2.getColumn(2).numFmt = "#,##0.00";
  sh2.columns = [{ width: 30 }, { width: 15 }];

  /*********** DESCARGA ***********/
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "Procesado.xlsx";
  a.click();
}
