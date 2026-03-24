// netlify/functions/pagar.js
// Recibe: { year, row, nFactura, fechaPago, montoPagado }
// Actualiza columna G (Estado) = "Pagada" e I (Fecha Pago) = fechaPago del usuario

const DRIVE_ID = "b!94YSNWupIUmh41_AtdOVSPieLm_WNpBEh9tqCJhq7-HE4RJxxbTATpTpoCXdSMrL";

const BASE_PATH = {
  2025: "Administración y Finanzas/2025/Contabilidad EST/Facturación/Facturación EST 2025.xlsx",
  2026: "Administración y Finanzas/2026/Contabilidad/Trayectoria EST/Facturación/Facturación EST 2026.xlsx",
};

const COL_ESTADO = "G";
const COL_FPAGO  = "I";

async function getToken() {
  const body = new URLSearchParams({
    grant_type:    "client_credentials",
    client_id:     process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope:         "https://graph.microsoft.com/.default"
  });
  const resp = await fetch(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", body }
  );
  const data = await resp.json();
  if (!data.access_token) throw new Error("Token error: " + JSON.stringify(data));
  return data.access_token;
}

async function patchCell(token, filePath, cellAddr, value) {
  const encoded = filePath.split('/').map(encodeURIComponent).join('/');
  const url = `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${encoded}:/workbook/worksheets/Sheet1/range(address='${cellAddr}')`;
  const resp = await fetch(url, {
    method: "PATCH",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type":  "application/json"
    },
    body: JSON.stringify({ values: [[value]] })
  });
  if (!resp.ok) throw new Error(`Graph PATCH error ${resp.status}: ${await resp.text()}`);
}

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") return { statusCode: 405, body: "Method not allowed" };

  let body;
  try { body = JSON.parse(event.body); } catch { return { statusCode: 400, body: "Invalid JSON" }; }

  const { year, row, nFactura, fechaPago, montoPagado } = body;
  if (!year || !row || !nFactura) return { statusCode: 400, body: "Missing fields: year, row, nFactura" };

  // Determinar ruta del archivo según año
  const filePath = BASE_PATH[parseInt(year)];
  if (!filePath) return { statusCode: 400, body: "Año no soportado: " + year };

  // Formatear fecha de pago
  let fechaFmt;
  if (fechaPago) {
    // Viene como "YYYY-MM-DD", convertir a "DD/MM/YYYY" para Excel
    const [y, m, d] = fechaPago.split('-');
    fechaFmt = `${d}/${m}/${y}`;
  } else {
    fechaFmt = new Date().toLocaleDateString("es-CL");
  }

  try {
    console.log(`Marcando pagada: ${nFactura} | año ${year} | fila ${row} | fecha ${fechaFmt}`);
    const token = await getToken();
    await patchCell(token, filePath, `${COL_ESTADO}${row}`, "Pagada");
    await patchCell(token, filePath, `${COL_FPAGO}${row}`, fechaFmt);
    console.log(`Factura ${nFactura} actualizada en SharePoint`);
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ok: true, nFactura, row, year, fechaFmt })
    };
  } catch (err) {
    console.error("Error:", err.message);
    return { statusCode: 500, body: JSON.stringify({ ok: false, error: err.message }) };
  }
};
