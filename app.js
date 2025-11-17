// ========================= CONFIGURACIÓN EMISOR =========================
// Basado en tu CFDI de ejemplo de NOMAD GENETICS (RFC, régimen, CP, etc.)
const EMISOR = {
  nombre: "NOMAD GENETICS",
  rfc: "NIN190116954",
  regimenFiscal: "601",          // General de Ley Personas Morales
  direccion: "LA MORENA No. 811, Col. NARVARTE PONIENTE, CP 03020, CDMX, MÉXICO",
  cp: "03020"
};

// ========================= ESTADO GLOBAL =========================
let satCatalogs = null;

let clientes = [];
let productos = [];
let conceptosFactura = [];

const LS_KEYS = {
  clientes: "factu_clientes",
  productos: "factu_productos"
};

// ========================= UTILIDADES =========================
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
}

function formatoMoneda(n) {
  return (n || 0).toFixed(2);
}

function escapeXml(str) {
  if (!str && str !== 0) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function hoyIso() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return (
    d.getFullYear() + "-" +
    pad(d.getMonth() + 1) + "-" +
    pad(d.getDate()) + "T" +
    pad(d.getHours()) + ":" +
    pad(d.getMinutes()) + ":" +
    pad(d.getSeconds())
  );
}

// ========================= CARGA INICIAL =========================
document.addEventListener("DOMContentLoaded", () => {
  initNavegacion();
  cargarDesdeLocalStorage();
  cargarSatCatalogs();
  initClientes();
  initProductos();
  initFacturacion();
});

// ========================= NAVEGACIÓN =========================
function initNavegacion() {
  const buttons = document.querySelectorAll(".nav-btn");
  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      buttons.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");

      const section = btn.dataset.section;
      document
        .querySelectorAll(".section")
        .forEach((sec) => sec.classList.remove("visible"));
      document
        .getElementById(`sec-${section}`)
        .classList.add("visible");
    });
  });
}

// ========================= LOCAL STORAGE =========================
function cargarDesdeLocalStorage() {
  try {
    const cli = localStorage.getItem(LS_KEYS.clientes);
    const prd = localStorage.getItem(LS_KEYS.productos);
    clientes = cli ? JSON.parse(cli) : [];
    productos = prd ? JSON.parse(prd) : [];
  } catch (e) {
    console.warn("Error leyendo localStorage", e);
    clientes = [];
    productos = [];
  }
}

function guardarEnLocalStorage() {
  localStorage.setItem(LS_KEYS.clientes, JSON.stringify(clientes));
  localStorage.setItem(LS_KEYS.productos, JSON.stringify(productos));
}

// ========================= SAT CATALOGS =========================
async function cargarSatCatalogs() {
  try {
    const res = await fetch("sat_catalogs.json");
    if (!res.ok) throw new Error("sat_catalogs.json no encontrado");
    satCatalogs = await res.json();
    console.log("SAT catalogs cargados", satCatalogs);
  } catch (e) {
    console.warn("No se pudieron cargar los catálogos SAT; uso valores mínimos", e);
    // Fallback mínimo
    satCatalogs = {
      formaPago: [
        { clave: "01", descripcion: "Efectivo" },
        { clave: "02", descripcion: "Cheque nominativo" },
        { clave: "03", descripcion: "Transferencia electrónica de fondos" }
      ],
      metodoPago: [
        { clave: "PUE", descripcion: "Pago en una sola exhibición" },
        { clave: "PPD", descripcion: "Pago en parcialidades o diferido" }
      ],
      regimenFiscal: [
        { clave: "601", descripcion: "General de Ley Personas Morales" },
        { clave: "605", descripcion: "Sueldos y Salarios e Ingresos Asimilados a Salarios" }
      ],
      usoCFDI: [
        { clave: "D01", descripcion: "Honorarios médicos, dentales y gastos hospitalarios." },
        { clave: "G03", descripcion: "Gastos en general" }
      ],
      prodServ: [],
      claveUnidad: []
    };
  }

  poblarCombosSat();
}

function poblarCombosSat() {
  // Cliente: régimen y uso
  const selReg = document.getElementById("cli-regimen");
  const selUso = document.getElementById("cli-uso-cfdi");
  selReg.innerHTML = `<option value="">Selecciona...</option>`;
  satCatalogs.regimenFiscal.forEach((r) => {
    const opt = document.createElement("option");
    opt.value = r.clave;
    opt.textContent = `${r.clave} - ${r.descripcion}`;
    selReg.appendChild(opt);
  });
  selUso.innerHTML = `<option value="">Selecciona...</option>`;
  satCatalogs.usoCFDI.forEach((u) => {
    const opt = document.createElement("option");
    opt.value = u.clave;
    opt.textContent = `${u.clave} - ${u.descripcion}`;
    selUso.appendChild(opt);
  });

  // Factura: forma de pago, método
  const selForma = document.getElementById("fac-forma-pago");
  const selMetodo = document.getElementById("fac-metodo-pago");
  selForma.innerHTML = "";
  satCatalogs.formaPago.forEach((f) => {
    const opt = document.createElement("option");
    opt.value = f.clave;
    opt.textContent = `${f.clave} - ${f.descripcion}`;
    selForma.appendChild(opt);
  });
  selMetodo.innerHTML = "";
  satCatalogs.metodoPago.forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m.clave;
    opt.textContent = `${m.clave} - ${m.descripcion}`;
    selMetodo.appendChild(opt);
  });

  // Producto: datalist de ClaveProdServ y ClaveUnidad (TODOS los valores del Excel)
  const dlProdServ = document.getElementById("lista-clave-prodserv");
  const dlUnidad = document.getElementById("lista-clave-unidad");
  dlProdServ.innerHTML = "";
  satCatalogs.prodServ.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = `${p.clave} - ${p.descripcion}`;
    dlProdServ.appendChild(opt);
  });
  dlUnidad.innerHTML = "";
  satCatalogs.claveUnidad.forEach((u) => {
    const texto = u.simbolo
      ? `${u.clave} - ${u.nombre} (${u.simbolo})`
      : `${u.clave} - ${u.nombre}`;
    const opt = document.createElement("option");
    opt.value = texto;
    dlUnidad.appendChild(opt);
  });
}

// ========================= CLIENTES =========================
function initClientes() {
  const form = document.getElementById("form-cliente");
  const btnLimpiar = document.getElementById("btn-cli-limpiar");

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const id = form.dataset.editId || uid();
    const data = {
      id,
      nombre: document.getElementById("cli-nombre").value.trim(),
      rfc: document.getElementById("cli-rfc").value.trim().toUpperCase(),
      regimen: document.getElementById("cli-regimen").value,
      usoCfdi: document.getElementById("cli-uso-cfdi").value,
      cp: document.getElementById("cli-cp").value.trim(),
      direccion: document.getElementById("cli-direccion").value.trim(),
      email: document.getElementById("cli-email").value.trim(),
      telefono: document.getElementById("cli-telefono").value.trim()
    };

    if (!data.nombre || !data.rfc || !data.regimen || !data.usoCfdi || !data.cp) {
      alert("Revisa los datos obligatorios del cliente.");
      return;
    }

    const idx = clientes.findIndex((c) => c.id === id);
    if (idx >= 0) {
      clientes[idx] = data;
    } else {
      clientes.push(data);
    }
    guardarEnLocalStorage();
    renderClientes();
    renderClientesEnFactura();
    limpiarFormCliente();
  });

  btnLimpiar.addEventListener("click", limpiarFormCliente);

  renderClientes();
  renderClientesEnFactura();
}

function renderClientes() {
  const tbody = document.querySelector("#tabla-clientes tbody");
  tbody.innerHTML = "";

  clientes.forEach((c) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${c.nombre}</td>
      <td>${c.rfc}</td>
      <td>${c.regimen}</td>
      <td>${c.usoCfdi}</td>
      <td>${c.cp}</td>
      <td class="actions">
        <button type="button" data-id="${c.id}" class="btn-cli-edit">Editar</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.querySelectorAll(".btn-cli-edit").forEach((btn) => {
    btn.addEventListener("click", () => {
      editarCliente(btn.dataset.id);
    });
  });
}

function editarCliente(id) {
  const c = clientes.find((x) => x.id === id);
  if (!c) return;
  const form = document.getElementById("form-cliente");
  form.dataset.editId = c.id;
  document.getElementById("cli-nombre").value = c.nombre;
  document.getElementById("cli-rfc").value = c.rfc;
  document.getElementById("cli-regimen").value = c.regimen;
  document.getElementById("cli-uso-cfdi").value = c.usoCfdi;
  document.getElementById("cli-cp").value = c.cp;
  document.getElementById("cli-direccion").value = c.direccion;
  document.getElementById("cli-email").value = c.email;
  document.getElementById("cli-telefono").value = c.telefono;
}

function limpiarFormCliente() {
  const form = document.getElementById("form-cliente");
  form.reset();
  delete form.dataset.editId;
}

// ========================= PRODUCTOS =========================
function initProductos() {
  const form = document.getElementById("form-producto");
  const btnLimpiar = document.getElementById("btn-prd-limpiar");

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const id = form.dataset.editId || uid();

    const claveProdServText = document.getElementById("prd-clave-prodserv").value.trim();
    const claveUnidadText = document.getElementById("prd-clave-unidad").value.trim();

    const claveProdServ = claveProdServText.split(" - ")[0] || "";
    const claveUnidad = claveUnidadText.split(" - ")[0] || "";

    const data = {
      id,
      descripcion: document.getElementById("prd-descripcion").value.trim(),
      claveProdServ,
      claveUnidad,
      unidad: document.getElementById("prd-unidad").value.trim(),
      precio: parseFloat(document.getElementById("prd-precio").value || "0"),
      gravaIva: document.getElementById("prd-grava-iva").checked
    };

    if (!data.descripcion || !data.precio) {
      alert("Revisa los datos del producto.");
      return;
    }

    const idx = productos.findIndex((p) => p.id === id);
    if (idx >= 0) {
      productos[idx] = data;
    } else {
      productos.push(data);
    }
    guardarEnLocalStorage();
    renderProductos();
    limpiarFormProducto();
  });

  btnLimpiar.addEventListener("click", limpiarFormProducto);

  renderProductos();
}

function renderProductos() {
  const tbody = document.querySelector("#tabla-productos tbody");
  tbody.innerHTML = "";

  productos.forEach((p) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${p.descripcion}</td>
      <td>${p.claveProdServ}</td>
      <td>${p.claveUnidad}</td>
      <td>${p.unidad}</td>
      <td>${formatoMoneda(p.precio)}</td>
      <td>${p.gravaIva ? "Sí" : "No"}</td>
      <td class="actions">
        <button type="button" data-id="${p.id}" class="btn-prd-edit">Editar</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.querySelectorAll(".btn-prd-edit").forEach((btn) => {
    btn.addEventListener("click", () => {
      editarProducto(btn.dataset.id);
    });
  });
}

function editarProducto(id) {
  const p = productos.find((x) => x.id === id);
  if (!p) return;
  const form = document.getElementById("form-producto");
  form.dataset.editId = p.id;
  document.getElementById("prd-descripcion").value = p.descripcion;
  document.getElementById("prd-clave-prodserv").value = p.claveProdServ;
  document.getElementById("prd-clave-unidad").value = p.claveUnidad;
  document.getElementById("prd-unidad").value = p.unidad;
  document.getElementById("prd-precio").value = p.precio;
  document.getElementById("prd-grava-iva").checked = !!p.gravaIva;
}

function limpiarFormProducto() {
  const form = document.getElementById("form-producto");
  form.reset();
  delete form.dataset.editId;
}

// ========================= FACTURACIÓN =========================
function initFacturacion() {
  // Fecha actual
  document.getElementById("fac-fecha").value = hoyIso();

  // Cliente en combo
  renderClientesEnFactura();

  // Botones
  document
    .getElementById("btn-agregar-concepto")
    .addEventListener("click", agregarConcepto);

  document
    .getElementById("btn-generar-xml")
    .addEventListener("click", generarXml);

  document
    .getElementById("btn-ver-impresion")
    .addEventListener("click", verImpresion);

  // iniciar con una fila de conceptos
  agregarConcepto();
}

function renderClientesEnFactura() {
  const sel = document.getElementById("fac-cliente");
  sel.innerHTML = "";
  if (clientes.length === 0) {
    const opt = document.createElement("option");
    opt.textContent = "No hay clientes, crea uno primero";
    opt.value = "";
    sel.appendChild(opt);
    return;
  }
  clientes.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = `${c.nombre} (${c.rfc})`;
    sel.appendChild(opt);
  });
}

function agregarConcepto() {
  const tbody = document.getElementById("fac-conceptos-body");
  const id = uid();
  const tr = document.createElement("tr");
  tr.dataset.rowId = id;

  const opcionesProductos = productos
    .map((p) => `<option value="${p.id}">${escapeXml(p.descripcion)}</option>`)
    .join("");

  tr.innerHTML = `
    <td>
      <select class="fac-prod-select">
        <option value="">Selecciona...</option>
        ${opcionesProductos}
      </select>
    </td>
    <td><input type="text" class="fac-descripcion" /></td>
    <td><input type="number" min="0" step="0.0001" class="fac-cantidad" value="1" /></td>
    <td><input type="text" class="fac-clave-prodserv" /></td>
    <td><input type="text" class="fac-clave-unidad" /></td>
    <td><input type="text" class="fac-unidad" /></td>
    <td><input type="number" min="0" step="0.000001" class="fac-precio" /></td>
    <td><input type="text" class="fac-importe" readonly /></td>
    <td><input type="text" class="fac-iva" readonly /></td>
    <td class="actions"><button type="button" class="fac-row-del">X</button></td>
  `;

  tbody.appendChild(tr);

  const selProd = tr.querySelector(".fac-prod-select");
  const inpDesc = tr.querySelector(".fac-descripcion");
  const inpCant = tr.querySelector(".fac-cantidad");
  const inpClaveProd = tr.querySelector(".fac-clave-prodserv");
  const inpClaveUni = tr.querySelector(".fac-clave-unidad");
  const inpUni = tr.querySelector(".fac-unidad");
  const inpPrecio = tr.querySelector(".fac-precio");
  const inpImporte = tr.querySelector(".fac-importe");
  const inpIva = tr.querySelector(".fac-iva");
  const btnDel = tr.querySelector(".fac-row-del");

  selProd.addEventListener("change", () => {
    const p = productos.find((x) => x.id === selProd.value);
    if (!p) return;
    inpDesc.value = p.descripcion;
    inpClaveProd.value = p.claveProdServ;
    inpClaveUni.value = p.claveUnidad;
    inpUni.value = p.unidad;
    inpPrecio.value = p.precio;
    calcularImportes();
  });

  [inpCant, inpPrecio].forEach((inp) => {
    inp.addEventListener("input", () => calcularImportes());
  });

  btnDel.addEventListener("click", () => {
    tr.remove();
    calcularImportes();
  });

  function calcularImportes() {
    const cant = parseFloat(inpCant.value || "0");
    const precio = parseFloat(inpPrecio.value || "0");
    const base = cant * precio;
    const p = productos.find((x) => x.id === selProd.value);
    const gravaIva = p ? p.gravaIva : true;
    const iva = gravaIva ? base * 0.16 : 0;
    inpImporte.value = formatoMoneda(base);
    inpIva.value = formatoMoneda(iva);
    recalcularTotales();
  }
}

function leerConceptosFactura() {
  const rows = Array.from(document.querySelectorAll("#fac-conceptos-body tr"));
  const conceptos = [];

  rows.forEach((tr) => {
    const cant = parseFloat(tr.querySelector(".fac-cantidad").value || "0");
    const precio = parseFloat(tr.querySelector(".fac-precio").value || "0");
    if (!cant || !precio) return;
    const base = cant * precio;
    const iva = parseFloat(tr.querySelector(".fac-iva").value || "0");

    conceptos.push({
      descripcion: tr.querySelector(".fac-descripcion").value.trim(),
      cantidad: cant,
      claveProdServ: tr.querySelector(".fac-clave-prodserv").value.trim(),
      claveUnidad: tr.querySelector(".fac-clave-unidad").value.trim(),
      unidad: tr.querySelector(".fac-unidad").value.trim(),
      valorUnitario: precio,
      importe: base,
      iva
    });
  });

  return conceptos;
}

function recalcularTotales() {
  conceptosFactura = leerConceptosFactura();
  let subtotal = 0;
  let ivaTotal = 0;
  conceptosFactura.forEach((c) => {
    subtotal += c.importe;
    ivaTotal += c.iva;
  });
  document.getElementById("fac-subtotal").textContent = formatoMoneda(subtotal);
  document.getElementById("fac-iva").textContent = formatoMoneda(ivaTotal);
  document.getElementById("fac-total").textContent = formatoMoneda(subtotal + ivaTotal);
}

// ========================= GENERAR XML CFDI =========================
function generarXml() {
  conceptosFactura = leerConceptosFactura();
  if (conceptosFactura.length === 0) {
    alert("Agrega al menos un concepto.");
    return;
  }

  const clienteId = document.getElementById("fac-cliente").value;
  const cliente = clientes.find((c) => c.id === clienteId);
  if (!cliente) {
    alert("Selecciona un cliente.");
    return;
  }

  const serie = document.getElementById("fac-serie").value.trim() || "";
  const folio = document.getElementById("fac-folio").value.trim() || "";
  const formaPago = document.getElementById("fac-forma-pago").value;
  const metodoPago = document.getElementById("fac-metodo-pago").value;
  const moneda = document.getElementById("fac-moneda").value.trim() || "MXN";
  const lugarExp = document.getElementById("fac-lugar-exp").value.trim() || EMISOR.cp;
  const tipoComp = document.getElementById("fac-tipo-comp").value.trim() || "I";
  const fecha = hoyIso();
  document.getElementById("fac-fecha").value = fecha;

  let subtotal = 0;
  let ivaTotal = 0;
  conceptosFactura.forEach((c) => {
    subtotal += c.importe;
    ivaTotal += c.iva;
  });
  const total = subtotal + ivaTotal;

  const conceptosXml = conceptosFactura
    .map((c) => {
      const base = c.importe;
      const importeIva = c.iva;
      return `
  <cfdi:Concepto ObjetoImp="02"
    ClaveProdServ="${escapeXml(c.claveProdServ || "01010101")}"
    Cantidad="${c.cantidad.toFixed(4)}"
    ClaveUnidad="${escapeXml(c.claveUnidad || "E48")}"
    Unidad="${escapeXml(c.unidad || "SERV")}"
    Descripcion="${escapeXml(c.descripcion)}"
    ValorUnitario="${c.valorUnitario.toFixed(6)}"
    Importe="${base.toFixed(2)}">
    <cfdi:Impuestos>
      <cfdi:Traslados>
        <cfdi:Traslado Base="${base.toFixed(2)}"
          Impuesto="002"
          TipoFactor="Tasa"
          TasaOCuota="0.160000"
          Importe="${importeIva.toFixed(2)}" />
      </cfdi:Traslados>
    </cfdi:Impuestos>
  </cfdi:Concepto>`;
    })
    .join("");

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante
  Version="4.0"
  Serie="${escapeXml(serie)}"
  Folio="${escapeXml(folio)}"
  Fecha="${fecha}"
  Moneda="${escapeXml(moneda)}"
  Exportacion="01"
  TipoDeComprobante="${escapeXml(tipoComp)}"
  FormaPago="${escapeXml(formaPago)}"
  MetodoPago="${escapeXml(metodoPago)}"
  LugarExpedicion="${escapeXml(lugarExp)}"
  SubTotal="${subtotal.toFixed(2)}"
  Total="${total.toFixed(2)}"
  xmlns:cfdi="http://www.sat.gob.mx/cfd/4"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="http://www.sat.gob.mx/cfd/4 http://www.sat.gob.mx/sitio_internet/cfd/4/cfdv40.xsd">
  <cfdi:Emisor
    Rfc="${escapeXml(EMISOR.rfc)}"
    Nombre="${escapeXml(EMISOR.nombre)}"
    RegimenFiscal="${escapeXml(EMISOR.regimenFiscal)}" />
  <cfdi:Receptor
    Rfc="${escapeXml(cliente.rfc)}"
    Nombre="${escapeXml(cliente.nombre)}"
    DomicilioFiscalReceptor="${escapeXml(cliente.cp)}"
    RegimenFiscalReceptor="${escapeXml(cliente.regimen)}"
    UsoCFDI="${escapeXml(cliente.usoCfdi)}" />
  <cfdi:Conceptos>${conceptosXml}
  </cfdi:Conceptos>
  <cfdi:Impuestos TotalImpuestosTrasladados="${ivaTotal.toFixed(2)}">
    <cfdi:Traslados>
      <cfdi:Traslado Base="${subtotal.toFixed(2)}"
        Impuesto="002"
        TipoFactor="Tasa"
        TasaOCuota="0.160000"
        Importe="${ivaTotal.toFixed(2)}" />
    </cfdi:Traslados>
  </cfdi:Impuestos>
  <!-- Complemento de timbrado se agrega cuando el PAC timbra este CFDI -->
</cfdi:Comprobante>`;

  document.getElementById("xml-output").value = xml;
  actualizarVistaImpresion(cliente, { serie, folio, fecha, formaPago, metodoPago, subtotal, ivaTotal, total });
}

// ========================= VISTA IMPRESIÓN =========================
function actualizarVistaImpresion(cliente, datosFactura) {
  const preview = document.getElementById("preview-impresion");
  const conceptosHtml = conceptosFactura
    .map((c, i) => {
      return `
      <tr>
        <td>${i + 1}</td>
        <td>${escapeXml(c.descripcion)}</td>
        <td>${c.cantidad.toFixed(2)}</td>
        <td>${escapeXml(c.unidad || "")}</td>
        <td>${formatoMoneda(c.valorUnitario)}</td>
        <td>${formatoMoneda(c.importe)}</td>
      </tr>`;
    })
    .join("");

  const html = `
  <div class="factura-pdf">
    <div class="factura-pdf-header">
      <div class="emisor">
        <h2>Datos del cliente</h2>
        <div class="factura-pdf-datos">
          <strong>${escapeXml(EMISOR.nombre)}</strong><br/>
          RFC: ${escapeXml(EMISOR.rfc)}<br/>
          ${escapeXml(EMISOR.direccion)}<br/>
          Régimen fiscal: ${escapeXml(EMISOR.regimenFiscal)}<br/>
          Lugar de expedición: ${escapeXml(EMISOR.cp)}
        </div>
      </div>
      <div class="receptor">
        <h2>Receptor</h2>
        <div class="factura-pdf-datos">
          <strong>${escapeXml(cliente.nombre)}</strong><br/>
          RFC: ${escapeXml(cliente.rfc)}<br/>
          CP: ${escapeXml(cliente.cp)}<br/>
          Uso CFDI: ${escapeXml(cliente.usoCfdi)}<br/>
          Régimen fiscal: ${escapeXml(cliente.regimen)}<br/>
          Dirección: ${escapeXml(cliente.direccion)}
        </div>
      </div>
    </div>

    <div class="factura-pdf-datos" style="margin-top:6px;">
      Comprobante fiscal digital - Serie: ${escapeXml(datosFactura.serie)} Folio: ${escapeXml(datosFactura.folio)}<br/>
      Fecha y hora: ${escapeXml(datosFactura.fecha)}<br/>
      Forma de pago: ${escapeXml(datosFactura.formaPago)} -
      Método de pago: ${escapeXml(datosFactura.metodoPago)}<br/>
      Moneda: MXN - Tipo de comprobante: I
    </div>

    <div class="factura-pdf-table">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Descripción</th>
            <th>Cantidad</th>
            <th>Unidad</th>
            <th>P/U</th>
            <th>Importe</th>
          </tr>
        </thead>
        <tbody>
          ${conceptosHtml}
        </tbody>
      </table>
    </div>

    <div class="factura-pdf-totales">
      Subtotal: ${formatoMoneda(datosFactura.subtotal)}<br/>
      IVA 16%: ${formatoMoneda(datosFactura.ivaTotal)}<br/>
      <strong>Total: ${formatoMoneda(datosFactura.total)}</strong>
    </div>

    <div class="factura-pdf-sellos">
      <strong>Sellos y cadena original</strong><br/>
      (Aquí puedes imprimir el sello del CFDI, sello SAT y cadena original una vez timbrado)
    </div>
  </div>`;

  preview.innerHTML = html;

  // también para el área de impresión
  document.getElementById("print-area").innerHTML = html;
}

function verImpresion() {
  const clienteId = document.getElementById("fac-cliente").value;
  const cliente = clientes.find((c) => c.id === clienteId);
  if (!cliente) {
    alert("Selecciona un cliente antes de imprimir.");
    return;
  }
  conceptosFactura = leerConceptosFactura();
  let subtotal = 0;
  let ivaTotal = 0;
  conceptosFactura.forEach((c) => {
    subtotal += c.importe;
    ivaTotal += c.iva;
  });
  const total = subtotal + ivaTotal;
  const datosFactura = {
    serie: document.getElementById("fac-serie").value.trim(),
    folio: document.getElementById("fac-folio").value.trim(),
    fecha: document.getElementById("fac-fecha").value.trim() || hoyIso(),
    formaPago: document.getElementById("fac-forma-pago").value,
    metodoPago: document.getElementById("fac-metodo-pago").value,
    subtotal,
    ivaTotal,
    total
  };
  actualizarVistaImpresion(cliente, datosFactura);
  window.print();
}
