/* ================================================================
   APP.JS — PARTE 1
   VARIABLES + CARGA DE ARCHIVOS + LECTURA EXCEL + AGRUPACIÓN
================================================================ */

/* ===========================
   ESTADO GLOBAL
=========================== */
let datosGlobales = [];

/* ===========================
   LISTAS DE VEHÍCULOS ESPERADOS
=========================== */
const vehiculosEsperados = {
  703: "700", 705: "700", 707: "700", 714: "700", 715: "700",
  717: "700", 719: "700", 721: "700", 723: "700", 725: "700",
  728: "700", 729: "700", 730: "700", 731: "700", 733: "700",
  735: "700", 742: "700", 744: "700", 746: "700", 747: "700",
  748: "700", 749: "700", 752: "700", 758: "700", 759: "700",
  708: "700", 709: "700", 757: "700", 736: "700", 722: "700",
  716: "700", 718: "700", 720: "700", 750: "700", 726: "700",
  727: "700", 732: "700", 734: "700", 741: "700", 724: "700",
  737: "700", 740: "700", 755: "700", 753: "700", 754: "700",
  751: "700", 739: "700", 743: "700", 745: "700", 738: "700",
  756: "700",

  // Especiales
  764: "ESPECIALES", 767: "ESPECIALES", 769: "ESPECIALES",
  766: "ESPECIALES", 768: "ESPECIALES"
};

/* ===========================
   MAPA DE BASES
=========================== */
const mapaBase = {
  703:4,705:4,707:4,708:5,709:3,714:3,715:4,716:3,717:4,718:3,
  719:2,720:3,721:4,722:3,723:3,724:3,725:4,726:3,727:3,728:4,
  729:1,730:1,731:4,732:1,733:5,734:3,735:4,736:8,737:3,738:3,
  739:3,740:3,741:3,742:3,744:3,745:3,746:4,747:5,748:2,749:2,
  750:3,751:3,752:3,753:3,754:3,755:3,757:5,758:3,759:6,756:0,
  764:0,767:0,769:0,766:0,768:0
};

/* ======================================================
   ACTUALIZAR DISPLAY DE FECHA
====================================================== */
document.getElementById("fechaReporte").addEventListener("change", function () {
  const fecha = this.value;
  const display = document.getElementById("fechaDisplay");
  if (!fecha) return display.textContent = "Seleccione una fecha";
  const [y,m,d] = fecha.split("-");
  display.textContent = `${d}/${m}/${y}`;
});

/* ======================================================
   MOSTRAR ESTADO DE ARCHIVOS
====================================================== */
function estadoArchivo(inputId, statusId) {
  const input = document.getElementById(inputId);
  const status = document.getElementById(statusId);
  input.addEventListener("change", () => {
    if (input.files.length > 0) {
      status.textContent = "✓ " + input.files[0].name;
      status.className = "file-status success";
    } else {
      status.textContent = "";
    }
  });
}

estadoArchivo("fileDetalle","statusDetalle");
estadoArchivo("fileMetro","statusMetro");

/* ======================================================
   LECTURA DE ARCHIVO EXCEL
====================================================== */
function leerExcel(archivo) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: "" });
        resolve(rows);
      } catch (error) { reject(error); }
    };
    reader.onerror = () => reject("Error al leer el archivo");
    reader.readAsArrayBuffer(archivo);
  });
}

/* ======================================================
   FUNCIÓN PARA CREAR REGISTROS COMPLETOS (incluye faltantes)
====================================================== */
function crearDatosCompletos(grupos, fechaClave, d, m, y, mapaMetro) {
  const listaFinal = [];
  const encontrados = Object.keys(grupos).map(Number);

  const orden = Object.keys(vehiculosEsperados)
    .map(Number)
    .sort((a, b) => a - b);

  orden.forEach(veh => {
    const ruta = vehiculosEsperados[veh];
    const presente = encontrados.includes(veh);

    if (presente) {
      listaFinal.push({
        ...grupos[veh],
        _idx: listaFinal.length,
        _ingresosBase: grupos[veh]._ingresosBase,
        _salidasP1Base: grupos[veh]._salidasP1Base,
        "Validación": "PRESENTE",
        "Ruta Validación": ruta
      });

    } else {
      listaFinal.push({
        "Vehiculo": veh,
        "Ingresos totales": 0,
        "Salidas P1": 0,
        "Salidas P2": 0,
        "Total salidas": 0,
        "Dif": 0,
        "Ruta": ruta,
        "Fecha": fechaClave,
        "Mes": obtenerMes(m),
        "Año": y,
        "# Día": obtenerDiaSemana(d,m,y),
        "Base": mapaBase[veh] || "SIN BASE",
        "Ajustes": 0,
        "Pasajeros metro": mapaMetro[veh] || 0,
        "Descuento>20": 0,
        "No labora": "FALTANTE",
        "Observaciones": "Vehículo no encontrado en archivo",
        "Sonar": 0,
        "Viajes realizados": 0,
        "SUB-RUTA": "",
        "Observaciones pasajero 20": "Vehículo faltante",
        "_ingresosBase": 0,
        "_salidasP1Base": 0,
        "_idx": listaFinal.length,
        "Validación": "FALTANTE",
        "Ruta Validación": ruta
      });
    }
  });

  return listaFinal;
}

/* ======================================================
   PROCESAR REPORTE (inicio)
====================================================== */
document.getElementById("btnProcesar").addEventListener("click", async () => {

  const archivoViajes = document.getElementById("fileDetalle").files[0];
  const archivoMetro  = document.getElementById("fileMetro").files[0];
  const fechaSel = document.getElementById("fechaReporte").value;

  if (!fechaSel) return alert("Seleccione la fecha");
  if (!archivoViajes) return alert("Cargue archivo de viajes");
  if (!archivoMetro)  return alert("Cargue archivo Metro");

  // Mostrar loading
  document.getElementById("loading").style.display = "block";

  const [Y, Mraw, Draw] = fechaSel.split("-");
  const m = Number(Mraw) - 1;
  const d = Number(Draw);
  const y = Number(Y);
  const fechaClave = `${Draw}/${Mraw}/${Y}`;

  try {
    // Leer archivos
    const viajes = await leerExcel(archivoViajes);
    const metro  = await leerExcel(archivoMetro);

    // Mapa Metro por vehículo
    const mapaMetro = {};
    metro.forEach(x => {
      const v = Number(x["Vehiculos"]);
      if (v) mapaMetro[v] = Number(x["Total general"] || 0);
    });

    /* ======== AGRUPAR POR VEHÍCULO ======== */
    const grupos = {};

    viajes.forEach(r => {
      const veh = r["Vehiculo"] || r["Vehículo descripción"] || r["Vehiculo descripción"];
      if (!veh) return;

      const ingresos = Number(r["Ingresos totales"] || 0);
      const salP1    = Number(r["Salidas P1"] || 0);
      const salP2    = Number(r["Salidas P2"] || r[" Salidas P2"] || 0);
      const salTot   = Number(r["Salidas totales"] || 0);
      const rutaNom  = String(r["Nombre de ruta"] || "").toLowerCase();

      const esSinRuta = rutaNom.includes("sin ruta");
      const excedente = (esSinRuta || ingresos <= 20) ? 0 : ingresos - 20;

      const ingDesc = ingresos - excedente;
      const p1Desc  = salP1 - excedente;
      const totDesc = salTot - excedente;

      if (!grupos[veh]) {
        grupos[veh] = {
          "Vehiculo": veh,
          "Ingresos totales": ingDesc,
          "Salidas P1": p1Desc,
          "Salidas P2": salP2,
          "Total salidas": totDesc,
          "Dif": ingDesc - totDesc,
          "Ruta": "700",
          "Fecha": fechaClave,
          "Mes": obtenerMes(m),
          "Año": y,
          "# Día": obtenerDiaSemana(d,m,y),
          "Base": mapaBase[veh] || "SIN BASE",
          "Ajustes": 0,
          "Pasajeros metro": mapaMetro[veh] || 0,
          "Descuento>20": excedente,
          "No labora": "",
          "Observaciones": "",
          "Sonar": ingresos,
          "Viajes realizados": esSinRuta ? 0 : 1,
          "SUB-RUTA": "",
          "Observaciones pasajero 20": "",
          "_ingresosBase": ingDesc,
          "_salidasP1Base": p1Desc
        };
      } else {
        grupos[veh]["Ingresos totales"] += ingDesc;
        grupos[veh]["Salidas P1"] += p1Desc;
        grupos[veh]["Salidas P2"] += salP2;
        grupos[veh]["Total salidas"] += totDesc;

        if (!esSinRuta) grupos[veh]["Viajes realizados"] += 1;

        grupos[veh]["Descuento>20"] += excedente;
        grupos[veh]["Sonar"] += ingresos;
        grupos[veh]["_ingresosBase"] += ingDesc;
        grupos[veh]["_salidasP1Base"] += p1Desc;
      }
    });

    // Crear registros completos (incluye faltantes)
    datosGlobales = crearDatosCompletos(grupos, fechaClave, d, m, y, mapaMetro);

    // Mostrar tablas
    mostrarTablasEnPestanas(datosGlobales);

    document.getElementById("btnExportar").disabled = false;

  } catch (err) {
    console.error(err);
    alert("Error procesando archivos");
  }

  document.getElementById("loading").style.display = "none";
});


/* ======================================================
   FUNCIONES AUXILIARES
====================================================== */
function obtenerMes(m) {
  return [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"
  ][m];
}

function obtenerDiaSemana(d,m,y) {
  return ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"]
    [ new Date(y,m,d).getDay() ];
}
/* ================================================================
   APP.JS — PARTE 2
   RENDERIZADO DE TABLAS + PESTAÑAS + CELDAS EDITABLES
================================================================ */

/* ======================================================
   MOSTRAR TABLAS EN LAS 3 PESTAÑAS
====================================================== */
function mostrarTablasEnPestanas(data) {
  const ruta700 = data.filter(r => r["Ruta Validación"] === "700");
  const especiales = data.filter(r => r["Ruta Validación"] === "ESPECIALES");

  mostrarTabla(ruta700, "tablaRuta700");
  mostrarTabla(especiales, "tablaEspeciales");
  mostrarTabla(data, "tablaTodos");

  // Contadores
  document.querySelector('[data-tab="ruta700"]').textContent = `Ruta 700 (${ruta700.length})`;
  document.querySelector('[data-tab="especiales"]').textContent = `Especiales (${especiales.length})`;
  document.querySelector('[data-tab="todos"]').textContent = `Todos (${data.length})`;

  document.getElementById("tabsContainer").style.display = "block";
}

/* ======================================================
   FUNCIÓN PRINCIPAL: RENDERIZAR UNA TABLA
====================================================== */
function mostrarTabla(lista, divID) {
  const cont = document.getElementById(divID);
  if (!lista.length) return cont.innerHTML = "<p>No hay datos.</p>";

  // Todas las columnas visibles (excluye las internas)
  const columnas = Object.keys(lista[0]).filter(c => !c.startsWith("_"));

  let html = `<table><thead><tr>`;
  columnas.forEach(col => html += `<th>${col}</th>`);
  html += `</tr></thead><tbody>`;

  lista.forEach(fila => {
    const esFaltante = fila["Validación"] === "FALTANTE";
    const claseFila = esFaltante ? "vehiculo-faltante" : "";
    const idx = fila._idx;

    html += `<tr class="${claseFila}">`;

    columnas.forEach(col => {
      const valor = fila[col];

      /* ==========================================
         VALIDACIÓN (Presente/Faltante)
      ========================================== */
      if (col === "Validación") {
        const clase = valor === "FALTANTE" ? "alerta-faltante" : "alerta-presente";
        html += `<td class="${clase}">${valor}</td>`;
      }

      /* ==========================================
         AJUSTES — SOLO EDITABLE SI PRESENTE
      ========================================== */
      else if (col === "Ajustes" && !esFaltante) {
        html += `
          <td class="editable">
            <input type="number"
                   class="editable-input ajuste-input"
                   data-idx="${idx}"
                   value="${valor}">
          </td>`;
      }

      /* ==========================================
         INGRESOS TOTALES / SALIDAS P1 — EDITABLES
      ========================================== */
      else if ((col === "Ingresos totales" || col === "Salidas P1") && !esFaltante) {
        html += `
          <td class="editable">
            <input type="number"
                   class="editable-input editable-numero"
                   data-col="${col}"
                   data-idx="${idx}"
                   value="${valor}">
          </td>`;
      }

      /* ==========================================
         CAMPOS EDITABLES: TEXTO
      ========================================== */
      else if (
        (col === "Observaciones" || col === "Observaciones pasajero 20" || col === "Viajes realizados")
        && !esFaltante
      ) {
        html += `
          <td class="editable">
            <input type="text"
                   class="editable-input editable-texto"
                   data-col="${col}"
                   data-idx="${idx}"
                   value="${valor}">
          </td>`;
      }

      /* ==========================================
         CAMPOS BLOQUEADOS — SOLO VISUALIZACIÓN
      ========================================== */
      else {
        html += `<td class="bloqueado">${valor}</td>`;
      }
    });

    html += `</tr>`;
  });

  html += `</tbody></table>`;
  cont.innerHTML = html;
}

/* ======================================================
   MANEJO DE PESTAÑAS
====================================================== */
document.addEventListener("click", function (e) {
  if (!e.target.classList.contains("tab")) return;

  const tabID = e.target.dataset.tab;

  // Quitar active de todas
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));

  // Activar la seleccionada
  e.target.classList.add("active");
  document.getElementById(`tab-${tabID}`).classList.add("active");
});
/* ================================================================
   APP.JS — PARTE 3
   AJUSTES DINÁMICOS + RECÁLCULO + EXPORTAR EXCEL
================================================================ */

/* ======================================================
   ESCUCHAR CAMBIOS EN TODAS LAS CELDAS EDITABLES
====================================================== */

document.addEventListener("input", function (e) {

  /* ===============================
     AJUSTES (SUMA / RESTA)
  =============================== */
  if (e.target.classList.contains("ajuste-input")) {

    const idx = Number(e.target.dataset.idx);
    const ajuste = Number(e.target.value) || 0;
    aplicarAjuste(idx, ajuste);
  }

  /* ===============================
     CAMPOS NUMÉRICOS EDITABLES
     - Ingresos totales
     - Salidas P1
     - Viajes realizados (si es número)
  =============================== */
  if (e.target.classList.contains("editable-numero")) {

    const idx = Number(e.target.dataset.idx);
    const col = e.target.dataset.col;
    const valor = Number(e.target.value) || 0;

    datosGlobales[idx][col] = valor;

    recalcularFila(idx);
  }

  /* ===============================
     CAMPOS DE TEXTO
     - Observaciones
     - Observaciones pasajero 20
  =============================== */
  if (e.target.classList.contains("editable-texto")) {

    const idx = Number(e.target.dataset.idx);
    const col = e.target.dataset.col;
    const texto = e.target.value;

    datosGlobales[idx][col] = texto;
  }
});


/* ======================================================
   APLICAR AJUSTE
====================================================== */
function aplicarAjuste(idx, ajuste) {

  const fila = datosGlobales[idx];
  if (!fila) return;

  fila["Ajustes"] = ajuste;

  // 🔵 Reset a los valores base (NO acumulado)
  const baseIng = Number(fila._ingresosBase || 0);
  const baseP1 = Number(fila._salidasP1Base || 0);

  fila["Ingresos totales"] = baseIng + ajuste;
  fila["Salidas P1"] = baseP1 + ajuste;

  // Recalcular DIF
  recalcularFila(idx);
}


/* ======================================================
   RECALCULAR DIF Y REDIBUJAR TABLES
====================================================== */
function recalcularFila(idx) {

  const fila = datosGlobales[idx];
  if (!fila) return;

  const ingresos = Number(fila["Ingresos totales"] || 0);
  const totalSalidas = Number(fila["Total salidas"] || 0);

  fila["Dif"] = ingresos - totalSalidas;

  // 🔵 Re-render COMPLETO de las pestañas
  mostrarTablasEnPestanas(datosGlobales);
}


/* ======================================================
   EXPORTAR EXCEL (SIN COLUMNAS INTERNAS)
====================================================== */
document.getElementById("btnExportar").addEventListener("click", () => {
  exportarExcel(datosGlobales);
});

function exportarExcel(data) {

  const fecha = document.getElementById("fechaReporte").value || "sin_fecha";
  const nombreArchivo = `Resumen_vehiculos_${fecha}.xlsx`;

  // Eliminar columnas internas _ingresosBase, _idx, etc.
  const limpio = data.map(fila => {
    const obj = {};
    Object.keys(fila).forEach(k => {
      if (!k.startsWith("_")) obj[k] = fila[k];
    });
    return obj;
  });

  const ws = XLSX.utils.json_to_sheet(limpio);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Resumen");

  XLSX.writeFile(wb, nombreArchivo);
}

// =============================
// CERRAR SESIÓN
// =============================
document.getElementById("logoutBtn").addEventListener("click", () => {
  firebase.auth().signOut();
  localStorage.removeItem("auth");
  window.location.href = "login.html";
});
