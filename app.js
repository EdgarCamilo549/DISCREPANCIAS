let datosActuales = [];

const FLOW_URL =
"https://default626c9bafa3094c4c8fe682c68a1ac9.88.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7eecdf32d22f4a3b8207199247753382/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=mygVtwPLe-haRd_34SKqDBspPnuSOCcJc3IZG6jS63M";

const MODO_PRUEBA = true;

/* =========================================
   CARGAR EXCEL
========================================= */
window.addEventListener(
  "DOMContentLoaded",
  cargarExcelAutomatico
);

async function cargarExcelAutomatico() {

  try {

    const response =
      await fetch(
        "Discrepancias_Consolidado.xlsx"
      );

    const arrayBuffer =
      await response.arrayBuffer();

    const workbook =
      XLSX.read(
        arrayBuffer,
        {
          type: "array",
          raw: false
        }
      );

    const hoja =
      workbook.Sheets[
        workbook.SheetNames[0]
      ];

    datosActuales =
      XLSX.utils.sheet_to_json(
        hoja,
        { raw: false }
      );

    if (!datosActuales.length) {

      alert(
        "El archivo está vacío."
      );

      return;
    }

    procesarDatos();

    crearFiltroTiendas();

    activarFiltroFecha();

    activarFiltroTarima();

    document
      .getElementById(
        "mensajeCarga"
      ).innerText =
      "Seleccione una tienda.";

    document
      .getElementById(
        "filtros"
      ).style.display =
      "flex";

  } catch (error) {

    console.error(error);

    alert(
      "No se pudo cargar el Excel."
    );
  }
}

/* =========================================
   PROCESAR DATOS
========================================= */
function procesarDatos() {

  datosActuales.forEach(
    fila => {

      fila.RDW_REF_1_ORIGINAL =
        String(
          fila.RDW_REF_1_ORIGINAL
        )
        .trim()
        .slice(1);

      fila.DAY =
        convertirFechaExcel(
          fila.DAY
        );

      fila.CD =
        calcularCD(
          fila.RDW_REF_1_ORIGINAL
        );

      fila.Tarima =
        calcularTarima(
          fila.CD,
          fila.RDW_REF_1_ORIGINAL
        );

      fila.Tienda =
        calcularTienda(
          fila.LOC_IDNT,
          fila.LOC_DESC
        );

      fila.EstatusUnidades =
        calcularEstatusUnidades(
          fila.UNIDADES
        );
    }
  );
}

/* =========================================
   FECHA
========================================= */
function convertirFechaExcel(
  valor
) {

  if (!valor)
    return "";

  if (isNaN(valor)) {

    const texto =
      String(valor).trim();

    if (
      texto.includes("/")
    ) {

      const partes =
        texto.split("/");

      if (
        partes.length === 3
      ) {

        let mes =
          partes[0]
            .padStart(2, "0");

        let dia =
          partes[1]
            .padStart(2, "0");

        let año =
          partes[2];

        if (
          año.length === 2
        ) {
          año =
            "20" + año;
        }

        return `${año}-${mes}-${dia}`;
      }
    }

    return texto;
  }

  const fecha =
    new Date(
      (Number(valor) - 25569)
      * 86400
      * 1000
    );

  return fecha
    .toISOString()
    .split("T")[0];
}

/* =========================================
   FORMULA CD
========================================= */
function calcularCD(
  valorTexto
) {

  if (!valorTexto)
    return "AJUSTE";

  valorTexto =
    String(
      valorTexto
    ).trim();

  const esNumerico =
    !isNaN(
      Number(valorTexto)
    );

  const valorNumerico =
    esNumerico
      ? Number(valorTexto)
      : null;

  const ultimos6 =
    esNumerico &&
    valorTexto.length >= 6
      ? Number(
          valorTexto.slice(-6)
        )
      : null;

  if (
    valorTexto.startsWith(
      "1000000"
    )
  )
    return "7021";

  if (
    ultimos6 >= 145000 &&
    ultimos6 <= 150000
  )
    return "7650";

  if (
    ultimos6 > 150000 &&
    ultimos6 <= 180000
  )
    return "7350";

  if (
    valorNumerico >= 7000 &&
    valorNumerico <= 55000
  )
    return "7430";

  if (
    ultimos6 >= 600000 &&
    ultimos6 <= 800000
  )
    return "7440";

  if (
    valorTexto
      .toUpperCase()
      .startsWith("RO")
  )
    return "AJUSTE RO";

  if (
    valorTexto
      .toUpperCase()
      .endsWith("/A")
  )
    return "AJUSTE/A";

  return "AJUSTE";
}

/* =========================================
   TARIMA
========================================= */
function calcularTarima(
  cd,
  rdwRef
) {

  if (
    String(cd)
      .includes("AJUSTE")
  ) {
    return "";
  }

  return (
    "PLT" +
    cd +
    String(
      rdwRef
    ).trim()
  );
}

/* =========================================
   TIENDA
========================================= */
function calcularTienda(
  no,
  desc
) {

  return (
    String(no) +
    " - " +
    String(desc)
  );
}

/* =========================================
   ESTATUS
========================================= */
function calcularEstatusUnidades(
  unidades
) {

  return Number(
    unidades
  ) > 0
    ? "Sobrantes"
    : "Faltantes";
}

/* =========================================
   FILTROS
========================================= */
function crearFiltroTiendas() {

  const lista =
    document.getElementById(
      "listaTiendas"
    );

  const tiendas =
    [...new Set(
      datosActuales.map(
        x => x.Tienda
      )
    )];

  tiendas.sort();

  lista.innerHTML = "";

  tiendas.forEach(
    tienda => {

      lista.innerHTML += `
        <option value="${tienda}">
      `;
    }
  );

  document
    .getElementById(
      "filtroTienda"
    )
    .addEventListener(
      "change",
      () => {

        crearFiltroTarimas();

        aplicarFiltros();
      }
    );
}

function crearFiltroTarimas() {

  const tienda =
    document
      .getElementById(
        "filtroTienda"
      )
      .value
      .trim();

  const lista =
    document.getElementById(
      "listaTarimas"
    );

  lista.innerHTML = "";

  if (!tienda)
    return;

  const tarimas =
    [...new Set(
      datosActuales
        .filter(
          x =>
            x.Tienda === tienda
        )
        .map(
          x => x.Tarima
        )
        .filter(Boolean)
    )];

  tarimas.sort();

  tarimas.forEach(
    tarima => {

      lista.innerHTML += `
        <option value="${tarima}">
      `;
    }
  );
}

function activarFiltroFecha() {

  document
    .getElementById(
      "filtroFecha"
    )
    .addEventListener(
      "change",
      aplicarFiltros
    );
}

function activarFiltroTarima() {

  document
    .getElementById(
      "filtroTarima"
    )
    .addEventListener(
      "change",
      aplicarFiltros
    );
}

/* =========================================
   FILTRAR
========================================= */
function aplicarFiltros() {

  const tienda =
    document
      .getElementById(
        "filtroTienda"
      )
      .value
      .trim();

  const fecha =
    document
      .getElementById(
        "filtroFecha"
      )
      .value;

  const tarima =
    document
      .getElementById(
        "filtroTarima"
      )
      .value
      .trim();

  if (!tienda) {

    mostrarTabla([]);

    return;
  }

  const filtrados =
    datosActuales.filter(
      fila =>

        fila.Tienda === tienda &&

        (!fecha ||
          fila.DAY.startsWith(
            fecha
          )) &&

        (!tarima ||
          fila.Tarima === tarima) &&

        !fila.CD.includes(
          "AJUSTE"
        )
    );

  mostrarTabla(
    filtrados
  );
}

/* =========================================
   TABLA
========================================= */
function mostrarTabla(
  datos
) {

  const thead =
    document.querySelector(
      "#tabla thead"
    );

  const tbody =
    document.querySelector(
      "#tabla tbody"
    );

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!datos.length)
    return;

  const columnas = [
    "Reset",
    "DAY",
    "Tarima",
    "SKU_IDNT",
    "SKU_DESC",
    "EstatusUnidades",
    "UNIDADES",
    "COSTO",
    "Accion",
    "Comentarios"
  ];

  let encabezado =
    "<tr>";

  columnas.forEach(
    col => {

      encabezado += `
        <th>${col}</th>
      `;
    }
  );

  encabezado += "</tr>";

  thead.innerHTML =
    encabezado;

  let filasHTML = "";

  datos.forEach(
    fila => {

      const llave =
        fila.DAY +
        "_" +
        fila.Tarima +
        "_" +
        fila.SKU_IDNT;

      const enviados =
        JSON.parse(
          localStorage.getItem(
            "filasEnviadas"
          ) || "[]"
        );

      const enviadosInfo =
        JSON.parse(
          localStorage.getItem(
            "filasEnviadasInfo"
          ) || "{}"
        );

      const yaEnviado =
        enviados.includes(
          llave
        );

      const tipoGuardado =
        enviadosInfo[llave];

      let colorFila = "";
      let opacidad = "";

      if (yaEnviado) {

        colorFila =
          "background:#d4edda;";

        opacidad =
          "opacity:0.8;";
      }

      filasHTML += `
        <tr
          style="${colorFila}${opacidad}"
          data-enviado="${
            yaEnviado
              ? "si"
              : "no"
          }"
        >
      `;

      columnas.forEach(
        col => {

          let valor =
            fila[col] ?? "";

          let estilo = "";

          if (col === "Reset") {

            if (yaEnviado) {

              valor = `
                <input
                  type="checkbox"
                  class="check-reset"
                  data-llave="${llave}"
                >
              `;

            } else {

              valor = "";
            }
          }

          if (col === "Accion") {

            if (yaEnviado) {

              valor = `
                <strong style="color:green;">
                  ✔ ${tipoGuardado}
                </strong>
              `;

            } else {

              valor = `
                <button
                  onclick="marcarDiscrepanciaReal(this)"
                  class="btn-real"
                >
                  Discrepancia Real
                </button>

                <button
                  onclick="marcarErrorCaptura(this)"
                  class="btn-error"
                >
                  Error de Captura
                </button>
              `;
            }
          }

          if (
            col === "Comentarios"
          ) {

            valor = `
              <input
                type="text"
                class="comentario"
                placeholder="Escribir comentario..."
                style="width:220px;"
                ${
                  yaEnviado
                    ? "disabled"
                    : ""
                }
              >
            `;
          }

          if (
            col === "COSTO"
          ) {

            const numero =
              Number(valor);

            if (
              !isNaN(numero)
            ) {

              valor =
                numero.toLocaleString(
                  "es-MX",
                  {
                    style: "currency",
                    currency: "MXN"
                  }
                );

              if (
                numero < 0
              ) {

                estilo =
                  "color:red;font-weight:bold;";
              }
            }
          }

          if (
            col ===
              "EstatusUnidades" &&
            valor ===
              "Faltantes"
          ) {

            estilo =
              "color:red;font-weight:bold;";
          }

          filasHTML += `
            <td style="${estilo}">
              ${valor}
            </td>
          `;
        }
      );

      filasHTML += "</tr>";
    }
  );

  tbody.innerHTML =
    filasHTML;

  document
    .getElementById(
      "mensajeCarga"
    ).innerText =
    `Registros encontrados: ${datos.length}`;
}

/* =========================================
   BOTONES
========================================= */
function marcarDiscrepanciaReal(
  boton
) {

  const fila =
    boton.closest("tr");

  fila.style.backgroundColor =
    "#cfe2ff";

  fila.dataset.tipo =
    "Discrepancia Real";

  validarSelecciones();
}

function marcarErrorCaptura(
  boton
) {

  const fila =
    boton.closest("tr");

  fila.style.backgroundColor =
    "#fff3cd";

  fila.dataset.tipo =
    "Error de Captura";

  validarSelecciones();
}

function marcarTodasDiscrepanciaReal() {

  document
    .querySelectorAll(
      "#tabla tbody tr"
    )
    .forEach(fila => {

      if (
        fila.dataset.enviado !== "si"
      ) {

        fila.style.backgroundColor =
          "#cfe2ff";

        fila.dataset.tipo =
          "Discrepancia Real";
      }
    });

  validarSelecciones();
}

function marcarTodasErrorCaptura() {

  document
    .querySelectorAll(
      "#tabla tbody tr"
    )
    .forEach(fila => {

      if (
        fila.dataset.enviado !== "si"
      ) {

        fila.style.backgroundColor =
          "#fff3cd";

        fila.dataset.tipo =
          "Error de Captura";
      }
    });

  validarSelecciones();
}

function borrarSelecciones() {

  document
    .querySelectorAll(
      "#tabla tbody tr"
    )
    .forEach(fila => {

      if (
        fila.dataset.enviado !== "si"
      ) {

        fila.style.backgroundColor =
          "";

        delete fila.dataset.tipo;
      }
    });

  validarSelecciones();
}

/* =========================================
   COMENTARIOS
========================================= */
function aplicarComentarioMasivo() {

  const comentario =
    document.getElementById(
      "comentarioMasivo"
    ).value;

  document
    .querySelectorAll(
      ".comentario"
    )
    .forEach(
      input => {

        if (
          !input.disabled
        ) {

          input.value =
            comentario;
        }
      }
    );
}

/* =========================================
   SHAREPOINT
========================================= */
async function guardarEnSharePoint() {

  const filas =
    document.querySelectorAll(
      "#tabla tbody tr"
    );

  const usuario =
    document
      .getElementById(
        "filtroTienda"
      )
      .value
      .trim();

  let enviados = 0;

  for (const fila of filas) {

    if (
      fila.dataset.enviado ===
      "si"
    ) {
      continue;
    }

    const tipo =
      fila.dataset.tipo;

    if (!tipo)
      continue;

    const celdas =
      fila.querySelectorAll(
        "td"
      );

    const fecha =
      celdas[1]
        .innerText
        .trim();

    const tarima =
      celdas[2]
        .innerText
        .trim();

    const sku =
      celdas[3]
        .innerText
        .trim();

    const costoTexto =
      celdas[7]
        .innerText
        .trim();

    const costo =
      Number(
        costoTexto
          .replace("$", "")
          .replaceAll(",", "")
      );

    const comentario =
      fila
        .querySelector(
          ".comentario"
        )
        .value
        .trim();

    const datos = {

      usuario:
        usuario,

      tarima:
        tarima,

      fecha:
        fecha,

      sku:
        sku,

      tipo:
        tipo,

      comentarios:
        comentario,

      costo:
        costo
    };

    try {

      const respuesta =
        await fetch(
          FLOW_URL,
          {
            method: "POST",

            headers: {
              "Content-Type":
                "application/json"
            },

            body:
              JSON.stringify(
                datos
              )
          }
        );

      if (
        respuesta.ok
      ) {

        enviados++;

        const llave =
          fecha +
          "_" +
          tarima +
          "_" +
          sku;

        let enviadosStorage =
          JSON.parse(
            localStorage.getItem(
              "filasEnviadas"
            ) || "[]"
          );

        let enviadosInfo =
          JSON.parse(
            localStorage.getItem(
              "filasEnviadasInfo"
            ) || "{}"
          );

        if (
          !enviadosStorage.includes(
            llave
          )
        ) {

          enviadosStorage.push(
            llave
          );

          localStorage.setItem(
            "filasEnviadas",
            JSON.stringify(
              enviadosStorage
            )
          );
        }

        enviadosInfo[
          llave
        ] = tipo;

        localStorage.setItem(
          "filasEnviadasInfo",
          JSON.stringify(
            enviadosInfo
          )
        );

        fila.dataset.enviado =
          "si";

        fila.style.backgroundColor =
          "#d4edda";

        fila.style.opacity =
          "0.8";
      }

    } catch (error) {

      console.error(error);
    }
  }

  alert(
    `${enviados} registros guardados`
  );

  aplicarFiltros();
}

/* =========================================
   VALIDAR BOTON
========================================= */
function validarSelecciones() {

  const filas =
    document.querySelectorAll(
      "#tabla tbody tr"
    );

  let existeSeleccion =
    false;

  filas.forEach(
    fila => {

      if (
        fila.dataset.tipo
      ) {

        existeSeleccion =
          true;
      }
    }
  );

  const btnGuardar =
    document.getElementById(
      "btnGuardar"
    );

  if (!btnGuardar)
    return;

  btnGuardar.disabled =
    !existeSeleccion;

  if (existeSeleccion) {

    btnGuardar.innerHTML =
      "✔ Guardar en SharePoint";

    btnGuardar.classList.remove(
      "bloqueado"
    );

    btnGuardar.classList.add(
      "desbloqueado"
    );

  } else {

    btnGuardar.innerHTML =
      "🔒 Guardar en SharePoint";

    btnGuardar.classList.add(
      "bloqueado"
    );

    btnGuardar.classList.remove(
      "desbloqueado"
    );
  }
}

/* =========================================
   RESET ADMIN
========================================= */
function resetearFilasSeleccionadas() {

  const password =
    prompt(
      "Ingrese contraseña administrador"
    );

  const PASSWORD_ADMIN =
    "Admin2025";

  if (
    password !==
    PASSWORD_ADMIN
  ) {

    alert(
      "Contraseña incorrecta"
    );

    return;
  }

  const checks =
    document.querySelectorAll(
      ".check-reset:checked"
    );

  if (!checks.length) {

    alert(
      "Seleccione filas"
    );

    return;
  }

  let enviados =
    JSON.parse(
      localStorage.getItem(
        "filasEnviadas"
      ) || "[]"
    );

  let enviadosInfo =
    JSON.parse(
      localStorage.getItem(
        "filasEnviadasInfo"
      ) || "{}"
    );

  checks.forEach(
    check => {

      const llave =
        check.dataset.llave;

      enviados =
        enviados.filter(
          x => x !== llave
        );

      delete enviadosInfo[
        llave
      ];
    }
  );

  localStorage.setItem(
    "filasEnviadas",
    JSON.stringify(
      enviados
    )
  );

  localStorage.setItem(
    "filasEnviadasInfo",
    JSON.stringify(
      enviadosInfo
    )
  );

  alert(
    "Filas desbloqueadas"
  );

  aplicarFiltros();
}
/* =========================================
   LIMPIAR FILTROS
========================================= */
function limpiarFiltros() {

  document
    .getElementById(
      "filtroTienda"
    ).value = "";

  document
    .getElementById(
      "filtroFecha"
    ).value = "";

  document
    .getElementById(
      "filtroTarima"
    ).value = "";

  document
    .getElementById(
      "listaTarimas"
    ).innerHTML = "";

  mostrarTabla([]);

  document
    .getElementById(
      "mensajeCarga"
    ).innerText =
    "Seleccione una tienda.";

  /* BLOQUEAR BOTON SHAREPOINT */
  const btnGuardar =
    document.getElementById(
      "btnGuardar"
    );

  if (btnGuardar) {

    btnGuardar.disabled =
      true;

    btnGuardar.innerHTML =
      "🔒 Guardar en SharePoint";

    btnGuardar.classList.add(
      "bloqueado"
    );

    btnGuardar.classList.remove(
      "desbloqueado"
    );
  }
}
/* =========================================
   BORRAR COMENTARIOS MASIVOS
========================================= */
function borrarComentarioMasivo() {

  document
    .getElementById(
      "comentarioMasivo"
    ).value = "";

  document
    .querySelectorAll(
      ".comentario"
    )
    .forEach(
      input => {

        /* SOLO LIMPIAR LOS NO ENVIADOS */
        if (
          !input.disabled
        ) {

          input.value = "";
        }
      }
    );
}
/* =========================================
   RESET TOTAL ADMIN
========================================= */
function resetearTodo() {

  const password =
    prompt(
      "Ingrese contraseña administrador"
    );

  const PASSWORD_ADMIN =
    "Admin2025";

  if (
    password !==
    PASSWORD_ADMIN
  ) {

    alert(
      "Contraseña incorrecta"
    );

    return;
  }

  const confirmar =
    confirm(
      "¿Desea desbloquear TODAS las filas enviadas?"
    );

  if (!confirmar)
    return;

  /* LIMPIAR STORAGE */
  localStorage.removeItem(
    "filasEnviadas"
  );

  localStorage.removeItem(
    "filasEnviadasInfo"
  );

  alert(
    "Todas las filas fueron desbloqueadas"
  );

  aplicarFiltros();
}