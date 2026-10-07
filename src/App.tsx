import { useEffect, useRef, useState } from "react";
import type { ChangeEvent } from "react";
import * as XLSX from "xlsx";
import "./styles.css";
import { supabase } from "./supabaseClient";

type Producto = {
  SKU: string;
  EAN: string | null;
  DESCRIPCION: string | null;
  UNIDAD_MEDIDA: string | null;
  FACTOR_WMS: string | null;
};
type AuditoriaPallet = {
  id: number;
  codigo_pallet: string;
  tienda: string | null;
  largo: number | null;
  ancho: number | null;
  alto: number | null;
  volumen_fisico: number | null;
  registrado_por: string | null;
  area: string | null;
  fecha_hora: string | null;
};
type Personal = {
  "DNI / CE": string | null;
  NOMBRE: string | null;
  CARGO: string | null;
  AREA: string | null;
  ESTADO: string | null;
};

type Pieza = {
  largo: string;
  ancho: string;
  alto: string;
};

type Medicion = {
  id: number;
  sku: string | null;
  ean: string | null;
  volumen_unitario: number | null;
  volumen_master: number | null;
  fecha_hora: string | null;
  REGISTRADO_POR?: string | null;
  AREA?: string | null;
};

type ItemCaja = {
  sku: string;
  ean: string | null;
  descripcion: string | null;
  cantidad: number;
  volumenUnitario: number;
  volumenTotal: number;
};

function App() {
  /* =========================================================
     CONSTANTES
  ========================================================= */

  const COLUMNA_DNI = "DNI / CE";
  const PAGE_SIZE_PERSONAL = 1000;

  const AREAS_EXCLUIDAS = new Set([
    "CONTROL DE RECURSOS",
    "DISTRIBUCION",
    "GERENCIA",
    "LOGISTICA",
    "RECURSOS HUMANOS",
    "SERVICIOS GENERALES",
    "SIG Y SEGURIDAD INTEGRAL",
  ]);

  /* =========================================================
     SESIÓN
  ========================================================= */

  const [sesionActiva, setSesionActiva] = useState(false);
  const [areaLogin, setAreaLogin] = useState("");
  const [personalLogin, setPersonalLogin] = useState<Personal | null>(null);
  const [dniLogin, setDniLogin] = useState("");
  const [mensajeLogin, setMensajeLogin] = useState("");
  const [ingresando, setIngresando] = useState(false);

  /* =========================================================
     NAVEGACIÓN
  ========================================================= */

  const [pantalla, setPantalla] = useState<
    "INICIO" | "PRODUCTO" | "CAJA" | "AUDITORIA_PALLET"
  >("INICIO");
  /* =========================================================
     ESTADOS - AUDITORÍA DE VOLÚMENES DE PALLETS
  ========================================================= */

  const [codigoPallet, setCodigoPallet] = useState("");
  const [proyectoPallet, setProyectoPallet] = useState("");

  const [palletLargo, setPalletLargo] = useState("");
  const [palletAncho, setPalletAncho] = useState("");
  const [palletAlto, setPalletAlto] = useState("");

  const volumenFisicoPallet =
    palletLargo && palletAncho && palletAlto
      ? (Number(palletLargo) * Number(palletAncho) * Number(palletAlto)) /
        1000000
      : 0;

  const [guardandoAuditoria, setGuardandoAuditoria] = useState(false);
  const [cargandoAuditoria, setCargandoAuditoria] = useState(false);
  const [exportandoAuditoria, setExportandoAuditoria] = useState(false);

  const [historialAuditoria, setHistorialAuditoria] = useState<
    AuditoriaPallet[]
  >([]);

  const [cajasAuditoria, setCajasAuditoria] = useState<any[]>([]);

  const [cajaSeleccionadaAuditoria, setCajaSeleccionadaAuditoria] = useState<
    any | null
  >(null);
  const [skuCajaBusqueda, setSkuCajaBusqueda] = useState("");
  const [skuCajaProducto, setSkuCajaProducto] = useState<any | null>(null);
  const [skuCajaCantidad, setSkuCajaCantidad] = useState("1");

  const [buscandoSkuCaja, setBuscandoSkuCaja] = useState(false);
  const [guardandoSkuCaja, setGuardandoSkuCaja] = useState(false);

  const [skusCajaRegistrados, setSkusCajaRegistrados] = useState<any[]>([]);
  const [agregandoCaja, setAgregandoCaja] = useState(false);
  const [auditoriaPalletId, setAuditoriaPalletId] = useState<number | null>(
    null
  );
  const [modoContenidoPallet, setModoContenidoPallet] = useState<
    "MENU" | "CAJA" | "SKU_SUELTO"
  >("MENU");
  const [skuSueltoBusqueda, setSkuSueltoBusqueda] = useState("");
  const [skuSueltoProducto, setSkuSueltoProducto] = useState<any | null>(null);
  const [skuSueltoCantidad, setSkuSueltoCantidad] = useState("1");

  const [buscandoSkuSuelto, setBuscandoSkuSuelto] = useState(false);
  const [guardandoSkuSuelto, setGuardandoSkuSuelto] = useState(false);

  const [skuSueltoRegistrados, setSkuSueltoRegistrados] = useState<any[]>([]);
  /* =========================================================
     PRODUCTO
  ========================================================= */

  const [busqueda, setBusqueda] = useState("");
  const [producto, setProducto] = useState<Producto | null>(null);

  const [largo, setLargo] = useState("");
  const [ancho, setAncho] = useState("");
  const [alto, setAlto] = useState("");

  const [esMultipieza, setEsMultipieza] = useState(false);

  const [cantidadPiezas, setCantidadPiezas] = useState(2);

  const [piezas, setPiezas] = useState<Pieza[]>(
    Array.from({ length: 6 }, () => ({
      largo: "",
      ancho: "",
      alto: "",
    }))
  );

  const [factorMedicion, setFactorMedicion] = useState("");

  const [mensaje, setMensaje] = useState("");
  const [cargando, setCargando] = useState(false);
  const [guardando, setGuardando] = useState(false);

  /* =========================================================
     EXPORTACIÓN
  ========================================================= */

  const [exportando, setExportando] = useState(false);
  const [fechaExportacion, setFechaExportacion] = useState(() => {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "America/Lima",
    }).format(new Date());
  });

  /* =========================================================
     ÚLTIMA MEDICIÓN
  ========================================================= */

  const [ultimaMedicion, setUltimaMedicion] = useState<Medicion | null>(null);

  const [cargandoUltimaMedicion, setCargandoUltimaMedicion] = useState(false);

  /* =========================================================
     ADVERTENCIA
  ========================================================= */

  const [mostrarAdvertencia, setMostrarAdvertencia] = useState(false);

  const [volumenAnteriorAdvertencia, setVolumenAnteriorAdvertencia] = useState<
    number | null
  >(null);

  /* =========================================================
     PERSONAL
  ========================================================= */

  const [personal, setPersonal] = useState<Personal[]>([]);
  const [areas, setAreas] = useState<string[]>([]);
  const [cargandoPersonal, setCargandoPersonal] = useState(false);

  const [actualizandoPersonal, setActualizandoPersonal] = useState(false);

  const inputPersonalRef = useRef<HTMLInputElement>(null);

  /* =========================================================
     TIPO MEDICIÓN
  ========================================================= */

  const [tipoMedicion, setTipoMedicion] = useState<"PRODUCTO" | "CAJA">(
    "PRODUCTO"
  );

  /* =========================================================
     CAJA MASTER
  ========================================================= */

  const [cajaLargo, setCajaLargo] = useState("");
  const [cajaAncho, setCajaAncho] = useState("");
  const [cajaAlto, setCajaAlto] = useState("");

  const [cantidadCaja, setCantidadCaja] = useState("1");

  const [itemsCaja, setItemsCaja] = useState<ItemCaja[]>([]);

  const [mostrarRevisionCaja, setMostrarRevisionCaja] = useState(false);

  const [cajaConfirmada, setCajaConfirmada] = useState(false);
  /* =========================================================
   AUDITORÍA DE VOL DE PALLETS
========================================================= */

  const entrarAuditoriaPallet = async () => {
    limpiarAuditoriaPallet();

    setPantalla("AUDITORIA_PALLET");

    await cargarHistorialAuditoria();
  };

  const limpiarAuditoriaPallet = () => {
    setCodigoPallet("");
    setProyectoPallet("");
    setPalletLargo("");
    setPalletAncho("");
    setPalletAlto("");

    // Al limpiar, se inicia una nueva auditoría desde el paso 1.
    setAuditoriaPalletId(null);
    setModoContenidoPallet("MENU");
    setCajasAuditoria([]);
    setCajaSeleccionadaAuditoria(null);
    setSkusCajaRegistrados([]);
    setSkuSueltoRegistrados([]);
    setCajaAncho("");
    setCajaAlto("");
    setCajaLargo("");
    setSkuCajaBusqueda("");
    setSkuCajaProducto(null);
    setSkuCajaCantidad("1");
    setSkuSueltoBusqueda("");
    setSkuSueltoProducto(null);
    setSkuSueltoCantidad("1");
    setMensaje("");
  };

  const guardarAuditoriaPallet = async () => {
    if (!sesionActiva) {
      setMensaje("Debe iniciar sesión antes de registrar una auditoría.");
      return;
    }

    if (!personalLogin) {
      setMensaje("No se encontró la información del usuario.");
      return;
    }

    const codigo = codigoPallet.trim().toUpperCase();
    const proyecto = proyectoPallet.trim();

    if (!codigo) {
      setMensaje("Ingrese el código del pallet.");
      return;
    }

    if (!proyecto) {
      setMensaje("Ingrese la Tienda.");
      return;
    }

    if (
      Number(palletLargo || 0) <= 0 ||
      Number(palletAncho || 0) <= 0 ||
      Number(palletAlto || 0) <= 0
    ) {
      setMensaje("Ingrese correctamente el ancho, alto y largo del pallet.");
      return;
    }

    if (volumenFisicoPallet <= 0) {
      setMensaje("El volumen físico debe ser mayor a cero.");
      return;
    }

    setGuardandoAuditoria(true);
    setMensaje("");

    try {
      const fechaHora = obtenerFechaHoraPeru();

      const datos = {
        codigo_pallet: codigo,
        tienda: proyecto,
        largo: Number(palletLargo),
        ancho: Number(palletAncho),
        alto: Number(palletAlto),
        volumen_fisico: volumenFisicoPallet,
        registrado_por: personalLogin.NOMBRE || "",
        area: personalLogin.AREA || "",
        fecha_hora: fechaHora,
      };

      const { data: auditoriaCreada, error } = await supabase
        .from("auditoria_pallet")
        .insert(datos)
        .select("id")
        .single();
      if (error) {
        throw new Error(`No se pudo guardar la auditoría: ${error.message}`);
      }
      if (!auditoriaCreada) {
        throw new Error("No se pudo obtener el ID de la auditoría creada.");
      }

      setAuditoriaPalletId(auditoriaCreada.id);
      setMensaje(
        `✓ Auditoría registrada correctamente.\n\n` +
          `Pallet: ${codigo}\n` +
          `Tienda: ${proyecto}\n` +
          `Volumen físico: ${volumenFisicoPallet.toFixed(4)} m³`
      );

      await cargarHistorialAuditoria();
    } catch (error: any) {
      console.error("ERROR GUARDANDO AUDITORÍA:", error);

      setMensaje(
        `❌ Error al guardar la auditoría: ${error?.message || String(error)}`
      );
    } finally {
      setGuardandoAuditoria(false);
    }
  };
  const agregarCajaAuditoria = async () => {
    if (!auditoriaPalletId) {
      setMensaje("Primero debe registrar el pallet.");
      return;
    }

    if (
      Number(cajaLargo || 0) <= 0 ||
      Number(cajaAncho || 0) <= 0 ||
      Number(cajaAlto || 0) <= 0
    ) {
      setMensaje("Ingrese correctamente el ancho, alto y largo de la caja.");
      return;
    }

    if (volumenCaja <= 0) {
      setMensaje("El volumen de la caja debe ser mayor a cero.");
      return;
    }

    setAgregandoCaja(true);
    setMensaje("");

    try {
      const numeroCaja = cajasAuditoria.length + 1;

      const datosCaja = {
        auditoria_pallet_id: auditoriaPalletId,
        numero_caja: numeroCaja,
        largo: Number(cajaLargo),
        ancho: Number(cajaAncho),
        alto: Number(cajaAlto),
        volumen_caja: volumenCaja,
      };

      const { data, error } = await supabase
        .from("auditoria_pallet_cajas")
        .insert(datosCaja)
        .select("*")
        .single();

      if (error) {
        throw new Error(`No se pudo registrar la caja: ${error.message}`);
      }

      if (!data) {
        throw new Error("No se pudo obtener la caja registrada.");
      }

      setCajasAuditoria((prev) => [...prev, data]);

      setCajaLargo("");
      setCajaAncho("");
      setCajaAlto("");

      setMensaje(
        `✓ Caja ${numeroCaja} registrada correctamente.\n\n` +
          `Volumen: ${Number(volumenCaja).toFixed(4)} m³`
      );

      // Nos quedamos en el formulario de cajas
      // para que el operario pueda registrar otra caja.
      setModoContenidoPallet("CAJA");
    } catch (error: any) {
      console.error("ERROR REGISTRANDO CAJA:", error);

      setMensaje(
        `❌ Error al registrar la caja: ${error?.message || String(error)}`
      );
    } finally {
      setAgregandoCaja(false);
    }
  };
  const abrirAgregarCaja = () => {
    if (!auditoriaPalletId) {
      setMensaje("Primero debe registrar el pallet.");
      return;
    }

    setMensaje("");
    setModoContenidoPallet("CAJA");
    const seleccionarCajaAuditoria = (caja: any) => {
      setCajaSeleccionadaAuditoria(caja);
      setMensaje("");
    };
  };
  const volverMenuContenido = () => {
    setMensaje("");
    setModoContenidoPallet("MENU");
  };
  const buscarProductoSkuSuelto = async () => {
    const busqueda = skuSueltoBusqueda.trim();

    if (!busqueda) {
      setMensaje("Ingrese un SKU o EAN para buscar.");
      return;
    }

    setBuscandoSkuSuelto(true);
    setMensaje("");
    setSkuSueltoProducto(null);

    try {
      // Primero buscamos por SKU
      const { data: productoPorSku, error: errorSku } = await supabase
        .from("PRODUCTOS")
        .select("SKU, EAN, DESCRIPCION, UNIDAD_MEDIDA, FACTOR_WMS")
        .eq("SKU", busqueda)
        .limit(1);

      if (errorSku) {
        throw new Error(`No se pudo buscar el producto: ${errorSku.message}`);
      }

      if (productoPorSku && productoPorSku.length > 0) {
        setSkuSueltoProducto(productoPorSku[0]);
        return;
      }

      // Si no encontró SKU, buscamos por EAN
      const { data: productoPorEan, error: errorEan } = await supabase
        .from("PRODUCTOS")
        .select("SKU, EAN, DESCRIPCION, UNIDAD_MEDIDA, FACTOR_WMS")
        .eq("EAN", busqueda)
        .limit(1);

      if (errorEan) {
        throw new Error(`No se pudo buscar el EAN: ${errorEan.message}`);
      }

      if (productoPorEan && productoPorEan.length > 0) {
        setSkuSueltoProducto(productoPorEan[0]);
        return;
      }

      setMensaje(`No se encontró ningún producto con SKU/EAN: ${busqueda}`);
    } catch (error: any) {
      console.error("ERROR BUSCANDO SKU SUELTO:", error);

      setMensaje(
        `❌ Error buscando producto: ${error?.message || String(error)}`
      );
    } finally {
      setBuscandoSkuSuelto(false);
    }
  };
  const buscarProductoSkuCaja = async () => {
    const busqueda = skuCajaBusqueda.trim();

    if (!busqueda) {
      setMensaje("Ingrese un SKU o EAN para buscar.");
      return;
    }

    setBuscandoSkuCaja(true);
    setMensaje("");
    setSkuCajaProducto(null);

    try {
      // =====================================================
      // BUSCAR PRIMERO POR SKU
      // =====================================================

      const { data: productoPorSku, error: errorSku } = await supabase
        .from("PRODUCTOS")
        .select("SKU, EAN, DESCRIPCION, UNIDAD_MEDIDA, FACTOR_WMS")
        .eq("SKU", busqueda)
        .limit(1);

      if (errorSku) {
        throw new Error(`No se pudo buscar el producto: ${errorSku.message}`);
      }

      if (productoPorSku && productoPorSku.length > 0) {
        setSkuCajaProducto(productoPorSku[0]);
        return;
      }

      // =====================================================
      // SI NO ENCUENTRA SKU, BUSCAR POR EAN
      // =====================================================

      const { data: productoPorEan, error: errorEan } = await supabase
        .from("PRODUCTOS")
        .select("SKU, EAN, DESCRIPCION, UNIDAD_MEDIDA, FACTOR_WMS")
        .eq("EAN", busqueda)
        .limit(1);

      if (errorEan) {
        throw new Error(`No se pudo buscar el EAN: ${errorEan.message}`);
      }

      if (productoPorEan && productoPorEan.length > 0) {
        setSkuCajaProducto(productoPorEan[0]);
        return;
      }

      setMensaje(`No se encontró ningún producto con SKU/EAN: ${busqueda}`);
    } catch (error: any) {
      console.error("ERROR BUSCANDO PRODUCTO DE CAJA:", error);

      setMensaje(
        `❌ Error buscando producto: ${error?.message || String(error)}`
      );
    } finally {
      setBuscandoSkuCaja(false);
    }
  };
  const agregarSkuCajaAuditoria = async () => {
    if (!auditoriaPalletId) {
      setMensaje("Primero debe registrar el pallet.");
      return;
    }

    if (!cajaSeleccionadaAuditoria) {
      setMensaje("Primero debe seleccionar una caja.");
      return;
    }

    if (!skuCajaProducto) {
      setMensaje("Primero debe buscar un SKU o EAN.");
      return;
    }

    const cantidad = Number(skuCajaCantidad || 0);

    if (cantidad <= 0) {
      setMensaje("Ingrese una cantidad mayor a cero.");
      return;
    }

    setGuardandoSkuCaja(true);
    setMensaje("");

    try {
      const sku = String(skuCajaProducto.SKU || "").trim();
      const ean = String(skuCajaProducto.EAN || "").trim();
      const descripcion = skuCajaProducto.DESCRIPCION || "";

      // Buscar volumen unitario en DATA_MAESTRO
      let volumenUnitario = 0;

      const { data: maestro, error: errorMaestro } = await supabase
        .from("DATA_MAESTRO")
        .select("VOLUMEN_UNITARIO")
        .eq("SKU", sku)
        .limit(1);

      if (errorMaestro) {
        console.warn(
          "No se pudo consultar DATA_MAESTRO:",
          errorMaestro.message
        );
      }

      if (
        maestro &&
        maestro.length > 0 &&
        maestro[0].VOLUMEN_UNITARIO != null
      ) {
        volumenUnitario = Number(maestro[0].VOLUMEN_UNITARIO);
      }

      const volumenTotal = volumenUnitario * cantidad;

      const datos = {
        auditoria_caja_id: cajaSeleccionadaAuditoria.id,
        sku,
        ean: ean || null,
        descripcion,
        cantidad,
        volumen_unitario_maestro: volumenUnitario,
        volumen_total: volumenTotal,
      };

      const { data, error } = await supabase
        .from("auditoria_pallet_caja_detalle")
        .insert(datos)
        .select("*")
        .single();

      if (error) {
        throw new Error(
          `No se pudo guardar el SKU en la caja: ${error.message}`
        );
      }

      if (!data) {
        throw new Error("No se pudo obtener el SKU registrado.");
      }

      setSkusCajaRegistrados((prev) => [...prev, data]);

      setSkuCajaBusqueda("");
      setSkuCajaProducto(null);
      setSkuCajaCantidad("1");

      setMensaje(
        `✓ SKU agregado correctamente a la Caja ${cajaSeleccionadaAuditoria.numero_caja}.\n\n` +
          `SKU: ${sku}\n` +
          `Cantidad: ${cantidad}\n` +
          `Volumen total: ${volumenTotal.toFixed(4)} m³`
      );
    } catch (error: any) {
      console.error("ERROR GUARDANDO SKU EN CAJA:", error);

      setMensaje(
        `❌ Error guardando SKU en la caja: ${error?.message || String(error)}`
      );
    } finally {
      setGuardandoSkuCaja(false);
    }
  };
  const agregarSkuSueltoAuditoria = async () => {
    if (!auditoriaPalletId) {
      setMensaje("Primero debe registrar el pallet.");
      return;
    }

    if (!skuSueltoProducto) {
      setMensaje("Primero debe buscar y seleccionar un producto.");
      return;
    }

    const cantidad = Number(skuSueltoCantidad || 0);

    if (cantidad <= 0) {
      setMensaje("Ingrese una cantidad mayor a cero.");
      return;
    }

    setGuardandoSkuSuelto(true);
    setMensaje("");

    try {
      const sku = String(skuSueltoProducto.SKU || "").trim();

      const ean = String(skuSueltoProducto.EAN || "").trim();

      const descripcion = skuSueltoProducto.DESCRIPCION || "";

      /*
        Por ahora tomamos el volumen unitario desde
        DATA_MAESTRO cuando exista.
      */

      let volumenUnitario = 0;

      const { data: maestro, error: errorMaestro } = await supabase
        .from("DATA_MAESTRO")
        .select("VOLUMEN_UNITARIO")
        .eq("SKU", sku)
        .limit(1);

      if (errorMaestro) {
        console.warn(
          "No se pudo consultar DATA_MAESTRO:",
          errorMaestro.message
        );
      }

      if (
        maestro &&
        maestro.length > 0 &&
        maestro[0].VOLUMEN_UNITARIO != null
      ) {
        volumenUnitario = Number(maestro[0].VOLUMEN_UNITARIO);
      }

      const volumenTotal = volumenUnitario * cantidad;

      const datos = {
        auditoria_pallet_id: auditoriaPalletId,
        sku,
        ean: ean || null,
        descripcion,
        cantidad,
        volumen_unitario_maestro: volumenUnitario,
        volumen_total: volumenTotal,
      };

      const { data, error } = await supabase
        .from("auditoria_pallet_sku_suelto")
        .insert(datos)
        .select("*")
        .single();

      if (error) {
        throw new Error(`No se pudo registrar el SKU suelto: ${error.message}`);
      }

      if (!data) {
        throw new Error("No se pudo obtener el registro creado.");
      }

      setSkuSueltoRegistrados((prev) => [...prev, data]);

      setSkuSueltoBusqueda("");
      setSkuSueltoProducto(null);
      setSkuSueltoCantidad("1");

      setMensaje(
        `✓ SKU suelto registrado correctamente.\n\n` +
          `SKU: ${sku}\n` +
          `Descripción: ${descripcion}\n` +
          `Cantidad: ${cantidad}`
      );
    } catch (error: any) {
      console.error("ERROR REGISTRANDO SKU SUELTO:", error);

      setMensaje(
        `❌ Error registrando SKU suelto: ${error?.message || String(error)}`
      );
    } finally {
      setGuardandoSkuSuelto(false);
    }
  };
  const cargarHistorialAuditoria = async () => {
    setCargandoAuditoria(true);

    try {
      const { data, error } = await supabase
        .from("auditoria_pallet")
        .select(
          `
          id,
          codigo_pallet,
          tienda,
          largo,
          ancho,
          alto,
          volumen_fisico,
          registrado_por,
          area,
          fecha_hora
        `
        )
        .order("fecha_hora", {
          ascending: false,
        })
        .limit(500);

      if (error) {
        throw new Error(`No se pudo cargar el historial: ${error.message}`);
      }

      setHistorialAuditoria((data || []) as AuditoriaPallet[]);
    } catch (error: any) {
      console.error("ERROR CARGANDO AUDITORÍA:", error);

      setMensaje(
        `No se pudo cargar el historial: ${error?.message || String(error)}`
      );
    } finally {
      setCargandoAuditoria(false);
    }
  };

  const exportarAuditoriaPallet = async () => {
    if (exportandoAuditoria) return;

    setExportandoAuditoria(true);
    setMensaje("");

    try {
      const { data, error } = await supabase
        .from("auditoria_pallet")
        .select(
          `
          id,
          codigo_pallet,
          tienda,
          largo,
          ancho,
          alto,
          volumen_fisico,
          registrado_por,
          area,
          fecha_hora
        `
        )
        .order("fecha_hora", {
          ascending: true,
        });

      if (error) {
        throw new Error(`No se pudo consultar la auditoría: ${error.message}`);
      }

      if (!data || data.length === 0) {
        setMensaje("No existen auditorías de pallets para exportar.");
        return;
      }

      const datosExcel = (data as AuditoriaPallet[]).map((registro) => ({
        "CÓDIGO PALLET": registro.codigo_pallet,

        TIENDA: registro.tienda || "",

        "ANCHO (cm.)": registro.ancho ?? "",

        "ALTO (cm.)": registro.alto ?? "",

        "LARGO (cm.)": registro.largo ?? "",

        "VOLUMEN FÍSICO (m³)": registro.volumen_fisico ?? "",

        "REGISTRADO POR": registro.registrado_por || "",

        ÁREA: registro.area || "",

        FECHA: registro.fecha_hora ? formatearFecha(registro.fecha_hora) : "",

        HORA: registro.fecha_hora ? formatearHora(registro.fecha_hora) : "",
      }));

      const libro = XLSX.utils.book_new();

      const hoja = XLSX.utils.json_to_sheet(datosExcel);

      XLSX.utils.book_append_sheet(libro, hoja, "AUDITORIA PALLETS");

      XLSX.writeFile(
        libro,
        `AUDITORIA_VOL_PALLETS_${new Date()
          .toLocaleDateString("es-PE", {
            timeZone: "America/Lima",
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
          })
          .replace(/\//g, "-")}.xlsx`
      );

      setMensaje(
        `✓ Excel generado correctamente.\n\n` +
          `Registros exportados: ${data.length.toLocaleString("es-PE")}`
      );
    } catch (error: any) {
      console.error("ERROR EXPORTANDO AUDITORÍA:", error);

      setMensaje(
        `❌ Error al generar Excel: ${error?.message || String(error)}`
      );
    } finally {
      setExportandoAuditoria(false);
    }
  };
  /* =========================================================
     REFERENCIAS
  ========================================================= */

  const temporizadorEAN = useRef<ReturnType<typeof setTimeout> | null>(null);

  const inputBusquedaRef = useRef<HTMLInputElement>(null);

  const inputDniRef = useRef<HTMLInputElement>(null);

  /* =========================================================
     FECHA HORA PERÚ
  ========================================================= */

  const obtenerFechaHoraPeru = () => {
    const ahora = new Date();

    const partes = new Intl.DateTimeFormat("en-CA", {
      timeZone: "America/Lima",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    }).formatToParts(ahora);

    const obtener = (tipo: string) =>
      partes.find((parte) => parte.type === tipo)?.value || "";

    return `${obtener("year")}-${obtener("month")}-${obtener("day")}T${obtener(
      "hour"
    )}:${obtener("minute")}:${obtener("second")}-05:00`;
  };

  const formatearFecha = (fechaHora: string | null) => {
    if (!fechaHora) return "";

    const fecha = new Date(fechaHora);

    if (isNaN(fecha.getTime())) return "";

    return fecha.toLocaleDateString("es-PE", {
      timeZone: "America/Lima",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
  };

  const formatearHora = (fechaHora: string | null) => {
    if (!fechaHora) return "";

    const fecha = new Date(fechaHora);

    if (isNaN(fecha.getTime())) return "";

    return fecha.toLocaleTimeString("es-PE", {
      timeZone: "America/Lima",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: false,
    });
  };

  /* =========================================================
     NORMALIZADORES
  ========================================================= */

  const normalizarEncabezado = (valor: any) => {
    return String(valor ?? "")
      .trim()
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  };

  const normalizarDNI = (valor: any) => {
    if (valor === null || valor === undefined) {
      return "";
    }

    let dni = String(valor).trim().toUpperCase().replace(/\s+/g, "");

    if (/^\d+\.0+$/.test(dni)) {
      dni = dni.split(".")[0];
    }

    return dni;
  };

  const obtenerDNI = (persona: Partial<Personal>) => {
    return normalizarDNI(persona[COLUMNA_DNI]);
  };

  const normalizarArea = (valor: any) => {
    return String(valor ?? "")
      .trim()
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  };

  /* =========================================================
     OBTENER TODO PERSONAL
  ========================================================= */

  const obtenerTodoElPersonal = async (): Promise<Personal[]> => {
    const todosLosRegistros: Personal[] = [];

    let desde = 0;

    while (true) {
      const hasta = desde + PAGE_SIZE_PERSONAL - 1;

      const { data, error } = await supabase
        .from("PERSONAL")
        .select('"DNI / CE",NOMBRE,CARGO,AREA,ESTADO')
        .range(desde, hasta);

      if (error) {
        throw new Error(`No se pudo consultar PERSONAL: ${error.message}`);
      }

      const bloque = (data || []) as Personal[];

      todosLosRegistros.push(...bloque);

      if (bloque.length < PAGE_SIZE_PERSONAL) {
        break;
      }

      desde += PAGE_SIZE_PERSONAL;
    }

    return todosLosRegistros;
  };

  /* =========================================================
     CARGAR PERSONAL ACTIVO OPERATIVO
  ========================================================= */

  const cargarPersonal = async () => {
    setCargandoPersonal(true);

    try {
      const datosCompletos = await obtenerTodoElPersonal();

      const personalActivo = datosCompletos.filter((persona) => {
        const estado = String(persona.ESTADO || "")
          .trim()
          .toUpperCase();

        const area = normalizarArea(persona.AREA);

        return estado === "ACTIVO" && !AREAS_EXCLUIDAS.has(area);
      });

      setPersonal([...personalActivo]);

      const mapaAreas = new Map<string, string>();

      personalActivo.forEach((persona) => {
        const areaOriginal = String(persona.AREA || "").trim();

        if (!areaOriginal) return;

        const areaNormalizada = normalizarArea(areaOriginal);

        if (!mapaAreas.has(areaNormalizada)) {
          mapaAreas.set(areaNormalizada, areaOriginal);
        }
      });

      const areasActualizadas = Array.from(mapaAreas.values()).sort((a, b) =>
        a.localeCompare(b, "es")
      );

      setAreas([...areasActualizadas]);

      setAreaLogin((areaActual) => {
        const sigueExistiendo = areasActualizadas.some(
          (area) => normalizarArea(area) === normalizarArea(areaActual)
        );

        return sigueExistiendo ? areaActual : "";
      });

      setPersonalLogin((personaActual) => {
        if (!personaActual) return null;

        const dniActual = obtenerDNI(personaActual);

        const personaSigueActiva = personalActivo.some(
          (persona) => obtenerDNI(persona) === dniActual
        );

        return personaSigueActiva ? personaActual : null;
      });
    } catch (error: any) {
      console.error("ERROR CARGANDO PERSONAL:", error);

      setMensajeLogin(
        `No se pudo cargar la dotación de personal: ${
          error?.message || String(error)
        }`
      );
    } finally {
      setCargandoPersonal(false);
    }
  };

  useEffect(() => {
    cargarPersonal();
  }, []);

  /* =========================================================
     ACTUALIZAR PERSONAL DESDE ARCHIVO
  ========================================================= */

  const actualizarPersonalDesdeArchivo = async (archivo: File) => {
    if (actualizandoPersonal) return;

    setActualizandoPersonal(true);
    setMensaje("");
    setMensajeLogin("");

    try {
      const extension = archivo.name.split(".").pop()?.toLowerCase();

      if (!extension || !["csv", "xlsx", "xls"].includes(extension)) {
        throw new Error("Seleccione un archivo CSV, XLSX o XLS.");
      }

      const buffer = await archivo.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: "array",
      });

      if (!workbook.SheetNames.length) {
        throw new Error("El archivo no contiene ninguna hoja.");
      }

      const hoja = workbook.Sheets[workbook.SheetNames[0]];

      if (!hoja) {
        throw new Error("No se pudo leer la primera hoja del archivo.");
      }

      const filas = XLSX.utils.sheet_to_json<Record<string, any>>(hoja, {
        defval: "",
      });

      if (!filas.length) {
        throw new Error("El archivo no contiene registros.");
      }

      /* =====================================================
         ENCABEZADOS
      ===================================================== */

      const primeraFila = filas[0];

      const encabezados = Object.keys(primeraFila);

      const mapaEncabezados: Record<string, string> = {};

      encabezados.forEach((encabezado) => {
        mapaEncabezados[normalizarEncabezado(encabezado)] = encabezado;
      });

      const obtenerEncabezado = (posibles: string[]) => {
        for (const posible of posibles) {
          const encontrado = mapaEncabezados[normalizarEncabezado(posible)];

          if (encontrado) {
            return encontrado;
          }
        }

        return null;
      };

      const encabezadoDNI = obtenerEncabezado([
        "DNI / CE",
        "DNI/CE",
        "DNI",
        "CE",
      ]);

      const encabezadoNombre = obtenerEncabezado([
        "NOMBRE",
        "NOMBRES",
        "NOMBRE COMPLETO",
      ]);

      const encabezadoCargo = obtenerEncabezado(["CARGO", "PUESTO"]);

      const encabezadoArea = obtenerEncabezado(["AREA", "ÁREA"]);

      const encabezadoEstado = obtenerEncabezado(["ESTADO", "STATUS"]);

      const faltantes: string[] = [];

      if (!encabezadoDNI) faltantes.push("DNI / CE");
      if (!encabezadoNombre) faltantes.push("NOMBRE");
      if (!encabezadoCargo) faltantes.push("CARGO");
      if (!encabezadoArea) faltantes.push("AREA");
      if (!encabezadoEstado) faltantes.push("ESTADO");

      if (faltantes.length > 0) {
        throw new Error(
          `Faltan las siguientes columnas en el archivo: ${faltantes.join(
            ", "
          )}`
        );
      }

      /* =====================================================
         PROCESAR ARCHIVO
      ===================================================== */

      const personalArchivo: Personal[] = [];

      const DNIsVistos = new Set<string>();

      let duplicadosArchivo = 0;
      let filasSinDNI = 0;
      let registrosAreasExcluidas = 0;

      for (const fila of filas) {
        const area = String(fila[encabezadoArea!] ?? "").trim();

        const areaNormalizada = normalizarArea(area);

        if (AREAS_EXCLUIDAS.has(areaNormalizada)) {
          registrosAreasExcluidas++;
          continue;
        }

        const dni = normalizarDNI(fila[encabezadoDNI!]);

        if (!dni) {
          filasSinDNI++;
          continue;
        }

        if (DNIsVistos.has(dni)) {
          duplicadosArchivo++;
          continue;
        }

        DNIsVistos.add(dni);

        const nombre = String(fila[encabezadoNombre!] ?? "").trim();

        const cargo = String(fila[encabezadoCargo!] ?? "").trim();

        const estadoArchivo = String(fila[encabezadoEstado!] ?? "")
          .trim()
          .toUpperCase();

        personalArchivo.push({
          [COLUMNA_DNI]: dni,
          NOMBRE: nombre || null,
          CARGO: cargo || null,
          AREA: area || null,
          ESTADO: estadoArchivo || "ACTIVO",
        });
      }

      if (!personalArchivo.length) {
        throw new Error(
          "No se encontraron registros operativos válidos con DNI / CE en el archivo."
        );
      }

      /* =====================================================
         PERSONAL ACTUAL EN SUPABASE
      ===================================================== */

      const personalActual = await obtenerTodoElPersonal();

      const registrosActuales = personalActual.map((persona) => ({
        [COLUMNA_DNI]: normalizarDNI(persona[COLUMNA_DNI]),

        NOMBRE:
          persona.NOMBRE === null || persona.NOMBRE === undefined
            ? null
            : String(persona.NOMBRE).trim(),

        CARGO:
          persona.CARGO === null || persona.CARGO === undefined
            ? null
            : String(persona.CARGO).trim(),

        AREA:
          persona.AREA === null || persona.AREA === undefined
            ? null
            : String(persona.AREA).trim(),

        ESTADO:
          persona.ESTADO === null || persona.ESTADO === undefined
            ? null
            : String(persona.ESTADO).trim().toUpperCase(),
      })) as Personal[];

      /* =====================================================
         MAPA ACTUAL
      ===================================================== */

      const mapaActual = new Map<string, Personal>();

      let duplicadosSupabase = 0;

      registrosActuales.forEach((persona) => {
        const dni = obtenerDNI(persona);

        if (!dni) return;

        if (!mapaActual.has(dni)) {
          mapaActual.set(dni, persona);
        } else {
          duplicadosSupabase++;
        }
      });

      /* =====================================================
         NUEVOS / EXISTENTES
      ===================================================== */

      const nuevos: Personal[] = [];

      const existentes: {
        actual: Personal;
        nuevo: Personal;
      }[] = [];

      personalArchivo.forEach((personaNueva) => {
        const dni = obtenerDNI(personaNueva);

        const personaActual = mapaActual.get(dni);

        if (!personaActual) {
          nuevos.push(personaNueva);
        } else {
          existentes.push({
            actual: personaActual,
            nuevo: personaNueva,
          });
        }
      });

      /* =====================================================
         INSERTAR NUEVOS
      ===================================================== */

      let cantidadNuevos = 0;

      if (nuevos.length > 0) {
        const { error: errorInsert } = await supabase
          .from("PERSONAL")
          .insert(nuevos);

        if (errorInsert) {
          throw new Error(
            "Error insertando personal nuevo: " + errorInsert.message
          );
        }

        cantidadNuevos = nuevos.length;
      }

      /* =====================================================
         ACTUALIZAR EXISTENTES
      ===================================================== */

      let cantidadActualizados = 0;

      const erroresActualizacion: string[] = [];

      for (const registro of existentes) {
        const actual = registro.actual;
        const nuevo = registro.nuevo;

        const dniActual = obtenerDNI(actual);

        if (!dniActual) continue;

        const huboCambios =
          String(actual.NOMBRE || "").trim() !==
            String(nuevo.NOMBRE || "").trim() ||
          String(actual.CARGO || "").trim() !==
            String(nuevo.CARGO || "").trim() ||
          normalizarArea(actual.AREA) !== normalizarArea(nuevo.AREA) ||
          String(actual.ESTADO || "")
            .trim()
            .toUpperCase() !==
            String(nuevo.ESTADO || "")
              .trim()
              .toUpperCase();

        if (!huboCambios) {
          continue;
        }

        const { error: errorUpdate } = await supabase
          .from("PERSONAL")
          .update({
            NOMBRE: nuevo.NOMBRE,
            CARGO: nuevo.CARGO,
            AREA: nuevo.AREA,
            ESTADO: nuevo.ESTADO,
          })
          .filter('"DNI / CE"', "eq", dniActual);

        if (errorUpdate) {
          erroresActualizacion.push(`${dniActual}: ${errorUpdate.message}`);
          continue;
        }

        cantidadActualizados++;
      }

      /* =====================================================
         INACTIVAR AUSENTES
      ===================================================== */

      let cantidadInactivados = 0;

      const erroresInactivacion: string[] = [];

      const DNIsParaInactivar = new Set<string>();

      registrosActuales.forEach((personaActual) => {
        const dniActual = obtenerDNI(personaActual);

        if (!dniActual) return;

        if (!DNIsVistos.has(dniActual)) {
          DNIsParaInactivar.add(dniActual);
        }
      });

      for (const dniInactivar of DNIsParaInactivar) {
        const { error: errorInactivar } = await supabase
          .from("PERSONAL")
          .update({
            ESTADO: "INACTIVO",
          })
          .filter('"DNI / CE"', "eq", dniInactivar);

        if (errorInactivar) {
          erroresInactivacion.push(
            `${dniInactivar}: ${errorInactivar.message}`
          );
          continue;
        }

        cantidadInactivados++;
      }

      /* =====================================================
         RECARGAR PERSONAL
      ===================================================== */

      await cargarPersonal();

      /* =====================================================
         MENSAJE FINAL
      ===================================================== */

      let mensajeFinal =
        `✓ PERSONAL procesado correctamente.\n\n` +
        `Archivo: ${filas.length.toLocaleString("es-PE")} registros\n` +
        `Válidos: ${personalArchivo.length.toLocaleString("es-PE")}\n` +
        `Excluidos por área: ${registrosAreasExcluidas.toLocaleString(
          "es-PE"
        )}\n` +
        `Nuevos: ${cantidadNuevos.toLocaleString("es-PE")}\n` +
        `Actualizados: ${cantidadActualizados.toLocaleString("es-PE")}\n` +
        `Inactivados: ${cantidadInactivados.toLocaleString("es-PE")}\n` +
        `Duplicados ignorados: ${duplicadosArchivo.toLocaleString("es-PE")}`;

      if (filasSinDNI > 0) {
        mensajeFinal += `\nFilas sin DNI / CE: ${filasSinDNI.toLocaleString(
          "es-PE"
        )}`;
      }

      if (duplicadosSupabase > 0) {
        mensajeFinal += `\n⚠️ Duplicados encontrados actualmente en Supabase: ${duplicadosSupabase}`;
      }

      if (erroresActualizacion.length > 0) {
        mensajeFinal +=
          `\n\n⚠️ Errores al actualizar: ${erroresActualizacion.length}` +
          `\nPrimer error:\n${erroresActualizacion[0]}`;

        console.error("ERRORES ACTUALIZACIÓN:", erroresActualizacion);
      }

      if (erroresInactivacion.length > 0) {
        mensajeFinal +=
          `\n\n⚠️ Errores al inactivar: ${erroresInactivacion.length}` +
          `\nPrimer error:\n${erroresInactivacion[0]}`;

        console.error("ERRORES INACTIVACIÓN:", erroresInactivacion);
      }

      setMensaje(mensajeFinal);
    } catch (error: any) {
      console.error("ERROR ACTUALIZANDO PERSONAL:", error);

      setMensaje(
        `❌ No se pudo actualizar PERSONAL.\n\n${
          error?.message || String(error)
        }`
      );
    } finally {
      setActualizandoPersonal(false);

      if (inputPersonalRef.current) {
        inputPersonalRef.current.value = "";
      }
    }
  };

  /* =========================================================
     ARCHIVO PERSONAL
  ========================================================= */

  const seleccionarArchivoPersonal = () => {
    inputPersonalRef.current?.click();
  };

  const manejarArchivoPersonal = (event: ChangeEvent<HTMLInputElement>) => {
    const archivo = event.target.files?.[0];

    if (!archivo) return;

    const extension = archivo.name.split(".").pop()?.toLowerCase();

    if (!extension || !["csv", "xlsx", "xls"].includes(extension)) {
      setMensaje("Seleccione un archivo CSV, XLSX o XLS.");

      event.target.value = "";

      return;
    }

    actualizarPersonalDesdeArchivo(archivo);
  };

  /* =========================================================
     LOGIN
  ========================================================= */

  const personalFiltradoLogin = personal.filter(
    (persona) =>
      normalizarArea(persona.AREA) === normalizarArea(areaLogin) &&
      !AREAS_EXCLUIDAS.has(normalizarArea(persona.AREA))
  );

  const cambiarAreaLogin = (area: string) => {
    setAreaLogin(area);
    setPersonalLogin(null);
    setDniLogin("");
    setMensajeLogin("");
  };

  const cambiarPersonalLogin = (dni: string) => {
    const dniNormalizado = normalizarDNI(dni);

    const seleccionado =
      personalFiltradoLogin.find(
        (persona) => obtenerDNI(persona) === dniNormalizado
      ) || null;

    setPersonalLogin(seleccionado);

    setDniLogin("");
    setMensajeLogin("");

    if (seleccionado) {
      setTimeout(() => {
        inputDniRef.current?.focus();
      }, 100);
    }
  };

  /* =========================================================
     INICIAR SESIÓN
  ========================================================= */

  const iniciarSesion = async () => {
    setMensajeLogin("");

    if (!areaLogin) {
      setMensajeLogin("Seleccione un área.");
      return;
    }

    if (!personalLogin) {
      setMensajeLogin("Seleccione el personal.");
      return;
    }

    const dniIngresado = normalizarDNI(dniLogin);

    if (!dniIngresado) {
      setMensajeLogin("Ingrese su DNI / CE.");
      return;
    }

    const dniRegistrado = obtenerDNI(personalLogin);

    if (dniIngresado !== dniRegistrado) {
      setMensajeLogin(
        "El DNI / CE ingresado no corresponde al personal seleccionado."
      );
      return;
    }

    const areaRegistrada = normalizarArea(personalLogin.AREA);
    const areaSeleccionada = normalizarArea(areaLogin);

    if (AREAS_EXCLUIDAS.has(areaRegistrada)) {
      setMensajeLogin("El área seleccionada no está habilitada para operar.");
      return;
    }

    if (areaRegistrada !== areaSeleccionada) {
      setMensajeLogin("El área seleccionada no corresponde al personal.");
      return;
    }

    const estado = String(personalLogin.ESTADO || "")
      .trim()
      .toUpperCase();

    if (estado !== "ACTIVO") {
      setMensajeLogin(
        `El usuario se encuentra ${estado || "SIN ESTADO"} y no puede ingresar.`
      );
      return;
    }

    setIngresando(true);

    try {
      setSesionActiva(true);
      setPantalla("INICIO");
      setMensajeLogin("");
      limpiarEstadosMedicion();
    } catch (error: any) {
      setMensajeLogin(
        `Error al iniciar sesión: ${error?.message || String(error)}`
      );
    } finally {
      setIngresando(false);
    }
  };

  /* =========================================================
     LIMPIAR ESTADOS
  ========================================================= */

  const limpiarEstadosMedicion = () => {
    setBusqueda("");
    setProducto(null);
    setUltimaMedicion(null);

    setLargo("");
    setAncho("");
    setAlto("");

    setEsMultipieza(false);
    setCantidadPiezas(2);

    setPiezas(
      Array.from({ length: 6 }, () => ({
        largo: "",
        ancho: "",
        alto: "",
      }))
    );

    setFactorMedicion("");

    setTipoMedicion("PRODUCTO");

    setCajaLargo("");
    setCajaAncho("");
    setCajaAlto("");
    setCantidadCaja("1");
    setItemsCaja([]);

    setMostrarRevisionCaja(false);
    setCajaConfirmada(false);

    setMensaje("");

    setMostrarAdvertencia(false);
    setVolumenAnteriorAdvertencia(null);
  };

  /* =========================================================
     EXPORTAR DATA MAESTRO
  ========================================================= */
  const exportarDataMaestro = async () => {
    if (exportando) return;

    if (!fechaExportacion) {
      setMensaje("Seleccione una fecha para exportar.");
      return;
    }

    setExportando(true);
    setMensaje("");

    try {
      const todasLasFilas: Record<string, any>[] = [];

      /*
       * =========================================================
       * CONVERTIR LA FECHA SELECCIONADA DE LIMA A UTC
       * =========================================================
       *
       * Ejemplo:
       *
       * fechaExportacion = 2026-09-25
       *
       * Inicio Lima:
       * 2026-09-25 00:00:00
       *
       * Fin Lima:
       * 2026-09-26 00:00:00
       *
       * Supabase guarda timestamptz en UTC, por eso usamos
       * los límites correspondientes.
       */

      const inicioDia = new Date(`${fechaExportacion}T00:00:00-05:00`);

      const finDia = new Date(`${fechaExportacion}T00:00:00-05:00`);

      finDia.setDate(finDia.getDate() + 1);

      const inicioISO = inicioDia.toISOString();
      const finISO = finDia.toISOString();

      /*
       * =========================================================
       * CONSULTAR MEDICIONES DEL DÍA
       * =========================================================
       *
       * IMPORTANTE:
       *
       * NO usamos DATA_MAESTRO porque esa VIEW solamente conserva
       * la última medición de cada SKU.
       *
       * Aquí necesitamos TODAS las mediciones realizadas en la
       * fecha seleccionada.
       */

      const pageSize = 1000;

      let desde = 0;

      while (true) {
        const hasta = desde + pageSize - 1;

        const { data, error } = await supabase
          .from("MEDICIONES")
          .select("*")
          .gte("fecha_hora", inicioISO)
          .lt("fecha_hora", finISO)
          .neq("REGISTRADO_POR", "MIGRACION MAESTRO")
          .neq("AREA", "MIGRACION")
          .order("fecha_hora", { ascending: true })
          .order("id", { ascending: true })
          .range(desde, hasta);

        if (error) {
          throw new Error(
            `No se pudieron consultar las mediciones: ${error.message}`
          );
        }

        if (!data || data.length === 0) {
          break;
        }

        todasLasFilas.push(...data);

        if (data.length < pageSize) {
          break;
        }

        desde += pageSize;
      }

      /*
       * =========================================================
       * VALIDAR RESULTADOS
       * =========================================================
       */

      if (todasLasFilas.length === 0) {
        const fechaMostrar = new Date(
          `${fechaExportacion}T12:00:00-05:00`
        ).toLocaleDateString("es-PE", {
          timeZone: "America/Lima",
          day: "2-digit",
          month: "2-digit",
          year: "numeric",
        });

        setMensaje(`No existen mediciones realizadas el ${fechaMostrar}.`);

        return;
      }

      /*
       * =========================================================
       * OBTENER PRODUCTOS
       * =========================================================
       *
       * Necesitamos EAN y DESCRIPCION de PRODUCTOS.
       */

      const skus = Array.from(
        new Set(
          todasLasFilas
            .map((fila) => String(fila.sku || "").trim())
            .filter(Boolean)
        )
      );

      const productosMap = new Map<string, Record<string, any>>();

      /*
       * Supabase tiene límites para filtros .in(), por eso
       * trabajamos por bloques.
       */

      const bloqueProductos = 500;

      for (let i = 0; i < skus.length; i += bloqueProductos) {
        const bloque = skus.slice(i, i + bloqueProductos);

        const { data: productos, error } = await supabase
          .from("PRODUCTOS")
          .select(`"SKU","EAN","DESCRIPCION","UNIDAD_MEDIDA","FACTOR_WMS"`)
          .in("SKU", bloque);

        if (error) {
          throw new Error(
            `No se pudieron consultar los productos: ${error.message}`
          );
        }

        if (productos) {
          productos.forEach((producto) => {
            const sku = String(producto["SKU"] || "").trim();

            if (!productosMap.has(sku)) {
              productosMap.set(sku, producto);
            }
          });
        }
      }

      /*
       * =========================================================
       * OBTENER ÚLTIMO MASTER POR SKU
       * =========================================================
       *
       * Esto replica la lógica de DATA_MAESTRO:
       *
       * último CAJAS_MASTER_DETALLE por fecha_hora + id.
       */

      const mastersMap = new Map<string, Record<string, any>>();

      const { data: detallesMaster, error: errorMaster } = await supabase
        .from("CAJAS_MASTER_DETALLE")
        .select(
          `
            sku,
            id,
            fecha_hora,
            caja_master_id,
            CAJAS_MASTER!inner(
              id,
              largo,
              ancho,
              alto
            )
          `
        )
        .in("sku", skus)
        .order("fecha_hora", { ascending: false })
        .order("id", { ascending: false });

      if (errorMaster) {
        throw new Error(
          `No se pudieron consultar las cajas master: ${errorMaster.message}`
        );
      }

      if (detallesMaster) {
        detallesMaster.forEach((detalle: any) => {
          const sku = String(detalle.sku || "").trim();

          if (!sku || mastersMap.has(sku)) {
            return;
          }

          const caja = detalle.CAJAS_MASTER;

          if (!caja) {
            return;
          }

          mastersMap.set(sku, {
            ancho: caja.ancho,
            largo: caja.largo,
            alto: caja.alto,
          });
        });
      }

      /*
       * =========================================================
       * COLUMNAS
       * =========================================================
       */

      const encabezadosExcel = [
        "SKU",
        "EAN",
        "DESCRIPCIÓN",
        "FACTOR MED",
        "FECHA",

        "ANCHO (cm.) MASTER",
        "LARGO (cm.) MASTER",
        "ALTO (cm.) MASTER",

        "ANCHO (cm.) UND",
        "LARGO (cm.) UND",
        "ALTO (cm.) UND",

        "VOL PZA 1",

        "ANCHO (cm.) PZA 2",
        "LARGO (cm.) PZA 2",
        "ALTO (cm.) PZA 2",
        "VOL PZA 2",

        "ANCHO (cm.) PZA 3",
        "LARGO (cm.) PZA 3",
        "ALTO (cm.) PZA 3",
        "VOL PZA 3",

        "ANCHO (cm.) PZA 4",
        "LARGO (cm.) PZA 4",
        "ALTO (cm.) PZA 4",
        "VOL PZA 4",

        "ANCHO (cm.) PZA 5",
        "LARGO (cm.) PZA 5",
        "ALTO (cm.) PZA 5",
        "VOL PZA 5",

        "ANCHO (cm.) PZA 6",
        "LARGO (cm.) PZA 6",
        "ALTO (cm.) PZA 6",
        "VOL PZA 6",
      ];

      /*
       * =========================================================
       * CONSTRUIR DATA PARA EXCEL
       * =========================================================
       */

      const datosExcel = todasLasFilas.map((medicion) => {
        const sku = String(medicion.sku || "").trim();

        const producto = productosMap.get(sku);
        const master = mastersMap.get(sku);

        /*
         * La fecha que se muestra en Excel será la fecha
         * correspondiente a Lima.
         */

        let fechaMostrar = "";

        if (medicion.fecha_hora) {
          fechaMostrar = new Intl.DateTimeFormat("es-PE", {
            timeZone: "America/Lima",
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
          }).format(new Date(medicion.fecha_hora));
        }

        return {
          SKU: sku,

          EAN: producto?.["EAN"] ?? medicion.ean ?? "",

          DESCRIPCIÓN: producto?.["DESCRIPCION"] ?? "",

          "FACTOR MED":
            medicion.factor_medicion ?? producto?.["FACTOR_WMS"] ?? "",

          FECHA: fechaMostrar,

          "ANCHO (cm.) MASTER": master?.ancho ?? "",

          "LARGO (cm.) MASTER": master?.largo ?? "",

          "ALTO (cm.) MASTER": master?.alto ?? "",

          "ANCHO (cm.) UND": medicion.ancho ?? "",

          "LARGO (cm.) UND": medicion.largo ?? "",

          "ALTO (cm.) UND": medicion.alto ?? "",

          "VOL PZA 1": medicion.pieza1_volumen ?? "",

          "ANCHO (cm.) PZA 2": medicion.pieza2_ancho ?? "",

          "LARGO (cm.) PZA 2": medicion.pieza2_largo ?? "",

          "ALTO (cm.) PZA 2": medicion.pieza2_alto ?? "",

          "VOL PZA 2": medicion.pieza2_volumen ?? "",

          "ANCHO (cm.) PZA 3": medicion.pieza3_ancho ?? "",

          "LARGO (cm.) PZA 3": medicion.pieza3_largo ?? "",

          "ALTO (cm.) PZA 3": medicion.pieza3_alto ?? "",

          "VOL PZA 3": medicion.pieza3_volumen ?? "",

          "ANCHO (cm.) PZA 4": medicion.pieza4_ancho ?? "",

          "LARGO (cm.) PZA 4": medicion.pieza4_largo ?? "",

          "ALTO (cm.) PZA 4": medicion.pieza4_alto ?? "",

          "VOL PZA 4": medicion.pieza4_volumen ?? "",

          "ANCHO (cm.) PZA 5": medicion.pieza5_ancho ?? "",

          "LARGO (cm.) PZA 5": medicion.pieza5_largo ?? "",

          "ALTO (cm.) PZA 5": medicion.pieza5_alto ?? "",

          "VOL PZA 5": medicion.pieza5_volumen ?? "",

          "ANCHO (cm.) PZA 6": medicion.pieza6_ancho ?? "",

          "LARGO (cm.) PZA 6": medicion.pieza6_largo ?? "",

          "ALTO (cm.) PZA 6": medicion.pieza6_alto ?? "",

          "VOL PZA 6": medicion.pieza6_volumen ?? "",
        };
      });

      /*
       * =========================================================
       * CREAR LIBRO EXCEL
       * =========================================================
       */

      const libro = XLSX.utils.book_new();

      const maxFilasPorHoja = 1048575;

      let numeroHoja = 1;

      for (
        let inicio = 0;
        inicio < datosExcel.length;
        inicio += maxFilasPorHoja
      ) {
        const bloque = datosExcel.slice(inicio, inicio + maxFilasPorHoja);

        const hoja = XLSX.utils.json_to_sheet(bloque, {
          header: encabezadosExcel,
        });

        XLSX.utils.book_append_sheet(libro, hoja, `DATA ${numeroHoja}`);

        numeroHoja++;
      }

      /*
       * =========================================================
       * NOMBRE DEL ARCHIVO
       * =========================================================
       */

      const partesFecha = fechaExportacion.split("-");

      const fechaArchivo =
        partesFecha.length === 3
          ? `${partesFecha[2]}-${partesFecha[1]}-${partesFecha[0]}`
          : fechaExportacion;

      /*
       * =========================================================
       * DESCARGAR
       * =========================================================
       */

      XLSX.writeFile(libro, `DATA_MAESTRO_${fechaArchivo}.xlsx`);

      const fechaMostrarMensaje = new Date(
        `${fechaExportacion}T12:00:00-05:00`
      ).toLocaleDateString("es-PE", {
        timeZone: "America/Lima",
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
      });

      setMensaje(
        `✓ DATA Maestro del ${fechaMostrarMensaje} exportado correctamente. ${datosExcel.length.toLocaleString(
          "es-PE"
        )} mediciones.`
      );
    } catch (error: any) {
      console.error("ERROR EXPORTANDO DATA MAESTRO:", error);

      setMensaje(
        `Error al exportar DATA Maestro: ${error?.message || String(error)}`
      );
    } finally {
      setExportando(false);
    }
  };

  /* =========================================================
     ENTRAR MEDICIÓN
  ========================================================= */

  const entrarAMedicion = (tipo: "PRODUCTO" | "CAJA") => {
    limpiarEstadosMedicion();

    setTipoMedicion(tipo);

    if (tipo === "PRODUCTO") {
      setPantalla("PRODUCTO");
    } else {
      setPantalla("CAJA");
    }

    setTimeout(() => {
      inputBusquedaRef.current?.focus();
    }, 150);
  };

  /* =========================================================
     VOLVER
  ========================================================= */

  const volverInicio = () => {
    setPantalla("INICIO");
    limpiarEstadosMedicion();
  };

  /* =========================================================
     CERRAR SESIÓN
  ========================================================= */

  const cerrarSesion = () => {
    setSesionActiva(false);

    setPantalla("INICIO");

    setAreaLogin("");
    setPersonalLogin(null);
    setDniLogin("");

    limpiarEstadosMedicion();

    setTimeout(() => {
      setMensajeLogin("Sesión cerrada correctamente.");
    }, 100);
  };

  /* =========================================================
     ÚLTIMA MEDICIÓN
  ========================================================= */

  const buscarUltimaMedicion = async (sku: string) => {
    setCargandoUltimaMedicion(true);

    setUltimaMedicion(null);

    try {
      const { data, error } = await supabase
        .from("MEDICIONES")
        .select(
          `
              id,
              sku,
              ean,
              volumen_unitario,
              volumen_master,
              fecha_hora,
              REGISTRADO_POR,
              AREA
            `
        )
        .eq("sku", sku)
        .order("fecha_hora", {
          ascending: false,
        })
        .limit(1);

      if (error) {
        setMensaje(
          "El producto fue encontrado, pero no se pudo consultar su última medición."
        );

        return null;
      }

      const ultima = data && data.length > 0 ? (data[0] as Medicion) : null;

      setUltimaMedicion(ultima);

      return ultima;
    } catch (error: any) {
      setMensaje(
        `No se pudo consultar la última medición: ${
          error?.message || String(error)
        }`
      );

      return null;
    } finally {
      setCargandoUltimaMedicion(false);
    }
  };

  /* =========================================================
     BUSCAR PRODUCTO
  ========================================================= */

  const buscarProducto = async (valorBusqueda?: string) => {
    try {
      const valor = String(
        valorBusqueda !== undefined ? valorBusqueda : busqueda
      ).trim();

      if (!valor) {
        setMensaje("Ingrese un SKU o EAN.");
        return;
      }

      setCargando(true);
      setMensaje("");
      setProducto(null);
      setUltimaMedicion(null);

      setLargo("");
      setAncho("");
      setAlto("");

      setEsMultipieza(false);
      setCantidadPiezas(2);

      setPiezas(
        Array.from({ length: 6 }, () => ({
          largo: "",
          ancho: "",
          alto: "",
        }))
      );

      setFactorMedicion("");

      const respuestaSKU = await supabase
        .from("PRODUCTOS")
        .select("SKU,EAN,DESCRIPCION,UNIDAD_MEDIDA,FACTOR_WMS")
        .eq("SKU", valor)
        .limit(1);

      if (respuestaSKU.error) {
        setMensaje(`Error al buscar SKU: ${respuestaSKU.error.message}`);

        return;
      }

      if (respuestaSKU.data && respuestaSKU.data.length > 0) {
        const encontrado = respuestaSKU.data[0] as Producto;

        setProducto(encontrado);

        setFactorMedicion(
          encontrado.FACTOR_WMS ? String(encontrado.FACTOR_WMS) : ""
        );

        await buscarUltimaMedicion(encontrado.SKU);

        return;
      }

      const respuestaEAN = await supabase
        .from("PRODUCTOS")
        .select("SKU,EAN,DESCRIPCION,UNIDAD_MEDIDA,FACTOR_WMS")
        .eq("EAN", valor)
        .limit(1);

      if (respuestaEAN.error) {
        setMensaje(`Error al buscar EAN: ${respuestaEAN.error.message}`);

        return;
      }

      if (respuestaEAN.data && respuestaEAN.data.length > 0) {
        const encontrado = respuestaEAN.data[0] as Producto;

        setProducto(encontrado);

        setFactorMedicion(
          encontrado.FACTOR_WMS ? String(encontrado.FACTOR_WMS) : ""
        );

        await buscarUltimaMedicion(encontrado.SKU);

        return;
      }

      setMensaje("No se encontró ningún producto con ese SKU o EAN.");
    } catch (error: any) {
      setMensaje(
        `Error al buscar el producto: ${error?.message || String(error)}`
      );
    } finally {
      setCargando(false);
    }
  };

  /* =========================================================
     BÚSQUEDA EAN
  ========================================================= */

  const manejarBusqueda = (valor: string) => {
    setBusqueda(valor);

    if (temporizadorEAN.current) {
      clearTimeout(temporizadorEAN.current);
    }

    const valorLimpio = valor.trim();

    if (/^\d{8,14}$/.test(valorLimpio)) {
      temporizadorEAN.current = setTimeout(() => {
        buscarProducto(valorLimpio);
      }, 300);
    }
  };

  /* =========================================================
     PIEZAS
  ========================================================= */

  const actualizarPieza = (
    indice: number,
    campo: keyof Pieza,
    valor: string
  ) => {
    setPiezas((actuales) =>
      actuales.map((pieza, i) =>
        i === indice
          ? {
              ...pieza,
              [campo]: valor,
            }
          : pieza
      )
    );
  };

  const cambiarTipoProducto = (multipieza: boolean) => {
    setEsMultipieza(multipieza);

    if (!multipieza) {
      setCantidadPiezas(2);
    }
  };

  const calcularVolumenPieza = (pieza: Pieza) => {
    return (
      (Number(pieza.largo || 0) *
        Number(pieza.ancho || 0) *
        Number(pieza.alto || 0)) /
      1000000
    );
  };

  const volumenMasterUnaPieza =
    (Number(largo || 0) * Number(ancho || 0) * Number(alto || 0)) / 1000000;

  const volumenMasterMultipieza = piezas
    .slice(0, cantidadPiezas)
    .reduce((total, pieza) => total + calcularVolumenPieza(pieza), 0);

  const volumenMaster = esMultipieza
    ? volumenMasterMultipieza
    : volumenMasterUnaPieza;

  const volumenUnitario =
    Number(factorMedicion || 0) > 0
      ? volumenMaster / Number(factorMedicion)
      : 0;

  /* =========================================================
     CAJA
  ========================================================= */

  const volumenCaja =
    (Number(cajaLargo || 0) * Number(cajaAncho || 0) * Number(cajaAlto || 0)) /
    1000000;

  const volumenProductosCaja = itemsCaja.reduce(
    (total, item) => total + item.volumenTotal,
    0
  );

  const cantidadTotalCaja = itemsCaja.reduce(
    (total, item) => total + item.cantidad,
    0
  );

  const diferenciaVolumenCaja = volumenCaja - volumenProductosCaja;

  const porcentajeOcupacionCaja =
    volumenCaja > 0 ? (volumenProductosCaja / volumenCaja) * 100 : 0;

  /* =========================================================
     AGREGAR PRODUCTO A CAJA
  ========================================================= */

  const agregarProductoACaja = () => {
    if (!producto) {
      setMensaje("Primero debe seleccionar un producto.");

      return;
    }

    const cantidad = Number(cantidadCaja || 0);

    if (!Number.isInteger(cantidad) || cantidad <= 0) {
      setMensaje("Ingrese una cantidad válida de unidades.");

      return;
    }

    const volumenUnitarioProducto = Number(
      ultimaMedicion?.volumen_unitario || 0
    );

    if (volumenUnitarioProducto <= 0) {
      setMensaje(
        "Este SKU no tiene una medición individual registrada. Primero debe medirse el producto."
      );

      return;
    }

    const existe = itemsCaja.findIndex((item) => item.sku === producto.SKU);

    if (existe >= 0) {
      setItemsCaja((actuales) =>
        actuales.map((item, indice) =>
          indice === existe
            ? {
                ...item,
                cantidad: item.cantidad + cantidad,
                volumenTotal: (item.cantidad + cantidad) * item.volumenUnitario,
              }
            : item
        )
      );
    } else {
      setItemsCaja((actuales) => [
        ...actuales,
        {
          sku: producto.SKU,
          ean: producto.EAN,
          descripcion: producto.DESCRIPCION,
          cantidad,
          volumenUnitario: volumenUnitarioProducto,
          volumenTotal: cantidad * volumenUnitarioProducto,
        },
      ]);
    }

    setCantidadCaja("1");
    setCajaConfirmada(false);
  };

  const eliminarItemCaja = (sku: string) => {
    setItemsCaja((actuales) => actuales.filter((item) => item.sku !== sku));

    setCajaConfirmada(false);
  };

  const limpiarCaja = () => {
    setCajaLargo("");
    setCajaAncho("");
    setCajaAlto("");
    setCantidadCaja("1");
    setItemsCaja([]);
    setMostrarRevisionCaja(false);
    setCajaConfirmada(false);
    setMensaje("");
  };

  const revisarCaja = () => {
    setMensaje("");

    if (
      Number(cajaLargo || 0) <= 0 ||
      Number(cajaAncho || 0) <= 0 ||
      Number(cajaAlto || 0) <= 0
    ) {
      setMensaje("Complete las dimensiones de la caja antes de revisar.");

      return;
    }

    if (itemsCaja.length === 0) {
      setMensaje("Agregue al menos un SKU a la caja antes de revisar.");

      return;
    }

    setCajaConfirmada(false);
    setMostrarRevisionCaja(true);
  };

  /* =========================================================
     CONFIRMAR CAJA
  ========================================================= */

  const confirmarCaja = async () => {
    if (!personalLogin) {
      setMensaje(
        "No se encontró la información del usuario para guardar la caja."
      );

      return;
    }

    if (itemsCaja.length === 0) {
      setMensaje("La caja debe tener al menos un SKU.");

      return;
    }

    if (
      Number(cajaLargo || 0) <= 0 ||
      Number(cajaAncho || 0) <= 0 ||
      Number(cajaAlto || 0) <= 0
    ) {
      setMensaje("Las dimensiones de la caja deben ser mayores a cero.");

      return;
    }

    setGuardando(true);
    setMensaje("");

    try {
      const fechaHora = obtenerFechaHoraPeru();

      const datosMaster = {
        largo: Number(cajaLargo),
        ancho: Number(cajaAncho),
        alto: Number(cajaAlto),
        volumen_caja: volumenCaja,
        cantidad_total: cantidadTotalCaja,
        volumen_productos: volumenProductosCaja,
        espacio_adicional: diferenciaVolumenCaja,
        porcentaje_ocupacion: porcentajeOcupacionCaja,
        registrado_por: personalLogin.NOMBRE || "",
        area: personalLogin.AREA || "",
        fecha_hora: fechaHora,
      };

      const { data: cajaMaster, error: errorMaster } = await supabase
        .from("CAJAS_MASTER")
        .insert(datosMaster)
        .select("*")
        .single();

      if (errorMaster) {
        throw new Error(
          `Error al guardar CAJAS_MASTER: ${errorMaster.message}`
        );
      }

      if (!cajaMaster || cajaMaster.id == null) {
        throw new Error("La caja se insertó, pero no se pudo obtener el ID.");
      }

      const datosDetalle = itemsCaja.map((item) => ({
        caja_master_id: cajaMaster.id,
        sku: item.sku,
        ean: item.ean,
        cantidad: item.cantidad,
        volumen_unitario: item.volumenUnitario,
        volumen_total: item.volumenTotal,
        registrado_por: personalLogin.NOMBRE || "",
        area: personalLogin.AREA || "",
        fecha_hora: fechaHora,
      }));

      const { error: errorDetalle } = await supabase
        .from("CAJAS_MASTER_DETALLE")
        .insert(datosDetalle);

      if (errorDetalle) {
        throw new Error(
          `La cabecera se guardó, pero ocurrió un error al guardar el detalle: ${errorDetalle.message}`
        );
      }

      setCajaLargo("");
      setCajaAncho("");
      setCajaAlto("");
      setCantidadCaja("1");
      setItemsCaja([]);
      setMostrarRevisionCaja(false);
      setCajaConfirmada(false);

      setMensaje(
        `✓ Caja guardada correctamente. Se registraron ${itemsCaja.length} SKU y ${cantidadTotalCaja} unidades.`
      );
    } catch (error: unknown) {
      const mensajeError =
        error instanceof Error
          ? error.message
          : "Ocurrió un error desconocido al guardar la caja.";

      setMensaje(mensajeError);
    } finally {
      setGuardando(false);
    }
  };

  /* =========================================================
     GUARDAR MEDICIÓN CONFIRMADA
  ========================================================= */

  const guardarMedicionConfirmada = async (
    aceptaVolumenMenor: boolean,
    volumenAnterior: number | null
  ) => {
    if (!producto || !personalLogin) {
      setMensaje("No se encontró la información necesaria para guardar.");

      return;
    }

    setGuardando(true);
    setMensaje("");
    setMostrarAdvertencia(false);

    const piezasSeleccionadas = esMultipieza
      ? piezas.slice(0, cantidadPiezas)
      : [];

    const fechaHora = obtenerFechaHoraPeru();

    const datosGuardar: Record<string, any> = {
      sku: producto.SKU,
      ean: producto.EAN,

      largo: esMultipieza
        ? Number(piezasSeleccionadas[0]?.largo || 0)
        : Number(largo),

      ancho: esMultipieza
        ? Number(piezasSeleccionadas[0]?.ancho || 0)
        : Number(ancho),

      alto: esMultipieza
        ? Number(piezasSeleccionadas[0]?.alto || 0)
        : Number(alto),

      factor_medicion: Number(factorMedicion),

      volumen_unitario: volumenUnitario,

      volumen_master: volumenMaster,

      VOLUMEN_ANTERIOR: volumenAnterior,

      ACEPTO_VOL_MENOR: aceptaVolumenMenor,

      cantidad_piezas: esMultipieza ? cantidadPiezas : 1,

      REGISTRADO_POR: personalLogin.NOMBRE,

      AREA: personalLogin.AREA,
    };

    if (esMultipieza) {
      piezasSeleccionadas.forEach((pieza, indice) => {
        const numero = indice + 1;

        datosGuardar[`pieza${numero}_largo`] = Number(pieza.largo);

        datosGuardar[`pieza${numero}_ancho`] = Number(pieza.ancho);

        datosGuardar[`pieza${numero}_alto`] = Number(pieza.alto);

        datosGuardar[`pieza${numero}_volumen`] = calcularVolumenPieza(pieza);
      });
    }

    try {
      const { data: medicionGuardada, error } = await supabase
        .from("MEDICIONES")
        .insert(datosGuardar)
        .select()
        .single();

      if (error) {
        setMensaje("No se pudo guardar la medición: " + error.message);

        return;
      }

      setMensaje(
        aceptaVolumenMenor
          ? "Medición guardada. Se aceptó registrar un volumen unitario menor."
          : "Medición guardada correctamente."
      );

      setLargo("");
      setAncho("");
      setAlto("");

      setEsMultipieza(false);
      setCantidadPiezas(2);

      setPiezas(
        Array.from({ length: 6 }, () => ({
          largo: "",
          ancho: "",
          alto: "",
        }))
      );

      setFactorMedicion(producto.FACTOR_WMS || "");

      setUltimaMedicion({
        id: medicionGuardada?.id || 0,
        sku: producto.SKU,
        ean: producto.EAN,
        volumen_unitario: volumenUnitario,
        volumen_master: volumenMaster,
        fecha_hora: medicionGuardada?.fecha_hora || null,
        REGISTRADO_POR: personalLogin.NOMBRE,
        AREA: personalLogin.AREA,
      });

      setVolumenAnteriorAdvertencia(null);
    } catch (error: any) {
      setMensaje(
        `Error al guardar la medición: ${error?.message || String(error)}`
      );
    } finally {
      setGuardando(false);

      setTimeout(() => {
        inputBusquedaRef.current?.focus();
      }, 100);
    }
  };

  /* =========================================================
     GUARDAR MEDICIÓN
  ========================================================= */

  const guardarMedicion = async () => {
    if (!sesionActiva) {
      setMensaje("Debe iniciar sesión antes de registrar una medición.");

      return;
    }

    if (!producto) {
      setMensaje("Primero debe seleccionar un producto.");

      return;
    }

    if (!personalLogin) {
      setMensaje("No se encontró el usuario de la sesión.");

      return;
    }

    if (Number(factorMedicion || 0) <= 0) {
      setMensaje("Ingrese un factor de medición válido.");

      return;
    }

    if (!esMultipieza) {
      if (
        Number(largo || 0) <= 0 ||
        Number(ancho || 0) <= 0 ||
        Number(alto || 0) <= 0
      ) {
        setMensaje("Ingrese largo, ancho y alto.");

        return;
      }
    } else {
      const piezasValidas = piezas
        .slice(0, cantidadPiezas)
        .every(
          (pieza) =>
            Number(pieza.largo || 0) > 0 &&
            Number(pieza.ancho || 0) > 0 &&
            Number(pieza.alto || 0) > 0
        );

      if (!piezasValidas) {
        setMensaje("Complete las dimensiones de todas las piezas.");

        return;
      }
    }

    const volumenAnterior = ultimaMedicion?.volumen_unitario;

    if (
      volumenAnterior !== null &&
      volumenAnterior !== undefined &&
      Number(volumenAnterior) > 0 &&
      volumenUnitario < Number(volumenAnterior)
    ) {
      setVolumenAnteriorAdvertencia(Number(volumenAnterior));

      setMostrarAdvertencia(true);

      return;
    }

    await guardarMedicionConfirmada(
      false,
      volumenAnterior !== null && volumenAnterior !== undefined
        ? Number(volumenAnterior)
        : null
    );
  };

  const aceptarVolumenMenor = async () => {
    await guardarMedicionConfirmada(true, volumenAnteriorAdvertencia);
  };

  const cancelarVolumenMenor = () => {
    setMostrarAdvertencia(false);

    setVolumenAnteriorAdvertencia(null);

    setMensaje(
      "Medición no guardada. Verifique el volumen unitario antes de continuar."
    );
  };

  const limpiar = () => {
    limpiarEstadosMedicion();

    setTimeout(() => {
      inputBusquedaRef.current?.focus();
    }, 100);
  };

  /* =========================================================
     LOGIN
  ========================================================= */

  if (!sesionActiva) {
    return (
      <div className="app">
        <div className="contenedor">
          <div className="medicion">
            <h1>Medición de Productos</h1>

            <h2>Inicio de sesión</h2>

            <div className="campo">
              <label>Área</label>

              <select
                value={areaLogin}
                onChange={(e) => cambiarAreaLogin(e.target.value)}
                disabled={cargandoPersonal}
              >
                <option value="">
                  {cargandoPersonal
                    ? "Cargando áreas..."
                    : "Seleccione su área"}
                </option>

                {areas.map((area) => (
                  <option key={area} value={area}>
                    {area}
                  </option>
                ))}
              </select>
            </div>

            <div className="campo">
              <label>Nombre</label>

              <select
                value={personalLogin?.[COLUMNA_DNI] || ""}
                onChange={(e) => cambiarPersonalLogin(e.target.value)}
                disabled={!areaLogin || cargandoPersonal}
              >
                <option value="">
                  {areaLogin
                    ? "Seleccione su nombre"
                    : "Primero seleccione un área"}
                </option>

                {personalFiltradoLogin.map((persona) => (
                  <option
                    key={persona[COLUMNA_DNI] || persona.NOMBRE || ""}
                    value={persona[COLUMNA_DNI] || ""}
                  >
                    {persona.NOMBRE}
                  </option>
                ))}
              </select>
            </div>

            <div className="campo">
              <label>DNI / CE</label>

              <input
                ref={inputDniRef}
                type="password"
                inputMode="numeric"
                value={dniLogin}
                onChange={(e) => setDniLogin(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    iniciarSesion();
                  }
                }}
                placeholder="Ingrese su DNI / CE"
                disabled={!personalLogin || ingresando}
              />
            </div>

            {mensajeLogin && (
              <div
                className="mensaje"
                style={{
                  whiteSpace: "pre-line",
                }}
              >
                {mensajeLogin}
              </div>
            )}

            <button
              className="guardar"
              onClick={iniciarSesion}
              disabled={ingresando || cargandoPersonal}
            >
              {ingresando ? "Validando..." : "Ingresar"}
            </button>
          </div>
        </div>
      </div>
    );
  }

  /* =========================================================
     INICIO
  ========================================================= */

  if (pantalla === "INICIO") {
    return (
      <div className="app">
        <div className="contenedor">
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              gap: "15px",
              flexWrap: "wrap",
              marginBottom: "25px",
            }}
          >
            <div>
              <h1
                style={{
                  marginBottom: "6px",
                }}
              >
                Medición de Productos
              </h1>

              <div
                style={{
                  fontSize: "14px",
                  opacity: 0.8,
                }}
              >
                Centro de medición y volumetría
              </div>
            </div>

            <button onClick={cerrarSesion}>Cerrar sesión</button>
          </div>

          <div
            className="producto"
            style={{
              marginBottom: "25px",
              textAlign: "center",
              padding: "30px 20px",
            }}
          >
            <div
              style={{
                fontSize: "46px",
                marginBottom: "10px",
              }}
            >
              👋
            </div>

            <h2
              style={{
                marginBottom: "8px",
              }}
            >
              Bienvenido
            </h2>

            <div
              style={{
                fontSize: "20px",
                fontWeight: 700,
                marginBottom: "6px",
              }}
            >
              {personalLogin?.NOMBRE || "-"}
            </div>

            <div
              style={{
                fontSize: "15px",
                opacity: 0.75,
              }}
            >
              Área: {personalLogin?.AREA || "-"}
            </div>
          </div>

          {/* =====================================================
              OPCIONES DE MEDICIÓN
              IMPORTANTE: estos botones son hermanos.
              No colocar un <button> dentro de otro <button>.
          ===================================================== */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
              gap: "20px",
            }}
          >
            {/* MEDIR PRODUCTO: un SKU individual, unidad y/o master */}
            <button
              type="button"
              onClick={() => entrarAMedicion("PRODUCTO")}
              style={{
                minHeight: "220px",
                borderRadius: "18px",
                padding: "25px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                alignItems: "center",
                gap: "12px",
                cursor: "pointer",
                border: "2px solid transparent",
                background: "linear-gradient(135deg, #eaf3ff, #ffffff)",
                boxShadow: "0 8px 25px rgba(0,0,0,0.08)",
              }}
            >
              <div style={{ fontSize: "58px", lineHeight: 1 }}>📏</div>

              <div style={{ fontSize: "22px", fontWeight: 800 }}>
                Medir producto
              </div>

              <div
                style={{
                  fontSize: "14px",
                  opacity: 0.7,
                  textAlign: "center",
                  lineHeight: 1.5,
                }}
              >
                Registre las mediciones del producto
                <br />
                <strong>por unidad y/o master</strong>.
                <br />
                Aplica para un SKU individual.
              </div>

              <div style={{ marginTop: "5px", fontWeight: 700 }}>
                Iniciar medición →
              </div>
            </button>

            {/* MEDIR CAJA MASTER: caja física que contiene varios SKU */}
            <button
              type="button"
              onClick={() => entrarAMedicion("CAJA")}
              style={{
                minHeight: "220px",
                borderRadius: "18px",
                padding: "25px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                alignItems: "center",
                gap: "12px",
                cursor: "pointer",
                border: "2px solid transparent",
                background: "linear-gradient(135deg, #fff7e8, #ffffff)",
                boxShadow: "0 8px 25px rgba(0,0,0,0.08)",
              }}
            >
              <div style={{ fontSize: "58px", lineHeight: 1 }}>📦</div>

              <div style={{ fontSize: "22px", fontWeight: 800 }}>
                Medir caja master
              </div>

              <div
                style={{
                  fontSize: "14px",
                  opacity: 0.7,
                  textAlign: "center",
                  lineHeight: 1.5,
                }}
              >
                Para una caja que contiene
                <br />
                <strong>varios SKU</strong>.
                <br />
                Mida la caja y registre los productos que contiene.
              </div>

              <div style={{ marginTop: "5px", fontWeight: 700 }}>
                Iniciar medición →
              </div>
            </button>

            <div
              style={{
                minHeight: "220px",
                borderRadius: "18px",
                padding: "25px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                alignItems: "center",
                gap: "12px",
                border: "2px solid transparent",
                background: "linear-gradient(135deg, #eafaf0, #ffffff)",
                boxShadow: "0 8px 25px rgba(0,0,0,0.08)",
              }}
            >
              <div
                style={{
                  fontSize: "58px",
                  lineHeight: 1,
                }}
              >
                📊
              </div>

              <div
                style={{
                  fontSize: "22px",
                  fontWeight: 800,
                }}
              >
                Exportar DATA Maestro
              </div>

              <div
                style={{
                  fontSize: "14px",
                  opacity: 0.7,
                  textAlign: "center",
                }}
              >
                Seleccione la fecha de las mediciones que desea descargar
              </div>

              <input
                type="date"
                value={fechaExportacion}
                onChange={(e) => setFechaExportacion(e.target.value)}
                disabled={exportando}
                style={{
                  padding: "10px 14px",
                  borderRadius: "10px",
                  border: "1px solid #ccc",
                  fontSize: "16px",
                  fontWeight: 600,
                  cursor: exportando ? "not-allowed" : "pointer",
                  background: "#ffffff",
                }}
              />

              <button
                onClick={exportarDataMaestro}
                disabled={exportando || !fechaExportacion}
                style={{
                  marginTop: "5px",
                  padding: "12px 22px",
                  borderRadius: "10px",
                  border: "none",
                  fontWeight: 700,
                  fontSize: "15px",
                  cursor: exportando || !fechaExportacion ? "wait" : "pointer",
                  background:
                    exportando || !fechaExportacion ? "#ccc" : "#198754",
                  color: "#ffffff",
                }}
              >
                {exportando ? "Preparando archivo..." : "Descargar Excel →"}
              </button>
            </div>
            <button
              onClick={entrarAuditoriaPallet}
              style={{
                minHeight: "220px",
                borderRadius: "18px",
                padding: "25px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                alignItems: "center",
                gap: "12px",
                cursor: "pointer",
                border: "2px solid transparent",
                background: "linear-gradient(135deg, #fff0f0, #ffffff)",
                boxShadow: "0 8px 25px rgba(0,0,0,0.08)",
              }}
            >
              <div
                style={{
                  fontSize: "58px",
                  lineHeight: 1,
                }}
              >
                📦
              </div>

              <div
                style={{
                  fontSize: "22px",
                  fontWeight: 800,
                }}
              >
                Auditoría de Vol. de Pallets
              </div>

              <div
                style={{
                  fontSize: "14px",
                  opacity: 0.7,
                  textAlign: "center",
                }}
              >
                Mida y registre el volumen físico real de los pallets
              </div>

              <div
                style={{
                  marginTop: "5px",
                  fontWeight: 700,
                }}
              >
                Iniciar auditoría →
              </div>
            </button>
            <button
              onClick={seleccionarArchivoPersonal}
              disabled={actualizandoPersonal}
              style={{
                minHeight: "220px",
                borderRadius: "18px",
                padding: "25px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                alignItems: "center",
                gap: "12px",
                cursor: actualizandoPersonal ? "wait" : "pointer",
                border: "2px solid transparent",
                background: "linear-gradient(135deg, #f2eaff, #ffffff)",
                boxShadow: "0 8px 25px rgba(0,0,0,0.08)",
                opacity: actualizandoPersonal ? 0.7 : 1,
              }}
            >
              <div
                style={{
                  fontSize: "58px",
                  lineHeight: 1,
                }}
              >
                {actualizandoPersonal ? "⏳" : "👥"}
              </div>

              <div
                style={{
                  fontSize: "22px",
                  fontWeight: 800,
                }}
              >
                {actualizandoPersonal
                  ? "Actualizando..."
                  : "Actualizar personal"}
              </div>

              <div
                style={{
                  fontSize: "14px",
                  opacity: 0.7,
                  textAlign: "center",
                }}
              >
                Cargue el CSV o Excel actualizado de la dotación
              </div>

              <div
                style={{
                  marginTop: "5px",
                  fontWeight: 700,
                }}
              >
                {actualizandoPersonal
                  ? "Procesando información..."
                  : "Seleccionar archivo →"}
              </div>
            </button>

            <input
              ref={inputPersonalRef}
              type="file"
              accept=".csv,.xlsx,.xls"
              style={{
                display: "none",
              }}
              onChange={manejarArchivoPersonal}
            />
          </div>

          {mensaje && (
            <div
              className="mensaje"
              style={{
                marginTop: "25px",
                textAlign: "center",
                padding: "16px",
                whiteSpace: "pre-line",
              }}
            >
              {mensaje}
            </div>
          )}

          {!mensaje && (
            <div
              className="mensaje"
              style={{
                marginTop: "25px",
                textAlign: "center",
                padding: "16px",
              }}
            >
              Seleccione una opción para comenzar su jornada de medición.
            </div>
          )}
        </div>
      </div>
    );
  }
  /* =========================================================
   PANTALLA AUDITORÍA DE VOL DE PALLETS
========================================================= */

  /* =========================================================
   PANTALLA AUDITORÍA DE VOL DE PALLETS
========================================================= */

  if (pantalla === "AUDITORIA_PALLET") {
    return (
      <div className="app">
        <div className="contenedor">
          {/* =====================================================
            ENCABEZADO
        ===================================================== */}

          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              gap: "12px",
              flexWrap: "wrap",
              marginBottom: "20px",
            }}
          >
            <div>
              <h1>Auditoría de Vol. de Pallets</h1>

              <div
                className="mensaje"
                style={{
                  marginTop: "8px",
                }}
              >
                <strong>Usuario:</strong> {personalLogin?.NOMBRE || "-"} —{" "}
                {personalLogin?.AREA || "-"}
              </div>
            </div>

            <div
              style={{
                display: "flex",
                gap: "8px",
                flexWrap: "wrap",
              }}
            >
              <button onClick={volverInicio}>← Inicio</button>

              <button onClick={cerrarSesion}>Cerrar sesión</button>
            </div>
          </div>

          {/* =====================================================
            1. REGISTRAR AUDITORÍA DEL PALLET
        ===================================================== */}

          {!auditoriaPalletId && (
            <div
              className="producto"
              style={{
                marginBottom: "20px",
              }}
            >
              <h2>① Datos del pallet</h2>
              <div
                className="mensaje"
                style={{
                  marginBottom: "20px",
                }}
              >
                Registre primero los datos y las dimensiones físicas del pallet.
                Luego podrá agregar las cajas y los SKU que contiene.
              </div>

              <div className="campo">
                <label>Código del pallet</label>

                <input
                  type="text"
                  value={codigoPallet}
                  onChange={(e) =>
                    setCodigoPallet(e.target.value.toUpperCase())
                  }
                  placeholder="Ej. PL954002839877"
                />
              </div>

              <div className="campo">
                <label>Tienda</label>

                <input
                  type="text"
                  value={proyectoPallet}
                  onChange={(e) => setProyectoPallet(e.target.value)}
                  placeholder="Ej. Cusco / Piura"
                />
              </div>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
                  gap: "15px",
                }}
              >
                <div className="campo">
                  <label>Ancho (cm)</label>

                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    value={palletAncho}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setPalletAncho(e.target.value)}
                    placeholder="Ej. 100"
                  />
                </div>

                <div className="campo">
                  <label>Alto (cm)</label>

                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    value={palletAlto}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setPalletAlto(e.target.value)}
                    placeholder="Ej. 180"
                  />
                </div>

                <div className="campo">
                  <label>Largo (cm)</label>

                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    value={palletLargo}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setPalletLargo(e.target.value)}
                    placeholder="Ej. 120"
                  />
                </div>
              </div>

              {/* VOLUMEN FÍSICO DEL PALLET */}

              <div
                className="volumen"
                style={{
                  marginTop: "20px",
                }}
              >
                <div>
                  <strong>Volumen físico del pallet</strong>

                  <span>{volumenFisicoPallet.toFixed(4)} m³</span>
                </div>
              </div>

              {/* BOTONES */}

              <div
                style={{
                  display: "flex",
                  gap: "10px",
                  flexWrap: "wrap",
                  marginTop: "20px",
                }}
              >
                <button
                  onClick={limpiarAuditoriaPallet}
                  disabled={guardandoAuditoria}
                >
                  Limpiar
                </button>

                <button
                  className="guardar"
                  onClick={guardarAuditoriaPallet}
                  disabled={guardandoAuditoria}
                >
                  {guardandoAuditoria ? "Registrando..." : "✓ Registrar pallet"}
                </button>
              </div>
            </div>
          )}

          {/* =====================================================
    2. CONTENIDO DEL PALLET
===================================================== */}

          {auditoriaPalletId && (
            <div
              className="producto"
              style={{
                marginBottom: "20px",
              }}
            >
              <h2>② Contenido del pallet</h2>

              <div
                className="mensaje"
                style={{
                  marginBottom: "20px",
                  border: "1px solid #d9d9d9",
                }}
              >
                <div>
                  <strong>Pallet en auditoría:</strong> {codigoPallet || "-"}
                </div>
                <div>
                  <strong>Tienda:</strong> {proyectoPallet || "-"}
                </div>
              </div>

              {/* =================================================
        MENÚ DE CONTENIDO
    ================================================= */}

              {modoContenidoPallet === "MENU" && (
                <>
                  <div
                    className="mensaje"
                    style={{
                      marginBottom: "20px",
                    }}
                  >
                    Seleccione qué desea registrar. Puede registrar cajas y SKU
                    sueltos en cualquier orden.
                  </div>

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns:
                        "repeat(auto-fit, minmax(220px, 1fr))",
                      gap: "15px",
                    }}
                  >
                    <button className="guardar" onClick={abrirAgregarCaja}>
                      📦 Agregar caja
                    </button>

                    <button
                      onClick={() => {
                        setModoContenidoPallet("SKU_SUELTO");
                        setMensaje("");
                      }}
                    >
                      🏷️ Agregar SKU suelto
                    </button>
                  </div>
                </>
              )}

              {/* =================================================
        FLUJO DE CAJA
    ================================================= */}

              {modoContenidoPallet === "CAJA" && (
                <>
                  <div
                    className="mensaje"
                    style={{
                      marginBottom: "20px",
                    }}
                  >
                    Registre las dimensiones físicas de la caja. Después podrá
                    agregar los SKU que contiene.
                  </div>

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns:
                        "repeat(auto-fit, minmax(180px, 1fr))",
                      gap: "15px",
                    }}
                  >
                    <div className="campo">
                      <label>Ancho (cm)</label>

                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        value={cajaAncho}
                        onWheel={(e) => e.currentTarget.blur()}
                        onChange={(e) => setCajaAncho(e.target.value)}
                        placeholder="Ej. 60"
                      />
                    </div>

                    <div className="campo">
                      <label>Alto (cm)</label>

                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        value={cajaAlto}
                        onWheel={(e) => e.currentTarget.blur()}
                        onChange={(e) => setCajaAlto(e.target.value)}
                        placeholder="Ej. 80"
                      />
                    </div>

                    <div className="campo">
                      <label>Largo (cm)</label>

                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        value={cajaLargo}
                        onWheel={(e) => e.currentTarget.blur()}
                        onChange={(e) => setCajaLargo(e.target.value)}
                        placeholder="Ej. 100"
                      />
                    </div>
                  </div>

                  {/* VOLUMEN DE LA CAJA */}

                  <div
                    className="volumen"
                    style={{
                      marginTop: "20px",
                    }}
                  >
                    <div>
                      <strong>Volumen de la caja</strong>

                      <span>{volumenCaja.toFixed(4)} m³</span>
                    </div>
                  </div>

                  {/* BOTONES */}

                  <div
                    style={{
                      display: "flex",
                      gap: "10px",
                      flexWrap: "wrap",
                      marginTop: "20px",
                    }}
                  >
                    <button
                      onClick={volverMenuContenido}
                      disabled={agregandoCaja}
                    >
                      ← Volver
                    </button>

                    <button
                      className="guardar"
                      onClick={agregarCajaAuditoria}
                      disabled={agregandoCaja}
                    >
                      {agregandoCaja ? "Registrando..." : "✓ Registrar caja"}
                    </button>
                  </div>

                  {/* CAJAS REGISTRADAS */}

                  {cajasAuditoria.length > 0 && (
                    <div
                      className="tabla-historial"
                      style={{
                        marginTop: "25px",
                      }}
                    >
                      <h3>Cajas registradas</h3>

                      <table>
                        <thead>
                          <tr>
                            <th>Caja</th>
                            <th>Ancho</th>
                            <th>Alto</th>
                            <th>Largo</th>
                            <th>Volumen</th>
                            <th>Contenido</th>
                          </tr>
                        </thead>

                        <tbody>
                          {cajasAuditoria.map((caja) => (
                            <tr key={caja.id}>
                              <td>
                                <strong>Caja {caja.numero_caja}</strong>
                              </td>

                              <td>{Number(caja.ancho).toFixed(2)} cm</td>

                              <td>{Number(caja.alto).toFixed(2)} cm</td>

                              <td>{Number(caja.largo).toFixed(2)} cm</td>

                              <td>
                                <strong>
                                  {Number(caja.volumen_caja || 0).toFixed(4)} m³
                                </strong>
                              </td>

                              <td>
                                <button
                                  onClick={() => {
                                    setCajaSeleccionadaAuditoria(caja);
                                    setMensaje("");
                                  }}
                                >
                                  📦 Ver contenido
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                  {cajaSeleccionadaAuditoria && (
                    <div
                      className="producto"
                      style={{
                        marginTop: "25px",
                        border: "2px solid #ddd",
                      }}
                    >
                      <h3>
                        📦 Contenido de Caja{" "}
                        {cajaSeleccionadaAuditoria.numero_caja}
                      </h3>

                      <div
                        className="mensaje"
                        style={{
                          marginBottom: "20px",
                        }}
                      >
                        <strong>Dimensiones de la caja:</strong>{" "}
                        {Number(cajaSeleccionadaAuditoria.ancho).toFixed(2)} ×{" "}
                        {Number(cajaSeleccionadaAuditoria.alto).toFixed(2)} ×{" "}
                        {Number(cajaSeleccionadaAuditoria.largo).toFixed(2)} cm
                        <br />
                        <strong>Volumen:</strong>{" "}
                        {Number(
                          cajaSeleccionadaAuditoria.volumen_caja || 0
                        ).toFixed(4)}{" "}
                        m³
                      </div>

                      {/* AGREGAR SKU A LA CAJA */}

                      <div
                        style={{
                          marginTop: "25px",
                          paddingTop: "20px",
                          borderTop: "1px solid #ddd",
                        }}
                      >
                        <h4>➕ Agregar SKU a esta caja</h4>

                        <div className="campo">
                          <label>SKU o EAN</label>

                          <div
                            style={{
                              display: "flex",
                              gap: "10px",
                              flexWrap: "wrap",
                            }}
                          >
                            <input
                              type="text"
                              value={skuCajaBusqueda}
                              onChange={(e) =>
                                setSkuCajaBusqueda(e.target.value)
                              }
                              onKeyDown={(e) => {
                                if (e.key === "Enter") {
                                  e.preventDefault();
                                  buscarProductoSkuCaja();
                                }
                              }}
                              placeholder="Escanee o ingrese SKU / EAN"
                              style={{
                                flex: 1,
                                minWidth: "220px",
                              }}
                            />

                            <button
                              type="button"
                              className="guardar"
                              onClick={(e) => {
                                e.preventDefault();
                                buscarProductoSkuCaja();
                              }}
                              disabled={buscandoSkuCaja}
                            >
                              {buscandoSkuCaja ? "Buscando..." : "🔎 Buscar"}
                            </button>
                          </div>
                        </div>

                        {/* PRODUCTO ENCONTRADO */}

                        {skuCajaProducto && (
                          <div
                            className="mensaje"
                            style={{
                              marginTop: "20px",
                            }}
                          >
                            <div>
                              <strong>SKU:</strong> {skuCajaProducto.SKU}
                            </div>

                            <div>
                              <strong>EAN:</strong> {skuCajaProducto.EAN || "-"}
                            </div>

                            <div>
                              <strong>Descripción:</strong>{" "}
                              {skuCajaProducto.DESCRIPCION || "-"}
                            </div>

                            <div>
                              <strong>Unidad:</strong>{" "}
                              {skuCajaProducto.UNIDAD_MEDIDA || "-"}
                            </div>
                          </div>
                        )}

                        {/* CANTIDAD */}

                        {skuCajaProducto && (
                          <div
                            className="campo"
                            style={{
                              marginTop: "20px",
                              maxWidth: "250px",
                            }}
                          >
                            <label>Cantidad</label>

                            <input
                              type="number"
                              min="1"
                              step="1"
                              value={skuCajaCantidad}
                              onChange={(e) =>
                                setSkuCajaCantidad(e.target.value)
                              }
                            />
                          </div>
                        )}
                      </div>
                      {skuCajaProducto && (
                        <button
                          type="button"
                          className="guardar"
                          onClick={agregarSkuCajaAuditoria}
                          disabled={guardandoSkuCaja}
                          style={{
                            marginTop: "20px",
                          }}
                        >
                          {guardandoSkuCaja
                            ? "Guardando..."
                            : "💾 Guardar SKU en esta caja"}
                        </button>
                      )}
                      {/* BOTONES */}

                      <div
                        style={{
                          display: "flex",
                          gap: "10px",
                          flexWrap: "wrap",
                          marginTop: "20px",
                        }}
                      >
                        <button
                          type="button"
                          onClick={() => {
                            setCajaSeleccionadaAuditoria(null);
                            setSkuCajaBusqueda("");
                            setSkuCajaProducto(null);
                            setSkuCajaCantidad("1");
                            setMensaje("");
                          }}
                        >
                          ← Volver a cajas
                        </button>
                      </div>
                    </div>
                  )}
                  {/* CAMBIAR DE TIPO DE REGISTRO */}

                  <div
                    style={{
                      marginTop: "20px",
                      paddingTop: "15px",
                      borderTop: "1px solid #ddd",
                    }}
                  >
                    <button
                      onClick={volverMenuContenido}
                      disabled={agregandoCaja}
                    >
                      ← Elegir otro tipo de registro
                    </button>
                  </div>
                </>
              )}

              {modoContenidoPallet === "SKU_SUELTO" && (
                <>
                  <h3>🏷️ Registrar SKU suelto</h3>

                  <div
                    className="mensaje"
                    style={{
                      marginBottom: "20px",
                    }}
                  >
                    Registre los productos que se encuentren sueltos
                    directamente dentro del pallet.
                  </div>

                  {/* BUSCAR SKU / EAN */}

                  <div className="campo">
                    <label>SKU o EAN</label>

                    <div
                      style={{
                        display: "flex",
                        gap: "10px",
                        flexWrap: "wrap",
                      }}
                    >
                      <input
                        type="text"
                        value={skuSueltoBusqueda}
                        onChange={(e) => setSkuSueltoBusqueda(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            buscarProductoSkuSuelto();
                          }
                        }}
                        placeholder="Escanee o ingrese SKU / EAN"
                        style={{
                          flex: 1,
                          minWidth: "220px",
                        }}
                      />

                      <button
                        className="guardar"
                        onClick={buscarProductoSkuSuelto}
                        disabled={buscandoSkuSuelto}
                      >
                        {buscandoSkuSuelto ? "Buscando..." : "🔎 Buscar"}
                      </button>
                    </div>
                  </div>

                  {/* PRODUCTO ENCONTRADO */}

                  {skuSueltoProducto && (
                    <div
                      className="mensaje"
                      style={{
                        marginTop: "20px",
                      }}
                    >
                      <div>
                        <strong>SKU:</strong> {skuSueltoProducto.SKU}
                      </div>

                      <div>
                        <strong>EAN:</strong> {skuSueltoProducto.EAN || "-"}
                      </div>

                      <div>
                        <strong>Descripción:</strong>{" "}
                        {skuSueltoProducto.DESCRIPCION || "-"}
                      </div>

                      <div>
                        <strong>Unidad:</strong>{" "}
                        {skuSueltoProducto.UNIDAD_MEDIDA || "-"}
                      </div>
                    </div>
                  )}

                  {/* CANTIDAD */}

                  {skuSueltoProducto && (
                    <div
                      className="campo"
                      style={{
                        marginTop: "20px",
                        maxWidth: "250px",
                      }}
                    >
                      <label>Cantidad</label>

                      <input
                        type="number"
                        min="1"
                        step="1"
                        value={skuSueltoCantidad}
                        onChange={(e) => setSkuSueltoCantidad(e.target.value)}
                      />
                    </div>
                  )}

                  {/* BOTONES */}

                  <div
                    style={{
                      display: "flex",
                      gap: "10px",
                      flexWrap: "wrap",
                      marginTop: "20px",
                    }}
                  >
                    <button
                      onClick={volverMenuContenido}
                      disabled={guardandoSkuSuelto}
                    >
                      ← Volver
                    </button>

                    {skuSueltoProducto && (
                      <button
                        className="guardar"
                        onClick={agregarSkuSueltoAuditoria}
                        disabled={guardandoSkuSuelto}
                      >
                        {guardandoSkuSuelto
                          ? "Registrando..."
                          : "✓ Registrar SKU"}
                      </button>
                    )}
                  </div>

                  {/* SKU SUELTOS REGISTRADOS */}

                  {skuSueltoRegistrados.length > 0 && (
                    <div
                      className="tabla-historial"
                      style={{
                        marginTop: "25px",
                      }}
                    >
                      <h3>SKU sueltos registrados</h3>

                      <table>
                        <thead>
                          <tr>
                            <th>SKU</th>
                            <th>EAN</th>
                            <th>Descripción</th>
                            <th>Cantidad</th>
                            <th>Volumen</th>
                          </tr>
                        </thead>

                        <tbody>
                          {skuSueltoRegistrados.map((registro) => (
                            <tr key={registro.id}>
                              <td>
                                <strong>{registro.sku}</strong>
                              </td>

                              <td>{registro.ean || "-"}</td>

                              <td>{registro.descripcion || "-"}</td>

                              <td>{registro.cantidad}</td>

                              <td>
                                {Number(registro.volumen_total || 0).toFixed(4)}{" "}
                                m³
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}

                  {/* CAMBIAR DE REGISTRO */}

                  <div
                    style={{
                      marginTop: "20px",
                      paddingTop: "15px",
                      borderTop: "1px solid #ddd",
                    }}
                  >
                    <button
                      onClick={volverMenuContenido}
                      disabled={guardandoSkuSuelto}
                    >
                      ← Elegir otro tipo de registro
                    </button>
                  </div>
                </>
              )}
            </div>
          )}

          {/* =====================================================
            EXPORTACIÓN DE AUDITORÍAS
            El historial de pallets NO se muestra a los operarios.
            Los registros permanecen guardados en Supabase y pueden
            exportarse mediante el botón de Excel.
          ===================================================== */}

          <div
            className="producto"
            style={{
              marginBottom: "20px",
            }}
          >
            <h2>③ Auditorías registradas</h2>

            <div
              style={{
                display: "flex",
                gap: "10px",
                flexWrap: "wrap",
                marginBottom: "10px",
              }}
            >
              <button
                className="guardar"
                onClick={exportarAuditoriaPallet}
                disabled={exportandoAuditoria}
              >
                {exportandoAuditoria
                  ? "Exportando..."
                  : "📊 Exportar auditorías a Excel"}
              </button>
            </div>

            <div className="mensaje">
              Las auditorías realizadas se guardan automáticamente.
              <br />
              El historial de pallets no se muestra en pantalla para evitar
              confusiones durante la operación.
            </div>
          </div>

          {/* =====================================================
            MENSAJE GENERAL
          ===================================================== */}

          {mensaje && (
            <div
              className="mensaje"
              style={{
                whiteSpace: "pre-line",
                marginBottom: "20px",
              }}
            >
              {mensaje}
            </div>
          )}
        </div>
      </div>
    );
  }
  /* =========================================================
     PANTALLA MEDICIÓN
  ========================================================= */

  return (
    <div className="app">
      <div className="contenedor">
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: "12px",
            flexWrap: "wrap",
            marginBottom: "15px",
          }}
        >
          <div>
            <h1>Medición de Productos</h1>

            <div
              className="mensaje"
              style={{
                marginTop: "8px",
              }}
            >
              <strong>Usuario:</strong> {personalLogin?.NOMBRE || "-"} —{" "}
              {personalLogin?.AREA || "-"}
            </div>
          </div>

          <div
            style={{
              display: "flex",
              gap: "8px",
              flexWrap: "wrap",
            }}
          >
            <button onClick={volverInicio}>← Inicio</button>

            <button onClick={cerrarSesion}>Cerrar sesión</button>
          </div>
        </div>

        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: "10px",
            marginBottom: "15px",
            padding: "12px 15px",
            borderRadius: "12px",
            background:
              tipoMedicion === "PRODUCTO"
                ? "rgba(30, 110, 220, 0.08)"
                : "rgba(230, 150, 30, 0.10)",
            border:
              tipoMedicion === "PRODUCTO"
                ? "1px solid rgba(30,110,220,0.18)"
                : "1px solid rgba(230,150,30,0.20)",
          }}
        >
          <span
            style={{
              fontSize: "24px",
            }}
          >
            {tipoMedicion === "PRODUCTO" ? "📏" : "📦"}
          </span>

          <div>
            <strong>
              {tipoMedicion === "PRODUCTO"
                ? "Medición de producto"
                : "Medición de caja / master"}
            </strong>

            <div
              style={{
                fontSize: "13px",
                opacity: 0.7,
              }}
            >
              {tipoMedicion === "PRODUCTO"
                ? "Escanee o busque un SKU para registrar su volumetría."
                : "Busque los SKU que contiene la caja y registre su volumen."}
            </div>
          </div>
        </div>

        <div className="buscador">
          <input
            ref={inputBusquedaRef}
            type="text"
            value={busqueda}
            onChange={(e) => manejarBusqueda(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                buscarProducto();
              }
            }}
            placeholder="Escanee EAN o ingrese SKU"
          />

          <button onClick={() => buscarProducto()}>Buscar</button>

          <button onClick={limpiar}>Limpiar</button>
        </div>

        {cargando && <div className="mensaje">Buscando producto...</div>}

        {mensaje && (
          <div
            className="mensaje"
            style={{
              whiteSpace: "pre-line",
            }}
          >
            {mensaje}
          </div>
        )}

        {producto && (
          <>
            <div className="producto">
              <h2>Información del producto</h2>

              <div className="dato">
                <strong>SKU</strong>

                <span>{producto.SKU}</span>
              </div>

              <div className="dato">
                <strong>EAN</strong>

                <span>{producto.EAN || "-"}</span>
              </div>

              <div className="dato">
                <strong>Descripción</strong>

                <span>{producto.DESCRIPCION || "-"}</span>
              </div>

              <div className="dato">
                <strong>Unidad</strong>

                <span>{producto.UNIDAD_MEDIDA || "-"}</span>
              </div>

              <div className="dato">
                <strong>Factor WMS</strong>

                <span>{producto.FACTOR_WMS || "-"}</span>
              </div>
            </div>

            <div className="producto">
              <h2>Última medición registrada</h2>

              {cargandoUltimaMedicion ? (
                <div className="mensaje">Consultando última medición...</div>
              ) : ultimaMedicion ? (
                <>
                  <div className="dato">
                    <strong>Fecha</strong>

                    <span>{formatearFecha(ultimaMedicion.fecha_hora)}</span>
                  </div>

                  <div className="dato">
                    <strong>Hora</strong>

                    <span>{formatearHora(ultimaMedicion.fecha_hora)}</span>
                  </div>

                  <div className="dato">
                    <strong>Volumen unitario</strong>

                    <span>
                      {Number(ultimaMedicion.volumen_unitario || 0).toFixed(4)}{" "}
                      m³
                    </span>
                  </div>

                  <div className="dato">
                    <strong>Registrado por</strong>

                    <span>{ultimaMedicion.REGISTRADO_POR || "-"}</span>
                  </div>

                  <div className="dato">
                    <strong>Área</strong>

                    <span>{ultimaMedicion.AREA || "-"}</span>
                  </div>
                </>
              ) : (
                <div className="mensaje">
                  Este SKU no tiene mediciones anteriores.
                </div>
              )}
            </div>

            <div className="medicion">
              <h2>Tipo de medición</h2>

              <div className="opciones">
                <button
                  className={
                    tipoMedicion === "PRODUCTO" ? "opcion activa" : "opcion"
                  }
                  onClick={() => entrarAMedicion("PRODUCTO")}
                >
                  📏 PRODUCTO INDIVIDUAL
                </button>

                <button
                  className={
                    tipoMedicion === "CAJA" ? "opcion activa" : "opcion"
                  }
                  onClick={() => entrarAMedicion("CAJA")}
                >
                  📦 CAJA / MASTER
                </button>
              </div>
            </div>

            {tipoMedicion === "PRODUCTO" && (
              <div className="medicion">
                <h2>Medición</h2>

                <div className="campo">
                  <label>¿El producto es multipieza?</label>

                  <div className="opciones">
                    <button
                      className={!esMultipieza ? "opcion activa" : "opcion"}
                      onClick={() => cambiarTipoProducto(false)}
                    >
                      NO
                    </button>

                    <button
                      className={esMultipieza ? "opcion activa" : "opcion"}
                      onClick={() => cambiarTipoProducto(true)}
                    >
                      SÍ
                    </button>
                  </div>
                </div>

                {esMultipieza ? (
                  <>
                    <div className="campo">
                      <label>Cantidad de piezas</label>

                      <select
                        value={cantidadPiezas}
                        onChange={(e) =>
                          setCantidadPiezas(Number(e.target.value))
                        }
                      >
                        <option value={2}>2 piezas</option>
                        <option value={3}>3 piezas</option>
                        <option value={4}>4 piezas</option>
                        <option value={5}>5 piezas</option>
                        <option value={6}>6 piezas</option>
                      </select>
                    </div>

                    {piezas.slice(0, cantidadPiezas).map((pieza, indice) => (
                      <div className="pieza" key={indice}>
                        <h3>Pieza {indice + 1}</h3>

                        {/* ORDEN VISUAL: ANCHO → ALTO → LARGO */}

                        <div className="campo">
                          <label>Ancho (cm)</label>

                          <input
                            type="number"
                            min="0"
                            value={pieza.ancho}
                            onWheel={(e) => e.currentTarget.blur()}
                            onChange={(e) =>
                              actualizarPieza(indice, "ancho", e.target.value)
                            }
                          />
                        </div>

                        <div className="campo">
                          <label>Alto (cm)</label>

                          <input
                            type="number"
                            min="0"
                            value={pieza.alto}
                            onWheel={(e) => e.currentTarget.blur()}
                            onChange={(e) =>
                              actualizarPieza(indice, "alto", e.target.value)
                            }
                          />
                        </div>

                        <div className="campo">
                          <label>Largo (cm)</label>

                          <input
                            type="number"
                            min="0"
                            value={pieza.largo}
                            onWheel={(e) => e.currentTarget.blur()}
                            onChange={(e) =>
                              actualizarPieza(indice, "largo", e.target.value)
                            }
                          />
                        </div>

                        <div className="volumen">
                          <div>
                            <strong>Volumen pieza</strong>

                            <span>
                              {calcularVolumenPieza(pieza).toFixed(4)} m³
                            </span>
                          </div>
                        </div>
                      </div>
                    ))}
                  </>
                ) : (
                  <>
                    {/* ORDEN VISUAL: ANCHO → ALTO → LARGO */}

                    <div className="campo">
                      <label>Ancho (cm)</label>

                      <input
                        type="number"
                        min="0"
                        value={ancho}
                        onWheel={(e) => e.currentTarget.blur()}
                        onChange={(e) => setAncho(e.target.value)}
                      />
                    </div>

                    <div className="campo">
                      <label>Alto (cm)</label>

                      <input
                        type="number"
                        min="0"
                        value={alto}
                        onWheel={(e) => e.currentTarget.blur()}
                        onChange={(e) => setAlto(e.target.value)}
                      />
                    </div>

                    <div className="campo">
                      <label>Largo (cm)</label>

                      <input
                        type="number"
                        min="0"
                        value={largo}
                        onWheel={(e) => e.currentTarget.blur()}
                        onChange={(e) => setLargo(e.target.value)}
                      />
                    </div>
                  </>
                )}

                <div className="campo">
                  <label>Factor de medición</label>

                  <input
                    type="number"
                    min="0"
                    value={factorMedicion}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setFactorMedicion(e.target.value)}
                  />
                </div>

                <div className="volumen">
                  <div>
                    <strong>Volumen Master</strong>

                    <span>{volumenMaster.toFixed(4)} m³</span>
                  </div>

                  <div>
                    <strong>Volumen Unitario</strong>

                    <span>{volumenUnitario.toFixed(4)} m³</span>
                  </div>
                </div>

                <button
                  className="guardar"
                  onClick={guardarMedicion}
                  disabled={guardando || cargandoUltimaMedicion}
                >
                  {guardando ? "Guardando..." : "Guardar medición"}
                </button>
              </div>
            )}

            {tipoMedicion === "CAJA" && (
              <div className="medicion">
                <h2>📦 Medición de Caja / Master</h2>

                <div className="mensaje">
                  Mida la caja completa, incluyendo todo el espacio que ocupa
                  físicamente.
                </div>

                {/* ORDEN VISUAL: ANCHO → ALTO → LARGO */}

                <div className="campo">
                  <label>Ancho de la caja (cm)</label>

                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    value={cajaAncho}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setCajaAncho(e.target.value)}
                    placeholder="Ej. 40"
                  />
                </div>

                <div className="campo">
                  <label>Alto de la caja (cm)</label>

                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    value={cajaAlto}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setCajaAlto(e.target.value)}
                    placeholder="Ej. 40"
                  />
                </div>

                <div className="campo">
                  <label>Largo de la caja (cm)</label>

                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    value={cajaLargo}
                    onWheel={(e) => e.currentTarget.blur()}
                    onChange={(e) => setCajaLargo(e.target.value)}
                    placeholder="Ej. 60"
                  />
                </div>

                <div className="volumen">
                  <div>
                    <strong>Volumen de la caja</strong>

                    <span>{volumenCaja.toFixed(4)} m³</span>
                  </div>
                </div>

                <div
                  className="producto"
                  style={{
                    marginTop: "20px",
                  }}
                >
                  <h2>Contenido de la caja</h2>

                  <div className="dato">
                    <strong>SKU actual</strong>

                    <span>{producto.SKU}</span>
                  </div>

                  <div className="dato">
                    <strong>Volumen unitario</strong>

                    <span>
                      {Number(ultimaMedicion?.volumen_unitario || 0).toFixed(4)}{" "}
                      m³
                    </span>
                  </div>

                  <div className="campo">
                    <label>Cantidad de unidades</label>

                    <input
                      type="number"
                      min="1"
                      step="1"
                      value={cantidadCaja}
                      onWheel={(e) => e.currentTarget.blur()}
                      onChange={(e) => setCantidadCaja(e.target.value)}
                    />
                  </div>

                  <button onClick={agregarProductoACaja}>
                    + Agregar SKU a la caja
                  </button>
                </div>

                {itemsCaja.length > 0 && (
                  <div
                    className="producto"
                    style={{
                      marginTop: "20px",
                    }}
                  >
                    <h2>Productos dentro de la caja</h2>

                    <div className="tabla-historial">
                      <table>
                        <thead>
                          <tr>
                            <th>SKU</th>
                            <th>Cantidad</th>
                            <th>Vol. unitario</th>
                            <th>Vol. total</th>
                            <th>Acción</th>
                          </tr>
                        </thead>

                        <tbody>
                          {itemsCaja.map((item) => (
                            <tr key={item.sku}>
                              <td>{item.sku}</td>

                              <td>{item.cantidad}</td>

                              <td>{item.volumenUnitario.toFixed(4)} m³</td>

                              <td>{item.volumenTotal.toFixed(4)} m³</td>

                              <td>
                                <button
                                  onClick={() => eliminarItemCaja(item.sku)}
                                >
                                  Quitar
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                <div
                  className="volumen"
                  style={{
                    marginTop: "20px",
                  }}
                >
                  <div>
                    <strong>Unidades dentro de la caja</strong>

                    <span>{cantidadTotalCaja}</span>
                  </div>

                  <div>
                    <strong>Volumen productos</strong>

                    <span>{volumenProductosCaja.toFixed(4)} m³</span>
                  </div>

                  <div>
                    <strong>Volumen caja</strong>

                    <span>{volumenCaja.toFixed(4)} m³</span>
                  </div>

                  <div>
                    <strong>Diferencia de volumen</strong>

                    <span>{diferenciaVolumenCaja.toFixed(4)} m³</span>
                  </div>

                  <div>
                    <strong>Ocupación de la caja</strong>

                    <span>{porcentajeOcupacionCaja.toFixed(1)} %</span>
                  </div>
                </div>

                <div
                  className="mensaje"
                  style={{
                    marginTop: "15px",
                  }}
                >
                  {itemsCaja.length === 0
                    ? "Agregue los SKU que contiene la caja."
                    : cajaConfirmada
                    ? "✓ Caja confirmada. Lista para guardar."
                    : "La información de la caja está lista para revisión."}
                </div>

                <div
                  style={{
                    display: "flex",
                    gap: "10px",
                    flexWrap: "wrap",
                    marginTop: "15px",
                  }}
                >
                  <button onClick={limpiarCaja}>Limpiar caja</button>

                  <button
                    className="guardar"
                    disabled={volumenCaja <= 0 || itemsCaja.length === 0}
                    onClick={revisarCaja}
                  >
                    Revisar caja
                  </button>
                </div>
              </div>
            )}
          </>
        )}

        {/* =====================================================
            ADVERTENCIA VOLUMEN MENOR
        ===================================================== */}

        {mostrarAdvertencia && (
          <div className="modal">
            <div className="modal-contenido">
              <div className="modal-cabecera">
                <h2>⚠️ Advertencia de volumetría</h2>

                <button onClick={cancelarVolumenMenor} disabled={guardando}>
                  ✕
                </button>
              </div>

              <div className="detalle-producto">
                <p>
                  El volumen unitario que está registrando es{" "}
                  <strong>menor</strong> al volumen registrado anteriormente
                  para este SKU.
                </p>

                <div>
                  <strong>SKU:</strong> {producto?.SKU || "-"}
                </div>

                <div>
                  <strong>Volumen anterior:</strong>{" "}
                  {Number(volumenAnteriorAdvertencia || 0).toFixed(4)} m³
                </div>

                <div>
                  <strong>Nuevo volumen:</strong> {volumenUnitario.toFixed(4)}{" "}
                  m³
                </div>
              </div>

              <div
                style={{
                  display: "flex",
                  gap: "10px",
                  justifyContent: "flex-end",
                  flexWrap: "wrap",
                  marginTop: "20px",
                }}
              >
                <button onClick={cancelarVolumenMenor} disabled={guardando}>
                  No, revisar
                </button>

                <button
                  className="guardar"
                  onClick={aceptarVolumenMenor}
                  disabled={guardando}
                >
                  {guardando ? "Guardando..." : "Sí, guardar"}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* =====================================================
            REVISIÓN CAJA
        ===================================================== */}

        {mostrarRevisionCaja && (
          <div className="modal">
            <div
              className="modal-contenido"
              style={{
                maxWidth: "900px",
              }}
            >
              <div className="modal-cabecera">
                <h2>📦 REVISIÓN DE CAJA MASTER</h2>

                <button onClick={() => setMostrarRevisionCaja(false)}>✕</button>
              </div>

              <div className="detalle-producto">
                <h3>Datos de la caja</h3>

                {/* ORDEN VISUAL: ANCHO → ALTO → LARGO */}

                <div>
                  <strong>Ancho:</strong> {Number(cajaAncho).toFixed(2)} cm
                </div>

                <div>
                  <strong>Alto:</strong> {Number(cajaAlto).toFixed(2)} cm
                </div>

                <div>
                  <strong>Largo:</strong> {Number(cajaLargo).toFixed(2)} cm
                </div>

                <div>
                  <strong>Volumen físico:</strong> {volumenCaja.toFixed(4)} m³
                </div>
              </div>

              <div
                className="tabla-piezas"
                style={{
                  marginTop: "20px",
                }}
              >
                <h3>Contenido de la caja</h3>

                <div className="tabla-historial">
                  <table>
                    <thead>
                      <tr>
                        <th>SKU</th>
                        <th>Descripción</th>
                        <th>Cantidad</th>
                        <th>Vol. unitario</th>
                        <th>Vol. total</th>
                      </tr>
                    </thead>

                    <tbody>
                      {itemsCaja.map((item) => (
                        <tr key={item.sku}>
                          <td>{item.sku}</td>

                          <td>{item.descripcion || "-"}</td>

                          <td>{item.cantidad}</td>

                          <td>{item.volumenUnitario.toFixed(4)} m³</td>

                          <td>{item.volumenTotal.toFixed(4)} m³</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div
                className="detalle-totales"
                style={{
                  marginTop: "20px",
                }}
              >
                <div>
                  <strong>Total unidades</strong>

                  <span>{cantidadTotalCaja}</span>
                </div>

                <div>
                  <strong>Volumen productos</strong>

                  <span>{volumenProductosCaja.toFixed(4)} m³</span>
                </div>

                <div>
                  <strong>Volumen caja</strong>

                  <span>{volumenCaja.toFixed(4)} m³</span>
                </div>

                <div>
                  <strong>Diferencia</strong>

                  <span>{diferenciaVolumenCaja.toFixed(4)} m³</span>
                </div>

                <div>
                  <strong>Ocupación</strong>

                  <span>{porcentajeOcupacionCaja.toFixed(1)} %</span>
                </div>
              </div>

              <div
                className="mensaje"
                style={{
                  marginTop: "20px",
                }}
              >
                {diferenciaVolumenCaja >= 0
                  ? "La caja tiene capacidad suficiente para el volumen registrado de sus productos."
                  : "⚠️ El volumen de los productos supera el volumen físico de la caja. Revise las cantidades o mediciones antes de confirmar."}
              </div>

              <div
                style={{
                  display: "flex",
                  gap: "10px",
                  justifyContent: "flex-end",
                  flexWrap: "wrap",
                  marginTop: "20px",
                }}
              >
                <button onClick={() => setMostrarRevisionCaja(false)}>
                  ← Volver a editar
                </button>

                <button
                  className="guardar"
                  onClick={confirmarCaja}
                  disabled={guardando}
                >
                  {guardando ? "Guardando caja..." : "✓ Confirmar caja"}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
