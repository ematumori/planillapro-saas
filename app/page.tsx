"use client";

import React, { useState, useRef, useEffect } from "react";
import ExcelJS from "exceljs";
import { supabase } from "@/lib/supabase";
import type { User } from "@supabase/supabase-js";
import {
  Download,
  Plus,
  Trash2,
  FileSpreadsheet,
  Building2,
  Calculator,
  Upload,
  Printer,
  CheckCircle2,
  TrendingUp,
  Layers,
  SlidersHorizontal,
  X,
  Cloud,
  LogIn,
  LogOut,
  FolderOpen,
  Save,
  Loader2,
  Crown,
  FileUp,
  Sparkles,
  Check,
} from "lucide-react";

type TemplateType =
  | "mercadolibre"
  | "stock"
  | "caja"
  | "presupuesto"
  | "cuentas"
  | "gastos"
  | "precios"
  | "pedidos";

type RubroType =
  | "general"
  | "indumentaria"
  | "electronica"
  | "gastronomia"
  | "ferreteria"
  | "farmacia"
  | "automotor"
  | "libreria"
  | "servicios"
  | "muebles";

interface RowData {
  id: number;
  concepto: string;
  detalleRubro: string;
  valorA: number;
  valorB: number;
  cantidad: number;
  stockMinimo: number;
  fecha: string;
  desplegable: string;
}

interface SavedPlanilla {
  id: string;
  nombre_negocio: string;
  subtitulo_negocio: string;
  plantilla: TemplateType;
  rubro: RubroType;
  regimen_fiscal: string;
  moneda: string;
  color_tema: string;
  incluir_envio: boolean;
  costo_envio_fijo: number;
  filas: RowData[];
  created_at: string;
}

// CONFIGURACIÓN DE LÍMITES FREEMIUM Y LINKS DE COBRO
const LIMITE_FILAS_GRATIS = 8;
const LIMITE_PLANILLAS_GRATIS = 2;
const LINK_PAGO_MERCADOPAGO = "https://www.mercadopago.com.ar/subscriptions"; // Reemplaza con tu link de suscripción MP
const LINK_PAGO_STRIPE = "https://buy.stripe.com/test_tu_enlace"; // Reemplaza con tu Stripe Payment Link

const TEMAS_VISUALES = [
  { id: "1E3A8A", nombre: "Azul Ejecutivo", bgClass: "bg-blue-900", ringClass: "ring-blue-900" },
  { id: "065F46", nombre: "Verde Finanzas", bgClass: "bg-emerald-800", ringClass: "ring-emerald-800" },
  { id: "0F172A", nombre: "Pizarra Oscuro", bgClass: "bg-slate-900", ringClass: "ring-slate-900" },
  { id: "7C2D12", nombre: "Terracota", bgClass: "bg-orange-900", ringClass: "ring-orange-900" },
  { id: "4C1D95", nombre: "Índigo Moderno", bgClass: "bg-violet-900", ringClass: "ring-violet-900" },
];

const RUBROS_CONFIG: Record<RubroType, { label: string; columnaExtra: string; ejemplo: string }> = {
  general: { label: "Comercio General", columnaExtra: "", ejemplo: "-" },
  indumentaria: { label: "Indumentaria y Calzado", columnaExtra: "Talle / Color", ejemplo: "42 / Negro" },
  electronica: { label: "Electrónica y Tecnología", columnaExtra: "Garantía / N° Serie", ejemplo: "12 meses / SN-884" },
  gastronomia: { label: "Gastronomía y Alimentos", columnaExtra: "Vencimiento / Lote", ejemplo: "15/12/2026 - L04" },
  ferreteria: { label: "Ferretería y Construcción", columnaExtra: "Unidad / Medida", ejemplo: "Caja x 100u / 2 pulg." },
  farmacia: { label: "Farmacia y Cosmética", columnaExtra: "Laboratorio / Lote", ejemplo: "Lab. Sur / B-902" },
  automotor: { label: "Automotor y Repuestos", columnaExtra: "Marca / Modelo", ejemplo: "VW Gol / 2018-2023" },
  libreria: { label: "Librería y Papelería", columnaExtra: "Presentación / Pack", ejemplo: "Resma 500 hojas A4" },
  servicios: { label: "Servicios Profesionales", columnaExtra: "Modalidad / Horas", ejemplo: "Mensual / 10 hs" },
  muebles: { label: "Muebles y Decoración", columnaExtra: "Material / Medidas", ejemplo: "Roble / 160x80 cm" },
};

const PLANTILLAS_CONFIG: Record<
  TemplateType,
  {
    label: string;
    descripcion: string;
    tituloExcel: string;
    colA: string;
    colB: string;
    colDropdown: string;
    colResultado: string;
    opcionesDropdown: string[];
    defaultDropdown: string;
  }
> = {
  mercadolibre: {
    label: "Rentabilidad y Márgenes (Mercado Libre)",
    descripcion: "Calcula ganancia neta descontando comisiones, envíos e impuestos.",
    tituloExcel: "CALCULADORA DE RENTABILIDAD - MERCADO LIBRE",
    colA: "Precio Venta",
    colB: "Costo Producto",
    colDropdown: "Tipo Publicación",
    colResultado: "Ganancia Neta Unit.",
    opcionesDropdown: ["Clásica (13%)", "Premium (28%)", "Full Clásica (13%)", "Full Premium (28%)"],
    defaultDropdown: "Clásica (13%)",
  },
  stock: {
    label: "Control de Stock e Inventario",
    descripcion: "Gestiona existencias, alertas de reposición y capital inmovilizado.",
    tituloExcel: "CONTROL DE STOCK Y VALORIZACIÓN DE INVENTARIO",
    colA: "Costo Unitario",
    colB: "Stock Actual",
    colDropdown: "Estado Stock",
    colResultado: "Valor Total Stock",
    opcionesDropdown: ["Disponible", "Stock Bajo", "Sin Stock", "En Camino"],
    defaultDropdown: "Disponible",
  },
  caja: {
    label: "Flujo de Caja Diario (Ingresos / Egresos)",
    descripcion: "Registra movimientos diarios por medio de pago y arqueo de caja.",
    tituloExcel: "FLUJO DE CAJA DIARIO Y ARQUEO",
    colA: "Monto Unitario",
    colB: "Cantidad",
    colDropdown: "Medio de Pago",
    colResultado: "Total Operación",
    opcionesDropdown: ["Efectivo", "Mercado Pago", "Pix / Transferencia", "Tarjeta Crédito", "Tarjeta Débito"],
    defaultDropdown: "Mercado Pago",
  },
  presupuesto: {
    label: "Presupuestos y Cotizaciones a Clientes",
    descripcion: "Arma cotizaciones formales con impuestos y plazos de entrega.",
    tituloExcel: "PRESUPUESTO COMERCIAL",
    colA: "Precio Unitario",
    colB: "Cantidad",
    colDropdown: "Condición Entrega",
    colResultado: "Subtotal Neto",
    opcionesDropdown: ["Inmediata", "48 hs hábiles", "A pedido (7 días)", "Retiro en local"],
    defaultDropdown: "Inmediata",
  },
  cuentas: {
    label: "Cuentas Corrientes (Clientes / Proveedores)",
    descripcion: "Controla saldos pendientes, pagos parciales y vencimientos.",
    tituloExcel: "ESTADO DE CUENTAS CORRIENTES",
    colA: "Monto Facturado",
    colB: "Monto Abonado",
    colDropdown: "Estado Cuenta",
    colResultado: "Saldo Pendiente",
    opcionesDropdown: ["Pendiente", "Pago Parcial", "Saldado", "Vencido"],
    defaultDropdown: "Pendiente",
  },
  gastos: {
    label: "Control de Gastos Fijos y Variables",
    descripcion: "Clasifica costos operativos mensuales por categoría.",
    tituloExcel: "REGISTRO DE GASTOS OPERATIVOS",
    colA: "Importe Base",
    colB: "Cantidad / Meses",
    colDropdown: "Categoría Gasto",
    colResultado: "Gasto Total",
    opcionesDropdown: ["Alquiler / Servicios", "Proveedores", "Impuestos / Tasas", "Publicidad / Ads", "Sueldos"],
    defaultDropdown: "Proveedores",
  },
  precios: {
    label: "Lista de Precios (Costo + Margen %)",
    descripcion: "Calcula precios de venta al público aplicando margen e impuestos.",
    tituloExcel: "LISTA DE PRECIOS ACTUALIZADA",
    colA: "Costo Neto",
    colB: "Margen (%)",
    colDropdown: "Canal Venta",
    colResultado: "Precio Público",
    opcionesDropdown: ["Minorista", "Mayorista", "Tienda Online", "Mercado Libre"],
    defaultDropdown: "Minorista",
  },
  pedidos: {
    label: "Seguimiento de Pedidos y Logística",
    descripcion: "Organiza despachos, costos de envío y estados de entrega.",
    tituloExcel: "CONTROL DE PEDIDOS Y DESPACHOS",
    colA: "Valor Pedido",
    colB: "Costo Envío",
    colDropdown: "Estado Envío",
    colResultado: "Total a Cobrar",
    opcionesDropdown: ["Preparando", "Despachado", "En Tránsito", "Entregado", "Cancelado"],
    defaultDropdown: "Preparando",
  },
};

export default function PlanillasGenerator() {
  // Estados de Autenticación, Suscripción y Nube
  const [user, setUser] = useState<User | null>(null);
  const [esPro, setEsPro] = useState<boolean>(false);
  const [emailUsuario, setEmailUsuario] = useState<string>("");
  const [inputEmail, setInputEmail] = useState<string>("");
  const [mensajeLogin, setMensajeLogin] = useState<string>("");
  const [showProModal, setShowProModal] = useState<boolean>(false);
  const [proMotivo, setProMotivo] = useState<string>("");

  const [showAuthModal, setShowAuthModal] = useState(false);
  const [isRegisterMode, setIsRegisterMode] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState<string | null>(null);

  const [guardandoNube, setGuardandoNube] = useState(false);
  const [mensajeNube, setMensajeNube] = useState<string | null>(null);
  const [misPlanillas, setMisPlanillas] = useState<SavedPlanilla[]>([]);
  const [showSavedDrawer, setShowSavedDrawer] = useState(false);

  // Referencias de archivos
  const fileInputRef = useRef<HTMLInputElement>(null);
  const importExcelInputRef = useRef<HTMLInputElement>(null);

  // Estados de la Planilla
  const [nombreNegocio, setNombreNegocio] = useState("Mi Negocio Comercial");
  const [subtituloNegocio, setSubtituloNegocio] = useState("CUIT / CNPJ: 30-71234567-8 • Ventas Online y Local");
  const [logoUrl, setLogoUrl] = useState<string | null>(null);

  const [plantilla, setPlantilla] = useState<TemplateType>("mercadolibre");
  const [rubro, setRubro] = useState<RubroType>("general");
  const [regimenFiscal, setRegimenFiscal] = useState<string>("AR_21");
  const [moneda, setMoneda] = useState<string>("$");
  const [incluirEnvio, setIncluirEnvio] = useState<boolean>(true);
  const [costoEnvioFijo, setCostoEnvioFijo] = useState<number>(4500);
  const [colorTema, setColorTema] = useState<string>("1E3A8A");

  const configActual = PLANTILLAS_CONFIG[plantilla];
  const rubroActual = RUBROS_CONFIG[rubro];
  const verificarEsPro = (correo: string): boolean => {
    const listaCruda = process.env.NEXT_PUBLIC_PRO_EMAILS || "";
    const correosAutorizados = listaCruda
      .split(",")
      .map((c) => c.trim().toLowerCase())
      .filter((c) => c.length > 0);

    return correosAutorizados.includes(correo.trim().toLowerCase());
  };

  useEffect(() => {
    const hoy = new Date().toISOString().split("T")[0];
    setFilas((prev) =>
      prev.map((f) => (f.fecha === "" ? { ...f, fecha: hoy } : f))
    );

    const emailGuardado = localStorage.getItem("calculopro_email");
    const esProPorLista = emailGuardado ? verificarEsPro(emailGuardado) : false;

    if (emailGuardado) {
      setEmailUsuario(emailGuardado);
      setInputEmail(emailGuardado);
      setEsPro(esProPorLista);
    }

    supabase.auth.getSession().then(({ data: { session } }) => {
      const u = session?.user ?? null;
      setUser(u);
      if (u) {
        cargarPlanillasUsuario();
        verificarEstadoPro(u.id, u.email || emailGuardado || "");
      }
    });

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, session) => {
      const u = session?.user ?? null;
      setUser(u);
      if (u) {
        cargarPlanillasUsuario();
        verificarEstadoPro(u.id, u.email || emailGuardado || "");
      } else {
        setMisPlanillas([]);
        // Solo apaga esPro si tampoco está autorizado por la lista de correos
        const correoActual = localStorage.getItem("calculopro_email") || "";
        setEsPro(verificarEsPro(correoActual));
      }
    });

    return () => subscription.unsubscribe();
  }, []);

  const verificarEstadoPro = async (userId: string, correo?: string) => {
    const correoCheck = correo || localStorage.getItem("calculopro_email") || "";
    if (verificarEsPro(correoCheck)) {
      setEsPro(true);
      return;
    }

    const { data } = await supabase
      .from("perfiles")
      .select("es_pro")
      .eq("user_id", userId)
      .maybeSingle();

    setEsPro(Boolean(data?.es_pro));
  };
const manejarAccesoEmail = (e: React.FormEvent) => {
    e.preventDefault();
    const correoLimpio = inputEmail.trim().toLowerCase();
    if (!correoLimpio || !correoLimpio.includes("@")) {
      setMensajeLogin("Ingresá un correo electrónico válido.");
      return;
    }

    localStorage.setItem("calculopro_email", correoLimpio);
    setEmailUsuario(correoLimpio);

    if (verificarEsPro(correoLimpio)) {
      setEsPro(true);
      setMensajeLogin("¡Licencia Pro activada correctamente!");
    } else {
      setEsPro(false);
      setMensajeLogin("Correo registrado en Plan Gratuito. Si ya compraste Pro, aguardá la activación.");
    }
  };

  const cerrarSesionEmail = () => {
    localStorage.removeItem("calculopro_email");
    setEmailUsuario("");
    setInputEmail("");
    setEsPro(false);
    setMensajeLogin("");
  };
  const abrirModalUpgrade = (motivo: string) => {
    setProMotivo(motivo);
    setShowProModal(true);
  };

  const handleAuthSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthLoading(true);
    setAuthError(null);

    try {
      if (isRegisterMode) {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        setShowAuthModal(false);
        mostrarNotificacion("Cuenta creada e iniciada correctamente.");
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        setShowAuthModal(false);
        mostrarNotificacion("Sesión iniciada correctamente.");
      }
      setEmail("");
      setPassword("");
    } catch (err: unknown) {
      setAuthError(err instanceof Error ? err.message : "Error de autenticación");
    } finally {
      setAuthLoading(false);
    }
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setShowSavedDrawer(false);
    setEsPro(false);
    mostrarNotificacion("Sesión cerrada.");
  };

  const mostrarNotificacion = (texto: string) => {
    setMensajeNube(texto);
    setTimeout(() => setMensajeNube(null), 4500);
  };

  const cargarPlanillasUsuario = async () => {
    const { data, error } = await supabase
      .from("planillas")
      .select("*")
      .order("created_at", { ascending: false });

    if (!error && data) {
      setMisPlanillas(data as SavedPlanilla[]);
    }
  };

  const guardarEnNube = async () => {
    if (!user) {
      setShowAuthModal(true);
      return;
    }

    if (!esPro && misPlanillas.length >= LIMITE_PLANILLAS_GRATIS) {
      abrirModalUpgrade(
        `El plan Gratuito permite guardar hasta ${LIMITE_PLANILLAS_GRATIS} planillas en la nube. Actualiza a CálculoPro para almacenamiento ilimitado.`
      );
      return;
    }

    setGuardandoNube(true);
    const { error } = await supabase.from("planillas").insert({
      user_id: user.id,
      nombre_negocio: nombreNegocio,
      subtitulo_negocio: subtituloNegocio,
      plantilla,
      rubro,
      regimen_fiscal: regimenFiscal,
      moneda,
      color_tema: colorTema,
      incluir_envio: incluirEnvio,
      costo_envio_fijo: costoEnvioFijo,
      filas,
    });

    setGuardandoNube(false);

    if (error) {
      mostrarNotificacion(`Error al guardar: ${error.message}`);
    } else {
      mostrarNotificacion("¡Planilla guardada en la nube!");
      cargarPlanillasUsuario();
    }
  };

  const abrirPlanillaGuardada = (p: SavedPlanilla) => {
    setNombreNegocio(p.nombre_negocio);
    setSubtituloNegocio(p.subtitulo_negocio || "");
    setPlantilla(p.plantilla);
    setRubro(p.rubro);
    setRegimenFiscal(p.regimen_fiscal);
    setMoneda(p.moneda);
    setColorTema(p.color_tema);
    setIncluirEnvio(p.incluir_envio);
    setCostoEnvioFijo(p.costo_envio_fijo);
    setFilas(p.filas);
    setShowSavedDrawer(false);
    mostrarNotificacion(`Planilla "${p.nombre_negocio}" cargada.`);
  };

  const eliminarPlanillaGuardada = async (id: string) => {
    const { error } = await supabase.from("planillas").delete().eq("id", id);
    if (!error) {
      setMisPlanillas(misPlanillas.filter((p) => p.id !== id));
      mostrarNotificacion("Planilla eliminada de la nube.");
    }
  };

// IMPORTACIÓN MASIVA DESDE EXCEL (.xlsx) O CSV (.csv)
  const handleImportarArchivo = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const nuevasFilas: RowData[] = [];
      const extension = file.name.split(".").pop()?.toLowerCase();
      const hoy = new Date().toISOString().split("T")[0];

      if (extension === "csv") {
        const text = await file.text();
        const lineas = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
        const separador = lineas[0]?.includes(";") ? ";" : ",";

        lineas.forEach((linea, index) => {
          const cols = linea.split(separador).map((c) => c.replace(/^"|"$/g, "").trim());
          const numA = parseFloat(cols[1]?.replace(/[^0-9.-]/g, "") || "");
          if (index === 0 && isNaN(numA)) return;

          if (cols[0]) {
            const numB = parseFloat(cols[2]?.replace(/[^0-9.-]/g, "") || "");
            nuevasFilas.push({
              id: Date.now() + index,
              concepto: cols[0],
              detalleRubro: cols[3] || rubroActual.ejemplo || "-",
              valorA: isNaN(numA) ? 0 : numA,
              valorB: isNaN(numB) ? (plantilla === "precios" ? 35 : 1) : numB,
              cantidad: 1,
              stockMinimo: 5,
              fecha: hoy,
              desplegable: configActual.defaultDropdown,
            });
          }
        });
      } else {
        const buffer = await file.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const sheet = workbook.worksheets[0];

        sheet.eachRow((row, rowNumber) => {
          const vals = row.values as (string | number | { result?: number } | null | undefined)[];
          const c1 = vals[1]?.toString().trim() || "";
          const rawA = typeof vals[2] === "object" && vals[2] !== null ? vals[2].result : vals[2];
          const rawB = typeof vals[3] === "object" && vals[3] !== null ? vals[3].result : vals[3];
          const c4 = vals[4]?.toString().trim() || "";

          const numA = parseFloat(String(rawA ?? "").replace(/[^0-9.-]/g, ""));
          if (rowNumber <= 3 && isNaN(numA)) return;

          if (c1 && !isNaN(numA)) {
            const numB = parseFloat(String(rawB ?? "").replace(/[^0-9.-]/g, ""));
            nuevasFilas.push({
              id: Date.now() + rowNumber,
              concepto: c1,
              detalleRubro: c4 || rubroActual.ejemplo || "-",
              valorA: numA,
              valorB: isNaN(numB) ? (plantilla === "precios" ? 35 : 1) : numB,
              cantidad: 1,
              stockMinimo: 5,
              fecha: hoy,
              desplegable: configActual.defaultDropdown,
            });
          }
        });
      }

      if (nuevasFilas.length === 0) {
        mostrarNotificacion("No se detectaron filas válidas. Usa columnas: Producto | Valor 1 | Valor 2.");
      } else if (!esPro && nuevasFilas.length > LIMITE_FILAS_GRATIS) {
        setFilas(nuevasFilas.slice(0, LIMITE_FILAS_GRATIS));
        abrirModalUpgrade(
          `Se importaron las primeras ${LIMITE_FILAS_GRATIS} filas (límite del plan Gratuito). Pásate a Pro para importar catálogos sin límite.`
        );
      } else {
        setFilas(nuevasFilas);
        mostrarNotificacion(`¡${nuevasFilas.length} productos importados correctamente!`);
      }
    } catch {
      mostrarNotificacion("Error al leer el archivo. Verifica que sea un .xlsx o .csv válido.");
    } finally {
      e.target.value = "";
    }
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      setLogoUrl(reader.result as string);
    };
    reader.readAsDataURL(file);
  };

  const cambiarRegimenFiscal = (nuevoRegimen: string) => {
    setRegimenFiscal(nuevoRegimen);
    if (nuevoRegimen.startsWith("BR_")) {
      setMoneda("R$ ");
    } else if (nuevoRegimen === "CO_19") {
      setMoneda("COP ");
    } else if (nuevoRegimen === "PE_18") {
      setMoneda("S/ ");
    } else if (nuevoRegimen === "EXENTO") {
      setMoneda("USD ");
    } else {
      setMoneda("$");
    }
  };

  const getPorcentajeImpuesto = (): number => {
    switch (regimenFiscal) {
      case "AR_21":
        return 21;
      case "AR_105":
        return 10.5;
      case "BR_LUCRO":
      case "PE_18":
        return 18;
      case "BR_SIMPLES":
        return 6;
      case "MX_16":
        return 16;
      case "CL_19":
      case "CO_19":
        return 19;
      default:
        return 0;
    }
  };

const [filas, setFilas] = useState<RowData[]>([
    {
      id: 1,
      concepto: "Zapatillas Urbanas Air Pro",
      detalleRubro: "42 / Negro",
      valorA: 68000,
      valorB: 35000,
      cantidad: 1,
      stockMinimo: 5,
      fecha: "",
      desplegable: "Clásica (13%)",
    },
    {
      id: 2,
      concepto: "Auriculares Inalámbricos ANC",
      detalleRubro: "12 meses / SN-884",
      valorA: 42900,
      valorB: 19500,
      cantidad: 1,
      stockMinimo: 5,
      fecha: "",
      desplegable: "Premium (28%)",
    },
    {
      id: 3,
      concepto: "Mochila Ejecutiva Impermeable",
      detalleRubro: "Estándar / Gris",
      valorA: 31000,
      valorB: 14500,
      cantidad: 1,
      stockMinimo: 5,
      fecha: "",
      desplegable: "Clásica (13%)",
    },
  ]);

  const calcularEstadoStockDinamico = (stockActual: number, stockMin: number): string => {
    if (stockActual <= 0) return "Sin Stock";
    if (stockActual <= stockMin) return "Stock Bajo";
    return "Disponible";
  };

  const cambiarPlantilla = (nueva: TemplateType) => {
    setPlantilla(nueva);
    const nuevaConfig = PLANTILLAS_CONFIG[nueva];
    setFilas(
      filas.map((f) => {
        const nuevoValorB =
          nueva === "precios"
            ? 40
            : nueva === "stock" || nueva === "caja" || nueva === "presupuesto"
            ? 5
            : f.valorB;

        return {
          ...f,
          valorB: nuevoValorB,
          desplegable:
            nueva === "stock"
              ? calcularEstadoStockDinamico(nuevoValorB, f.stockMinimo ?? 5)
              : nuevaConfig.defaultDropdown,
        };
      })
    );
  };

  const agregarFila = () => {
    if (!esPro && filas.length >= LIMITE_FILAS_GRATIS) {
      abrirModalUpgrade(
        `Has alcanzado el límite de ${LIMITE_FILAS_GRATIS} filas del plan Gratuito. Desbloquea filas ilimitadas con MargenReal Pro.`
      );
      return;
    }

    setFilas([
      ...filas,
      {
        id: Date.now(),
        concepto: "Nuevo producto o concepto",
        detalleRubro: rubroActual.ejemplo || "-",
        valorA: 20000,
        valorB: plantilla === "precios" ? 35 : 2,
        cantidad: 1,
        stockMinimo: 5,
        fecha: new Date().toISOString().split("T")[0],
        desplegable: configActual.defaultDropdown,
      },
    ]);
  };

  const eliminarFila = (id: number) => {
    setFilas(filas.filter((f) => f.id !== id));
  };

  const registrarVentaItem = (item: RowData, unidadesVendidas: number = 1) => {
    if (item.valorB <= 0) {
      mostrarNotificacion(`"${item.concepto}" no tiene stock disponible.`);
      return;
    }

    const cantidadFinal = Math.min(unidadesVendidas, item.valorB);
    const nuevoStock = Math.max(0, item.valorB - cantidadFinal);
    const nuevoEstado =
      item.desplegable === "En Camino"
        ? "En Camino"
        : calcularEstadoStockDinamico(nuevoStock, item.stockMinimo ?? 5);

    setFilas(
      filas.map((f) =>
        f.id === item.id
          ? { ...f, valorB: nuevoStock, desplegable: nuevoEstado }
          : f
      )
    );

    if (user) {
      const hoy = new Date().toISOString().split("T")[0];
      const movimientoCaja: RowData = {
        id: Date.now(),
        concepto: `Venta: ${item.concepto}`,
        detalleRubro: item.detalleRubro,
        valorA: item.valorA,
        valorB: cantidadFinal,
        cantidad: cantidadFinal,
        stockMinimo: 0,
        fecha: hoy,
        desplegable: "Efectivo",
      };

      const planillaCajaExistente = misPlanillas.find((p) => p.plantilla === "caja");
      if (planillaCajaExistente) {
        supabase
          .from("planillas")
          .update({ filas: [movimientoCaja, ...planillaCajaExistente.filas] })
          .eq("id", planillaCajaExistente.id)
          .then(() => cargarPlanillasUsuario());
      }
    }

    mostrarNotificacion(
      `Venta impactada: -${cantidadFinal} u. de "${item.concepto}" (Stock restante: ${nuevoStock}).`
    );
  };

  const actualizarFila = (id: number, campo: keyof RowData, valor: string | number) => {
    setFilas(filas.map((f) => (f.id === id ? { ...f, [campo]: valor } : f)));
  };

  const calcularDetalleImpuesto = (f: RowData) => {
    const imp = getPorcentajeImpuesto();
    const factor = imp / 100;

    switch (plantilla) {
      case "mercadolibre": {
        const montoImpuesto = imp > 0 ? (f.valorA * imp) / (100 + imp) : 0;
        return { montoImpuesto, precioUnitarioFinal: f.valorA };
      }
      case "presupuesto": {
        const subtotalNeto = f.valorA * f.valorB;
        const montoImpuesto = subtotalNeto * factor;
        return { montoImpuesto, precioUnitarioFinal: f.valorA };
      }
      case "stock":
      case "caja":
      case "gastos": {
        const base = f.valorA * f.valorB;
        return { montoImpuesto: base * factor, precioUnitarioFinal: f.valorA };
      }
      case "cuentas": {
        return { montoImpuesto: f.valorA * factor, precioUnitarioFinal: f.valorA };
      }
      case "precios": {
        const conMargen = f.valorA * (1 + f.valorB / 100);
        return { montoImpuesto: conMargen * factor, precioUnitarioFinal: f.valorA };
      }
      case "pedidos": {
        const cant = f.cantidad || 1;
        const subtotalPedido = f.valorA * cant;
        return { montoImpuesto: subtotalPedido * factor, precioUnitarioFinal: f.valorA };
      }
      default:
        return { montoImpuesto: 0, precioUnitarioFinal: f.valorA };
    }
  };

  const calcularResultadoFila = (f: RowData): number => {
    const { montoImpuesto } = calcularDetalleImpuesto(f);

    switch (plantilla) {
      case "mercadolibre": {
        const porcML = f.desplegable.includes("28%") ? 0.28 : 0.13;
        const comision = f.valorA * porcML;
        const envio = incluirEnvio ? costoEnvioFijo : 0;
        return f.valorA - f.valorB - comision - envio - montoImpuesto;
      }
      case "stock":
      case "caja":
      case "gastos":
        return f.valorA * f.valorB - montoImpuesto;
      case "presupuesto":
        // Subtotal Neto (el Total Final con impuesto se calcula sumando montoImpuesto)
        return f.valorA * f.valorB;
      case "cuentas":
        return f.valorA - f.valorB - montoImpuesto;
      case "precios": {
        const conMargen = f.valorA * (1 + f.valorB / 100);
        return conMargen + montoImpuesto;
      }
      case "pedidos": {
        const cant = f.cantidad || 1;
        return f.valorA * cant + f.valorB + montoImpuesto;
      }
      default:
        return 0;
    }
  };

  const exportarExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Planilla Gestión");
    const imp = getPorcentajeImpuesto();
    const tieneColumnaRubro = rubro !== "general";
    const esPresupuesto = plantilla === "presupuesto";
    const esPedidos = plantilla === "pedidos";
    const esSumaImpuesto = esPresupuesto || plantilla === "precios" || esPedidos;

    // Calcula la última letra según columnas extra (Rubro, Cantidad en Pedidos, Total Final en Presupuestos)
    const columnasExtra =
      (tieneColumnaRubro ? 1 : 0) + (esPedidos ? 1 : 0) + (esPresupuesto ? 1 : 0);
    const ultimaLetra = String.fromCharCode(70 + columnasExtra); // F, G o H

    sheet.mergeCells(`A1:${ultimaLetra}1`);
    const titleCell = sheet.getCell("A1");
    titleCell.value = `${nombreNegocio.toUpperCase()} — ${configActual.tituloExcel}`;
    titleCell.font = { size: 13, bold: true, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: `FF${colorTema}` },
    };
    titleCell.alignment = { vertical: "middle", horizontal: "center" };
    sheet.getRow(1).height = 34;

    sheet.mergeCells(`A2:${ultimaLetra}2`);
    const subCell = sheet.getCell("A2");
    subCell.value = `${subtituloNegocio} | Rubro: ${rubroActual.label} | Impuesto: ${imp}%`;
    subCell.font = { size: 9.5, italic: true, color: { argb: "FF475569" } };
    subCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(2).height = 22;

    if (logoUrl) {
      try {
        const ext = logoUrl.includes("image/png") ? "png" : "jpeg";
        const imageId = workbook.addImage({
          base64: logoUrl,
          extension: ext,
        });
        sheet.addImage(imageId, {
          tl: { col: 0.1, row: 0.1 },
          ext: { width: 38, height: 38 },
        });
      } catch {
        // Continuar si el formato de imagen no es compatible
      }
    }

    const headers: string[] = ["Producto / Concepto"];
    if (tieneColumnaRubro) headers.push(rubroActual.columnaExtra);
    headers.push(configActual.colA);
    if (esPedidos) headers.push("Cantidad");
    headers.push(
      configActual.colB,
      configActual.colDropdown,
      esSumaImpuesto ? `Impuesto (${imp}%)` : `Desc. Impuesto (${imp}%)`,
      configActual.colResultado
    );
    if (esPresupuesto) headers.push("Total Final (c/ Imp.)");

    const headerRow = sheet.addRow(headers);
    headerRow.height = 24;
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1E293B" },
    };
    headerRow.alignment = { vertical: "middle" };

    const offsetRubro = tieneColumnaRubro ? 1 : 0;
    const offsetTotal = offsetRubro + (esPedidos ? 1 : 0);

    const colALetter = String.fromCharCode(66 + offsetRubro);
    const colCantLetter = String.fromCharCode(67 + offsetRubro);
    const colBLetter = String.fromCharCode(67 + offsetTotal);
    const colDropLetter = String.fromCharCode(68 + offsetTotal);
    const colImpLetter = String.fromCharCode(69 + offsetTotal);
    const colResLetter = String.fromCharCode(70 + offsetTotal);

    filas.forEach((f, idx) => {
      const r = idx + 4;
      const rowValues: (string | number | { formula: string })[] = [f.concepto];
      if (tieneColumnaRubro) rowValues.push(f.detalleRubro);

      rowValues.push(f.valorA);
      if (esPedidos) rowValues.push(f.cantidad || 1);

      let formulaImpuesto = "0";
      let formulaResultado = "";

      switch (plantilla) {
        case "mercadolibre": {
          const envioVal = incluirEnvio ? costoEnvioFijo : 0;
          formulaImpuesto = imp > 0 ? `${colALetter}${r}*${imp}/${100 + imp}` : "0";
          formulaResultado = `${colALetter}${r}-${colBLetter}${r}-IF(ISNUMBER(SEARCH("28%",${colDropLetter}${r})),${colALetter}${r}*0.28,${colALetter}${r}*0.13)-${envioVal}-${colImpLetter}${r}`;
          break;
        }
        case "stock":
        case "caja":
        case "gastos":
          formulaImpuesto = `(${colALetter}${r}*${colBLetter}${r})*${imp / 100}`;
          formulaResultado = `(${colALetter}${r}*${colBLetter}${r})-${colImpLetter}${r}`;
          break;
        case "presupuesto":
          formulaImpuesto = `(${colALetter}${r}*${colBLetter}${r})*${imp / 100}`;
          formulaResultado = `${colALetter}${r}*${colBLetter}${r}`;
          break;
        case "cuentas":
          formulaImpuesto = `${colALetter}${r}*${imp / 100}`;
          formulaResultado = `${colALetter}${r}-${colBLetter}${r}-${colImpLetter}${r}`;
          break;
        case "precios":
          formulaImpuesto = `(${colALetter}${r}*(1+${colBLetter}${r}/100))*${imp / 100}`;
          formulaResultado = `(${colALetter}${r}*(1+${colBLetter}${r}/100))+${colImpLetter}${r}`;
          break;
        case "pedidos":
          formulaImpuesto = `(${colALetter}${r}*${colCantLetter}${r})*${imp / 100}`;
          formulaResultado = `(${colALetter}${r}*${colCantLetter}${r})+${colBLetter}${r}+${colImpLetter}${r}`;
          break;
      }

      rowValues.push(
        f.valorB,
        f.desplegable,
        { formula: formulaImpuesto },
        { formula: formulaResultado }
      );

      if (esPresupuesto) {
        rowValues.push({ formula: `${colResLetter}${r}+${colImpLetter}${r}` });
      }

      const addedRow = sheet.addRow(rowValues);
      addedRow.height = 22;

      const dropdownCell = addedRow.getCell(4 + offsetTotal);
      dropdownCell.dataValidation = {
        type: "list",
        allowBlank: false,
        formulae: [`"${configActual.opcionesDropdown.join(",")}"`],
      };

      addedRow.getCell(2 + offsetRubro).numFmt = `"${moneda}"#,##0.00`;
      if (
        plantilla !== "stock" &&
        plantilla !== "caja" &&
        plantilla !== "presupuesto" &&
        plantilla !== "precios"
      ) {
        addedRow.getCell(3 + offsetTotal).numFmt = `"${moneda}"#,##0.00`;
      }
      addedRow.getCell(5 + offsetTotal).numFmt = `"${moneda}"#,##0.00`;
      addedRow.getCell(6 + offsetTotal).numFmt = `"${moneda}"#,##0.00`;

      if (esPresupuesto) {
        addedRow.getCell(7 + offsetTotal).numFmt = `"${moneda}"#,##0.00`;
        addedRow.getCell(7 + offsetTotal).font = { bold: true };
      } else {
        addedRow.getCell(6 + offsetTotal).font = { bold: true };
      }
    });

    sheet.columns.forEach((col) => {
      col.width = 24;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${nombreNegocio.replace(/\s+/g, "_")}_${plantilla}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const totalGeneral = filas.reduce((acc, f) => {
    const base = calcularResultadoFila(f);
    if (plantilla === "presupuesto") {
      return acc + base + calcularDetalleImpuesto(f).montoImpuesto;
    }
    return acc + base;
  }, 0);
  const promedioPorFila = filas.length > 0 ? totalGeneral / filas.length : 0;
  const impActual = getPorcentajeImpuesto();

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* NOTIFICACIÓN FLOTANTE */}
      {mensajeNube && (
        <div className="fixed bottom-5 right-5 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-lg text-sm font-medium flex items-center gap-2 border border-slate-700">
          <Cloud className="w-4 h-4 text-emerald-400" />
          <span>{mensajeNube}</span>
        </div>
      )}

      {/* MODAL DE SUSCRIPCIÓN PRO (MONETIZACIÓN) */}
      {showProModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-lg w-full overflow-hidden">
            <div className="bg-gradient-to-r from-amber-500 via-orange-500 to-amber-600 p-6 text-white flex justify-between items-start">
              <div>
                <div className="inline-flex items-center gap-1.5 bg-white/20 px-2.5 py-1 rounded-full text-xs font-bold mb-2">
                  <Crown className="w-3.5 h-3.5" /> CÁLCULOPRO ILIMITADO
                </div>
                <h3 className="font-extrabold text-2xl">Escala tu gestión comercial</h3>
                <p className="text-xs text-amber-50 mt-1">
                  {proMotivo || "Desbloquea todo el potencial para tu empresa sin restricciones."}
                </p>
              </div>
              <button
                onClick={() => setShowProModal(false)}
                className="text-white/80 hover:text-white cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 space-y-5">
              <div className="flex items-baseline justify-between border-b pb-4">
                <div>
                  <span className="text-3xl font-extrabold text-slate-900">$25000 ARS</span>
                  <span className="text-xs text-slate-500 font-medium"> / mes (o equivalente en moneda local)</span>
                </div>
                <span className="text-xs font-semibold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-full">
                  Cancela cuando quieras
                </span>
              </div>

              <ul className="space-y-2.5 text-sm text-slate-700">
                <li className="flex items-center gap-2">
                  <Check className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span><strong>Filas e ítems ilimitados</strong> en todas tus planillas</span>
                </li>
                <li className="flex items-center gap-2">
                  <Check className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span><strong>Almacenamiento ilimitado</strong> en la nube de tus modelos</span>
                </li>
                <li className="flex items-center gap-2">
                  <Check className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span><strong>Importación masiva</strong> de catálogos completos desde Excel y CSV</span>
                </li>
                <li className="flex items-center gap-2">
                  <Check className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span>Exportación Excel con fórmulas vivas, desplegables y logo de marca</span>
                </li>
              </ul>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                <a
                  href={LINK_PAGO_MERCADOPAGO}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="bg-sky-500 hover:bg-sky-600 text-white font-bold text-sm py-3 px-4 rounded-xl text-center transition shadow-xs"
                >
                  Pagar con Mercado Pago
                </a>
                <a
                  href={LINK_PAGO_STRIPE}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm py-3 px-4 rounded-xl text-center transition shadow-xs"
                >
                  Tarjeta Internacional (Stripe)
                </a>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL DE INICIO DE SESIÓN / REGISTRO */}
      {showAuthModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl border border-slate-200 max-w-md w-full p-6 space-y-5">
            <div className="flex justify-between items-center border-b pb-3">
              <div className="flex items-center gap-2">
                <Cloud className="w-5 h-5 text-blue-600" />
                <h3 className="font-bold text-lg text-slate-900">
                  {isRegisterMode ? "Crear Cuenta en CálculoPro" : "Iniciar Sesión"}
                </h3>
              </div>
              <button
                onClick={() => setShowAuthModal(false)}
                className="text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <p className="text-xs text-slate-500">
              Guarda tus planillas personalizadas en la nube y accede a ellas desde cualquier dispositivo.
            </p>

            {authError && (
              <div className="p-3 rounded-lg bg-red-50 border border-red-200 text-xs text-red-700 font-medium">
                {authError}
              </div>
            )}

            <form onSubmit={handleAuthSubmit} className="space-y-4">
              <div>
                <label className="block text-xs font-bold uppercase text-slate-600 mb-1">
                  Correo Electrónico
                </label>
                <input
                  type="email"
                  required
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="tu@empresa.com"
                  className="w-full px-3 py-2 text-sm border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>

              <div>
                <label className="block text-xs font-bold uppercase text-slate-600 mb-1">
                  Contraseña (mínimo 6 caracteres)
                </label>
                <input
                  type="password"
                  required
                  minLength={6}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  className="w-full px-3 py-2 text-sm border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>

              <button
                type="submit"
                disabled={authLoading}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2.5 rounded-xl text-sm flex items-center justify-center gap-2 transition cursor-pointer disabled:opacity-60"
              >
                {authLoading && <Loader2 className="w-4 h-4 animate-spin" />}
                {isRegisterMode ? "Registrarme Gratis" : "Entrar a mi Cuenta"}
              </button>
            </form>

            <div className="text-center pt-2 border-t">
              <button
                type="button"
                onClick={() => {
                  setIsRegisterMode(!isRegisterMode);
                  setAuthError(null);
                }}
                className="text-xs font-semibold text-blue-600 hover:underline cursor-pointer"
              >
                {isRegisterMode
                  ? "¿Ya tienes cuenta? Inicia sesión aquí"
                  : "¿No tienes cuenta? Créala gratis aquí"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* PANEL DESLIZABLE DE PLANILLAS GUARDADAS */}
      {showSavedDrawer && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex justify-end">
          <div className="bg-white w-full max-w-md h-full shadow-2xl border-l border-slate-200 p-6 flex flex-col">
            <div className="flex justify-between items-center border-b pb-4">
              <div className="flex items-center gap-2">
                <FolderOpen className="w-5 h-5 text-blue-600" />
                <h3 className="font-bold text-lg text-slate-900">Mis Planillas en la Nube</h3>
              </div>
              <button
                onClick={() => setShowSavedDrawer(false)}
                className="text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {!esPro && (
              <div className="mt-3 p-3 rounded-xl bg-amber-50 border border-amber-200 flex items-center justify-between gap-2">
                <span className="text-xs font-medium text-amber-900">
                  Plan Gratis: {misPlanillas.length}/{LIMITE_PLANILLAS_GRATIS} planillas usadas
                </span>
                <button
                  onClick={() => {
                    setShowSavedDrawer(false);
                    abrirModalUpgrade("Actualiza a Pro para guardar planillas ilimitadas.");
                  }}
                  className="text-xs font-bold text-amber-800 underline cursor-pointer"
                >
                  Ampliar
                </button>
              </div>
            )}

            <div className="flex-1 overflow-y-auto py-4 space-y-3">
              {misPlanillas.length === 0 ? (
                <div className="text-center py-12 text-slate-400 text-sm">
                  Aún no tienes planillas guardadas. Haz clic en &quot;Guardar en Nube&quot; para almacenar la actual.
                </div>
              ) : (
                misPlanillas.map((p) => (
                  <div
                    key={p.id}
                    className="p-4 rounded-xl border border-slate-200 hover:border-blue-400 bg-slate-50/60 flex items-center justify-between gap-3 transition"
                  >
                    <div
                      onClick={() => abrirPlanillaGuardada(p)}
                      className="flex-1 cursor-pointer"
                    >
                      <h4 className="font-bold text-sm text-slate-900">{p.nombre_negocio}</h4>
                      <p className="text-xs text-blue-700 font-medium mt-0.5">
                        {PLANTILLAS_CONFIG[p.plantilla]?.label || p.plantilla}
                      </p>
                      <p className="text-[11px] text-slate-400 mt-1">
                        {p.filas?.length || 0} filas • {new Date(p.created_at).toLocaleDateString()}
                      </p>
                    </div>
                    <button
                      onClick={() => eliminarPlanillaGuardada(p.id)}
                      className="text-slate-400 hover:text-red-600 p-2 rounded-lg hover:bg-red-50 cursor-pointer"
                      title="Eliminar planilla"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}

      {/* BARRA SUPERIOR */}
      <header className="bg-white border-b border-slate-200 px-6 py-3.5 flex flex-wrap items-center justify-between gap-4 sticky top-0 z-20 print:hidden">

       
        <div className="flex items-center gap-3">
          <div className="bg-gradient-to-br from-blue-600 to-indigo-700 p-2.5 rounded-xl text-white shadow-sm">
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="font-bold text-lg tracking-tight text-slate-900">CálculoPro</h1>
              {esPro ? (
                <span className="inline-flex items-center gap-1 text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 px-2 py-0.5 rounded-full">
                  <Crown className="w-3 h-3" /> PRO
                </span>
              ) : (
                <button
                  onClick={() => abrirModalUpgrade("")}
                  className="inline-flex items-center gap-1 text-[11px] font-bold bg-gradient-to-r from-amber-500 to-orange-500 text-white px-2.5 py-0.5 rounded-full hover:opacity-95 transition cursor-pointer"
                >
                  <Sparkles className="w-3 h-3" /> Pasar a Pro
                </button>
              )}
            </div>
            <p className="text-xs text-slate-500">Generador inteligente de control y gestión comercial</p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          {/* Importar CSV / Excel */}
          <input
            ref={importExcelInputRef}
            type="file"
            accept=".xlsx,.csv"
            onChange={handleImportarArchivo}
            className="hidden"
          />
          {/* INICIO BLOQUE ACCESO POR CORREO */}
            <div className="flex flex-col items-end">
              <div className="flex items-center gap-2 bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm">
                {emailUsuario ? (
                  <>
                    <span
                      className="text-slate-700 font-medium truncate max-w-[160px]"
                      title={emailUsuario}
                    >
                      {emailUsuario}
                    </span>
                    <span
                      className={`px-2 py-0.5 rounded-full text-xs font-bold ${
                        esPro
                          ? "bg-emerald-100 text-emerald-700 border border-emerald-300"
                          : "bg-amber-100 text-amber-700 border border-amber-300"
                      }`}
                    >
                      {esPro ? "PRO" : "GRATIS"}
                    </span>
                    <button
                      type="button"
                      onClick={cerrarSesionEmail}
                      className="text-xs text-slate-400 hover:text-red-600 ml-1 underline"
                      title="Cambiar correo"
                    >
                      Salir
                    </button>
                  </>
                ) : (
                  <form onSubmit={manejarAccesoEmail} className="flex items-center gap-1.5">
                    <input
                      type="email"
                      placeholder="Tu correo electrónico..."
                      value={inputEmail}
                      onChange={(e) => setInputEmail(e.target.value)}
                      className="px-2 py-1 text-xs border border-slate-300 rounded bg-white text-slate-800 focus:outline-none focus:border-blue-600 w-44"
                      required
                    />
                    <button
                      type="submit"
                      className="px-2.5 py-1 text-xs font-semibold bg-blue-700 text-white rounded hover:bg-blue-800 transition"
                    >
                      Activar
                    </button>
                  </form>
                )}
              </div>
              {mensajeLogin && (
                <span className={`text-[11px] mt-1 font-medium ${esPro ? "text-emerald-700" : "text-amber-700"}`}>
                  {mensajeLogin}
                </span>
              )}
            </div>
            {/* FIN BLOQUE ACCESO POR CORREO */}
          <button
            onClick={() => importExcelInputRef.current?.click()}
            title="Importar productos desde Excel (.xlsx) o CSV"
            className="flex items-center gap-1.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 font-medium text-sm px-3 py-2.5 rounded-xl transition cursor-pointer"
          >
            <FileUp className="w-4 h-4 text-indigo-600" />
            <span className="hidden xl:inline">Importar Excel/CSV</span>
          </button>

          {/* Botones de Nube y Usuario */}
          <button
            onClick={guardarEnNube}
            disabled={guardandoNube}
            className="flex items-center gap-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 font-semibold text-sm px-3.5 py-2.5 rounded-xl transition cursor-pointer"
          >
            {guardandoNube ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Save className="w-4 h-4" />
            )}
            <span>Guardar</span>
          </button>

          {user ? (
            <>
              <button
                onClick={() => {
                  cargarPlanillasUsuario();
                  setShowSavedDrawer(true);
                }}
                className="flex items-center gap-1.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 font-medium text-sm px-3.5 py-2.5 rounded-xl transition cursor-pointer"
              >
                <FolderOpen className="w-4 h-4 text-blue-600" />
                <span>Mis Planillas ({misPlanillas.length})</span>
              </button>

              <button
                onClick={handleLogout}
                title={`Cerrar sesión (${user.email})`}
                className="flex items-center gap-1.5 bg-white hover:bg-red-50 text-slate-600 hover:text-red-600 border border-slate-200 text-sm px-3 py-2.5 rounded-xl transition cursor-pointer"
              >
                <LogOut className="w-4 h-4" />
              </button>
            </>
          ) : (
            <button
              onClick={() => setShowAuthModal(true)}
              className="flex items-center gap-1.5 bg-slate-900 hover:bg-slate-800 text-white font-medium text-sm px-4 py-2.5 rounded-xl transition cursor-pointer"
            >
              <LogIn className="w-4 h-4" />
              <span>Iniciar Sesión</span>
            </button>
          )}

          <div className="h-6 w-px bg-slate-200 hidden sm:block" />

          <button
            onClick={() => window.print()}
            className="flex items-center gap-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 font-medium text-sm px-3 py-2.5 rounded-xl transition cursor-pointer"
          >
            <Printer className="w-4 h-4 text-slate-500" />
            <span className="hidden md:inline">PDF</span>
          </button>

          <button
            onClick={exportarExcel}
            className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-sm px-4 py-2.5 rounded-xl shadow-sm transition cursor-pointer"
          >
            <Download className="w-4 h-4" />
            <span>Excel (.xlsx)</span>
          </button>
        </div>
      </header>

      <div className="flex-1 flex flex-col lg:flex-row">
        {/* PANEL IZQUIERDO */}
        <aside className="w-full lg:w-[380px] bg-white border-r border-slate-200 p-6 flex flex-col gap-5 shrink-0 print:hidden">
          <div className="flex items-center gap-2 text-slate-800 font-semibold text-sm border-b border-slate-100 pb-3">
            <SlidersHorizontal className="w-4 h-4 text-blue-600" />
            <span>Personalización de la Planilla</span>
          </div>

          {/* Identidad de Marca */}
          <div className="bg-slate-50 p-4 rounded-xl border border-slate-200/80 space-y-3">
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500">
              Identidad de tu Negocio
            </label>

            <div className="flex items-center gap-3">
              <input
                ref={fileInputRef}
                type="file"
                accept="image/png, image/jpeg"
                onChange={handleLogoUpload}
                className="hidden"
              />
              {logoUrl ? (
                <div className="relative w-12 h-12 rounded-lg border bg-white p-1 flex items-center justify-center shrink-0">
                  <img src={logoUrl} alt="Logo" className="max-h-full max-w-full object-contain" />
                  <button
                    onClick={() => setLogoUrl(null)}
                    className="absolute -top-1.5 -right-1.5 bg-red-500 text-white rounded-full p-0.5 hover:bg-red-600 cursor-pointer"
                    title="Quitar logo"
                  >
                    <X className="w-3 h-3" />
                  </button>
                </div>
              ) : (
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="w-12 h-12 rounded-lg border-2 border-dashed border-slate-300 hover:border-blue-500 bg-white flex flex-col items-center justify-center text-slate-400 hover:text-blue-600 transition shrink-0 cursor-pointer"
                  title="Subir logotipo"
                >
                  <Upload className="w-4 h-4" />
                  <span className="text-[9px] font-medium mt-0.5">Logo</span>
                </button>
              )}

              <div className="flex-1 space-y-1.5">
                <div className="relative">
                  <Building2 className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
                  <input
                    type="text"
                    value={nombreNegocio}
                    onChange={(e) => setNombreNegocio(e.target.value)}
                    placeholder="Nombre del negocio"
                    className="w-full pl-8 pr-2.5 py-1.5 text-sm font-medium border border-slate-300 rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                </div>
                <input
                  type="text"
                  value={subtituloNegocio}
                  onChange={(e) => setSubtituloNegocio(e.target.value)}
                  placeholder="Datos fiscales / contacto"
                  className="w-full px-2.5 py-1 text-xs border border-slate-200 rounded-lg bg-white text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
            </div>
          </div>

          {/* 1. Modelo de Planilla */}
          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">
              1. Modelo de Planilla
            </label>
            <select
              value={plantilla}
              onChange={(e) => cambiarPlantilla(e.target.value as TemplateType)}
              className="w-full p-2.5 text-sm border border-slate-300 rounded-xl bg-white font-medium text-slate-800 shadow-2xs focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
            >
              {Object.entries(PLANTILLAS_CONFIG).map(([key, cfg]) => (
                <option key={key} value={key}>
                  {cfg.label}
                </option>
              ))}
            </select>
            <p className="text-xs text-slate-500 mt-1.5 leading-relaxed">
              {configActual.descripcion}
            </p>
          </div>

          {/* 2. Rubro */}
          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">
              2. Rubro / Sector Comercial
            </label>
            <select
              value={rubro}
              onChange={(e) => setRubro(e.target.value as RubroType)}
              className="w-full p-2.5 text-sm border border-slate-300 rounded-xl bg-white font-medium text-slate-800 shadow-2xs focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
            >
              {Object.entries(RUBROS_CONFIG).map(([key, cfg]) => (
                <option key={key} value={key}>
                  {cfg.label} {cfg.columnaExtra ? `(+${cfg.columnaExtra})` : ""}
                </option>
              ))}
            </select>
          </div>

          {/* 3. País y Régimen Fiscal */}
          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">
              3. País y Régimen Fiscal
            </label>
            <select
              value={regimenFiscal}
              onChange={(e) => cambiarRegimenFiscal(e.target.value)}
              className="w-full p-2.5 text-sm border border-slate-300 rounded-xl bg-white font-medium text-slate-800 shadow-2xs focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
            >
              <optgroup label="Argentina">
                <option value="AR_21">Argentina — Resp. Inscripto (IVA 21%)</option>
                <option value="AR_105">Argentina — Alícuota Reducida (IVA 10.5%)</option>
                <option value="AR_MONO">Argentina — Monotributo (0% desglose)</option>
              </optgroup>
              <optgroup label="Brasil">
                <option value="BR_SIMPLES">Brasil — Simples Nacional (DAS ~6%)</option>
                <option value="BR_LUCRO">Brasil — Lucro Presumido / Real (~18%)</option>
                <option value="BR_MEI">Brasil — MEI Microempreendedor (0%)</option>
              </optgroup>
              <optgroup label="México, Chile, Colombia y Perú">
                <option value="MX_16">México — General / RESICO (IVA 16%)</option>
                <option value="CL_19">Chile — Régimen General (IVA 19%)</option>
                <option value="CO_19">Colombia — Régimen Común (IVA 19%)</option>
                <option value="PE_18">Perú — Régimen MYPE / General (IGV 18%)</option>
              </optgroup>
              <optgroup label="Internacional">
                <option value="EXENTO">Exento / Exportación USD (0%)</option>
              </optgroup>
            </select>
          </div>

          {/* 4. Moneda y Color Marca */}
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">
                4. Moneda
              </label>
              <select
                value={moneda}
                onChange={(e) => setMoneda(e.target.value)}
                className="w-full p-2.5 text-sm border border-slate-300 rounded-xl bg-white font-medium cursor-pointer"
              >
                <option value="$">$ (ARS/MXN/CLP)</option>
                <option value="R$ ">R$ (Brasil)</option>
                <option value="USD ">USD (Dólares)</option>
                <option value="COP ">COP (Colombia)</option>
                <option value="S/ ">S/ (Perú)</option>
              </select>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">
                5. Color Marca
              </label>
              <div className="flex items-center gap-2 pt-1.5">
                {TEMAS_VISUALES.map((tema) => (
                  <button
                    key={tema.id}
                    onClick={() => setColorTema(tema.id)}
                    title={tema.nombre}
                    className={`w-6 h-6 rounded-full ${tema.bgClass} transition cursor-pointer ${
                      colorTema === tema.id ? `ring-2 ring-offset-2 ${tema.ringClass} scale-110` : "opacity-75 hover:opacity-100"
                    }`}
                  />
                ))}
              </div>
            </div>
          </div>

          {/* Opciones extra Mercado Libre */}
          {plantilla === "mercadolibre" && (
            <div className="p-3.5 bg-amber-50/80 border border-amber-200 rounded-xl space-y-2.5">
              <label className="flex items-center gap-2 text-sm font-semibold text-amber-950 cursor-pointer">
                <input
                  type="checkbox"
                  checked={incluirEnvio}
                  onChange={(e) => setIncluirEnvio(e.target.checked)}
                  className="w-4 h-4 rounded text-amber-600 focus:ring-amber-500"
                />
                Descontar Mercado Envíos
              </label>
              {incluirEnvio && (
                <div>
                  <span className="text-xs font-medium text-amber-800">
                    Costo envío promedio por venta ({moneda}):
                  </span>
                  <input
                    type="number"
                    value={costoEnvioFijo}
                    onChange={(e) => setCostoEnvioFijo(Number(e.target.value))}
                    className="w-full mt-1 p-2 text-sm border border-amber-300 rounded-lg bg-white font-medium"
                  />
                </div>
              )}
            </div>
          )}
        </aside>

        {/* PANEL DERECHO */}
        <main className="flex-1 p-6 md:p-8 overflow-x-auto">
          <div className="max-w-5xl mx-auto space-y-6">
            {/* Tarjetas KPI */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 print:hidden">
              <div className="bg-white p-4 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-3.5">
                <div className="p-2.5 rounded-xl bg-blue-50 text-blue-600">
                  <Layers className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-xs font-medium text-slate-500">Ítems en Planilla</p>
                  <p className="text-xl font-bold text-slate-900">
                    {filas.length} {!esPro && `/ ${LIMITE_FILAS_GRATIS}`}
                  </p>
                </div>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-3.5">
                <div className="p-2.5 rounded-xl bg-indigo-50 text-indigo-600">
                  <TrendingUp className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-xs font-medium text-slate-500">Promedio por Ítem</p>
                  <p className="text-xl font-bold text-slate-900">
                    {moneda}
                    {promedioPorFila.toLocaleString("es-AR", {
                      minimumFractionDigits: 2,
                      maximumFractionDigits: 2,
                    })}
                  </p>
                </div>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200/90 shadow-2xs flex items-center gap-3.5">
                <div
                  className={`p-2.5 rounded-xl ${
                    totalGeneral < 0 ? "bg-red-50 text-red-600" : "bg-emerald-50 text-emerald-600"
                  }`}
                >
                  <Calculator className="w-5 h-5" />
                </div>
                <div>
                  <p className="text-xs font-medium text-slate-500">{configActual.colResultado} Total</p>
                  <p
                    className={`text-xl font-bold ${
                      totalGeneral < 0 ? "text-red-600" : "text-emerald-700"
                    }`}
                  >
                    {moneda}
                    {totalGeneral.toLocaleString("es-AR", {
                      minimumFractionDigits: 2,
                      maximumFractionDigits: 2,
                    })}
                  </p>
                </div>
              </div>
            </div>

            {/* DOCUMENTO / PLANILLA VISUAL */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="p-6 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-white">
                <div className="flex items-center gap-4">
                  {logoUrl && (
                    <img
                      src={logoUrl}
                      alt="Logo Empresa"
                      className="w-14 h-14 object-contain rounded-lg border border-slate-100 p-1"
                    />
                  )}
                  <div>
                    <h2 className="text-xl font-bold text-slate-900">{nombreNegocio}</h2>
                    <p className="text-xs text-slate-500 mt-0.5">{subtituloNegocio}</p>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2 text-xs">
                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full font-medium bg-slate-100 text-slate-700">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                    {rubroActual.label}
                  </span>
                  <span className="px-2.5 py-1 rounded-full font-medium bg-blue-50 text-blue-700">
                    Impuesto: {impActual}%
                  </span>
                </div>
              </div>

              <div
                style={{ backgroundColor: `#${colorTema}` }}
                className="px-6 py-3 text-white font-semibold text-xs tracking-wider uppercase flex justify-between items-center"
              >
                <span>{configActual.tituloExcel}</span>
                <span className="opacity-80 hidden sm:inline">Fórmulas activas en tiempo real</span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-sm text-left">
                  <thead className="bg-slate-50 text-slate-600 text-xs uppercase font-bold border-b border-slate-200">
                    <tr>
                      <th className="py-3.5 px-4">Producto / Concepto</th>
                      {rubro !== "general" && (
                        <th className="py-3.5 px-4 text-blue-700">{rubroActual.columnaExtra}</th>
                      )}
                      <th className="py-3.5 px-4">{configActual.colA}</th>
                      {plantilla === "pedidos" && (
                        <th className="py-3.5 px-4">Cantidad</th>
                      )}
                      <th className="py-3.5 px-4">{configActual.colB}</th>
                      <th className="py-3.5 px-4">{configActual.colDropdown}</th>
                      <th className="py-3.5 px-4 text-right">
                        {plantilla === "presupuesto" ||
                        plantilla === "precios" ||
                        plantilla === "pedidos"
                          ? `Impuesto (${impActual}%)`
                          : `Desc. Impuesto (${impActual}%)`}
                      </th>
                      <th className="py-3.5 px-4 text-right">{configActual.colResultado}</th>
                      {plantilla === "presupuesto" && (
                        <th className="py-3.5 px-4 text-right text-blue-800">
                          Total Final (c/ Imp.)
                        </th>
                      )}
                      <th className="py-3.5 px-3 w-10 print:hidden"></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filas.map((f) => {
                      const resultado = calcularResultadoFila(f);
                      const { montoImpuesto } = calcularDetalleImpuesto(f);
                      const esPresupuesto = plantilla === "presupuesto";
                      const esPedidos = plantilla === "pedidos";
                      const esCaja = plantilla === "caja";
                      const esStock = plantilla === "stock";
                      const esSumaImpuesto =
                        esPresupuesto || plantilla === "precios" || esPedidos;
                      const totalFinalPresupuesto = resultado + montoImpuesto;

                      return (
                        <tr key={f.id} className="hover:bg-slate-50/80 transition-colors">
                          <td className="py-2.5 px-3">
                            <input
                              type="text"
                              value={f.concepto}
                              onChange={(e) => actualizarFila(f.id, "concepto", e.target.value)}
                              className="w-full min-w-[170px] px-2.5 py-1.5 font-medium text-slate-800 border border-transparent hover:border-slate-300 focus:border-blue-500 focus:bg-white rounded-lg bg-transparent transition"
                            />
                          </td>

                          {rubro !== "general" && (
                            <td className="py-2.5 px-3">
                              <input
                                type="text"
                                value={f.detalleRubro}
                                onChange={(e) =>
                                  actualizarFila(f.id, "detalleRubro", e.target.value)
                                }
                                className="w-36 px-2.5 py-1.5 text-xs font-medium border border-blue-200 bg-blue-50/40 rounded-lg text-blue-900 focus:outline-none focus:ring-2 focus:ring-blue-500"
                              />
                            </td>
                          )}

                          <td className="py-2.5 px-3">
                            <div className="relative flex items-center">
                              <span className="text-xs text-slate-400 mr-1">{moneda}</span>
                              <input
                                type="number"
                                value={f.valorA}
                                onChange={(e) =>
                                  actualizarFila(f.id, "valorA", Number(e.target.value))
                                }
                                className="w-24 px-2.5 py-1.5 border border-slate-200 hover:border-slate-300 focus:border-blue-500 rounded-lg bg-white text-slate-800 font-medium"
                              />
                            </div>
                          </td>

                          {esPedidos && (
                            <td className="py-2.5 px-3">
                              <input
                                type="number"
                                min={1}
                                value={f.cantidad ?? 1}
                                onChange={(e) =>
                                  actualizarFila(
                                    f.id,
                                    "cantidad",
                                    Math.max(1, Number(e.target.value))
                                  )
                                }
                                className="w-20 px-2.5 py-1.5 border border-slate-200 hover:border-slate-300 focus:border-blue-500 rounded-lg bg-white text-slate-800 font-medium"
                              />
                            </td>
                          )}

                          <td className="py-2.5 px-3">
                            <input
                              type="number"
                              value={f.valorB}
                              onChange={(e) =>
                                actualizarFila(f.id, "valorB", Number(e.target.value))
                              }
                              className="w-24 px-2.5 py-1.5 border border-slate-200 hover:border-slate-300 focus:border-blue-500 rounded-lg bg-white text-slate-800 font-medium"
                            />
                          </td>

                          <td className="py-2.5 px-3">
                            <select
                              value={f.desplegable}
                              onChange={(e) =>
                                actualizarFila(f.id, "desplegable", e.target.value)
                              }
                              className="px-2.5 py-1.5 border border-slate-200 hover:border-slate-300 rounded-lg bg-slate-50 font-medium text-xs text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer"
                            >
                              {configActual.opcionesDropdown.map((op) => (
                                <option key={op} value={op}>
                                  {op}
                                </option>
                              ))}
                            </select>
                          </td>

                          <td className="py-2.5 px-4 text-right whitespace-nowrap">
                            <span
                              className={`inline-block px-2 py-0.5 rounded-md font-semibold text-xs ${
                                esSumaImpuesto
                                  ? "bg-slate-100 text-slate-700"
                                  : "bg-amber-50 text-amber-700"
                              }`}
                            >
                              {esSumaImpuesto ? "+" : "-"}
                              {moneda}
                              {montoImpuesto.toLocaleString("es-AR", {
                                minimumFractionDigits: 2,
                                maximumFractionDigits: 2,
                              })}
                            </span>
                          </td>

                          <td className="py-2.5 px-4 text-right whitespace-nowrap">
                            <span
                              className={`inline-block px-2.5 py-1 rounded-lg font-bold text-sm ${
                                resultado < 0
                                  ? "bg-red-50 text-red-700"
                                  : esPresupuesto
                                  ? "bg-slate-100 text-slate-800"
                                  : "bg-emerald-50/70 text-emerald-800"
                              }`}
                            >
                              {moneda}
                              {resultado.toLocaleString("es-AR", {
                                minimumFractionDigits: 2,
                                maximumFractionDigits: 2,
                              })}
                            </span>
                          </td>

                          {esPresupuesto && (
                            <td className="py-2.5 px-4 text-right whitespace-nowrap">
                              <span className="inline-block px-2.5 py-1 rounded-lg font-bold text-sm bg-emerald-50 text-emerald-800 border border-emerald-200">
                                {moneda}
                                {totalFinalPresupuesto.toLocaleString("es-AR", {
                                  minimumFractionDigits: 2,
                                  maximumFractionDigits: 2,
                                })}
                              </span>
                            </td>
                          )}

                          <td className="py-2.5 px-3 text-center print:hidden">
                            <div className="inline-flex items-center gap-1.5">
                              {plantilla === "stock" && (
                                <button
                                  type="button"
                                  onClick={() => registrarVentaItem(f, 1)}
                                  disabled={f.valorB <= 0}
                                  className={`px-2.5 py-1 rounded-lg text-xs font-bold transition cursor-pointer ${
                                    f.valorB <= 0
                                      ? "bg-slate-100 text-slate-400 cursor-not-allowed"
                                      : "bg-emerald-600 hover:bg-emerald-700 text-white shadow-xs"
                                  }`}
                                  title="Descontar 1 unidad vendida del stock"
                                >
                                  -1 Venta
                                </button>
                              )}
                              <button
                                onClick={() => eliminarFila(f.id)}
                                className="text-slate-400 hover:text-red-600 hover:bg-red-50 p-1.5 rounded-lg transition cursor-pointer"
                                title="Eliminar fila"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex flex-col sm:flex-row justify-between items-center gap-4">
                <button
                  onClick={agregarFila}
                  className="flex items-center gap-2 text-sm font-semibold text-blue-700 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 px-4 py-2 rounded-xl transition cursor-pointer print:hidden"
                >
                  <Plus className="w-4 h-4" />
                  Agregar nueva fila
                </button>

                <div className="flex items-center gap-3">
                  <span className="text-xs font-bold uppercase tracking-wider text-slate-500">
                    Total Consolidado:
                  </span>
                  <span
                    className={`text-xl font-extrabold ${
                      totalGeneral < 0 ? "text-red-600" : "text-slate-900"
                    }`}
                  >
                    {moneda}
                    {totalGeneral.toLocaleString("es-AR", {
                      minimumFractionDigits: 2,
                      maximumFractionDigits: 2,
                    })}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}