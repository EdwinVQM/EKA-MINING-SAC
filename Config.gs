/***** CONFIG CENTRALIZADA (ÚNICA FUENTE) *****/
const ADMIN_EMAIL = 'presupuestos@ekamining.com';

// Nombres “raw”
const SHEET_MASTER = 'Cotizaciones_Master';
const SHEET_USERS  = 'Usuarios';
const SHEET_COLS   = 'Columnas';
const SHEET_VIEW   = 'Vista_Usuario';

// Columnas canónicas del master
const MASTER_COLUMNS = [
  'id', 'fecha_registro',
  'cliente','unidad_trabajo', 'proyecto', 'referencia',
  'responsable_tecnico', 'responsable_economico',
  'estado', 'fecha_presentacion', 'fecha_limite',
  'monto_estimado', 'monto_ofertado', 'probabilidad',
  'moneda', 'descripcion',
  'correo_message_id', 'correo_thread_id', 'origen_invitacion',
  'industria', 'area_cliente', 'tipo_servicio', 'prioridad',
  'fecha_invitacion', 'fecha_confirmacion_participacion',
  'fecha_visita_tecnica', 'fecha_reporte_visita_tecnica',
  'fecha_limite_consultas', 'fecha_absolucion',
  'fecha_limite_presentacion', 'fecha_resultado',
  'tipo_presupuesto', 'razon_no_adjudicado',
  'registro_validado', 'validado_por', 'validado_en',
  'kpi_score', 'tags',
  // nuevas
  'tipo_gasto',
  'valor_contrato',
  'costo_presupuestado',
  'acumulado_coste_real',
  'acumulado_avance_valorizado',
  'eac', 'etc', 'cpi', 'spi',
  'mes_inicio_forecast', 'mes_fin_forecast'
];

// Columnas editables mínimas
const DEFAULT_EDITABLE = ['estado', 'fecha_presentacion'];

/***** NOMBRES BONITOS + COLORES *****/
const NICE_SHEET_MASTER = '📦 Cotizaciones_Master';
const NICE_SHEET_USERS  = '👥 Usuarios';
const NICE_SHEET_COLS   = '🧱 Columnas';
const NICE_SHEET_VIEW   = '👁️ Vista_Usuario';

const TAB_COLORS = {
  [NICE_SHEET_MASTER]:  {r:0.90,g:0.79,b:0.24},
  [NICE_SHEET_USERS]:   {r:0.31,g:0.64,b:0.91},
  [NICE_SHEET_COLS]:    {r:0.62,g:0.62,b:0.62},
  [NICE_SHEET_VIEW]:    {r:0.41,g:0.84,b:0.52}
};

const NAME_MAP = new Map([
  [SHEET_MASTER, NICE_SHEET_MASTER],
  [SHEET_USERS,  NICE_SHEET_USERS],
  [SHEET_COLS,   NICE_SHEET_COLS],
  [SHEET_VIEW,   NICE_SHEET_VIEW],
]);

/***** SUPABASE / NOCO Y DRIVE ******/
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const PARENT_FOLDER_ID = SCRIPT_PROPS.getProperty('PARENT_FOLDER_ID') || '1tbpO3aJieCfbcZib5AdgrilgXoJewlbM';

const SB_URL   = SCRIPT_PROPS.getProperty('SB_URL')   || '';
const SB_KEY   = SCRIPT_PROPS.getProperty('SB_KEY')   || '';
const SB_TABLE = SCRIPT_PROPS.getProperty('SB_TABLE') || 'cotizaciones';

// Si aún no definiste estos, déjalos vacíos (no rompe)
const NOCO_URL     = SCRIPT_PROPS.getProperty('NOCO_URL')     || '';
const NOCO_PROJECT = SCRIPT_PROPS.getProperty('NOCO_PROJECT') || '';
const NOCO_TABLE   = SCRIPT_PROPS.getProperty('NOCO_TABLE')   || 'cotizaciones';

// Subcarpetas de planificación
const PLAN_SUBFOLDERS = [
  '00_Entrada de Licitación (invitación, TDR, anexos)',
  '05_Alcance (PGA, Requisitos, EDT/WBS, diccionario)',
  '06_Cronograma (MSP nivel 7, Curva S, línea base)',
  '07_Costos (estimaciones, presupuesto, supuestos)',
  '08_Calidad (Plan de Calidad, formatos de control)',
  '09_Recursos (RRHH, RACI, perfiles críticos)',
  '10_Comunicaciones (matriz, formatos de reporte)',
  '11_Riesgos (registro, análisis cuali/cuant, respuesta)',
  '12_Adquisiciones (Plan)',
  '12_Adquisiciones/Pre-OC (cotizaciones de materiales)',
  '12_Adquisiciones/Subcontratos (pre-adjudicación)',
  '12_Adquisiciones/Requerimientos (prediseños)',
  '13_Interesados (registro, estrategia)',
  'Control Documentario (índices, dossier preliminar)'
];
