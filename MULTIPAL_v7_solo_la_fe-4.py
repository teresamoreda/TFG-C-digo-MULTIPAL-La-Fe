"""
=============================================================================
  MULTIPAL v7 — Análisis Multicriterio · Cuidados Paliativos Domiciliarios
  VERSIÓN FILTRADA AL HOSPITAL LA FE
=============================================================================
  Esta versión procesa ÚNICAMENTE los registros del Hospital La Fe.

  Tratamiento de las unidades de La Fe (revisión metodológica):
       · Las subunidades de Hospitalización a Domicilio (Adultos,
         Pediatría, Salud Mental, Salud Infanto-Juvenil) se ANALIZAN EN
         CONJUNTO, como una única unidad de HaD, SIN desagregar, porque
         por separado tienen muy pocas respuestas.
       · Oncología y Hematología de La Fe se analizan POR SEPARADO, con
         su propio informe detallado, ya que son entornos muy distintos
         a la UHD y NO se comparan con las medias de las UHD.
  El resto de hospitales (Manises, Arnau, etc.) se ignoran en esta versión.
  En esta versión NO se contrasta con medias generales (una sola unidad).
  INSTRUCCIONES (Mac / VS Code)
  ─────────────────────────────
  1. Instala dependencias (solo una vez):
       pip install pandas numpy matplotlib seaborn scipy scikit-posthocs statsmodels openpyxl

  2. Coloca el Excel exportado de Serviencuestas en:
       /Users/martamunoz/DocumentosPython/AnalisisMulticriterio/dataBase.xlsx

  3. Ejecuta:
       python MULTIPAL_v2.py
     o con otro archivo:
       python MULTIPAL_v2.py --file /ruta/nuevo_export.xlsx

  4. Resultados (PNG + CSV + TXT) en:
       /Users/martamunoz/DocumentosPython/AnalisisMulticriterio/outputs/

  NOVEDAD v2: el análisis es 100 % dinámico.
  • Detecta automáticamente todos los hospitales/unidades presentes en la BD.
  • Si hay un hospital nuevo, genera su propio bloque de análisis y MCDA.
  • Basta reemplazar dataBase.xlsx y volver a ejecutar.
=============================================================================
"""

# ── IMPORTACIONES ─────────────────────────────────────────────────────────────
# Librerías estándar de Python
import os, sys, re, warnings, textwrap, argparse
from pathlib import Path
from datetime import datetime

# Librerías numéricas y de análisis
import numpy as np                         # cálculos numéricos y álgebra lineal
import pandas as pd                        # manejo de datos tabulares (DataFrames)

# Matplotlib: motor de gráficos principal
import matplotlib
matplotlib.use("Agg")                      # backend sin pantalla (genera PNG sin abrir ventana)
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches     # parches para leyendas personalizadas
import matplotlib.gridspec as gridspec    # layouts de subplots complejos
from matplotlib.backends.backend_pdf import PdfPages  # generación de informe PDF multipágina
from matplotlib.image import imread        # incrustar PNGs ya generados en el PDF

# Seaborn: gráficos estadísticos (heatmaps, etc.)
import seaborn as sns

# Estadística científica
from scipy import stats
from scipy.stats import kruskal, mannwhitneyu, spearmanr  # tests no paramétricos y correlación
import scikit_posthocs as sp              # post-hoc de Dunn tras Kruskal-Wallis

warnings.filterwarnings("ignore")        # suprime advertencias menores durante la ejecución


# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN GLOBAL: RUTAS, ETIQUETAS Y PALETAS DE COLOR
# ══════════════════════════════════════════════════════════════════════════════

# ── Rutas por defecto ──────────────────────────────────────────────────────────
# Se puede sobreescribir desde línea de comandos con --file y --outdir
BASE_DIR   = Path("/Users/martamunoz/DocumentosPython/AnalisisMulticriterio")
INPUT_FILE = BASE_DIR / "dataBase.csv"   # CSV (o Excel) exportado de la encuesta
OUT_DIR    = BASE_DIR / "outputsLaFe"     # carpeta específica para esta versión filtrada

# ── Comparación con medias generales ─────────────────────────────────────────
# En el análisis COMPLETO (todos los hospitales) cada UHD se compara con la
# media del conjunto de UHD comparables. En la versión "solo La Fe" no procede
# comparar con medias (se trabaja una única unidad), por lo que este switch se
# pone a False en ese archivo.
COMPARAR_CON_MEDIAS = False

# ── Etiquetas de las 5 dimensiones del cuestionario ──────────────────────────
# P1–P5 son las 5 dimensiones clave de los Cuidados Paliativos evaluadas
DIM   = ["P1","P2","P3","P4","P5"]

# Etiqueta corta (para ejes de gráficas)
DLBL  = {"P1":"P1: Control síntomas","P2":"P2: Comunicación",
         "P3":"P3: Psicosocial","P4":"P4: Coordinación","P5":"P5: Autocuidado equipo"}

# Versión aún más corta (solo el nombre, sin el código) para informes
# en lenguaje sencillo dirigidos a los equipos
DSHORT = {"P1": "Control síntomas",
          "P2": "Comunicación",
          "P3": "Psicosocial",
          "P4": "Coordinación",
          "P5": "Autocuidado equipo"}

# Etiqueta larga (para informes de texto)
DFULL = {"P1":"Control efectivo de síntomas",
         "P2":"Comunicación y decisiones compartidas",
         "P3":"Atención psicosocial y espiritual",
         "P4":"Coordinación y continuidad asistencial",
         "P5":"Cuidado y apoyo al equipo (burnout)"}

# Color asignado a cada dimensión (azul, verde, naranja, morado, rojo)
DCOL  = ["#003D7C","#2E9E6B","#E8A020","#9B59B6","#E74C3C"]

# ── Paletas de color fijas ────────────────────────────────────────────────────
# Se amplían dinámicamente si hay más grupos de los previstos
_BASE_COLORS = ["#003D7C","#E8A020","#2E9E6B","#9B59B6","#E74C3C",
                "#F39C12","#1ABC9C","#8E44AD","#D35400","#2C3E50"]

# Paleta por rol profesional
PAL_ROL = {"Médico":"#003D7C","Enfermera":"#2E9E6B","Fisioterapeuta":"#E8A020",
           "Psicologo":"#9B59B6","TCAE":"#E74C3C","Trabajador Social":"#7F8C8D"}

# Paleta por género
PAL_GEN = {"Femenino":"#E91E8C","Masculino":"#003D7C"}

# Paleta por grupo de edad
PAL_AGE = {"≤35":"#003D7C","36-45":"#2E9E6B","46-55":"#E8A020",">55":"#E74C3C"}

# Paleta por años de experiencia
PAL_EXP = {"<1 año":"#BDC3C7","1-3 años":"#5D6D7E","3-5 años":"#1F618D",">5 años":"#003D7C"}


def make_palette(keys):
    """
    Genera una paleta de colores dinámica para cualquier lista de grupos.
    Recorre _BASE_COLORS en ciclo si hay más grupos que colores predefinidos.
    """
    return {str(k): _BASE_COLORS[i % len(_BASE_COLORS)] for i,k in enumerate(keys)}


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 1 — UTILIDADES GENERALES
# ══════════════════════════════════════════════════════════════════════════════

def save(fig, name, outdir):
    """Guarda una figura matplotlib como PNG en la carpeta de salida y la cierra."""
    p = outdir / f"{name}.png"
    fig.savefig(p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"   ✓ {p.name}")

def sig(p):
    """
    Convierte un p-valor en su símbolo de significación estadística:
    *** p<0.001 | ** p<0.01 | * p<0.05 | ns = no significativo
    """
    if pd.isna(p): return "—"
    return "***" if p<.001 else "**" if p<.01 else "*" if p<.05 else "ns"

def radar_angles():
    """
    Calcula los ángulos para los 5 vértices de un gráfico radar (araña).
    Añade el primer ángulo al final para cerrar el polígono.
    """
    a = np.linspace(0, 2*np.pi, 5, endpoint=False).tolist()
    return a + a[:1]

def kendall_w(mat):
    """
    Calcula el coeficiente W de Kendall a partir de una matriz de rangos.
    Mide el grado de acuerdo entre evaluadores (0=sin acuerdo, 1=acuerdo total).
    mat: matriz donde filas=evaluadores, columnas=criterios
    """
    n, m = mat.shape
    R    = mat.sum(axis=0)          # suma de rangos por criterio
    S    = ((R - R.mean())**2).sum()  # varianza de las sumas de rangos
    return float(np.clip(12*S / (n**2 * (m**3 - m)), 0, 1))


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 2 — CARGA Y LIMPIEZA DE DATOS (100% DINÁMICA)
# ══════════════════════════════════════════════════════════════════════════════

def _corregir_valor_escala(v):
    """
    Corrige un valor de las preguntas de pesos / importancia / preparación,
    que SOLO admiten valores de 0 a 100.

    Reglas de limpieza acordadas con el equipo:
      · Si el valor ya está en [0, 100] → se conserva.
      · Si es un número "duplicado" por error de tecleo (la misma cifra
        repetida dos veces: 8080 → 80, 9090 → 90, 5050 → 50, 100100 → 100)
        → se recupera la mitad.
      · Cualquier otro valor fuera de rango se recorta al límite válido
        (p. ej. 199 → 100, 1001 → 100, valores negativos → 0).
      · Los valores ausentes (NaN) se devuelven tal cual.
    """
    if pd.isna(v):
        return v
    try:
        f = float(v)
    except (TypeError, ValueError):
        return np.nan
    if 0 <= f <= 100:
        return f
    # Patrón de cifra duplicada (p. ej. "8080" = "80"+"80")
    if f > 0 and float(f).is_integer():
        s = str(int(f))
        if len(s) % 2 == 0:
            mitad = len(s) // 2
            if s[:mitad] == s[mitad:]:
                h = int(s[:mitad])
                if 0 <= h <= 100:
                    return float(h)
    # Recorte al rango válido
    if f < 0:
        return 0.0
    return 100.0


def load(filepath):
    """
    Lee la base de datos exportada de la encuesta (CSV o Excel) y devuelve
    un DataFrame limpio.

    Columnas renombradas:
      · id, pagina, edad, rol, genero, experiencia, formacion, provincia, unidad
      · W_P1–W_P5  : pesos de prioridad asignados a cada dimensión (%)
      · I_P1–I_P5  : valoración de importancia percibida (0–100)
      · R_P1–R_P5  : valoración de percepción de preparación (0–100)

    Transformaciones aplicadas:
      · Filtrado: solo se conservan respuestas de la página 2 del cuestionario
      · Coerción numérica y eliminación de filas con pesos missing
      · Clipping de I y R al rango [0, 100]
      · Inferencia dinámica del hospital a partir del campo "unidad"
      · Variables derivadas: grupo_edad, exp_label, exp_num, G_Pi (brecha), rol_bin, gen_bin
    """
    # Lectura robusta: admite tanto el CSV exportado de LimeSurvey como
    # el Excel clásico de Serviencuestas. Se detecta por la extensión.
    if str(filepath).lower().endswith(".csv"):
        raw = pd.read_csv(filepath, encoding="utf-8-sig")
    else:
        raw = pd.read_excel(filepath)

    # Renombrado posicional de columnas según el diseño del Excel de Serviencuestas
    rename = {
        raw.columns[0]:"id",    raw.columns[2]:"pagina",
        raw.columns[8]:"edad",  raw.columns[9]:"rol",
        raw.columns[10]:"genero", raw.columns[11]:"experiencia",
        raw.columns[12]:"formacion", raw.columns[13]:"provincia",
        raw.columns[14]:"unidad",
        **{raw.columns[15+i]:f"W_P{i+1}" for i in range(5)},   # columnas 15–19
        **{raw.columns[20+i]:f"I_P{i+1}" for i in range(5)},   # columnas 20–24
        **{raw.columns[25+i]:f"R_P{i+1}" for i in range(5)},   # columnas 25–29
    }
    df = raw.rename(columns=rename)

    # ── LIMPIEZA DE LA BASE DE DATOS (primer paso) ───────────────────────────
    # Solo se analizan respuestas que llegaron al módulo de valoración
    df = df[df["pagina"] == 2].copy()

    # Columnas de las tres preguntas centrales (pesos, importancia, preparación)
    cols_W = [f"W_P{i}" for i in range(1, 6)]
    cols_I = [f"I_P{i}" for i in range(1, 6)]
    cols_R = [f"R_P{i}" for i in range(1, 6)]
    cols_core = cols_W + cols_I + cols_R

    # 1) Coerción a numérico de todas las variables cuantitativas
    for c in cols_core + ["edad"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # 2) Corrección de valores fuera del rango válido [0, 100] en las tres
    #    preguntas (pesos, importancia y preparación). Se reparan errores de
    #    tecleo del tipo 8080 → 80 y se recortan valores como 199 → 100.
    n_corregidos = 0
    for c in cols_core:
        original = df[c].copy()
        df[c] = df[c].apply(_corregir_valor_escala)
        n_corregidos += int(((original != df[c]) & original.notna()).sum())

    # 3) Eliminación de filas con las preguntas de pesos, importancia y
    #    preparación incompletas. Para que una respuesta sea analizable
    #    (cálculo de brechas Imp−Prep y MCDA con pesos que sumen 100) se
    #    exige que las TRES preguntas estén completas en sus 5 dimensiones.
    n_antes = len(df)
    df = df.dropna(subset=cols_core)
    n_eliminadas = n_antes - len(df)

    print(f"\n  Limpieza de la base de datos:")
    print(f"     · {n_corregidos} celdas con valores fuera de rango corregidas "
          f"(p. ej. 8080→80, 199→100)")
    print(f"     · {n_eliminadas} filas eliminadas por respuestas incompletas "
          f"en pesos / importancia / preparación")
    print(f"     · {len(df)} respuestas válidas conservadas")

    # Limpieza de variables categóricas
    for c in ["rol","genero","unidad","experiencia","formacion","provincia"]:
        df[c] = df[c].astype(str).str.strip().replace("nan", np.nan)

    # ── DETECCIÓN AUTOMÁTICA DE HOSPITALES / GRUPOS DE ANÁLISIS ──────────────
    # Cada unidad recogida en el Excel se clasifica en:
    #   · "grupo": nombre legible del grupo de análisis (lo que antes era
    #              "hospital"). Define al colectivo cuyo informe individual
    #              se generará. Una unidad ≡ un grupo de análisis.
    #   · "tipo":  categoría clínica de la unidad. Solo los grupos de tipo
    #              "cp_adultos" entran en el comparativo entre hospitales,
    #              porque las unidades de Pediatría, Salud Mental o
    #              Salud Infanto-Juvenil atienden a poblaciones distintas
    #              y no son clínicamente comparables.
    #   · "comparable": True si el grupo entra en el comparativo entre
    #              hospitales (UHD-Paliativos Adultos).
    #
    # El mapeo es explícito para las unidades conocidas y se complementa
    # con una heurística por palabras clave que permite incorporar
    # automáticamente nuevos centros (Arnau, Xàtiva, etc.) sin tocar el
    # código.
    unidades_unicas = df["unidad"].dropna().unique().tolist()

    MAPEO_UNIDADES = {
        # Cadena en BD                  →  (nombre legible,                  tipo,                   comparable)
        "La Fe":                           ("La Fe — Adultos",               "cp_adultos",           True),
        "Manises":                         ("Manises",                       "cp_adultos",           True),
        "Arnau":                           ("Arnau",                         "cp_adultos",           True),
        "Xàtiva":                          ("Xàtiva",                        "cp_adultos",           True),
        "Xativa":                          ("Xàtiva",                        "cp_adultos",           True),
        "Pediatria La Fe":                 ("La Fe — Pediatría",             "cp_pediatria",         False),
        "Pediatría La Fe":                 ("La Fe — Pediatría",             "cp_pediatria",         False),
        "Salud Mental La Fe":              ("La Fe — Salud Mental",          "salud_mental",         False),
        "Salud Infantojuvenil La Fe":      ("La Fe — Salud Infanto-Juvenil", "salud_infantojuvenil", False),
        "Salud Infanto-Juvenil La Fe":     ("La Fe — Salud Infanto-Juvenil", "salud_infantojuvenil", False),
        # Unidades hospitalarias de La Fe distintas de la UHD. Reciben informe
        # detallado propio, pero NO entran en el comparativo entre UHD ni se
        # contrastan con las medias generales (son entornos muy distintos).
        "Oncologia La Fe":                 ("La Fe — Oncología",             "oncologia",            False),
        "Oncología La Fe":                 ("La Fe — Oncología",             "oncologia",            False),
        "Hematologia La Fe":               ("La Fe — Hematología",           "hematologia",          False),
        "Hematología La Fe":               ("La Fe — Hematología",           "hematologia",          False),
    }

    def clasificar_unidad(u):
        """
        Devuelve (grupo, tipo, comparable) para una cadena de unidad.

        1) Primero busca coincidencia exacta en MAPEO_UNIDADES.
        2) Si no la hay, aplica heurística por palabras clave para
           detectar automáticamente la categoría clínica.
        3) Si no encaja en ninguna categoría especializada, se asume
           que es una UHD-Paliativos Adultos (cp_adultos, comparable).
        """
        if pd.isna(u):
            return (np.nan, np.nan, False)
        u = str(u).strip()
        # Mapeo explícito
        if u in MAPEO_UNIDADES:
            return MAPEO_UNIDADES[u]
        # Heurística para unidades nuevas
        u_low = u.lower()
        hosps_conocidos = ["La Fe", "Manises", "Arnau", "Xàtiva", "Xativa"]
        if "pediatr" in u_low:
            for h in hosps_conocidos:
                if h.lower() in u_low:
                    return (f"{h} — Pediatría", "cp_pediatria", False)
            return (f"{u} — Pediatría", "cp_pediatria", False)
        if "mental" in u_low:
            for h in hosps_conocidos:
                if h.lower() in u_low:
                    return (f"{h} — Salud Mental", "salud_mental", False)
            return (f"{u} — Salud Mental", "salud_mental", False)
        if "infanto" in u_low or "infantojuv" in u_low:
            for h in hosps_conocidos:
                if h.lower() in u_low:
                    return (f"{h} — Salud Infanto-Juvenil",
                            "salud_infantojuvenil", False)
            return (f"{u} — Salud Infanto-Juvenil",
                    "salud_infantojuvenil", False)
        if "oncolog" in u_low:
            for h in hosps_conocidos:
                if h.lower() in u_low:
                    return (f"{h} — Oncología", "oncologia", False)
            return (f"{u} — Oncología", "oncologia", False)
        if "hematolog" in u_low:
            for h in hosps_conocidos:
                if h.lower() in u_low:
                    return (f"{h} — Hematología", "hematologia", False)
            return (f"{u} — Hematología", "hematologia", False)
        # Por defecto: UHD-Paliativos Adultos, comparable
        return (u, "cp_adultos", True)

    # Aplicar la clasificación
    clas = df["unidad"].apply(clasificar_unidad)
    df["hospital"]    = clas.apply(lambda x: x[0])  # mantenemos el nombre por compatibilidad
    df["tipo_unidad"] = clas.apply(lambda x: x[1])
    df["comparable"]  = clas.apply(lambda x: x[2])

    # Guardar la cadena original por si se necesita auditoría
    df["unidad_raw"] = df["unidad"]
    df["unidad"]     = df["hospital"]   # cada grupo es ahora una unidad asistencial

    # ── HOSPITAL PADRE ───────────────────────────────────────────────────────
    # Para los informes generales que muestran resultados a nivel de
    # hospital (agregando todas las subunidades de un mismo centro),
    # extraemos el nombre raíz del hospital. Por ejemplo,
    #   "La Fe — Adultos"           → "La Fe"
    #   "La Fe — Pediatría"          → "La Fe"
    #   "La Fe — Salud Mental"       → "La Fe"
    #   "Manises"                    → "Manises"
    # De este modo, las cuatro subunidades de La Fe pueden tratarse en
    # el informe general como un único hospital agregado, manteniendo a
    # su vez el análisis desagregado por subunidad.
    def _hospital_padre(grupo):
        if pd.isna(grupo):
            return np.nan
        return grupo.split(" — ")[0] if " — " in grupo else grupo
    df["hospital_padre"] = df["hospital"].apply(_hospital_padre)

    # ── FILTRO: SOLO HOSPITAL LA FE ──────────────────────────────────────────
    n_antes_filtro = len(df)
    df = df[df["hospital_padre"] == "La Fe"].copy()
    n_descartados = n_antes_filtro - len(df)
    if n_descartados > 0:
        print(f"\n  FILTRO LA FE: {len(df)} registros retenidos "
              f"(de {n_antes_filtro} totales · {n_descartados} descartados "
              f"por no ser de La Fe)")
    if len(df) == 0:
        raise SystemExit("\n❌  No se han encontrado registros del "
                         "Hospital La Fe en la base de datos. "
                         "Verifica el campo 'unidad'.")

    # ── COMBINAR SUBUNIDADES EN UNA ÚNICA UNIDAD DE HaD ─────────────────────────
    # Por revisión metodológica: Oncología y Hematología de La Fe se mantienen
    # SEPARADAS (entornos muy distintos a la UHD). El resto de subunidades de
    # La Fe (Adultos, Pediatría, Salud Mental, Salud Infanto-Juvenil) se agregan
    # en UNA Única unidad de Hospitalización a Domicilio, sin desagregar, porque
    # por separado tienen muy pocas respuestas.
    TIPOS_SEPARADOS = {"oncologia", "hematologia"}
    NOMBRE_HAD = "Hospitalización a Domicilio La Fe"
    mask_had = ~df["tipo_unidad"].isin(TIPOS_SEPARADOS)
    df.loc[mask_had, "hospital"]    = NOMBRE_HAD
    df.loc[mask_had, "unidad"]      = NOMBRE_HAD
    df.loc[mask_had, "tipo_unidad"] = "cp_adultos"
    df.loc[mask_had, "comparable"]  = True   # HaD La Fe (única UHD aquí)
    df.loc[~mask_had, "comparable"] = False  # Oncología / Hematología aparte
    df["hospital_padre"] = "La Fe"
    print(f"  Subunidades de HaD agregadas en «La Fe»: "
          f"n={int(mask_had.sum())}")
    sep = sorted(df.loc[~mask_had, "hospital"].dropna().unique())
    if sep:
        print(f"  Unidades analizadas por separado: {', '.join(sep)}")
    df = df.reset_index(drop=True)


    # Excluir registros sin unidad asignada (no se pueden clasificar)
    n_sin_unidad = df["hospital"].isna().sum()
    if n_sin_unidad > 0:
        print(f"  ⚠ {n_sin_unidad} registros sin unidad asignada — se "
              f"excluyen de los análisis por grupo.")

    # Paletas de color generadas dinámicamente para los grupos detectados
    hospitales_unicos = sorted(df["hospital"].dropna().unique().tolist())
    unidades_unicas   = sorted(df["unidad"].dropna().unique().tolist())
    df._pal_hosp  = make_palette(hospitales_unicos)
    df._pal_unit  = make_palette(unidades_unicas)

    # Reporte por consola del mapeo realizado
    print(f"\n  Clasificación de unidades detectadas:")
    for g in hospitales_unicos:
        sub = df[df["hospital"] == g]
        tipo = sub["tipo_unidad"].iloc[0] if len(sub) else "?"
        comp = " (comparable)" if sub["comparable"].iloc[0] else " (especializada)"
        print(f"     · {g:<40} n={len(sub):<3}  tipo={tipo}{comp}")

    # ── VARIABLES DERIVADAS ───────────────────────────────────────────────────
    # Grupos de edad en 4 intervalos
    df["grupo_edad"] = pd.cut(df["edad"], bins=[0,35,45,55,120],
                               labels=["≤35","36-45","46-55",">55"], right=True)

    # Experiencia: texto libre → etiqueta ordenada
    exp_ord = ["Menos de un año","Entre uno y tres años","Entre 3 y 5 años","Mas de 5 años"]
    exp_lbl = ["<1 año","1-3 años","3-5 años",">5 años"]
    df["exp_label"] = pd.Categorical(
        df["experiencia"].map(dict(zip(exp_ord, exp_lbl))),
        categories=exp_lbl, ordered=True)
    df["exp_num"] = df["experiencia"].map(
        {"Menos de un año":1,"Entre uno y tres años":2,"Entre 3 y 5 años":3,"Mas de 5 años":4})

    # Brecha = Importancia − Percepción de preparación (valores positivos = déficit percibido)
    for i in range(1,6):
        df[f"G_P{i}"] = df[f"I_P{i}"] - df[f"R_P{i}"]

    # Variables binarias para correlaciones
    df["rol_bin"] = df["rol"].map({"Médico":1,"Enfermera":0})
    df["gen_bin"] = df["genero"].map({"Masculino":1,"Femenino":0})

    _report(df)
    return df


def _report(df):
    """Imprime en consola un resumen de la muestra cargada: n total, por hospital, unidad, género y edad."""
    sep = "="*62
    print(f"\n{sep}\n  MULTIPAL v2.0 · {datetime.now():%d/%m/%Y %H:%M}\n{sep}")
    print(f"  Respuestas completas : {len(df)}")
    print(f"\n  Por hospital:")
    for h,n in df["hospital"].value_counts().items():
        print(f"    {h}: {n}")
    print(f"\n  Por unidad:")
    for u,n in df["unidad"].value_counts().items():
        warn = " ⚠" if n<5 else ""
        print(f"    {u}: {n}{warn}")
    print(f"\n  Género  : {dict(df['genero'].value_counts(dropna=True))}")
    print(f"  G. edad : {dict(df['grupo_edad'].value_counts().sort_index())}")
    print(sep+"\n")


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 3 — FUNCIONES GRÁFICAS BASE (reutilizadas en todos los análisis)
# ══════════════════════════════════════════════════════════════════════════════

def _radar(df, group_col, titulo, palette, outdir, fname, min_n=1):
    """
    Genera un panel de gráficos radar (araña) con un subplot por grupo.
    Cada radar superpone la curva de Importancia (azul) y Percepción de preparación (naranja).
    Grupos con n<5 se marcan como "orientativos" en el título del subplot.
    min_n: tamaño mínimo de grupo para incluirlo en el análisis.
    """
    groups = [g for g in df[group_col].dropna().unique()
              if (df[group_col]==g).sum() >= min_n]
    n = len(groups)
    if n == 0: return
    ncols = min(n, 3); nrows = (n+ncols-1)//ncols
    fig, axs = plt.subplots(nrows, ncols, figsize=(5.5*ncols, 5*nrows),
                             subplot_kw=dict(polar=True))
    fig.suptitle(titulo, fontsize=13, fontweight="bold", y=1.01)
    axs_flat = np.array(axs).flatten() if n>1 else [axs]
    angles = radar_angles()
    for ax, g in zip(axs_flat, groups):
        sub  = df[df[group_col]==g]
        imp  = [sub[f"I_P{i}"].mean() for i in range(1,6)] + [sub["I_P1"].mean()]
        rend = [sub[f"R_P{i}"].mean() for i in range(1,6)] + [sub["R_P1"].mean()]
        ax.plot(angles, imp,  color="#003D7C", lw=2.2, label="Importancia")
        ax.fill(angles, imp,  color="#003D7C", alpha=.12)
        ax.plot(angles, rend, color="#E8A020", lw=2.2, ls="--", label="Percepción de preparación")
        ax.fill(angles, rend, color="#E8A020", alpha=.10)
        ax.set_xticks(angles[:-1]); ax.set_xticklabels(DIM, fontsize=9)
        ax.set_ylim(0, 100); ax.set_yticks([20,40,60,80,100])
        ax.set_yticklabels(["20","40","60","80","100"], fontsize=7)
        col = palette.get(str(g), "#555"); ng = len(sub)
        ax.set_title(f"{g}\n(n={ng}{'  ⚠ orientativo' if ng<5 else ''})",
                     fontsize=9, fontweight="bold", color=col, pad=12)
        ax.legend(loc="lower right", fontsize=7)
    for ax in axs_flat[n:]: ax.set_visible(False)
    plt.tight_layout(); save(fig, fname, outdir)


def _radar_brecha_rol(df, titulo, outdir, fname, min_n=3, return_path=False):
    """
    Gráfico de araña de la BRECHA (Importancia − Percepción de preparación)
    por dimensión, con un polígono por ROL profesional (médicos vs
    enfermería, y otros roles si tienen muestra suficiente).

    A diferencia del radar clásico (que superpone importancia y
    preparación), aquí cada vértice representa la BRECHA media de una
    dimensión: valores altos = mayor déficit percibido (la importancia
    supera a la preparación), valores en torno a 0 = sin déficit.

    Devuelve la ruta del PNG si return_path=True (para incrustarlo en
    los informes PDF por unidad); en caso contrario None.
    """
    # Roles con muestra suficiente, dando prioridad a Médico y Enfermera
    orden_roles = ["Médico", "Enfermera", "Médico Residente (MIR)", "TCAE",
                   "Psicologo", "Fisioterapeuta", "Trabajador Social"]
    presentes = [r for r in df["rol"].dropna().unique()]
    roles = [r for r in orden_roles if r in presentes
             and (df["rol"] == r).sum() >= min_n]
    roles += [r for r in presentes if r not in orden_roles
              and (df["rol"] == r).sum() >= min_n]
    if len(roles) == 0:
        return None

    angles = radar_angles()
    fig = plt.figure(figsize=(8.5, 7.5))
    ax = fig.add_subplot(111, polar=True)
    fig.suptitle(titulo, fontsize=13, fontweight="bold", y=1.0)

    todos = []
    for r in roles:
        sub = df[df["rol"] == r]
        brecha = [sub[f"G_P{i}"].mean() for i in range(1, 6)]
        todos += [b for b in brecha if not pd.isna(b)]
        brecha_cerrado = brecha + brecha[:1]
        col = PAL_ROL.get(r, "#555")
        ax.plot(angles, brecha_cerrado, color=col, lw=2.4,
                label=f"{r} (n={len(sub)})")
        ax.fill(angles, brecha_cerrado, color=col, alpha=0.12)

    # Límites radiales adaptados (la brecha puede ser negativa)
    lo = min(todos + [0]); hi = max(todos + [0])
    pad = max(2.0, (hi - lo) * 0.15)
    ax.set_ylim(lo - pad, hi + pad)
    # Círculo de referencia en brecha = 0 (sin déficit)
    ax.plot(np.linspace(0, 2*np.pi, 200), [0]*200,
            color="#333", lw=1.1, ls="--", alpha=0.7)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels([f"{p}\n{DSHORT[p]}" for p in DIM], fontsize=9)
    ax.tick_params(axis="y", labelsize=7)
    ax.set_title("Brecha media Importancia − Percepción de preparación "
                 "(↑ = mayor déficit)", fontsize=9, color="#555", pad=18)
    ax.legend(loc="upper right", bbox_to_anchor=(1.28, 1.10), fontsize=9)
    plt.tight_layout()

    if return_path:
        ruta = outdir / f"{fname}.png"
        fig.savefig(ruta, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return ruta
    save(fig, fname, outdir)
    return outdir / f"{fname}.png"


def _barras(df, group_col, prefix, titulo, ylabel, palette, outdir, fname, ylim=115):
    """
    Gráfico de barras agrupadas comparando las 5 dimensiones entre grupos.
    Muestra barras de error (SEM) encima de cada barra.
    prefix: 'R' (percepción de preparación), 'I' (importancia) o 'W' (pesos)
    """
    groups = sorted(df[group_col].dropna().unique())
    x = np.arange(5); w = 0.8/len(groups)
    fig, ax = plt.subplots(figsize=(14, 6))
    ax.set_title(titulo, fontsize=12, fontweight="bold")
    for j, g in enumerate(groups):
        sub = df[df[group_col]==g]
        m   = [sub[f"{prefix}_P{i}"].mean() for i in range(1,6)]
        e   = [sub[f"{prefix}_P{i}"].sem()  for i in range(1,6)]
        off = (j - len(groups)/2 + .5) * w
        col = palette.get(str(g), "#999")
        lbl = f"{g} (n={len(sub)})" + (" ⚠" if len(sub)<5 else "")
        ax.bar(x+off, m, w, label=lbl, color=col, alpha=.87, yerr=e, capsize=3)
    ax.set_xticks(x)
    ax.set_xticklabels([DLBL[d] for d in DIM], rotation=14, ha="right", fontsize=9)
    ax.set_ylabel(ylabel); ax.set_ylim(0, ylim)
    ax.legend(fontsize=9, loc="upper right"); ax.grid(axis="y", alpha=.3)
    plt.tight_layout(); save(fig, fname, outdir)


def _heatmap_gap(df, group_col, titulo, outdir, fname):
    """
    Heatmap de la brecha (Importancia − Percepción de preparación) por grupo y dimensión.
    Colores: rojo = brecha alta (déficit), verde = brecha pequeña o negativa (exceso).
    Útil para identificar de un vistazo qué dimensiones tienen mayor déficit.
    """
    groups = df[group_col].dropna().unique()
    mat = pd.DataFrame(
        [[df[df[group_col]==g][f"G_P{i}"].mean() for i in range(1,6)] for g in groups],
        index=[f"{g} (n={int((df[group_col]==g).sum())})" for g in groups],
        columns=[DLBL[d] for d in DIM])
    fig, ax = plt.subplots(figsize=(11, max(3, len(groups)*1.1)))
    sns.heatmap(mat, annot=True, fmt=".1f", cmap="RdYlGn_r", center=0,
                vmin=-15, vmax=40, linewidths=.6, ax=ax,
                cbar_kws={"label":"Brecha · Importancia − Percepción de preparación"})
    ax.set_title(titulo, fontsize=12, fontweight="bold")
    plt.tight_layout(); save(fig, fname, outdir)


def _mw(df, group_col, g1, g2, prefix):
    """
    Test de Mann-Whitney U para comparar dos grupos en las 5 dimensiones.
    Devuelve un DataFrame con U, p-valor, tamaño del efecto r y significación.
    r = (2U)/(n1·n2) − 1  →  rango [−1, 1], similar a correlación de Pearson.
    Requiere n≥3 en cada grupo para ejecutarse.
    """
    rows = []
    for i in range(1,6):
        a = df[df[group_col]==g1][f"{prefix}_P{i}"].dropna().values
        b = df[df[group_col]==g2][f"{prefix}_P{i}"].dropna().values
        if len(a)<3 or len(b)<3:
            rows.append({"Dim":f"P{i}","U":np.nan,"p":np.nan,"r":np.nan,"Sig.":"—"}); continue
        U, p = mannwhitneyu(a, b, alternative="two-sided")
        r    = (2*U)/(len(a)*len(b)) - 1
        rows.append({"Dim":f"P{i}","U":round(U,1),"p":round(p,4),"r":round(r,3),"Sig.":sig(p)})
    return pd.DataFrame(rows)


def _kruskal_dunn(df, group_col, prefix, min_n=3):
    """
    Test de Kruskal-Wallis (alternativa no paramétrica a ANOVA) para k grupos.
    Si el resultado es significativo (p<0.05), aplica post-hoc de Dunn
    con corrección de Bonferroni para identificar qué pares de grupos difieren.
    Devuelve: (DataFrame con resultados KW, dict de tablas Dunn por dimensión).
    """
    groups = [g for g in df[group_col].dropna().unique()
              if (df[group_col]==g).sum() >= min_n]
    kw_rows = []; dunn_all = {}
    for i in range(1,6):
        col = f"{prefix}_P{i}"
        samples = {str(g): df[df[group_col]==g][col].dropna().values for g in groups}
        samples  = {k:v for k,v in samples.items() if len(v)>=min_n}
        if len(samples) < 2:
            kw_rows.append({"Dim":f"P{i}","H":np.nan,"p":np.nan,"Sig.":"—"}); continue
        H, p = kruskal(*samples.values())
        kw_rows.append({"Dim":f"P{i}","H":round(H,3),"p":round(p,4),"Sig.":sig(p)})
        if p < .05:
            all_v = np.concatenate(list(samples.values()))
            all_l = np.concatenate([[k]*len(v) for k,v in samples.items()])
            ser   = pd.DataFrame({col:all_v, group_col:all_l})
            try:
                dunn_all[f"P{i}"] = sp.posthoc_dunn(ser, val_col=col,
                                                      group_col=group_col, p_adjust="bonferroni")
            except: pass
    return pd.DataFrame(kw_rows), dunn_all


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 4 — ANÁLISIS COMPARATIVO ENTRE HOSPITALES
# ══════════════════════════════════════════════════════════════════════════════

def analisis_hospitales(df, outdir):
    """
    Genera el bloque de comparativa general entre los hospitales
    *comparables* (UHD-Paliativos Adultos). Las unidades especializadas
    (Pediatría, Salud Mental, Salud Infanto-Juvenil) se procesan
    individualmente en informes propios — no entran en este comparativo
    por atender a poblaciones clínicamente distintas.

    Genera:
      · Figuras 02–07: radares, barras de percepción de preparación/importancia/pesos,
        heatmap de brecha y figura de perfiles en línea.
      · CSVs con test Mann-Whitney U para cada par de hospitales.
    Funciona automáticamente con cualquier número de hospitales nuevos.
    """
    # Filtrar a los grupos comparables
    df_comp = df[df["comparable"] == True].copy()
    if df_comp["hospital"].nunique() < 2:
        print("  ⚠ Menos de 2 unidades comparables (UHD): se omite la "
              "comparativa entre unidades. Cada unidad se describe en su "
              "informe individual.")
        return

    pal_all = df._pal_hosp
    pal = {k: v for k, v in pal_all.items() if k in df_comp["hospital"].unique()}
    print(f"  Hospitales comparables: {list(pal.keys())}")
    grupos_no_comp = sorted(df.loc[df["comparable"] == False,
                                   "hospital"].dropna().unique())
    if grupos_no_comp:
        print(f"  Excluidos del comparativo (analizados aparte): "
              f"{', '.join(grupos_no_comp)}")

    # Radares Importancia vs Percepción de preparación por hospital
    _radar(df_comp, "hospital", "MULTIPAL — Radar por Hospital", pal, outdir, "02_radar_hospitales")

    # Barras comparativas por dimensión
    _barras(df_comp, "hospital", "R", "Percepción de preparación por hospital",  "Percepción de preparación (0-100)", pal, outdir, "03_percepción de preparación_hospital")
    _barras(df_comp, "hospital", "I", "Importancia por hospital",  "Importancia (0-100)", pal, outdir, "04_importancia_hospital")
    _barras(df_comp, "hospital", "W", "Pesos de prioridad por hospital", "Peso (%)", pal, outdir, "05_pesos_hospital", ylim=60)

    # Heatmap de brecha
    _heatmap_gap(df_comp, "hospital", "Brecha (Imp−Prep) por hospital", outdir, "06_gap_hospital")

    # Mann-Whitney U para cada par de hospitales comparables
    hospitales = sorted(df_comp["hospital"].dropna().unique())
    for i in range(len(hospitales)):
        for j in range(i+1, len(hospitales)):
            h1, h2 = hospitales[i], hospitales[j]
            tag = f"{h1[:6].replace(' ','')}_vs_{h2[:6].replace(' ','')}"
            for pref, lbl in [("R","preparacion"),("I","importancia")]:
                t = _mw(df_comp, "hospital", h1, h2, pref)
                t.to_csv(outdir/f"06_mw_{tag}_{lbl}.csv", index=False)

    # Figura de perfiles en línea
    fig, axs = plt.subplots(1, 3, figsize=(18, 6))
    fig.suptitle("MULTIPAL — Comparativa entre Hospitales (UHD-Paliativos Adultos)",
                 fontsize=13, fontweight="bold")
    for ax, (prefix, tit, yl) in zip(axs, [("R","Percepción de preparación",115),
                                             ("I","Importancia percibida",115),
                                             ("W","Pesos de prioridad (%)",60)]):
        for h in hospitales:
            sub = df_comp[df_comp["hospital"]==h]
            v   = [sub[f"{prefix}_P{i}"].mean() for i in range(1,6)]
            e   = [sub[f"{prefix}_P{i}"].sem()  for i in range(1,6)]
            col = pal.get(h, "#777")
            ax.errorbar(DIM, v, yerr=e, fmt="o-", color=col, lw=2, ms=7,
                        capsize=3, label=f"{h} (n={len(sub)})")
        ax.set_ylim(0, yl); ax.set_title(tit, fontweight="bold")
        ax.set_ylabel("Puntuación"); ax.legend(fontsize=9); ax.grid(alpha=.3)
    plt.tight_layout(); save(fig, "07_comparativa_hospitales", outdir)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 5 — ANÁLISIS POR SUBUNIDADES (DINÁMICO POR HOSPITAL)
# ══════════════════════════════════════════════════════════════════════════════

def analisis_subunidades(df, outdir):
    """
    [OBSOLETO desde la nueva taxonomía v4]

    Tras el refactor del mapeo unidad → grupo de análisis, cada unidad
    es ya su propio grupo de análisis y no contiene subunidades
    internas. La función se conserva como no-op para mantener la
    interfaz del pipeline. El análisis específico de cada unidad se
    realiza ahora en `generar_pdfs_por_unidad` (paso final del run),
    que produce un informe individual siguiendo la estructura
    estándar de 5 apartados.
    """
    print("  (Omitido: la nueva taxonomía trata cada unidad como su "
          "propio grupo. Ver informes individuales al final del pipeline.)")
    return


def _OBSOLETO_analisis_subunidades(df, outdir):
    """Versión antigua, conservada únicamente como referencia interna."""
    pal_unit  = df._pal_unit
    hospitales = sorted(df["hospital"].dropna().unique())
    counter    = 8   # número de figura inicial para este bloque

    for hosp in hospitales:
        sub_hosp = df[df["hospital"]==hosp].copy()
        unidades  = sub_hosp["unidad"].dropna().unique()

        # Si solo hay una unidad no hay nada que comparar dentro del hospital
        if len(unidades) < 2:
            print(f"  [{hosp}] Una sola unidad — se omite análisis desagregado")
            continue

        counts = sub_hosp["unidad"].value_counts()
        print(f"\n  [{hosp}] Subunidades:")
        for u in unidades:
            n = counts.get(u,0); warn = " ⚠ orientativo" if n<5 else ""
            print(f"    {u}: n={n}{warn}")

        tag  = hosp[:12].replace(" ","_")
        base = f"{counter:02d}_{tag}"

        # Radares, barras y heatmap de brecha para las subunidades del hospital
        _radar(sub_hosp, "unidad",
               f"MULTIPAL — Radares Subunidades {hosp}\n(Importancia vs Percepción de preparación)",
               pal_unit, outdir, f"{base}_a_radar", min_n=1)
        _barras(sub_hosp, "unidad", "R", f"Subunidades {hosp} — Percepción de preparación",
                "Percepción de preparación (0-100)", pal_unit, outdir, f"{base}_b_preparacion")
        _barras(sub_hosp, "unidad", "I", f"Subunidades {hosp} — Importancia",
                "Importancia (0-100)", pal_unit, outdir, f"{base}_c_importancia")
        _barras(sub_hosp, "unidad", "W", f"Subunidades {hosp} — Pesos de prioridad",
                "Peso (%)", pal_unit, outdir, f"{base}_d_pesos", ylim=60)
        _heatmap_gap(sub_hosp, "unidad",
                     f"Subunidades {hosp} — Brecha (Imp−Prep)",
                     outdir, f"{base}_e_gap")

        # Figura de perfiles en línea: Percepción de preparación / Importancia / Brecha
        fig, axs = plt.subplots(1, 3, figsize=(18, 6))
        fig.suptitle(f"MULTIPAL — Perfiles por Subunidad · {hosp}",
                     fontsize=13, fontweight="bold")
        for ax, (prefix, tit, yl) in zip(axs, [("R","Percepción de preparación",115),
                                                 ("I","Importancia",115),
                                                 ("G","Brecha (Imp−Prep)",45)]):
            for u in unidades:
                s   = sub_hosp[sub_hosp["unidad"]==u]
                v   = [s[f"{prefix}_P{i}"].mean() for i in range(1,6)]
                col = pal_unit.get(str(u), "#777"); nu = len(s)
                st  = "-" if nu>=5 else "--"   # línea discontinua si n<5
                ax.plot(DIM, v, f"o{st}", color=col, lw=2, ms=7,
                        label=f"{u} (n={nu}{'⚠' if nu<5 else ''})")
            if prefix == "G": ax.axhline(0, color="black", lw=1.2)
            ax.set_ylim(-10 if prefix=="G" else 0, yl)
            ax.set_title(tit, fontweight="bold")
            ax.set_ylabel("Puntos" if prefix=="G" else "Puntuación")
            ax.legend(fontsize=8); ax.grid(alpha=.3)
        plt.tight_layout(); save(fig, f"{base}_f_perfiles", outdir)

        # CSV con tabla de medias (W, I, R, G) para todas las subunidades
        rows = []
        for u in unidades:
            s = sub_hosp[sub_hosp["unidad"]==u]; row = {"Subunidad":u,"n":len(s)}
            for i in range(1,6):
                for p in ["W","I","R","G"]:
                    row[f"{p}_P{i}"] = round(s[f"{p}_P{i}"].mean(), 2)
            rows.append(row)
        pd.DataFrame(rows).to_csv(outdir/f"{base}_tabla_medias.csv", index=False)

        # Kruskal-Wallis entre subunidades (solo si hay n≥3 en al menos 2 subunidades)
        uok = [u for u in unidades if counts.get(u,0) >= 3]
        if len(uok) >= 2:
            df_ok = sub_hosp[sub_hosp["unidad"].isin(uok)]
            for pref, lbl in [("R","preparacion"),("I","importancia")]:
                kw, dunn = _kruskal_dunn(df_ok, "unidad", pref, min_n=3)
                kw.to_csv(outdir/f"{base}_kruskal_{lbl}.csv", index=False)
                for k, ddf in dunn.items():
                    ddf.round(4).to_csv(outdir/f"{base}_dunn_{lbl}_{k}.csv")

        counter += 1


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 6 — ANÁLISIS POR ROL PROFESIONAL
# ══════════════════════════════════════════════════════════════════════════════

def analisis_rol(df, outdir):
    """
    Genera radares, barras, heatmap y tests estadísticos comparando los roles
    profesionales (Médico, Enfermera, Fisioterapeuta, TCAE, etc.).
    Incluye Mann-Whitney U Médico vs Enfermera y Kruskal-Wallis global.
    """
    pal = make_palette(df["rol"].dropna().unique())
    pal.update(PAL_ROL)  # sobreescribe con los colores corporativos definidos
    _radar(df, "rol", "MULTIPAL — Radares por Rol Profesional", pal, outdir, "19_radar_rol")
    # Radar de la BRECHA (Imp − Preparación) por rol — un polígono por rol
    _radar_brecha_rol(df, "MULTIPAL — Brecha Importancia − Preparación por Rol",
                      outdir, "19_radar_brecha_rol")
    _barras(df, "rol", "R", "Percepción de preparación por rol", "Percepción de preparación (0-100)", pal, outdir, "20_preparacion_rol")
    _barras(df, "rol", "I", "Importancia por rol", "Importancia (0-100)", pal, outdir, "21_importancia_rol")
    _barras(df, "rol", "W", "Pesos de prioridad por rol", "Peso (%)", pal, outdir, "22_pesos_rol", ylim=60)
    _heatmap_gap(df, "rol", "Brecha (Imp−Prep) por Rol", outdir, "23_gap_rol")

    # Comparación específica Médico vs Enfermera (los roles más frecuentes)
    for pref, lbl in [("R","rend"),("I","imp")]:
        t = _mw(df, "rol", "Médico", "Enfermera", pref)
        t.to_csv(outdir/f"23_mw_medico_enfermera_{lbl}.csv", index=False)

    # Kruskal-Wallis para todos los roles con n≥2
    roles_ok = [r for r in df["rol"].dropna().unique() if (df["rol"]==r).sum()>=2]
    if len(roles_ok) >= 2:
        kw, _ = _kruskal_dunn(df[df["rol"].isin(roles_ok)], "rol", "R", min_n=2)
        kw.to_csv(outdir/"23_kruskal_rol.csv", index=False)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 7 — ANÁLISIS POR GÉNERO
# ══════════════════════════════════════════════════════════════════════════════

def analisis_genero(df, outdir):
    """
    Análisis completo por género (Femenino vs Masculino).
    Requiere n≥5 en cada género para ejecutarse (de lo contrario se omite).
    Genera una figura de 6 paneles con:
      · Barras comparativas de percepción de preparación, importancia y pesos
      · Brecha por género
      · Boxplot de distribución del percepción de preparación global
      · Tabla resumen con resultados Mann-Whitney U
    """
    nF = (df["genero"]=="Femenino").sum(); nM = (df["genero"]=="Masculino").sum()
    if nF<5 or nM<5:
        print(f"  [Género] ⚠ Sin suficientes participantes (F={nF}, M={nM})"); return
    print(f"  [Género]  Femenino n={nF} · Masculino n={nM}")

    fig, axs = plt.subplots(2, 3, figsize=(18, 10))
    fig.suptitle(f"MULTIPAL — Análisis por Género   Femenino n={nF} · Masculino n={nM}",
                 fontsize=13, fontweight="bold")

    mw_res = {}
    # Fila superior: barras para Percepción de preparación, Importancia y Pesos
    for ci, (prefix, tit, yl) in enumerate([("R","Percepción de preparación (0-100)",115),
                                              ("I","Importancia (0-100)",115),
                                              ("W","Pesos de prioridad (%)",60)]):
        ax = axs[0,ci]; x = np.arange(5)
        for ji, (g, col) in enumerate([("Femenino","#E91E8C"),("Masculino","#003D7C")]):
            sub = df[df["genero"]==g]
            m   = [sub[f"{prefix}_P{i}"].mean() for i in range(1,6)]
            e   = [sub[f"{prefix}_P{i}"].sem()  for i in range(1,6)]
            ax.bar(x+(ji-.5)*.38, m, .38, color=col, label=f"{g} (n={len(sub)})",
                   yerr=e, capsize=3, alpha=.85)
        ax.set_xticks(x); ax.set_xticklabels(DIM)
        ax.set_title(tit, fontweight="bold"); ax.set_ylim(0, yl)
        ax.legend(fontsize=9); ax.grid(axis="y", alpha=.3)

        # Anota estrellas de significación sobre las barras donde hay diferencias
        mw = _mw(df, "genero", "Femenino", "Masculino", prefix)
        mw_res[prefix] = mw; mw.to_csv(outdir/f"14_mw_genero_{prefix}.csv", index=False)
        for idx2, row in mw.iterrows():
            if row["Sig."] not in ["ns","—"]:
                ax.text(idx2, yl*.95, row["Sig."], ha="center", fontsize=11,
                        color="red", fontweight="bold")

    # Panel: Brecha por género
    ax = axs[1,0]
    for g, col in [("Femenino","#E91E8C"),("Masculino","#003D7C")]:
        sub = df[df["genero"]==g]
        ax.plot(DIM, [sub[f"G_P{i}"].mean() for i in range(1,6)],
                "o-", color=col, lw=2, ms=7, label=f"{g} (n={len(sub)})")
    ax.axhline(0, color="black", lw=1.2)
    ax.set_title("Brecha (Imp−Prep) por género", fontweight="bold")
    ax.set_ylabel("Puntos"); ax.legend(); ax.grid(alpha=.3)

    # Panel: Boxplot distribución global de percepción de preparación por género
    ax2 = axs[1,1]
    for ji, (g, col) in enumerate([("Femenino","#E91E8C"),("Masculino","#003D7C")]):
        vals = df[df["genero"]==g][[f"R_P{i}" for i in range(1,6)]].values.flatten()
        ax2.boxplot(vals, positions=[ji], patch_artist=True, widths=.4,
                    boxprops=dict(facecolor=col, alpha=.5),
                    medianprops=dict(color="black", lw=2))
    ax2.set_xticks([0,1]); ax2.set_xticklabels(["Femenino","Masculino"])
    ax2.set_title("Distribución percepción de preparación global", fontweight="bold")
    ax2.set_ylabel("Percepción de preparación (0-100)"); ax2.set_ylim(0, 110)

    # Panel: Tabla resumen de resultados Mann-Whitney
    axs[1,2].axis("off")
    tabla = [["Dim","U Prep","Sig. R","U Imp","Sig. I"]]
    for _, r in mw_res["R"].iterrows():
        ri = mw_res["I"][mw_res["I"]["Dim"]==r["Dim"]].iloc[0]
        tabla.append([r["Dim"], f"{r['U']:.0f}" if not pd.isna(r['U']) else "—",
                      r["Sig."], f"{ri['U']:.0f}" if not pd.isna(ri['U']) else "—", ri["Sig."]])
    t = axs[1,2].table(cellText=tabla[1:], colLabels=tabla[0],
                        cellLoc="center", loc="center", bbox=[0,.05,1,.85])
    t.auto_set_font_size(False); t.set_fontsize(10)
    axs[1,2].set_title("Test Mann-Whitney U · Género\n(*p<.05  **p<.01  ***p<.001)",
                        fontsize=9, fontweight="bold")
    plt.tight_layout(); save(fig, "14_analisis_genero", outdir)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 8 — ANÁLISIS POR GRUPOS DE EDAD
# ══════════════════════════════════════════════════════════════════════════════

def analisis_edad(df, outdir):
    """
    Análisis por grupos de edad (≤35, 36-45, 46-55, >55).
    Genera:
      · Boxplots de Percepción de preparación, Importancia y Pesos por grupo de edad (figura 15)
      · Heatmap de medias por grupo de edad (figura 16)
      · Barras de percepción de preparación e importancia (figuras 17 y 18)
      · CSVs con Kruskal-Wallis y post-hoc de Dunn
    Anota en rojo las dimensiones con diferencias significativas entre grupos.
    """
    age_cats = ["≤35","36-45","46-55",">55"]
    age_ok   = [g for g in age_cats if g in df["grupo_edad"].cat.categories]
    counts   = {g: int((df["grupo_edad"]==g).sum()) for g in age_ok}
    print(f"  [Edad]  {counts}")

    # Tests estadísticos guardados como CSV
    for pref, lbl in [("R","preparacion"),("I","importancia"),("W","pesos")]:
        kw, dunn = _kruskal_dunn(df, "grupo_edad", pref, min_n=2)
        kw.to_csv(outdir/f"15_kruskal_edad_{lbl}.csv", index=False)
        for k, ddf in dunn.items():
            ddf.round(4).to_csv(outdir/f"15_dunn_edad_{lbl}_{k}.csv")

    # Panel 3×5: boxplot por dimensión y grupo de edad
    fig, axs = plt.subplots(3, 5, figsize=(22, 14))
    fig.suptitle("MULTIPAL — Análisis por Grupos de Edad\n"
                 "(Kruskal-Wallis + Dunn post-hoc)", fontsize=14, fontweight="bold")
    age_lbl = [f"{g}\n(n={counts.get(g,0)})" for g in age_ok]
    for ci, dim in enumerate(DIM):
        i = ci+1
        for ri, (prefix, yl, ylabel) in enumerate([("R",115,"Percepción de preparación"),
                                                    ("I",115,"Importancia"),
                                                    ("W",60,"Peso %")]):
            ax   = axs[ri,ci]
            data = [df[df["grupo_edad"]==g][f"{prefix}_P{i}"].dropna().values for g in age_ok]
            bp   = ax.boxplot(data, patch_artist=True, notch=False, widths=.55)
            for patch, g in zip(bp["boxes"], age_ok):
                patch.set_facecolor(PAL_AGE.get(g,"#999")); patch.set_alpha(.6)
            for med in bp["medians"]: med.set(color="black", lw=2)
            ax.set_xticks(range(1,len(age_ok)+1))
            ax.set_xticklabels(age_lbl, fontsize=7.5, rotation=20)
            ax.set_title(DLBL[dim], fontsize=8.5); ax.set_ylim(0, yl)
            if ci==0: ax.set_ylabel(ylabel, fontsize=9)
            # Indica significación estadística en el título del panel
            kw, _ = _kruskal_dunn(df, "grupo_edad", prefix, min_n=2)
            row_kw = kw[kw["Dim"]==dim]
            if not row_kw.empty and row_kw["Sig."].values[0] not in ["ns","—"]:
                ax.set_title(DLBL[dim]+f"\n{row_kw['Sig.'].values[0]}",
                             fontsize=8.5, color="red")
    plt.tight_layout(); save(fig, "15_edad_boxplots", outdir)

    # Heatmap de medias por grupo de edad
    fig2, axs2 = plt.subplots(1, 3, figsize=(18, 4.5))
    fig2.suptitle("MULTIPAL — Medias por Grupo de Edad", fontsize=12, fontweight="bold")
    for ax2, (pref, tit, cm) in zip(axs2, [("R","Percepción de preparación","Blues"),
                                             ("I","Importancia","Greens"),
                                             ("W","Pesos (%)","Oranges")]):
        mat = pd.DataFrame(
            [[df[df["grupo_edad"]==g][f"{pref}_P{i}"].mean() for i in range(1,6)]
             for g in age_ok],
            index=[f"{g}(n={counts.get(g,0)})" for g in age_ok], columns=DIM)
        sns.heatmap(mat, annot=True, fmt=".1f", cmap=cm, linewidths=.5, ax=ax2,
                    cbar_kws={"label":"Media"})
        ax2.set_title(tit, fontweight="bold")
    plt.tight_layout(); save(fig2, "16_heatmap_edad", outdir)

    _barras(df, "grupo_edad", "R", "Percepción de preparación por grupo edad", "Percepción de preparación (0-100)",
            PAL_AGE, outdir, "17_percepción de preparación_edad")
    _barras(df, "grupo_edad", "I", "Importancia por grupo edad", "Importancia (0-100)",
            PAL_AGE, outdir, "18_importancia_edad")


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 9 — ANÁLISIS POR AÑOS DE EXPERIENCIA
# ══════════════════════════════════════════════════════════════════════════════

def analisis_experiencia(df, outdir):
    """
    Análisis de la relación entre experiencia profesional y las valoraciones.
    Incluye:
      · Kruskal-Wallis entre grupos de experiencia (<1, 1-3, 3-5, >5 años)
      · Correlación de Spearman entre experiencia (ordinal numérica) y puntuaciones
      · Boxplots (figura 24)
    La correlación de Spearman detecta tendencias monotónicas aunque no lineales.
    """
    for pref, lbl in [("R","preparacion"),("I","importancia")]:
        kw, _ = _kruskal_dunn(df, "exp_label", pref, min_n=2)
        kw.to_csv(outdir/f"24_kruskal_exp_{lbl}.csv", index=False)

        # Correlación de Spearman: experiencia numérica (1-4) vs puntuación
        rows = []
        for i in range(1,6):
            v = df[["exp_num", f"{pref}_P{i}"]].dropna()
            if len(v)<5:
                rows.append({"Dim":f"P{i}","rho":np.nan,"p":np.nan,"Sig.":"—"}); continue
            rho, p = spearmanr(v["exp_num"], v[f"{pref}_P{i}"])
            rows.append({"Dim":f"P{i}","rho":round(rho,3),"p":round(p,4),"Sig.":sig(p)})
        pd.DataFrame(rows).to_csv(outdir/f"24_spearman_exp_{lbl}.csv", index=False)

    # Boxplots de Percepción de preparación e Importancia por grupo de experiencia
    exp_cats = [e for e in ["<1 año","1-3 años","3-5 años",">5 años"]
                if e in df["exp_label"].cat.categories]
    fig, axs = plt.subplots(2, 5, figsize=(22, 9))
    fig.suptitle("MULTIPAL — Percepción de preparación e Importancia por Experiencia",
                 fontsize=13, fontweight="bold")
    for ci, dim in enumerate(DIM):
        i = ci+1
        for ri, (prefix, ylabel) in enumerate([("R","Percepción de preparación"),("I","Importancia")]):
            ax   = axs[ri,ci]
            data = [df[df["exp_label"]==e][f"{prefix}_P{i}"].dropna().values for e in exp_cats]
            bp   = ax.boxplot(data, patch_artist=True, widths=.55)
            for patch, e in zip(bp["boxes"], exp_cats):
                patch.set_facecolor(PAL_EXP.get(e,"#999")); patch.set_alpha(.6)
            for med in bp["medians"]: med.set(color="black", lw=2)
            ax.set_xticks(range(1,len(exp_cats)+1))
            ax.set_xticklabels(exp_cats, fontsize=7.5, rotation=20)
            ax.set_title(DLBL[dim], fontsize=8.5); ax.set_ylim(0, 115)
            if ci==0: ax.set_ylabel(ylabel)
    plt.tight_layout(); save(fig, "24_experiencia_boxplots", outdir)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 10 — CONSENSO W DE KENDALL
# ══════════════════════════════════════════════════════════════════════════════

def analisis_consenso(df, outdir):
    """
    Calcula el W de Kendall (coeficiente de concordancia) para medir
    el grado de acuerdo interno en la asignación de pesos de prioridad.
    Se calcula globalmente y para cada subgrupo (hospital, unidad, rol, edad, experiencia).

    Interpretación:
      W < 0.3  → acuerdo bajo
      W ≈ 0.5  → acuerdo moderado
      W ≥ 0.7  → acuerdo alto (los profesionales priorizan de forma similar)

    Genera figura 25 con barras por hospital, rol y grupo de edad,
    con líneas de referencia en W=0.5 y W=0.7.
    """
    def _w(sub):
        mat = sub[[f"W_P{i}" for i in range(1,6)]].dropna().values
        if len(mat) < 2: return np.nan
        # Convertir pesos a rangos por fila (inverso: mayor peso → rango 1)
        ranks = np.apply_along_axis(
            lambda x: len(x)+1-stats.rankdata(x,"average"), 1, mat)
        return kendall_w(ranks)

    rows = [{"Agrupación":"Global","Grupo":"Todos","n":len(df),"W":_w(df)}]
    for gc in ["hospital","unidad","rol","grupo_edad","exp_label"]:
        for g, sub in df.groupby(gc, observed=True):
            rows.append({"Agrupación":gc,"Grupo":str(g),"n":len(sub),"W":_w(sub)})
    w_df = pd.DataFrame(rows)
    w_df.to_csv(outdir/"25_consenso_kendall.csv", index=False)

    # Figura con barras del W por hospital, rol y grupo de edad
    fig, axs = plt.subplots(1, 3, figsize=(18, 6))
    fig.suptitle("MULTIPAL — Consenso interprofesional (W de Kendall)\n"
                 "Interpretación: 0 = sin consenso · 0.5 = moderado · ≥0.7 = alto",
                 fontsize=13, fontweight="bold")
    for ax, (agr, tit) in zip(axs, [("hospital","Por hospital"),
                                      ("rol","Por rol profesional"),
                                      ("grupo_edad","Por grupo de edad")]):
        sub = w_df[w_df["Agrupación"]==agr].dropna(subset=["W"])
        if len(sub)==0: ax.set_visible(False); continue
        pal_dyn = make_palette(sub["Grupo"].tolist())
        pal_dyn.update({**df._pal_hosp, **PAL_ROL, **PAL_AGE})
        colors  = [pal_dyn.get(g,"#777") for g in sub["Grupo"]]
        ax.bar(sub["Grupo"], sub["W"], color=colors, width=.5)
        ax.axhline(.7, color="red",    ls="--", lw=1.3, label="Alto (≥0.7)")
        ax.axhline(.5, color="orange", ls="--", lw=1.3, label="Moderado (≥0.5)")
        ax.set_ylim(0,1); ax.set_title(tit, fontweight="bold")
        ax.set_ylabel("W de Kendall"); ax.legend(fontsize=8)
        ax.tick_params(axis="x", rotation=20)
        for b in ax.patches:
            ax.text(b.get_x()+b.get_width()/2, b.get_height()+.02,
                    f"{b.get_height():.2f}", ha="center", fontsize=10)
    plt.tight_layout(); save(fig, "25_consenso_kendall", outdir)
    print(f"     W global = {_w(df):.3f}")


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 11 — MÉTODOS MCDA: TOPSIS · AHP · PROMETHEE II
# ══════════════════════════════════════════════════════════════════════════════

# ── TOPSIS ───────────────────────────────────────────────────────────────────
def topsis(dm, weights):
    """
    TOPSIS — Technique for Order Preference by Similarity to Ideal Solution.
    Todos los criterios son de beneficio (mayor percepción de preparación = mejor).

    Pasos:
     1. Normalización vectorial  rij = dij / √(Σ dkj²)
     2. Matriz ponderada          vij = wj × rij
     3. Ideal positivo A+ = max_i(vij),  Ideal negativo A- = min_i(vij)
     4. Distancia euclídea d+i = ‖vi − A+‖,  d−i = ‖vi − A−‖
     5. Coeficiente de cercanía  Ci = d−i / (d+i + d−i)
     6. Ranking: mayor Ci → mejor alternativa

    Interpretación de Ci:
     · Ci = 1.0  →  la alternativa ES la solución ideal
     · Ci = 0.0  →  la alternativa ES la peor solución posible
     · Ci > 0.5  →  más cerca del ideal que del anti-ideal  ✓
     · Ci < 0.5  →  más cerca del anti-ideal                ✗
    """
    dm = dm.astype(float)
    norms = np.sqrt((dm**2).sum(axis=0)); norms[norms==0] = 1e-12
    R = dm/norms; V = R*weights
    A_pos = V.max(axis=0); A_neg = V.min(axis=0)
    d_pos = np.sqrt(((V - A_pos)**2).sum(axis=1))
    d_neg = np.sqrt(((V - A_neg)**2).sum(axis=1))
    C     = d_neg / (d_pos + d_neg + 1e-12)
    return C, d_pos, d_neg, stats.rankdata(-C, method="min").astype(int)


# ── AHP ──────────────────────────────────────────────────────────────────────
def ahp_calcular(weights_pct):
    """
    AHP — Analytic Hierarchy Process (valoración directa con pesos porcentuales).
    Los encuestados asignaron pesos % directamente → equivale al vector AHP.

    Pasos:
     1. Normalizar:  w = pcts / Σpcts
     2. Matriz implícita de comparación por pares:  A[i,j] = w[i]/w[j]
     3. Vector de consistencia:  cv[i] = (A·w)[i] / w[i]
     4. λmax = media(cv)
     5. CI   = (λmax − n) / (n−1)
     6. CR   = CI / RI   [RI_5 = 1.12 según Saaty 1980]
     7. Si CR ≤ 0.10 → juicios consistentes
     8. Score_alternativa = Σ w_j × percepción de preparación_j

    Interpretación:
     · CR < 0.10:  consistencia aceptable (los pesos son coherentes entre sí)
     · CR > 0.10:  los encuestados fueron incoherentes al repartir pesos
     · Score: percepción de preparación global ponderada por importancia → mayor = mejor
    """
    n = len(weights_pct)
    w = weights_pct/weights_pct.sum(); w = np.where(w==0, 1e-12, w)
    A = np.outer(w, 1.0/w)                # matriz de comparación por pares implícita
    lam = np.mean((A@w)/w)                # autovalor principal λmax
    CI  = (lam - n)/(n - 1)              # índice de consistencia
    RI  = {3:.58, 4:.90, 5:1.12, 6:1.24, 7:1.32, 8:1.41, 9:1.45, 10:1.49}
    CR  = CI / RI.get(n, 1.12)           # ratio de consistencia (< 0.10 = ok)
    return {"w":w, "lambda_max":lam, "CI":CI, "CR":CR, "consistent":CR<=.10, "A":A}


def aplicar_ahp(df, group_col, outdir, tag):
    """
    Aplica AHP a cada grupo de la columna group_col usando los pesos medios W_Pi.
    Calcula el Score AHP ponderando la percepción de preparación por los pesos normalizados.
    Devuelve un DataFrame con los resultados y los guarda en CSV.
    """
    groups = df[group_col].dropna().unique()
    recs   = []
    for g in groups:
        sub  = df[df[group_col]==g]
        wm   = np.array([sub[f"W_P{i}"].mean() for i in range(1,6)])
        res  = ahp_calcular(wm)
        score = sum(res["w"][i] * sub[f"R_P{i+1}"].mean() for i in range(5))
        recs.append({"Grupo":g, "n":len(sub),
                     **{f"w_AHP_P{i+1}":round(res["w"][i]*100, 2) for i in range(5)},
                     "λmax":round(res["lambda_max"],4),
                     "CI":round(res["CI"],4), "CR":round(res["CR"],4),
                     "Consistente":"✓" if res["consistent"] else "✗",
                     "AHP_Score":round(score, 4)})
    df_ahp = pd.DataFrame(recs)
    df_ahp["AHP_Rank"] = stats.rankdata(-df_ahp["AHP_Score"], method="min").astype(int)
    df_ahp.sort_values("AHP_Rank").to_csv(outdir/f"{tag}_ahp.csv", index=False)
    return df_ahp


# ── PROMETHEE II ─────────────────────────────────────────────────────────────
def promethee_ii(dm, weights):
    """
    PROMETHEE II — Preference Ranking Organization Method for Enrichment Evaluations.
    Función de preferencia LINEAL (Tipo V de Brans, sin umbral de indiferencia).

    Pasos por criterio k:
     1. dk(a,b) = performance_k(a) − performance_k(b)
     2. Pk(a,b) = max(0, min(1, dk(a,b) / rango_k))   [preferencia normalizada]
     3. π(a,b)  = Σk wk · Pk(a,b)                      [índice de preferencia global]
     4. φ+(a)   = 1/(m−1) · Σb π(a,b)                  [flujo saliente]
     5. φ−(a)   = 1/(m−1) · Σb π(b,a)                  [flujo entrante]
     6. φ(a)    = φ+(a) − φ−(a)                         [flujo neto]
     7. Ranking: mayor φ → mejor alternativa

    Interpretación de φ:
     · φ > 0:  la alternativa SUPERA a más de las que es superada
     · φ < 0:  la alternativa ES SUPERADA por más de las que supera
     · φ = 0:  equilibrio perfecto
    """
    m, n = dm.shape
    # Índice de preferencia agregado π(a,b): grado en que a se prefiere a b
    pi = np.zeros((m, m))
    for k in range(n):
        col = dm[:, k].astype(float)
        rng = col.max() - col.min()
        if rng < 1e-10:
            continue                       # criterio sin variabilidad → se ignora
        for a in range(m):
            for b in range(m):
                if a == b:
                    continue
                d = col[a] - col[b]
                pref = max(0.0, min(1.0, d / rng))   # P_k(a,b), tipo V (lineal)
                pi[a, b] += weights[k] * pref
    # Flujo saliente φ+(a) = media de π(a,·); flujo entrante φ−(a) = media de π(·,a)
    phi_p = pi.sum(axis=1) / (m - 1)
    phi_n = pi.sum(axis=0) / (m - 1)
    phi_net = phi_p - phi_n
    return phi_p, phi_n, phi_net, stats.rankdata(-phi_net, method="min").astype(int)


# ── APLICACIÓN Y COMPARACIÓN DE LOS 3 MÉTODOS ────────────────────────────────
def analisis_mcda(df, group_col, palette, outdir, tag, titulo):
    """
    Orquesta la aplicación de TOPSIS, AHP y PROMETHEE II al mismo conjunto de grupos.
    Genera tres figuras y varios CSVs:
      · {tag}_mcda.png              : figura principal de 8 paneles
      · {tag}_mcda_interpretacion.png : figura interpretativa con barras explicadas
      · {tag}_ahp_consistencia.png  : figura de consistencia AHP
      · {tag}_mcda_rankings.csv     : tabla consolidada de rankings
      · {tag}_ahp.csv               : detalle de scores y consistencia AHP

    La tabla de rankings colorea verde (acuerdo), amarillo (parcial) o rojo (diverge)
    según si los tres métodos coinciden en el ranking de cada grupo.
    """
    groups_ok = [g for g in df[group_col].dropna().unique()
                 if (df[group_col]==g).sum() >= 1]
    if len(groups_ok) < 2:
        print(f"     ⚠ {tag}: menos de 2 grupos con n≥2, omitido"); return

    sub = df[df[group_col].isin(groups_ok)].copy()

    # Matriz de decisión: filas=grupos, columnas=5 dimensiones (percepción de preparación media)
    dm = np.array([[sub[sub[group_col]==g][f"R_P{i}"].mean()
                    for i in range(1,6)] for g in groups_ok])
    # Vector de pesos normalizados (media de todos los encuestados)
    wm = np.array([sub[f"W_P{i}"].mean() for i in range(1,6)])
    wn = wm / wm.sum()

    # Ejecutar los tres métodos
    C, d_pos, d_neg, t_ranks      = topsis(dm, wn)
    ahp_df                         = aplicar_ahp(sub, group_col, outdir, tag)
    phi_p, phi_n, phi_net, p_ranks = promethee_ii(dm, wn)
    ahp_scores = {r["Grupo"]:r["AHP_Score"] for _,r in ahp_df.iterrows()}
    ahp_ranks  = {r["Grupo"]:r["AHP_Rank"]  for _,r in ahp_df.iterrows()}

    # DataFrame consolidado de resultados
    result = pd.DataFrame({
        "Grupo":          groups_ok,
        "n":              [(sub[group_col]==g).sum() for g in groups_ok],
        "TOPSIS_Ci":      C.round(4),
        "TOPSIS_d+":      d_pos.round(4),
        "TOPSIS_d-":      d_neg.round(4),
        "TOPSIS_Rank":    t_ranks,
        "AHP_Score":      [ahp_scores.get(g, np.nan) for g in groups_ok],
        "AHP_Rank":       [ahp_ranks.get(g, np.nan)  for g in groups_ok],
        "PROMETHEE_phi+": phi_p.round(4),
        "PROMETHEE_phi-": phi_n.round(4),
        "PROMETHEE_phi":  phi_net.round(4),
        "PROMETHEE_Rank": p_ranks,
    })
    result.to_csv(outdir/f"{tag}_mcda_rankings.csv", index=False)

    colors = [palette.get(str(g),"#777") for g in groups_ok]

    # ── FIGURA PRINCIPAL MCDA (8 paneles) ────────────────────────────────────
    fig = plt.figure(figsize=(22, 16))
    fig.suptitle(f"MULTIPAL — Análisis Multicriterio (TOPSIS · AHP · PROMETHEE II)\n{titulo}",
                 fontsize=14, fontweight="bold")
    gs = gridspec.GridSpec(3, 3, figure=fig, hspace=.52, wspace=.40)

    # Panel 1: Barras TOPSIS con coeficiente Ci
    ax1 = fig.add_subplot(gs[0,0])
    t_s  = result.sort_values("TOPSIS_Rank")
    ax1.barh(t_s["Grupo"].astype(str), t_s["TOPSIS_Ci"], color=colors)
    ax1.axvline(.5, color="gray", ls="--", lw=1.2, label="Umbral 0.5")
    ax1.set_xlim(0, 1.1); ax1.set_title("TOPSIS — Coeficiente Ci", fontweight="bold", fontsize=11)
    ax1.set_xlabel("Ci  (0=peor · 1=ideal · >0.5=bueno)")
    ax1.legend(fontsize=8)
    for b, r in zip(ax1.patches, t_s["TOPSIS_Rank"]):
        v = b.get_width()
        ax1.text(v+.01, b.get_y()+b.get_height()/2, f"{v:.3f}  (#{int(r)})",
                 va="center", fontsize=9, fontweight="bold")

    # Panel 2: Barras AHP Score ponderado
    ax2 = fig.add_subplot(gs[0,1])
    a_s  = result.sort_values("AHP_Rank")
    ax2.barh(a_s["Grupo"].astype(str), a_s["AHP_Score"], color=colors)
    ax2.set_title("AHP — Score ponderado", fontweight="bold", fontsize=11)
    ax2.set_xlabel("Score = Σ(w_AHP × percepción de preparación)  [0-100]")
    for b, r in zip(ax2.patches, a_s["AHP_Rank"]):
        v = b.get_width()
        ax2.text(v+.3, b.get_y()+b.get_height()/2, f"{v:.1f}  (#{int(r)})",
                 va="center", fontsize=9, fontweight="bold")

    # Panel 3: Barras PROMETHEE II flujo neto φ
    ax3 = fig.add_subplot(gs[0,2])
    p_s  = result.sort_values("PROMETHEE_Rank")
    ax3.barh(p_s["Grupo"].astype(str), p_s["PROMETHEE_phi"], color=colors)
    ax3.axvline(0, color="black", lw=1.2)
    ax3.set_title("PROMETHEE II — Flujo neto φ", fontweight="bold", fontsize=11)
    ax3.set_xlabel("φ neto  (>0=supera · <0=es superado)")
    for b, r in zip(ax3.patches, p_s["PROMETHEE_Rank"]):
        v = b.get_width()
        ax3.text(v+.003, b.get_y()+b.get_height()/2, f"{v:.3f}  (#{int(r)})",
                 va="center", fontsize=9, fontweight="bold")

    # Panel 4: Tabla consolidada de rankings con semáforo de consenso
    ax4 = fig.add_subplot(gs[1,0:2]); ax4.axis("off")
    tabla = [["Grupo","n","TOPSIS Ci","TOPSIS #","AHP Score","AHP #","PROMETHEE φ","PROMETHEE #","Consenso"]]
    for _, r in result.sort_values("TOPSIS_Rank").iterrows():
        ranks   = [int(r["TOPSIS_Rank"]), int(r["AHP_Rank"]), int(r["PROMETHEE_Rank"])]
        consenso = "✓ Acuerdo" if len(set(ranks))==1 else \
                   "≈ Parcial"  if max(ranks)-min(ranks)<=1 else "✗ Diverge"
        tabla.append([str(r["Grupo"]), int(r["n"]),
                      f"{r['TOPSIS_Ci']:.3f}", f"#{int(r['TOPSIS_Rank'])}",
                      f"{r['AHP_Score']:.1f}", f"#{int(r['AHP_Rank'])}",
                      f"{r['PROMETHEE_phi']:.3f}", f"#{int(r['PROMETHEE_Rank'])}",
                      consenso])
    tbl = ax4.table(cellText=tabla[1:], colLabels=tabla[0],
                    cellLoc="center", loc="center", bbox=[0,0,1,1])
    tbl.auto_set_font_size(False); tbl.set_fontsize(9.5)
    tbl.auto_set_column_width(list(range(len(tabla[0]))))
    for i in range(1, len(tabla)):
        cons  = tabla[i][-1]
        color = "#D5F5E3" if "Acuerdo" in cons else \
                "#FEF9E7" if "Parcial" in cons else "#FADBD8"
        for j in range(len(tabla[0])): tbl[i,j].set_facecolor(color)
    ax4.set_title("Rankings consolidados y consenso entre métodos",
                  fontweight="bold", pad=12, fontsize=11)

    # Panel 5: Heatmap correlación Spearman entre rankings de los 3 métodos
    ax5 = fig.add_subplot(gs[1,2])
    rm   = result[["TOPSIS_Rank","AHP_Rank","PROMETHEE_Rank"]].dropna().astype(float)
    corr = rm.corr(method="spearman")
    corr.index   = ["TOPSIS","AHP","PROMETHEE"]
    corr.columns = ["TOPSIS","AHP","PROMETHEE"]
    sns.heatmap(corr, annot=True, fmt=".3f", cmap="RdYlGn", vmin=-1, vmax=1,
                linewidths=.5, ax=ax5, annot_kws={"size":13,"weight":"bold"})
    ax5.set_title("Correlación Spearman\nentre rankings de los 3 métodos",
                  fontweight="bold", fontsize=10)

    # Panel 6: TOPSIS — distancias al ideal positivo y negativo
    ax6 = fig.add_subplot(gs[2,0])
    x   = np.arange(len(groups_ok)); w2 = .35
    ax6.bar(x-.175, result["TOPSIS_d+"], w2, color="#E74C3C", alpha=.75, label="d+ (distancia al ideal)")
    ax6.bar(x+.175, result["TOPSIS_d-"], w2, color="#2E9E6B", alpha=.75, label="d- (distancia al anti-ideal)")
    ax6.set_xticks(x); ax6.set_xticklabels([str(g) for g in groups_ok], rotation=15, fontsize=8)
    ax6.set_title("TOPSIS — Distancias a soluciones ideales", fontweight="bold", fontsize=10)
    ax6.legend(fontsize=8); ax6.grid(axis="y", alpha=.3)
    ax6.set_ylabel("Distancia euclidea ponderada")

    # Panel 7: PROMETHEE II — flujos φ+ y φ−
    ax7 = fig.add_subplot(gs[2,1])
    ax7.bar(x-.175, result["PROMETHEE_phi+"], w2, color="#003D7C", alpha=.75, label="φ+ (flujo saliente)")
    ax7.bar(x+.175, result["PROMETHEE_phi-"], w2, color="#E8A020", alpha=.75, label="φ- (flujo entrante)")
    ax7.set_xticks(x); ax7.set_xticklabels([str(g) for g in groups_ok], rotation=15, fontsize=8)
    ax7.set_title("PROMETHEE II — Flujos φ+ y φ−", fontweight="bold", fontsize=10)
    ax7.legend(fontsize=8); ax7.grid(axis="y", alpha=.3)
    ax7.set_ylabel("Flujo de preferencia")

    # Panel 8: AHP — radar de pesos por grupo
    ax8 = fig.add_subplot(gs[2,2], polar=True)
    angles = radar_angles()
    for _, r in ahp_df.iterrows():
        g   = r["Grupo"]; col = palette.get(str(g),"#777")
        wg  = [r[f"w_AHP_P{i}"] for i in range(1,6)] + [r["w_AHP_P1"]]
        ax8.plot(angles, wg, "o-", color=col, lw=2, label=str(g))
        ax8.fill(angles, wg, color=col, alpha=.07)
    ax8.set_xticks(angles[:-1]); ax8.set_xticklabels(DIM)
    ax8.set_ylim(0, 55)
    ax8.set_title("AHP — Pesos por grupo (%)", fontweight="bold", pad=15, fontsize=10)
    ax8.legend(loc="lower right", fontsize=7)

    save(fig, f"{tag}_mcda", outdir)

    # Figuras adicionales: consistencia AHP e interpretación numérica
    _fig_ahp(ahp_df, titulo, palette, outdir, f"{tag}_ahp_consistencia")
    _fig_interpretacion_mcda(result, wn, groups_ok, palette, titulo, outdir, f"{tag}_mcda_interpretacion")

    return result


def _fig_ahp(ahp_df, titulo, palette, outdir, fname):
    """
    Figura auxiliar de AHP: barras de pesos normalizados por grupo y
    tabla con λmax, CI, CR y score final. Colorea verde si CR<0.10 (consistente).
    """
    fig, axs = plt.subplots(1, 2, figsize=(14, 5))
    fig.suptitle(f"MULTIPAL — AHP: Pesos y Consistencia\n{titulo}",
                 fontsize=12, fontweight="bold")
    ax = axs[0]; x = np.arange(5); ng = len(ahp_df); w2 = .7/ng
    for j, (_, row) in enumerate(ahp_df.iterrows()):
        g  = row["Grupo"]
        wv = [row[f"w_AHP_P{i}"] for i in range(1,6)]
        off = (j - ng/2 + .5) * w2
        ax.bar(x+off, wv, w2, color=palette.get(str(g),"#777"), alpha=.85,
               label=f"{g} (n={int(row['n'])})")
    ax.axhline(20, color="gray", ls="--", lw=1.2, label="Uniforme 20%")
    ax.set_xticks(x); ax.set_xticklabels(DIM); ax.set_ylim(0, 65)
    ax.set_title("Pesos AHP por dimensión", fontweight="bold")
    ax.set_ylabel("Peso AHP (%)"); ax.legend(fontsize=8); ax.grid(axis="y", alpha=.3)

    # Tabla de consistencia
    axs[1].axis("off")
    cols = ["Grupo","n","λmax","CI","CR","Consistente","AHP Score","AHP Rank"]
    data = [[str(r["Grupo"]), int(r["n"]), f"{r['λmax']:.4f}",
             f"{r['CI']:.4f}", f"{r['CR']:.4f}", r["Consistente"],
             f"{r['AHP_Score']:.2f}", f"#{int(r['AHP_Rank'])}"]
            for _, r in ahp_df.sort_values("AHP_Rank").iterrows()]
    tbl  = axs[1].table(cellText=data, colLabels=cols, cellLoc="center",
                         loc="center", bbox=[0,.05,1,.85])
    tbl.auto_set_font_size(False); tbl.set_fontsize(10)
    for i, row in enumerate(data):
        col = "#D5F5E3" if row[5]=="✓" else "#FADBD8"
        for j in range(len(cols)): tbl[i+1,j].set_facecolor(col)
    axs[1].set_title("Índice de Consistencia AHP\n(CR < 0.10 = aceptable · Saaty 1980)",
                      fontweight="bold", pad=20, fontsize=10)
    plt.tight_layout(); save(fig, fname, outdir)


def _fig_interpretacion_mcda(result, wn, groups_ok, palette, titulo, outdir, fname):
    """
    Figura complementaria con los valores numéricos de cada método y su
    interpretación textual anotada en el mismo gráfico.
    TOPSIS: verde si Ci≥0.5, rojo si Ci<0.5.
    AHP:    degradado de azul oscuro (mejor) a azul claro.
    PROMETHEE: verde si φ≥0 (supera), rojo si φ<0 (es superado).
    """
    m = len(groups_ok)
    fig, axs = plt.subplots(1, 3, figsize=(18, max(5, m*1.4+3)))
    fig.suptitle(f"MULTIPAL — Interpretación de resultados MCDA\n{titulo}",
                 fontsize=13, fontweight="bold")

    # Panel TOPSIS
    ax  = axs[0]
    t_s = result.sort_values("TOPSIS_Rank", ascending=False)
    y   = np.arange(len(t_s))
    ax.barh(y, t_s["TOPSIS_Ci"],
            color=["#2E9E6B" if v>=.5 else "#E74C3C" for v in t_s["TOPSIS_Ci"]],
            alpha=.85, height=.6)
    ax.axvline(.5, color="black", ls="--", lw=1.5, label="Umbral 0.5")
    ax.set_yticks(y); ax.set_yticklabels(t_s["Grupo"].astype(str), fontsize=9)
    ax.set_xlim(0, 1.25); ax.set_xlabel("Ci"); ax.set_title("TOPSIS", fontweight="bold", fontsize=12)
    ax.legend(fontsize=9)
    for bar, ci, rank in zip(ax.patches, t_s["TOPSIS_Ci"], t_s["TOPSIS_Rank"]):
        interp = "Cerca del ideal ✓" if ci>=.5 else "Cerca del anti-ideal ✗"
        ax.text(bar.get_width()+.01, bar.get_y()+bar.get_height()/2,
                f"{ci:.3f}  #{int(rank)}\n{interp}", va="center", fontsize=8)
    ax.text(.02, .98, "Ci > 0.5 = cerca del ideal\nCi < 0.5 = cerca del anti-ideal",
            transform=ax.transAxes, fontsize=8, va="top",
            bbox=dict(boxstyle="round", facecolor="lightyellow", alpha=.8))

    # Panel AHP
    ax2 = axs[1]
    a_s  = result.sort_values("AHP_Rank", ascending=False)
    y2   = np.arange(len(a_s))
    ax2.barh(y2, a_s["AHP_Score"],
             color=["#003D7C" if i==0 else "#5D8AA8" for i in range(len(a_s))],
             alpha=.85, height=.6)
    ax2.set_yticks(y2); ax2.set_yticklabels(a_s["Grupo"].astype(str), fontsize=9)
    ax2.set_xlabel("Score (0-100)"); ax2.set_title("AHP", fontweight="bold", fontsize=12)
    for bar, sc, rank in zip(ax2.patches, a_s["AHP_Score"], a_s["AHP_Rank"]):
        ax2.text(bar.get_width()+.3, bar.get_y()+bar.get_height()/2,
                 f"{sc:.1f}/100  #{int(rank)}", va="center", fontsize=9, fontweight="bold")
    ax2.text(.02, .98, "Score = percepción de preparación\nponderado por pesos AHP\n(mayor = mejor gestión global)",
             transform=ax2.transAxes, fontsize=8, va="top",
             bbox=dict(boxstyle="round", facecolor="lightyellow", alpha=.8))

    # Panel PROMETHEE II
    ax3 = axs[2]
    p_s  = result.sort_values("PROMETHEE_Rank", ascending=False)
    y3   = np.arange(len(p_s))
    ax3.barh(y3, p_s["PROMETHEE_phi"],
             color=["#2E9E6B" if v>=0 else "#E74C3C" for v in p_s["PROMETHEE_phi"]],
             alpha=.85, height=.6)
    ax3.axvline(0, color="black", lw=1.5)
    ax3.set_yticks(y3); ax3.set_yticklabels(p_s["Grupo"].astype(str), fontsize=9)
    ax3.set_xlabel("Flujo neto φ"); ax3.set_title("PROMETHEE II", fontweight="bold", fontsize=12)
    for bar, phi, rank in zip(ax3.patches, p_s["PROMETHEE_phi"], p_s["PROMETHEE_Rank"]):
        interp = "Supera a más ✓" if phi>=0 else "Es superado ✗"
        ax3.text(bar.get_width()+.003, bar.get_y()+bar.get_height()/2,
                 f"{phi:.3f}  #{int(rank)}\n{interp}", va="center", fontsize=8)
    ax3.text(.02, .98, "φ > 0 = supera a más alternativas\nφ < 0 = es superado por más\nφ = 0 = equilibrio",
             transform=ax3.transAxes, fontsize=8, va="top",
             bbox=dict(boxstyle="round", facecolor="lightyellow", alpha=.8))

    plt.tight_layout(); save(fig, fname, outdir)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 11.B — MCDA DE DIMENSIONES (refactor metodológico v6)
# ══════════════════════════════════════════════════════════════════════════════
#
# Nueva semántica del análisis multicriterio, según lo acordado con el
# equipo médico:
#
#   · ALTERNATIVAS = las 5 DIMENSIONES del cuidado paliativo (P1-P5)
#   · CRITERIOS    = las 3 PREGUNTAS del cuestionario:
#                       W = peso/prioridad asignado (suma 100 entre las 5)
#                       I = importancia individual (0-100, independiente)
#                       R = preparación percibida (0-100)
#   · PESOS de los criterios: equiponderados (1/3, 1/3, 1/3)
#   · Matriz de decisión: para cada DIMENSIÓN, media de las respuestas de
#                          los profesionales de esa subunidad en W, I y R
#
# Aplicamos TOPSIS, AHP y PROMETHEE II a esta matriz 5×3 para cada
# subunidad y obtenemos 3 rankings de las dimensiones. La consistencia
# entre métodos se mide con el coeficiente de correlación de Kendall
# entre los tres pares de rankings.
#
# Interpretación clínica del ranking: indica qué dimensión es más
# "central" o "preponderante" en la valoración del equipo, combinando
# la prioridad relativa (W), la importancia atribuida (I) y la
# preparación percibida (R). Una dimensión en posición #1 es aquella
# que más capta la atención del equipo en conjunto.
# ══════════════════════════════════════════════════════════════════════════════

# ──────────────────────────────────────────────────────────────────────────────
#  HELPERS PARA EL MCDA REFACTORIZADO (v7) — criterios W y R, W de Kendall
# ──────────────────────────────────────────────────────────────────────────────

def kendall_w_rankings(rankings_matriz):
    """
    Coeficiente W de Kendall para concordancia entre múltiples
    evaluadores (jueces). Recibe una matriz k×n donde k es el número
    de jueces (métodos MCDA en nuestro caso) y n el número de
    alternativas (las 5 dimensiones).

    Devuelve W ∈ [0, 1]:
      · W ≥ 0,7 → concordancia alta (rankings convergen)
      · 0,5 ≤ W < 0,7 → concordancia moderada
      · W < 0,5 → concordancia baja (divergencia)
    """
    ranks = np.asarray(rankings_matriz, dtype=float)
    k, n = ranks.shape
    if n <= 1 or k <= 1:
        return np.nan
    R = ranks.sum(axis=0)   # suma de rangos por objeto
    S = np.sum((R - np.mean(R)) ** 2)
    denom = k ** 2 * (n ** 3 - n)
    if denom == 0:
        return np.nan
    return 12 * S / denom


def analisis_mcda_dimensiones(df_unit, hosp, slug, outdir):
    """
    NUEVO MCDA REFACTORIZADO (v7, según indicaciones de los médicos)
    ────────────────────────────────────────────────────────────────────
    Aplica TOPSIS, AHP y PROMETHEE II a las 5 dimensiones (alternativas)
    usando solo DOS criterios (W y R, no I), con el objetivo de
    identificar QUÉ DIMENSIONES MÁS REQUIEREN INTERVENCIÓN dentro de
    cada unidad.

    Criterios de entrada:
      · W = peso medio asignado a la dimensión (refleja la prioridad
            relativa bajo restricción de reparto). Criterio de
            BENEFICIO: más peso → más prioritaria.
      · R = percepción de preparación media percibido en la dimensión. Criterio de
            COSTE: menos percepción de preparación → más necesidad de intervención
            (se transforma a 100-R para tratarlo como beneficio).

    Se excluye explícitamente la importancia (I) como criterio porque,
    al no estar sometida a restricción de reparto, permite puntuar
    todas las dimensiones igual de alto sin forzar el trade-off
    informativo que el MCDA necesita.

    Concordancia entre métodos:
      · Coeficiente W de Kendall entre los 3 rankings (no Kendall τ
        pairwise como en v6). Umbral W ≥ 0,7 = convergencia.

    Ranking consensuado final:
      · Media de los 3 rankings (TOPSIS, AHP, PROMETHEE) reordenada.
    """
    n = len(df_unit)
    if n < 1:
        return {"error": "Sin datos en la subunidad"}

    # ── Matriz de decisión 5 dimensiones × 2 criterios (W, 100-R) ───────
    medias_W = np.array([df_unit[f"W_P{i}"].mean() for i in range(1, 6)])
    medias_R = np.array([df_unit[f"R_P{i}"].mean() for i in range(1, 6)])

    if np.any(np.isnan(medias_W)) or np.any(np.isnan(medias_R)):
        return {"error": "Datos insuficientes para construir la matriz"}

    # Matriz original para mostrar en la tabla (W, R sin transformar)
    matriz_original = np.column_stack([medias_W, medias_R])
    matriz_df = pd.DataFrame(matriz_original, index=DIM,
                              columns=["W (peso medio %)",
                                       "R (preparación media)"])

    # Matriz transformada para los métodos MCDA (W beneficio, 100-R beneficio)
    matriz_mcda = np.column_stack([medias_W, 100 - medias_R])

    # Pesos de los criterios: equiponderados
    pesos_criterios = np.array([0.5, 0.5])

    # ── TOPSIS ──────────────────────────────────────────────────────────
    C_topsis, dpos, dneg, rank_topsis = topsis(matriz_mcda, pesos_criterios)

    # ── AHP simplificado: score ponderado normalizado ──────────────────
    matriz_norm = matriz_mcda / matriz_mcda.sum(axis=0)
    score_ahp = matriz_norm @ pesos_criterios * 100
    rank_ahp = stats.rankdata(-score_ahp, method="ordinal").astype(int)

    # ── PROMETHEE II ────────────────────────────────────────────────────
    phi_plus, phi_minus, phi_net, rank_prom = promethee_ii(matriz_mcda,
                                                            pesos_criterios)

    # ── Tabla consolidada de rankings ──────────────────────────────────
    rankings_df = pd.DataFrame({
        "Dimensión": DIM,
        "W_medio": medias_W,
        "R_medio": medias_R,
        "Brecha (Imp-Prep)": [df_unit[f"G_P{i}"].mean() for i in range(1, 6)],
        "TOPSIS_Ci": C_topsis,
        "TOPSIS_Rank": rank_topsis,
        "AHP_Score": score_ahp,
        "AHP_Rank": rank_ahp,
        "PROMETHEE_phi": phi_net,
        "PROMETHEE_Rank": rank_prom,
    })

    # Ranking consensuado: media de los 3 rankings, luego se reordena
    rank_promedio = (rank_topsis + rank_ahp + rank_prom) / 3.0
    rank_consenso = stats.rankdata(rank_promedio, method="ordinal").astype(int)
    rankings_df["Rank_Consenso_Final"] = rank_consenso
    rankings_df["Rank_Promedio"] = rank_promedio

    # ── Concordancia: W de Kendall entre los 3 rankings ────────────────
    matriz_rankings = np.array([rank_topsis, rank_ahp, rank_prom],
                                dtype=float)
    W_kendall = kendall_w_rankings(matriz_rankings)

    # Etiquetas interpretativas
    if np.isnan(W_kendall):
        nivel_concordancia = "indeterminada"
    elif W_kendall >= 0.7:
        nivel_concordancia = "alta (rankings convergentes)"
    elif W_kendall >= 0.5:
        nivel_concordancia = "moderada"
    else:
        nivel_concordancia = "baja (rankings divergentes)"

    # Guardar CSVs
    matriz_df.to_csv(outdir / f"unit_{slug}_mcda_dim_matriz.csv")
    rankings_df.to_csv(outdir / f"unit_{slug}_mcda_dim_rankings.csv",
                       index=False)
    with open(outdir / f"unit_{slug}_mcda_dim_concordancia.txt", "w") as f:
        f.write(f"W de Kendall entre TOPSIS, AHP y PROMETHEE: {W_kendall:.3f}\n")
        f.write(f"Interpretación: {nivel_concordancia}\n")
        f.write("Umbrales: W≥0,7 alta | 0,5≤W<0,7 moderada | W<0,5 baja\n")

    # ── Figura panorámica ──────────────────────────────────────────────
    fig = plt.figure(figsize=(14, 9))
    fig.suptitle(f"Ranking de dimensiones (MCDA refactor v7) — {hosp}\n"
                 f"Criterios: peso (W) y percepción de preparación (R, como coste)",
                 fontsize=13, fontweight="bold")

    gs = gridspec.GridSpec(2, 2, figure=fig, height_ratios=[1.5, 1],
                            hspace=0.40, wspace=0.30,
                            left=0.05, right=0.97, top=0.88, bottom=0.06)

    # Panel A: Tabla de rankings
    axA = fig.add_subplot(gs[0, :])
    axA.axis("off")
    tabla_datos = [["Dim.", "Nombre",
                     "W (%)", "R (0-100)", "Brecha",
                     "T#", "A#", "P#", "Final"]]
    for i, dim in enumerate(DIM):
        tabla_datos.append([
            dim, DSHORT[dim],
            f"{medias_W[i]:.1f}",
            f"{medias_R[i]:.1f}",
            f"{rankings_df['Brecha (Imp-Prep)'].iloc[i]:+.1f}",
            f"#{int(rank_topsis[i])}",
            f"#{int(rank_ahp[i])}",
            f"#{int(rank_prom[i])}",
            f"#{int(rank_consenso[i])}",
        ])
    tbl = axA.table(cellText=tabla_datos[1:], colLabels=tabla_datos[0],
                     cellLoc="center", loc="center",
                     bbox=[0.05, 0.05, 0.90, 0.85])
    tbl.auto_set_font_size(False); tbl.set_fontsize(10)
    # Cabecera azul
    for col in range(len(tabla_datos[0])):
        tbl[0, col].set_facecolor("#003D7C")
        tbl[0, col].set_text_props(color="white", fontweight="bold")
    # Resaltar la fila #1 final
    for i in range(5):
        if rank_consenso[i] == 1:
            for col in range(len(tabla_datos[0])):
                tbl[i+1, col].set_facecolor("#E8F4D8")
        if rank_topsis[i] == 1:
            tbl[i+1, 5].set_facecolor("#E8F4D8")
        if rank_ahp[i] == 1:
            tbl[i+1, 6].set_facecolor("#E8F4D8")
        if rank_prom[i] == 1:
            tbl[i+1, 7].set_facecolor("#E8F4D8")
    axA.set_title("A. Rankings por método y ranking consensuado final "
                  "(verde = posición #1)",
                  fontweight="bold", fontsize=11, loc="left", pad=10)

    # Panel B: W de Kendall (gráfico de barras tipo "termómetro")
    axB = fig.add_subplot(gs[1, 0])
    valor = W_kendall if not np.isnan(W_kendall) else 0
    color_w = ("#2E9E6B" if valor >= 0.7
               else "#E8A020" if valor >= 0.5
               else "#E74C3C")
    axB.barh([""], [valor], color=color_w, edgecolor="black", linewidth=0.8,
              height=0.5)
    axB.set_xlim(0, 1)
    axB.axvline(0.5, color="orange", ls="--", lw=1, label="0,5 moderada")
    axB.axvline(0.7, color="green", ls="--", lw=1, label="0,7 alta")
    axB.set_title(f"B. Concordancia entre métodos — W de Kendall = "
                  f"{W_kendall:.3f}\n({nivel_concordancia})",
                  fontweight="bold", fontsize=10)
    axB.set_xlabel("W de Kendall")
    axB.legend(loc="lower right", fontsize=8)
    if not np.isnan(W_kendall):
        axB.text(valor / 2, 0, f"{W_kendall:.3f}",
                  ha="center", va="center", fontweight="bold",
                  color="white", fontsize=13)

    # Panel C: Ranking consensuado final (barras horizontales ordenadas)
    axC = fig.add_subplot(gs[1, 1])
    order = np.argsort(rank_consenso)   # del #1 al #5
    colores = [DCOL[i] for i in order]
    bars = axC.barh([DIM[i] for i in order],
                     [6 - rank_consenso[i] for i in order],
                     color=colores, edgecolor="black", linewidth=0.5)
    axC.invert_yaxis()
    axC.set_xlim(0, 6)
    axC.set_xlabel("← Más prioritaria (#1 a la izquierda)")
    axC.set_title("C. Ranking consensuado final",
                  fontweight="bold", fontsize=11)
    for i, idx in enumerate(order):
        axC.text(0.2, i, f"#{int(rank_consenso[idx])}  {DSHORT[DIM[idx]]}",
                  va="center", color="white", fontsize=10,
                  fontweight="bold")

    fig_path = outdir / f"unit_{slug}_mcda_dimensiones.png"
    plt.savefig(fig_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    # Dimensión #1 según consenso
    dim_top = DIM[int(np.argmin(rank_consenso))]

    return {
        "matriz_decision": matriz_df,
        "rankings": rankings_df,
        "topsis": {"Ci": C_topsis, "rank": rank_topsis},
        "ahp": {"score": score_ahp, "rank": rank_ahp},
        "promethee": {"phi": phi_net, "rank": rank_prom},
        "rank_consenso": rank_consenso,
        "rank_promedio": rank_promedio,
        "dim_top_consenso": dim_top,
        "kendall_w": W_kendall,
        "nivel_concordancia": nivel_concordancia,
        "figura": fig_path,
        # Aliases para compatibilidad con código antiguo
        "rank_avg": rank_promedio,
        "consistencia_global": W_kendall,
    }


# __NUEVA_MCDA_DIMENSIONES__


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 11.C — REGRESIÓN: VARIABLES EXPLICATIVAS DE LAS PUNTUACIONES
# ══════════════════════════════════════════════════════════════════════════════
#
# Para cada una de las 15 variables dependientes (3 preguntas × 5
# dimensiones = W_P1..W_P5, I_P1..I_P5, R_P1..R_P5) se ajustan dos
# modelos a nivel GLOBAL del estudio (todos los profesionales):
#
#   · LINEAL:    predice la puntuación 0-100 en función de las variables
#                sociodemográficas. Aporta magnitudes (coeficientes β).
#   · LOGÍSTICA: dicotomiza la puntuación por la mediana (alto / bajo)
#                y predice la probabilidad de estar en el grupo alto.
#                Aporta odds ratios.
#
# Variables independientes:
#   · edad        (continua, años)
#   · genero      (binario, 1=Masculino, 0=Femenino)
#   · rol         (categórico, one-hot encoded)
#   · experiencia (categórico ordinal: <1, 1-3, 3-5, >5 años → 1,2,3,4)
#   · formacion   (continua, horas de cursos en CP; si está disponible)
#
# Los resultados se presentan como dos heatmaps de p-valores
# (filas = variables independientes, columnas = puntuaciones)
# que permiten ver de un vistazo qué variables explican qué.
# ══════════════════════════════════════════════════════════════════════════════

def _regresion_lineal_manual(X, y, return_ci=False):
    """
    Regresión lineal múltiple con numpy + cálculo de p-valores via
    estadística t. Devuelve (coeficientes, p_valores, R²).
    Si return_ci=True devuelve también (IC95_low, IC95_high, SE).

    No depende de statsmodels: funciona en cualquier entorno con
    numpy + scipy. Implementación matemática estándar:
       β = (X'X)⁻¹ X'y
       Var(β) = σ² (X'X)⁻¹
       t = β / SE(β)
       p = 2 · (1 - cdf_t(|t|, df))
       IC95% = β ± t_crit(0.025, df) · SE(β)
    """
    n = len(y)
    X_int = np.column_stack([np.ones(n), X])
    k = X_int.shape[1]
    nan_vec = lambda: np.full(X.shape[1], np.nan)
    if n <= k:
        if return_ci:
            return (nan_vec(), nan_vec(), np.nan,
                    nan_vec(), nan_vec(), nan_vec())
        return nan_vec(), nan_vec(), np.nan
    try:
        XtX_inv = np.linalg.inv(X_int.T @ X_int)
        beta = XtX_inv @ X_int.T @ y
        y_pred = X_int @ beta
        resid = y - y_pred
        ss_res = (resid ** 2).sum()
        ss_tot = ((y - y.mean()) ** 2).sum()
        r2 = 1 - ss_res / ss_tot if ss_tot > 0 else np.nan
        df = n - k
        sigma2 = ss_res / df if df > 0 else np.nan
        se = np.sqrt(np.diag(sigma2 * XtX_inv))
        t_stat = beta / se
        p_vals = 2 * (1 - stats.t.cdf(np.abs(t_stat), df))
        if return_ci:
            t_crit = stats.t.ppf(0.975, df)
            ic_low = beta - t_crit * se
            ic_high = beta + t_crit * se
            return (beta[1:], p_vals[1:], r2,
                    ic_low[1:], ic_high[1:], se[1:])
        return beta[1:], p_vals[1:], r2
    except (np.linalg.LinAlgError, ValueError):
        if return_ci:
            return (nan_vec(), nan_vec(), np.nan,
                    nan_vec(), nan_vec(), nan_vec())
        return nan_vec(), nan_vec(), np.nan


def _regresion_logistica_manual(X, y, max_iter=100, tol=1e-6):
    """
    Regresión logística múltiple por método de Newton-Raphson (IRLS).
    Devuelve (coeficientes, p_valores, odds_ratios).

    Implementación matemática estándar:
       p = 1 / (1 + exp(-Xβ))
       β <- β + (X'WX)⁻¹ X'(y - p)        donde W = diag(p(1-p))
       Var(β) = (X'WX)⁻¹
       Wald: z = β / SE(β),  p = 2·(1-Φ(|z|))
    """
    n = len(y)
    X_int = np.column_stack([np.ones(n), X])
    k = X_int.shape[1]
    if n <= k or y.std() == 0:
        return (np.full(X.shape[1], np.nan),
                np.full(X.shape[1], np.nan),
                np.full(X.shape[1], np.nan))
    try:
        beta = np.zeros(k)
        for _ in range(max_iter):
            z = X_int @ beta
            z = np.clip(z, -30, 30)
            p_hat = 1.0 / (1.0 + np.exp(-z))
            W = p_hat * (1 - p_hat)
            grad = X_int.T @ (y - p_hat)
            H_inv = np.linalg.inv(X_int.T @ (X_int * W[:, None]) + np.eye(k) * 1e-8)
            delta = H_inv @ grad
            beta_new = beta + delta
            if np.max(np.abs(beta_new - beta)) < tol:
                beta = beta_new
                break
            beta = beta_new
        # Errores estándar e inferencia Wald
        z = X_int @ beta
        z = np.clip(z, -30, 30)
        p_hat = 1.0 / (1.0 + np.exp(-z))
        W = p_hat * (1 - p_hat)
        Cov = np.linalg.inv(X_int.T @ (X_int * W[:, None]) + np.eye(k) * 1e-8)
        se = np.sqrt(np.diag(Cov))
        z_stat = beta / se
        p_vals = 2 * (1 - stats.norm.cdf(np.abs(z_stat)))
        odds = np.exp(beta)
        return beta[1:], p_vals[1:], odds[1:]
    except (np.linalg.LinAlgError, ValueError):
        return (np.full(X.shape[1], np.nan),
                np.full(X.shape[1], np.nan),
                np.full(X.shape[1], np.nan))


def analisis_regresion(df, outdir):
    """
    Ajusta 15 modelos de regresión lineal y 15 logísticos sobre las
    puntuaciones del cuestionario, y genera dos figuras de heatmap
    + dos CSVs con los coeficientes detallados.

    Implementación independiente de statsmodels (usa numpy/scipy).
    """
    # ── Preparar variables independientes ──────────────────────────────
    df_reg = df.copy()
    df_reg["X_edad"] = pd.to_numeric(df_reg["edad"], errors="coerce")
    df_reg["X_genero"] = df_reg["genero"].map({"Masculino": 1, "Femenino": 0})
    exp_map = {"<1 año": 1, "1-3 años": 2, "3-5 años": 3, ">5 años": 4}
    df_reg["X_exp"] = df_reg["exp_label"].astype(str).map(exp_map)
    df_reg["X_rol_medico"]    = (df_reg["rol"] == "Médico").astype(int)
    df_reg["X_rol_enfermera"] = (df_reg["rol"] == "Enfermera").astype(int)

    # formación: si existe columna numérica de horas de formación, la
    # usamos. Excluimos "estudios" porque suele ser categórica/texto.
    formacion_col = None
    for cand in ["formacion", "horas_cp", "horas_formacion", "formación"]:
        if cand in df_reg.columns:
            # Verificar que sea numérica (al menos parcialmente)
            test = pd.to_numeric(df_reg[cand], errors="coerce")
            if test.notna().sum() >= 10:
                formacion_col = cand
                break
    if formacion_col is not None:
        df_reg["X_formacion"] = pd.to_numeric(df_reg[formacion_col],
                                               errors="coerce")
        variables_X = ["X_edad", "X_genero", "X_exp", "X_rol_medico",
                       "X_rol_enfermera", "X_formacion"]
        nombres_X = ["Edad", "Género (M)", "Experiencia",
                     "Rol: Médico", "Rol: Enfermera", "Formación CP"]
    else:
        variables_X = ["X_edad", "X_genero", "X_exp", "X_rol_medico",
                       "X_rol_enfermera"]
        nombres_X = ["Edad", "Género (M)", "Experiencia",
                     "Rol: Médico", "Rol: Enfermera"]
        print("     ⚠ Sin columna numérica de horas de formación en CP. "
              "Se omite esa variable.")

    # ── Variables dependientes (15) ────────────────────────────────────
    deps = []
    nombres_deps = []
    for pref, etiq in [("W", "Peso"), ("I", "Importancia"),
                        ("R", "Preparación")]:
        for i in range(1, 6):
            deps.append(f"{pref}_P{i}")
            nombres_deps.append(f"{etiq} {DIM[i-1]}")

    # ── Matrices de resultados ─────────────────────────────────────────
    pvals_lin = pd.DataFrame(np.nan, index=nombres_X,
                              columns=nombres_deps, dtype=float)
    coefs_lin = pvals_lin.copy()
    pvals_log = pvals_lin.copy()
    odds_log = pvals_lin.copy()
    r2_lin = pd.Series(np.nan, index=nombres_deps, dtype=float)

    n_validos = 0
    # ── Loop principal sobre las 15 variables dependientes ─────────────
    for dep, dep_nombre in zip(deps, nombres_deps):
        sub = df_reg[variables_X + [dep]].dropna()
        if len(sub) < 10:
            continue
        X = sub[variables_X].values.astype(float)
        y_continua = sub[dep].values.astype(float)
        n_validos = max(n_validos, len(sub))

        # Lineal
        beta, p_lin, r2 = _regresion_lineal_manual(X, y_continua)
        for idx, nombre_X in enumerate(nombres_X):
            coefs_lin.loc[nombre_X, dep_nombre] = beta[idx]
            pvals_lin.loc[nombre_X, dep_nombre] = p_lin[idx]
        r2_lin[dep_nombre] = r2

        # Logística (dicotomizar por la mediana)
        med = np.median(y_continua)
        y_bin = (y_continua > med).astype(float)
        if y_bin.std() > 0:
            beta_log, p_log, odds = _regresion_logistica_manual(X, y_bin)
            for idx, nombre_X in enumerate(nombres_X):
                odds_log.loc[nombre_X, dep_nombre] = odds[idx]
                pvals_log.loc[nombre_X, dep_nombre] = p_log[idx]

    # ── Guardar CSVs ────────────────────────────────────────────────────
    coefs_lin.to_csv(outdir / "regresion_lineal_coeficientes.csv")
    pvals_lin.to_csv(outdir / "regresion_lineal_pvalores.csv")
    odds_log.to_csv(outdir / "regresion_logistica_odds.csv")
    pvals_log.to_csv(outdir / "regresion_logistica_pvalores.csv")
    r2_lin.to_csv(outdir / "regresion_lineal_r2.csv", header=["R2"])

    # ── Figura: heatmap de p-valores ───────────────────────────────────
    fig, axs = plt.subplots(2, 1, figsize=(15, 9))
    fig.suptitle(f"Influencia de las variables sociodemográficas sobre "
                 f"las {len(deps)} puntuaciones (regresión global, "
                 f"n={n_validos})",
                 fontsize=13, fontweight="bold")

    for ax, pvals, titulo in zip(
            axs, [pvals_lin, pvals_log],
            ["A. Regresión LINEAL — −log₁₀(p)   (asteriscos = p<0,05)",
             "B. Regresión LOGÍSTICA — −log₁₀(p)   (asteriscos = p<0,05)"]):
        with np.errstate(divide="ignore", invalid="ignore"):
            data = -np.log10(pvals.values.astype(float))
            data[np.isinf(data)] = np.nan
        im = ax.imshow(data, aspect="auto", cmap="Reds", vmin=0, vmax=3.5)
        ax.set_xticks(range(len(nombres_deps)))
        ax.set_xticklabels(nombres_deps, rotation=60, ha="right",
                            fontsize=8)
        ax.set_yticks(range(len(nombres_X)))
        ax.set_yticklabels(nombres_X, fontsize=9)
        ax.set_title(titulo, fontweight="bold", fontsize=11, loc="left")

        for i in range(len(nombres_X)):
            for j in range(len(nombres_deps)):
                p = pvals.values[i, j]
                if pd.notna(p) and p < 0.001:
                    ax.text(j, i, "***", ha="center", va="center",
                            fontsize=8, color="white", fontweight="bold")
                elif pd.notna(p) and p < 0.01:
                    ax.text(j, i, "**", ha="center", va="center",
                            fontsize=8, color="white", fontweight="bold")
                elif pd.notna(p) and p < 0.05:
                    ax.text(j, i, "*", ha="center", va="center",
                            fontsize=8, color="white", fontweight="bold")

        plt.colorbar(im, ax=ax, fraction=0.02, pad=0.01)

    plt.tight_layout()
    fig_path = outdir / "regresion_heatmap.png"
    plt.savefig(fig_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"   ✓ regresion_heatmap.png")

    # ── Resumen de hallazgos significativos ────────────────────────────
    hallazgos = []
    for nombre_X in nombres_X:
        sig_lin = pvals_lin.loc[nombre_X]
        sig_lin = sig_lin[sig_lin < 0.05].sort_values()
        for dep in sig_lin.index:
            hallazgos.append({
                "modelo": "lineal", "variable": nombre_X,
                "puntuacion": dep, "p": sig_lin[dep],
                "efecto": coefs_lin.loc[nombre_X, dep],
            })
        sig_log = pvals_log.loc[nombre_X]
        sig_log = sig_log[sig_log < 0.05].sort_values()
        for dep in sig_log.index:
            hallazgos.append({
                "modelo": "logística", "variable": nombre_X,
                "puntuacion": dep, "p": sig_log[dep],
                "efecto": odds_log.loc[nombre_X, dep],
            })
    pd.DataFrame(hallazgos).to_csv(
        outdir / "regresion_hallazgos_significativos.csv", index=False)

    return {
        "figura": fig_path,
        "pvals_lin": pvals_lin, "coefs_lin": coefs_lin,
        "pvals_log": pvals_log, "odds_log": odds_log,
        "r2_lin": r2_lin,
        "hallazgos": hallazgos,
        "n_modelos": len(deps),
        "n_significativos": len(hallazgos),
        "n_observaciones": n_validos,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 11.D — REGRESIÓN MÚLTIPLE DE LA BRECHA (refactor v7)
# ══════════════════════════════════════════════════════════════════════════════
#
# Tras la revisión metodológica con el equipo médico, se añade un
# análisis de regresión lineal múltiple ESPECÍFICO de la brecha
# (importancia − percepción de preparación) para cada una de las 5 dimensiones del
# cuidado paliativo.
#
# Variable dependiente: brecha_Pi = importancia_Pi − percepción de preparación_Pi
# Predictores:           rol + género + grupo_edad + experiencia + formación CP
#
# Objetivo: distinguir entre predictores INDEPENDIENTES de la brecha
# (variables que mantienen efecto significativo tras ajustar por las
# demás) y CONFUSORES (variables que pierden significación al
# ajustar). Ejemplo típico: la experiencia acumulada correlaciona con
# haber recibido formación específica, así que sin ajuste no se sabe
# cuál de las dos explica realmente las diferencias en la brecha.
#
# Se reportan:
#   · β estandarizados (efecto comparable entre variables de
#     distintas unidades)
#   · IC95% de cada coeficiente
#   · p-valor
#   · Marca explícita de variables significativas tras ajuste
# ══════════════════════════════════════════════════════════════════════════════

def analisis_regresion_brecha(df, outdir):
    """
    Ajusta 5 modelos de regresión lineal múltiple, uno por dimensión,
    con la brecha (Imp - Prep) como variable dependiente. Devuelve
    coeficientes β estandarizados con IC95% y p-valores.
    """
    df_reg = df.copy()
    df_reg["X_edad"] = pd.to_numeric(df_reg["edad"], errors="coerce")
    df_reg["X_genero"] = df_reg["genero"].map({"Masculino": 1, "Femenino": 0})
    exp_map = {"<1 año": 1, "1-3 años": 2, "3-5 años": 3, ">5 años": 4}
    df_reg["X_exp"] = df_reg["exp_label"].astype(str).map(exp_map)
    df_reg["X_rol_medico"]    = (df_reg["rol"] == "Médico").astype(int)
    df_reg["X_rol_enfermera"] = (df_reg["rol"] == "Enfermera").astype(int)

    # Formación específica en CP como predictor ORDINAL.
    # En la encuesta es categórica (texto); se mapea a una escala 0–3.
    form_map = {
        "Ninguna": 0,
        "Hasta 20 horas": 1, "Hasta 20h": 1, "<20 horas": 1,
        "Entre 20 y 60 horas": 2, "20-60 horas": 2,
        "Mas de 60 horas": 3, "Más de 60 horas": 3, ">60 horas": 3,
    }
    formacion_disponible = False
    if "formacion" in df_reg.columns:
        f_txt = df_reg["formacion"].astype(str).str.strip()
        # 1) intento por mapeo categórico ordinal
        x_form = f_txt.map(form_map)
        # 2) si la columna fuese numérica (otra exportación), se usa tal cual
        if x_form.notna().sum() < 10:
            x_form = pd.to_numeric(df_reg["formacion"], errors="coerce")
        if x_form.notna().sum() >= 10:
            df_reg["X_formacion"] = x_form
            formacion_disponible = True

    if formacion_disponible:
        variables_X = ["X_edad", "X_genero", "X_exp",
                       "X_rol_medico", "X_rol_enfermera", "X_formacion"]
        nombres_X = ["Edad", "Género (M)", "Experiencia",
                     "Rol: Médico", "Rol: Enfermera", "Formación CP"]
    else:
        variables_X = ["X_edad", "X_genero", "X_exp",
                       "X_rol_medico", "X_rol_enfermera"]
        nombres_X = ["Edad", "Género (M)", "Experiencia",
                     "Rol: Médico", "Rol: Enfermera"]

    # Término de interacción rol × formación. Solo se incluye si la muestra
    # es suficientemente grande (datos globales); con muestras pequeñas
    # (p. ej. una sola unidad) no suele dar y se omite para evitar
    # sobreajuste o matrices singulares.
    incluir_interaccion = False
    UMBRAL_INTERACCION = 60
    if formacion_disponible and len(df_reg.dropna(subset=variables_X)) >= UMBRAL_INTERACCION:
        df_reg["X_rolMed_x_form"] = df_reg["X_rol_medico"] * df_reg["X_formacion"]
        # comprobamos que el término tenga variación
        if df_reg["X_rolMed_x_form"].nunique(dropna=True) > 1:
            variables_X = variables_X + ["X_rolMed_x_form"]
            nombres_X = nombres_X + ["Rol Médico × Formación"]
            incluir_interaccion = True

    # Resultados: para cada dimensión, una tabla de coeficientes
    resultados_por_dim = {}
    todos_los_resultados = []

    for i in range(1, 6):
        dim = f"P{i}"
        sub = df_reg[variables_X + [f"G_P{i}"]].dropna()
        if len(sub) < 10:
            resultados_por_dim[dim] = None
            continue

        X = sub[variables_X].values.astype(float)
        y = sub[f"G_P{i}"].values.astype(float)

        # Estandarizar X e y para obtener β estandarizados
        X_means = X.mean(axis=0)
        X_stds = X.std(axis=0, ddof=1)
        X_stds[X_stds == 0] = 1.0   # evitar división por cero
        X_std = (X - X_means) / X_stds

        y_mean = y.mean()
        y_std_val = y.std(ddof=1)
        if y_std_val == 0:
            resultados_por_dim[dim] = None
            continue
        y_std = (y - y_mean) / y_std_val

        beta, p_vals, r2, ic_low, ic_high, se = _regresion_lineal_manual(
            X_std, y_std, return_ci=True)

        tabla = pd.DataFrame({
            "variable": nombres_X,
            "beta_estandarizado": beta,
            "IC95_inf": ic_low,
            "IC95_sup": ic_high,
            "SE": se,
            "p_valor": p_vals,
            "significativo": p_vals < 0.05,
        })
        resultados_por_dim[dim] = {
            "tabla": tabla,
            "n": len(sub),
            "R2": r2,
            "y_mean": y_mean,
            "y_std": y_std_val,
        }

        for _, row in tabla.iterrows():
            todos_los_resultados.append({
                "dimension": dim,
                **row.to_dict()
            })

    # CSV consolidado
    todos_df = pd.DataFrame(todos_los_resultados)
    todos_df.to_csv(outdir / "regresion_brecha_coeficientes.csv", index=False)

    # ── ¿Hay asociaciones significativas? ──────────────────────────────
    n_dim = sum(1 for v in resultados_por_dim.values() if v is not None)
    sig_count = 0
    for dim, res in resultados_por_dim.items():
        if res is None:
            continue
        sig_count += int(res["tabla"]["significativo"].sum())

    # ── Interpretación en lenguaje sencillo de los predictores
    #    significativos (lo que pidió el equipo: p. ej. qué rangos de
    #    edad priorizan más unas dimensiones que otras) ────────────────
    interpretacion = []
    # Pesos medios por grupo de edad, para describir patrones de prioridad
    pesos_edad = {}
    if "grupo_edad" in df.columns:
        for i in range(1, 6):
            try:
                pesos_edad[f"P{i}"] = (df.groupby("grupo_edad", observed=True)
                                         [f"W_P{i}"].mean())
            except Exception:
                pesos_edad[f"P{i}"] = None

    for dim, res in resultados_por_dim.items():
        if res is None:
            continue
        tabla = res["tabla"]
        for _, fila in tabla[tabla["significativo"]].iterrows():
            var = fila["variable"]
            beta = fila["beta_estandarizado"]
            p = fila["p_valor"]
            signo = "mayor" if beta > 0 else "menor"
            if var == "Edad":
                frase = (f"{dim} ({DSHORT[dim]}): a mayor edad, {signo} "
                         f"brecha importancia−preparación "
                         f"(β={beta:+.2f}; p={p:.3f}).")
                serie = pesos_edad.get(dim)
                if serie is not None and serie.notna().any():
                    grupo_top = serie.idxmax()
                    frase += (f" El grupo de edad que más prioriza esta "
                              f"dimensión (mayor peso medio asignado) es "
                              f"«{grupo_top}» ({serie.max():.0f} %).")
                interpretacion.append(frase)
            elif var in ("Rol: Médico", "Rol: Enfermera"):
                rol = "los médicos" if "Médico" in var else "la enfermería"
                interpretacion.append(
                    f"{dim} ({DSHORT[dim]}): {rol} presentan una brecha "
                    f"{signo} en esta dimensión, de forma independiente "
                    f"del resto de variables (β={beta:+.2f}; p={p:.3f}).")
            elif var == "Experiencia":
                interpretacion.append(
                    f"{dim} ({DSHORT[dim]}): a mayor experiencia, {signo} "
                    f"brecha (β={beta:+.2f}; p={p:.3f}).")
            elif var.startswith("Género"):
                interpretacion.append(
                    f"{dim} ({DSHORT[dim]}): se observa una brecha {signo} "
                    f"asociada al género (β={beta:+.2f}; p={p:.3f}).")
            elif var == "Formación CP":
                interpretacion.append(
                    f"{dim} ({DSHORT[dim]}): a mayor formación específica "
                    f"en CP, {signo} brecha (β={beta:+.2f}; p={p:.3f}).")
            elif var.startswith("Rol Médico × Formación"):
                interpretacion.append(
                    f"{dim} ({DSHORT[dim]}): el efecto de la formación sobre "
                    f"la brecha DIFIERE entre médicos y enfermería (término "
                    f"de interacción significativo; β={beta:+.2f}; p={p:.3f}).")
            else:
                interpretacion.append(
                    f"{dim} ({DSHORT[dim]}): {var} se asocia a una brecha "
                    f"{signo} (β={beta:+.2f}; p={p:.3f}).")

    n_obs = max((r["n"] for r in resultados_por_dim.values()
                 if r is not None), default=0)

    # Si NO hay asociaciones significativas, el gráfico no aporta
    # información relevante: se OMITE (no se genera la figura) y el
    # informe lo indicará con una nota breve.
    if n_dim == 0 or sig_count == 0:
        return {
            "resultados": resultados_por_dim,
            "tabla_completa": todos_df,
            "figura": None,
            "n_significativos": 0,
            "n_observaciones": n_obs,
            "nombres_predictores": nombres_X,
            "interpretacion": interpretacion,
        }

    # ── Figura: forest plot de β estandarizados por dimensión ──────────
    fig, axs = plt.subplots(1, 5, figsize=(18, 8), sharey=True)
    fig.suptitle("Regresión múltiple de la BRECHA (Importancia − "
                 "Percepción de preparación) por dimensión\n"
                 "β estandarizados con IC95%",
                 fontsize=13, fontweight="bold")

    y_pos = np.arange(len(nombres_X))
    for ax, i, dim in zip(axs, range(1, 6), DIM):
        res = resultados_por_dim[dim]
        ax.set_title(f"{dim}: {DSHORT[dim]}", fontweight="bold", fontsize=10)
        ax.axvline(0, color="gray", lw=1, ls="-")
        if res is None:
            ax.text(0.5, 0.5, "Sin datos\nsuficientes",
                    ha="center", va="center",
                    transform=ax.transAxes, color="#999")
            ax.set_xlim(-1, 1)
            continue

        tabla = res["tabla"]
        for idx in range(len(nombres_X)):
            beta_val = tabla.iloc[idx]["beta_estandarizado"]
            ic_l = tabla.iloc[idx]["IC95_inf"]
            ic_u = tabla.iloc[idx]["IC95_sup"]
            sigv = tabla.iloc[idx]["significativo"]
            color = "#E74C3C" if sigv else "#888"
            ax.errorbar(beta_val, idx,
                         xerr=[[beta_val - ic_l], [ic_u - beta_val]],
                         fmt="o", color=color, ms=8, capsize=4, lw=1.5)
            if sigv:
                ax.text(beta_val, idx + 0.25, f"p={tabla.iloc[idx]['p_valor']:.3f}",
                        fontsize=7, ha="center", color="#E74C3C",
                        fontweight="bold")
        ax.set_yticks(y_pos)
        if i == 1:
            ax.set_yticklabels(nombres_X, fontsize=9)
            ax.set_ylabel("Predictor")
        ax.set_xlabel("β estandarizado")
        ax.grid(axis="x", alpha=0.3)
        # Anotación R²
        ax.text(0.02, 0.97, f"R² = {res['R2']:.3f}\nn = {res['n']}",
                 transform=ax.transAxes, fontsize=8, va="top",
                 bbox=dict(boxstyle="round", facecolor="white", alpha=0.8))

    plt.tight_layout()
    fig_path = outdir / "regresion_brecha_forest.png"
    plt.savefig(fig_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"   ✓ regresion_brecha_forest.png")

    return {
        "resultados": resultados_por_dim,
        "tabla_completa": todos_df,
        "figura": fig_path,
        "n_significativos": sig_count,
        "n_observaciones": n_obs,
        "nombres_predictores": nombres_X,
        "interpretacion": interpretacion,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 11.E — DESALINEACIÓN ENTRE ROLES (test MW de brechas)
# ══════════════════════════════════════════════════════════════════════════════
#
# A petición del equipo médico, se añade un análisis específico que
# señala explícitamente en qué dimensiones los médicos y la enfermería
# difieren significativamente en la BRECHA (no solo en percepción de preparación o
# importancia por separado), y en qué dirección se produce esa
# desalineación.
#
# Para cada dimensión Pi se realiza:
#   · Test de Mann-Whitney U sobre la brecha (Imp - Prep) entre roles
#   · Test de MW sobre la percepción de preparación (R) y la importancia (I) por
#     separado, para identificar cuál es la "fuente" de la diferencia
# ══════════════════════════════════════════════════════════════════════════════

def analisis_desalineacion_roles(df, outdir):
    """
    Analiza la desalineación entre médicos y enfermería en la BRECHA
    de cada dimensión, e identifica si la diferencia proviene del
    percepción de preparación, de la importancia o de ambas fuentes.
    """
    med = df[df["rol"] == "Médico"]
    enf = df[df["rol"] == "Enfermera"]
    if len(med) < 3 or len(enf) < 3:
        return {"resultados": [], "figura": None,
                "aviso": (f"Muestra insuficiente para comparar médicos "
                          f"(n={len(med)}) vs enfermería (n={len(enf)}).")}

    resultados = []
    for i in range(1, 6):
        dim = f"P{i}"
        g_med = med[f"G_P{i}"].dropna()
        g_enf = enf[f"G_P{i}"].dropna()
        r_med = med[f"R_P{i}"].dropna()
        r_enf = enf[f"R_P{i}"].dropna()
        im_med = med[f"I_P{i}"].dropna()
        im_enf = enf[f"I_P{i}"].dropna()

        if len(g_med) < 2 or len(g_enf) < 2:
            continue

        try:
            _, p_gap = mannwhitneyu(g_med, g_enf, alternative="two-sided")
        except ValueError:
            p_gap = np.nan
        try:
            _, p_rend = mannwhitneyu(r_med, r_enf, alternative="two-sided")
        except ValueError:
            p_rend = np.nan
        try:
            _, p_imp = mannwhitneyu(im_med, im_enf, alternative="two-sided")
        except ValueError:
            p_imp = np.nan

        # Identificar la fuente de la desalineación
        sig_rend = (not pd.isna(p_rend)) and p_rend < 0.05
        sig_imp = (not pd.isna(p_imp)) and p_imp < 0.05
        if sig_rend and sig_imp:
            fuente = "Percepción de preparación e Importancia"
        elif sig_rend:
            fuente = "Solo Percepción de preparación"
        elif sig_imp:
            fuente = "Solo Importancia"
        else:
            fuente = "Ninguno (sin desalineación clara)"

        resultados.append({
            "dimension": dim,
            "nombre": DSHORT[dim],
            "n_medicos": len(g_med),
            "n_enfermeria": len(g_enf),
            "brecha_medicos": g_med.mean(),
            "brecha_enfermeria": g_enf.mean(),
            "diferencia": g_med.mean() - g_enf.mean(),
            "p_brecha": p_gap,
            "p_preparacion": p_rend,
            "p_importancia": p_imp,
            "fuente_desalineacion": fuente,
            "significativo_brecha": (not pd.isna(p_gap)) and p_gap < 0.05,
        })

    # Guardar CSV
    df_res = pd.DataFrame(resultados)
    df_res.to_csv(outdir / "desalineacion_roles.csv", index=False)

    # ── Figura: barras de brecha por rol y dimensión + p-valores ───────
    fig, axs = plt.subplots(1, 2, figsize=(15, 6),
                              gridspec_kw={"width_ratios": [1.3, 1]})
    fig.suptitle("Desalineación entre roles: brecha Imp−Prep "
                 "(Médicos vs Enfermería)",
                 fontsize=13, fontweight="bold")

    # Panel A: barras dobles de brecha por dimensión
    axA = axs[0]
    x = np.arange(len(resultados))
    bm = [r["brecha_medicos"] for r in resultados]
    be = [r["brecha_enfermeria"] for r in resultados]
    bars1 = axA.bar(x - 0.2, bm, 0.4, color="#003D7C",
                     label="Médicos", edgecolor="black", lw=0.5)
    bars2 = axA.bar(x + 0.2, be, 0.4, color="#2E9E6B",
                     label="Enfermería", edgecolor="black", lw=0.5)
    # Marcar significancia
    for i, r in enumerate(resultados):
        if r["significativo_brecha"]:
            top = max(r["brecha_medicos"], r["brecha_enfermeria"]) + 2
            axA.text(i, top, f"*  p={r['p_brecha']:.3f}",
                      ha="center", color="red", fontweight="bold",
                      fontsize=9)
    axA.set_xticks(x)
    axA.set_xticklabels([f"{r['dimension']}\n{r['nombre']}"
                          for r in resultados], fontsize=8)
    axA.set_ylabel("Brecha media (Imp − Prep)")
    axA.set_title("A. Brecha media por rol y dimensión",
                  fontweight="bold", fontsize=11)
    axA.axhline(0, color="black", lw=0.8)
    axA.legend(loc="upper right"); axA.grid(axis="y", alpha=0.3)

    # Panel B: tabla de fuentes de desalineación
    axB = axs[1]
    axB.axis("off")
    tabla = [["Dim.", "Δbrecha", "p", "Fuente"]]
    for r in resultados:
        sig = "*" if r["significativo_brecha"] else ""
        tabla.append([
            r["dimension"],
            f"{r['diferencia']:+.1f}",
            f"{r['p_brecha']:.3f}{sig}",
            r["fuente_desalineacion"],
        ])
    tbl = axB.table(cellText=tabla[1:], colLabels=tabla[0],
                     cellLoc="center", loc="center",
                     bbox=[0.0, 0.05, 1.0, 0.85])
    tbl.auto_set_font_size(False); tbl.set_fontsize(9)
    for col in range(len(tabla[0])):
        tbl[0, col].set_facecolor("#003D7C")
        tbl[0, col].set_text_props(color="white", fontweight="bold")
    # Resaltar filas significativas
    for i, r in enumerate(resultados):
        if r["significativo_brecha"]:
            for col in range(len(tabla[0])):
                tbl[i + 1, col].set_facecolor("#FDEDEC")
    axB.set_title("B. Origen de la desalineación",
                  fontweight="bold", fontsize=11)

    plt.tight_layout()
    fig_path = outdir / "desalineacion_roles.png"
    plt.savefig(fig_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"   ✓ desalineacion_roles.png")

    n_sig = sum(1 for r in resultados if r["significativo_brecha"])
    return {
        "resultados": resultados,
        "figura": fig_path,
        "n_significativos": n_sig,
        "n_medicos": len(med),
        "n_enfermeria": len(enf),
    }


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 12 — RESUMEN EJECUTIVO (figura panorámica con 8 paneles)
# ══════════════════════════════════════════════════════════════════════════════

def fig_hospital_panorama(df, hosp_padre, outdir, tag):
    """
    Genera una figura panorámica para un hospital (que puede contener
    una o varias subunidades). Cuatro paneles:
      A. Pesos de prioridad medios del hospital
      B. Importancia vs Percepción de preparación por dimensión
      C. Brecha (Importancia − Percepción de preparación) por dimensión
      D. Perfiles de percepción de preparación por subunidad

    Se utiliza en el informe general para presentar cada hospital
    como bloque agregado antes de mostrar sus subunidades por separado.
    """
    sub = df[df["hospital_padre"] == hosp_padre].copy()
    if len(sub) == 0:
        return None

    n = len(sub)
    subunidades = sorted(sub["hospital"].dropna().unique())

    fig = plt.figure(figsize=(16, 10))
    fig.suptitle(f"Hospital {hosp_padre} — Panorámica agregada (n={n})",
                 fontsize=15, fontweight="bold", y=0.98)

    gs = gridspec.GridSpec(2, 2, hspace=0.4, wspace=0.3,
                            top=0.92, bottom=0.06,
                            left=0.06, right=0.97)

    # A. Pesos globales del hospital
    ax1 = fig.add_subplot(gs[0, 0])
    pesos = [sub[f"W_P{i}"].mean() for i in range(1, 6)]
    sems_w = [sub[f"W_P{i}"].sem() if len(sub) > 1 else 0
              for i in range(1, 6)]
    bars = ax1.bar(DIM, pesos, color=DCOL, yerr=sems_w, capsize=4,
                    edgecolor="black", linewidth=0.5)
    ax1.axhline(20, color="gray", ls="--", lw=1, label="Distribución uniforme")
    ax1.set_ylim(0, max(60, max(pesos)*1.2))
    ax1.set_title("A. Pesos de prioridad globales (%)", fontweight="bold")
    ax1.set_ylabel("Peso (%)")
    ax1.legend(fontsize=8); ax1.grid(axis="y", alpha=0.3)
    for b, v in zip(bars, pesos):
        ax1.text(b.get_x() + b.get_width()/2, b.get_height() + 1,
                 f"{v:.1f}%", ha="center", fontsize=9, fontweight="bold")

    # B. Importancia vs Percepción de preparación
    ax2 = fig.add_subplot(gs[0, 1])
    imp = [sub[f"I_P{i}"].mean() for i in range(1, 6)]
    ren = [sub[f"R_P{i}"].mean() for i in range(1, 6)]
    x = np.arange(5)
    ax2.bar(x - 0.2, imp, 0.4, color="#003D7C", label="Importancia")
    ax2.bar(x + 0.2, ren, 0.4, color="#E8A020", label="Percepción de preparación")
    ax2.set_xticks(x); ax2.set_xticklabels(DIM)
    ax2.set_ylim(0, 110)
    ax2.set_title("B. Importancia vs Percepción de preparación", fontweight="bold")
    ax2.set_ylabel("Puntuación (0-100)")
    ax2.legend(fontsize=9); ax2.grid(axis="y", alpha=0.3)

    # C. Brecha por dimensión con código semáforo
    ax3 = fig.add_subplot(gs[1, 0])
    gaps = [sub[f"G_P{i}"].mean() for i in range(1, 6)]
    colors_gap = ["#2E9E6B" if g < 5 else "#E8A020" if g < 10 else "#E74C3C"
                  for g in gaps]
    bars3 = ax3.bar(DIM, gaps, color=colors_gap, edgecolor="black", linewidth=0.5)
    ax3.set_title("C. Brecha (Importancia − Percepción de preparación)", fontweight="bold")
    ax3.set_ylabel("Puntos")
    ax3.axhline(0, color="black", lw=1)
    ax3.grid(axis="y", alpha=0.3)
    for b, v in zip(bars3, gaps):
        ax3.text(b.get_x() + b.get_width()/2, b.get_height() + 0.3,
                 f"{v:+.1f}", ha="center", fontsize=9, fontweight="bold")
    ax3.legend(handles=[
        mpatches.Patch(color="#E74C3C", label=">10 alto"),
        mpatches.Patch(color="#E8A020", label="5-10 moderado"),
        mpatches.Patch(color="#2E9E6B", label="<5 adecuado"),
    ], fontsize=8, loc="upper right")

    # D. Perfiles de percepción de preparación por subunidad
    ax4 = fig.add_subplot(gs[1, 1])
    pal_sub = make_palette(subunidades)
    for su in subunidades:
        sub2 = sub[sub["hospital"] == su]
        if len(sub2) == 0:
            continue
        m = [sub2[f"R_P{i}"].mean() for i in range(1, 6)]
        ax4.plot(DIM, m, "o-", lw=2, ms=7, color=pal_sub.get(su, "#777"),
                 label=f"{su} (n={len(sub2)})")
    ax4.set_ylim(0, 105)
    ax4.set_title("D. Percepción de preparación por subunidad", fontweight="bold")
    ax4.set_ylabel("Percepción de preparación")
    ax4.legend(fontsize=8, loc="best"); ax4.grid(alpha=0.3)

    save(fig, tag, outdir)
    return outdir / f"{tag}.png"


def fig_resumen_ejecutivo(df, outdir):
    """
    Genera la figura de resumen ejecutivo (figura 00), pensada para
    presentar los resultados globales de forma comprensible de un vistazo.
    Incluye 8 paneles:
      A. Pesos de prioridad globales (%)
      B. Importancia vs Percepción de preparación global
      C. Brecha (Imp − Prep) global con semáforo de déficit
      D. Percepción de preparación por hospital (perfiles en línea)
      E. Rankings MCDA: TOPSIS, AHP y PROMETHEE II por hospital (tabla)
      F. Brecha por rol principal (Médico vs Enfermera)
      G. Brecha por género
      H. Brecha por grupo de edad
    """
    pal_hosp   = df._pal_hosp
    hospitales = sorted(df["hospital"].dropna().unique())
    # Subset de comparables (UHD-Paliativos Adultos)
    hospitales_comp = sorted(df.loc[df["comparable"] == True,
                                    "hospital"].dropna().unique())

    # Precalcular MCDA para incluir tabla en el resumen.
    # Solo se incluyen los grupos comparables.
    groups_ok = [g for g in hospitales_comp if (df["hospital"]==g).sum() >= 1]
    mcda_data = {}
    if len(groups_ok) >= 2:
        sub  = df[df["hospital"].isin(groups_ok)]
        dm   = np.array([[sub[sub["hospital"]==g][f"R_P{i}"].mean()
                          for i in range(1,6)] for g in groups_ok])
        wm   = np.array([sub[f"W_P{i}"].mean() for i in range(1,6)])
        wn   = wm / wm.sum()
        C,_,_,tr     = topsis(dm, wn)
        ahp_df_r     = aplicar_ahp(sub, "hospital", outdir, "resumen")
        _,_,phi,pr   = promethee_ii(dm, wn)
        for j, g in enumerate(groups_ok):
            ahp_r = ahp_df_r[ahp_df_r["Grupo"]==g]["AHP_Rank"].values
            ahp_s = ahp_df_r[ahp_df_r["Grupo"]==g]["AHP_Score"].values
            mcda_data[g] = {"topsis_ci":C[j], "topsis_rank":tr[j],
                             "ahp_score":ahp_s[0] if len(ahp_s) else np.nan,
                             "ahp_rank":ahp_r[0]  if len(ahp_r) else np.nan,
                             "prom_phi":phi[j], "prom_rank":pr[j]}

    fig = plt.figure(figsize=(20, 14))
    fig.suptitle("MULTIPAL — Resumen Ejecutivo\n"
                 "Percepciones entre profesionales sobre dimensiones clave de CP domiciliarios",
                 fontsize=14, fontweight="bold")
    gs = gridspec.GridSpec(3, 3, figure=fig, hspace=.50, wspace=.38)

    # A: Pesos globales
    ax = fig.add_subplot(gs[0,0])
    wm  = [df[f"W_P{i}"].mean() for i in range(1,6)]
    ws  = [df[f"W_P{i}"].std()  for i in range(1,6)]
    bars = ax.bar(DIM, wm, color=DCOL, yerr=ws, capsize=5)
    ax.axhline(20, color="gray", ls="--", lw=1.2, label="Distribución uniforme")
    ax.set_ylim(0, 62); ax.set_title("A. Pesos de prioridad globales (%)", fontweight="bold")
    ax.set_ylabel("%"); ax.legend(fontsize=8)
    for b, v in zip(bars, wm):
        ax.text(b.get_x()+b.get_width()/2, b.get_height()+.5,
                f"{v:.1f}%", ha="center", fontsize=9.5, fontweight="bold")

    # B: Barras dobles Importancia vs Percepción de preparación global
    ax2 = fig.add_subplot(gs[0,1])
    rm = [df[f"R_P{i}"].mean() for i in range(1,6)]
    im = [df[f"I_P{i}"].mean() for i in range(1,6)]
    re = [df[f"R_P{i}"].sem()  for i in range(1,6)]
    ie = [df[f"I_P{i}"].sem()  for i in range(1,6)]
    x  = np.arange(5)
    ax2.bar(x-.2, im, .4, color="#003D7C", label="Importancia", yerr=ie, capsize=3)
    ax2.bar(x+.2, rm, .4, color="#E8A020", label="Percepción de preparación", yerr=re, capsize=3)
    ax2.set_xticks(x); ax2.set_xticklabels(DIM); ax2.set_ylim(0, 115)
    ax2.set_title("B. Importancia vs Percepción de preparación (global)", fontweight="bold")
    ax2.set_ylabel("Puntuación (0-100)"); ax2.legend(); ax2.grid(axis="y", alpha=.3)

    # C: Brecha con semáforo (rojo>10, naranja 5-10, verde<5)
    ax3 = fig.add_subplot(gs[0,2])
    gm  = [df[f"G_P{i}"].mean() for i in range(1,6)]
    gs2 = [df[f"G_P{i}"].std()  for i in range(1,6)]
    clrs = ["#E74C3C" if g>10 else "#E8A020" if g>5 else "#2E9E6B" for g in gm]
    ax3.bar(DIM, gm, color=clrs, yerr=gs2, capsize=4)
    ax3.axhline(0, color="black", lw=1.2); ax3.set_ylim(-5, 35)
    ax3.set_title("C. Brecha (Importancia − Percepción de preparación)", fontweight="bold")
    ax3.set_ylabel("Puntos"); ax3.grid(axis="y", alpha=.3)
    ax3.legend(handles=[mpatches.Patch(color="#E74C3C", label=">10 · déficit alto"),
                         mpatches.Patch(color="#E8A020", label="5-10 · moderado"),
                         mpatches.Patch(color="#2E9E6B", label="<5 · adecuado")], fontsize=8)

    # D: Perfiles de percepción de preparación por hospital (solo comparables)
    ax4 = fig.add_subplot(gs[1,0])
    for h in hospitales_comp:
        sub2 = df[df["hospital"]==h]
        ax4.plot(DIM, [sub2[f"R_P{i}"].mean() for i in range(1,6)],
                 "o-", color=pal_hosp.get(h,"#777"), lw=2.2, ms=7, label=f"{h} (n={len(sub2)})")
    ax4.set_ylim(0, 105); ax4.set_title("D. Percepción de preparación por hospital", fontweight="bold")
    ax4.set_ylabel("Percepción de preparación"); ax4.legend(fontsize=8); ax4.grid(alpha=.3)

    # E: Tabla MCDA con los tres métodos por hospital
    ax5 = fig.add_subplot(gs[1,1:3]); ax5.axis("off")
    if mcda_data:
        tabla_h = [["Hospital","n","TOPSIS Ci","T#","AHP Score","A#","PROMETHEE φ","P#","Consenso"]]
        for g in groups_ok:
            d  = mcda_data[g]; ns = (df["hospital"]==g).sum()
            ranks = [int(d["topsis_rank"]), int(d["ahp_rank"]), int(d["prom_rank"])]
            cons  = "✓ Acuerdo" if len(set(ranks))==1 else \
                    "≈ Parcial"  if max(ranks)-min(ranks)<=1 else "✗ Diverge"
            tabla_h.append([str(g), int(ns),
                             f"{d['topsis_ci']:.3f}", f"#{int(d['topsis_rank'])}",
                             f"{d['ahp_score']:.1f}", f"#{int(d['ahp_rank'])}",
                             f"{d['prom_phi']:.3f}", f"#{int(d['prom_rank'])}", cons])
        tbl = ax5.table(cellText=tabla_h[1:], colLabels=tabla_h[0],
                         cellLoc="center", loc="center", bbox=[0,.15,1,.75])
        tbl.auto_set_font_size(False); tbl.set_fontsize(10.5)
        tbl.auto_set_column_width(list(range(len(tabla_h[0]))))
        for i, row in enumerate(tabla_h[1:]):
            col = "#D5F5E3" if "Acuerdo" in row[-1] else \
                  "#FEF9E7" if "Parcial" in row[-1] else "#FADBD8"
            for j in range(len(tabla_h[0])): tbl[i+1,j].set_facecolor(col)
    ax5.set_title("E. Rankings MCDA: TOPSIS · AHP · PROMETHEE II (por hospital)\n"
                  "Verde=consenso · Amarillo=parcial · Rojo=divergencia",
                  fontweight="bold", fontsize=10, pad=12)

    # F, G, H: Brechas por rol, género y edad
    ax6 = fig.add_subplot(gs[2,0])
    for rol, col in [("Médico","#003D7C"),("Enfermera","#2E9E6B")]:
        sub3 = df[df["rol"]==rol]
        if len(sub3) < 3: continue
        ax6.plot(DIM, [sub3[f"G_P{i}"].mean() for i in range(1,6)],
                 "o-", color=col, lw=2, ms=7, label=f"{rol} (n={len(sub3)})")
    ax6.axhline(0, color="black", lw=1.2)
    ax6.set_title("F. Brecha por rol principal", fontweight="bold")
    ax6.set_ylabel("Brecha (puntos)"); ax6.legend(fontsize=9); ax6.grid(alpha=.3)

    ax7 = fig.add_subplot(gs[2,1])
    for gen, col in [("Femenino","#E91E8C"),("Masculino","#003D7C")]:
        sg = df[df["genero"]==gen]
        if len(sg) < 5: continue
        ax7.plot(DIM, [sg[f"G_P{i}"].mean() for i in range(1,6)],
                 "o-", color=col, lw=2, ms=7, label=f"{gen} (n={len(sg)})")
    ax7.axhline(0, color="black", lw=1.2)
    ax7.set_title("G. Brecha por género", fontweight="bold")
    ax7.set_ylabel("Brecha (puntos)"); ax7.legend(fontsize=9); ax7.grid(alpha=.3)

    ax8 = fig.add_subplot(gs[2,2])
    for ag, col in PAL_AGE.items():
        sa = df[df["grupo_edad"]==ag]
        if len(sa) < 3: continue
        ax8.plot(DIM, [sa[f"G_P{i}"].mean() for i in range(1,6)],
                 "o-", color=col, lw=2, ms=7, label=f"{ag} (n={len(sa)})")
    ax8.axhline(0, color="black", lw=1.2)
    ax8.set_title("H. Brecha por grupo de edad", fontweight="bold")
    ax8.set_ylabel("Brecha (puntos)"); ax8.legend(fontsize=8); ax8.grid(alpha=.3)

    plt.tight_layout(); save(fig, "00_resumen_ejecutivo", outdir)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 13 — INFORME DE TEXTO
# ══════════════════════════════════════════════════════════════════════════════

def informe_texto(df, outdir):
    """
    Genera un informe de texto plano (MULTIPAL_informe_resumen.txt) con:
      · Descripción de la muestra (n por hospital y unidad)
      · Pesos de prioridad medios ± DE
      · Percepción de preparación media ± DE por dimensión
      · Brechas con clasificación automática (DÉFICIT ALTO / moderado / adecuado)
      · W de Kendall global y por subgrupo
      · Rankings MCDA (TOPSIS, AHP, PROMETHEE) por hospital con interpretación
    También imprime el informe en la consola para seguimiento inmediato.
    """
    def _w(sub):
        mat = sub[[f"W_P{i}" for i in range(1,6)]].dropna().values
        if len(mat) < 2: return np.nan
        ranks = np.apply_along_axis(lambda x: len(x)+1-stats.rankdata(x,"average"), 1, mat)
        return kendall_w(ranks)

    lines = ["="*65, "MULTIPAL v2.0 — INFORME RESUMEN",
             f"Generado: {datetime.now():%d/%m/%Y %H:%M}", "="*65, "",
             f"MUESTRA: {len(df)} profesionales"]
    for h, n in df["hospital"].value_counts().items():
        lines.append(f"  Hospital {h}: n={n}")
        for u, nu in df[df["hospital"]==h]["unidad"].value_counts().items():
            lines.append(f"    · {u}: n={nu}" + (" ⚠" if nu<5 else ""))

    lines += ["", "PESOS DE PRIORIDAD (media ± DE):"]
    for i in range(1,6):
        lines.append(f"  P{i} {DFULL[f'P{i}']}: {df[f'W_P{i}'].mean():.1f}% ±{df[f'W_P{i}'].std():.1f}")

    lines += ["", "RENDIMIENTO (media ± DE):"]
    for i in range(1,6):
        lines.append(f"  P{i}: {df[f'R_P{i}'].mean():.1f} ±{df[f'R_P{i}'].std():.1f}")

    lines += ["", "BRECHAS (Importancia − Percepción de preparación):"]
    for i in range(1,6):
        g   = df[f"G_P{i}"].mean()
        est = "DÉFICIT ALTO" if g>10 else "moderado" if g>5 else "adecuado"
        lines.append(f"  P{i}: {g:+.1f} pts  [{est}]")

    lines += ["", "CONSENSO W de KENDALL:"]
    lines.append(f"  Global: {_w(df):.3f}")
    for gc in ["hospital","rol"]:
        for g, sub in df.groupby(gc, observed=True):
            w = _w(sub)
            if not pd.isna(w):
                lines.append(f"  {gc}={g}: W={w:.3f} (n={len(sub)})")

    # Rankings MCDA por hospital (solo comparables UHD-Paliativos Adultos)
    hospitales_comp = sorted(df.loc[df["comparable"] == True,
                                    "hospital"].dropna().unique())
    groups_ok = [g for g in hospitales_comp if (df["hospital"]==g).sum() >= 1]
    if len(groups_ok) >= 2:
        sub  = df[df["hospital"].isin(groups_ok)]
        dm   = np.array([[sub[sub["hospital"]==g][f"R_P{i}"].mean()
                          for i in range(1,6)] for g in groups_ok])
        wm   = np.array([sub[f"W_P{i}"].mean() for i in range(1,6)]); wn = wm/wm.sum()
        C,_,_,tr     = topsis(dm, wn)
        ahp_df_r     = aplicar_ahp(sub, "hospital", outdir, "informe")
        _,_,phi,pr   = promethee_ii(dm, wn)
        lines += ["", "RANKINGS MCDA POR HOSPITAL:"]
        lines.append(f"  {'Hospital':30s} {'TOPSIS Ci':>12} {'#T':>4} {'AHP Score':>12} {'#A':>4} {'PROM. φ':>12} {'#P':>4}")
        lines.append("  "+"-"*80)
        for j, g in enumerate(groups_ok):
            ar  = ahp_df_r[ahp_df_r["Grupo"]==g]["AHP_Rank"].values[0]
            as_ = ahp_df_r[ahp_df_r["Grupo"]==g]["AHP_Score"].values[0]
            lines.append(f"  {str(g):30s} {C[j]:>12.3f} {tr[j]:>4} {as_:>12.2f} {int(ar):>4} {phi[j]:>12.3f} {pr[j]:>4}")
        lines += ["", "INTERPRETACIÓN:"]
        lines += ["  TOPSIS Ci: 0=peor posible · >0.5=cerca del ideal · 1=ideal perfecto"]
        lines += ["  AHP Score: percepción de preparación global ponderada (0-100, mayor=mejor)"]
        lines += ["  PROMETHEE φ: >0=supera a más alternativas · <0=es superado"]

    lines += ["", "="*65, "Nota: grupos con n<5 son orientativos.", "="*65]
    txt  = "\n".join(str(l) for l in lines)
    (outdir/"MULTIPAL_informe_resumen.txt").write_text(txt, encoding="utf-8")
    print("     ✓ MULTIPAL_informe_resumen.txt")
    print(f"\n{txt}\n")


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 14 — DESCRIPTIVOS GENERALES DE LA MUESTRA
# ══════════════════════════════════════════════════════════════════════════════

def fig_descriptivos(df, outdir):
    """
    Figura 01: caracterización sociodemográfica y profesional de la muestra.
    6 paneles:
      · Distribución por rol profesional (barras)
      · Distribución por unidad hospitalaria (barras horizontales)
      · Histograma de edad con cortes en 35, 45, 55 años
      · Gráfico de tarta por género
      · Barras por grupo de edad
      · Pesos de prioridad medios globales (%)
    """
    fig, axs = plt.subplots(2, 3, figsize=(17, 9))
    fig.suptitle("MULTIPAL — Caracterización de la muestra", fontsize=14, fontweight="bold")

    rc = df["rol"].value_counts()
    axs[0,0].bar(rc.index, rc.values, color=[PAL_ROL.get(r,"#999") for r in rc.index])
    axs[0,0].set_title("Rol profesional"); axs[0,0].set_ylabel("N")
    axs[0,0].tick_params(axis="x", rotation=25)
    for b, v in zip(axs[0,0].patches, rc.values):
        axs[0,0].text(b.get_x()+b.get_width()/2, b.get_height()+.15, str(v), ha="center")

    uc = df["unidad"].value_counts()
    axs[0,1].barh(uc.index, uc.values, color=[df._pal_unit.get(u,"#777") for u in uc.index])
    axs[0,1].set_title("Unidad"); axs[0,1].set_xlabel("N")
    for b, v in zip(axs[0,1].patches, uc.values):
        axs[0,1].text(b.get_width()+.1, b.get_y()+b.get_height()/2, str(v), va="center")

    axs[0,2].hist(df["edad"].dropna(), bins=12, color="#E8A020", edgecolor="white")
    for cut in [35, 45, 55]: axs[0,2].axvline(cut, color="#003D7C", ls="--", lw=1.2)
    axs[0,2].set_title("Distribución de edad"); axs[0,2].set_xlabel("Edad"); axs[0,2].set_ylabel("Frecuencia")

    gc = df["genero"].value_counts(dropna=True)
    axs[1,0].pie(gc.values, labels=gc.index, autopct="%1.0f%%",
                 colors=[PAL_GEN.get(g,"#999") for g in gc.index], startangle=90)
    axs[1,0].set_title("Género")

    gc2 = df["grupo_edad"].value_counts().sort_index()
    axs[1,1].bar(gc2.index.astype(str), gc2.values,
                 color=[PAL_AGE.get(g,"#999") for g in gc2.index.astype(str)])
    axs[1,1].set_title("Grupo de edad"); axs[1,1].set_ylabel("N")
    for b, v in zip(axs[1,1].patches, gc2.values):
        axs[1,1].text(b.get_x()+b.get_width()/2, b.get_height()+.1, str(v), ha="center")

    wm   = [df[f"W_P{i}"].mean() for i in range(1,6)]
    bars = axs[1,2].bar(DIM, wm, color=DCOL)
    axs[1,2].axhline(20, color="gray", ls="--", lw=1.2, label="Uniforme")
    axs[1,2].set_ylim(0, 60); axs[1,2].set_title("Pesos de prioridad medios (%)")
    axs[1,2].set_ylabel("%"); axs[1,2].legend(fontsize=8)
    for b, v in zip(bars, wm):
        axs[1,2].text(b.get_x()+b.get_width()/2, b.get_height()+.5,
                      f"{v:.1f}%", ha="center", fontsize=9)
    plt.tight_layout(); save(fig, "01_descriptivos", outdir)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 15 — INFORME PDF PROFESIONAL CON TODAS LAS GRÁFICAS Y COMENTARIOS
# ══════════════════════════════════════════════════════════════════════════════
#
# Genera un documento PDF multipágina que recopila todas las figuras producidas
# por el pipeline de análisis y las acompaña de comentarios redactados en
# estilo formal apto para presentación institucional o académica.
#
# El informe se estructura en los siguientes apartados:
#   - Portada con metadatos de la muestra
#   - Índice de contenidos
#   - Resumen ejecutivo
#   - Caracterización de la muestra
#   - Análisis comparativo entre hospitales
#   - Análisis por subunidades (un bloque por hospital)
#   - Análisis por rol profesional
#   - Análisis por género
#   - Análisis por grupo de edad
#   - Análisis por experiencia profesional
#   - Consenso interprofesional (W de Kendall)
#   - Análisis multicriterio (TOPSIS, AHP, PROMETHEE II)
#   - Conclusiones e implicaciones
# ══════════════════════════════════════════════════════════════════════════════


# Comentarios profesionales asociados a cada figura del informe.
# Las claves coinciden con el nombre de archivo PNG (sin extensión).
COMENTARIOS_FIGURAS = {
    "00_resumen_ejecutivo": (
        "La figura panorámica resume los hallazgos más relevantes del estudio "
        "en ocho paneles complementarios. El panel A muestra los pesos de "
        "prioridad medios asignados por los profesionales a cada una de las "
        "cinco dimensiones evaluadas, contrastándolos con la línea de "
        "distribución uniforme (20 %). El panel B contrapone la importancia "
        "percibida con el percepción de preparación autoinformada para cada dimensión, "
        "permitiendo identificar visualmente los desajustes. El panel C "
        "cuantifica esos desajustes mediante la brecha (Importancia − "
        "Percepción de preparación), aplicando un código semáforo que diferencia los "
        "déficits altos (>10 puntos), moderados (5-10) y áreas adecuadas "
        "(<5). El panel D representa los perfiles de percepción de preparación por "
        "hospital. La tabla del panel E sintetiza los rankings derivados de "
        "los tres métodos multicriterio (TOPSIS, AHP y PROMETHEE II), "
        "señalando con código de color el grado de consenso entre métodos. "
        "Los paneles F, G y H descomponen la brecha por rol profesional, "
        "género y grupo de edad, lo que facilita la detección de "
        "perspectivas diferenciadas entre subgrupos."
    ),
    "01_descriptivos": (
        "La caracterización de la muestra se presenta en seis paneles que "
        "describen la composición sociodemográfica y profesional de los "
        "participantes. Se muestran las distribuciones por rol profesional, "
        "unidad de adscripción, edad (con cortes en 35, 45 y 55 años), "
        "género y grupo de edad agregado. El último panel anticipa la "
        "distribución global de los pesos de prioridad asignados a las cinco "
        "dimensiones, permitiendo apreciar de manera preliminar las "
        "preferencias colectivas del conjunto de la muestra."
    ),
    "02_radar_hospitales": (
        "Los gráficos radar comparan, para cada hospital, los perfiles de "
        "importancia percibida (en azul) y percepción de preparación autoinformada (en "
        "naranja) a lo largo de las cinco dimensiones. La separación entre "
        "ambas curvas constituye una representación visual directa de la "
        "brecha entre lo que se considera relevante y lo que efectivamente "
        "se está consiguiendo. Los grupos con un tamaño muestral inferior a "
        "5 se etiquetan como orientativos, ya que sus medias presentan una "
        "elevada variabilidad."
    ),
    "03_percepción de preparación_hospital": (
        "Comparativa del percepción de preparación por dimensión entre "
        "hospitales. Las barras representan las medias muestrales y los "
        "intervalos verticales el error estándar de la media (SEM), "
        "permitiendo valorar la precisión de cada estimación. Las "
        "diferencias entre hospitales reflejan la heterogeneidad operativa "
        "de las unidades y constituyen el punto de partida para los tests "
        "no paramétricos por pares (Mann-Whitney U) que figuran en los CSV "
        "asociados."
    ),
    "04_importancia_hospital": (
        "Comparativa de la importancia atribuida a cada dimensión por los "
        "profesionales de cada hospital. Una elevada coincidencia entre "
        "centros indica un consenso profesional sobre las prioridades del "
        "cuidado paliativo domiciliario; las divergencias, en cambio, "
        "pueden reflejar diferencias contextuales o de modelo asistencial."
    ),
    "05_pesos_hospital": (
        "Distribución de los pesos de prioridad (en porcentaje) que cada "
        "hospital concede a las dimensiones del cuestionario. A diferencia "
        "de la importancia (que puede valorarse independientemente para "
        "cada dimensión), los pesos exigen un reparto que suma 100 %, lo "
        "que obliga a establecer un orden de relevancia. Este reparto "
        "alimenta los algoritmos MCDA presentados en la sección final."
    ),
    "06_gap_hospital": (
        "Mapa de calor de la brecha (Importancia − Percepción de preparación) por "
        "hospital y dimensión. Los tonos rojos identifican déficits "
        "percibidos por los profesionales (la dimensión se considera muy "
        "importante pero el percepción de preparación autoinformada es insuficiente); "
        "los tonos verdes señalan equilibrio o áreas en que la percepción de preparación "
        "iguala o supera la importancia atribuida. Este gráfico permite "
        "priorizar las áreas de mejora con un criterio explícito y "
        "cuantitativo."
    ),
    "07_comparativa_hospitales": (
        "Visualización conjunta de los tres indicadores principales "
        "—percepción de preparación, importancia y pesos— para cada hospital. La "
        "yuxtaposición de los tres paneles facilita la interpretación "
        "integrada: una dimensión bien valorada (alta importancia y peso) "
        "pero con baja percepción de preparación es candidata prioritaria a intervención, "
        "mientras que una dimensión bien resuelta (alto percepción de preparación) en un "
        "área de baja prioridad sugiere una asignación eficiente de "
        "recursos."
    ),
    "19_radar_rol": (
        "Comparación por rol profesional de los perfiles de importancia y "
        "percepción de preparación. Las diferencias entre roles permiten identificar "
        "lecturas profesionales distintas sobre el mismo proceso "
        "asistencial: por ejemplo, una mayor sensibilidad de un colectivo "
        "ante determinadas dimensiones psicosociales o de coordinación "
        "puede reflejar su posición específica en el equipo "
        "multidisciplinar."
    ),
    "19_radar_brecha_rol": (
        "Gráfico de araña de la BRECHA (Importancia − Percepción de "
        "preparación) por dimensión, con un polígono por rol profesional. "
        "Cada vértice es la brecha media de una dimensión: cuanto más "
        "lejos del centro (línea discontinua = brecha 0), mayor es el "
        "déficit percibido (la importancia supera a la preparación). "
        "Permite ver de un vistazo en qué dimensiones médicos y enfermería "
        "perciben mayor desajuste y si sus perfiles de déficit coinciden o "
        "divergen."
    ),
    "20_preparacion_rol": (
        "Comparativa de la percepción de preparación entre roles profesionales. "
        "Las diferencias observadas, contrastadas estadísticamente "
        "mediante Mann-Whitney U y Kruskal-Wallis (resultados en CSV "
        "anexos), permiten identificar discrepancias significativas entre "
        "colectivos. Una divergencia marcada entre médicos y personal de "
        "enfermería, por ejemplo, puede tener implicaciones para la "
        "coordinación del equipo."
    ),
    "21_importancia_rol": (
        "Importancia atribuida a cada dimensión por rol profesional. "
        "Permite evaluar el grado de alineamiento entre colectivos sobre "
        "lo que constituye una atención paliativa de calidad."
    ),
    "22_pesos_rol": (
        "Reparto de pesos de prioridad por rol profesional. Las "
        "diferencias en este reparto reflejan jerarquías de valor "
        "asistencial específicas de cada perfil, y son insumo directo "
        "para los análisis MCDA por rol."
    ),
    "23_gap_rol": (
        "Heatmap de la brecha por rol profesional. La identificación "
        "específica de los déficits percibidos por cada colectivo "
        "facilita el diseño de intervenciones formativas o "
        "organizativas dirigidas a perfiles concretos."
    ),
    "25_consenso_kendall": (
        "Coeficiente de concordancia W de Kendall, que cuantifica el "
        "grado de acuerdo entre los profesionales en el reparto de "
        "pesos de prioridad. Un valor próximo a 1 indica acuerdo casi "
        "unánime; un valor próximo a 0 refleja heterogeneidad en las "
        "preferencias. Las líneas de referencia marcan los umbrales "
        "habituales de interpretación: W = 0,5 (acuerdo moderado) y "
        "W = 0,7 (acuerdo alto). Los paneles muestran el consenso "
        "descompuesto por hospital, por rol profesional y por grupo de "
        "edad."
    ),
    "24_experiencia_boxplots": (
        "Boxplots de percepción de preparación e importancia por grupo de experiencia "
        "profesional. La línea central representa la mediana, la caja el "
        "rango intercuartílico y los bigotes la dispersión hasta 1,5 "
        "veces ese rango. Esta representación permite detectar tendencias "
        "asociadas a la antigüedad profesional, complementadas con la "
        "correlación de Spearman que aparece en los CSV adjuntos."
    ),
    "14_analisis_genero": (
        "Análisis comparativo por género estructurado en seis paneles. "
        "La fila superior muestra las medias de percepción de preparación, importancia "
        "y pesos de prioridad por dimensión para profesionales mujeres "
        "y hombres, con barras de error correspondientes al SEM. Las "
        "diferencias estadísticamente significativas detectadas mediante "
        "Mann-Whitney U se señalan con asteriscos sobre las barras "
        "(*** p<0,001; ** p<0,01; * p<0,05). La fila inferior muestra "
        "la brecha media por género, la distribución global del "
        "percepción de preparación mediante boxplot, y una tabla resumen con los "
        "resultados de las pruebas estadísticas."
    ),
    "15_edad_boxplots": (
        "Distribución del percepción de preparación por grupo de edad "
        "mediante boxplots para cada dimensión. Permite identificar "
        "diferencias generacionales en la valoración del propio "
        "percepción de preparación. Las medianas y los rangos intercuartílicos "
        "facilitan una comparación robusta entre grupos sin asumir "
        "normalidad."
    ),
    "16_heatmap_edad": (
        "Heatmap de la brecha (Importancia − Percepción de preparación) por grupo "
        "de edad y dimensión. Las divergencias entre grupos pueden "
        "reflejar diferentes niveles de exigencia personal o "
        "perspectivas distintas sobre la atención paliativa asociadas "
        "a la trayectoria profesional acumulada."
    ),
    "17_percepción de preparación_edad": (
        "Comparativa del percepción de preparación por grupo de edad. "
        "Las variaciones observadas entre cohortes pueden estar "
        "moduladas por la experiencia acumulada, el nivel formativo y "
        "los cambios generacionales en el modelo asistencial."
    ),
    "18_importancia_edad": (
        "Importancia atribuida a cada dimensión por grupo de edad. "
        "El grado de alineamiento entre cohortes constituye un "
        "indicador de la solidez del marco conceptual compartido sobre "
        "los cuidados paliativos en la organización."
    ),
}


def _comentario_figura(nombre_fig, hospital=None):
    """
    Devuelve el comentario profesional asociado a una figura.

    Para figuras dinámicas (subunidades por hospital, MCDA por grupo,
    etc.) no contempladas explícitamente, se genera un comentario por
    patrón a partir del nombre del archivo.
    """
    # Coincidencia directa por nombre exacto
    if nombre_fig in COMENTARIOS_FIGURAS:
        return COMENTARIOS_FIGURAS[nombre_fig]

    # Subunidades por hospital (patrón 08_<hospital>_<sufijo>)
    if "_a_radar" in nombre_fig:
        return (f"Perfiles radar de importancia y percepción de preparación para las "
                f"subunidades del hospital {hospital or ''}. Las áreas con "
                f"mayor distancia entre las curvas azul (importancia) y "
                f"naranja (percepción de preparación) identifican los déficits percibidos "
                f"con mayor intensidad por los profesionales de cada unidad.")
    if "_b_preparacion" in nombre_fig:
        return (f"Comparativa del percepción de preparación entre las "
                f"subunidades de {hospital or 'este hospital'}. Las barras "
                f"de error representan el SEM; las unidades con menos de "
                f"5 respuestas se identifican como orientativas.")
    if "_c_importancia" in nombre_fig:
        return (f"Importancia atribuida a cada dimensión por subunidad "
                f"dentro de {hospital or 'el hospital'}. El alineamiento "
                f"entre unidades sobre lo que se considera prioritario es "
                f"un indicador de coherencia organizativa interna.")
    if "_d_pesos" in nombre_fig:
        return (f"Reparto de pesos de prioridad por subunidad. Refleja la "
                f"adaptación de los criterios de relevancia al perfil de "
                f"pacientes y al modelo de cuidados específico de cada "
                f"unidad de {hospital or 'el hospital'}.")
    if "_e_gap" in nombre_fig:
        return (f"Heatmap de brecha por subunidad. Los tonos rojos "
                f"señalan áreas de mejora prioritaria; los tonos verdes "
                f"indican equilibrio. La comparación entre subunidades "
                f"de un mismo centro permite identificar buenas prácticas "
                f"potencialmente transferibles.")
    if "_f_perfiles" in nombre_fig:
        return (f"Perfiles consolidados de percepción de preparación, importancia y "
                f"brecha por subunidad de {hospital or 'el hospital'}. La "
                f"visión integrada facilita la priorización de "
                f"intervenciones internas.")

    # MCDA
    if nombre_fig.endswith("_mcda"):
        return ("Figura principal del análisis multicriterio que integra "
                "los tres métodos aplicados (TOPSIS, AHP y PROMETHEE II) en "
                "ocho paneles. Las barras horizontales muestran el "
                "percepción de preparación relativo de cada grupo según cada método, "
                "junto con las distancias al ideal positivo y negativo "
                "(TOPSIS), los flujos de preferencia (PROMETHEE) y los "
                "pesos por dimensión derivados (AHP). La convergencia de "
                "los tres rankings refuerza la robustez de las "
                "conclusiones.")
    if nombre_fig.endswith("_mcda_interpretacion"):
        return ("Visualización interpretativa de los resultados MCDA. Cada "
                "panel etiqueta gráficamente la lectura de cada método: "
                "TOPSIS clasifica como verde (Ci ≥ 0,5, cercano al ideal) "
                "o rojo (Ci < 0,5, cercano al anti-ideal); AHP ordena los "
                "scores ponderados de mayor a menor; PROMETHEE II señala "
                "con verde los grupos con flujo neto positivo (superan a "
                "más alternativas que las que los superan) y con rojo los "
                "de flujo negativo.")
    if nombre_fig.endswith("_ahp_consistencia"):
        return ("Diagnóstico de consistencia del análisis AHP. La tabla "
                "incluye el autovalor principal (λmax), el índice de "
                "consistencia (CI) y el ratio de consistencia (CR). Según "
                "el criterio de Saaty (1980), un CR inferior a 0,10 "
                "indica que los juicios emitidos son razonablemente "
                "coherentes; valores superiores sugieren revisar las "
                "ponderaciones para asegurar la validez del ranking.")

    # Comentario genérico cuando no hay patrón identificado
    return ("Esta figura forma parte del bloque analítico correspondiente y "
            "complementa las visualizaciones principales con detalle "
            "adicional sobre la dimensión examinada.")


def _titulo_legible_figura(fig_name, hospitales=None):
    """
    Convierte el nombre del archivo de figura en un título profesional
    apto para presentación. Reemplaza guiones bajos, expande siglas
    (MCDA, AHP, etc.), resuelve los tags abreviados de hospital y
    elimina los prefijos numéricos internos.

    Parameters
    ----------
    fig_name : str
        Nombre del archivo PNG sin extensión.
    hospitales : list[str] | None
        Lista de nombres completos de hospitales detectados, usada
        para resolver tags abreviados como "Hospital_La_" → "Hospital La Fe".
    """
    stem = fig_name

    # 1. Eliminar el prefijo numérico de orden interno
    #    (e.g. "08_Hospital_La__a_radar" → "Hospital_La__a_radar")
    stem = re.sub(r"^\d+_", "", stem)

    # 2. Resolver el tag abreviado del hospital por su nombre completo
    #    Los tags son hosp[:12].replace(" ", "_"), por lo que pueden
    #    terminar en guion bajo si el nombre tiene exactamente 12 chars
    #    con espacios. Probamos los tags de mayor longitud primero para
    #    evitar coincidencias parciales.
    if hospitales:
        tags_ordenados = sorted(
            [(h[:12].replace(" ", "_"), h) for h in hospitales],
            key=lambda x: -len(x[0]))
        for tag, nombre in tags_ordenados:
            if tag in stem:
                stem = stem.replace(tag, nombre.replace(" ", "_"))
                break

    # 3. Sustituciones de siglas y sufijos técnicos
    sustituciones = {
        "_a_radar":      " — radar de importancia y percepción de preparación",
        "radar_brecha_rol": "Radar de brecha (Importancia − Preparación) por rol",
        "_b_preparacion":" — comparativa de percepción de preparación",
        "_c_importancia":" — comparativa de importancia",
        "_d_pesos":      " — pesos de prioridad",
        "_e_gap":        " — brecha (Importancia − Percepción de preparación)",
        "_f_perfiles":   " — perfiles consolidados",
        "_mcda_interpretacion": " — interpretación MCDA",
        "_ahp_consistencia":    " — consistencia AHP",
        "_mcda":         " — análisis MCDA (TOPSIS · AHP · PROMETHEE II)",
        "_subunidades":  " — subunidades",
    }
    for orig, nuevo in sustituciones.items():
        stem = stem.replace(orig, nuevo)

    # 4. Limpieza general: guiones bajos por espacios, espacios múltiples
    stem = stem.replace("_", " ")
    stem = re.sub(r"\s+", " ", stem).strip()

    # 5. Capitalizar solo la primera letra
    if stem:
        stem = stem[0].upper() + stem[1:]
    return stem


def _añadir_pagina_imagen(pdf, ruta_png, titulo_pagina, comentario,
                          numero_pagina=None, total_paginas=None):
    """
    Añade al PDF una página A4 horizontal compuesta por un encabezado con
    título, la imagen PNG centrada, y un bloque de comentario al pie.
    """
    if not Path(ruta_png).exists():
        return False

    fig = plt.figure(figsize=(11.69, 8.27))  # A4 apaisado en pulgadas
    fig.patch.set_facecolor("white")

    # Encabezado superior con franja institucional
    ax_head = fig.add_axes([0, 0.94, 1, 0.06])
    ax_head.set_facecolor("#003D7C")
    ax_head.text(0.02, 0.5, "MULTIPAL  ·  Percepciones entre profesionales "
                            "sobre los cuidados paliativos domiciliarios",
                 va="center", ha="left", color="white",
                 fontsize=9.5, fontweight="bold",
                 transform=ax_head.transAxes)
    if numero_pagina is not None and total_paginas is not None:
        ax_head.text(0.98, 0.5, f"Página {numero_pagina} de {total_paginas}",
                     va="center", ha="right", color="white", fontsize=8.5,
                     transform=ax_head.transAxes)
    ax_head.set_xticks([]); ax_head.set_yticks([])
    for s in ax_head.spines.values(): s.set_visible(False)

    # Título de la figura
    ax_tit = fig.add_axes([0.05, 0.87, 0.9, 0.05])
    ax_tit.text(0, 0.5, titulo_pagina, va="center", ha="left",
                fontsize=13, fontweight="bold", color="#003D7C",
                transform=ax_tit.transAxes)
    ax_tit.axis("off")

    # Imagen
    ax_img = fig.add_axes([0.05, 0.28, 0.9, 0.58])
    try:
        img = imread(str(ruta_png))
        ax_img.imshow(img)
    except Exception as e:
        ax_img.text(0.5, 0.5, f"[No se pudo cargar la figura: {e}]",
                    ha="center", va="center", fontsize=10, color="gray",
                    transform=ax_img.transAxes)
    ax_img.axis("off")

    # Bloque de comentario al pie
    ax_com = fig.add_axes([0.05, 0.04, 0.9, 0.22])
    ax_com.axis("off")
    ax_com.text(0, 0.95, "Lectura interpretativa",
                va="top", ha="left", fontsize=10, fontweight="bold",
                color="#003D7C", transform=ax_com.transAxes)
    comentario_envuelto = textwrap.fill(comentario, width=145)
    ax_com.text(0, 0.78, comentario_envuelto,
                va="top", ha="left", fontsize=9, color="#222",
                linespacing=1.45, transform=ax_com.transAxes)

    # Línea separadora superior del comentario
    ax_com.axhline(0.88, color="#003D7C", lw=0.7,
                   xmin=0, xmax=1)

    # Pie de página
    ax_foot = fig.add_axes([0, 0, 1, 0.025])
    ax_foot.axis("off")
    ax_foot.text(0.5, 0.5,
                 f"Informe generado el {datetime.now():%d/%m/%Y a las %H:%M}",
                 ha="center", va="center", fontsize=7.5, color="gray",
                 transform=ax_foot.transAxes)

    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)
    return True


def _pagina_texto(pdf, titulo, contenido, numero_pagina=None,
                  total_paginas=None, es_portada=False, subtitulo=None):
    """
    Añade una página de texto al PDF (portada, índice, separadores de
    sección o conclusiones).

    subtitulo: texto del encabezado superior. Si es None se usa el del
    estudio de cuidados paliativos domiciliarios; las unidades que no son
    UHD (oncología, hematología, etc.) pasan aquí su propio descriptor.
    """
    if subtitulo is None:
        subtitulo = ("Percepciones entre profesionales sobre los cuidados "
                     "paliativos domiciliarios")
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")

    if es_portada:
        # Portada con diseño institucional
        ax = fig.add_axes([0, 0, 1, 1]); ax.axis("off")
        # Banda superior gruesa
        fig.patches.append(mpatches.Rectangle((0, 0.78), 1, 0.22,
                                              transform=fig.transFigure,
                                              facecolor="#003D7C", zorder=0))
        # Banda inferior fina
        fig.patches.append(mpatches.Rectangle((0, 0), 1, 0.08,
                                              transform=fig.transFigure,
                                              facecolor="#E8A020", zorder=0))
        fig.text(0.5, 0.91, "MULTIPAL",
                 ha="center", va="center", fontsize=38,
                 fontweight="bold", color="white")
        fig.text(0.5, 0.83, "Informe de Resultados",
                 ha="center", va="center", fontsize=16,
                 color="white", style="italic")
        fig.text(0.5, 0.62,
                 "Percepciones entre profesionales sanitarios\n"
                 "sobre las dimensiones clave de los cuidados\n"
                 "paliativos domiciliarios:\n"
                 "un análisis multicriterio comparativo",
                 ha="center", va="center", fontsize=17,
                 fontweight="bold", color="#003D7C", linespacing=1.5)
        fig.text(0.5, 0.40, contenido, ha="center", va="center",
                 fontsize=11, color="#333", linespacing=1.8)
        fig.text(0.5, 0.18,
                 "Métodos integrados:\nTOPSIS  ·  AHP  ·  PROMETHEE II",
                 ha="center", va="center", fontsize=12,
                 fontweight="bold", color="#003D7C", linespacing=1.5)
        fig.text(0.5, 0.10, f"Generado el {datetime.now():%d de %B de %Y}",
                 ha="center", va="center", fontsize=10, color="white")
    else:
        # Encabezado estándar
        ax_head = fig.add_axes([0, 0.94, 1, 0.06])
        ax_head.set_facecolor("#003D7C")
        ax_head.text(0.02, 0.5, f"MULTIPAL  ·  {subtitulo}",
                     va="center", ha="left", color="white",
                     fontsize=9.5, fontweight="bold",
                     transform=ax_head.transAxes)
        if numero_pagina is not None and total_paginas is not None:
            ax_head.text(0.98, 0.5, f"Página {numero_pagina} de {total_paginas}",
                         va="center", ha="right", color="white", fontsize=8.5,
                         transform=ax_head.transAxes)
        ax_head.set_xticks([]); ax_head.set_yticks([])
        for s in ax_head.spines.values(): s.set_visible(False)

        # Título
        ax_tit = fig.add_axes([0.05, 0.85, 0.9, 0.07])
        ax_tit.text(0, 0.5, titulo, va="center", ha="left",
                    fontsize=18, fontweight="bold", color="#003D7C",
                    transform=ax_tit.transAxes)
        ax_tit.axhline(0.0, color="#E8A020", lw=2)
        ax_tit.axis("off")

        # Contenido (envuelto automáticamente respetando los saltos
        # de línea explícitos del original)
        ancho_max = 130
        parrafos = contenido.split("\n")
        contenido_envuelto = "\n".join(
            textwrap.fill(p, width=ancho_max,
                          subsequent_indent="      " if p.startswith("  ") else "")
            if p.strip() else p
            for p in parrafos
        )

        ax_body = fig.add_axes([0.07, 0.07, 0.86, 0.76])
        ax_body.axis("off")
        ax_body.text(0, 1, contenido_envuelto, va="top", ha="left",
                     fontsize=10.5, color="#222", linespacing=1.7,
                     transform=ax_body.transAxes,
                     family="DejaVu Sans")

        # Pie
        ax_foot = fig.add_axes([0, 0, 1, 0.025])
        ax_foot.axis("off")
        ax_foot.text(0.5, 0.5,
                     f"Informe generado el {datetime.now():%d/%m/%Y a las %H:%M}",
                     ha="center", va="center", fontsize=7.5, color="gray",
                     transform=ax_foot.transAxes)

    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _resumir_hallazgos(df):
    """
    Construye una lista de hallazgos clave del estudio en formato narrativo
    para la sección de conclusiones del informe.
    """
    hallazgos = []

    # 1. Dimensión con mayor brecha
    gaps = {f"P{i}": df[f"G_P{i}"].mean() for i in range(1, 6)}
    p_max = max(gaps, key=gaps.get)
    hallazgos.append(
        f"La dimensión con mayor brecha entre importancia y percepción de preparación es "
        f"{p_max} ({DFULL[p_max]}), con una diferencia media de "
        f"{gaps[p_max]:+.1f} puntos sobre 100. Constituye, por tanto, el "
        f"área de mejora prioritaria identificada por el conjunto de la "
        f"muestra."
    )

    # 2. Dimensión mejor resuelta
    p_min = min(gaps, key=gaps.get)
    hallazgos.append(
        f"La dimensión mejor resuelta es {p_min} ({DFULL[p_min]}), con una "
        f"brecha media de {gaps[p_min]:+.1f} puntos, lo que indica un "
        f"alineamiento adecuado entre la relevancia atribuida y el "
        f"percepción de preparación."
    )

    # 3. Peso máximo
    pesos = {f"P{i}": df[f"W_P{i}"].mean() for i in range(1, 6)}
    p_top_w = max(pesos, key=pesos.get)
    hallazgos.append(
        f"La dimensión que recibe mayor peso de prioridad es {p_top_w} "
        f"({DFULL[p_top_w]}), con una asignación media del "
        f"{pesos[p_top_w]:.1f} %, frente al 20 % que correspondería a una "
        f"distribución uniforme."
    )

    # 4. Consenso W de Kendall
    mat = df[[f"W_P{i}" for i in range(1, 6)]].dropna().values
    if len(mat) >= 2:
        ranks = np.apply_along_axis(
            lambda x: len(x) + 1 - stats.rankdata(x, "average"), 1, mat)
        w = kendall_w(ranks)
        nivel = ("alto" if w >= 0.7 else
                 "moderado" if w >= 0.5 else
                 "bajo")
        hallazgos.append(
            f"El coeficiente de concordancia W de Kendall global asciende a "
            f"{w:.3f}, lo que se interpreta como un nivel de acuerdo "
            f"{nivel} entre los profesionales en la jerarquización de las "
            f"dimensiones evaluadas."
        )

    # 5. Hospital mejor posicionado
    hospitales = sorted(df["hospital"].dropna().unique())
    groups_ok = [g for g in hospitales if (df["hospital"] == g).sum() >= 1]
    if len(groups_ok) >= 2:
        sub = df[df["hospital"].isin(groups_ok)]
        dm = np.array([[sub[sub["hospital"] == g][f"R_P{i}"].mean()
                        for i in range(1, 6)] for g in groups_ok])
        wm = np.array([sub[f"W_P{i}"].mean() for i in range(1, 6)])
        wn = wm / wm.sum()
        C, _, _, tr = topsis(dm, wn)
        idx_mejor = int(np.argmin(tr))
        hallazgos.append(
            f"En el análisis multicriterio TOPSIS, el grupo mejor "
            f"posicionado es {groups_ok[idx_mejor]}, con un coeficiente de "
            f"cercanía al ideal Ci = {C[idx_mejor]:.3f}. La convergencia "
            f"con los rankings AHP y PROMETHEE II (detallados en el "
            f"informe) refuerza la robustez de esta clasificación."
        )

    return hallazgos




# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 16 — DETECCIÓN DE DIFERENCIAS ESTADÍSTICAMENTE SIGNIFICATIVAS
# ══════════════════════════════════════════════════════════════════════════════

def _hay_diferencias_significativas(df, group_col, alpha=0.05, min_n=5):
    """
    Determina si existe al menos UNA diferencia estadísticamente
    significativa (p<alpha) entre los grupos definidos por `group_col`
    en alguna de las cinco dimensiones de percepción de preparación o importancia.

    Devuelve:
      {
        "significativo": bool,
        "tests": [ {"dim":..., "var":..., "p":..., "test":...}, ... ]
      }
    Para 2 grupos se aplica Mann-Whitney U; para 3+ Kruskal-Wallis.
    Solo se procesan grupos con n>=min_n.
    """
    tests = []
    if group_col not in df.columns:
        return {"significativo": False, "tests": tests}

    grupos = [g for g in df[group_col].dropna().unique()
              if (df[group_col] == g).sum() >= min_n]
    if len(grupos) < 2:
        return {"significativo": False, "tests": tests}

    sub = df[df[group_col].isin(grupos)]
    for prefix, lbl in [("R", "Percepción de preparación"), ("I", "Importancia")]:
        for i in range(1, 6):
            col = f"{prefix}_P{i}"
            datos_por_grupo = [sub[sub[group_col] == g][col].dropna().values
                               for g in grupos]
            if any(len(x) < min_n for x in datos_por_grupo):
                continue
            try:
                if len(grupos) == 2:
                    _, p = mannwhitneyu(*datos_por_grupo,
                                        alternative="two-sided")
                    test_name = "Mann-Whitney U"
                else:
                    _, p = kruskal(*datos_por_grupo)
                    test_name = "Kruskal-Wallis"
                if p < alpha:
                    tests.append({
                        "dim": f"P{i} ({DSHORT[f'P{i}']})",
                        "var": lbl, "p": p, "test": test_name,
                    })
            except Exception:
                continue
    return {"significativo": len(tests) > 0, "tests": tests}


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 17 — INFORME GENERAL ULTRA-CONCISO (8-12 páginas)
# ══════════════════════════════════════════════════════════════════════════════

def generar_informe_general(df, outdir):
    """
    Genera el PDF general del estudio con la estructura solicitada:

      Bloque A — Análisis global:
        1. Resumen ejecutivo
        2. Caracterización de la muestra
        3. Análisis global por rol profesional
        4. Análisis global por género
        5. Análisis global por grupo de edad
        6. Análisis global por experiencia profesional
        7. Consenso interprofesional

      Bloque B — Resultados por hospital y subunidades:
        8.  Hospital La Fe (panorámica agregada)
        9.  Subunidades del Hospital La Fe
        10. Hospital de Manises (panorámica agregada)

      Bloque C — Comparativa y ranking:
        11. Comparativa entre hospitales
        12. Ranking MULTIPAL (TOPSIS · AHP · PROMETHEE II)

      Bloque D — Conclusiones:
        13. Conclusiones e implicaciones

    Las secciones de análisis de subgrupos (3-7) solo añaden páginas si
    hay diferencias estadísticamente significativas o si la sección
    aporta información descriptiva relevante. En el resto se incluye
    una mención breve.
    """
    pdf_path = outdir / "LaFe_MULTIPAL_informe_general.pdf"
    pngs_existentes = {p.stem: p for p in outdir.glob("*.png")}

    # ── Detección de subgrupos con diferencias significativas ───────────
    sig_rol    = _hay_diferencias_significativas(df, "rol")
    sig_genero = _hay_diferencias_significativas(df, "genero")
    sig_edad   = _hay_diferencias_significativas(df, "grupo_edad")
    sig_exp    = _hay_diferencias_significativas(df, "exp_label")

    # ── Lista ordenada de hospitales padre (La Fe antes de Manises) ─────
    hospitales_padre = df["hospital_padre"].dropna().unique().tolist()
    orden_preferido = ["La Fe", "Manises"]
    hospitales_padre = (
        [h for h in orden_preferido if h in hospitales_padre] +
        [h for h in hospitales_padre if h not in orden_preferido]
    )

    # ── Cálculo dinámico del total de páginas ───────────────────────────
    # Bloque A
    paginas_A = 1 + 1   # resumen + caracterización
    paginas_A += 1 + 5  # cabecera rol + 5 figuras
    paginas_A += 1 + 1  # cabecera género + 1 figura
    paginas_A += 1 + 4  # cabecera edad + ~4 figuras disponibles
    paginas_A += 1 + 1  # cabecera experiencia + 1 figura
    paginas_A += 1 + 1  # cabecera consenso + 1 figura

    # Bloque B (cada hospital padre = 1 panorámica + N subunidades)
    paginas_B = 0
    for hp in hospitales_padre:
        paginas_B += 1   # página panorámica
        subunits = df[df["hospital_padre"] == hp]["hospital"].nunique()
        if subunits > 1:
            paginas_B += 1   # cabecera subunidades + figura comparativa

    # Bloque C
    paginas_C = 2  # comparativa hospitales + ranking MULTIPAL

    # Bloque D
    paginas_D = 1  # conclusiones

    # Total = portada + A + B + C + D
    total_paginas = 1 + paginas_A + paginas_B + paginas_C + paginas_D

    with PdfPages(pdf_path) as pdf:
        pag = 1

        # ═══════════════════════════════════════════════════════════════
        # PORTADA
        # ═══════════════════════════════════════════════════════════════
        df_validos = df[df["hospital"].notna()]
        n_total = len(df_validos)
        n_comp = (df_validos["comparable"] == True).sum()
        n_esp = (df_validos["comparable"] == False).sum()
        n_hosp = df_validos["hospital_padre"].nunique()
        portada_txt = (
            f"Muestra analizada: {n_total} profesionales sanitarios\n\n"
            f"Hospitales: {n_hosp}    ·    "
            f"Subunidades: {df_validos['hospital'].nunique()}\n\n"
            f"UHD-Paliativos Adultos (comparables): {n_comp} respuestas\n"
            f"Unidades especializadas: {n_esp} respuestas\n\n"
            f"(Los informes detallados por unidad se entregan como "
            f"documentos PDF independientes en la misma carpeta.)"
        )
        _pagina_texto(pdf, "", portada_txt, numero_pagina=pag,
                      total_paginas=total_paginas, es_portada=True)
        pag += 1

        # ═══════════════════════════════════════════════════════════════
        # BLOQUE A — ANÁLISIS GLOBAL
        # ═══════════════════════════════════════════════════════════════

        # 1. Resumen ejecutivo
        if "00_resumen_ejecutivo" in pngs_existentes:
            _añadir_pagina_imagen(
                pdf, pngs_existentes["00_resumen_ejecutivo"],
                "1. Resumen Ejecutivo",
                COMENTARIOS_FIGURAS["00_resumen_ejecutivo"],
                numero_pagina=pag, total_paginas=total_paginas)
            pag += 1

        # 2. Caracterización
        if "01_descriptivos" in pngs_existentes:
            _añadir_pagina_imagen(
                pdf, pngs_existentes["01_descriptivos"],
                "2. Caracterización de la Muestra",
                COMENTARIOS_FIGURAS["01_descriptivos"],
                numero_pagina=pag, total_paginas=total_paginas)
            pag += 1

        # 3. Rol profesional (cabecera + 5 figuras)
        figs_rol = [k for k in ["19_radar_rol", "19_radar_brecha_rol",
                                 "20_preparacion_rol",
                                 "21_importancia_rol", "22_pesos_rol",
                                 "23_gap_rol"] if k in pngs_existentes]
        if figs_rol:
            if sig_rol["significativo"]:
                tests_txt = "; ".join([f"{t['var']} {t['dim']}: p={t['p']:.3f}"
                                        for t in sig_rol["tests"][:5]])
                intro = (f"Diferencias entre los distintos perfiles "
                         f"profesionales del equipo. Se identifican "
                         f"diferencias estadísticamente significativas "
                         f"(p<0,05) en: {tests_txt}.")
            else:
                intro = ("Comparativa entre roles profesionales. No se "
                         "han detectado diferencias estadísticamente "
                         "significativas entre los colectivos.")
            _pagina_texto(pdf, "3. Análisis Global por Rol Profesional",
                          intro, numero_pagina=pag,
                          total_paginas=total_paginas)
            pag += 1
            for k in figs_rol:
                _añadir_pagina_imagen(
                    pdf, pngs_existentes[k],
                    f"3. Análisis por Rol Profesional — {_titulo_legible_figura(k)}",
                    COMENTARIOS_FIGURAS.get(k, "Comparativa por rol profesional."),
                    numero_pagina=pag, total_paginas=total_paginas)
                pag += 1

        # 4. Género (cabecera + 1 figura)
        if "14_analisis_genero" in pngs_existentes:
            if sig_genero["significativo"]:
                tests_txt = "; ".join([f"{t['var']} {t['dim']}: p={t['p']:.3f}"
                                        for t in sig_genero["tests"][:5]])
                intro = (f"Diferencias entre profesionales mujeres y "
                         f"hombres. Detectadas (p<0,05) en: {tests_txt}.")
            else:
                intro = ("Comparativa entre profesionales mujeres y "
                         "hombres. No se detectaron diferencias "
                         "estadísticamente significativas.")
            _pagina_texto(pdf, "4. Análisis Global por Género", intro,
                          numero_pagina=pag, total_paginas=total_paginas)
            pag += 1
            _añadir_pagina_imagen(
                pdf, pngs_existentes["14_analisis_genero"],
                "4. Análisis por Género",
                COMENTARIOS_FIGURAS.get("14_analisis_genero",
                                         "Comparativa por género."),
                numero_pagina=pag, total_paginas=total_paginas)
            pag += 1

        # 5. Grupo de edad (cabecera + figuras)
        figs_edad = sorted([k for k in pngs_existentes
                            if k.startswith(("15_", "16_", "17_", "18_"))])
        if figs_edad:
            if sig_edad["significativo"]:
                tests_txt = "; ".join([f"{t['var']} {t['dim']}: p={t['p']:.3f}"
                                        for t in sig_edad["tests"][:5]])
                intro = (f"Diferencias entre grupos de edad (≤35, "
                         f"36-45, 46-55 y >55 años). Detectadas: {tests_txt}.")
            else:
                intro = ("Comparativa entre grupos de edad. No se "
                         "detectaron diferencias estadísticamente "
                         "significativas.")
            _pagina_texto(pdf, "5. Análisis Global por Grupo de Edad",
                          intro, numero_pagina=pag,
                          total_paginas=total_paginas)
            pag += 1
            for k in figs_edad:
                _añadir_pagina_imagen(
                    pdf, pngs_existentes[k],
                    f"5. Análisis por Grupo de Edad — {_titulo_legible_figura(k)}",
                    COMENTARIOS_FIGURAS.get(k,
                                             "Comparativa por grupo de edad."),
                    numero_pagina=pag, total_paginas=total_paginas)
                pag += 1

        # 6. Experiencia (cabecera + 1 figura)
        if "24_experiencia_boxplots" in pngs_existentes:
            if sig_exp["significativo"]:
                tests_txt = "; ".join([f"{t['var']} {t['dim']}: p={t['p']:.3f}"
                                        for t in sig_exp["tests"][:5]])
                intro = (f"Efecto de la experiencia profesional sobre "
                         f"las valoraciones. Detectadas en: {tests_txt}.")
            else:
                intro = ("Comparativa por grupos de experiencia profesional. "
                         "No se detectaron diferencias estadísticamente "
                         "significativas.")
            _pagina_texto(pdf, "6. Análisis Global por Experiencia Profesional",
                          intro, numero_pagina=pag,
                          total_paginas=total_paginas)
            pag += 1
            _añadir_pagina_imagen(
                pdf, pngs_existentes["24_experiencia_boxplots"],
                "6. Análisis por Experiencia Profesional",
                COMENTARIOS_FIGURAS.get("24_experiencia_boxplots",
                                         "Boxplots por experiencia."),
                numero_pagina=pag, total_paginas=total_paginas)
            pag += 1

        # 7. Consenso interprofesional (cabecera + 1 figura)
        if "25_consenso_kendall" in pngs_existentes:
            mat = df[[f"W_P{i}" for i in range(1, 6)]].dropna().values
            if len(mat) >= 2:
                ranks = np.apply_along_axis(
                    lambda x: len(x) + 1 - stats.rankdata(x, "average"),
                    1, mat)
                w = kendall_w(ranks)
                nivel = ("alto" if w >= 0.7 else
                         "moderado" if w >= 0.5 else "bajo")
                intro = (f"Grado de acuerdo entre los profesionales en "
                         f"el reparto de pesos de prioridad. Coeficiente "
                         f"W de Kendall global: {w:.3f} (nivel {nivel}).")
            else:
                intro = "Análisis del consenso interprofesional."
            _pagina_texto(pdf, "7. Consenso Interprofesional",
                          intro, numero_pagina=pag,
                          total_paginas=total_paginas)
            pag += 1
            _añadir_pagina_imagen(
                pdf, pngs_existentes["25_consenso_kendall"],
                "7. Consenso Interprofesional — W de Kendall",
                COMENTARIOS_FIGURAS.get("25_consenso_kendall",
                                         "W de Kendall por grupos."),
                numero_pagina=pag, total_paginas=total_paginas)
            pag += 1

        # ═══════════════════════════════════════════════════════════════
        # BLOQUE B — RESULTADOS POR HOSPITAL Y SUBUNIDADES
        # ═══════════════════════════════════════════════════════════════
        num_seccion = 8
        for hp in hospitales_padre:
            # Hospital agregado
            tag_pano = f"99_panorama_{hp.replace(' ', '_')}"
            fig_pano = fig_hospital_panorama(df, hp, outdir, tag_pano)
            if fig_pano and Path(fig_pano).exists():
                subs = sorted(df[df["hospital_padre"] == hp]
                              ["hospital"].dropna().unique())
                subs_txt = ", ".join(subs)
                _añadir_pagina_imagen(
                    pdf, fig_pano,
                    f"{num_seccion}. Hospital {hp}",
                    f"Panorámica agregada del Hospital {hp}, que en la "
                    f"muestra actual incluye las siguientes subunidades: "
                    f"{subs_txt}. La figura combina los pesos de "
                    f"prioridad, la comparación entre importancia y "
                    f"percepción de preparación, las brechas con código semáforo y el "
                    f"perfil de percepción de preparación de cada subunidad.",
                    numero_pagina=pag, total_paginas=total_paginas)
                pag += 1
                num_seccion += 1

            # Sub-sección de subunidades del hospital (si hay más de una)
            subs = sorted(df[df["hospital_padre"] == hp]
                          ["hospital"].dropna().unique())
            if len(subs) > 1:
                # Subunidades de La Fe: se referencia el desglose
                lista = "\n".join([f"  · {s} (n={(df['hospital']==s).sum()})"
                                    for s in subs])
                intro = (
                    f"El Hospital {hp} cuenta con {len(subs)} subunidades "
                    f"asistenciales en la muestra analizada:\n\n{lista}\n\n"
                    f"Cada subunidad dispone de un informe PDF "
                    f"individual con su estructura completa de 10 "
                    f"apartados (5 clínicos + 5 de análisis de "
                    f"subgrupos). Consultar los archivos "
                    f"MULTIPAL_unidad_*.pdf en la misma carpeta de "
                    f"salida para el detalle por subunidad."
                )
                _pagina_texto(pdf, f"{num_seccion}. Subunidades del Hospital {hp}",
                              intro, numero_pagina=pag,
                              total_paginas=total_paginas)
                pag += 1
                num_seccion += 1

        # ═══════════════════════════════════════════════════════════════
        # BLOQUE C — COMPARATIVA Y RANKING
        # ═══════════════════════════════════════════════════════════════

        # Comparativa entre hospitales
        if "07_comparativa_hospitales" in pngs_existentes:
            _añadir_pagina_imagen(
                pdf, pngs_existentes["07_comparativa_hospitales"],
                f"{num_seccion}. Comparativa entre Hospitales",
                "Comparativa visual entre los hospitales comparables del "
                "estudio (UHD-Paliativos Adultos). Los tres paneles "
                "—percepción de preparación, importancia y pesos de prioridad— se "
                "muestran conjuntamente, permitiendo identificar qué "
                "centros valoran más cada dimensión y dónde se "
                "encuentran los déficits más marcados. Las unidades "
                "especializadas se analizan en sus informes "
                "individuales.",
                numero_pagina=pag, total_paginas=total_paginas)
            pag += 1
            num_seccion += 1

        # Ranking MULTIPAL
        if "26_hospital_mcda" in pngs_existentes:
            _añadir_pagina_imagen(
                pdf, pngs_existentes["26_hospital_mcda"],
                f"{num_seccion}. Ranking MULTIPAL · TOPSIS · AHP · PROMETHEE II",
                "Integración de los tres métodos de análisis "
                "multicriterio. TOPSIS evalúa la proximidad a una "
                "solución ideal; AHP construye un score ponderado y "
                "verifica la consistencia; PROMETHEE II ordena por "
                "flujos de preferencia bilaterales. Cuando los tres "
                "métodos coinciden, la conclusión es robusta.",
                numero_pagina=pag, total_paginas=total_paginas)
            pag += 1
            num_seccion += 1

        # ═══════════════════════════════════════════════════════════════
        # BLOQUE D — CONCLUSIONES
        # ═══════════════════════════════════════════════════════════════
        hallazgos = _resumir_hallazgos(df)
        texto_conclusiones = (
            "A partir del análisis integrado de los datos recogidos, "
            "se identifican los siguientes hallazgos relevantes:\n\n"
        )
        for i, h in enumerate(hallazgos, 1):
            texto_conclusiones += f"  {i}.  {h}\n\n"
        texto_conclusiones += (
            "\nLos informes individuales por unidad asistencial "
            "(MULTIPAL_unidad_*.pdf) recogen el detalle específico de "
            "cada grupo —incluidos los análisis por rol profesional, "
            "género, grupo de edad, experiencia y consenso "
            "interprofesional dentro de cada unidad— y constituyen el "
            "material principal para la devolución de resultados a los "
            "equipos participantes."
        )
        _pagina_texto(pdf, f"{num_seccion}. Conclusiones e Implicaciones",
                      texto_conclusiones,
                      numero_pagina=pag, total_paginas=total_paginas)

        meta = pdf.infodict()
        meta["Title"]    = ("MULTIPAL — Informe General · Percepciones entre "
                            "profesionales sanitarios sobre los cuidados "
                            "paliativos domiciliarios")
        meta["Author"]   = "MULTIPAL"
        meta["Subject"]  = ("Informe general del estudio MULTIPAL — "
                            "ordenado por bloques (global, La Fe, Manises, "
                            "comparativa)")
        meta["Keywords"] = ("MULTIPAL, MCDA, TOPSIS, AHP, PROMETHEE, "
                            "cuidados paliativos, domiciliarios")
        meta["CreationDate"] = datetime.now()

    print(f"     ✓ {pdf_path.name}  ({pag} páginas)")
    return pdf_path


# Alias para compatibilidad
generar_informe_general_conciso = generar_informe_general







# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 18 — INFORMES INDIVIDUALES POR UNIDAD (uno por hospital/grupo)
# ══════════════════════════════════════════════════════════════════════════════
#
# Cada unidad asistencial recibe su propio informe PDF con un mensaje
# claro y accionable para el equipo. La estructura sigue lo propuesto
# en la revisión metodológica:
#
#   1. Pesos asignados a cada dimensión
#   2. Importancia percibida
#   3. Percepción de preparación (percepción de preparación)
#   4. Brecha (GAP) por dimensión, ordenada de mayor a menor
#   5. Análisis MULTIPAL: ranking por método (formato tabla)
#
# Además:
#   · Mensaje resumen al inicio (lenguaje sencillo)
#   · Comparativa médicos vs enfermeras si la unidad tiene ambos
#     colectivos con n suficiente y existen diferencias significativas
#   · Variables sexo / edad / formación solo si arrojan hallazgo
# ══════════════════════════════════════════════════════════════════════════════

def _safe_mean(x):
    """Media segura: devuelve NaN si la serie está vacía."""
    s = pd.Series(x).dropna()
    return s.mean() if len(s) > 0 else np.nan


def _safe_sem(x):
    """Error estándar seguro: devuelve 0 si n<2 (no hay variabilidad estimable)."""
    s = pd.Series(x).dropna()
    return s.sem() if len(s) >= 2 else 0.0


def _pagina_unidad_resumen(pdf, hosp, sub, total_paginas, numero_pagina,
                            df_global):
    """Página 1 del informe por unidad: mensaje resumen accionable."""
    n  = len(sub)
    aviso = ""
    if n < 5:
        aviso = (f"\n\nAVISO: Esta unidad cuenta únicamente con {n} "
                 f"{'respuesta válida' if n==1 else 'respuestas válidas'}. "
                 f"Los resultados deben interpretarse con extrema cautela "
                 f"y considerarse exploratorios: medias e índices "
                 f"derivados de muestras tan reducidas pueden no ser "
                 f"representativos del conjunto del equipo. Se recomienda "
                 f"ampliar la muestra antes de tomar decisiones basadas "
                 f"en estos resultados.")

    pesos = {f"P{i}": _safe_mean(sub[f"W_P{i}"]) for i in range(1, 6)}
    gaps  = {f"P{i}": _safe_mean(sub[f"G_P{i}"]) for i in range(1, 6)}

    # Filtrar dimensiones sin datos válidos
    pesos_v = {k: v for k, v in pesos.items() if not pd.isna(v)}
    gaps_v  = {k: v for k, v in gaps.items()  if not pd.isna(v)}

    if not pesos_v or not gaps_v:
        texto = (f"La unidad {hosp} no tiene datos suficientes para "
                 f"generar el resumen accionable.{aviso}")
        _pagina_texto(pdf, f"{hosp} — Informe de Unidad", texto,
                      numero_pagina=numero_pagina,
                      total_paginas=total_paginas)
        return

    p_top_w   = max(pesos_v, key=pesos_v.get)
    p_top_gap = max(gaps_v,  key=gaps_v.get)
    gaps_ord  = sorted(gaps_v.items(), key=lambda x: -x[1])

    # Posición en el ranking MULTIPAL (si la unidad es comparable)
    posicion_txt = ""
    if sub["comparable"].iloc[0]:
        df_comp = df_global[df_global["comparable"] == True]
        grupos = sorted(df_comp["hospital"].dropna().unique())
        if len(grupos) >= 2 and hosp in grupos:
            try:
                dm = np.array([[df_comp[df_comp["hospital"]==g][f"R_P{i}"].mean()
                                for i in range(1,6)] for g in grupos])
                wm = np.array([df_comp[f"W_P{i}"].mean() for i in range(1,6)])
                wn = wm / wm.sum()
                C, _, _, tr = topsis(dm, wn)
                idx = grupos.index(hosp)
                posicion_txt = (
                    f"\n\nPosición en el ranking MULTIPAL entre los "
                    f"{len(grupos)} hospitales comparables: "
                    f"#{int(tr[idx])} (TOPSIS Ci = {C[idx]:.3f}; "
                    f"cuanto más cercano a 1, mejor posicionado "
                    f"respecto al ideal teórico)."
                )
            except Exception:
                pass
    else:
        posicion_txt = (
            "\n\nEsta unidad no participa en el ranking comparativo "
            "entre hospitales por tratarse de una unidad especializada "
            "(no UHD-Paliativos Adultos). Su análisis es individual."
        )

    texto = (
        f"Este informe recoge los resultados específicos de la unidad "
        f"{hosp}, con {n} {'respuesta válida' if n==1 else 'respuestas válidas'}. "
        f"Su objetivo es ofrecer al equipo una devolución clara y "
        f"accionable.\n\n"
        f"────────────────────────────────────────────────\n"
        f"MENSAJES CLAVE\n"
        f"────────────────────────────────────────────────\n\n"
        f"1. Peso principal:\n"
        f"   El equipo otorga mayor importancia relativa a la dimensión "
        f"{p_top_w} ({DFULL[p_top_w]}), con un peso medio del "
        f"{pesos_v[p_top_w]:.1f} % sobre el reparto total.\n\n"
        f"2. Mayor déficit detectado (GAP):\n"
        f"   La diferencia más marcada entre importancia atribuida y "
        f"percepción de preparación se encuentra en {p_top_gap} "
        f"({DFULL[p_top_gap]}), con una brecha media de "
        f"{gaps_v[p_top_gap]:+.1f} puntos sobre 100. Es, por tanto, "
        f"el área prioritaria de mejora para este equipo.\n\n"
        f"3. Brechas ordenadas de mayor a menor (para priorizar "
        f"actuaciones):\n"
    )
    for i, (p, g) in enumerate(gaps_ord, 1):
        texto += f"     {i}. {p} ({DSHORT[p]}): {g:+.1f} puntos\n"
    texto += posicion_txt + aviso

    _pagina_texto(pdf, f"{hosp} — Informe de Unidad", texto,
                  numero_pagina=numero_pagina,
                  total_paginas=total_paginas)


def _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas):
    """Cabecera estándar de página para los informes por unidad."""
    ax_head = fig.add_axes([0, 0.94, 1, 0.06])
    ax_head.set_facecolor("#003D7C")
    ax_head.text(0.02, 0.5, f"MULTIPAL  ·  {hosp}", va="center",
                 ha="left", color="white", fontsize=9.5,
                 fontweight="bold", transform=ax_head.transAxes)
    ax_head.text(0.98, 0.5, f"Página {numero_pagina} de {total_paginas}",
                 va="center", ha="right", color="white", fontsize=8.5,
                 transform=ax_head.transAxes)
    ax_head.set_xticks([]); ax_head.set_yticks([])
    for s in ax_head.spines.values(): s.set_visible(False)


def _pie_pagina_unidad(fig):
    """Pie estándar de página."""
    ax_foot = fig.add_axes([0, 0, 1, 0.025])
    ax_foot.axis("off")
    ax_foot.text(0.5, 0.5,
                 f"Informe generado el {datetime.now():%d/%m/%Y a las %H:%M}",
                 ha="center", va="center", fontsize=7.5, color="gray",
                 transform=ax_foot.transAxes)


def _titulo_pagina_unidad(fig, titulo):
    """Bloque de título con línea naranja."""
    ax_tit = fig.add_axes([0.05, 0.86, 0.9, 0.06])
    ax_tit.text(0, 0.5, titulo, va="center", ha="left",
                fontsize=15, fontweight="bold", color="#003D7C",
                transform=ax_tit.transAxes)
    ax_tit.axhline(0, color="#E8A020", lw=2)
    ax_tit.axis("off")


def _comentario_pagina_unidad(fig, comentario):
    """Bloque de comentario interpretativo al pie."""
    ax_com = fig.add_axes([0.07, 0.05, 0.86, 0.21])
    ax_com.axis("off")
    ax_com.text(0, 0.95, "Lectura interpretativa",
                va="top", ha="left", fontsize=10, fontweight="bold",
                color="#003D7C", transform=ax_com.transAxes)
    ax_com.text(0, 0.78, textwrap.fill(comentario, width=140),
                va="top", ha="left", fontsize=9.5, color="#222",
                linespacing=1.45, transform=ax_com.transAxes)
    ax_com.axhline(0.88, color="#003D7C", lw=0.7)


def _pagina_unidad_grafica(pdf, sub, hosp, titulo, prefix, ylabel,
                            comentario, numero_pagina, total_paginas,
                            ylim=115, color="#003D7C", df_global=None):
    """
    Página de gráfica de barras por dimensión, con comparación opcional
    contra la media global del estudio.

    Usada para apartados 1, 2 y 3 (pesos, importancia, preparación).
    Tolera muestras de n=1 sin error.

    Si se pasa `df_global`, se añade una segunda serie de barras (más
    estrechas, con patrón rayado) con la media de TODOS los profesionales
    del estudio para que cada subunidad pueda ver cómo se aleja o
    acerca del conjunto global.
    """
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig, titulo)

    ax = fig.add_axes([0.10, 0.32, 0.80, 0.50])
    medias = [_safe_mean(sub[f"{prefix}_P{i}"]) for i in range(1, 6)]
    sems   = [_safe_sem(sub[f"{prefix}_P{i}"])  for i in range(1, 6)]
    colores = DCOL if prefix == "W" else [color]*5

    # Si hay df_global, dibujamos barras dobles (unidad vs global)
    if df_global is not None and len(df_global) > 0:
        medias_glob = [_safe_mean(df_global[f"{prefix}_P{i}"])
                       for i in range(1, 6)]
        x = np.arange(5)
        width = 0.38
        bars = ax.bar(x - width/2, medias, width, color=colores,
                       yerr=sems, capsize=4, edgecolor="black", linewidth=0.6,
                       label=f"Esta unidad (n={len(sub)})")
        bars_g = ax.bar(x + width/2, medias_glob, width,
                         color="lightgray", edgecolor="black", linewidth=0.6,
                         hatch="//", alpha=0.7,
                         label=f"Media global (n={len(df_global)})")
        # Etiquetas numéricas
        for b, v in zip(bars, medias):
            if pd.isna(v):
                continue
            ax.text(b.get_x() + b.get_width()/2,
                    b.get_height() + ylim*0.015,
                    f"{v:.1f}" + ("%" if prefix == "W" else ""),
                    ha="center", fontsize=9, fontweight="bold")
        for b, v in zip(bars_g, medias_glob):
            if pd.isna(v):
                continue
            ax.text(b.get_x() + b.get_width()/2,
                    b.get_height() + ylim*0.015,
                    f"{v:.1f}" + ("%" if prefix == "W" else ""),
                    ha="center", fontsize=8, color="#555")
        ax.set_xticks(x)
        ax.legend(fontsize=9, loc="upper right", framealpha=0.95)
    else:
        # Versión clásica (sin comparativa)
        bars = ax.bar(DIM, medias, color=colores, yerr=sems, capsize=5,
                      edgecolor="black", linewidth=0.6)
        for b, v in zip(bars, medias):
            if pd.isna(v):
                continue
            ax.text(b.get_x() + b.get_width()/2,
                    b.get_height() + ylim*0.015,
                    f"{v:.1f}" + ("%" if prefix == "W" else ""),
                    ha="center", fontsize=10.5, fontweight="bold")
        ax.set_xticks(range(5))

    ax.set_xticklabels([f"{p}\n{DSHORT[p]}" for p in DIM], fontsize=9)
    ax.set_ylim(0, ylim)
    ax.set_ylabel(ylabel, fontsize=11)
    ax.grid(axis="y", alpha=0.3)

    _comentario_pagina_unidad(fig, comentario)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_unidad_gap(pdf, sub, hosp, numero_pagina, total_paginas):
    """
    Apartado 4: gráfica de GAPs ordenados de mayor a menor con código
    semáforo y comentario en lenguaje sencillo.
    """
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig,
        "4. Brecha (Importancia − Preparación) por dimensión")

    ax = fig.add_axes([0.20, 0.32, 0.70, 0.50])
    gaps = [(f"P{i}", _safe_mean(sub[f"G_P{i}"])) for i in range(1, 6)]
    gaps = [(p, v) for p, v in gaps if not pd.isna(v)]
    if not gaps:
        ax.text(0.5, 0.5, "Sin datos suficientes para calcular las brechas.",
                ha="center", va="center", fontsize=12, color="#555",
                transform=ax.transAxes)
        ax.axis("off")
        _pie_pagina_unidad(fig)
        pdf.savefig(fig, dpi=150, facecolor="white")
        plt.close(fig)
        return

    gaps_ord = sorted(gaps, key=lambda x: x[1])  # plot de menor a mayor
    etiquetas = [f"{p}: {DSHORT[p]}" for p, _ in gaps_ord]
    valores   = [g for _, g in gaps_ord]
    colores   = ["#2E9E6B" if v < 5 else "#E8A020" if v < 10 else "#E74C3C"
                 for v in valores]
    bars = ax.barh(etiquetas, valores, color=colores, edgecolor="black",
                   linewidth=0.6)
    ax.axvline(0, color="black", lw=1)
    ax.set_xlabel("Brecha (puntos · Importancia − Preparación)", fontsize=11)
    ax.grid(axis="x", alpha=0.3)
    for b, v in zip(bars, valores):
        ax.text(b.get_width() + (0.3 if v >= 0 else -0.3),
                b.get_y() + b.get_height()/2,
                f"{v:+.1f}", va="center",
                ha="left" if v >= 0 else "right",
                fontsize=11, fontweight="bold")
    ax.legend(handles=[
        mpatches.Patch(color="#E74C3C", label="Déficit alto (>10)"),
        mpatches.Patch(color="#E8A020", label="Déficit moderado (5-10)"),
        mpatches.Patch(color="#2E9E6B", label="Adecuado (<5)"),
    ], loc="lower right", fontsize=9, framealpha=0.95)

    p_top = max(gaps, key=lambda x: x[1])
    com = (
        f"La gráfica muestra la diferencia entre la importancia que el "
        f"equipo otorga a cada dimensión y la preparación que percibe "
        f"para abordarla, ordenada de menor a mayor brecha. Las "
        f"dimensiones rojas señalan un déficit importante: el equipo "
        f"considera que son temas relevantes pero no se siente "
        f"suficientemente preparado para afrontarlos. Las verdes "
        f"reflejan equilibrio. La dimensión con mayor brecha es "
        f"{p_top[0]} ({DFULL[p_top[0]]}), con {p_top[1]:+.1f} puntos, "
        f"y debería ser el foco principal de las acciones de mejora "
        f"que se propongan en la unidad."
    )
    _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_unidad_multipal(pdf, hosp, sub, df_global, outdir,
                             numero_pagina, total_paginas):
    """
    Apartado 5: tabla del ranking MULTIPAL.
    Si la unidad no es comparable, se muestra una explicación
    metodológica en lugar de la tabla.
    """
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig,
        "5. Posición en el análisis multicriterio MULTIPAL")

    ax = fig.add_axes([0.05, 0.32, 0.9, 0.52])
    ax.axis("off")

    if not sub["comparable"].iloc[0]:
        # Unidad no comparable: explicación metodológica
        msg = (
            f"La unidad {hosp} es una unidad especializada (no UHD-"
            f"Paliativos Adultos) y, por tanto, no participa en el "
            f"ranking comparativo del estudio: comparar su percepción de preparación "
            f"con el de otras UHD-Paliativos Adultos no sería "
            f"metodológicamente correcto, ya que atienden a "
            f"poblaciones clínicamente distintas, con perfiles de "
            f"paciente y prioridades asistenciales propios.\n\n"
            f"En su lugar, este informe presenta los resultados "
            f"individuales de la unidad (apartados 1 a 4), que "
            f"constituyen un material de devolución específico para "
            f"el equipo. Cuando se disponga de datos de unidades "
            f"análogas (otras unidades de Pediatría / Salud Mental / "
            f"Salud Infanto-Juvenil), será posible construir un "
            f"ranking comparativo dentro de la misma categoría."
        )
        ax.text(0.05, 0.95, textwrap.fill(msg, width=110),
                va="top", ha="left", fontsize=11, color="#222",
                linespacing=1.6, transform=ax.transAxes)
    else:
        # Unidad comparable: calcular ranking
        df_comp = df_global[df_global["comparable"] == True]
        grupos = sorted(df_comp["hospital"].dropna().unique())
        if len(grupos) >= 2:
            try:
                dm = np.array([[df_comp[df_comp["hospital"]==g][f"R_P{i}"].mean()
                                for i in range(1,6)] for g in grupos])
                wm = np.array([df_comp[f"W_P{i}"].mean() for i in range(1,6)])
                wn = wm / wm.sum()
                C, _, _, tr = topsis(dm, wn)
                ahp_df = aplicar_ahp(df_comp, "hospital", outdir, "unidad_aux")
                _, _, phi, pr = promethee_ii(dm, wn)

                datos = [["Hospital", "n", "TOPSIS Ci", "T#",
                           "AHP Score", "A#", "PROMETHEE φ", "P#"]]
                for j, g in enumerate(grupos):
                    n_g = (df_comp["hospital"]==g).sum()
                    ahp_r = ahp_df[ahp_df["Grupo"]==g]["AHP_Rank"].values
                    ahp_s = ahp_df[ahp_df["Grupo"]==g]["AHP_Score"].values
                    datos.append([str(g), int(n_g),
                                  f"{C[j]:.3f}", f"#{int(tr[j])}",
                                  f"{ahp_s[0]:.1f}" if len(ahp_s) else "-",
                                  f"#{int(ahp_r[0])}" if len(ahp_r) else "-",
                                  f"{phi[j]:.3f}", f"#{int(pr[j])}"])
                tbl = ax.table(cellText=datos[1:], colLabels=datos[0],
                                cellLoc="center", loc="center",
                                bbox=[0.05, 0.25, 0.9, 0.70])
                tbl.auto_set_font_size(False); tbl.set_fontsize(11)
                tbl.auto_set_column_width(list(range(len(datos[0]))))
                # Resaltar la fila de la unidad actual
                for j, g in enumerate(grupos):
                    color = "#E8F4D8" if g == hosp else "white"
                    for col in range(len(datos[0])):
                        tbl[j+1, col].set_facecolor(color)
                # Cabecera de tabla
                for col in range(len(datos[0])):
                    tbl[0, col].set_facecolor("#003D7C")
                    tbl[0, col].set_text_props(color="white",
                                               fontweight="bold")
            except Exception as e:
                ax.text(0.5, 0.5,
                        f"No se pudo calcular el ranking: {e}",
                        ha="center", va="center", fontsize=11,
                        color="#555", transform=ax.transAxes)
        else:
            ax.text(0.5, 0.5,
                    "No hay suficientes hospitales comparables (≥2) en "
                    "la muestra actual para construir el ranking "
                    "MULTIPAL. Cuando se incorporen los datos de "
                    "Arnau / Xàtiva / otros centros se generará "
                    "automáticamente.",
                    ha="center", va="center", fontsize=11, color="#555",
                    transform=ax.transAxes, wrap=True)

    # Comentario interpretativo SIEMPRE con wrap (correcciones pendientes)
    if sub["comparable"].iloc[0]:
        com = (
            "Interpretación de los indicadores: TOPSIS Ci toma valores "
            "entre 0 (peor) y 1 (ideal); valores ≥ 0,5 indican "
            "cercanía al ideal. AHP Score es el percepción de preparación global "
            "ponderado por los pesos asignados (escala 0-100, mayor "
            "es mejor). PROMETHEE φ es el flujo neto de preferencias: "
            "positivo si el grupo supera a más alternativas de las "
            "que le superan. La fila resaltada corresponde a esta "
            "unidad. La convergencia de los tres métodos refuerza la "
            "robustez del ranking; las divergencias señalan grupos "
            "cuya posición depende del criterio metodológico."
        )
        _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_unidad_medicos_enfermeras(pdf, hosp, sub, numero_pagina,
                                       total_paginas):
    """
    Página opcional: comparativa médicos vs enfermeras dentro de la
    unidad. Solo se llama si hay diferencias significativas (p<0,05).
    """
    med = sub[sub["rol"] == "Médico"]
    enf = sub[sub["rol"] == "Enfermera"]
    if len(med) < 3 or len(enf) < 3:
        return False

    sig_dims = []
    for prefix, lbl in [("R", "Preparación"), ("I", "Importancia"),
                         ("W", "Pesos")]:
        for i in range(1, 6):
            try:
                _, p = mannwhitneyu(med[f"{prefix}_P{i}"].dropna(),
                                    enf[f"{prefix}_P{i}"].dropna(),
                                    alternative="two-sided")
                if p < 0.05:
                    sig_dims.append((prefix, i, lbl, p))
            except Exception:
                continue
    if not sig_dims:
        return False

    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig, "6. Alineamiento entre médicos y enfermeras")

    axs = [fig.add_axes([0.07+0.31*k, 0.35, 0.26, 0.45]) for k in range(3)]
    for ax, prefix, tit, ylim in zip(axs, ["R", "I", "W"],
                                       ["Preparación", "Importancia", "Pesos (%)"],
                                       [115, 115, 60]):
        m = [med[f"{prefix}_P{i}"].mean() for i in range(1, 6)]
        e = [enf[f"{prefix}_P{i}"].mean() for i in range(1, 6)]
        x = np.arange(5)
        ax.bar(x-0.2, m, 0.4, color="#003D7C",
               label=f"Médicos (n={len(med)})")
        ax.bar(x+0.2, e, 0.4, color="#2E9E6B",
               label=f"Enfermeras (n={len(enf)})")
        for prf, i, _, p in sig_dims:
            if prf == prefix:
                ax.text(i-1, max(m[i-1], e[i-1]) + ylim*0.04, "*",
                        ha="center", fontsize=18, color="red",
                        fontweight="bold")
        ax.set_xticks(x); ax.set_xticklabels(DIM, fontsize=9)
        ax.set_ylim(0, ylim); ax.set_title(tit, fontweight="bold")
        ax.legend(fontsize=8); ax.grid(axis="y", alpha=0.3)

    sig_txt = ", ".join([f"{lbl} {DIM[i-1]} (p={p:.3f})"
                          for _, i, lbl, p in sig_dims[:5]])
    com = (
        f"Se detectaron diferencias estadísticamente significativas "
        f"(p<0,05, test de Mann-Whitney U) entre médicos y enfermeras "
        f"de la unidad en: {sig_txt}. Los asteriscos en rojo señalan "
        f"esas dimensiones. Estas divergencias pueden tener "
        f"implicaciones prácticas para la coordinación interna del "
        f"equipo y suelen merecer una conversación específica con el "
        f"colectivo, especialmente cuando la dimensión coincide con "
        f"un déficit (GAP) elevado."
    )
    _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)
    return True


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 19 — FIGURAS DE SUBGRUPOS POR UNIDAD (rol, género, edad, exp, W)
# ══════════════════════════════════════════════════════════════════════════════
#
# Estas funciones replican el análisis por subgrupos del informe global
# (rol profesional, género, grupo de edad, experiencia, consenso W de
# Kendall) pero restringido al subconjunto de datos de cada unidad
# asistencial. Las figuras producidas se incorporan al PDF individual
# de cada unidad como apartados 3 a 7.
#
# Cuando la muestra de la unidad no permite realizar un análisis
# determinado (porque ningún subgrupo alcanza n≥2, o porque solo hay
# una categoría representada), se omite el análisis correspondiente y
# se incluye una página de aviso explicativo en su lugar.
# ══════════════════════════════════════════════════════════════════════════════

def _fig_genero_unidad(sub_g, hosp, slug, outdir):
    """Figura compacta de género (6 paneles): R, I, W medias por género."""
    generos = sorted(sub_g["genero"].dropna().unique())
    fig, axs = plt.subplots(1, 3, figsize=(15, 5))
    fig.suptitle(f"Análisis por Género — {hosp}",
                 fontsize=13, fontweight="bold")
    colores = {"Femenino": "#9B59B6", "Masculino": "#3498DB"}
    for ax, (prefix, tit, yl) in zip(axs,
            [("R", "Percepción de preparación", 115), ("I", "Importancia", 115),
             ("W", "Pesos (%)", 60)]):
        x = np.arange(5)
        width = 0.35
        for j, g in enumerate(generos):
            sg = sub_g[sub_g["genero"] == g]
            m = [sg[f"{prefix}_P{i}"].mean() for i in range(1, 6)]
            e = [sg[f"{prefix}_P{i}"].sem() if len(sg) > 1 else 0
                 for i in range(1, 6)]
            offset = (j - (len(generos)-1)/2) * width
            ax.bar(x + offset, m, width, yerr=e, capsize=3,
                   color=colores.get(g, "#777"),
                   label=f"{g} (n={len(sg)})", edgecolor="black", lw=0.5)
        ax.set_xticks(x); ax.set_xticklabels(DIM, fontsize=9)
        ax.set_title(tit, fontweight="bold")
        ax.set_ylim(0, yl); ax.grid(axis="y", alpha=0.3)
        ax.legend(fontsize=8)
    plt.tight_layout()
    save(fig, f"unit_{slug}_genero", outdir)


def _fig_boxplots_edad_unidad(sub_e, hosp, slug, outdir):
    """Boxplots de percepción de preparación por grupo de edad y dimensión."""
    edades = sorted(sub_e["grupo_edad"].dropna().unique(),
                    key=lambda x: ["≤35","36-45","46-55",">55"].index(x))
    fig, axs = plt.subplots(2, 5, figsize=(18, 8))
    fig.suptitle(f"Percepción de preparación e Importancia por Grupo de Edad — {hosp}",
                 fontsize=13, fontweight="bold")
    for col_i, dim in enumerate(DIM):
        for row_i, (prefix, ylabel) in enumerate(
                [("R", "Percepción de preparación"), ("I", "Importancia")]):
            ax = axs[row_i, col_i]
            data = [sub_e[sub_e["grupo_edad"] == e][f"{prefix}_{dim}"]
                       .dropna().values for e in edades]
            ax.boxplot(data, tick_labels=edades, patch_artist=True,
                       boxprops=dict(facecolor="#AED6F1"))
            ax.set_title(f"{dim}: {DSHORT[dim]}", fontsize=10)
            if col_i == 0:
                ax.set_ylabel(ylabel, fontsize=10)
            ax.tick_params(axis="x", rotation=30, labelsize=8)
            ax.set_ylim(0, 105); ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    save(fig, f"unit_{slug}_edad_box", outdir)


def _fig_experiencia_unidad(sub_x, hosp, slug, outdir):
    """Boxplots de percepción de preparación e importancia por grupos de experiencia."""
    orden = ["<1 año", "1-3 años", "3-5 años", ">5 años"]
    exps = [e for e in orden if e in sub_x["exp_label"].dropna().unique()]
    fig, axs = plt.subplots(2, 5, figsize=(18, 8))
    fig.suptitle(f"Percepción de preparación e Importancia por Experiencia — {hosp}",
                 fontsize=13, fontweight="bold")
    for col_i, dim in enumerate(DIM):
        for row_i, (prefix, ylabel) in enumerate(
                [("R", "Percepción de preparación"), ("I", "Importancia")]):
            ax = axs[row_i, col_i]
            data = [sub_x[sub_x["exp_label"] == e][f"{prefix}_{dim}"]
                       .dropna().values for e in exps]
            ax.boxplot(data, tick_labels=exps, patch_artist=True,
                       boxprops=dict(facecolor="#F9E79F"))
            ax.set_title(f"{dim}: {DSHORT[dim]}", fontsize=10)
            if col_i == 0:
                ax.set_ylabel(ylabel, fontsize=10)
            ax.tick_params(axis="x", rotation=30, labelsize=8)
            ax.set_ylim(0, 105); ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    save(fig, f"unit_{slug}_exp", outdir)


def _fig_consenso_unidad(sub, hosp, slug, outdir):
    """
    Figura del consenso W de Kendall en la unidad, desagregado por
    rol cuando hay datos suficientes. Devuelve W global.
    """
    fig, axs = plt.subplots(1, 2, figsize=(13, 5))
    fig.suptitle(f"Consenso Interprofesional (W de Kendall) — {hosp}",
                 fontsize=13, fontweight="bold")

    def _w(d):
        mat = d[[f"W_P{i}" for i in range(1, 6)]].dropna().values
        if len(mat) < 2:
            return np.nan
        ranks = np.apply_along_axis(
            lambda x: len(x) + 1 - stats.rankdata(x, "average"), 1, mat)
        return kendall_w(ranks)

    # Panel A: W global de la unidad
    w_global = _w(sub)
    axs[0].bar(["Unidad completa"], [w_global if not pd.isna(w_global) else 0],
               color="#003D7C", width=0.5)
    axs[0].axhline(0.7, color="red", ls="--", lw=1, label="Alto (≥0,7)")
    axs[0].axhline(0.5, color="orange", ls="--", lw=1, label="Moderado (≥0,5)")
    axs[0].set_ylim(0, 1)
    axs[0].set_title("Consenso global", fontweight="bold")
    axs[0].set_ylabel("W de Kendall")
    if not pd.isna(w_global):
        axs[0].text(0, w_global + 0.03, f"{w_global:.3f}",
                    ha="center", fontweight="bold", fontsize=12)
    axs[0].legend(fontsize=8); axs[0].grid(axis="y", alpha=0.3)

    # Panel B: W por rol profesional
    roles = [r for r in sub["rol"].dropna().unique()
             if (sub["rol"] == r).sum() >= 2]
    if roles:
        ws = [_w(sub[sub["rol"] == r]) for r in roles]
        ns = [(sub["rol"] == r).sum() for r in roles]
        colores = [PAL_ROL.get(r, "#777") for r in roles]
        bars = axs[1].bar([f"{r}\n(n={n})" for r, n in zip(roles, ns)],
                          [w if not pd.isna(w) else 0 for w in ws],
                          color=colores)
        for b, w in zip(bars, ws):
            if not pd.isna(w):
                axs[1].text(b.get_x() + b.get_width()/2, w + 0.03,
                            f"{w:.2f}", ha="center", fontweight="bold")
        axs[1].axhline(0.7, color="red", ls="--", lw=1)
        axs[1].axhline(0.5, color="orange", ls="--", lw=1)
        axs[1].set_ylim(0, 1)
        axs[1].set_title("Consenso por rol profesional", fontweight="bold")
        axs[1].tick_params(axis="x", labelsize=8)
        axs[1].grid(axis="y", alpha=0.3)
    else:
        axs[1].text(0.5, 0.5, "Sin roles con n≥2 en esta unidad",
                    ha="center", va="center",
                    transform=axs[1].transAxes, color="#777")
        axs[1].axis("off")
    plt.tight_layout()
    save(fig, f"unit_{slug}_consenso", outdir)
    return w_global


def _generar_figuras_subgrupos_unidad(sub, hosp, slug, outdir):
    """
    Orquestador: genera todas las figuras de subgrupos para una unidad.
    Devuelve un dict estructurado con listas de figuras y avisos.
    """
    resultado = {
        "rol":         {"figs": [], "aviso": None, "tests": []},
        "genero":      {"figs": [], "aviso": None, "tests": []},
        "edad":        {"figs": [], "aviso": None, "tests": []},
        "experiencia": {"figs": [], "aviso": None, "tests": []},
        "consenso":    {"fig": None, "aviso": None, "w": np.nan},
    }

    # ── ROL PROFESIONAL (5 figuras) ───────────────────────────────────
    roles_v = [r for r in sub["rol"].dropna().unique()
               if (sub["rol"] == r).sum() >= 2]
    if len(roles_v) >= 2:
        sub_r = sub[sub["rol"].isin(roles_v)].copy()
        pal_r = {r: PAL_ROL.get(r, make_palette(roles_v)[r])
                 for r in roles_v}
        try:
            _radar(sub_r, "rol", f"Radar por Rol — {hosp}",
                   pal_r, outdir, f"unit_{slug}_rol_radar", min_n=2)
            resultado["rol"]["figs"].append(
                ("Radar de Importancia vs Percepción de preparación",
                 outdir / f"unit_{slug}_rol_radar.png"))
        except Exception: pass
        try:
            _barras(sub_r, "rol", "R", f"Percepción de preparación por Rol — {hosp}",
                    "Percepción de preparación (0-100)", pal_r, outdir,
                    f"unit_{slug}_rol_R")
            resultado["rol"]["figs"].append(
                ("Percepción de preparación por rol",
                 outdir / f"unit_{slug}_rol_R.png"))
        except Exception: pass
        try:
            _barras(sub_r, "rol", "I", f"Importancia por Rol — {hosp}",
                    "Importancia (0-100)", pal_r, outdir,
                    f"unit_{slug}_rol_I")
            resultado["rol"]["figs"].append(
                ("Importancia por rol",
                 outdir / f"unit_{slug}_rol_I.png"))
        except Exception: pass
        try:
            _barras(sub_r, "rol", "W", f"Pesos por Rol — {hosp}",
                    "Peso (%)", pal_r, outdir,
                    f"unit_{slug}_rol_W", ylim=60)
            resultado["rol"]["figs"].append(
                ("Pesos de prioridad por rol",
                 outdir / f"unit_{slug}_rol_W.png"))
        except Exception: pass
        try:
            _heatmap_gap(sub_r, "rol", f"Brecha por Rol — {hosp}",
                         outdir, f"unit_{slug}_rol_gap")
            resultado["rol"]["figs"].append(
                ("Brecha (Importancia − Percepción de preparación) por rol",
                 outdir / f"unit_{slug}_rol_gap.png"))
        except Exception: pass
        sig = _hay_diferencias_significativas(sub_r, "rol", min_n=2)
        resultado["rol"]["tests"] = sig["tests"]
    else:
        resultado["rol"]["aviso"] = (
            f"Muestra insuficiente para análisis por rol profesional en "
            f"esta unidad: no hay al menos 2 roles representados con n≥2 "
            f"(roles detectados con n≥2: {len(roles_v)})."
        )

    # ── GÉNERO (1 figura) ─────────────────────────────────────────────
    gens_v = [g for g in sub["genero"].dropna().unique()
              if (sub["genero"] == g).sum() >= 2]
    if len(gens_v) >= 2:
        sub_g = sub[sub["genero"].isin(gens_v)].copy()
        try:
            _fig_genero_unidad(sub_g, hosp, slug, outdir)
            resultado["genero"]["figs"].append(
                ("Comparativa por género",
                 outdir / f"unit_{slug}_genero.png"))
        except Exception: pass
        sig = _hay_diferencias_significativas(sub_g, "genero", min_n=2)
        resultado["genero"]["tests"] = sig["tests"]
    else:
        resultado["genero"]["aviso"] = (
            "Muestra insuficiente para análisis por género en esta "
            "unidad (no hay al menos 2 géneros representados con n≥2)."
        )

    # ── GRUPO DE EDAD (7 figuras) ─────────────────────────────────────
    edades_v = [e for e in sub["grupo_edad"].dropna().unique()
                if (sub["grupo_edad"] == e).sum() >= 2]
    if len(edades_v) >= 2:
        sub_e = sub[sub["grupo_edad"].isin(edades_v)].copy()
        pal_e = make_palette([str(e) for e in edades_v])
        # Convertir claves a Categorical para evitar warnings
        pal_e = {e: pal_e[str(e)] for e in edades_v}
        try:
            _radar(sub_e, "grupo_edad", f"Radar por Edad — {hosp}",
                   pal_e, outdir, f"unit_{slug}_edad_radar", min_n=2)
            resultado["edad"]["figs"].append(
                ("Radar por grupo de edad",
                 outdir / f"unit_{slug}_edad_radar.png"))
        except Exception: pass
        try:
            _barras(sub_e, "grupo_edad", "R",
                    f"Percepción de preparación por Edad — {hosp}",
                    "Percepción de preparación (0-100)", pal_e, outdir,
                    f"unit_{slug}_edad_R")
            resultado["edad"]["figs"].append(
                ("Percepción de preparación por grupo de edad",
                 outdir / f"unit_{slug}_edad_R.png"))
        except Exception: pass
        try:
            _barras(sub_e, "grupo_edad", "I",
                    f"Importancia por Edad — {hosp}",
                    "Importancia (0-100)", pal_e, outdir,
                    f"unit_{slug}_edad_I")
            resultado["edad"]["figs"].append(
                ("Importancia por grupo de edad",
                 outdir / f"unit_{slug}_edad_I.png"))
        except Exception: pass
        try:
            _barras(sub_e, "grupo_edad", "W",
                    f"Pesos por Edad — {hosp}",
                    "Peso (%)", pal_e, outdir,
                    f"unit_{slug}_edad_W", ylim=60)
            resultado["edad"]["figs"].append(
                ("Pesos de prioridad por grupo de edad",
                 outdir / f"unit_{slug}_edad_W.png"))
        except Exception: pass
        try:
            _heatmap_gap(sub_e, "grupo_edad",
                         f"Brecha por Edad — {hosp}",
                         outdir, f"unit_{slug}_edad_gap")
            resultado["edad"]["figs"].append(
                ("Brecha por grupo de edad",
                 outdir / f"unit_{slug}_edad_gap.png"))
        except Exception: pass
        try:
            _fig_boxplots_edad_unidad(sub_e, hosp, slug, outdir)
            resultado["edad"]["figs"].append(
                ("Distribución (boxplots) por grupo de edad",
                 outdir / f"unit_{slug}_edad_box.png"))
        except Exception: pass
        # Séptima figura: comparativa global de medias edad por dimensión
        try:
            fig, ax = plt.subplots(1, 1, figsize=(13, 6))
            fig.suptitle(f"Perfil de medias por Edad — {hosp}",
                         fontsize=13, fontweight="bold")
            x = np.arange(5)
            for e in edades_v:
                se = sub_e[sub_e["grupo_edad"] == e]
                m = [se[f"R_P{i}"].mean() for i in range(1, 6)]
                ax.plot(x, m, "o-", lw=2, ms=8,
                        color=pal_e.get(e, "#777"),
                        label=f"{e} (n={len(se)})")
            ax.set_xticks(x); ax.set_xticklabels(DIM)
            ax.set_ylim(0, 105); ax.set_ylabel("Percepción de preparación media")
            ax.legend(fontsize=9); ax.grid(alpha=0.3)
            plt.tight_layout()
            save(fig, f"unit_{slug}_edad_perfiles", outdir)
            resultado["edad"]["figs"].append(
                ("Perfiles de percepción de preparación por grupo de edad",
                 outdir / f"unit_{slug}_edad_perfiles.png"))
        except Exception: pass
        sig = _hay_diferencias_significativas(sub_e, "grupo_edad", min_n=2)
        resultado["edad"]["tests"] = sig["tests"]
    else:
        resultado["edad"]["aviso"] = (
            "Muestra insuficiente para análisis por grupo de edad en "
            "esta unidad (no hay al menos 2 grupos con n≥2)."
        )

    # ── EXPERIENCIA PROFESIONAL (1 figura) ────────────────────────────
    exps_v = [e for e in sub["exp_label"].dropna().unique()
              if (sub["exp_label"] == e).sum() >= 2]
    if len(exps_v) >= 2:
        sub_x = sub[sub["exp_label"].isin(exps_v)].copy()
        try:
            _fig_experiencia_unidad(sub_x, hosp, slug, outdir)
            resultado["experiencia"]["figs"].append(
                ("Distribución por experiencia profesional",
                 outdir / f"unit_{slug}_exp.png"))
        except Exception: pass
        sig = _hay_diferencias_significativas(sub_x, "exp_label", min_n=2)
        resultado["experiencia"]["tests"] = sig["tests"]
    else:
        resultado["experiencia"]["aviso"] = (
            "Muestra insuficiente para análisis por experiencia "
            "profesional en esta unidad (no hay al menos 2 grupos "
            "con n≥2)."
        )

    # ── CONSENSO W DE KENDALL (1 figura) ──────────────────────────────
    if len(sub) >= 3:
        try:
            w = _fig_consenso_unidad(sub, hosp, slug, outdir)
            resultado["consenso"]["fig"] = (outdir /
                                            f"unit_{slug}_consenso.png")
            resultado["consenso"]["w"] = w
        except Exception: pass
    else:
        resultado["consenso"]["aviso"] = (
            f"Muestra insuficiente para análisis de consenso en esta "
            f"unidad (n={len(sub)}; se necesitan al menos 3 respuestas)."
        )

    return resultado


def _pagina_seccion_unidad(pdf, hosp, titulo, intro, pagina, total):
    """
    Página de cabecera de una sección dentro del informe por unidad.
    """
    _pagina_texto(pdf, f"{hosp} — {titulo}", intro,
                  numero_pagina=pagina, total_paginas=total)


def _pagina_aviso_seccion(pdf, hosp, titulo, mensaje, pagina, total):
    """
    Página de aviso cuando una sección no se puede generar
    (muestra insuficiente).
    """
    contenido = (
        f"{mensaje}\n\n"
        "Cuando se incorporen más respuestas que permitan el análisis, "
        "este apartado se generará automáticamente al volver a ejecutar "
        "el programa."
    )
    _pagina_texto(pdf, f"{hosp} — {titulo}", contenido,
                  numero_pagina=pagina, total_paginas=total)


def _añadir_seccion_subgrupos(pdf, hosp, slug, sub, datos_sub,
                                clave_grupo, titulo_seccion, descripcion,
                                pagina, total_paginas):
    """
    Añade al PDF las páginas correspondientes a una sección de subgrupo
    (rol, género, edad, experiencia). Si la sección no se pudo generar
    (datos_sub[clave_grupo]['aviso'] no None), añade una página de aviso.
    Devuelve el número de página siguiente.
    """
    info = datos_sub[clave_grupo]
    # Página de aviso si la sección no se pudo generar
    if info.get("aviso"):
        _pagina_aviso_seccion(pdf, hosp, titulo_seccion, info["aviso"],
                               pagina, total_paginas)
        return pagina + 1
    # Sin figuras: nada que añadir (debería estar el aviso, pero por
    # seguridad chequeamos)
    if not info["figs"]:
        return pagina

    # Cabecera de la sección con descripción y resumen de tests
    tests = info.get("tests", [])
    if tests:
        tests_txt = "; ".join(
            [f"{t['var']} {t['dim']}: p={t['p']:.3f}"
             for t in tests[:5]])
        intro = (f"{descripcion}\n\n"
                 f"Diferencias estadísticamente significativas detectadas "
                 f"(p<0,05): {tests_txt}.")
    else:
        intro = (f"{descripcion}\n\n"
                 f"No se detectaron diferencias estadísticamente "
                 f"significativas (p<0,05) entre los subgrupos analizados.")
    _pagina_seccion_unidad(pdf, hosp, titulo_seccion, intro,
                            pagina, total_paginas)
    pagina += 1

    # Una página por figura
    for desc_fig, ruta_fig in info["figs"]:
        if not Path(ruta_fig).exists():
            continue
        _añadir_pagina_imagen(
            pdf, ruta_fig,
            f"{titulo_seccion} — {desc_fig}",
            f"Figura de apoyo del apartado '{titulo_seccion}' "
            f"correspondiente a la unidad {hosp}. {desc_fig}.",
            numero_pagina=pagina, total_paginas=total_paginas)
        pagina += 1

    return pagina


def _añadir_seccion_consenso(pdf, hosp, sub, datos_sub, pagina,
                              total_paginas):
    """Añade la sección de consenso W de Kendall al PDF de la unidad."""
    info = datos_sub["consenso"]
    if info.get("aviso"):
        _pagina_aviso_seccion(pdf, hosp, "10. Consenso Interprofesional",
                               info["aviso"], pagina, total_paginas)
        return pagina + 1
    if info.get("fig") is None:
        return pagina

    w = info.get("w", np.nan)
    nivel = ("alto" if (not pd.isna(w) and w >= 0.7) else
             "moderado" if (not pd.isna(w) and w >= 0.5) else
             "bajo")
    intro = (
        "El coeficiente de concordancia W de Kendall cuantifica el "
        "grado de acuerdo entre los profesionales de la unidad en el "
        "reparto de pesos de prioridad entre las cinco dimensiones "
        "evaluadas. Un valor próximo a 1 indica un acuerdo "
        "casi unánime; valores próximos a 0 reflejan heterogeneidad "
        "en las preferencias.\n\n"
        f"Consenso global de la unidad: W = "
        f"{f'{w:.3f}' if not pd.isna(w) else 'no calculable'} "
        f"(nivel: {nivel})."
    )
    _pagina_seccion_unidad(pdf, hosp, "10. Consenso Interprofesional",
                            intro, pagina, total_paginas)
    pagina += 1

    if Path(info["fig"]).exists():
        _añadir_pagina_imagen(
            pdf, info["fig"],
            "10. Consenso Interprofesional — W de Kendall",
            f"Coeficiente W de Kendall en la unidad {hosp}. El panel "
            f"izquierdo muestra el consenso global; el derecho lo "
            f"descompone por rol profesional cuando hay datos "
            f"suficientes. Las líneas horizontales marcan los umbrales "
            f"de interpretación habituales (0,5 = moderado, 0,7 = alto).",
            numero_pagina=pagina, total_paginas=total_paginas)
        pagina += 1

    return pagina


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 19.B — NUEVAS PÁGINAS (contexto, resumen, MCDA dimensiones, regresión)
# ══════════════════════════════════════════════════════════════════════════════
#
# Funciones que construyen las páginas adicionales del nuevo esquema
# de informe acordado con el equipo médico:
#
#   · _pagina_contexto_objetivos:  página de introducción con los 4
#                                   objetivos del estudio y su carácter
#                                   exploratorio en el ámbito de los
#                                   cuidados paliativos domiciliarios
#   · _pagina_resumen_resultados:  página de hallazgos clave de la
#                                   subunidad (cifras resumidas)
#   · _pagina_mcda_dimensiones:    inserción del MCDA refactor en el PDF
#   · _pagina_regresion:           inserción del análisis de regresión
# ══════════════════════════════════════════════════════════════════════════════

TEXTO_CONTEXTO = (
    "El presente estudio tiene como finalidad explorar las percepciones "
    "de los profesionales sanitarios que trabajan en cuidados paliativos "
    "domiciliarios sobre cinco dimensiones clave de esta actividad "
    "asistencial: control efectivo de síntomas (P1), comunicación sobre "
    "pronóstico y decisiones compartidas (P2), atención psicosocial y "
    "espiritual (P3), coordinación y continuidad asistencial (P4) y "
    "cuidado y apoyo al equipo profesional (P5).\n\n"
    "OBJETIVOS\n"
    "────────────────────────────────────────────────────────────────\n"
    "  1.  Identificar y comparar las percepciones de médicos y "
    "enfermeros sobre las cinco dimensiones clave de los cuidados "
    "paliativos domiciliarios.\n\n"
    "  2.  Establecer una jerarquización de prioridades asistenciales "
    "mediante métodos multicriterio (TOPSIS, AHP y PROMETHEE II).\n\n"
    "  3.  Analizar el grado de consenso entre ambos grupos "
    "profesionales en la valoración de las dimensiones.\n\n"
    "  4.  Evaluar la influencia de la experiencia profesional y otras "
    "variables sociodemográficas (edad, género, rol, formación "
    "específica) sobre las percepciones recogidas.\n\n"
    "CARÁCTER EXPLORATORIO\n"
    "────────────────────────────────────────────────────────────────\n"
    "A pesar del creciente uso de métodos de análisis multicriterio "
    "(MCDA) en el ámbito sanitario, hasta la fecha no se han "
    "identificado aplicaciones publicadas específicamente en cuidados "
    "paliativos domiciliarios. El presente trabajo tiene, por tanto, "
    "un carácter exploratorio y novedoso."
)


# ── Descriptores por TIPO de unidad ──────────────────────────────────────────
# Oncología, Hematología, Salud Mental, etc. NO son unidades de
# hospitalización a domicilio (UHD) ni de cuidados paliativos domiciliarios;
# por eso sus informes NO deben titularse ni enmarcarse como tales. Cada tipo
# de unidad tiene su propio descriptor para los títulos y textos del informe.
TIPO_UNIDAD_DESC = {
    "cp_adultos": {
        "nombre":   "Unidad de Hospitalización a Domicilio (Cuidados Paliativos)",
        "contexto": "los profesionales sanitarios que trabajan en cuidados "
                    "paliativos domiciliarios",
        "ambito":   "los cuidados paliativos domiciliarios",
        "es_paliativo_uhd": True,
    },
    "cp_pediatria": {
        "nombre":   "Unidad de Cuidados Paliativos Pediátricos",
        "contexto": "los profesionales sanitarios que atienden cuidados "
                    "paliativos pediátricos",
        "ambito":   "los cuidados paliativos pediátricos",
        "es_paliativo_uhd": False,
    },
    "oncologia": {
        "nombre":   "Unidad de Oncología",
        "contexto": "los profesionales sanitarios de la Unidad de Oncología",
        "ambito":   "la Unidad de Oncología",
        "es_paliativo_uhd": False,
    },
    "hematologia": {
        "nombre":   "Unidad de Hematología",
        "contexto": "los profesionales sanitarios de la Unidad de Hematología",
        "ambito":   "la Unidad de Hematología",
        "es_paliativo_uhd": False,
    },
    "salud_mental": {
        "nombre":   "Unidad de Salud Mental",
        "contexto": "los profesionales sanitarios de la Unidad de Salud Mental",
        "ambito":   "la Unidad de Salud Mental",
        "es_paliativo_uhd": False,
    },
    "salud_infantojuvenil": {
        "nombre":   "Unidad de Salud Infanto-Juvenil",
        "contexto": "los profesionales sanitarios de la Unidad de Salud "
                    "Infanto-Juvenil",
        "ambito":   "la Unidad de Salud Infanto-Juvenil",
        "es_paliativo_uhd": False,
    },
}


def _desc_unidad(tipo):
    """Descriptor del tipo de unidad (con valores por defecto = UHD adultos)."""
    return TIPO_UNIDAD_DESC.get(tipo, TIPO_UNIDAD_DESC["cp_adultos"])


def _subtitulo_cabecera(tipo="cp_adultos"):
    """Subtítulo del encabezado de las páginas de texto, según el tipo de unidad."""
    d = _desc_unidad(tipo)
    if tipo == "cp_adultos":
        return ("Percepciones entre profesionales sobre los cuidados "
                "paliativos domiciliarios")
    if tipo == "cp_pediatria":
        return ("Percepciones entre profesionales sobre los cuidados "
                "paliativos pediátricos")
    return f"Percepciones entre profesionales · {d['nombre']}"


def _texto_contexto(tipo="cp_adultos"):
    """
    Construye el texto de contexto y objetivos adaptado al tipo de unidad.
    Las UHD de cuidados paliativos adultos conservan el texto original;
    el resto de unidades (oncología, hematología, etc.) se enmarcan según
    su propio ámbito asistencial, SIN etiquetarlas como UHD ni como
    cuidados paliativos domiciliarios.
    """
    if tipo == "cp_adultos":
        return TEXTO_CONTEXTO
    d = _desc_unidad(tipo)
    ctx, ambito = d["contexto"], d["ambito"]
    return (
        f"El presente informe explora las percepciones de {ctx} sobre cinco "
        "dimensiones clave de la práctica asistencial: control efectivo de "
        "síntomas (P1), comunicación sobre pronóstico y decisiones "
        "compartidas (P2), atención psicosocial y espiritual (P3), "
        "coordinación y continuidad asistencial (P4) y cuidado y apoyo al "
        "equipo profesional (P5).\n\n"
        "OBJETIVOS\n"
        "────────────────────────────────────────────────────────────────\n"
        f"  1.  Identificar y comparar las percepciones de médicos y "
        f"enfermeros sobre las cinco dimensiones clave en {ambito}.\n\n"
        "  2.  Establecer una jerarquización de prioridades asistenciales "
        "mediante métodos multicriterio (TOPSIS, AHP y PROMETHEE II).\n\n"
        "  3.  Analizar el grado de consenso entre ambos grupos "
        "profesionales en la valoración de las dimensiones.\n\n"
        "  4.  Evaluar la influencia de la experiencia profesional y otras "
        "variables sociodemográficas (edad, género, rol, formación "
        "específica) sobre las percepciones recogidas.\n\n"
        "NOTA SOBRE EL ALCANCE\n"
        "────────────────────────────────────────────────────────────────\n"
        f"{d['nombre']}. Esta unidad no forma parte de las unidades de "
        "hospitalización a domicilio (UHD) del estudio, por lo que se "
        "analiza de forma independiente y sus resultados no se contrastan "
        "con las medias generales de las UHD."
    )


def _pagina_contexto_objetivos(pdf, total_paginas, numero_pagina,
                                titulo="Contexto y objetivos del estudio",
                                tipo_unidad="cp_adultos"):
    """Página de introducción con los objetivos, adaptada al tipo de unidad."""
    d = _desc_unidad(tipo_unidad)
    if not d["es_paliativo_uhd"] and tipo_unidad in TIPO_UNIDAD_DESC \
            and tipo_unidad != "cp_adultos":
        titulo = f"Contexto y objetivos — {d['nombre']}"
    _pagina_texto(pdf, titulo, _texto_contexto(tipo_unidad),
                  numero_pagina=numero_pagina,
                  total_paginas=total_paginas,
                  subtitulo=_subtitulo_cabecera(tipo_unidad))


def _pagina_resumen_resultados(pdf, hosp, sub, df_global,
                                total_paginas, numero_pagina,
                                resultado_mcda=None):
    """
    Página de resumen de resultados clave de una subunidad, con las
    cifras más relevantes y la comparación frente al global.
    """
    _subt = _subtitulo_cabecera(sub["tipo_unidad"].iloc[0]) if len(sub) else None
    n = len(sub)
    comparar = df_global is not None and len(df_global) > 0
    n_global = len(df_global) if comparar else 0

    # Cifras clave de la subunidad
    pesos = {f"P{i}": _safe_mean(sub[f"W_P{i}"]) for i in range(1, 6)}
    gaps = {f"P{i}": _safe_mean(sub[f"G_P{i}"]) for i in range(1, 6)}
    pesos_v = {k: v for k, v in pesos.items() if not pd.isna(v)}
    gaps_v = {k: v for k, v in gaps.items() if not pd.isna(v)}

    if not pesos_v or not gaps_v:
        _pagina_texto(pdf, f"{hosp} — Resumen de resultados",
                      f"Esta unidad cuenta con datos insuficientes "
                      f"(n={n}) para construir un resumen significativo.",
                      numero_pagina=numero_pagina,
                      total_paginas=total_paginas, subtitulo=_subt)
        return

    p_top_w = max(pesos_v, key=pesos_v.get)
    p_top_gap = max(gaps_v, key=gaps_v.get)

    # Cifras del global para comparar (solo en unidades comparables)
    if comparar:
        p_top_w_global = max(
            {f"P{i}": df_global[f"W_P{i}"].mean() for i in range(1, 6)},
            key=lambda k: df_global[f"W_{k}"].mean()
        )
        p_top_gap_global = max(
            {f"P{i}": df_global[f"G_P{i}"].mean() for i in range(1, 6)},
            key=lambda k: df_global[f"G_{k}"].mean()
        )
        coincide_w = "✓ coincide con la media de las UHD" if p_top_w == p_top_w_global \
                      else f"≠ media UHD ({p_top_w_global})"
        coincide_gap = "✓ coincide con la media de las UHD" if p_top_gap == p_top_gap_global \
                        else f"≠ media UHD ({p_top_gap_global})"
        intro = (
            f"Esta página resume los hallazgos principales obtenidos para "
            f"la unidad {hosp}, contrastándolos con la media del conjunto "
            f"de UHD del estudio (n={n_global} profesionales).\n\n"
        )
    else:
        coincide_w = "(unidad analizada por separado, sin comparación con UHD)"
        coincide_gap = "(unidad analizada por separado, sin comparación con UHD)"
        intro = (
            f"Esta página resume los hallazgos principales obtenidos para "
            f"la unidad {hosp}. Por tratarse de un entorno asistencial muy "
            f"distinto al de las UHD, se analiza de forma independiente y "
            f"no se contrasta con las medias generales del estudio.\n\n"
        )

    texto = (
        intro +
        f"DATOS DE LA SUBUNIDAD\n"
        f"────────────────────────────────────────────────────────────\n"
        f"  · Número de respuestas válidas: {n}\n"
        f"  · Tipo de unidad: {sub['tipo_unidad'].iloc[0]}\n\n"
        f"PRIORIDAD MÁXIMA (peso medio mayor)\n"
        f"────────────────────────────────────────────────────────────\n"
        f"  · En esta subunidad: {p_top_w} ({DFULL[p_top_w]}), con un "
        f"peso medio del {pesos_v[p_top_w]:.1f} %.\n"
        f"  · {coincide_w}\n\n"
        f"DÉFICIT MÁS MARCADO (mayor brecha entre importancia y "
        f"preparación)\n"
        f"────────────────────────────────────────────────────────────\n"
        f"  · En esta subunidad: {p_top_gap} ({DFULL[p_top_gap]}), con "
        f"una brecha de {gaps_v[p_top_gap]:+.1f} puntos.\n"
        f"  · {coincide_gap}\n\n"
        f"BRECHAS ORDENADAS DE MAYOR A MENOR (para priorizar "
        f"actuaciones)\n"
        f"────────────────────────────────────────────────────────────\n"
    )
    gaps_ord = sorted(gaps_v.items(), key=lambda x: -x[1])
    for i, (p, g) in enumerate(gaps_ord, 1):
        if comparar:
            g_global = df_global[f"G_{p}"].mean()
            delta = g - g_global
            signo = "↑" if delta > 0 else ("↓" if delta < 0 else "=")
            texto += (f"  {i}. {p} ({DSHORT[p]}): {g:+.1f} puntos  "
                      f"[media UHD: {g_global:+.1f} · {signo} {abs(delta):.1f}]\n")
        else:
            texto += (f"  {i}. {p} ({DSHORT[p]}): {g:+.1f} puntos\n")

    # Si tenemos resultado MCDA, añadimos el ranking principal
    if resultado_mcda and "error" not in resultado_mcda:
        rank_avg = resultado_mcda["rank_avg"]
        cons = resultado_mcda["consistencia_global"]
        nivel_cons = ("alta" if cons >= 0.7 else
                      "moderada" if cons >= 0.4 else "baja")
        # Top 3 dimensiones según ranking promedio
        ord_idx = np.argsort(rank_avg)
        top3 = [(DIM[i], DSHORT[DIM[i]], rank_avg[i]) for i in ord_idx[:3]]
        texto += (
            f"\nRANKING MULTICRITERIO DE DIMENSIONES (top-3)\n"
            f"────────────────────────────────────────────────────────────\n"
        )
        for i, (dim, desc, r) in enumerate(top3, 1):
            texto += (f"  {i}. {dim} ({desc}) — ranking promedio "
                      f"entre métodos: {r:.1f}\n")
        texto += (f"\n  Consistencia entre métodos (TOPSIS, AHP, "
                  f"PROMETHEE): {nivel_cons} (τ = {cons:+.2f}).\n")

    if n < 5:
        texto += (f"\n⚠ AVISO IMPORTANTE\n"
                  f"────────────────────────────────────────────────────────────\n"
                  f"Esta unidad cuenta con muy pocas respuestas (n={n}). "
                  f"Los resultados deben interpretarse con extrema "
                  f"cautela y se consideran únicamente orientativos.")

    _pagina_texto(pdf, f"{hosp} — Resumen de resultados", texto,
                  numero_pagina=numero_pagina,
                  total_paginas=total_paginas, subtitulo=_subt)


def _pagina_mcda_dimensiones(pdf, hosp, resultado_mcda, numero_pagina,
                              total_paginas):
    """Inserta el MCDA de dimensiones como página del informe individual."""
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig,
        "Ranking de dimensiones (MCDA · TOPSIS · AHP · PROMETHEE)")

    if "error" in resultado_mcda:
        ax = fig.add_axes([0.10, 0.32, 0.80, 0.50])
        ax.axis("off")
        ax.text(0.5, 0.5,
                f"No se pudo calcular el MCDA de dimensiones:\n"
                f"{resultado_mcda['error']}",
                ha="center", va="center", fontsize=12, color="#555",
                transform=ax.transAxes)
        _pie_pagina_unidad(fig)
        pdf.savefig(fig, dpi=150, facecolor="white")
        plt.close(fig)
        return

    # Cargar e insertar la imagen
    ruta = resultado_mcda["figura"]
    if Path(ruta).exists():
        import matplotlib.image as mpimg
        img = mpimg.imread(ruta)
        ax_img = fig.add_axes([0.05, 0.28, 0.90, 0.58])
        ax_img.imshow(img)
        ax_img.axis("off")

    # Texto interpretativo
    cons = resultado_mcda["consistencia_global"]
    nivel = resultado_mcda.get("nivel_concordancia",
                                ("alta (rankings convergentes)" if cons >= 0.7
                                 else "moderada" if cons >= 0.5 else "baja"))
    rank_consenso = resultado_mcda.get("rank_consenso",
                                         resultado_mcda.get("rank_avg"))
    top1 = DIM[int(np.argmin(rank_consenso))]
    com = (
        f"Aplicación del análisis multicriterio a las 5 dimensiones "
        f"evaluadas (alternativas), usando como criterios "
        f"el peso medio asignado (W, beneficio) y la percepción de preparación "
        f"percibido (R, tratado como coste para que un bajo "
        f"percepción de preparación señale más necesidad de intervención). La "
        f"importancia (I) se ha excluido como criterio: al no estar "
        f"sometida a restricción de reparto, no fuerza el trade-off "
        f"informativo que el MCDA necesita. La dimensión consensuada "
        f"como más prioritaria por el conjunto del equipo es {top1} "
        f"({DFULL[top1]}). La concordancia entre los tres métodos "
        f"medida con el coeficiente W de Kendall = "
        f"{cons:.3f} es {nivel}; cuando W ≥ 0,7 los rankings "
        f"convergen y la conclusión es robusta."
    )
    _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_regresion(pdf, hosp, resultado_reg, numero_pagina,
                       total_paginas):
    """Inserta el heatmap de regresión global como página del PDF."""
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig,
        "Variables explicativas (regresión lineal y logística global)")

    ruta = resultado_reg.get("figura")
    if ruta and Path(ruta).exists():
        import matplotlib.image as mpimg
        img = mpimg.imread(ruta)
        ax_img = fig.add_axes([0.04, 0.28, 0.92, 0.58])
        ax_img.imshow(img)
        ax_img.axis("off")

    # Hallazgos significativos para esta unidad o globales
    n_sig = len([h for h in resultado_reg.get("hallazgos", [])
                  if h["p"] < 0.05])
    com = (
        f"Análisis de regresión global del estudio (n="
        f"{resultado_reg.get('n_modelos', 0)} modelos por tipo). "
        f"El panel A muestra la regresión LINEAL (¿qué variables "
        f"predicen la magnitud de las puntuaciones 0-100?); el "
        f"panel B muestra la regresión LOGÍSTICA (¿qué variables "
        f"predicen estar en el grupo alto vs. bajo respecto a la "
        f"mediana?). Las celdas con asteriscos señalan asociaciones "
        f"estadísticamente significativas (* p<0,05, ** p<0,01, "
        f"*** p<0,001). En el conjunto de las 30 regresiones se "
        f"identificaron {n_sig} asociaciones significativas. El "
        f"detalle numérico (coeficientes β y odds ratio) se "
        f"encuentra en los ficheros CSV anexos."
    )
    _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_regresion_brecha(pdf, hosp, resultado_brecha, numero_pagina,
                              total_paginas):
    """
    Inserta la regresión múltiple de la BRECHA. Si NO hay asociaciones
    significativas, el forest plot no aporta información y se OMITE,
    sustituyéndose por una nota breve. Si las hay, se muestra el gráfico
    junto con la interpretación en lenguaje sencillo (p. ej. qué rangos
    de edad priorizan más cada dimensión).
    """
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig,
        "Regresión múltiple de la brecha (Importancia − Percepción de preparación)")

    n_sig = resultado_brecha.get("n_significativos", 0) if resultado_brecha else 0
    n_obs = resultado_brecha.get("n_observaciones", 0) if resultado_brecha else 0
    ruta = resultado_brecha.get("figura") if resultado_brecha else None
    interp = resultado_brecha.get("interpretacion", []) if resultado_brecha else []

    if not ruta or not Path(ruta).exists() or n_sig == 0:
        # Sin asociaciones significativas → se omite el gráfico
        ax = fig.add_axes([0.10, 0.34, 0.80, 0.46])
        ax.axis("off")
        ax.text(0.5, 0.62,
                "No se han identificado asociaciones estadísticamente\n"
                "significativas entre las características de los\n"
                "profesionales (edad, sexo, rol, experiencia o formación)\n"
                "y la brecha importancia − percepción de preparación.",
                ha="center", va="center", fontsize=12, color="#333",
                transform=ax.transAxes)
        ax.text(0.5, 0.22,
                "Por ello, el gráfico de regresión múltiple se omite en\n"
                "este informe al no aportar información relevante. La\n"
                "brecha de cada dimensión no depende de forma "
                "independiente\nde ninguna de esas variables.",
                ha="center", va="center", fontsize=10.5, color="#666",
                transform=ax.transAxes)
        _pie_pagina_unidad(fig)
        pdf.savefig(fig, dpi=150, facecolor="white")
        plt.close(fig)
        return

    # Hay asociaciones significativas → mostrar figura + interpretación
    import matplotlib.image as mpimg
    img = mpimg.imread(ruta)
    ax_img = fig.add_axes([0.04, 0.40, 0.92, 0.44])
    ax_img.imshow(img)
    ax_img.axis("off")

    com = (
        f"Regresión lineal múltiple (n={n_obs}), un modelo por dimensión, "
        f"con la BRECHA (Importancia − Percepción de preparación) como "
        f"variable dependiente y como predictores el rol, el género, la "
        f"edad, la experiencia y la formación específica en CP. Los puntos "
        f"rojos son predictores INDEPENDIENTES (p<0,05 tras ajustar por el "
        f"resto); los demás actúan como confusores. Se identificaron "
        f"{n_sig} asociaciones significativas:"
    )
    # Construir el bloque de interpretación
    if interp:
        com += "\n\n" + "\n".join(f"  • {s}" for s in interp[:8])
    _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_desalineacion(pdf, hosp, resultado_desal, numero_pagina,
                            total_paginas):
    """
    Inserta la figura de desalineación entre roles (médicos vs
    enfermería) con el test MW de la brecha por dimensión y la
    identificación de la fuente.
    """
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig,
        "Desalineación entre roles (Médicos vs Enfermería)")

    if not resultado_desal or not resultado_desal.get("figura"):
        ax = fig.add_axes([0.10, 0.32, 0.80, 0.50])
        ax.axis("off")
        ax.text(0.5, 0.5,
                f"No se pudo calcular la desalineación entre roles.\n"
                f"{resultado_desal.get('aviso', '') if resultado_desal else ''}",
                ha="center", va="center", fontsize=11, color="#555",
                transform=ax.transAxes)
        _pie_pagina_unidad(fig)
        pdf.savefig(fig, dpi=150, facecolor="white")
        plt.close(fig)
        return

    ruta = resultado_desal.get("figura")
    if ruta and Path(ruta).exists():
        import matplotlib.image as mpimg
        img = mpimg.imread(ruta)
        ax_img = fig.add_axes([0.04, 0.28, 0.92, 0.58])
        ax_img.imshow(img)
        ax_img.axis("off")

    n_med = resultado_desal.get("n_medicos", 0)
    n_enf = resultado_desal.get("n_enfermeria", 0)
    n_sig = resultado_desal.get("n_significativos", 0)
    com = (
        f"Análisis global de la desalineación entre médicos (n={n_med}) "
        f"y enfermería (n={n_enf}) en la BRECHA (Imp−Prep) de cada "
        f"dimensión. El panel A compara las brechas medias entre "
        f"ambos roles; los asteriscos rojos señalan diferencias "
        f"estadísticamente significativas (Mann-Whitney U, p<0,05). "
        f"El panel B identifica la FUENTE de la desalineación cuando "
        f"existe: si proviene únicamente de la percepción de preparación (los roles "
        f"perciben distinto nivel de preparación), únicamente de la "
        f"importancia (valoran distinto el dominio) o de ambas. En el "
        f"conjunto se identificaron {n_sig} dimensiones con "
        f"desalineación significativa, lo que orienta el diseño de "
        f"intervenciones específicas para alinear la percepción "
        f"entre los dos colectivos."
    )
    _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_brecha_rol_unidad(pdf, sub, hosp, slug, outdir,
                               numero_pagina, total_paginas):
    """
    Página con el gráfico de araña de la BRECHA (Importancia − Percepción
    de preparación) por dimensión y rol profesional dentro de la unidad.
    Si no hay roles con muestra suficiente (≥3), muestra una nota.
    """
    fig = plt.figure(figsize=(11.69, 8.27))
    fig.patch.set_facecolor("white")
    _cabecera_pagina_unidad(fig, hosp, numero_pagina, total_paginas)
    _titulo_pagina_unidad(fig,
        "Brecha Importancia − Percepción de preparación por rol (araña)")

    ruta = _radar_brecha_rol(sub,
        f"Brecha por dimensión y rol — {hosp}",
        outdir, f"unit_{slug}_brecha_rol", min_n=3, return_path=True)

    if ruta is None or not Path(ruta).exists():
        ax = fig.add_axes([0.10, 0.34, 0.80, 0.46]); ax.axis("off")
        ax.text(0.5, 0.5,
                "No hay suficientes profesionales por rol (mínimo 3 por "
                "colectivo)\npara representar la brecha por rol en esta unidad.",
                ha="center", va="center", fontsize=12, color="#555",
                transform=ax.transAxes)
        _pie_pagina_unidad(fig)
        pdf.savefig(fig, dpi=150, facecolor="white")
        plt.close(fig)
        return

    import matplotlib.image as mpimg
    img = mpimg.imread(ruta)
    ax_img = fig.add_axes([0.06, 0.30, 0.88, 0.56])
    ax_img.imshow(img); ax_img.axis("off")
    com = (
        "Cada polígono representa un rol profesional; cada vértice es la "
        "brecha media (Importancia − Percepción de preparación) de una "
        "dimensión. Cuanto más alejado del centro (círculo discontinuo = "
        "brecha 0), mayor es el déficit percibido en esa dimensión. "
        "Comparar las formas de médicos y enfermería permite ver si ambos "
        "colectivos sienten el mayor déficit en las mismas dimensiones "
        "(perfiles solapados) o en dimensiones distintas (perfiles "
        "divergentes), lo que orienta acciones formativas diferenciadas."
    )
    _comentario_pagina_unidad(fig, com)
    _pie_pagina_unidad(fig)
    pdf.savefig(fig, dpi=150, facecolor="white")
    plt.close(fig)


def _pagina_conclusiones_unidad(pdf, hosp, sub, df_global,
                                  resultado_mcda, numero_pagina,
                                  total_paginas):
    """Página final de conclusiones específicas de la subunidad."""
    _subt = _subtitulo_cabecera(sub["tipo_unidad"].iloc[0]) if len(sub) else None
    n = len(sub)
    pesos = {f"P{i}": _safe_mean(sub[f"W_P{i}"]) for i in range(1, 6)}
    gaps = {f"P{i}": _safe_mean(sub[f"G_P{i}"]) for i in range(1, 6)}
    pesos_v = {k: v for k, v in pesos.items() if not pd.isna(v)}
    gaps_v = {k: v for k, v in gaps.items() if not pd.isna(v)}

    if not pesos_v or not gaps_v:
        texto = "Datos insuficientes para emitir conclusiones."
    else:
        p_top_w = max(pesos_v, key=pesos_v.get)
        p_top_gap = max(gaps_v, key=gaps_v.get)
        texto = (
            f"CONCLUSIONES PARA LA UNIDAD {hosp}\n"
            f"────────────────────────────────────────────────────────\n\n"
            f"1.  Prioridad principal del equipo: dimensión {p_top_w} "
            f"({DFULL[p_top_w]}), con un peso medio del "
            f"{pesos_v[p_top_w]:.1f} % sobre el reparto total. "
            f"Esto indica el área a la que el equipo otorga mayor "
            f"relevancia en su práctica habitual.\n\n"
            f"2.  Área prioritaria de mejora: dimensión {p_top_gap} "
            f"({DFULL[p_top_gap]}), con una brecha de "
            f"{gaps_v[p_top_gap]:+.1f} puntos entre importancia y "
            f"preparación. Es la dimensión donde el equipo percibe "
            f"mayor desajuste y, por tanto, el foco recomendado para "
            f"las acciones de mejora.\n\n"
        )
        if resultado_mcda and "error" not in resultado_mcda:
            cons = resultado_mcda["consistencia_global"]
            nivel = ("alta" if cons >= 0.7 else
                     "moderada" if cons >= 0.4 else "baja")
            rank_avg = resultado_mcda["rank_avg"]
            top1 = DIM[int(np.argmin(rank_avg))]
            texto += (
                f"3.  Ranking multicriterio: combinando los tres "
                f"métodos (TOPSIS, AHP y PROMETHEE), la dimensión más "
                f"prioritaria resulta {top1} ({DFULL[top1]}). La "
                f"consistencia entre métodos es {nivel} "
                f"(τ = {cons:+.2f}), lo que "
                f"{'confirma' if cons >= 0.4 else 'matiza'} la "
                f"robustez del ranking.\n\n"
            )
        if df_global is not None and len(df_global) > 0:
            texto += (
                f"4.  Comparación con las UHD: revisar los apartados "
                f"anteriores para identificar las dimensiones en las que "
                f"esta unidad se aparta más de la media del conjunto de "
                f"UHD del estudio (n={len(df_global)}). Las divergencias "
                f"mayores de 10 puntos respecto a esa media merecen "
                f"análisis cualitativo en el equipo.\n\n"
            )
        else:
            texto += (
                f"4.  Análisis independiente: por tratarse de un entorno "
                f"asistencial muy distinto al de las UHD, esta unidad se "
                f"interpreta por separado y no se contrasta con las medias "
                f"generales del estudio. Las conclusiones se basan en su "
                f"propio perfil interno de prioridades y brechas.\n\n"
            )
        if n < 5:
            texto += (
                f"⚠ AVISO METODOLÓGICO\n"
                f"────────────────────────────────────────────────────────\n"
                f"Con solo {n} respuestas, los resultados deben "
                f"considerarse exploratorios. Se recomienda ampliar "
                f"la muestra antes de tomar decisiones organizativas."
            )

    _pagina_texto(pdf, f"{hosp} — Conclusiones", texto,
                  numero_pagina=numero_pagina,
                  total_paginas=total_paginas, subtitulo=_subt)


# ══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 20 — INFORMES INDIVIDUALES POR UNIDAD (con apartados 3-7 añadidos)
# ══════════════════════════════════════════════════════════════════════════════

def generar_pdfs_por_unidad(df, outdir, resultado_regresion=None,
                              resultado_regresion_brecha=None,
                              resultado_desalineacion=None):
    """
    Genera un PDF individual para cada grupo de análisis (unidad)
    siguiendo el esquema acordado con el equipo médico (v7):

      p1.  Contexto y objetivos del estudio
      p2.  Resumen de resultados clave
      p3.  Pesos asignados a cada dimensión (vs media global)
      p4.  Importancia atribuida (vs media global)
      p5.  Percepción de preparación (vs media global)
      p6.  Brecha (GAP) por dimensión ordenada
      p7.  Ranking multicriterio de dimensiones (TOPSIS + AHP +
            PROMETHEE) con criterios W y R, concordancia por W de
            Kendall y ranking consensuado
      p8.  Regresión múltiple de la BRECHA (5 modelos, β estandarizados
            con IC95%) — identificación de predictores independientes
            vs confusores
      p9.  Desalineación entre roles: test MW de la brecha por
            dimensión y identificación de la fuente (percepción de preparación,
            importancia o ambas)
      p10. Variables explicativas (regresión global secundaria)
      p11. Conclusiones específicas de la subunidad
    """
    pdfs_generados = []
    grupos = sorted(df["hospital"].dropna().unique())

    # Referencia de comparación: media del conjunto de UHD comparables.
    # Las unidades NO comparables (Oncología, Hematología, Pediatría, etc.)
    # se analizan por separado y no se contrastan con esta media.
    df_uhd = df[df["comparable"] == True].copy()

    for hosp in grupos:
        sub = df[df["hospital"] == hosp].copy()
        if len(sub) == 0:
            continue

        es_comparable = bool(sub["comparable"].iloc[0])
        df_ref = df_uhd if (es_comparable and COMPARAR_CON_MEDIAS
                            and len(df_uhd) > 0) else None

        # Marco asistencial adaptado al tipo de unidad (oncología,
        # hematología, etc. NO se enmarcan como cuidados paliativos / UHD).
        _tipo_u = sub["tipo_unidad"].iloc[0]
        _d_u = _desc_unidad(_tipo_u)
        if _d_u["es_paliativo_uhd"]:
            _amb_dim = "las cinco dimensiones del cuidado paliativo"
            _amb_complej = "la complejidad del cuidado paliativo"
        else:
            _amb_dim = "las cinco dimensiones de la práctica asistencial"
            _amb_complej = "la complejidad de la atención al paciente"

        # Texto de comparación adaptado a si realmente se muestra comparación
        if df_ref is not None:
            _ref_w   = ("junto con la media del conjunto de UHD del estudio "
                        "(barras rayadas grises). Una distribución muy "
                        "desigual respecto a esa media indica que esta "
                        "unidad prioriza claramente alguna dimensión")
            _ref_i   = ("junto con la media del conjunto de UHD. Valores "
                        "altos en todas las dimensiones indican un equipo "
                        f"que reconoce {_amb_complej}")
            _ref_r   = ("contrastado con la media del conjunto de UHD. "
                        "Valores bajos respecto a esa media señalan")
        else:
            _ref_w   = ("Esta unidad se analiza por separado, sin comparación "
                        "con las UHD por tratarse de un entorno asistencial "
                        "muy distinto. Una distribución muy desigual indica "
                        "que el equipo prioriza claramente alguna dimensión")
            _ref_i   = ("Esta unidad se interpreta de forma independiente, "
                        "sin contraste con las UHD. Valores altos en todas "
                        f"las dimensiones indican un equipo que reconoce "
                        f"{_amb_complej}")
            _ref_r   = ("Esta unidad se interpreta de forma independiente, "
                        "sin contraste con las UHD. Valores bajos señalan")

        # Nombre de archivo seguro
        slug = (hosp.lower()
                    .replace("á","a").replace("é","e").replace("í","i")
                    .replace("ó","o").replace("ú","u").replace("ñ","n")
                    .replace("—","-").replace(" ","_").replace("/","-"))
        pdf_path = outdir / f"LaFe_MULTIPAL_unidad_{slug}.pdf"

        # MCDA de dimensiones para esta subunidad
        resultado_mcda = analisis_mcda_dimensiones(sub, hosp, slug, outdir)

        # 11 páginas fijas en el nuevo esquema v7
        total_paginas = 12

        with PdfPages(pdf_path) as pdf:
            pag = 1

            # p1) Contexto y objetivos (adaptado al tipo de unidad)
            tipo_u = sub["tipo_unidad"].iloc[0]
            _pagina_contexto_objetivos(
                pdf, total_paginas, pag,
                titulo=f"{hosp} — Contexto y objetivos del estudio",
                tipo_unidad=tipo_u)
            pag += 1

            # p2) Resumen de resultados
            _pagina_resumen_resultados(
                pdf, hosp, sub, df_ref, total_paginas, pag,
                resultado_mcda=resultado_mcda)
            pag += 1

            # p3) Pesos (con comparación vs UHD solo en unidades comparables)
            _pagina_unidad_grafica(
                pdf, sub, hosp,
                "1. Pesos asignados a cada dimensión",
                "W", "Peso (%) — distribución de la prioridad",
                "Cómo reparte este equipo el 100 % de la importancia "
                f"entre {_amb_dim}, "
                f"{_ref_w} sobre el resto.",
                pag, total_paginas, ylim=60, color="#003D7C",
                df_global=df_ref)
            pag += 1

            # p4) Importancia
            _pagina_unidad_grafica(
                pdf, sub, hosp,
                "2. Importancia atribuida a cada dimensión",
                "I", "Importancia (0–100)",
                "Cuán importante considera cada dimensión esta "
                "unidad (de forma independiente, sin reparto de "
                f"100 puntos). {_ref_i} en su conjunto.",
                pag, total_paginas, ylim=115, color="#003D7C",
                df_global=df_ref)
            pag += 1

            # p5) Preparación
            _pagina_unidad_grafica(
                pdf, sub, hosp,
                "3. Percepción de preparación del equipo",
                "R", "Percepción de preparación (0–100)",
                "Cuán preparado se siente el equipo para abordar cada "
                f"dimensión. {_ref_r} necesidades formativas u "
                "organizativas concretas en esta unidad.",
                pag, total_paginas, ylim=115, color="#E8A020",
                df_global=df_ref)
            pag += 1

            # p6) GAP ordenado
            _pagina_unidad_gap(pdf, sub, hosp, pag, total_paginas)
            pag += 1

            # p7) MCDA de dimensiones (refactor v7)
            _pagina_mcda_dimensiones(pdf, hosp, resultado_mcda,
                                      pag, total_paginas)
            pag += 1

            # p8) Regresión múltiple de la BRECHA (β estandarizados, IC95%)
            if resultado_regresion_brecha is not None:
                _pagina_regresion_brecha(pdf, hosp,
                                          resultado_regresion_brecha,
                                          pag, total_paginas)
            else:
                _pagina_texto(pdf, f"{hosp} — Regresión de la brecha",
                               "El análisis no se ha podido completar.",
                               numero_pagina=pag,
                               total_paginas=total_paginas)
            pag += 1

            # p9) Desalineación entre roles (médicos vs enfermería)
            if resultado_desalineacion is not None:
                _pagina_desalineacion(pdf, hosp, resultado_desalineacion,
                                       pag, total_paginas)
            else:
                _pagina_texto(pdf, f"{hosp} — Desalineación entre roles",
                               "El análisis no se ha podido completar.",
                               numero_pagina=pag,
                               total_paginas=total_paginas)
            pag += 1

            # p10) Brecha (Imp − Preparación) por ROL — gráfico de araña
            _pagina_brecha_rol_unidad(pdf, sub, hosp, slug, outdir,
                                       pag, total_paginas)
            pag += 1

            # p11) Regresión global (secundaria — 15 puntuaciones)
            if resultado_regresion is not None:
                _pagina_regresion(pdf, hosp, resultado_regresion,
                                   pag, total_paginas)
            else:
                _pagina_texto(pdf, f"{hosp} — Variables explicativas",
                               "El análisis de regresión no se ha "
                               "podido completar para esta unidad.",
                               numero_pagina=pag,
                               total_paginas=total_paginas)
            pag += 1

            # p12) Conclusiones
            _pagina_conclusiones_unidad(pdf, hosp, sub, df_ref,
                                          resultado_mcda,
                                          pag, total_paginas)

            # Metadatos
            meta = pdf.infodict()
            meta["Title"] = f"MULTIPAL — Informe de unidad · {hosp}"
            meta["Author"] = "MULTIPAL"
            meta["Subject"] = ("Informe individual de unidad asistencial "
                                "(12 apartados: contexto, resumen, pesos, "
                                "importancia, preparación, GAP, MCDA "
                                "dimensiones refactor, regresión de la "
                                "brecha, desalineación roles, brecha por "
                                "rol (araña), regresión global, "
                                "conclusiones)")
            meta["Keywords"] = (f"MULTIPAL, MCDA, cuidados paliativos, "
                                f"{hosp}")
            meta["CreationDate"] = datetime.now()

        pdfs_generados.append(pdf_path)
        print(f"     ✓ {pdf_path.name}  ({total_paginas} págs, n={len(sub)})")

    return pdfs_generados







# ══════════════════════════════════════════════════════════════════════════════
#  PIPELINE PRINCIPAL — orquesta todos los análisis en orden
# ══════════════════════════════════════════════════════════════════════════════

def run(filepath, outdir):
    """
    Pipeline principal de MULTIPAL v4:
      0.  Resumen ejecutivo (con TOPSIS+AHP+PROMETHEE integrados)
      1.  Descriptivos generales de la muestra
      2.  Comparativa entre hospitales comparables (UHD-Paliativos Adultos)
      3.  [omitido] Subunidades — cada unidad es ya su propio grupo
      4.  Análisis por rol profesional
      5.  Análisis por género
      6.  Análisis por grupos de edad
      7.  Análisis por experiencia profesional
      8.  Consenso W de Kendall
      9.  MCDA entre hospitales comparables
      10. MCDA por rol profesional
      11. Informe de texto resumen
      12. PDF general ULTRA-CONCISO (solo lo crítico y significativo)
      13. PDFs individuales por unidad (5 apartados estándar)
    """
    outdir.mkdir(parents=True, exist_ok=True)
    print(f"\n📂 Salida → {outdir}\n")

    print("▶  Cargando datos...")
    df = load(filepath)
    df.to_csv(outdir/"datos_limpios.csv", index=False)

    print("\n▶  [0]  Resumen ejecutivo (con TOPSIS+AHP+PROMETHEE)")
    fig_resumen_ejecutivo(df, outdir)

    print("\n▶  [1]  Descriptivos generales")
    fig_descriptivos(df, outdir)

    print("\n▶  [2]  Hospitales — comparativa entre comparables")
    analisis_hospitales(df, outdir)

    print("\n▶  [3]  Subunidades (obsoleto en la nueva taxonomía)")
    analisis_subunidades(df, outdir)

    print("\n▶  [4]  Rol profesional")
    analisis_rol(df, outdir)

    print("\n▶  [5]  Género")
    analisis_genero(df, outdir)

    print("\n▶  [6]  Grupos de edad")
    analisis_edad(df, outdir)

    print("\n▶  [7]  Experiencia")
    analisis_experiencia(df, outdir)

    print("\n▶  [8]  Consenso W de Kendall")
    analisis_consenso(df, outdir)

    print("\n▶  [9]  MCDA: TOPSIS + AHP + PROMETHEE II — hospitales comparables")
    pal_h = df._pal_hosp
    df_comp = df[df["comparable"] == True].copy()
    df_comp._pal_hosp = pal_h
    analisis_mcda(df_comp, "hospital", pal_h, outdir, "26_hospital",
                  "Comparativa Hospitales (UHD-Paliativos Adultos)")

    print("\n▶  [10] MCDA por rol profesional")
    analisis_mcda(df, "rol", {**make_palette(df["rol"].dropna().unique()), **PAL_ROL},
                  outdir, "28_rol", "Por Rol Profesional")

    print("\n▶  [11] Análisis de regresión (variables explicativas)")
    resultado_regresion = analisis_regresion(df, outdir)
    print(f"   ✓ {resultado_regresion['n_modelos']} modelos lineales + "
          f"{resultado_regresion['n_modelos']} logísticos · "
          f"{resultado_regresion['n_significativos']} asociaciones "
          f"estadísticamente significativas")

    print("\n▶  [11.B] Regresión múltiple de la BRECHA (5 modelos por dimensión)")
    resultado_regresion_brecha = analisis_regresion_brecha(df, outdir)
    print(f"   ✓ {resultado_regresion_brecha['n_significativos']} "
          f"predictores significativos identificados (β estandarizados "
          f"con IC95%)")

    print("\n▶  [11.C] Desalineación entre roles (Médicos vs Enfermería)")
    resultado_desalineacion = analisis_desalineacion_roles(df, outdir)
    if resultado_desalineacion.get("figura"):
        print(f"   ✓ {resultado_desalineacion['n_significativos']} "
              f"dimensiones con desalineación estadísticamente "
              f"significativa en la brecha")
    else:
        print(f"   ⚠ {resultado_desalineacion.get('aviso', '')}")

    print("\n▶  [12] Informe texto")
    informe_texto(df, outdir)

    print("\n▶  [13] Informe general")
    generar_informe_general_conciso(df, outdir)

    print("\n▶  [14] Informes individuales por unidad")
    generar_pdfs_por_unidad(df, outdir,
                              resultado_regresion=resultado_regresion,
                              resultado_regresion_brecha=resultado_regresion_brecha,
                              resultado_desalineacion=resultado_desalineacion)

    print(f"\n{'='*62}\n  ✅  COMPLETADO — {outdir.resolve()}\n{'='*62}\n")


# ══════════════════════════════════════════════════════════════════════════════
#  PUNTO DE ENTRADA — ejecución desde línea de comandos
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MULTIPAL v7 — Análisis Multicriterio CP (TODOS los hospitales)")
    parser.add_argument("--file", "-f", default=str(INPUT_FILE),
                        help="Ruta al CSV (o Excel) exportado de la encuesta")
    parser.add_argument("--outdir", "-o", default=str(OUT_DIR),
                        help="Carpeta de salida para PNGs, CSVs y TXT")
    args = parser.parse_args()

    print("\n" + "="*70)
    print("  MULTIPAL v7 — ANÁLISIS SOLO HOSPITAL LA FE")
    print("="*70)

    # Localizar la base de datos: si no existe la ruta indicada, se busca
    # automáticamente la base de datos con la otra extensión (.csv / .xlsx)
    # en la misma carpeta. Así funciona tanto si el archivo es CSV como Excel.
    f = Path(args.file)
    if not f.exists():
        _here = Path(__file__).resolve().parent
        candidatos = [f.with_suffix(".csv"), f.with_suffix(".xlsx"),
                      f.parent / "dataBase.csv", f.parent / "dataBase.xlsx",
                      _here / "dataBase.csv", _here / "dataBase.xlsx"]
        encontrado = next((c for c in candidatos if c.exists()), None)
        if encontrado is None:
            print(f"  Archivo de entrada : {args.file}")
            print("="*70 + "\n")
            print(f"\n❌  No se encuentra la base de datos: {f}")
            print(f"   Coloca 'dataBase.csv' (o 'dataBase.xlsx') en:\n   {f.parent}\n")
            sys.exit(1)
        f = encontrado

    print(f"  Archivo de entrada : {f}")
    print(f"  Carpeta de salida  : {args.outdir}/")
    print(f"  Prefijo archivos   : MULTIPAL_*")
    print("="*70 + "\n")
    print("  ⓘ  Esta versión filtra los datos al Hospital La Fe.")
    print("     Las subunidades de HaD se analizan en conjunto; Oncología y")
    print("     Hematología, por separado. Archivos con prefijo 'LaFe_'.\n")

    run(f, Path(args.outdir))
