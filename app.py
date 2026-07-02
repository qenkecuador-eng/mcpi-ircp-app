from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

try:
    from docx import Document
except Exception:
    Document = None

APP_DIR = Path(__file__).parent
BASE_PATH = APP_DIR / "tipos_penales_base_ampliada.csv"
SUBCRITERIOS_PATH = APP_DIR / "subcriterios_MCPI_IRCP_app_unico.xlsx"@echo off
title MCPI-IRCP-I v6 subcriterios
cd /d "%~dp0"
py -m pip install -r requirements.txt
py -m streamlit run app.py# Matriz MCPI-IRCP-I v6 con subcriterios
streamlit>=1.31
pandas>=2.0
matplotlib>=3.7
openpyxl>=3.1
python-docx>=1.1País,Norma,Artículo,Tipo penal,Categoría,Estado,Contexto típico de protesta,Legalidad,Claridad normativa,Lesividad,Mínima intervención penal,Idoneidad,Necesidad penal,Peso derechos indígenas,Peso interés estatal,Observación metodológica
Argentina,Código Penal,art. 226 CP,Rebelión / sublevación,Orden constitucional,Base tesis,"Protesta indígena, defensa territorial o conflicto político-institucional.",70,65,60,45,65,55,36,32,Requiere interpretación restrictiva si se invoca frente a protesta.
Argentina,Código Penal,art. 230 CP,Sedición,Orden público / institucional,Base tesis,Protesta indígena o conflicto institucional sin violencia grave acreditada.,60,55,50,35,55,45,40,24,Riesgo por amplitud si no existe alzamiento grave e idóneo.
Argentina,Código Penal,art. 213 CP,Apología,Expresión / discurso,Base tesis,"Discurso crítico, protesta indígena, expresión colectiva o defensa territorial.",40,35,40,25,40,30,48,16,Riesgo alto de afectación a libertad de expresión.
Argentina,Código Penal,art. 239 CP,Resistencia o desobediencia a la autoridad,Autoridad pública,Base tesis,"Resistencia comunitaria, protesta no violenta o conflicto de autoridad.",50,45,45,30,50,35,40,20,Riesgo por uso expansivo ante desobediencias simbólicas o colectivas.
Argentina,Código Penal,art. 194 CP,Obstrucción de vías o servicios,Circulación / servicios,Base tesis,Corte de ruta en contexto de protesta indígena.,55,50,50,40,55,40,40,24,Tipo central en criminalización de cortes de ruta.
Argentina,Código Penal,art. 210 CP,Asociación ilícita,Organización colectiva,Base tesis,"Organización comunitaria indígena, liderazgo social o acción colectiva territorial.",35,30,30,10,35,20,48,16,Riesgo alto si se asimila organización social a estructura criminal.
Argentina,Código Penal,art. 181 CP,Usurpación,Territorio / propiedad,Base tesis,"Recuperación territorial, defensa de territorio ancestral o conflicto de propiedad/posesión.",50,45,45,30,50,35,44,20,Riesgo elevado en conflictos de posesión o territorio ancestral.
Argentina,Código Penal,arts. 237–238 CP,Atentado contra la autoridad,Autoridad pública,Ampliado preliminar,"Protesta, resistencia frente a operativo policial o autoridad pública.",52,47,45,30,50,35,40,22,Exige violencia o intimidación concreta e individualizada.
Argentina,Código Penal,arts. 149 bis / 149 ter CP,Amenazas y coacciones,Libertad individual,Ampliado preliminar,Reclamos colectivos interpretados como presión ilegítima.,55,50,45,30,50,35,42,22,Riesgo por confundir presión política legítima con amenaza penal.
Argentina,Código Penal,arts. 183–184 CP,Daños y daños agravados,Propiedad / bienes,Ampliado preliminar,"Roturas, pintadas, daños menores o afectaciones materiales en protestas.",60,55,50,40,55,45,36,28,"Requiere daño probado, individualización y proporcionalidad."
Argentina,Código Penal,arts. 186–189 CP,"Incendio, explosión u otros estragos",Seguridad pública,Ampliado preliminar,"Uso de fuego, barricadas o hechos de riesgo en contexto de protesta.",65,60,65,55,70,60,32,42,Solo debería operar con peligro común grave y probado.
Argentina,Código Penal,art. 209 CP,Instigación a cometer delitos,Expresión / convocatoria,Ampliado preliminar,"Convocatorias públicas, discursos de dirigentes o llamados a movilización.",45,40,35,25,40,30,48,18,Alto riesgo si se sanciona discurso político sin nexo directo con delito concreto.
Argentina,Código Penal,art. 211 CP,Intimidación pública,Orden público / alarma pública,Ampliado preliminar,"Mensajes, comunicados o acciones de protesta interpretados como alarma pública.",42,38,35,25,40,30,48,18,Tipo riesgoso si se aplica a expresiones de protesta.
Argentina,Código Penal,art. 41 quinquies CP,Agravante de finalidad terrorista,Terrorismo / agravante,Ampliado preliminar,Calificación agravada de protestas como actos con finalidad terrorista.,30,25,25,10,30,20,55,15,Debe excluir reclamo legítimo sin violencia grave.
Argentina,Código Penal,arts. 89–92 CP,Lesiones,Integridad personal,Ampliado preliminar,Confrontaciones en protestas u operativos policiales.,70,65,65,55,70,60,30,45,Requiere prueba individualizada de autoría y nexo causal.
Argentina,Código Penal,art. 141 CP,Privación ilegítima de libertad,Libertad individual,Ampliado preliminar,"Retenciones temporales, bloqueos o conflictos con autoridades/empresas.",60,55,50,40,55,45,38,32,"Distinguir protesta, bloqueo y verdadera privación ilegítima."
Ecuador,COIP,art. 336 COIP,Rebelión,Orden constitucional,Base tesis,"Protesta indígena, defensa territorial o conflicto político-institucional.",70,65,55,45,60,50,40,28,Riesgo si se usa frente a protesta política sin alzamiento real.
Ecuador,COIP,Sedición militar / policial,Sedición militar / policial,Orden institucional militar/policial,Base tesis,Ámbito institucional calificado; no extensible por analogía a civiles.,75,70,75,70,78,65,24,48,Riesgo bajo si se mantiene en ámbito institucional específico.
Ecuador,COIP,art. 365 COIP,Apología,Expresión / discurso,Base tesis,"Discurso crítico, protesta indígena, expresión colectiva o defensa territorial.",40,35,40,25,40,30,48,16,Riesgo alto para libertad de expresión si se usa contra discurso de protesta.
Ecuador,COIP,arts. 283 y 282 COIP,Ataque o resistencia / incumplimiento de decisiones,Autoridad pública,Base tesis,"Resistencia comunitaria, protesta no violenta o conflicto de autoridad.",45,40,40,25,45,30,44,20,Riesgo frente a actos de resistencia civil o incumplimiento en protesta.
Ecuador,COIP,art. 346 COIP,Paralización de servicio público,Servicios públicos,Base tesis,"Protesta indígena, movilización social o conflicto socioambiental.",55,50,50,40,55,40,40,24,Tipo central para protestas que afectan servicios o movilidad.
Ecuador,COIP,art. 370 COIP,Asociación ilícita,Organización colectiva,Base tesis,"Organización comunitaria indígena, liderazgo social o acción colectiva territorial.",35,30,30,10,35,20,48,16,Riesgo alto si se asocia organización social con criminalidad.
Ecuador,COIP,art. 345 COIP,Sabotaje,Seguridad pública / economía,Base tesis,"Protesta socioambiental, defensa territorial o conflicto extractivo.",45,40,35,20,45,25,48,20,Tipo de alto riesgo en conflictos extractivos o territoriales.
Ecuador,COIP,art. 366 COIP,Terrorismo,Terrorismo / seguridad pública,Ampliado preliminar,Protesta social calificada como terror o amenaza grave a la población.,30,25,25,10,30,20,55,15,Riesgo muy alto si se aplica a protesta sin violencia grave y finalidad terrorista.
Ecuador,COIP,art. 367 COIP,Financiación del terrorismo,Terrorismo / financiación,Ampliado preliminar,"Apoyo económico, logístico o comunitario a movilizaciones calificado como financiamiento ilícito.",30,25,20,10,30,20,55,15,Riesgo extremo si confunde solidaridad comunitaria con financiamiento terrorista.
Ecuador,COIP,art. 369 COIP,Delincuencia organizada,Organización colectiva,Ampliado preliminar,Liderazgos comunitarios o grupos de protesta tratados como grupo estructurado delictivo.,35,30,25,10,35,20,55,16,Riesgo alto si se criminaliza organización social.
Ecuador,COIP,art. 363 COIP,Instigación,Expresión / convocatoria,Ampliado preliminar,"Convocatorias públicas, discursos o llamados a movilización.",45,40,35,25,40,30,48,18,Riesgo si se sanciona convocatoria política sin nexo directo con delito concreto.
Ecuador,COIP,art. 200 COIP,Usurpación,Territorio / propiedad,Ampliado preliminar,"Recuperación territorial, ocupación, posesión comunitaria o conflicto de tierra.",50,45,45,30,50,35,44,20,Riesgo en conflictos territoriales indígenas o socioambientales.
Ecuador,COIP,art. 204 COIP,Daño a bien ajeno,Propiedad / bienes,Ampliado preliminar,"Daños materiales, pintadas, afectación de bienes públicos o privados en protesta.",60,55,50,40,55,45,36,28,"Requiere daño concreto, autoría individual y proporcionalidad."
Ecuador,COIP,art. 154 COIP,Intimidación,Libertad individual,Ampliado preliminar,Mensajes o acciones colectivas interpretadas como amenaza.,55,50,45,30,50,35,42,22,Riesgo por equiparar presión social o política con intimidación penal.
Ecuador,COIP,arts. 152 y ss. COIP,Lesiones,Integridad personal,Ampliado preliminar,Confrontaciones en protestas u operativos de fuerza pública.,70,65,65,55,70,60,30,45,"Requiere prueba individualizada, nexo causal y proporcionalidad."
Ecuador,COIP,arts. 161–162 COIP,Secuestro / retención ilegal,Libertad individual,Ampliado preliminar,"Retenciones, bloqueos o conflictos comunitarios con autoridades o empresas.",60,55,50,40,55,45,38,32,"Diferenciar protesta, bloqueo y verdadera privación de libertad."


Versión actualizada de la aplicación Streamlit para evaluar riesgo de criminalización penal intercultural.

## Novedades v6

- Módulo nuevo: Rúbrica de subcriterios.
- Cálculo de variables principales desde subcriterios ponderados.
- Lectura del archivo `subcriterios_MCPI_IRCP_app_unico.xlsx`.
- Edición directa de subcriterios en la app.
- Exportación de rúbrica actualizada.
- Mantiene base ampliada de delitos de Argentina y Ecuador.

## Archivos requeridos

- `app.py`
- `requirements.txt`
- `tipos_penales_base_ampliada.csv`
- `subcriterios_MCPI_IRCP_app_unico.xlsx`

## Streamlit Cloud

- Repository: `qenkecuador-eng/mcpi-ircp-app`
- Branch: `main`
- Main file path: `app.py`

pause


st.set_page_config(
    page_title="Matriz MCPI-IRCP-I v6",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

COLUMNAS_BASE = [
    "País", "Norma", "Artículo", "Tipo penal", "Categoría", "Estado",
    "Contexto típico de protesta", "Legalidad", "Claridad normativa",
    "Lesividad", "Mínima intervención penal", "Idoneidad", "Necesidad penal",
    "Peso derechos indígenas", "Peso interés estatal", "Observación metodológica"
]

NUMERIC_COLS = [
    "Legalidad", "Claridad normativa", "Lesividad", "Mínima intervención penal",
    "Idoneidad", "Necesidad penal", "Peso derechos indígenas", "Peso interés estatal"
]

RUBRICA_COLS = [
    "Modulo", "Variable", "Codigo", "Subcriterio", "Peso_sobre_variable",
    "Valor_editable_0_100", "Aporte_calculado", "Formula_de_aporte",
    "Formula_de_variable", "Explicacion_metodologica", "Puntaje_alto_75_100",
    "Puntaje_medio_40_74", "Puntaje_bajo_0_39", "Fuente_metodologica_base"
]

VAR_CANONICA = {
    "Legalidad / tipicidad estricta": "Legalidad",
    "Legalidad": "Legalidad",
    "Claridad normativa": "Claridad normativa",
    "Lesividad": "Lesividad",
    "Mínima intervención penal": "Mínima intervención penal",
    "Minima intervencion penal": "Mínima intervención penal",
    "Idoneidad": "Idoneidad",
    "Necesidad penal": "Necesidad penal",
    "Peso derechos indígenas": "Peso derechos indígenas",
    "Peso derechos indigenas": "Peso derechos indígenas",
    "Peso interés estatal": "Peso interés estatal",
    "Peso interes estatal": "Peso interés estatal",
}

ORDEN_VARIABLES = [
    "Legalidad / tipicidad estricta",
    "Claridad normativa",
    "Lesividad",
    "Mínima intervención penal",
    "Idoneidad",
    "Necesidad penal",
    "Peso derechos indígenas",
    "Peso interés estatal",
]

def normalizar_base(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in COLUMNAS_BASE:
        if c not in df.columns:
            df[c] = 0 if c in NUMERIC_COLS else ""
    df = df[COLUMNAS_BASE]
    for c in NUMERIC_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).clip(0, 100)
    for c in [x for x in COLUMNAS_BASE if x not in NUMERIC_COLS]:
        df[c] = df[c].fillna("").astype(str)
    return df

def normalizar_rubrica(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in RUBRICA_COLS:
        if c not in df.columns:
            df[c] = 0 if c in ["Peso_sobre_variable", "Valor_editable_0_100", "Aporte_calculado"] else ""
    df = df[RUBRICA_COLS]
    df["Peso_sobre_variable"] = pd.to_numeric(df["Peso_sobre_variable"], errors="coerce").fillna(0)
    df["Valor_editable_0_100"] = pd.to_numeric(df["Valor_editable_0_100"], errors="coerce").fillna(0).clip(0, 100)
    df["Aporte_calculado"] = (df["Valor_editable_0_100"] * df["Peso_sobre_variable"]).round(2)
    for c in [x for x in RUBRICA_COLS if x not in ["Peso_sobre_variable", "Valor_editable_0_100", "Aporte_calculado"]]:
        df[c] = df[c].fillna("").astype(str)
    return df

def cargar_base_default() -> pd.DataFrame:
    if BASE_PATH.exists():
        return normalizar_base(pd.read_csv(BASE_PATH))
    return normalizar_base(pd.DataFrame(columns=COLUMNAS_BASE))

def cargar_rubrica_default() -> pd.DataFrame:
    if SUBCRITERIOS_PATH.exists():
        return normalizar_rubrica(pd.read_excel(SUBCRITERIOS_PATH, sheet_name=0))
    return normalizar_rubrica(pd.DataFrame(columns=RUBRICA_COLS))

def nivel_compatibilidad(valor):
    valor = float(valor)
    if valor >= 75:
        return "Compatibilidad alta", "Verde"
    if valor >= 60:
        return "Compatibilidad condicionada", "Amarillo"
    if valor >= 40:
        return "Compatibilidad débil", "Naranja"
    return "Incompatibilidad", "Rojo"

def nivel_riesgo(valor):
    valor = float(valor)
    if valor >= 70:
        return "Riesgo alto", "Rojo"
    if valor >= 50:
        return "Riesgo medio-alto", "Naranja"
    if valor >= 35:
        return "Riesgo medio", "Amarillo"
    return "Riesgo bajo", "Verde"

def calcular_indice_desde_variables(valores: dict) -> dict:
    legalidad = float(valores.get("Legalidad", 0))
    claridad = float(valores.get("Claridad normativa", 0))
    lesividad = float(valores.get("Lesividad", 0))
    minima = float(valores.get("Mínima intervención penal", 0))
    idoneidad = float(valores.get("Idoneidad", 0))
    necesidad = float(valores.get("Necesidad penal", 0))
    pdi = float(valores.get("Peso derechos indígenas", 0))
    pie = float(valores.get("Peso interés estatal", 0))

    formal = round(legalidad * 0.60 + claridad * 0.40, 2)
    material = round(lesividad * 0.60 + minima * 0.40, 2)
    europeo = round(idoneidad * 0.50 + necesidad * 0.50, 2)
    alexy = round((pie / (pdi + pie)) * 100, 2) if (pdi + pie) else 0
    peso_relativo = round(pdi / pie, 2) if pie else 999
    cpi = round((formal * 0.25) + (material * 0.30) + (europeo * 0.25) + (alexy * 0.20), 2)
    ircp = round(100 - cpi, 2)
    nivel, sem = nivel_riesgo(ircp)

    return {
        "Legalidad": legalidad,
        "Claridad normativa": claridad,
        "Lesividad": lesividad,
        "Mínima intervención penal": minima,
        "Idoneidad": idoneidad,
        "Necesidad penal": necesidad,
        "Peso derechos indígenas": pdi,
        "Peso interés estatal": pie,
        "Formal": formal,
        "Material": material,
        "Europeo": europeo,
        "Alexy": alexy,
        "Peso relativo": peso_relativo,
        "CPI": cpi,
        "IRCP-I": ircp,
        "Nivel": nivel,
        "Semáforo": sem,
    }

def calcular_fila(row) -> dict:
    valores = {c: row[c] for c in NUMERIC_COLS}
    res = calcular_indice_desde_variables(valores)
    return {
        "País": row["País"],
        "Artículo": row["Artículo"],
        "Tipo penal": row["Tipo penal"],
        "Categoría": row["Categoría"],
        "Estado": row["Estado"],
        "Formal": res["Formal"],
        "Material": res["Material"],
        "Europeo": res["Europeo"],
        "Peso relativo": res["Peso relativo"],
        "Alexy": res["Alexy"],
        "CPI": res["CPI"],
        "IRCP-I": res["IRCP-I"],
        "Nivel": res["Nivel"],
        "Semáforo": res["Semáforo"],
    }

def base_resultados(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return pd.DataFrame([calcular_fila(row) for _, row in df.iterrows()])

def to_excel_base_bytes(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        normalizar_base(df).to_excel(writer, index=False, sheet_name="Base editable")
        base_resultados(normalizar_base(df)).to_excel(writer, index=False, sheet_name="Resultados")
    bio.seek(0)
    return bio.getvalue()

def to_excel_rubrica_bytes(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        normalizar_rubrica(df).to_excel(writer, index=False, sheet_name="Subcriterios_App")
    bio.seek(0)
    return bio.getvalue()

def generar_word(resumen, matrices):
    if Document is None:
        return None
    doc = Document()
    doc.add_heading("Informe MCPI-IRCP-I", 0)
    for k in ["País", "Artículo", "Tipo penal", "Contexto"]:
        doc.add_paragraph(f"{k}: {resumen.get(k, '')}")
    doc.add_heading("Resultado global", level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Indicador"
    table.rows[0].cells[1].text = "Valor"
    for k in ["Formal", "Material", "Europeo", "Peso relativo", "Alexy", "CPI", "IRCP-I", "Nivel", "Semáforo"]:
        cells = table.add_row().cells
        cells[0].text = k
        cells[1].text = str(resumen.get(k, ""))
    for nombre, df in matrices.items():
        doc.add_heading(nombre, level=1)
        cols = [c for c in ["Variable", "Codigo", "Subcriterio", "Valor", "Peso_sobre_variable", "Aporte", "Justificación"] if c in df.columns]
        if not cols:
            cols = [c for c in df.columns[:5]]
        t = doc.add_table(rows=1, cols=len(cols))
        t.style = "Table Grid"
        for i, h in enumerate(cols):
            t.rows[0].cells[i].text = h
        for _, row in df.iterrows():
            cells = t.add_row().cells
            for i, h in enumerate(cols):
                cells[i].text = str(row.get(h, ""))
    doc.add_heading("Conclusión", level=1)
    doc.add_paragraph(resumen.get("Conclusión", ""))
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def variable_desde_rubrica(variable_rubrica: str, row: dict, key: str):
    rub = st.session_state.rubrica_actual.copy()
    sub = rub[rub["Variable"] == variable_rubrica].copy()

    if sub.empty:
        st.warning(f"No existen subcriterios para: {variable_rubrica}")
        canon = VAR_CANONICA.get(variable_rubrica, variable_rubrica)
        return float(row.get(canon, 0)), pd.DataFrame()

    canon = VAR_CANONICA.get(variable_rubrica, variable_rubrica)
    valor_base = float(row.get(canon, 0))

    sub["Valor"] = sub["Valor_editable_0_100"]
    if float(sub["Valor"].sum()) == 0:
        sub["Valor"] = valor_base

    mostrar = sub[[
        "Codigo", "Subcriterio", "Peso_sobre_variable", "Valor",
        "Explicacion_metodologica", "Puntaje_alto_75_100",
        "Puntaje_medio_40_74", "Puntaje_bajo_0_39"
    ]].copy()

    edited = st.data_editor(
        mostrar,
        key=key,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Codigo": st.column_config.TextColumn("Código", disabled=True),
            "Subcriterio": st.column_config.TextColumn("Subcriterio", disabled=True, width="medium"),
            "Peso_sobre_variable": st.column_config.NumberColumn("Peso", disabled=True, format="%.2f"),
            "Valor": st.column_config.NumberColumn("Valor 0-100", min_value=0, max_value=100, step=1, format="%.0f"),
            "Explicacion_metodologica": st.column_config.TextColumn("Explicación metodológica", disabled=True, width="large"),
            "Puntaje_alto_75_100": st.column_config.TextColumn("Alto 75-100", disabled=True, width="medium"),
            "Puntaje_medio_40_74": st.column_config.TextColumn("Medio 40-74", disabled=True, width="medium"),
            "Puntaje_bajo_0_39": st.column_config.TextColumn("Bajo 0-39", disabled=True, width="medium"),
        },
        disabled=[
            "Codigo", "Subcriterio", "Peso_sobre_variable", "Explicacion_metodologica",
            "Puntaje_alto_75_100", "Puntaje_medio_40_74", "Puntaje_bajo_0_39"
        ],
    )

    edited["Valor"] = pd.to_numeric(edited["Valor"], errors="coerce").fillna(0).clip(0, 100)
    edited["Peso_sobre_variable"] = pd.to_numeric(edited["Peso_sobre_variable"], errors="coerce").fillna(0)
    edited["Aporte"] = (edited["Valor"] * edited["Peso_sobre_variable"]).round(2)
    total = round(float(edited["Aporte"].sum()), 2)

    st.dataframe(
        edited[["Codigo", "Subcriterio", "Valor", "Peso_sobre_variable", "Aporte"]],
        hide_index=True,
        use_container_width=True,
    )

    res, sem = nivel_compatibilidad(total)
    st.success(f"{canon}: {total} — {res} / {sem}")

    edited.insert(0, "Variable", canon)
    return total, edited

def render_simple_variable_table(row: dict):
    data = pd.DataFrame([
        {"Variable": "Legalidad", "Peso dentro del control": 0.60, "Puntaje": row["Legalidad"], "Módulo": "Control formal"},
        {"Variable": "Claridad normativa", "Peso dentro del control": 0.40, "Puntaje": row["Claridad normativa"], "Módulo": "Control formal"},
        {"Variable": "Lesividad", "Peso dentro del control": 0.60, "Puntaje": row["Lesividad"], "Módulo": "Control material"},
        {"Variable": "Mínima intervención penal", "Peso dentro del control": 0.40, "Puntaje": row["Mínima intervención penal"], "Módulo": "Control material"},
        {"Variable": "Idoneidad", "Peso dentro del control": 0.50, "Puntaje": row["Idoneidad"], "Módulo": "Test europeo"},
        {"Variable": "Necesidad penal", "Peso dentro del control": 0.50, "Puntaje": row["Necesidad penal"], "Módulo": "Test europeo"},
        {"Variable": "Peso derechos indígenas", "Peso dentro del control": 1.00, "Puntaje": row["Peso derechos indígenas"], "Módulo": "Alexy"},
        {"Variable": "Peso interés estatal", "Peso dentro del control": 1.00, "Puntaje": row["Peso interés estatal"], "Módulo": "Alexy"},
    ])
    edited = st.data_editor(
        data,
        key="variables_simples",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Variable": st.column_config.TextColumn(disabled=True),
            "Módulo": st.column_config.TextColumn(disabled=True),
            "Peso dentro del control": st.column_config.NumberColumn(disabled=True, format="%.2f"),
            "Puntaje": st.column_config.NumberColumn(min_value=0, max_value=100, step=1, format="%.0f"),
        },
        disabled=["Variable", "Módulo", "Peso dentro del control"],
    )
    valores = {r["Variable"]: float(r["Puntaje"]) for _, r in edited.iterrows()}
    return valores, edited

def reglas_de_cierre(res: dict):
    reglas = []
    if res["Legalidad"] < 40:
        reglas.append("Legalidad menor a 40: presunción de incompatibilidad formal.")
    if res["Lesividad"] < 40:
        reglas.append("Lesividad menor a 40: no supera el control material.")
    if res["Idoneidad"] < 75:
        reglas.append("Idoneidad menor a 75: la criminalización penal no queda justificada.")
    if res["Necesidad penal"] < 60:
        reglas.append("Necesidad penal menor a 60: debe preferirse una alternativa no penal.")
    if res["Peso relativo"] > 1:
        reglas.append("Alexy favorece derechos indígenas: exige motivación reforzada.")
    if res["IRCP-I"] >= 70:
        reglas.append("IRCP-I alto: riesgo severo de criminalización penal intercultural.")
    elif res["IRCP-I"] >= 50:
        reglas.append("IRCP-I medio-alto: riesgo relevante de criminalización penal intercultural.")
    return reglas

if "base_actual" not in st.session_state:
    st.session_state.base_actual = cargar_base_default()

if "rubrica_actual" not in st.session_state:
    st.session_state.rubrica_actual = cargar_rubrica_default()

st.title("Matriz MCPI-IRCP-I v6")
st.caption("Versión con subcriterios: calcula variables desde rúbricas editables, agrega delitos, importa/exporta base y recalcula el índice global.")

menu = st.sidebar.radio(
    "Módulos",
    [
        "Evaluación en matriz",
        "Rúbrica de subcriterios",
        "Base editable y ampliada",
        "Añadir delito",
        "Importar / Exportar",
        "Dashboard comparativo",
        "Metodología y reglas",
    ],
)

base = st.session_state.base_actual

if menu == "Evaluación en matriz":
    st.header("Evaluación en matriz")

    if base.empty:
        st.error("La base está vacía. Ingrese a 'Añadir delito' o importe un CSV.")
        st.stop()

    opciones = [
        f"{i} | {r['País']} | {r['Tipo penal']} | {r['Artículo']}"
        for i, r in base.iterrows()
    ]
    default_idx = next((i for i, x in enumerate(opciones) if "194" in x), 0)
    seleccion = st.selectbox("Seleccione delito/tipo penal", opciones, index=default_idx)
    idx = int(seleccion.split("|")[0].strip())
    row = base.loc[idx].to_dict()

    st.subheader("Datos del caso")
    c1, c2, c3 = st.columns(3)
    with c1:
        pais = st.text_input("País", row["País"])
    with c2:
        norma = st.text_input("Norma", row["Norma"])
    with c3:
        articulo = st.text_input("Artículo", row["Artículo"])
    tipo = st.text_input("Tipo penal", row["Tipo penal"])
    contexto = st.text_area("Contexto típico de protesta", row["Contexto típico de protesta"], height=80)
    observacion = st.text_area("Observación metodológica", row["Observación metodológica"], height=80)

    modo = st.radio(
        "Modo de cálculo",
        ["Calcular desde subcriterios", "Editar variables principales"],
        horizontal=True,
    )

    matrices_word = {}

    if modo == "Calcular desde subcriterios":
        st.info("Edite los valores 0-100 de cada subcriterio. La app calcula automáticamente la variable y luego el índice MCPI-IRCP-I.")

        valores = {}
        matriz_subcriterios = []

        grupos = {
            "1. Control formal": ["Legalidad / tipicidad estricta", "Claridad normativa"],
            "2. Control material": ["Lesividad", "Mínima intervención penal"],
            "3. Test europeo de proporcionalidad": ["Idoneidad", "Necesidad penal"],
            "4. Ponderación de Alexy": ["Peso derechos indígenas", "Peso interés estatal"],
        }

        for titulo, variables in grupos.items():
            st.subheader(titulo)
            for variable in variables:
                with st.expander(f"{variable}", expanded=(variable in ["Legalidad / tipicidad estricta", "Claridad normativa"])):
                    total, edited = variable_desde_rubrica(variable, row, f"sub_{variable}")
                    canon = VAR_CANONICA.get(variable, variable)
                    valores[canon] = total
                    if not edited.empty:
                        matriz_subcriterios.append(edited)
                        matrices_word[canon] = edited

    else:
        st.info("Edite directamente las variables principales. Este modo es más rápido, pero no muestra el detalle de subcriterios.")
        valores, edited_vars = render_simple_variable_table(row)
        matrices_word["Variables principales"] = edited_vars

    res = calcular_indice_desde_variables(valores)

    st.subheader("5. Índice global de criminalización penal intercultural")

    componentes = pd.DataFrame([
        {"Componente": "Control formal", "Resultado": res["Formal"], "Peso": 0.25, "Aporte": round(res["Formal"] * 0.25, 2)},
        {"Componente": "Control material", "Resultado": res["Material"], "Peso": 0.30, "Aporte": round(res["Material"] * 0.30, 2)},
        {"Componente": "Test europeo", "Resultado": res["Europeo"], "Peso": 0.25, "Aporte": round(res["Europeo"] * 0.25, 2)},
        {"Componente": "Puntaje Alexy", "Resultado": res["Alexy"], "Peso": 0.20, "Aporte": round(res["Alexy"] * 0.20, 2)},
    ])
    st.dataframe(componentes, use_container_width=True, hide_index=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Compatibilidad Penal Intercultural (CPI)", res["CPI"])
    m2.metric("IRCP-I", res["IRCP-I"])
    m3.metric("Nivel / semáforo", f"{res['Nivel']} / {res['Semáforo']}")

    st.code(
        f"Formal = ({res['Legalidad']} × 0,60) + ({res['Claridad normativa']} × 0,40) = {res['Formal']}\n"
        f"Material = ({res['Lesividad']} × 0,60) + ({res['Mínima intervención penal']} × 0,40) = {res['Material']}\n"
        f"Europeo = ({res['Idoneidad']} × 0,50) + ({res['Necesidad penal']} × 0,50) = {res['Europeo']}\n"
        f"Alexy = [{res['Peso interés estatal']} ÷ ({res['Peso derechos indígenas']} + {res['Peso interés estatal']})] × 100 = {res['Alexy']}\n"
        f"CPI = ({res['Formal']} × 0,25) + ({res['Material']} × 0,30) + ({res['Europeo']} × 0,25) + ({res['Alexy']} × 0,20) = {res['CPI']}\n"
        f"IRCP-I = 100 − {res['CPI']} = {res['IRCP-I']}",
        language="text"
    )

    st.subheader("Reglas de cierre")
    reglas = reglas_de_cierre(res)
    if reglas:
        for r in reglas:
            st.warning(r)
    else:
        st.success("No se activan reglas críticas de cierre.")

    conclusion = (
        f"El tipo penal evaluado ({tipo}, {articulo}, {pais}) presenta una Compatibilidad Penal Intercultural "
        f"de {res['CPI']} y un Índice de Riesgo de Criminalización Penal Intercultural de {res['IRCP-I']}, "
        f"equivalente a {res['Nivel']} / semáforo {res['Semáforo']}. En el contexto analizado, la aplicación penal "
        f"exige control de convencionalidad reforzado, interpretación restrictiva, prueba individualizada, daño concreto, "
        f"riesgo real y ausencia de alternativas menos lesivas."
    )
    st.subheader("Conclusión automática")
    st.write(conclusion)

    chart_df = pd.DataFrame({
        "Componente": ["Formal", "Material", "Europeo", "Alexy", "CPI", "IRCP-I"],
        "Valor": [res["Formal"], res["Material"], res["Europeo"], res["Alexy"], res["CPI"], res["IRCP-I"]],
    })
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.bar(chart_df["Componente"], chart_df["Valor"])
    ax.set_ylim(0, 100)
    ax.set_ylabel("Puntaje")
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    st.pyplot(fig)

    word = generar_word(
        {
            "País": pais, "Artículo": articulo, "Tipo penal": tipo, "Contexto": contexto,
            "Formal": res["Formal"], "Material": res["Material"], "Europeo": res["Europeo"],
            "Peso relativo": res["Peso relativo"], "Alexy": res["Alexy"], "CPI": res["CPI"],
            "IRCP-I": res["IRCP-I"], "Nivel": res["Nivel"], "Semáforo": res["Semáforo"],
            "Conclusión": conclusion,
        },
        matrices_word
    )
    if word:
        st.download_button(
            "Descargar informe Word",
            data=word,
            file_name="informe_MCPI_IRCP.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

elif menu == "Rúbrica de subcriterios":
    st.header("Rúbrica de subcriterios")
    st.info("Esta tabla define los subcriterios, pesos y explicación metodológica. Puede editarla en sesión y descargarla. Para dejarla permanente, suba el Excel actualizado a GitHub.")

    rubrica = st.data_editor(
        st.session_state.rubrica_actual,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Peso_sobre_variable": st.column_config.NumberColumn("Peso", min_value=0.0, max_value=1.0, step=0.01, format="%.2f"),
            "Valor_editable_0_100": st.column_config.NumberColumn("Valor base 0-100", min_value=0, max_value=100, step=1),
            "Aporte_calculado": st.column_config.NumberColumn("Aporte", disabled=True, format="%.2f"),
        },
    )

    rubrica_norm = normalizar_rubrica(rubrica)

    if st.button("Aplicar cambios de rúbrica a la sesión"):
        st.session_state.rubrica_actual = rubrica_norm
        st.success("Rúbrica actualizada en la sesión.")

    st.download_button(
        "Descargar rúbrica Excel",
        data=to_excel_rubrica_bytes(rubrica_norm),
        file_name="subcriterios_MCPI_IRCP_app_unico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        "Descargar rúbrica CSV",
        data=rubrica_norm.to_csv(index=False).encode("utf-8-sig"),
        file_name="subcriterios_MCPI_IRCP_app_unico.csv",
        mime="text/csv",
    )

elif menu == "Base editable y ampliada":
    st.header("Base editable y ampliada")
    st.info("Edite valores o agregue filas. Para conservar cambios, descargue CSV/Excel y súbalo a GitHub reemplazando el archivo base.")
    edited = st.data_editor(
        st.session_state.base_actual,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={c: st.column_config.NumberColumn(c, min_value=0, max_value=100, step=1) for c in NUMERIC_COLS},
    )
    edited_norm = normalizar_base(edited)
    if st.button("Aplicar cambios a la sesión"):
        st.session_state.base_actual = edited_norm
        st.success("Base actualizada en la sesión.")

    st.download_button(
        "Descargar base CSV",
        data=edited_norm.to_csv(index=False).encode("utf-8-sig"),
        file_name="tipos_penales_base_ampliada.csv",
        mime="text/csv",
    )
    st.download_button(
        "Descargar base Excel",
        data=to_excel_base_bytes(edited_norm),
        file_name="tipos_penales_base_ampliada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

elif menu == "Añadir delito":
    st.header("Añadir nuevo delito/tipo penal")
    with st.form("nuevo_delito"):
        c1, c2, c3 = st.columns(3)
        with c1:
            pais = st.selectbox("País", ["Argentina", "Ecuador", "Otro"])
            norma = st.text_input("Norma", "Código Penal" if pais == "Argentina" else "COIP")
        with c2:
            articulo = st.text_input("Artículo", "")
            categoria = st.text_input("Categoría", "Orden público")
        with c3:
            estado = st.selectbox("Estado", ["Ampliado preliminar", "Base tesis", "Validado por investigador"])
            tipo = st.text_input("Tipo penal", "")
        contexto = st.text_area("Contexto típico de protesta", "")
        obs = st.text_area("Observación metodológica", "")
        st.subheader("Puntajes iniciales")
        p1, p2, p3, p4 = st.columns(4)
        with p1:
            legalidad = st.number_input("Legalidad", 0, 100, 50)
            claridad = st.number_input("Claridad normativa", 0, 100, 45)
        with p2:
            lesividad = st.number_input("Lesividad", 0, 100, 45)
            minima = st.number_input("Mínima intervención penal", 0, 100, 35)
        with p3:
            idoneidad = st.number_input("Idoneidad", 0, 100, 50)
            necesidad = st.number_input("Necesidad penal", 0, 100, 35)
        with p4:
            pdi = st.number_input("Peso derechos indígenas", 0, 100, 40)
            pie = st.number_input("Peso interés estatal", 0, 100, 24)
        submitted = st.form_submit_button("Agregar a la base")

    if submitted:
        nueva = pd.DataFrame([{
            "País": pais,
            "Norma": norma,
            "Artículo": articulo,
            "Tipo penal": tipo,
            "Categoría": categoria,
            "Estado": estado,
            "Contexto típico de protesta": contexto,
            "Legalidad": legalidad,
            "Claridad normativa": claridad,
            "Lesividad": lesividad,
            "Mínima intervención penal": minima,
            "Idoneidad": idoneidad,
            "Necesidad penal": necesidad,
            "Peso derechos indígenas": pdi,
            "Peso interés estatal": pie,
            "Observación metodológica": obs,
        }])
        st.session_state.base_actual = normalizar_base(pd.concat([st.session_state.base_actual, nueva], ignore_index=True))
        st.success("Delito agregado a la base de la sesión.")

elif menu == "Importar / Exportar":
    st.header("Importar / Exportar")
    st.subheader("Base de tipos penales")
    up = st.file_uploader("Importar base CSV o Excel", type=["csv", "xlsx"], key="upload_base")
    modo = st.radio("Modo de importación de base", ["Reemplazar base actual", "Agregar a base actual"], horizontal=True)
    if up is not None:
        try:
            nueva = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            nueva = normalizar_base(nueva)
            st.dataframe(nueva, hide_index=True, use_container_width=True)
            if st.button("Confirmar importación de base"):
                if modo == "Reemplazar base actual":
                    st.session_state.base_actual = nueva
                else:
                    st.session_state.base_actual = normalizar_base(pd.concat([st.session_state.base_actual, nueva], ignore_index=True))
                st.success("Importación de base aplicada.")
        except Exception as e:
            st.error(f"No se pudo importar la base: {e}")

    st.download_button(
        "Descargar CSV actual de base",
        data=st.session_state.base_actual.to_csv(index=False).encode("utf-8-sig"),
        file_name="tipos_penales_base_actual.csv",
        mime="text/csv",
    )
    st.download_button(
        "Descargar Excel actual de base",
        data=to_excel_base_bytes(st.session_state.base_actual),
        file_name="tipos_penales_base_actual.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("Rúbrica de subcriterios")
    up_r = st.file_uploader("Importar rúbrica Excel", type=["xlsx"], key="upload_rubrica")
    if up_r is not None:
        try:
            nueva_r = normalizar_rubrica(pd.read_excel(up_r, sheet_name=0))
            st.dataframe(nueva_r, hide_index=True, use_container_width=True)
            if st.button("Confirmar importación de rúbrica"):
                st.session_state.rubrica_actual = nueva_r
                st.success("Rúbrica importada a la sesión.")
        except Exception as e:
            st.error(f"No se pudo importar la rúbrica: {e}")

    st.download_button(
        "Descargar rúbrica actual Excel",
        data=to_excel_rubrica_bytes(st.session_state.rubrica_actual),
        file_name="subcriterios_MCPI_IRCP_app_unico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

elif menu == "Dashboard comparativo":
    st.header("Dashboard comparativo")
    resultados = base_resultados(st.session_state.base_actual)
    if resultados.empty:
        st.warning("No hay datos para mostrar.")
        st.stop()
    st.dataframe(resultados, hide_index=True, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Delitos en base", len(resultados))
    c2.metric("IRCP-I promedio", round(resultados["IRCP-I"].mean(), 2))
    c3.metric("Mayor riesgo", resultados.loc[resultados["IRCP-I"].idxmax(), "Tipo penal"])

    st.subheader("IRCP-I por tipo penal")
    plot = resultados.sort_values("IRCP-I")
    fig, ax = plt.subplots(figsize=(11, max(5, len(plot) * 0.28)))
    ax.barh(plot["País"] + " - " + plot["Tipo penal"], plot["IRCP-I"])
    ax.set_xlim(0, 100)
    ax.set_xlabel("IRCP-I")
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    st.pyplot(fig)

    st.subheader("Promedio por país")
    prom = resultados.groupby("País", as_index=False)["IRCP-I"].mean()
    fig2, ax2 = plt.subplots(figsize=(7, 4))
    ax2.bar(prom["País"], prom["IRCP-I"])
    ax2.set_ylim(0, 100)
    ax2.set_ylabel("IRCP-I promedio")
    ax2.grid(axis="y", linestyle="--", alpha=0.4)
    st.pyplot(fig2)

else:
    st.header("Metodología y reglas")
    st.markdown("""
### Lógica de cálculo

La app ahora permite dos niveles de cálculo:

1. **Subcriterios**: cada variable se calcula como suma ponderada de sus subcriterios.
2. **Variables principales**: cada variable puede editarse directamente si se requiere un cálculo rápido.

### Fórmulas de variables

**Legalidad** = SUMA(LEG-i × Peso)  
**Claridad normativa** = SUMA(CLA-i × Peso)  
**Lesividad** = SUMA(LES-i × Peso)  
**Mínima intervención penal** = SUMA(MIN-i × Peso)  
**Idoneidad** = SUMA(IDO-i × Peso)  
**Necesidad penal** = SUMA(NEC-i × Peso)  
**Peso derechos indígenas** = SUMA(PDI-i × Peso)  
**Peso interés estatal** = SUMA(PIE-i × Peso)

### Fórmulas principales

**Control formal** = (Legalidad × 0,60) + (Claridad normativa × 0,40)

**Control material** = (Lesividad × 0,60) + (Mínima intervención penal × 0,40)

**Test europeo** = (Idoneidad × 0,50) + (Necesidad penal × 0,50)

**Puntaje Alexy** = Peso interés estatal / (Peso derechos indígenas + Peso interés estatal) × 100

**CPI** = (Formal × 0,25) + (Material × 0,30) + (Europeo × 0,25) + (Alexy × 0,20)

**IRCP-I** = 100 − CPI

### Reglas de cierre

- Legalidad menor a 40: presunción de incompatibilidad formal.
- Lesividad menor a 40: no supera control material.
- Idoneidad menor a 75: la criminalización penal no queda justificada.
- Necesidad penal menor a 60: debe preferirse alternativa no penal.
- Peso relativo de Alexy mayor a 1: prevalecen derechos indígenas y se exige motivación reforzada.
- IRCP-I mayor o igual a 70: riesgo alto.

### Persistencia en Streamlit Cloud

Los cambios realizados por usuarios quedan en la sesión del navegador. Para conservarlos, descargue la base CSV/Excel o la rúbrica Excel. Para hacerlos permanentes en la app pública, suba el archivo actualizado al repositorio de GitHub reemplazando el archivo correspondiente.
""")

st.caption("Herramienta académica orientativa. No sustituye análisis jurídico definitivo ni control judicial del caso concreto.")
