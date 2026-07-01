from pathlib import Path
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from io import BytesIO

try:
    from docx import Document
except Exception:
    Document = None

st.set_page_config(
    page_title="Matriz MCPI-IRCP-I",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).parent / "tipos_penales_base.csv"

@st.cache_data
def cargar_base():
    return pd.read_csv(BASE)

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

def calc_ponderado(df):
    df = df.copy()
    df["Puntaje"] = pd.to_numeric(df["Puntaje"], errors="coerce").fillna(0)
    df["Peso"] = pd.to_numeric(df["Peso"], errors="coerce").fillna(0)
    df["Resultado ponderado"] = (df["Puntaje"] * df["Peso"]).round(2)
    total = round(float(df["Resultado ponderado"].sum()), 2)
    return df, total

def editor_matriz(titulo, data, key):
    st.subheader(titulo)
    df = pd.DataFrame(data)
    edited = st.data_editor(
        df,
        key=key,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Variable": st.column_config.TextColumn("Variable", disabled=True),
            "Descripción": st.column_config.TextColumn("Descripción", disabled=True),
            "Peso": st.column_config.NumberColumn("Peso", disabled=True, format="%.2f"),
            "Puntaje": st.column_config.NumberColumn("Puntaje", min_value=0, max_value=100, step=1, format="%.0f"),
        },
        disabled=["Variable", "Descripción", "Peso"],
    )
    calculated, total = calc_ponderado(edited)
    st.dataframe(calculated[["Variable", "Puntaje", "Peso", "Resultado ponderado"]], use_container_width=True, hide_index=True)
    res, sem = nivel_compatibilidad(total)
    st.success(f"Resultado de {titulo}: {total} — {res} / {sem}")
    return calculated, total, res, sem

def generar_word(resultado):
    if Document is None:
        return None
    doc = Document()
    doc.add_heading("Informe MCPI-IRCP-I", 0)
    doc.add_paragraph(f"País: {resultado['pais']}")
    doc.add_paragraph(f"Tipo penal: {resultado['tipo_penal']}")
    doc.add_paragraph(f"Artículo: {resultado['articulo']}")
    doc.add_paragraph(f"Contexto: {resultado['contexto']}")

    doc.add_heading("Resultado final", level=1)
    tabla = doc.add_table(rows=1, cols=2)
    tabla.style = "Table Grid"
    tabla.rows[0].cells[0].text = "Indicador"
    tabla.rows[0].cells[1].text = "Valor"
    for k, v in [
        ("Control formal", resultado["formal"]),
        ("Control material", resultado["material"]),
        ("Test europeo", resultado["europeo"]),
        ("Puntaje Alexy", resultado["alexy"]),
        ("Compatibilidad Penal Intercultural", resultado["cpi"]),
        ("IRCP-I", resultado["ircp"]),
        ("Nivel de riesgo", resultado["nivel_riesgo"]),
        ("Semáforo", resultado["semaforo"]),
    ]:
        row = tabla.add_row().cells
        row[0].text = str(k)
        row[1].text = str(v)

    doc.add_heading("Conclusión", level=1)
    doc.add_paragraph(resultado["conclusion"])

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

base = cargar_base()

st.title("Matriz MCPI-IRCP-I")
st.caption("Versión con celdas/matrices desplegadas para control formal, control material, test europeo, Alexy e índice global.")

menu = st.sidebar.radio(
    "Módulos",
    ["Evaluación en matriz", "Dashboard comparativo", "Base de tipos penales", "Metodología"],
)

if menu == "Evaluación en matriz":
    st.header("1. Datos del tipo penal")

    opciones = [f"{r['País']} | {r['Tipo penal']} | {r['Artículo']}" for _, r in base.iterrows()]
    default_idx = next((i for i, x in enumerate(opciones) if "194" in x), 0)
    seleccion = st.selectbox("Seleccione el tipo penal", opciones, index=default_idx)
    row = base.iloc[opciones.index(seleccion)].to_dict()

    c1, c2, c3 = st.columns(3)
    with c1:
        pais = st.text_input("País", row["País"])
    with c2:
        norma = st.text_input("Norma", row["Norma"])
    with c3:
        articulo = st.text_input("Artículo", row["Artículo"])

    tipo_penal = st.text_input("Tipo penal", row["Tipo penal"])
    contexto = st.text_area("Contexto", row["Contexto"], height=80)

    st.header("2. Matrices de evaluación")

    formal_df, formal, formal_res, formal_sem = editor_matriz(
        "Control formal",
        [
            {"Variable": "Legalidad", "Descripción": "Ley previa, precisión de conducta, sujetos, elementos del tipo y prohibición de analogía.", "Peso": 0.60, "Puntaje": float(row["Legalidad"])},
            {"Variable": "Claridad normativa", "Descripción": "Comprensibilidad, verbos rectores, bien jurídico, umbral de gravedad y salvaguardas.", "Peso": 0.40, "Puntaje": float(row["Claridad normativa"])},
        ],
        "formal_editor"
    )

    material_df, material, material_res, material_sem = editor_matriz(
        "Control material",
        [
            {"Variable": "Lesividad", "Descripción": "Daño real, gravedad, riesgo concreto y prueba individualizada.", "Peso": 0.60, "Puntaje": float(row["Lesividad"])},
            {"Variable": "Mínima intervención penal", "Descripción": "Alternativas no penales, diálogo, consulta, mediación y última ratio.", "Peso": 0.40, "Puntaje": float(row["Mínima intervención penal"])},
        ],
        "material_editor"
    )

    europeo_df, europeo, europeo_res, europeo_sem = editor_matriz(
        "Test europeo de proporcionalidad",
        [
            {"Variable": "Idoneidad", "Descripción": "Conexión real entre persecución penal y finalidad constitucional legítima.", "Peso": 0.50, "Puntaje": float(row["Idoneidad"])},
            {"Variable": "Necesidad penal", "Descripción": "Inexistencia de alternativas menos lesivas y necesidad estricta de vía penal.", "Peso": 0.50, "Puntaje": float(row["Necesidad penal"])},
        ],
        "europeo_editor"
    )

    st.subheader("Ponderación de Alexy")
    alexy_df = pd.DataFrame([
        {"Variable": "Peso derechos indígenas", "Descripción": "Peso de protesta indígena, territorio, consulta previa, autodeterminación y participación.", "Valor": float(row["Peso derechos indígenas"])},
        {"Variable": "Peso interés estatal", "Descripción": "Peso del orden público, circulación, seguridad o prestación de servicios.", "Valor": float(row["Peso interés estatal"])},
    ])
    alexy_edit = st.data_editor(
        alexy_df,
        key="alexy_editor",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Variable": st.column_config.TextColumn("Variable", disabled=True),
            "Descripción": st.column_config.TextColumn("Descripción", disabled=True),
            "Valor": st.column_config.NumberColumn("Valor", min_value=0, max_value=100, step=1, format="%.0f"),
        },
        disabled=["Variable", "Descripción"],
    )
    peso_di = float(alexy_edit.loc[alexy_edit["Variable"] == "Peso derechos indígenas", "Valor"].iloc[0])
    peso_ie = float(alexy_edit.loc[alexy_edit["Variable"] == "Peso interés estatal", "Valor"].iloc[0])
    peso_relativo = round(peso_di / peso_ie, 2) if peso_ie else 999.0
    alexy = round((peso_ie / (peso_di + peso_ie)) * 100, 2) if (peso_di + peso_ie) else 0.0
    alexy_res, alexy_sem = nivel_compatibilidad(alexy)

    alexy_result = pd.DataFrame([
        {"Indicador": "Peso relativo", "Fórmula": "Derechos indígenas / interés estatal", "Resultado": peso_relativo},
        {"Indicador": "Puntaje Alexy", "Fórmula": "Interés estatal / (derechos indígenas + interés estatal) × 100", "Resultado": alexy},
    ])
    st.dataframe(alexy_result, use_container_width=True, hide_index=True)
    st.success(f"Resultado de Alexy: {alexy} — {alexy_res} / {alexy_sem}")

    st.header("3. Índice global de criminalización penal intercultural")

    cpi = round((formal * 0.25) + (material * 0.30) + (europeo * 0.25) + (alexy * 0.20), 2)
    ircp = round(100 - cpi, 2)
    nivel, semaforo = nivel_riesgo(ircp)

    indice_df = pd.DataFrame([
        {"Componente": "Control formal", "Resultado": formal, "Peso": 0.25, "Aporte": round(formal * 0.25, 2)},
        {"Componente": "Control material", "Resultado": material, "Peso": 0.30, "Aporte": round(material * 0.30, 2)},
        {"Componente": "Test europeo", "Resultado": europeo, "Peso": 0.25, "Aporte": round(europeo * 0.25, 2)},
        {"Componente": "Puntaje Alexy", "Resultado": alexy, "Peso": 0.20, "Aporte": round(alexy * 0.20, 2)},
    ])
    st.dataframe(indice_df, use_container_width=True, hide_index=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Compatibilidad Penal Intercultural (CPI)", cpi)
    m2.metric("IRCP-I", ircp)
    m3.metric("Nivel / semáforo", f"{nivel} / {semaforo}")

    st.code(
        f"CPI = ({formal} × 0,25) + ({material} × 0,30) + ({europeo} × 0,25) + ({alexy} × 0,20) = {cpi}\n"
        f"IRCP-I = 100 − {cpi} = {ircp}",
        language="text"
    )

    st.header("4. Conclusión automática")
    conclusion = (
        f"El tipo penal evaluado ({tipo_penal}, {articulo}, {pais}) presenta una Compatibilidad Penal "
        f"Intercultural de {cpi} y un Índice de Riesgo de Criminalización Penal Intercultural de {ircp}, "
        f"equivalente a {nivel} / semáforo {semaforo}. El resultado exige control de convencionalidad "
        f"reforzado, interpretación restrictiva y verificación de violencia grave, daño concreto, riesgo real, "
        f"prueba individualizada y ausencia de alternativas menos lesivas."
    )
    st.write(conclusion)

    st.header("5. Gráfico")
    chart_df = pd.DataFrame({
        "Componente": ["Formal", "Material", "Europeo", "Alexy", "CPI", "IRCP-I"],
        "Valor": [formal, material, europeo, alexy, cpi, ircp],
    })
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.bar(chart_df["Componente"], chart_df["Valor"])
    ax.set_ylim(0, 100)
    ax.set_ylabel("Puntaje")
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    st.pyplot(fig)

    resultado_word = {
        "pais": pais, "tipo_penal": tipo_penal, "articulo": articulo, "contexto": contexto,
        "formal": formal, "material": material, "europeo": europeo, "alexy": alexy,
        "cpi": cpi, "ircp": ircp, "nivel_riesgo": nivel, "semaforo": semaforo,
        "conclusion": conclusion,
    }
    word_bytes = generar_word(resultado_word)
    if word_bytes:
        st.download_button(
            "Descargar informe Word",
            data=word_bytes,
            file_name="informe_MCPI_IRCP_matriz.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

elif menu == "Dashboard comparativo":
    st.header("Dashboard comparativo")
    resultados = []
    for _, r in base.iterrows():
        formal = round(r["Legalidad"] * 0.60 + r["Claridad normativa"] * 0.40, 2)
        material = round(r["Lesividad"] * 0.60 + r["Mínima intervención penal"] * 0.40, 2)
        europeo = round(r["Idoneidad"] * 0.50 + r["Necesidad penal"] * 0.50, 2)
        alexy = round((r["Peso interés estatal"] / (r["Peso derechos indígenas"] + r["Peso interés estatal"])) * 100, 2)
        cpi = round((formal * 0.25) + (material * 0.30) + (europeo * 0.25) + (alexy * 0.20), 2)
        ircp = round(100 - cpi, 2)
        nivel, sem = nivel_riesgo(ircp)
        resultados.append({
            "País": r["País"], "Tipo penal": r["Tipo penal"], "Artículo": r["Artículo"],
            "Formal": formal, "Material": material, "Europeo": europeo, "Alexy": alexy,
            "CPI": cpi, "IRCP-I": ircp, "Nivel": nivel, "Semáforo": sem
        })
    res_df = pd.DataFrame(resultados)
    st.dataframe(res_df, use_container_width=True, hide_index=True)

    fig, ax = plt.subplots(figsize=(11, 6))
    plot = res_df.sort_values("IRCP-I")
    ax.barh(plot["País"] + " - " + plot["Tipo penal"], plot["IRCP-I"])
    ax.set_xlim(0, 100)
    ax.set_xlabel("IRCP-I")
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    st.pyplot(fig)

elif menu == "Base de tipos penales":
    st.header("Base de tipos penales")
    st.dataframe(base, use_container_width=True, hide_index=True)

else:
    st.header("Metodología")
    st.markdown("""
### Fórmulas principales

**Control formal** = (Legalidad × 0,60) + (Claridad normativa × 0,40)

**Control material** = (Lesividad × 0,60) + (Mínima intervención penal × 0,40)

**Test europeo** = (Idoneidad × 0,50) + (Necesidad penal × 0,50)

**Puntaje Alexy** = Peso interés estatal / (Peso derechos indígenas + Peso interés estatal) × 100

**CPI** = (Formal × 0,25) + (Material × 0,30) + (Europeo × 0,25) + (Alexy × 0,20)

**IRCP-I** = 100 − CPI
""")
