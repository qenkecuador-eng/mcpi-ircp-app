
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
st.set_page_config(page_title="Matriz MCPI-IRCP-I v5", page_icon="⚖️", layout="wide", initial_sidebar_state="expanded")

COLUMNAS_BASE = ["País","Norma","Artículo","Tipo penal","Categoría","Estado","Contexto típico de protesta","Legalidad","Claridad normativa","Lesividad","Mínima intervención penal","Idoneidad","Necesidad penal","Peso derechos indígenas","Peso interés estatal","Observación metodológica"]
NUMERIC_COLS = ["Legalidad","Claridad normativa","Lesividad","Mínima intervención penal","Idoneidad","Necesidad penal","Peso derechos indígenas","Peso interés estatal"]

def normalizar_base(df):
    df=df.copy()
    for c in COLUMNAS_BASE:
        if c not in df.columns: df[c]=0 if c in NUMERIC_COLS else ""
    df=df[COLUMNAS_BASE]
    for c in NUMERIC_COLS: df[c]=pd.to_numeric(df[c], errors="coerce").fillna(0).clip(0,100)
    for c in [x for x in COLUMNAS_BASE if x not in NUMERIC_COLS]: df[c]=df[c].fillna("").astype(str)
    return df

def cargar_default():
    return normalizar_base(pd.read_csv(BASE_PATH))

if "base_actual" not in st.session_state:
    st.session_state.base_actual=cargar_default()

def nivel_compatibilidad(v):
    v=float(v)
    if v>=75: return "Compatibilidad alta","Verde"
    if v>=60: return "Compatibilidad condicionada","Amarillo"
    if v>=40: return "Compatibilidad débil","Naranja"
    return "Incompatibilidad","Rojo"

def nivel_riesgo(v):
    v=float(v)
    if v>=70: return "Riesgo alto","Rojo"
    if v>=50: return "Riesgo medio-alto","Naranja"
    if v>=35: return "Riesgo medio","Amarillo"
    return "Riesgo bajo","Verde"

def calc_ponderado(df):
    df=df.copy(); df["Puntaje"]=pd.to_numeric(df["Puntaje"], errors="coerce").fillna(0).clip(0,100); df["Peso"]=pd.to_numeric(df["Peso"], errors="coerce").fillna(0)
    df["Aporte"]=(df["Puntaje"]*df["Peso"]).round(2); total=round(float(df["Aporte"].sum()),2); res,sem=nivel_compatibilidad(total)
    return df,total,res,sem

def calcular_fila(r):
    formal=round(float(r["Legalidad"])*0.60+float(r["Claridad normativa"])*0.40,2)
    material=round(float(r["Lesividad"])*0.60+float(r["Mínima intervención penal"])*0.40,2)
    europeo=round(float(r["Idoneidad"])*0.50+float(r["Necesidad penal"])*0.50,2)
    pdi=float(r["Peso derechos indígenas"]); pie=float(r["Peso interés estatal"])
    alexy=round((pie/(pdi+pie))*100,2) if (pdi+pie) else 0; rel=round(pdi/pie,2) if pie else 999
    cpi=round((formal*0.25)+(material*0.30)+(europeo*0.25)+(alexy*0.20),2); ircp=round(100-cpi,2); nivel,sem=nivel_riesgo(ircp)
    return {"País":r["País"],"Artículo":r["Artículo"],"Tipo penal":r["Tipo penal"],"Categoría":r["Categoría"],"Estado":r["Estado"],"Formal":formal,"Material":material,"Europeo":europeo,"Peso relativo":rel,"Alexy":alexy,"CPI":cpi,"IRCP-I":ircp,"Nivel":nivel,"Semáforo":sem}

def base_resultados(df):
    return pd.DataFrame([calcular_fila(r) for _,r in df.iterrows()])

def matriz_editor(titulo,data,key):
    st.subheader(titulo)
    edited=st.data_editor(pd.DataFrame(data), key=key, hide_index=True, use_container_width=True, column_config={"Variable":st.column_config.TextColumn(disabled=True),"Descripción":st.column_config.TextColumn(disabled=True),"Peso":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Puntaje":st.column_config.NumberColumn(min_value=0,max_value=100,step=1,format="%.0f"),"Justificación":st.column_config.TextColumn(width="large")}, disabled=["Variable","Descripción","Peso"])
    calc,total,res,sem=calc_ponderado(edited); st.dataframe(calc[["Variable","Puntaje","Peso","Aporte","Justificación"]], hide_index=True, use_container_width=True); st.success(f"{titulo}: {total} — {res} / {sem}")
    return calc,total,res,sem

def to_excel_bytes(df):
    bio=BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        df.to_excel(w,index=False,sheet_name="Base editable"); base_resultados(df).to_excel(w,index=False,sheet_name="Resultados")
    bio.seek(0); return bio.getvalue()

def generar_word(resumen, matrices):
    if Document is None: return None
    doc=Document(); doc.add_heading("Informe MCPI-IRCP-I",0)
    for k in ["País","Artículo","Tipo penal","Contexto"]: doc.add_paragraph(f"{k}: {resumen[k]}")
    doc.add_heading("Resultado global",level=1); t=doc.add_table(rows=1,cols=2); t.style="Table Grid"; t.rows[0].cells[0].text="Indicador"; t.rows[0].cells[1].text="Valor"
    for k in ["Formal","Material","Europeo","Peso relativo","Alexy","CPI","IRCP-I","Nivel","Semáforo"]:
        c=t.add_row().cells; c[0].text=k; c[1].text=str(resumen[k])
    for nombre,df in matrices.items():
        doc.add_heading(nombre,level=1); tt=doc.add_table(rows=1,cols=5); tt.style="Table Grid"; hs=["Variable","Puntaje","Peso","Aporte","Justificación"]
        for i,h in enumerate(hs): tt.rows[0].cells[i].text=h
        for _,row in df.iterrows():
            c=tt.add_row().cells
            for i,h in enumerate(hs): c[i].text=str(row.get(h,""))
    doc.add_heading("Conclusión",level=1); doc.add_paragraph(resumen["Conclusión"])
    bio=BytesIO(); doc.save(bio); return bio.getvalue()

st.title("Matriz MCPI-IRCP-I v5")
st.caption("Versión editable: modificar valores, agregar delitos, importar/exportar base y recalcular el índice global.")
menu=st.sidebar.radio("Módulos",["Evaluación en matriz","Base editable y ampliada","Añadir delito","Importar / Exportar","Dashboard comparativo","Metodología y reglas"])
base=st.session_state.base_actual

if menu=="Evaluación en matriz":
    st.header("Evaluación en matriz")
    if base.empty: st.error("La base está vacía."); st.stop()
    opciones=[f"{i} | {r['País']} | {r['Tipo penal']} | {r['Artículo']}" for i,r in base.iterrows()]
    default=next((i for i,x in enumerate(opciones) if "194" in x),0); sel=st.selectbox("Seleccione delito/tipo penal",opciones,index=default); idx=int(sel.split("|")[0].strip()); row=base.loc[idx].to_dict()
    st.subheader("Datos del caso"); c1,c2,c3=st.columns(3)
    with c1: pais=st.text_input("País",row["País"])
    with c2: norma=st.text_input("Norma",row["Norma"])
    with c3: articulo=st.text_input("Artículo",row["Artículo"])
    tipo=st.text_input("Tipo penal",row["Tipo penal"]); contexto=st.text_area("Contexto típico de protesta",row["Contexto típico de protesta"],height=80); obs=st.text_area("Observación metodológica",row["Observación metodológica"],height=80)
    formal_df,formal,_,_=matriz_editor("1. Control formal",[{"Variable":"Legalidad","Descripción":"Ley previa, precisión de conducta, sujetos, elementos del tipo y prohibición de analogía.","Peso":0.60,"Puntaje":row["Legalidad"],"Justificación":"Evaluar si el tipo es previo, estricto y no permite aplicación analógica."},{"Variable":"Claridad normativa","Descripción":"Comprensibilidad, verbos rectores, bien jurídico, umbral de gravedad y salvaguardas.","Peso":0.40,"Puntaje":row["Claridad normativa"],"Justificación":"Evaluar si la norma es clara y previsible frente a protesta protegida."}],"formal")
    material_df,material,_,_=matriz_editor("2. Control material",[{"Variable":"Lesividad","Descripción":"Daño real, gravedad, riesgo concreto y prueba individualizada.","Peso":0.60,"Puntaje":row["Lesividad"],"Justificación":"Evaluar si existe daño real y concreto, no mera afectación abstracta."},{"Variable":"Mínima intervención penal","Descripción":"Alternativas no penales, diálogo, consulta, mediación y última ratio.","Peso":0.40,"Puntaje":row["Mínima intervención penal"],"Justificación":"Evaluar si el derecho penal era indispensable."}],"material")
    europeo_df,europeo,_,_=matriz_editor("3. Test europeo de proporcionalidad",[{"Variable":"Idoneidad","Descripción":"Conexión entre persecución penal y finalidad constitucional legítima.","Peso":0.50,"Puntaje":row["Idoneidad"],"Justificación":"Evaluar si la vía penal protege realmente el bien invocado."},{"Variable":"Necesidad penal","Descripción":"Ausencia de medios menos restrictivos igualmente eficaces.","Peso":0.50,"Puntaje":row["Necesidad penal"],"Justificación":"Evaluar si diálogo, consulta u otras vías eran suficientes."}],"europeo")
    st.subheader("4. Ponderación de Alexy")
    alexy_in=pd.DataFrame([{"Variable":"Peso derechos indígenas","Descripción":"Territorio, consulta previa, autodeterminación, protesta, participación y expresión colectiva.","Valor":row["Peso derechos indígenas"],"Justificación":"Valorar intensidad de afectación de derechos indígenas."},{"Variable":"Peso interés estatal","Descripción":"Orden público, seguridad, circulación, propiedad, servicios públicos o integridad.","Valor":row["Peso interés estatal"],"Justificación":"Valorar intensidad del interés estatal realmente afectado."}])
    alexy_edit=st.data_editor(alexy_in,key="alexy",hide_index=True,use_container_width=True,column_config={"Variable":st.column_config.TextColumn(disabled=True),"Descripción":st.column_config.TextColumn(disabled=True),"Valor":st.column_config.NumberColumn(min_value=0,max_value=100,step=1,format="%.0f"),"Justificación":st.column_config.TextColumn(width="large")},disabled=["Variable","Descripción"])
    pdi=float(alexy_edit.loc[alexy_edit["Variable"]=="Peso derechos indígenas","Valor"].iloc[0]); pie=float(alexy_edit.loc[alexy_edit["Variable"]=="Peso interés estatal","Valor"].iloc[0]); rel=round(pdi/pie,2) if pie else 999; alexy=round((pie/(pdi+pie))*100,2) if (pdi+pie) else 0; ar,asem=nivel_compatibilidad(alexy)
    st.dataframe(pd.DataFrame([{"Indicador":"Peso relativo","Fórmula":"Derechos indígenas / interés estatal","Resultado":rel},{"Indicador":"Puntaje Alexy","Fórmula":"Interés estatal / (derechos indígenas + interés estatal) × 100","Resultado":alexy}]),hide_index=True,use_container_width=True); st.success(f"Alexy: {alexy} — {ar} / {asem}")
    st.subheader("5. Índice global")
    ind=pd.DataFrame([{"Componente":"Control formal","Resultado":formal,"Peso":0.25,"Aporte":round(formal*0.25,2)},{"Componente":"Control material","Resultado":material,"Peso":0.30,"Aporte":round(material*0.30,2)},{"Componente":"Test europeo","Resultado":europeo,"Peso":0.25,"Aporte":round(europeo*0.25,2)},{"Componente":"Puntaje Alexy","Resultado":alexy,"Peso":0.20,"Aporte":round(alexy*0.20,2)}])
    cpi=round(float(ind["Aporte"].sum()),2); ircp=round(100-cpi,2); nivel,sem=nivel_riesgo(ircp); st.dataframe(ind,hide_index=True,use_container_width=True)
    m1,m2,m3=st.columns(3); m1.metric("Compatibilidad Penal Intercultural (CPI)",cpi); m2.metric("IRCP-I",ircp); m3.metric("Nivel / semáforo",f"{nivel} / {sem}")
    st.code(f"CPI = ({formal} × 0,25) + ({material} × 0,30) + ({europeo} × 0,25) + ({alexy} × 0,20) = {cpi}\nIRCP-I = 100 − {cpi} = {ircp}",language="text")
    st.subheader("Reglas de cierre")
    reglas=[]
    if formal_df.loc[formal_df["Variable"]=="Legalidad","Puntaje"].iloc[0]<40: reglas.append("Legalidad menor a 40: presunción de incompatibilidad formal.")
    if material_df.loc[material_df["Variable"]=="Lesividad","Puntaje"].iloc[0]<40: reglas.append("Lesividad menor a 40: no supera el control material.")
    if europeo_df.loc[europeo_df["Variable"]=="Idoneidad","Puntaje"].iloc[0]<75: reglas.append("Idoneidad menor a 75: la criminalización penal no queda justificada.")
    if europeo_df.loc[europeo_df["Variable"]=="Necesidad penal","Puntaje"].iloc[0]<60: reglas.append("Necesidad penal menor a 60: debe preferirse una alternativa no penal.")
    if rel>1: reglas.append("Alexy favorece derechos indígenas: exige motivación reforzada.")
    if ircp>=70: reglas.append("IRCP-I alto: riesgo severo de criminalización penal intercultural.")
    elif ircp>=50: reglas.append("IRCP-I medio-alto: riesgo relevante de criminalización penal intercultural.")
    for r in reglas: st.warning(r)
    conclusion=f"El tipo penal evaluado ({tipo}, {articulo}, {pais}) presenta una Compatibilidad Penal Intercultural de {cpi} y un Índice de Riesgo de Criminalización Penal Intercultural de {ircp}, equivalente a {nivel} / semáforo {sem}. En el contexto analizado, la aplicación penal exige control de convencionalidad reforzado, interpretación restrictiva, prueba individualizada, daño concreto, riesgo real y ausencia de alternativas menos lesivas."
    st.subheader("Conclusión automática"); st.write(conclusion)
    chart=pd.DataFrame({"Componente":["Formal","Material","Europeo","Alexy","CPI","IRCP-I"],"Valor":[formal,material,europeo,alexy,cpi,ircp]}); fig,ax=plt.subplots(figsize=(9,4)); ax.bar(chart["Componente"],chart["Valor"]); ax.set_ylim(0,100); ax.set_ylabel("Puntaje"); ax.grid(axis="y",linestyle="--",alpha=.4); st.pyplot(fig)
    word=generar_word({"País":pais,"Artículo":articulo,"Tipo penal":tipo,"Contexto":contexto,"Formal":formal,"Material":material,"Europeo":europeo,"Peso relativo":rel,"Alexy":alexy,"CPI":cpi,"IRCP-I":ircp,"Nivel":nivel,"Semáforo":sem,"Conclusión":conclusion},{"Control formal":formal_df,"Control material":material_df,"Test europeo":europeo_df,"Alexy":alexy_edit})
    if word: st.download_button("Descargar informe Word",data=word,file_name="informe_MCPI_IRCP.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

elif menu=="Base editable y ampliada":
    st.header("Base editable y ampliada"); st.info("Edite valores o agregue filas. Para conservar cambios, descargue CSV/Excel y súbalo a GitHub reemplazando el archivo base.")
    edited=st.data_editor(st.session_state.base_actual,num_rows="dynamic",hide_index=True,use_container_width=True,column_config={c:st.column_config.NumberColumn(c,min_value=0,max_value=100,step=1) for c in NUMERIC_COLS})
    if st.button("Aplicar cambios a la sesión"):
        st.session_state.base_actual=normalizar_base(edited); st.success("Base actualizada en la sesión.")
    st.download_button("Descargar base CSV",data=normalizar_base(edited).to_csv(index=False).encode("utf-8-sig"),file_name="tipos_penales_base_ampliada.csv",mime="text/csv")
    st.download_button("Descargar base Excel",data=to_excel_bytes(normalizar_base(edited)),file_name="tipos_penales_base_ampliada.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu=="Añadir delito":
    st.header("Añadir nuevo delito/tipo penal")
    with st.form("nuevo"):
        c1,c2,c3=st.columns(3)
        with c1: pais=st.selectbox("País",["Argentina","Ecuador","Otro"]); norma=st.text_input("Norma","Código Penal" if pais=="Argentina" else "COIP")
        with c2: articulo=st.text_input("Artículo",""); categoria=st.text_input("Categoría","Orden público")
        with c3: estado=st.selectbox("Estado",["Ampliado preliminar","Base tesis","Validado por investigador"]); tipo=st.text_input("Tipo penal","")
        contexto=st.text_area("Contexto típico de protesta",""); obs=st.text_area("Observación metodológica","")
        p1,p2,p3,p4=st.columns(4)
        with p1: legalidad=st.number_input("Legalidad",0,100,50); claridad=st.number_input("Claridad normativa",0,100,45)
        with p2: lesividad=st.number_input("Lesividad",0,100,45); minima=st.number_input("Mínima intervención penal",0,100,35)
        with p3: idoneidad=st.number_input("Idoneidad",0,100,50); necesidad=st.number_input("Necesidad penal",0,100,35)
        with p4: pdi=st.number_input("Peso derechos indígenas",0,100,40); pie=st.number_input("Peso interés estatal",0,100,24)
        submitted=st.form_submit_button("Agregar a la base")
    if submitted:
        nueva=pd.DataFrame([{ "País":pais,"Norma":norma,"Artículo":articulo,"Tipo penal":tipo,"Categoría":categoria,"Estado":estado,"Contexto típico de protesta":contexto,"Legalidad":legalidad,"Claridad normativa":claridad,"Lesividad":lesividad,"Mínima intervención penal":minima,"Idoneidad":idoneidad,"Necesidad penal":necesidad,"Peso derechos indígenas":pdi,"Peso interés estatal":pie,"Observación metodológica":obs}])
        st.session_state.base_actual=normalizar_base(pd.concat([st.session_state.base_actual,nueva],ignore_index=True)); st.success("Delito agregado a la base de la sesión.")

elif menu=="Importar / Exportar":
    st.header("Importar / Exportar base"); up=st.file_uploader("Importar CSV o Excel",type=["csv","xlsx"]); modo=st.radio("Modo",["Reemplazar base actual","Agregar a base actual"],horizontal=True)
    if up is not None:
        try:
            nueva=pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up); nueva=normalizar_base(nueva); st.dataframe(nueva,hide_index=True,use_container_width=True)
            if st.button("Confirmar importación"):
                st.session_state.base_actual=nueva if modo=="Reemplazar base actual" else normalizar_base(pd.concat([st.session_state.base_actual,nueva],ignore_index=True)); st.success("Importación aplicada.")
        except Exception as e: st.error(f"No se pudo importar: {e}")
    st.download_button("Descargar CSV actual",data=st.session_state.base_actual.to_csv(index=False).encode("utf-8-sig"),file_name="tipos_penales_base_actual.csv",mime="text/csv")
    st.download_button("Descargar Excel actual",data=to_excel_bytes(st.session_state.base_actual),file_name="tipos_penales_base_actual.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu=="Dashboard comparativo":
    st.header("Dashboard comparativo"); res=base_resultados(st.session_state.base_actual); st.dataframe(res,hide_index=True,use_container_width=True)
    c1,c2,c3=st.columns(3); c1.metric("Delitos en base",len(res)); c2.metric("IRCP-I promedio",round(res["IRCP-I"].mean(),2)); c3.metric("Mayor riesgo",res.loc[res["IRCP-I"].idxmax(),"Tipo penal"])
    plot=res.sort_values("IRCP-I"); fig,ax=plt.subplots(figsize=(11,max(5,len(plot)*0.28))); ax.barh(plot["País"]+" - "+plot["Tipo penal"],plot["IRCP-I"]); ax.set_xlim(0,100); ax.set_xlabel("IRCP-I"); ax.grid(axis="x",linestyle="--",alpha=.4); st.pyplot(fig)
    prom=res.groupby("País",as_index=False)["IRCP-I"].mean(); fig2,ax2=plt.subplots(figsize=(7,4)); ax2.bar(prom["País"],prom["IRCP-I"]); ax2.set_ylim(0,100); ax2.set_ylabel("IRCP-I promedio"); ax2.grid(axis="y",linestyle="--",alpha=.4); st.pyplot(fig2)
else:
    st.header("Metodología y reglas")
    st.markdown("""
**Control formal** = (Legalidad × 0,60) + (Claridad normativa × 0,40)  
**Control material** = (Lesividad × 0,60) + (Mínima intervención penal × 0,40)  
**Test europeo** = (Idoneidad × 0,50) + (Necesidad penal × 0,50)  
**Puntaje Alexy** = Peso interés estatal / (Peso derechos indígenas + Peso interés estatal) × 100  
**CPI** = (Formal × 0,25) + (Material × 0,30) + (Europeo × 0,25) + (Alexy × 0,20)  
**IRCP-I** = 100 − CPI

Reglas de cierre: legalidad < 40, lesividad < 40, idoneidad < 75, necesidad penal < 60, Alexy > 1, IRCP-I >= 70.

En Streamlit Cloud, los cambios quedan en la sesión. Para hacerlos permanentes: descargar la base editada y reemplazar `tipos_penales_base_ampliada.csv` en GitHub.
""")
st.caption("Herramienta académica orientativa. No sustituye análisis jurídico definitivo ni control judicial del caso concreto.")
