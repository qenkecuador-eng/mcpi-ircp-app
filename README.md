# MCPI-IRCP-I v9 · Pesos editables, supuestos ampliados y base jurisprudencial reforzada

Herramienta Streamlit de apoyo judicial para evaluar compatibilidad penal intercultural y riesgo de criminalización de protesta indígena.

## Novedades v9

1. **Supuestos de hecho ampliados**
   - Cada tipo penal cuenta con escenarios comunes de criminalización de protesta indígena.
   - Incluye movilización pacífica, corte temporal, recuperación territorial, operativo/desalojo, discurso/convocatoria, daños aislados no individualizados y escenario grave de control.

2. **Pesos editables**
   - En `Evaluación desde hechos`, los pesos de los subcriterios son editables.
   - La app valida si los pesos por variable suman 1.00.
   - Incluye botón para normalizar pesos automáticamente.
   - Se puede exportar la matriz de pesos en CSV y Excel.

3. **Base jurisprudencial y estándares ampliada**
   - Incluye casos de la tesis, estándares del Sistema Interamericano, precedentes de consulta previa, protesta, criminalización, derechos indígenas, proporcionalidad y control de convencionalidad.
   - Diferencia casos documentados de criminalización y precedentes jurisprudenciales.

4. **Estado editable en hechos críticos**
   - El módulo de control probatorio permite editar el estado del hecho crítico.
   - Opciones: Sí, No, No consta, No aplica, No / deficiente, Parcial, Alegado por Fiscalía, Alegado por defensa, Controvertido.

5. **Ideas para operadores de justicia**
   - Módulo con mejoras funcionales y criterios de uso institucional.

## Archivos requeridos en la raíz del repositorio

- app.py
- requirements.txt
- tipos_penales_base_ampliada.csv
- subcriterios_MCPI_IRCP_app_unico.xlsx
- supuestos_hecho_base.csv
- tipicidad_por_delito.csv
- estandares_convencionalidad.csv
- jurisprudencia_base.csv
- catalogo_descripciones_tipo_penal.csv
- README.md

## Streamlit Cloud

Main file path: `app.py`.

## Nota metodológica

La app no decide casos. Estructura el razonamiento jurídico y exige control humano, prueba, contradicción, motivación, tipicidad estricta, proporcionalidad, mínima intervención y enfoque intercultural.


## Novedad v10

- Se agrega `catalogo_descripciones_tipo_penal.csv`.
- En `1. Expediente del caso`, apartado `Descripción`, se puede cargar un catálogo exhaustivo de supuestos comunes de criminalización para cada tipo penal.
- El catálogo es editable desde `Administración de bases`.


## Corrección v11 del catálogo de descripciones

El catálogo fue depurado para evitar confusión: ya no se cargan supuestos genéricos iguales para todos los delitos.  
Cada tipo penal tiene ahora únicamente supuestos compatibles con su núcleo típico.  
Ejemplo: apología contiene solo discursos, comunicados, defensa pública o convocatoria; usurpación contiene recuperación territorial, permanencia o ingreso a predio; sabotaje contiene infraestructura, daño grave o afectación esencial.
