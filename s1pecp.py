import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Simulador de Evaluación de Proyectos", layout="wide")

st.title("🎮 Simulador de Evaluación de Proyectos")

# ---------------------------------------------------
# CONTROL DE EXPLICACIÓN ABIERTA
# ---------------------------------------------------

if "open_explainer" not in st.session_state:
    st.session_state.open_explainer = None

# ---------------------------------------------------
# CASO COMPLETO POR SECCIONES
# ---------------------------------------------------

case_sections = {
    "Contexto inicial": {
        "texto": """
Una empresa de retail ha experimentado un aumento significativo en retrasos en sus entregas durante los últimos 6 meses.

Esto ha generado:
- Quejas constantes de clientes
- Aumento en costos logísticos
- Pérdida de ventas

La gerencia decide lanzar un proyecto para resolver el problema lo antes posible.
""",
        "factores": []
    },

    "Definición del problema": {
        "texto": """
No se ha realizado un diagnóstico formal del problema.

Algunas áreas creen que el problema es tecnológico, otras que es operativo.

No todos tienen la misma visión del problema.
""",
        "factores": [
            "Claridad del problema",
            "Visión compartida"
        ]
    },

    "Requerimientos": {
        "texto": """
No existe una lista clara de requerimientos.

Se espera resolver la mayoría de los problemas en una sola iniciativa.
""",
        "factores": [
            "Requerimientos definidos",
            "Requerimientos realistas"
        ]
    },

    "Tiempo y planificación": {
        "texto": """
El proyecto debe completarse en 2 meses por presión de la gerencia.

Existe un plan general, pero sin mucho detalle.
""",
        "factores": [
            "Tiempo realista",
            "Plan inicial"
        ]
    },

    "Organización": {
        "texto": """
No está completamente claro quién es el sponsor del proyecto.

La gerencia apoya el proyecto, pero no participa activamente.

Logística y ventas tienen prioridades diferentes.
""",
        "factores": [
            "Sponsor definido",
            "Apoyo del sponsor",
            "Alineación entre áreas"
        ]
    },

    "Comunicación": {
        "texto": """
La comunicación existe, pero es informal y desordenada.

La información no siempre se comparte a tiempo.
""",
        "factores": [
            "Canales definidos",
            "Flujo de información"
        ]
    },

    "Riesgos": {
        "texto": """
No se han identificado formalmente los riesgos.

No se ha evaluado el impacto de posibles problemas.
""",
        "factores": [
            "Identificación de riesgos",
            "Impacto de riesgos"
        ]
    },

    "Recursos": {
        "texto": """
El equipo asignado trabaja en múltiples proyectos.

El proyecto depende fuertemente de un analista clave.
""",
        "factores": [
            "Capacidad del equipo",
            "Dependencia crítica"
        ]
    },

    "Condición final": {
        "texto": """
La gerencia insiste en iniciar el proyecto inmediatamente, sin realizar análisis adicional, debido a la presión del mercado.
""",
        "factores": []
    }
}

# ---------------------------------------------------
# EXPLICACIONES COMPLETAS (15)
# ---------------------------------------------------

explicaciones = {

"Claridad del problema": """
### 📘 Concepto
Un proyecto debe comenzar con una definición clara del problema que se busca resolver.

### ⚠️ Problema típico
Muchas organizaciones comienzan directamente con soluciones sin entender la causa raíz.

### 💥 Consecuencia
Se implementan soluciones incorrectas que no atacan el problema real.

### 🧠 Ejemplo
Se compra software nuevo cuando el problema es desorganización operativa.

### 🎯 Conclusión
Sin claridad, el proyecto pierde dirección desde el inicio.
""",

"Visión compartida": """
### 📘 Concepto
Todos los interesados deben compartir una misma comprensión del problema.

### ⚠️ Problema típico
Cada área interpreta el problema desde su perspectiva.

### 💥 Consecuencia
Se generan conflictos, retrabajo y falta de coherencia.

### 🧠 Ejemplo
Ventas cree que el problema es demanda, logística cree que es distribución.

### 🎯 Conclusión
Sin visión compartida, no hay integración del proyecto.
""",

"Requerimientos definidos": """
### 📘 Concepto
Los requerimientos establecen qué debe lograr el proyecto.

### ⚠️ Problema típico
No se documenta lo que se espera lograr.

### 💥 Consecuencia
No se puede medir avance ni controlar alcance.

### 🧠 Ejemplo
Se sigue trabajando sin saber cuándo se termina.

### 🎯 Conclusión
Sin requerimientos claros, el proyecto es incontrolable.
""",

"Requerimientos realistas": """
### 📘 Concepto
Los proyectos deben tener objetivos alcanzables.

### ⚠️ Problema típico
Se intenta resolver todo en una sola iniciativa.

### 💥 Consecuencia
Se sobrecarga el proyecto y se reduce su probabilidad de éxito.

### 🧠 Ejemplo
Intentar cambiar todo el sistema logístico de una vez.

### 🎯 Conclusión
La ambición excesiva reduce la viabilidad.
""",

"Tiempo realista": """
### 📘 Concepto
El tiempo es una de las principales restricciones del proyecto.

### ⚠️ Problema típico
Los plazos se imponen sin análisis.

### 💥 Consecuencia
Se generan errores, presión y baja calidad.

### 🧠 Ejemplo
Un proyecto de meses se intenta hacer en semanas.

### 🎯 Conclusión
Un plazo irreal es una de las principales causas de fracaso.
""",

"Plan inicial": """
### 📘 Concepto
Todo proyecto necesita al menos una estructura básica de planificación.

### ⚠️ Problema típico
Se inicia sin plan formal.

### 💥 Consecuencia
El proyecto se vuelve reactivo.

### 🧠 Ejemplo
Se trabaja “día a día” sin dirección.

### 🎯 Conclusión
Planificar es reducir incertidumbre.
""",

"Sponsor definido": """
### 📘 Concepto
El sponsor es el responsable estratégico del proyecto.

### ⚠️ Problema típico
No hay claridad de quién toma decisiones.

### 💥 Consecuencia
Falta liderazgo y dirección.

### 🧠 Ejemplo
Nadie aprueba cambios importantes.

### 🎯 Conclusión
Sin sponsor, no hay proyecto.
""",

"Apoyo del sponsor": """
### 📘 Concepto
El sponsor debe involucrarse activamente.

### ⚠️ Problema típico
El sponsor solo “aprueba”, pero no participa.

### 💥 Consecuencia
Los conflictos no se resuelven.

### 🧠 Ejemplo
El equipo queda solo ante problemas.

### 🎯 Conclusión
El apoyo activo es crítico.
""",

"Alineación entre áreas": """
### 📘 Concepto
Los proyectos cruzan múltiples áreas.

### ⚠️ Problema típico
Cada área tiene prioridades distintas.

### 💥 Consecuencia
Bloqueos, retrasos y conflictos.

### 🧠 Ejemplo
Ventas empuja, logística no puede cumplir.

### 🎯 Conclusión
Sin alineación, no hay ejecución efectiva.
""",

"Canales definidos": """
### 📘 Concepto
La comunicación debe estar estructurada.

### ⚠️ Problema típico
Comunicación informal.

### 💥 Consecuencia
Información inconsistente.

### 🧠 Ejemplo
Mensajes por múltiples medios sin control.

### 🎯 Conclusión
La comunicación debe diseñarse.
""",

"Flujo de información": """
### 📘 Concepto
La información debe llegar a tiempo.

### ⚠️ Problema típico
Retrasos en comunicación.

### 💥 Consecuencia
Decisiones incorrectas.

### 🧠 Ejemplo
Se detecta un problema demasiado tarde.

### 🎯 Conclusión
La información es crítica para el control.
""",

"Identificación de riesgos": """
### 📘 Concepto
Los riesgos deben anticiparse.

### ⚠️ Problema típico
No se consideran.

### 💥 Consecuencia
Surgen como crisis.

### 🧠 Ejemplo
Problemas inesperados paralizan el proyecto.

### 🎯 Conclusión
Identificar riesgos es obligatorio.
""",

"Impacto de riesgos": """
### 📘 Concepto
No todos los riesgos son iguales.

### ⚠️ Problema típico
No se priorizan.

### 💥 Consecuencia
Se pierde enfoque.

### 🧠 Ejemplo
Se atienden riesgos menores y se ignoran críticos.

### 🎯 Conclusión
El impacto define la prioridad.
""",

"Capacidad del equipo": """
### 📘 Concepto
El equipo tiene límites.

### ⚠️ Problema típico
Sobrecarga.

### 💥 Consecuencia
Errores y retrasos.

### 🧠 Ejemplo
Personas en múltiples proyectos.

### 🎯 Conclusión
La capacidad define la velocidad real.
""",

"Dependencia crítica": """
### 📘 Concepto
Evitar puntos únicos de falla.

### ⚠️ Problema típico
Dependencia de una persona.

### 💥 Consecuencia
Riesgo alto de interrupción.

### 🧠 Ejemplo
Un experto clave se ausenta.

### 🎯 Conclusión
La redundancia es necesaria.
"""
}

# ---------------------------------------------------
# SESSION STATE
# ---------------------------------------------------

if "scores" not in st.session_state:
    st.session_state.scores = {}

if "section_index" not in st.session_state:
    st.session_state.section_index = 0

section_names = list(case_sections.keys())

# ---------------------------------------------------
# BOTÓN HOME
# ---------------------------------------------------

if st.button("🏠 Volver al inicio"):
    st.session_state.section_index = 0

# ---------------------------------------------------
# NAVEGACIÓN
# ---------------------------------------------------

col1, col2, col3 = st.columns([1,2,1])

with col1:
    if st.button("⬅️ Anterior") and st.session_state.section_index > 0:
        st.session_state.section_index -= 1

with col3:
    if st.button("Siguiente ➡️") and st.session_state.section_index < len(section_names) - 1:
        st.session_state.section_index += 1

# ---------------------------------------------------
# SECCIÓN ACTUAL
# ---------------------------------------------------

current_section = section_names[st.session_state.section_index]
section_data = case_sections[current_section]

st.subheader(f"📌 {current_section}")
st.info(section_data["texto"])

# ---------------------------------------------------
# FACTORES + EXPLICACIÓN CONTROLADA
# ---------------------------------------------------

all_factors = [f for sec in case_sections.values() for f in sec["factores"]]

if section_data["factores"]:

    for factor in section_data["factores"]:
        index = all_factors.index(factor) + 1
        total_factors = len(all_factors)

        st.markdown(f"### {index} de {total_factors} — {factor}")

        value = st.selectbox(
            "Evaluación",
            [0,1,2],
            format_func=lambda x: {0:"❌ Deficiente",1:"⚠️ Parcial",2:"✔️ Sólido"}[x],
            key=factor
        )

        st.session_state.scores[factor] = value

        if st.button(f"📘 Ver explicación — {factor}", key=f"btn_{factor}"):
            st.session_state.open_explainer = factor

        if st.session_state.open_explainer == factor:
            st.markdown(explicaciones[factor])

# ---------------------------------------------------
# RESULTADO
# ---------------------------------------------------

st.markdown("---")
st.subheader("📊 Resultado")

if st.session_state.scores:

    df = pd.DataFrame(list(st.session_state.scores.items()), columns=["Factor","Score"])
    total = df["Score"].sum()
    indice = total / 30

    st.dataframe(df, use_container_width=True)

    st.metric("Score total", total)
    st.metric("Índice", f"{indice:.2f}")

    st.markdown("""
### 📘 Interpretación del índice

- 🔴 0 – 10 → Alto riesgo  
- 🟡 11 – 20 → Riesgo medio  
- 🟢 21 – 30 → Buenas condiciones  

Este valor representa qué tan bien está planteado el proyecto.
""")

# ---------------------------------------------------
# EXPORTAR
# ---------------------------------------------------

def export_excel(df, total, indice):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
        resumen = pd.DataFrame({
            "Métrica": ["Total", "Índice"],
            "Valor": [total, indice]
        })
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
    return output.getvalue()

if st.button("📥 Exportar a Excel"):
    excel_data = export_excel(df, total, indice)
    st.download_button("Descargar Excel", excel_data, file_name="evaluacion.xlsx")