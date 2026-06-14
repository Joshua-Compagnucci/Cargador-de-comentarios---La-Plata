
import streamlit as st
import pandas as pd
import re
from pathlib import Path
from datetime import datetime

# ======================================================
# CONFIG
# ======================================================

st.set_page_config(
    page_title="Cargador de comentarios la plata",
    layout="wide"
)

ARCHIVO_EXCEL = "comentarios_etiquetados_sentimiento.xlsx"

st.markdown("""
<style>
.stApp { background-color: #0F172A; color: white; }
.block-container { padding-top: 1rem; }
.card {
    background-color: #111827;
    border-radius: 12px;
    padding: 16px;
    border: 1px solid #334155;
    margin-bottom: 15px;
}
</style>
""", unsafe_allow_html=True)

# ======================================================
# DICCIONARIOS
# ======================================================

negativas_fuertes = [
    "desastre", "vergüenza", "una vergüenza", "abandono", "abandonado",
    "horrible", "lamentable", "impresentable", "asco", "bronca",
    "indignante", "no se puede", "no va más", "todo mal",
    "son un desastre", "está destruido", "esta destruido",
    "se cae a pedazos", "nadie hace nada", "no hacen nada",
    "nos roban", "roban", "chorros", "chorro", "inseguridad",
    "miedo", "peligroso", "sucio", "basura", "baches", "pozos",
    "no hay plata", "no alcanza", "carísimo", "carisimo",
    "cerraron", "cerró", "cerro", "crisis", "fundieron",
    "fundió", "fundio", "no venden", "no funciona", "tasas",
    "impuestos", "aumentos", "aumento", "culpa", "mentira",
    "mentirosos", "corruptos", "corrupto", "inútiles", "inutiles",
    "estafadores", "dejen de robar", "roban todo"
]

negativas_suaves = [
    "mal", "feo", "caro", "problema", "problemas", "queja",
    "reclamo", "falta", "faltan", "falta de", "sin luz",
    "sin agua", "sin seguridad", "sin control", "sin obras",
    "no arreglan", "no solucionan", "no limpian", "cada vez peor",
    "peor", "complicado", "difícil", "dificil", "triste",
    "baja", "cae", "caída", "caida"
]

positivas_fuertes = [
    "excelente", "muy bien", "buenísimo", "buenisimo", "mejoró",
    "mejoro", "mejorando", "gran trabajo", "felicitaciones",
    "apoyo", "acompaño", "gracias", "solución", "solucion",
    "ordenado", "progreso", "avance", "bien hecho",
    "sigan así", "sigan asi", "por fin", "funciona perfecto"
]

positivas_suaves = [
    "bien", "bueno", "buena", "lindo", "linda", "funciona",
    "obra", "obras", "positivo", "me gusta", "correcto",
    "mejor", "arreglaron", "limpiaron"
]

temas = {
    "economia_consumo": [
        "consumo", "no hay plata", "no alcanza", "ventas", "comercio",
        "comercios", "local", "locales", "crisis", "cerraron", "cierre",
        "fundieron", "caro", "carísimo", "carisimo", "mercado libre", "internet"
    ],
    "impuestos_tasas": [
        "impuestos", "tasas", "municipal", "municipio", "estacionamiento",
        "vtv", "aumento", "suba", "cobrar"
    ],
    "alquileres": ["alquiler", "alquileres"],
    "inseguridad": [
        "inseguridad", "robo", "roban", "chorros", "miedo", "seguridad",
        "policía", "policia", "patrullero", "peligroso"
    ],
    "servicios_urbanos": [
        "bache", "baches", "pozo", "pozos", "calle", "calles",
        "luz", "alumbrado", "basura", "recolección", "recoleccion",
        "agua", "zanja", "arroyo", "cloaca", "cloacas", "plaza", "limpieza"
    ],
    "politica_ideologia": [
        "milei", "kicillof", "alak", "cristina", "cfk", "macri",
        "intendente", "gobernador", "presidente", "municipalidad",
        "kirchnerista", "libertario", "peronista", "zurdo", "kuka"
    ]
}

barrios_zonas = {
    "Centro": ["centro", "plaza san martín", "plaza san martin", "plaza moreno", "7 y 50", "8 y 50", "12 y 51"],
    "Tolosa": ["tolosa"],
    "Los Hornos": ["los hornos"],
    "Villa Elvira": ["villa elvira"],
    "San Carlos": ["san carlos"],
    "Gonnet": ["gonnet"],
    "City Bell": ["city bell"],
    "Ringuelet": ["ringuelet"],
    "Altos de San Lorenzo": ["altos de san lorenzo"],
    "Melchor Romero": ["romero", "melchor romero"],
    "Olmos": ["olmos", "lisandro olmos"],
    "Abasto": ["abasto"],
    "Etcheverry": ["etcheverry"],
    "Berisso Centro": ["berisso centro", "montevideo", "calle nueva york"],
    "Villa Nueva": ["villa nueva"],
    "El Carmen": ["el carmen"],
    "Barrio Obrero": ["barrio obrero"],
    "Villa Argüello": ["villa argüello", "villa arguello"],
    "Los Talas": ["los talas", "altos los talas"],
    "Palo Blanco": ["palo blanco"],
    "La Franja": ["la franja"]
}

politicos_alias = {
    "Milei": ["milei", "javier", "presidente", "libertario", "gobierno nacional"],
    "Kicillof": ["kicillof", "axel", "gobernador", "provincia"],
    "Alak": ["alak", "intendente", "municipio", "municipalidad"],
    "Cristina / CFK": ["cristina", "cfk", "kirchner"],
    "Macri": ["macri", "mauricio"]
}

perfiles = {
    "comerciante": ["mi local", "mi comercio", "soy comerciante", "vendo", "ventas", "clientes", "comerciantes"],
    "militante": ["viva milei", "milei carajo", "zurdo", "kuka", "kirchnerista", "libertario", "peronista", "cfk"],
    "vecino": ["vivo en", "soy de", "acá en", "aca en", "mi barrio", "mi zona", "vecino"],
    "anti_gestion": ["intendente", "municipio", "municipalidad", "gobernador", "presidente", "tasas", "impuestos"],
    "analitico": ["también", "tambien", "además", "ademas", "depende", "no es solo", "hay varios", "por un lado", "por otro"]
}

# ======================================================
# FUNCIONES
# ======================================================

def limpiar_linea(texto):
    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def detectar_por_diccionario(texto, diccionario, default="otro"):
    t = texto.lower()
    encontrados = []

    for categoria, palabras in diccionario.items():
        for p in palabras:
            if p in t:
                encontrados.append(categoria)
                break

    return ", ".join(sorted(set(encontrados))) if encontrados else default


def detectar_politicos(texto):
    return detectar_por_diccionario(texto, politicos_alias, default="Sin mención")


def detectar_barrio(texto):
    return detectar_por_diccionario(texto, barrios_zonas, default="Sin detectar")


def detectar_sentimiento(texto):
    t = texto.lower()

    score_neg = sum(1 for p in negativas_fuertes if p in t) * 3
    score_neg += sum(1 for p in negativas_suaves if p in t)

    score_pos = sum(1 for p in positivas_fuertes if p in t) * 3
    score_pos += sum(1 for p in positivas_suaves if p in t)

    if "??" in texto or "!!!" in texto:
        score_neg += 1

    if any(x in t for x in ["jajaja", "jaja", "😂", "🤣"]) and score_neg > 0:
        score_neg += 1

    if re.search(r"\bno\b.{0,12}\b(bien|bueno|buena|mejor|lindo|linda)\b", t):
        score_neg += 3
        score_pos = max(0, score_pos - 1)

    if re.search(r"\bno\b.{0,12}\b(mal|horrible|desastre)\b", t):
        score_pos += 2
        score_neg = max(0, score_neg - 1)

    if score_neg >= score_pos + 1 and score_neg >= 1:
        return "Negativo"
    if score_pos >= score_neg + 1 and score_pos >= 1:
        return "Positivo"

    return "Neutral"


def calcular_impacto(me_gustas, respuestas_detectadas):
    try:
        me_gustas = int(me_gustas)
    except Exception:
        me_gustas = 0

    try:
        respuestas_detectadas = int(respuestas_detectadas)
    except Exception:
        respuestas_detectadas = 0

    return me_gustas + respuestas_detectadas


def nivel_impacto(total_interacciones):
    try:
        total_interacciones = int(total_interacciones)
    except Exception:
        total_interacciones = 0

    if total_interacciones >= 100:
        return "Alto"
    if total_interacciones >= 20:
        return "Medio"
    if total_interacciones >= 1:
        return "Bajo"
    return "Sin impacto"


def limpiar_metadata_instagram(linea):
    texto = limpiar_linea(linea)

    reemplazos = {
        "Me gustaResponder": " Me gusta Responder",
        "me gustaResponder": " me gusta Responder",
        "Me gustaresponder": " Me gusta Responder",
        "megusta": "me gusta",
    }

    for viejo, nuevo in reemplazos.items():
        texto = texto.replace(viejo, nuevo)

    patrones_finales = [
        r"\s+\d+\s*(s|min|m|h|d|sem|semana|semanas)\s+\d+\s*(me gusta|likes?|like)\s*responder\s*$",
        r"\s+\d+\s*(s|min|m|h|d|sem|semana|semanas)\s+\d+\s*(me gusta|likes?|like)\s*$",
        r"\s+\d+\s*(s|min|m|h|d|sem|semana|semanas)\s*responder\s*$",
        r"\s+\d+\s*(me gusta|likes?|like)\s*responder\s*$",
        r"\s+\d+\s*(me gusta|likes?|like)\s*$",
        r"\s+responder\s*$",
    ]

    limpio = texto

    for patron in patrones_finales:
        limpio = re.sub(patron, "", limpio, flags=re.IGNORECASE).strip()

    return limpiar_linea(limpio)


def es_dato_basura(linea):
    t = limpiar_linea(linea).lower()
    t = t.replace("me gustaresponder", "me gusta responder")
    t = t.replace("megusta", "me gusta")
    t = re.sub(r"\s+", " ", t).strip()

    if not t:
        return True

    basura_exacta = [
        "responder", "ver traducción", "ver traduccion", "me gusta",
        "me gusta responder", "likes", "like", "ocultar respuestas",
        "ver respuestas", "seguir", "más", "mas", "ver más", "ver mas"
    ]

    if t in basura_exacta:
        return True

    if re.fullmatch(r"\d+\s*(s|min|m|h|d|sem|semana|semanas)", t):
        return True

    if re.fullmatch(r"\d+\s*(me gusta|likes?|like)", t):
        return True

    if re.fullmatch(
        r"\d+\s*(s|min|m|h|d|sem|semana|semanas)\s+\d+\s*(me gusta|likes?|like)\s*(responder)?",
        t,
    ):
        return True

    if re.fullmatch(r"\d+\s*(s|min|m|h|d|sem|semana|semanas)\s*responder", t):
        return True

    if re.fullmatch(r"\d+\s*(s|min|m|h|d|sem)\s*\d+\s*(s|min|m|h|d|sem)?", t):
        return True

    if len(t) < 4:
        return True

    letras = re.findall(r"[a-záéíóúñü]+", t)
    utiles = [
        x for x in letras
        if x not in ["s", "m", "min", "h", "d", "sem", "semana", "semanas",
                     "me", "gusta", "likes", "like", "responder"]
    ]

    return len(utiles) == 0


def es_respuesta_arroba(linea):
    """
    Detecta respuestas que empiezan con @usuario.
    Si menciona político, NO se cuenta como simple respuesta.
    """
    t = limpiar_linea(linea).lower()

    if not t.startswith("@"):
        return False

    if detectar_politicos(t) != "Sin mención":
        return False

    return True


def procesar_bloque_comentarios(texto_pegado):
    """
    Devuelve solo comentarios principales.
    Las líneas que empiezan con @usuario se cuentan como respuestas/interacciones
    del comentario principal anterior.
    """
    lineas = texto_pegado.splitlines()
    registros = []
    indice_actual = None

    for linea in lineas:
        linea = limpiar_metadata_instagram(linea)
        linea = limpiar_linea(linea)

        if es_dato_basura(linea):
            continue

        if es_respuesta_arroba(linea):
            if indice_actual is not None:
                registros[indice_actual]["respuestas_detectadas"] += 1
                registros[indice_actual]["respuestas_texto"].append(linea)
            continue

        registros.append({
            "comentario": linea,
            "respuestas_detectadas": 0,
            "respuestas_texto": []
        })
        indice_actual = len(registros) - 1

    return registros


def cargar_excel():
    columnas = [
        "fecha_carga",
        "fuente",
        "publicacion",
        "usuario",
        "comentario",
        "sentimiento",
        "tema",
        "perfil_usuario",
        "barrio_estimado",
        "politicos_mencionados",
        "me_gustas",
        "respuestas_detectadas",
        "total_interacciones",
        "nivel_impacto",
        "respuestas_texto",
        "observaciones",
    ]

    if Path(ARCHIVO_EXCEL).exists():
        df = pd.read_excel(ARCHIVO_EXCEL)

        for col in columnas:
            if col not in df.columns:
                df[col] = ""

        return df

    return pd.DataFrame(columns=columnas)


def guardar_excel(df):
    df.to_excel(ARCHIVO_EXCEL, index=False)

# ======================================================
# APP
# ======================================================

st.title("Cargador de comentarios - La Plata")
st.caption("Carga comentarios y cuenta interacciones detectadas por respuestas @usuario.")

df_existente = cargar_excel()

m1, m2, m3 = st.columns(3)
m1.metric("Comentarios guardados", len(df_existente))
m2.metric("Archivo", ARCHIVO_EXCEL)
m3.metric("Fecha", datetime.now().strftime("%d/%m/%Y"))

st.markdown('<div class="card">', unsafe_allow_html=True)

col_a, col_b, col_c = st.columns(3)

with col_a:
    fuente = st.selectbox("Fuente", ["Instagram", "Facebook", "TikTok", "Twitter/X", "Medio digital", "Otro"])

with col_b:
    publicacion = st.text_input("Publicación / tema", placeholder="Ej: nota sobre comercios de La Plata")

with col_c:
    usuario_manual = st.text_input("Usuario opcional", placeholder="Si todos son del mismo usuario")

texto_pegado = st.text_area(
    "Pegá acá los comentarios",
    height=260,
    placeholder="Pegá comentarios copiados de Instagram/Facebook.\nLas respuestas @usuario se cuentan como interacciones del comentario anterior."
)

st.markdown('</div>', unsafe_allow_html=True)

if "df_preview" not in st.session_state:
    st.session_state.df_preview = pd.DataFrame()

if st.button("Procesar comentarios"):
    registros = procesar_bloque_comentarios(texto_pegado)
    nuevos = []

    for r in registros:
        comentario = r["comentario"]
        me_gustas = 0
        respuestas_detectadas = int(r["respuestas_detectadas"])
        total_interacciones = calcular_impacto(me_gustas, respuestas_detectadas)

        nuevos.append({
            "fecha_carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "fuente": fuente,
            "publicacion": publicacion,
            "usuario": usuario_manual if usuario_manual else "sin_usuario",
            "comentario": comentario,
            "sentimiento": detectar_sentimiento(comentario),
            "tema": detectar_por_diccionario(comentario, temas, default="otro"),
            "perfil_usuario": detectar_por_diccionario(comentario, perfiles, default="usuario_general"),
            "barrio_estimado": detectar_barrio(comentario),
            "politicos_mencionados": detectar_politicos(comentario),
            "me_gustas": me_gustas,
            "respuestas_detectadas": respuestas_detectadas,
            "total_interacciones": total_interacciones,
            "nivel_impacto": nivel_impacto(total_interacciones),
            "respuestas_texto": " | ".join(r["respuestas_texto"]),
            "observaciones": "",
        })

    st.session_state.df_preview = pd.DataFrame(nuevos)

if not st.session_state.df_preview.empty:
    st.subheader("Vista previa antes de guardar")

    resumen_sent = st.session_state.df_preview["sentimiento"].value_counts()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Negativos", int(resumen_sent.get("Negativo", 0)))
    c2.metric("Neutrales", int(resumen_sent.get("Neutral", 0)))
    c3.metric("Positivos", int(resumen_sent.get("Positivo", 0)))
    c4.metric("Respuestas @ detectadas", int(st.session_state.df_preview["respuestas_detectadas"].sum()))
    c5.metric("Interacciones totales", int(st.session_state.df_preview["total_interacciones"].sum()))

    st.info("Los @usuario no se guardan como comentarios separados: suman en respuestas_detectadas del comentario anterior.")

    df_editado = st.data_editor(
        st.session_state.df_preview,
        use_container_width=True,
        num_rows="dynamic",
        height=460,
        column_config={
            "sentimiento": st.column_config.SelectboxColumn(
                "sentimiento",
                options=["Negativo", "Neutral", "Positivo"],
            ),
            "me_gustas": st.column_config.NumberColumn(
                "me_gustas",
                min_value=0,
                step=1,
            ),
            "respuestas_detectadas": st.column_config.NumberColumn(
                "respuestas_detectadas",
                min_value=0,
                step=1,
            ),
            "nivel_impacto": st.column_config.SelectboxColumn(
                "nivel_impacto",
                options=["Sin impacto", "Bajo", "Medio", "Alto"],
            ),
        },
    )

    df_editado["total_interacciones"] = df_editado.apply(
        lambda row: calcular_impacto(row.get("me_gustas", 0), row.get("respuestas_detectadas", 0)),
        axis=1
    )

    df_editado["nivel_impacto"] = df_editado["total_interacciones"].apply(nivel_impacto)

    col_guardar, col_descartar = st.columns(2)

    with col_guardar:
        if st.button("Guardar en Excel"):
            df_actual = cargar_excel()
            df_final = pd.concat([df_actual, df_editado], ignore_index=True)
            guardar_excel(df_final)
            st.success(f"Guardado. Ahora el Excel tiene {len(df_final)} comentarios.")
            st.session_state.df_preview = pd.DataFrame()

    with col_descartar:
        if st.button("Descartar vista previa"):
            st.session_state.df_preview = pd.DataFrame()
            st.warning("Vista previa descartada.")

st.subheader("Base actual")

if not df_existente.empty:
    st.dataframe(df_existente.tail(50), use_container_width=True, height=350)
else:
    st.write("Todavía no hay comentarios guardados.")
