"""Interfaz Streamlit: cargar Excel, ingresar uno o varios rangos con su peso, calcular y descargar."""
from datetime import datetime

import streamlit as st

import calculo as C

st.set_page_config(page_title="Tenor de nuevo lote", page_icon="⛏️", layout="centered")
st.title("⛏️ Tenor de un nuevo lote compuesto")
st.caption(f"Lee únicamente la hoja **{C.HOJA}** · columnas **{C.COL_LOTE}** y **{C.COL_TENOR}**")

if "n_rangos" not in st.session_state:
    st.session_state["n_rangos"] = 2


def _mas():
    st.session_state["n_rangos"] += 1


def _menos():
    st.session_state["n_rangos"] = max(2, st.session_state["n_rangos"] - 1)


# 1) Archivo ------------------------------------------------------------------
archivo = st.file_uploader("1. Carga el archivo Excel", type=["xlsx", "xlsm"])
df = None
if archivo is not None:
    try:
        df = C.leer_hoja_oficial(archivo)
        st.success(f"Hoja {C.HOJA} leída: {len(df)} lotes encontrados.")
    except C.ErrorDatos as e:
        st.error(str(e))

# 2) Entradas ---------------------------------------------------------------------
st.subheader("2. Datos del nuevo lote")
tipo = st.radio(
    "¿Cómo se compone el nuevo lote?",
    ["Un solo rango o lista (un peso total, repartido en partes iguales entre sus lotes)",
     "Varios rangos, cada uno con su propia cantidad de material"],
)
multiple = tipo.startswith("Varios")

modo_txt = st.radio(
    "Los números ingresados corresponden a:",
    ["Número de LOTE (valor de la columna LOTE)", "Número de FILA de la hoja de Excel"],
    horizontal=True,
)
modo = "lote" if modo_txt.startswith("Número de LOTE") else "fila"

grupos_in: list[tuple[str, str]] = []
if not multiple:
    t_lotes = st.text_input("Lotes anteriores que componen el nuevo lote",
                            placeholder="601-607   ó   601, 605, 610")
    t_peso = st.text_input("Peso aproximado del nuevo lote", placeholder="1500 ó 1500,5")
    grupos_in.append((t_lotes, t_peso))
else:
    st.caption("Cada rango acepta rangos y listas, p. ej. `601-612` o `601, 605, 610-615`. "
               "Un lote no puede estar en dos rangos.")
    for i in range(st.session_state["n_rangos"]):
        c1, c2 = st.columns([3, 2])
        t_lotes = c1.text_input(f"Rango {i + 1} – lotes", key=f"lotes_{i}",
                                placeholder="601-612" if i == 0 else "615-649")
        t_peso = c2.text_input(f"Rango {i + 1} – cantidad de material", key=f"peso_{i}", placeholder="800")
        grupos_in.append((t_lotes, t_peso))
    b1, b2, _ = st.columns([1, 1, 2])
    b1.button("➕ Agregar rango", on_click=_mas)
    b2.button("➖ Quitar último", on_click=_menos, disabled=st.session_state["n_rangos"] <= 2)

c1, c2 = st.columns([2, 1])
unidad = c2.text_input("Unidad del peso (opcional)", placeholder="kg, t…")
id_nuevo = c1.text_input("Identificador del nuevo lote (opcional)",
                         placeholder="Si lo dejas vacío se genera uno automático")

# 3) Cálculo ------------------------------------------------------------------------
if st.button("3. Calcular", type="primary", disabled=df is None):
    try:
        st.session_state["res"] = C.procesar(df, grupos_in, modo)
        st.session_state["id"] = id_nuevo.strip() or "NL-" + datetime.now().strftime("%Y%m%d-%H%M%S")
        st.session_state["unidad"] = unidad.strip()
    except C.ErrorDatos as e:
        st.session_state.pop("res", None)
        st.error(str(e))

res = st.session_state.get("res")
if res:
    st.divider()
    st.subheader("Resultado")
    for a in res.avisos:
        st.warning(a)
    u = f" {st.session_state['unidad']}" if st.session_state["unidad"] else ""
    m1, m2, m3 = st.columns(3)
    m1.metric("Lotes tomados", res.cantidad)
    m2.metric("Peso aproximado total", f"{res.peso_total:,.2f}{u}")
    m3.metric("Tenor nuevo lote", f"{res.tenor_nuevo:.4f} g/t")

    for g in res.grupos:
        st.write(f"**{g.nombre}** ({g.texto}) · {g.cantidad} lotes · peso {g.peso:,.2f}{u} → "
                 f"{', '.join(g.lotes)}" if len(res.grupos) > 1 else
                 f"**Lotes que componen el nuevo lote:** {', '.join(g.lotes)}")

    with st.expander("Detalle del cálculo"):
        st.dataframe(
            {"Rango": [g.nombre for g in res.grupos for _ in g.lotes],
             "Lote": res.lotes,
             "Tenor (g/t)": res.tenores,
             "Peso por lote": [g.peso_por_lote for g in res.grupos for _ in g.lotes],
             "Aporte": [a for g in res.grupos for a in g.aportes]},
            use_container_width=True, hide_index=True,
        )
        if len(res.grupos) > 1:
            st.write("Tenor de cada rango por separado: "
                     + " · ".join(f"{g.nombre}: {g.tenor_grupo:.4f} g/t" for g in res.grupos))
        st.write(f"Σ aportes = {res.suma_aportes:,.4f} → tenor = Σ aportes / peso total = "
                 f"**{res.tenor_nuevo:.4f} g/t**")

    # 4) Documento
    doc = C.generar_documento(res, st.session_state["id"], st.session_state["unidad"])
    st.download_button(
        "4. Descargar documento (Hojas de cálculo)",
        data=doc,
        file_name=f"Nuevo_lote_{st.session_state['id']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Para Google Sheets: sube el archivo a Google Drive → clic derecho → Abrir con → Hojas de cálculo de Google.")
