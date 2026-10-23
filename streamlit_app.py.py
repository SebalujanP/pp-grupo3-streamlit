"""
Cancelación y reinscripción de profesionales  ·  PP Grupo 3
App de Streamlit que recorre, paso a paso, el análisis del notebook PP_Grupo3.ipynb.

Ejecutar:   streamlit run app.py
"""
from io import BytesIO
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from matplotlib.patches import Patch
from matplotlib.ticker import PercentFormatter
from scipy import stats

# ──────────────────────────────────────────────────────────────────────────────
# Configuración general
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Cancelación y reinscripción de profesionales",
                   page_icon="📊", layout="wide")

FECHA_REF = pd.Timestamp("2026-07-01")  # fecha a la que se calcula la edad
ORDEN_RANGOS = ["Menores de 30", "30 a 39", "40 a 49", "50 a 59", "60 a 69", "70 y más"]
FOCO_DEFECTO = ["40 a 49", "50 a 59"]
ORDEN_TRAY = ["Activo, nunca canceló", "Activo, canceló y volvió",
              "No activo, volvió y se fue otra vez", "No activo, no volvió"]
COL_CANC = {"Sí canceló": "#d62728", "Nunca canceló": "#2ca02c"}
COL_SEX = {"Femenino": "#9467bd", "Masculino": "#17becf"}

NOMBRE_ACTIVOS = "Base de profesionales activos  v4.xlsx"
NOMBRE_CANCELADOS = "Base de prof con canc y susp v4.xlsx"

plt.rcParams.update({"figure.dpi": 100, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.titlesize": 12, "axes.titleweight": "bold"})


# ──────────────────────────────────────────────────────────────────────────────
# Carga y preparación (misma lógica que el notebook, secciones 2 a 5)
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Leyendo los Excel…")
def leer_excel(contenido: bytes) -> pd.DataFrame:
    # header=4: el encabezado real de las columnas está en la fila 5
    return pd.read_excel(BytesIO(contenido), header=4)


@st.cache_data(show_spinner="Unificando y limpiando las bases…")
def preparar(activos: pd.DataFrame, cancelados: pd.DataFrame):
    """Devuelve (df, info) con la base unificada, limpia y con todas las variables creadas."""
    info = {}
    activos = activos.rename(columns={"Cantidad total préstamos": "Cantidad prestamos"}).copy()
    cancelados = cancelados.copy()
    activos["Estado"] = "Activo"
    cancelados = cancelados.drop(columns=["Jubilado"], errors="ignore")

    info["n_activos"], info["n_cancelados"] = len(activos), len(cancelados)
    info["repetidos"] = len(set(activos["Nro interno"]) & set(cancelados["Nro interno"]))

    df = pd.concat([activos, cancelados], ignore_index=True, sort=False)
    info["n_total"], info["id_unico"] = len(df), bool(df["Nro interno"].is_unique)

    # Fecha de nacimiento: anular fechas de relleno (1899/1901) y calcular edad al 01/07/2026
    df["Fecha nacimiento"] = pd.to_datetime(df["Fecha nacimiento"], errors="coerce")
    relleno = df["Fecha nacimiento"].dt.year <= 1901
    info["fechas_relleno"] = int(relleno.sum())
    df.loc[relleno, "Fecha nacimiento"] = pd.NaT
    nac = df["Fecha nacimiento"]
    df["Edad"] = (FECHA_REF.year - nac.dt.year
                  - ((nac.dt.month > FECHA_REF.month)
                     | ((nac.dt.month == FECHA_REF.month) & (nac.dt.day > FECHA_REF.day))).astype(float))
    df["Rango etario"] = pd.cut(df["Edad"], bins=[0, 30, 40, 50, 60, 70, 200],
                                labels=ORDEN_RANGOS, right=False)
    df["Rango etario"] = df["Rango etario"].cat.add_categories("Sin dato de edad").fillna("Sin dato de edad")

    # Deuda de aportes → número
    df["Deuda de aportes"] = (df["Deuda de aportes"].astype(str).str.replace(".", "", regex=False)
                              .str.replace(",", ".", regex=False))
    df["Deuda de aportes"] = pd.to_numeric(df["Deuda de aportes"], errors="coerce")

    # Actuaciones: vacío = sin actuaciones registradas
    col_act = [f"Cantidad Actuaciones {a}" for a in range(2021, 2026)]
    df[col_act] = df[col_act].fillna(0)
    df["Actuaciones 2021-2025"] = df[col_act].sum(axis=1)

    # Préstamos: vacío = ningún préstamo
    df["Cantidad prestamos"] = df["Cantidad prestamos"].fillna(0)
    df["Pidio prestamo"] = df["Cantidad prestamos"] > 0
    col_fecha_p = [c for c in df.columns if c.startswith("Fecha prestamo")]
    df["Prestamos vigentes"] = df[col_fecha_p].notna().sum(axis=1)

    for c in ["Localidad", "Delegacion de pertenencia", "Provincia",
              "Titulo", "Segundo titulo", "Tercer titulo", "Cuarto titulo"]:
        df[c] = df[c].astype("string").str.strip()

    df["Es activo"] = df["Estado"].eq("Activo")
    df["Tiene deuda"] = df["Deuda de aportes"].fillna(0) > 0
    df["Jubilado"] = df["Tipo jubilacion"].notna()

    # Sexo
    mapa = {"F": "Femenino", "FEMENINO": "Femenino", "MUJER": "Femenino",
            "M": "Masculino", "MASCULINO": "Masculino", "VARON": "Masculino",
            "VARÓN": "Masculino", "HOMBRE": "Masculino"}
    df["Sexo"] = df["Sexo"].astype("string").str.strip().str.upper().map(mapa).fillna("Sin dato")
    df["Femenino"] = np.where(df["Sexo"] == "Femenino", 1.0, np.where(df["Sexo"] == "Masculino", 0.0, np.nan))

    # Reinscripción, cancelación y trayectoria
    df["Cantidad reinscripcion"] = pd.to_numeric(df["Cantidad reinscripcion"], errors="coerce").fillna(0)
    df["Reinscripto"] = df["Cantidad reinscripcion"] > 0
    df["Cancelo"] = (~df["Es activo"]) | df["Reinscripto"]
    df["Cancelo_num"] = df["Cancelo"].astype(int)
    df["Canceló"] = np.where(df["Cancelo"], "Sí canceló", "Nunca canceló")
    df["Trayectoria"] = np.select(
        [df["Es activo"] & ~df["Reinscripto"], df["Es activo"] & df["Reinscripto"],
         ~df["Es activo"] & df["Reinscripto"], ~df["Es activo"] & ~df["Reinscripto"]],
        ORDEN_TRAY, default="Sin clasificar")

    df["Fecha matricula"] = pd.to_datetime(df["Fecha matricula"], errors="coerce")
    df["Años matricula"] = (FECHA_REF - df["Fecha matricula"]).dt.days / 365.25
    df["Deuda"] = df["Deuda de aportes"].fillna(0)
    df["Sin actuaciones"] = df["Actuaciones 2021-2025"] == 0
    return df, info


# ──────────────────────────────────────────────────────────────────────────────
# Funciones auxiliares (tablas y gráficos)
# ──────────────────────────────────────────────────────────────────────────────
def cruce_cancelo(data, by):
    """Conteo de Sí / No canceló según una o más variables y el % que canceló."""
    idx = ([data[c].astype("string").fillna("Sin dato").astype(str) for c in by]
           if isinstance(by, list) else data[by].astype("string").fillna("Sin dato").astype(str))
    t = pd.crosstab(idx, data["Cancelo"]).reindex(columns=[False, True], fill_value=0)
    t.columns = ["Nunca canceló", "Sí canceló"]
    t["Total"] = t.sum(axis=1)
    t["% canceló"] = t["Sí canceló"] / t["Total"]
    return t


def tasa_volvio(data, by):
    """Entre los que cancelaron: cuántos se reinscribieron y su %."""
    d = data.copy()
    keys = by if isinstance(by, list) else [by]
    for c in keys:
        d[c] = d[c].astype("string").fillna("Sin dato").astype(str)
    g = d.groupby(keys)["Reinscripto"]
    t = pd.DataFrame({"Total": g.size(), "Volvieron": g.sum()})
    t["% volvió"] = t["Volvieron"] / t["Total"]
    return t


def ordenar_rangos(t, con_sexo=False):
    if con_sexo:
        nuevo = pd.MultiIndex.from_product([["Femenino", "Masculino"], ORDEN_RANGOS], names=t.index.names)
    else:
        nuevo = pd.Index(ORDEN_RANGOS, name=t.index.name)
    return t.reindex(nuevo).dropna(subset=["Total"])


def pct_fmt(t, cols):
    f = t.copy()
    for c in cols:
        f[c] = f[c].map("{:.1%}".format)
    return f


def agrupar_top(s, n=8):
    s = s.astype("string").fillna("Sin dato")
    top = s.value_counts().head(n).index
    return s.where(s.isin(top), "Otros")


def mostrar(fig):
    st.pyplot(fig, clear_figure=True)
    plt.close(fig)


def nota(texto):
    st.caption(texto)


def serie(t, col, sx):
    return [t.loc[(sx, r), col] if (sx, r) in t.index else 0 for r in ORDEN_RANGOS]


# ── Gráficos de la sección «delegación y título» ──────────────────────────────
def panel_sexo(ax, data, orden, col, sexo, ref, mostrar_y, min_n):
    d = data[data["Sexo"] == sexo].copy()
    d[col] = d[col].astype("string").fillna("Sin dato").astype(str)
    g = d.groupby(col)["Cancelo"].agg(["mean", "size", "sum"]).reindex(orden)
    for i, (_, f) in enumerate(g.iterrows()):
        if pd.isna(f["size"]):
            continue
        p, m, si = f["mean"], int(f["size"]), int(f["sum"])
        poco = m < min_n
        ax.barh(i, p, color=COL_SEX[sexo], alpha=0.35 if poco else 1.0)
        ax.text(p + 0.01, i, f"{p:.0%}  (sí {si:,} · no {m - si:,})" + ("*" if poco else ""),
                va="center", fontsize=8)
    ax.set_yticks(range(len(orden)))
    ax.set_yticklabels(orden, fontsize=9)
    ax.set_ylim(len(orden) - 0.5, -0.5)
    if not mostrar_y:
        ax.tick_params(axis="y", labelleft=False)
    ax.axvline(ref, color="k", ls="--", lw=1)
    ax.xaxis.set_major_formatter(PercentFormatter(1.0))
    ax.set_xlim(0, 1.2)
    ax.set_title(sexo, color=COL_SEX[sexo], fontweight="bold")


def panel_edad(ax, data, orden, col, min_n):
    d = data.copy()
    d[col] = d[col].astype("string").fillna("Sin dato").astype(str)
    pct = (pd.crosstab(d[col], d["Rango etario"], values=d["Cancelo_num"], aggfunc="mean")
           .reindex(index=orden, columns=ORDEN_RANGOS))
    cnt = (pd.crosstab(d[col], d["Rango etario"]).reindex(index=orden, columns=ORDEN_RANGOS).fillna(0))
    vals = pct.where(cnt >= min_n)
    cmap = plt.cm.Reds.copy()
    cmap.set_bad("#eeeeee")
    ax.imshow(np.ma.masked_invalid(vals.values.astype(float)), cmap=cmap, vmin=0, vmax=1, aspect="auto")
    for i in range(len(orden)):
        for j in range(len(ORDEN_RANGOS)):
            v = vals.iloc[i, j]
            if pd.isna(v):
                ax.text(j, i, "–", ha="center", va="center", color="#999999", fontsize=9)
            else:
                ax.text(j, i, f"{v:.0%}", ha="center", va="center", fontsize=8,
                        color="white" if v > 0.55 else "black")
    ax.set_xticks(range(len(ORDEN_RANGOS)))
    ax.set_xticklabels(ORDEN_RANGOS, rotation=40, ha="right", fontsize=8)
    ax.set_yticks(range(len(orden)))
    ax.tick_params(axis="y", labelleft=False)
    ax.set_title("Por edad (ambos sexos)", fontweight="bold")


def figura_completa(data_barras, data_edad, col, titulo, ref, min_n, n=12, min_cat=0):
    d = data_barras.copy()
    d[col] = d[col].astype("string").fillna("Sin dato").astype(str)
    tot = d.groupby(col)["Cancelo"].agg(["mean", "size"])
    tot = tot[tot["size"] >= min_cat].sort_values("size", ascending=False).head(n)
    orden = list(tot.sort_values("mean", ascending=False).index)
    if not orden:
        st.warning("Ninguna categoría alcanza el tamaño mínimo elegido en la barra lateral.")
        return
    alto = max(3.0, 0.45 * len(orden) + 1.5)
    fig, axes = plt.subplots(1, 3, figsize=(17, alto), gridspec_kw={"width_ratios": [3, 3, 2.2]})
    panel_sexo(axes[0], data_barras, orden, col, "Femenino", ref, True, min_n)
    panel_sexo(axes[1], data_barras, orden, col, "Masculino", ref, False, min_n)
    panel_edad(axes[2], data_edad, orden, col, min_n)
    axes[0].text(ref, -0.9, f" total base {ref:.0%}", fontsize=8, va="bottom")
    fig.suptitle(f"{titulo}   (* y celdas grises = menos de {min_n} personas)",
                 fontweight="bold", x=0.01, ha="left")
    plt.tight_layout()
    mostrar(fig)


def perfil(data, by):
    g = data.groupby(by, observed=True)
    out = g.agg(Personas=("Nro interno", "count"),
                Pct_pidio_prestamo=("Pidio prestamo", "mean"),
                Pct_con_deuda=("Tiene deuda", "mean"),
                Act_promedio=("Actuaciones 2021-2025", "mean"),
                Act_mediana=("Actuaciones 2021-2025", "median"),
                Pct_sin_actuaciones=("Sin actuaciones", "mean"))
    out["Deuda_mediana (si debe)"] = data[data["Tiene deuda"]].groupby(by, observed=True)["Deuda"].median()
    return out


def comparar(data, variables):
    filas = []
    for v in variables:
        d = data[[v, "Reinscripto"]].dropna()
        a = d.loc[d["Reinscripto"], v].astype(float)
        b = d.loc[~d["Reinscripto"], v].astype(float)
        if len(a) < 2 or len(b) < 2:
            continue
        filas.append({"Variable": v,
                      "Volvió: promedio": a.mean(), "No volvió: promedio": b.mean(),
                      "Volvió: mediana": a.median(), "No volvió: mediana": b.median()})
    return pd.DataFrame(filas).set_index("Variable")


# ──────────────────────────────────────────────────────────────────────────────
# Secciones de la app
# ──────────────────────────────────────────────────────────────────────────────
def pregunta(texto):
    st.info(f"**Pregunta que responde esta sección:** {texto}")


def sec_intro(df, info):
    st.title("Cancelación y reinscripción de profesionales")
    st.markdown(
        "Esta aplicación presenta, paso a paso, el análisis del notebook **PP_Grupo3**. "
        "Se unifican dos bases de datos (profesionales **activos** y **cancelados/suspendidos**) y se "
        "sigue siempre el mismo esquema, de lo general a lo particular. "
        "Usá el menú de la izquierda para recorrer cada paso."
    )
    st.graphviz_chart("""
    digraph {
      rankdir=LR; node [shape=box, style="rounded,filled", fillcolor="#eef3fb", color="#4c72b0", fontname="Helvetica"];
      a [label="¿Canceló?\\nSí / No"]; b [label="Sexo\\nF / M"]; c [label="Rango\\netario"];
      d [label="Delegación\\ny título"]; e [label="¿Por qué canceló?\\npréstamos · deuda · actuaciones"];
      f [label="¿Volvió?"]; g [label="¿Por qué volvió?"];
      a -> b -> c -> d -> e -> f -> g;
    }
    """)

    # Resumen dinámico: siempre coincide con los datos cargados
    canc = df[df["Cancelo"]]
    sx = cruce_cancelo(df[df["Sexo"] != "Sin dato"], "Sexo")
    ce = df[df["Rango etario"] != "Sin dato de edad"]
    tv = tasa_volvio(canc[canc["Rango etario"] != "Sin dato de edad"], "Rango etario").reindex(ORDEN_RANGOS).dropna(subset=["Total"])
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Profesionales analizados", f"{len(df):,}")
    c2.metric("Canceló alguna vez", f"{df['Cancelo'].mean():.1%}")
    c3.metric("De los que cancelaron, volvieron", f"{canc['Reinscripto'].mean():.1%}")
    c4.metric("Edad calculada al", FECHA_REF.strftime("%d/%m/%Y"))

    st.subheader("Resumen de resultados")
    st.markdown(
        f"- **{df['Cancelo'].mean():.1%}** de los profesionales cancelaron al menos una vez.\n"
        f"- Por sexo, cancelaron **{sx.loc['Femenino', '% canceló']:.1%}** de las mujeres y "
        f"**{sx.loc['Masculino', '% canceló']:.1%}** de los varones.\n"
        f"- El rango etario donde más gente vuelve es **{tv['% volvió'].idxmax()}** "
        f"({tv['% volvió'].max():.1%} de los que cancelaron).\n"
        f"- Solo **{canc['Reinscripto'].mean():.1%}** de quienes cancelaron se reinscribió alguna vez."
    )

    st.subheader("Definiciones que usa todo el análisis")
    st.table(pd.DataFrame({
        "Término": ["Canceló = Sí", "Canceló = No", "Volvió", "Edad", "Trayectoria"],
        "Definición": [
            "Hoy figura Cancelado o Suspendido, **o** se reinscribió alguna vez (canceló y volvió, hoy está Activo).",
            "Hoy es Activo y nunca se reinscribió.",
            "Se reinscribió al menos una vez (`Cantidad reinscripcion` > 0).",
            f"Se calcula al {FECHA_REF.strftime('%d/%m/%Y')} para todos, no es la edad a la que cancelaron.",
            "Combina estado actual y reinscripción en 4 grupos (ver paso 1)."]}).set_index("Término"))


def sec_datos(df, info):
    st.header("Datos y limpieza")
    pregunta("¿Con qué datos se trabaja y qué se corrigió antes de analizarlos?")
    c1, c2, c3 = st.columns(3)
    c1.metric("Base de activos", f"{info['n_activos']:,}")
    c2.metric("Base de cancelados y suspendidos", f"{info['n_cancelados']:,}")
    c3.metric("Base unificada", f"{info['n_total']:,}")
    st.markdown(
        f"- Profesionales repetidos entre ambas bases: **{info['repetidos']}**.\n"
        f"- `Nro interno` es único en la base unificada: **{'sí' if info['id_unico'] else 'no'}**.\n"
        f"- Fechas de nacimiento de relleno anuladas (1899/1901, edad imposible): **{info['fechas_relleno']:,}**."
    )
    e = df["Estado"].value_counts().to_frame("Profesionales")
    e["%"] = e["Profesionales"] / e["Profesionales"].sum()
    st.dataframe(pct_fmt(e, ["%"]), use_container_width=False)

    st.subheader("Decisiones de limpieza")
    st.table(pd.DataFrame({
        "Variable": ["Fecha de nacimiento", "Edad / rango etario", "Deuda de aportes", "Actuaciones 2021-2025",
                     "Préstamos", "Sexo", "Reinscripto"],
        "Qué se hizo": [
            "Las fechas de relleno pasan a vacío y esas personas quedan fuera de los rangos.",
            "Edad al 01/07/2026; rangos: <30, 30-39, 40-49, 50-59, 60-69, 70+.",
            "Se convierte a número (se quitan separadores de miles).",
            "Los vacíos se interpretan como «sin actuaciones» y se completan con 0; se suman los 5 años.",
            "Cantidad vacía = no tomó ningún préstamo.",
            "Se unifican los valores (F/M) en Femenino / Masculino.",
            "`Cantidad reinscripcion` > 0 (vacíos = 0)."]}).set_index("Variable"))
    with st.expander("Ver una muestra de la base unificada"):
        cols = ["Nro interno", "Estado", "Sexo", "Edad", "Rango etario", "Delegacion de pertenencia", "Titulo",
                "Cantidad reinscripcion", "Cantidad prestamos", "Deuda", "Actuaciones 2021-2025"]
        st.dataframe(df[cols].head(200), use_container_width=True)


def sec_cancelo(df, info):
    st.header("Paso 1 · ¿Canceló? Sí / No")
    pregunta("¿Cuántos profesionales cancelaron al menos una vez y cuántos nunca lo hicieron?")
    t1 = df["Canceló"].value_counts().reindex(["Sí canceló", "Nunca canceló"]).to_frame("Profesionales")
    t1["%"] = t1["Profesionales"] / t1["Profesionales"].sum()
    t1b = df["Trayectoria"].value_counts().reindex(ORDEN_TRAY).fillna(0).to_frame("Profesionales")
    t1b["%"] = t1b["Profesionales"] / t1b["Profesionales"].sum()

    c1, c2 = st.columns([1, 1.3])
    with c1:
        st.dataframe(pct_fmt(t1, ["%"]))
        st.dataframe(pct_fmt(t1b, ["%"]))
    with c2:
        st.markdown(
            "La **trayectoria** separa a los que hoy están activos en dos grupos (los que nunca cancelaron y "
            "los que cancelaron y volvieron) y a los no activos en otros dos (los que volvieron y se fueron "
            "otra vez, y los que nunca volvieron).")

    fig, axes = plt.subplots(1, 2, figsize=(13, 4))
    tot = t1["Profesionales"].sum()
    axes[0].bar(t1.index, t1["Profesionales"], color=[COL_CANC[i] for i in t1.index])
    for i, v in enumerate(t1["Profesionales"]):
        axes[0].text(i, v + tot * 0.01, f"{v:,}  ({v / tot:.0%})", ha="center")
    axes[0].set_title("¿Canceló alguna vez?")
    vals = t1b["Profesionales"]
    axes[1].barh(range(4), vals.values, color=["#2ca02c", "#f2c14e", "#ff7f0e", "#d62728"])
    axes[1].set_yticks(range(4))
    axes[1].set_yticklabels(ORDEN_TRAY)
    axes[1].invert_yaxis()
    for i, v in enumerate(vals.values):
        axes[1].text(v + vals.max() * 0.01, i, f"{int(v):,}", va="center", fontsize=9)
    axes[1].set_title("Trayectoria")
    axes[1].set_xlim(0, vals.max() * 1.15)
    plt.tight_layout()
    mostrar(fig)


def sec_sexo(df, info):
    st.header("Paso 2 · Por sexo")
    pregunta("¿Cancelan más las mujeres o los varones?")
    con_sexo = df[df["Sexo"] != "Sin dato"]
    t2 = cruce_cancelo(con_sexo, "Sexo")
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.dataframe(pct_fmt(t2, ["% canceló"]))
        nota(f"Sin dato de sexo: {(df['Sexo'] == 'Sin dato').sum():,} profesionales (quedan fuera de este paso).")
    with c2:
        fig, ax = plt.subplots(figsize=(5, 3.6))
        ax.bar(t2.index, t2["% canceló"], color=[COL_SEX[s] for s in t2.index])
        for i, v in enumerate(t2["% canceló"]):
            ax.text(i, v + 0.01, f"{v:.1%}", ha="center", fontweight="bold")
        ax.yaxis.set_major_formatter(PercentFormatter(1.0))
        ax.set_ylim(0, min(1.05, t2["% canceló"].max() + 0.15))
        ax.set_title("% que canceló, según sexo")
        plt.tight_layout()
        mostrar(fig)


def sec_edad(df, info, min_n):
    st.header("Paso 3 · Rango etario dentro de cada sexo")
    pregunta("¿Cómo cambia el % de cancelación con la edad, y es igual en mujeres y varones?")
    con_edad = df[df["Rango etario"] != "Sin dato de edad"]
    t3 = ordenar_rangos(cruce_cancelo(con_edad[con_edad["Sexo"] != "Sin dato"], ["Sexo", "Rango etario"]),
                        con_sexo=True)
    t3["Aviso"] = np.where(t3[["Nunca canceló", "Sí canceló"]].min(axis=1) < min_n, f"grupo < {min_n}", "")
    st.dataframe(pct_fmt(t3, ["% canceló"]), use_container_width=True)
    st.warning(
        f"**Cómo leerlo:** la columna *Aviso* marca las celdas donde alguno de los dos grupos (Sí / No canceló) "
        f"tiene menos de {min_n} personas: hay que tomarlas con cuidado. Además, la edad es la del 01/07/2026 "
        "para todos, incluso para quienes cancelaron hace años.")

    x = np.arange(len(ORDEN_RANGOS))
    w = 0.38
    fig, axes = plt.subplots(1, 3, figsize=(17, 4.2))
    for ax, col in zip(axes[:2], ["Sí canceló", "Nunca canceló"]):
        for k, sx in enumerate(["Femenino", "Masculino"]):
            ax.bar(x + (k - 0.5) * w, serie(t3, col, sx), w, label=sx, color=COL_SEX[sx])
        ax.set_xticks(x)
        ax.set_xticklabels(ORDEN_RANGOS, rotation=30)
        ax.set_title(f"{col}: cantidad")
        ax.legend()
    for k, sx in enumerate(["Femenino", "Masculino"]):
        v = serie(t3, "% canceló", sx)
        axes[2].bar(x + (k - 0.5) * w, v, w, label=sx, color=COL_SEX[sx])
        for xi, vi in zip(x + (k - 0.5) * w, v):
            axes[2].text(xi, vi + 0.01, f"{vi:.0%}", ha="center", fontsize=7)
    axes[2].set_xticks(x)
    axes[2].set_xticklabels(ORDEN_RANGOS, rotation=30)
    axes[2].yaxis.set_major_formatter(PercentFormatter(1.0))
    axes[2].set_title("% que canceló")
    axes[2].legend()
    plt.tight_layout()
    mostrar(fig)


def sec_delegacion(df, info, min_n, min_grupo, foco):
    st.header("Paso 4 · Delegación y título")
    pregunta("¿Cambia el % de cancelación según la delegación de pertenencia o el título?")
    if not foco:
        st.warning("Elegí al menos un rango etario de foco en la barra lateral.")
        return
    st.markdown(
        "Para que la edad no distorsione la comparación (en 70 y más cancelan casi todos) se mira solo a los "
        f"rangos de foco: **{' y '.join(foco)}**. Se muestran delegaciones con al menos **{min_grupo}** personas; "
        "en título se dejan los 8 más frecuentes y el resto va a «Otros».")

    con_edad = df[df["Rango etario"] != "Sin dato de edad"]
    base4 = con_edad[con_edad["Rango etario"].isin(foco)].copy()
    base4["Titulo grupo"] = agrupar_top(base4["Titulo"], 8)
    ref4 = base4["Cancelo"].mean()
    st.metric("% que canceló en la base de foco", f"{ref4:.1%}", help=f"{len(base4):,} profesionales")

    base_edad = con_edad.copy()
    tops = set(base4["Titulo grupo"].unique())
    base_edad["Titulo grupo"] = (base_edad["Titulo"].astype("string").fillna("Sin dato")
                                 .where(lambda s: s.isin(tops), "Otros"))

    t4d = cruce_cancelo(base4, "Delegacion de pertenencia")
    t4d = t4d[t4d["Total"] >= min_grupo].sort_values("% canceló", ascending=False)
    t4t = cruce_cancelo(base4, "Titulo grupo").sort_values("% canceló", ascending=False)

    tab1, tab2 = st.tabs(["Por delegación", "Por título"])
    with tab1:
        figura_completa(base4, base_edad, "Delegacion de pertenencia",
                        "% que canceló por delegación: sexo y edad", ref4, min_n, n=12, min_cat=min_grupo)
        with st.expander("Ver tabla"):
            st.dataframe(pct_fmt(t4d, ["% canceló"]), use_container_width=True)
    with tab2:
        figura_completa(base4, base_edad, "Titulo grupo",
                        "% que canceló por título: sexo y edad", ref4, min_n, n=9)
        with st.expander("Ver tabla"):
            st.dataframe(pct_fmt(t4t, ["% canceló"]), use_container_width=True)
    nota("Cómo leer los gráficos: los dos paneles de barras muestran el % que canceló en mujeres y en varones "
         "(rango de foco); la línea punteada es el total de la base; el mapa de calor muestra el % por rango "
         "etario para todas las edades.")


def sec_motivos(df, info):
    st.header("Paso 5 · ¿Por qué canceló? Préstamos, deudas y actuaciones")
    pregunta("¿Los que cancelaron se parecen o se diferencian de los que nunca cancelaron en préstamos, deuda y actuaciones?")
    st.warning(
        "**Cuidado con las actuaciones:** quienes cancelaron casi no tienen actuaciones porque dejaron de ejercer. "
        "Eso es en gran parte *consecuencia* de la cancelación, no su causa. Para ver si baja *antes* de cancelar "
        "haría falta la fecha de cancelación.")
    con_edad = df[df["Rango etario"] != "Sin dato de edad"]
    pf = perfil(con_edad, ["Canceló", "Rango etario"])
    pf = pf[pf["Personas"] > 0]
    fmt = ["Pct_pidio_prestamo", "Pct_con_deuda", "Pct_sin_actuaciones"]

    x = np.arange(len(ORDEN_RANGOS))
    w = 0.38
    fig, axes = plt.subplots(1, 3, figsize=(17, 4.2))
    for ax, (m, tit) in zip(axes, [("Pct_pidio_prestamo", "% que alguna vez pidió préstamo"),
                                   ("Pct_con_deuda", "% con deuda de aportes"),
                                   ("Pct_sin_actuaciones", "% sin actuaciones (2021-2025)")]):
        for k, c in enumerate(["Sí canceló", "Nunca canceló"]):
            v = [pf.loc[(c, r), m] if (c, r) in pf.index else 0 for r in ORDEN_RANGOS]
            ax.bar(x + (k - 0.5) * w, v, w, label=c, color=COL_CANC[c])
        ax.set_xticks(x)
        ax.set_xticklabels(ORDEN_RANGOS, rotation=30)
        ax.yaxis.set_major_formatter(PercentFormatter(1.0))
        ax.set_title(tit)
        ax.legend(fontsize=8)
    plt.tight_layout()
    mostrar(fig)

    tab1, tab2 = st.tabs(["Por rango etario", "Por sexo"])
    with tab1:
        st.dataframe(pct_fmt(pf, fmt), use_container_width=True)
    with tab2:
        pf_sx = perfil(con_edad[con_edad["Sexo"] != "Sin dato"], ["Canceló", "Sexo"])
        st.dataframe(pct_fmt(pf_sx, fmt), use_container_width=True)


def sec_volvio(df, info):
    st.header("Paso 6 · ¿Volvió?")
    pregunta("De los que cancelaron, ¿cuántos se reinscribieron?")
    st.warning(
        "**Cómo leerlo:** un cancelado con 0 reinscripciones *todavía podría volver*; «no volvió» significa "
        "«no volvió hasta hoy». Además la base no trae fecha de reinscripción, por lo que **no se puede medir "
        "cuánto tardaron en volver**.")
    canc = df[df["Cancelo"]].copy()
    canc_edad = canc[canc["Rango etario"] != "Sin dato de edad"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Cancelaron", f"{len(canc):,}")
    c2.metric("Volvieron", f"{int(canc['Reinscripto'].sum()):,}")
    c3.metric("% que volvió", f"{canc['Reinscripto'].mean():.1%}")

    tv_sr = ordenar_rangos(tasa_volvio(canc_edad[canc_edad["Sexo"] != "Sin dato"], ["Sexo", "Rango etario"]),
                           con_sexo=True)
    x = np.arange(len(ORDEN_RANGOS))
    w = 0.38
    fig, ax = plt.subplots(figsize=(9, 4))
    for k, sx in enumerate(["Femenino", "Masculino"]):
        v = [tv_sr.loc[(sx, r), "% volvió"] if (sx, r) in tv_sr.index else 0 for r in ORDEN_RANGOS]
        ax.bar(x + (k - 0.5) * w, v, w, label=sx, color=COL_SEX[sx])
        for xi, vi in zip(x + (k - 0.5) * w, v):
            ax.text(xi, vi + 0.005, f"{vi:.0%}", ha="center", fontsize=7)
    ax.set_xticks(x)
    ax.set_xticklabels(ORDEN_RANGOS, rotation=30)
    ax.yaxis.set_major_formatter(PercentFormatter(1.0))
    ax.set_title("% de los que cancelaron que volvió, por sexo y rango etario")
    ax.legend()
    plt.tight_layout()
    mostrar(fig)

    t_estado, t_sexo, t_rango = st.tabs(["Por estado actual", "Por sexo", "Por rango etario"])
    with t_estado:
        st.dataframe(pct_fmt(tasa_volvio(canc, "Estado"), ["% volvió"]))
        nota("Los «Activos» que cancelaron son, por definición, todos reinscriptos (100%).")
        veces = (canc.loc[canc["Reinscripto"], "Cantidad reinscripcion"].clip(upper=3)
                 .map({1.0: "1 vez", 2.0: "2 veces", 3.0: "3 o más"}).value_counts().to_frame("Profesionales"))
        st.markdown("**Cantidad de reinscripciones** (solo los que volvieron)")
        st.dataframe(veces)
    with t_sexo:
        st.dataframe(pct_fmt(tasa_volvio(canc[canc["Sexo"] != "Sin dato"], "Sexo"), ["% volvió"]))
    with t_rango:
        tv_r = tasa_volvio(canc_edad, "Rango etario").reindex(ORDEN_RANGOS).dropna(subset=["Total"])
        st.dataframe(pct_fmt(tv_r, ["% volvió"]))
        st.markdown("**Por sexo y rango etario**")
        st.dataframe(pct_fmt(tv_sr, ["% volvió"]))


def sec_porque(df, info):
    st.header("Paso 7 · ¿Por qué volvieron? ¿Hay un patrón?")
    pregunta("¿En qué se diferencian los que volvieron de los que no volvieron?")
    st.markdown(
        "Quienes **volvieron y hoy están activos** tienen, casi por definición, más actuaciones y menos deuda que "
        "quienes siguen cancelados. Comparar a ambos grupos en esas variables no dice *por qué* volvieron. "
        "Para evitar ese sesgo se hacen **dos comparaciones**:")
    st.markdown(
        "- **A. Todos los que cancelaron** (volvió vs. no volvió): solo variables que no dependen de la situación "
        "actual, *edad, años de matrícula, jubilado y sexo*.\n"
        "- **B. Solo los que hoy NO están activos** (volvieron y se fueron otra vez vs. nunca volvieron): como los "
        "dos grupos están hoy en la misma situación, se pueden comparar también *préstamos, deuda y actuaciones*.")
    canc = df[df["Cancelo"]].copy()
    for c in ["Jubilado", "Pidio prestamo", "Tiene deuda"]:
        canc[c] = canc[c].astype(float)
    vars_a = ["Edad", "Años matricula", "Jubilado", "Femenino"]
    vars_b = vars_a + ["Pidio prestamo", "Cantidad prestamos", "Tiene deuda", "Deuda", "Actuaciones 2021-2025"]

    st.subheader(f"A. Todos los que cancelaron ({len(canc):,})")
    st.dataframe(comparar(canc, vars_a).round(3), use_container_width=True)
    canc_na = canc[~canc["Es activo"]]
    st.subheader(f"B. Solo los que hoy NO están activos ({len(canc_na):,}; "
                 f"volvieron y se fueron otra vez: {int(canc_na['Reinscripto'].sum()):,})")
    st.dataframe(comparar(canc_na, vars_b).round(3), use_container_width=True)
    nota("Las variables 0/1 (Jubilado, Femenino, Pidió préstamo, Tiene deuda) se leen como proporción: "
         "0,25 = 25% del grupo.")


def sec_tabla(df, info):
    st.header("Tabla final")
    pregunta("¿Cómo queda una fila por profesional con todas las variables del análisis?")
    st.markdown(
        "Una fila por profesional: ID, título, delegación, reinscripciones, meses activos, préstamos, deuda y "
        "actuaciones. **Meses activos** se cuenta desde la matriculación hasta hoy solo para los activos; en los "
        "no activos queda vacío porque el notebook no usa ninguna fecha de baja.")
    fin = pd.Series(FECHA_REF, index=df.index).where(df["Es activo"], pd.NaT)
    meses = ((fin - df["Fecha matricula"]).dt.days / 30.4375).clip(lower=0)
    tabla = pd.DataFrame({
        "ID": df["Nro interno"], "Título": df["Titulo"], "Sexo": df["Sexo"],
        "Delegación": df["Delegacion de pertenencia"], "Estado": df["Estado"],
        "Reinscripciones": df["Cantidad reinscripcion"].astype(int),
        "Reinscripto": np.where(df["Reinscripto"], "Sí", "No"),
        "Meses activos": meses.round(1),
        "Cantidad de préstamos": df["Cantidad prestamos"].astype(int),
        "Deuda ($)": df["Deuda"],
        "Actuaciones (2021-2025)": df["Actuaciones 2021-2025"].astype(int)})
    c1, c2 = st.columns(2)
    estados = c1.multiselect("Estado", sorted(tabla["Estado"].unique()), default=sorted(tabla["Estado"].unique()))
    reins = c2.multiselect("Reinscripto", ["Sí", "No"], default=["Sí", "No"])
    vista = tabla[tabla["Estado"].isin(estados) & tabla["Reinscripto"].isin(reins)]
    st.caption(f"{len(vista):,} de {len(tabla):,} filas")
    st.dataframe(vista.head(1000), use_container_width=True)
    if len(vista) > 1000:
        nota("Se muestran las primeras 1.000 filas; la descarga incluye todas las filtradas.")
    st.download_button("Descargar CSV", vista.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                       file_name="tabla_final.csv", mime="text/csv")


# ──────────────────────────────────────────────────────────────────────────────
# Barra lateral y navegación
# ──────────────────────────────────────────────────────────────────────────────
def obtener_bytes(subido, nombre_local):
    if subido is not None:
        return subido.getvalue()
    for carpeta in (Path(__file__).parent, Path(__file__).parent / "data"):
        p = carpeta / nombre_local
        if p.exists():
            return p.read_bytes()
    return None


def main():
    with st.sidebar:
        st.header("📂 Datos")
        up_act = st.file_uploader("Base de profesionales activos (.xlsx)", type="xlsx")
        up_can = st.file_uploader("Base de cancelados y suspendidos (.xlsx)", type="xlsx")
        st.caption("Si los archivos están en la misma carpeta que app.py (o en `data/`), se cargan solos.")

    b_act, b_can = obtener_bytes(up_act, NOMBRE_ACTIVOS), obtener_bytes(up_can, NOMBRE_CANCELADOS)
    if b_act is None or b_can is None:
        st.title("Cancelación y reinscripción de profesionales")
        st.info("Subí las dos bases de datos en la barra lateral para comenzar.")
        st.stop()

    df, info = preparar(leer_excel(b_act), leer_excel(b_can))

    with st.sidebar:
        st.header("🧭 Navegación")
        pasos = {
            "Introducción": sec_intro, "Datos y limpieza": sec_datos,
            "1 · ¿Canceló?": sec_cancelo, "2 · Sexo": sec_sexo, "3 · Rango etario": sec_edad,
            "4 · Delegación y título": sec_delegacion, "5 · ¿Por qué canceló?": sec_motivos,
            "6 · ¿Volvió?": sec_volvio, "7 · ¿Por qué volvió?": sec_porque, "Tabla final": sec_tabla}
        paso = st.radio("Sección", list(pasos), label_visibility="collapsed")
        st.header("⚙️ Parámetros")
        min_n = st.number_input("Mínimo de personas por grupo (aviso)", 1, 500, 30,
                                help="Los grupos más chicos se marcan con * o en gris.")
        min_grupo = st.number_input("Mínimo por delegación", 1, 2000, 100,
                                    help="Delegaciones con menos personas no se muestran en el paso 4.")
        foco = st.multiselect("Rangos de foco (paso 4)", ORDEN_RANGOS, default=FOCO_DEFECTO)

    if paso == "4 · Delegación y título":
        sec_delegacion(df, info, min_n, min_grupo, foco)
    elif paso == "3 · Rango etario":
        sec_edad(df, info, min_n)
    else:
        pasos[paso](df, info)


main()
