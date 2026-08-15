#!/usr/bin/env python3
"""
Dashboard Ejecutivo · Retail Omnicanal · Reto A
¿Dónde debe concentrarse el crecimiento por región, canal y categoría
para aumentar ventas sin destruir margen ni incumplir metas?

Requisitos cumplidos (PROYECTO FINAL):
- Resumen Ejecutivo
- Diagnóstico
- KPIs de gestión
- Insights (mín. 3)
- Recomendaciones concretas
- Plan de acción
- Datos limpios reproducibles (no se modifica RAW a mano)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path

RUTA_CASO = Path("Caso_1_Retail_Omnicanal.xlsx")
RUTA_ENTREGA = Path("Caso_1_Retail_Omnicanal_Entrega.xlsx")

PREGUNTAS_NEGOCIO = [
    "¿Qué regiones y canales concentran ventas y margen sin destruir rentabilidad?",
    "¿El cumplimiento de meta es real o la base/metas están mal calibradas?",
    "¿Qué combinaciones Región×Canal×Categoría crecen con margen y baja devolución?",
    "¿Los descuentos están erosionando margen más de lo que aportan en volumen?",
    "¿El inventario disponible respalda o limita la expansión propuesta?",
]

# ──────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Retail Omnicanal · Dashboard Ejecutivo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Colores corporativos
COLORES = {
    "primario": "#1B4F72",
    "secundario": "#2874A6",
    "acento": "#148F77",
    "alerta": "#C0392B",
    "advertencia": "#D68910",
    "neutro": "#5D6D7E",
    "fondo": "#F8F9F9",
}

# ──────────────────────────────────────────────
# CARGA DE DATOS (cacheada)
# ──────────────────────────────────────────────
def _tabla_excel_a_texto(df: pd.DataFrame) -> pd.DataFrame:
    """Evita errores de PyArrow cuando el Excel mezcla textos y números en una columna."""
    return df.fillna("").astype(str)


@st.cache_data(show_spinner="Cargando datos desde el Excel de entrega…")
def cargar_datos():
    if not RUTA_ENTREGA.exists():
        from generar_entregables import generar_excel

        generar_excel(RUTA_ENTREGA)

    ventas = pd.read_excel(RUTA_ENTREGA, sheet_name="02_Ventas_Limpias")
    eliminadas = pd.read_excel(RUTA_ENTREGA, sheet_name="03_Ventas_Eliminadas")
    metas = pd.read_excel(RUTA_ENTREGA, sheet_name="04_Metas_Limpias", header=2)
    
    # Cargar auditoría de las 14 combinaciones
    try:
        auditoria_metas = pd.read_excel(RUTA_ENTREGA, sheet_name="04b_Auditoria_14_Metas", header=2)
    except Exception:
        from preparar_datos import construir_auditoria_14_metas
        auditoria_metas = construir_auditoria_14_metas()

    inventario = pd.read_excel(RUTA_ENTREGA, sheet_name="05_Inventario_Mensual", header=2)
    registro_calidad = pd.read_excel(RUTA_ENTREGA, sheet_name="06_Registro_y_Analisis", header=None)
    kpis_tabla = pd.read_excel(RUTA_ENTREGA, sheet_name="09_KPI_Detalle")
    plan_accion = pd.read_excel(RUTA_ENTREGA, sheet_name="08_Plan_Accion")
    resumen_eliminados = pd.read_excel(RUTA_ENTREGA, sheet_name="03b_Resumen_Eliminados", header=None)

    ventas["Fecha_Limpia"] = pd.to_datetime(ventas["Fecha_Limpia"], errors="coerce")
    ventas["Año"] = ventas["Fecha_Limpia"].dt.year
    ventas["Mes_Num"] = ventas["Fecha_Limpia"].dt.month
    ventas["Año_Mes"] = ventas["Fecha_Limpia"].dt.to_period("M").astype(str)

    metas = metas.dropna(subset=["Año_Mes", "Region_Limpia", "Canal_Limpio"])
    col_meta = "Meta_Corregida_BOB" if "Meta_Corregida_BOB" in metas.columns else "Meta_BOB"
    metas["Meta_BOB"] = pd.to_numeric(metas[col_meta], errors="coerce")

    return ventas, metas, auditoria_metas, eliminadas, inventario, registro_calidad, kpis_tabla, plan_accion, resumen_eliminados


ventas, metas, auditoria_metas, ventas_eliminadas, inventario, registro_calidad, kpis_tabla, plan_accion, resumen_eliminados = cargar_datos()
ventas_validas = ventas.copy()

# ──────────────────────────────────────────────
# SIDEBAR · FILTROS
# ──────────────────────────────────────────────
st.sidebar.title("Filtros del Directorio")
st.sidebar.markdown("---")

regiones = ["Todas"] + sorted(ventas_validas["Region_Limpia"].dropna().unique().tolist())
canales = ["Todos"] + sorted(ventas_validas["Canal_Limpio"].dropna().unique().tolist())
categorias = ["Todas"] + sorted(
    [c for c in ventas_validas["Categoria"].dropna().unique() if c != "No identificada"]
)

filtro_region = st.sidebar.selectbox("Región", regiones, index=0)
filtro_canal = st.sidebar.selectbox("Canal", canales, index=0)
filtro_categoria = st.sidebar.selectbox("Categoría", categorias, index=0)

st.sidebar.markdown("---")
st.sidebar.caption(f"**Fuente activa:** `{RUTA_ENTREGA.name}`")
st.sidebar.caption(f"Ventas limpias: {len(ventas):,} · Eliminadas: {len(ventas_eliminadas):,}")
with st.sidebar.expander("Preguntas de negocio (Reto A)"):
    for i, p in enumerate(PREGUNTAS_NEGOCIO, 1):
        st.markdown(f"{i}. {p}")

# Aplicar filtros
df = ventas_validas.copy()
if filtro_region != "Todas":
    df = df[df["Region_Limpia"] == filtro_region]
if filtro_canal != "Todos":
    df = df[df["Canal_Limpio"] == filtro_canal]
if filtro_categoria != "Todas":
    df = df[df["Categoria"] == filtro_categoria]

# ──────────────────────────────────────────────
# KPIs PRINCIPALES (fórmulas de gestión)
# ──────────────────────────────────────────────
def calcular_kpis(data: pd.DataFrame):
    ventas_netas = data["Venta_Calculada_BOB"].sum()
    utilidad = data["Utilidad_BOB"].sum()
    margen = utilidad / ventas_netas if ventas_netas != 0 else 0

    completados = data[data["Estado_Limpio"] == "Completado"]
    pedidos_comp = completados["Venta_ID"].nunique()
    ticket = completados["Venta_Calculada_BOB"].sum() / pedidos_comp if pedidos_comp else 0

    devueltos = data[data["Estado_Limpio"] == "Devuelto"]["Venta_ID"].nunique()
    tasa_dev = devueltos / (devueltos + pedidos_comp) if (devueltos + pedidos_comp) else 0

    venta_sin_desc = (data["Cantidad"] * data["Precio_Unit_BOB"]).sum()
    desc_pond = 1 - (ventas_netas / venta_sin_desc) if venta_sin_desc else 0

    return {
        "ventas_netas": ventas_netas,
        "utilidad": utilidad,
        "margen": margen,
        "ticket": ticket,
        "tasa_dev": tasa_dev,
        "pedidos_comp": pedidos_comp,
        "unidades": data["Cantidad"].sum(),
        "clientes": data["Cliente_ID"].nunique(),
        "desc_pond": desc_pond,
    }


def metricas_semestre(data: pd.DataFrame, inicio: str, fin: str) -> dict:
    sub = data[(data["Fecha_Limpia"] >= inicio) & (data["Fecha_Limpia"] <= fin)]
    k = calcular_kpis(sub)
    return {"ventas": k["ventas_netas"], "margen": k["margen"]}


def calcular_metricas_globales(data: pd.DataFrame, metas_df: pd.DataFrame) -> dict:
    k = calcular_kpis(data)
    meta_total = metas_df["Meta_BOB"].sum()
    k["cumplimiento"] = k["ventas_netas"] / meta_total if meta_total else 0
    k["meta_total"] = meta_total

    h1 = metricas_semestre(data, "2026-01-01", "2026-06-30")
    h2 = metricas_semestre(data, "2025-07-01", "2025-12-31")
    k["crecimiento_sem"] = (h1["ventas"] - h2["ventas"]) / h2["ventas"] if h2["ventas"] else 0
    k["margen_h1"] = h1["margen"]
    k["margen_h2"] = h2["margen"]

    canal_stats = []
    for canal in data["Canal_Limpio"].dropna().unique():
        sub = data[data["Canal_Limpio"] == canal]
        ck = calcular_kpis(sub)
        canal_stats.append(
            {
                "canal": canal,
                "pct_ventas": ck["ventas_netas"] / k["ventas_netas"] if k["ventas_netas"] else 0,
                "margen": ck["margen"],
            }
        )
    k["canales"] = sorted(canal_stats, key=lambda x: x["margen"], reverse=True)

    reg_dev = []
    for reg in data["Region_Limpia"].dropna().unique():
        sub = data[data["Region_Limpia"] == reg]
        reg_dev.append({"region": reg, "tasa_dev": calcular_kpis(sub)["tasa_dev"]})
    k["region_mayor_dev"] = max(reg_dev, key=lambda x: x["tasa_dev"]) if reg_dev else None

    cat = (
        data.groupby("Categoria")
        .agg(Ventas=("Venta_Calculada_BOB", "sum"), Utilidad=("Utilidad_BOB", "sum"))
        .reset_index()
    )
    top_cat = cat.sort_values("Ventas", ascending=False).iloc[0]
    k["cat_top"] = top_cat["Categoria"]
    k["cat_top_pct"] = top_cat["Ventas"] / k["ventas_netas"] if k["ventas_netas"] else 0
    k["cat_top_margen"] = top_cat["Utilidad"] / top_cat["Ventas"] if top_cat["Ventas"] else 0

    return k


globales = calcular_metricas_globales(ventas_validas, metas)
canal_tienda = next(
    (c for c in globales["canales"] if c["canal"] == "Tienda"),
    {"margen": 0, "pct_ventas": 0},
)
canal_whatsapp = next((c for c in globales["canales"] if c["canal"] == "WhatsApp"), None)

kpis = calcular_kpis(df)

# Meta filtrada (solo región + canal; categoría no aplica)
metas_f = metas.copy()
if filtro_region != "Todas":
    metas_f = metas_f[metas_f["Region_Limpia"] == filtro_region]
if filtro_canal != "Todos":
    metas_f = metas_f[metas_f["Canal_Limpio"] == filtro_canal]
meta_total = metas_f["Meta_BOB"].sum()
cumplimiento = kpis["ventas_netas"] / meta_total if meta_total else 0

# ──────────────────────────────────────────────
# NAVEGACIÓN
# ──────────────────────────────────────────────
pagina = st.sidebar.radio(
    "Navegación",
    [
        "1 · Resumen Ejecutivo",
        "2 · Diagnóstico",
        "3 · Oportunidades de crecimiento",
        "4 · Insights y decisiones",
        "5 · Plan de acción",
        "6 · Calidad de datos",
    ],
)

# ──────────────────────────────────────────────
# PÁGINA 1 · RESUMEN EJECUTIVO
# ──────────────────────────────────────────────
if pagina == "1 · Resumen Ejecutivo":
    st.title("Retail Omnicanal · Dashboard Ejecutivo")
    st.caption("Reto A · ¿Dónde concentrar el crecimiento por región, canal y categoría sin destruir margen?")

    st.markdown(
        f"**Contexto:** análisis sobre **{len(ventas):,} ventas limpias** "
        f"(de {len(ventas) + len(ventas_eliminadas):,} filas RAW; "
        f"{len(ventas_eliminadas):,} excluidas y documentadas en el Excel de entrega)."
    )

    st.markdown("### KPIs de gestión (10 indicadores · base filtrada)")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Ventas netas (BOB)", f"{kpis['ventas_netas']:,.0f}")
    c2.metric("Utilidad (BOB)", f"{kpis['utilidad']:,.0f}")
    c3.metric("Margen %", f"{kpis['margen']:.1%}")
    c4.metric("Ticket promedio", f"{kpis['ticket']:,.0f}")
    c5.metric("Cumplimiento meta*", f"{cumplimiento:.2%}")

    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("Pedidos completados", f"{kpis['pedidos_comp']:,}")
    c7.metric("Tasa devolución", f"{kpis['tasa_dev']:.1%}")
    c8.metric("Unidades netas", f"{kpis['unidades']:,.0f}")
    c9.metric("Clientes únicos", f"{kpis['clientes']:,}")
    c10.metric("Descuento ponderado", f"{kpis['desc_pond']:.1%}")

    st.caption(
        "* El cumplimiento de meta responde a región y canal; el filtro categoría no modifica la meta "
        "porque ese nivel no existe en la fuente."
    )

    st.markdown("---")

    # Dos columnas: ventas por región/canal y por categoría
    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Ventas netas por Región × Canal")
        pivot_rc = (
            df.groupby(["Region_Limpia", "Canal_Limpio"])["Venta_Calculada_BOB"]
            .sum()
            .reset_index()
        )
        fig_rc = px.bar(
            pivot_rc,
            x="Region_Limpia",
            y="Venta_Calculada_BOB",
            color="Canal_Limpio",
            barmode="group",
            labels={"Venta_Calculada_BOB": "Ventas netas (BOB)", "Region_Limpia": "Región"},
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig_rc.update_layout(height=380, margin=dict(t=30, b=40), legend_title_text="Canal")
        st.plotly_chart(fig_rc, use_container_width=True)

    with col_b:
        st.subheader("Ventas y margen por Categoría")
        cat = (
            df.groupby("Categoria")
            .agg(
                Ventas=("Venta_Calculada_BOB", "sum"),
                Utilidad=("Utilidad_BOB", "sum"),
            )
            .reset_index()
        )
        cat = cat[cat["Categoria"] != "No identificada"]
        cat["Margen"] = cat["Utilidad"] / cat["Ventas"]
        fig_cat = make_subplots(specs=[[{"secondary_y": True}]])
        fig_cat.add_trace(
            go.Bar(x=cat["Categoria"], y=cat["Ventas"], name="Ventas", marker_color=COLORES["primario"]),
            secondary_y=False,
        )
        fig_cat.add_trace(
            go.Scatter(
                x=cat["Categoria"],
                y=cat["Margen"],
                name="Margen %",
                mode="lines+markers",
                marker_color=COLORES["acento"],
            ),
            secondary_y=True,
        )
        fig_cat.update_layout(height=380, margin=dict(t=30, b=40), legend=dict(orientation="h"))
        fig_cat.update_yaxes(title_text="Ventas (BOB)", secondary_y=False)
        fig_cat.update_yaxes(title_text="Margen", tickformat=".0%", secondary_y=True)
        st.plotly_chart(fig_cat, use_container_width=True)

    # Tendencia mensual
    st.subheader("Evolución mensual de ventas netas y margen")
    mens = (
        df.groupby("Año_Mes")
        .agg(Ventas=("Venta_Calculada_BOB", "sum"), Utilidad=("Utilidad_BOB", "sum"))
        .reset_index()
        .sort_values("Año_Mes")
    )
    mens["Margen"] = mens["Utilidad"] / mens["Ventas"]
    fig_tend = make_subplots(specs=[[{"secondary_y": True}]])
    fig_tend.add_trace(
        go.Bar(x=mens["Año_Mes"], y=mens["Ventas"], name="Ventas", marker_color=COLORES["secundario"]),
        secondary_y=False,
    )
    fig_tend.add_trace(
        go.Scatter(
            x=mens["Año_Mes"],
            y=mens["Margen"],
            name="Margen %",
            mode="lines+markers",
            line=dict(color=COLORES["acento"], width=2),
        ),
        secondary_y=True,
    )
    fig_tend.update_layout(height=360, margin=dict(t=20), xaxis_tickangle=-45)
    fig_tend.update_yaxes(title_text="Ventas (BOB)", secondary_y=False)
    fig_tend.update_yaxes(title_text="Margen", tickformat=".0%", secondary_y=True)
    st.plotly_chart(fig_tend, use_container_width=True)

    st.info(
        f"**Lectura ejecutiva rápida**  \n"
        f"• Cumplimiento global **{globales['cumplimiento']:.1%}** "
        f"(Bs {globales['ventas_netas']:,.0f} vs meta Bs {globales['meta_total']:,.0f}).  \n"
        f"• Ene–Jun 2026 crece **{globales['crecimiento_sem']:.1%}** vs Jul–Dic 2025; "
        f"margen mejora de **{globales['margen_h2']:.1%}** a **{globales['margen_h1']:.1%}**.  \n"
        f"• Crecer solo donde margen y devoluciones permanezcan dentro de guardarraíles."
    )

    with st.expander("Ver definición de KPIs (fórmula y decisión asociada)"):
        st.dataframe(kpis_tabla, use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────
# PÁGINA 2 · DIAGNÓSTICO
# ──────────────────────────────────────────────
elif pagina == "2 · Diagnóstico":
    st.title("Diagnóstico · Qué está provocando los resultados")

    st.markdown("### 1 · Cumplimiento de meta por Región × Canal")
    st.caption("Meta existe solo a este nivel. Valores < 3 % se muestran en rojo.")

    # Cumplimiento real
    ventas_rc = (
        ventas_validas.groupby(["Region_Limpia", "Canal_Limpio"])["Venta_Calculada_BOB"]
        .sum()
        .reset_index()
        .rename(columns={"Venta_Calculada_BOB": "Ventas"})
    )
    meta_rc = (
        metas.groupby(["Region_Limpia", "Canal_Limpio"])["Meta_BOB"]
        .sum()
        .reset_index()
    )
    cumple = ventas_rc.merge(meta_rc, on=["Region_Limpia", "Canal_Limpio"], how="outer").fillna(0)
    cumple["Cumplimiento"] = cumple["Ventas"] / cumple["Meta_BOB"].replace(0, np.nan)

    pivot_c = cumple.pivot(index="Region_Limpia", columns="Canal_Limpio", values="Cumplimiento")
    fig_heat = px.imshow(
        pivot_c,
        text_auto=".2%",
        color_continuous_scale=["#C0392B", "#F5B041", "#1ABC9C"],
        aspect="auto",
        labels=dict(color="Cumplimiento"),
    )
    fig_heat.update_layout(height=320, margin=dict(t=20))
    st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("### 2 · Margen y tasa de devolución por Canal")
    canal_list = []
    for canal in df["Canal_Limpio"].unique():
        sub = df[df["Canal_Limpio"] == canal]
        k = calcular_kpis(sub)
        canal_list.append(
            {
                "Canal": canal,
                "Ventas": k["ventas_netas"],
                "Margen": k["margen"],
                "Tasa_Dev": k["tasa_dev"],
                "Ticket": k["ticket"],
            }
        )
    canal_df = pd.DataFrame(canal_list).sort_values("Ventas", ascending=False)

    fig_canal = make_subplots(rows=1, cols=2, subplot_titles=("Margen % por canal", "Tasa devolución por canal"))
    fig_canal.add_trace(
        go.Bar(x=canal_df["Canal"], y=canal_df["Margen"], marker_color=COLORES["acento"], name="Margen"),
        row=1, col=1,
    )
    fig_canal.add_trace(
        go.Bar(x=canal_df["Canal"], y=canal_df["Tasa_Dev"], marker_color=COLORES["alerta"], name="Devolución"),
        row=1, col=2,
    )
    fig_canal.update_yaxes(tickformat=".0%", row=1, col=1)
    fig_canal.update_yaxes(tickformat=".0%", row=1, col=2)
    fig_canal.update_layout(height=350, showlegend=False)
    st.plotly_chart(fig_canal, use_container_width=True)

    st.markdown("### 3 · Descuento ponderado y su impacto en margen")
    # Descuento ponderado por ventas
    df_temp = df.copy()
    df_temp["Venta_sin_desc"] = df_temp["Cantidad"] * df_temp["Precio_Unit_BOB"]
    desc_pond = 1 - (df_temp["Venta_Calculada_BOB"].sum() / df_temp["Venta_sin_desc"].sum()) if df_temp["Venta_sin_desc"].sum() else 0
    st.metric("Descuento ponderado actual", f"{desc_pond:.1%}")

    # Margen por tramo de descuento
    df_temp["Tramo_Desc"] = pd.cut(
        df_temp["Descuento_Limpio"],
        bins=[-0.01, 0, 0.05, 0.10, 0.15, 0.25, 1],
        labels=["0 %", "1–5 %", "6–10 %", "11–15 %", "16–25 %", ">25 %"],
    )
    tramo = (
        df_temp.groupby("Tramo_Desc", observed=True)
        .agg(Ventas=("Venta_Calculada_BOB", "sum"), Utilidad=("Utilidad_BOB", "sum"))
        .reset_index()
    )
    tramo["Margen"] = tramo["Utilidad"] / tramo["Ventas"]
    fig_desc = px.bar(
        tramo,
        x="Tramo_Desc",
        y="Margen",
        text=[f"{m:.0%}" for m in tramo["Margen"]],
        color="Margen",
        color_continuous_scale="RdYlGn",
        labels={"Tramo_Desc": "Tramo de descuento", "Margen": "Margen"},
    )
    fig_desc.update_layout(height=320, showlegend=False)
    st.plotly_chart(fig_desc, use_container_width=True)

    tienda = canal_tienda
    reg_alerta = globales["region_mayor_dev"]

    st.warning(
        f"**Alertas que condicionan la decisión**  \n"
        f"• **{reg_alerta['region']}** tiene la mayor tasa regional de devolución "
        f"(**{reg_alerta['tasa_dev']:.1%}**).  \n"
        f"• El inventario dispone de **{len(inventario)}** snapshots únicos; lectura direccional, no operativa.  \n"
        f"• Ene–jun 2026 crece **{globales['crecimiento_sem']:.1%}** y el margen mejora de "
        f"**{globales['margen_h2']:.1%}** a **{globales['margen_h1']:.1%}**.  \n"
        f"• Descuento ponderado actual: **{kpis['desc_pond']:.1%}**; evitar promociones generalizadas."
    )

# ──────────────────────────────────────────────
# PÁGINA 3 · OPORTUNIDADES
# ──────────────────────────────────────────────
elif pagina == "3 · Oportunidades de crecimiento":
    st.title("Oportunidades de crecimiento · Base robusta")
    st.markdown(
        "Combinaciones Región × Canal × Categoría con **≥ 10 pedidos** y **≥ Bs 5.000** "
        "en los últimos 6 meses (Ene–Jun 2026) para evitar conclusiones por muestras pequeñas."
    )

    # Definir periodos
    df_2026h1 = ventas_validas[
        (ventas_validas["Fecha_Limpia"] >= "2026-01-01")
        & (ventas_validas["Fecha_Limpia"] <= "2026-06-30")
    ]
    df_2025h2 = ventas_validas[
        (ventas_validas["Fecha_Limpia"] >= "2025-07-01")
        & (ventas_validas["Fecha_Limpia"] <= "2025-12-31")
    ]

    def resumen_combo(data):
        g = (
            data.groupby(["Region_Limpia", "Canal_Limpio", "Categoria"])
            .agg(
                Ventas=("Venta_Calculada_BOB", "sum"),
                Utilidad=("Utilidad_BOB", "sum"),
                Pedidos=("Venta_ID", "nunique"),
            )
            .reset_index()
        )
        g["Margen"] = g["Utilidad"] / g["Ventas"]
        return g

    r26 = resumen_combo(df_2026h1)
    r25 = resumen_combo(df_2025h2)[["Region_Limpia", "Canal_Limpio", "Categoria", "Ventas"]].rename(
        columns={"Ventas": "Ventas_ant"}
    )
    opp = r26.merge(r25, on=["Region_Limpia", "Canal_Limpio", "Categoria"], how="left")
    opp["Ventas_ant"] = opp["Ventas_ant"].fillna(0)
    opp["Crecimiento"] = np.where(
        opp["Ventas_ant"] > 0, (opp["Ventas"] - opp["Ventas_ant"]) / opp["Ventas_ant"], np.nan
    )

    # Tasa devolución por combo
    def tasa_dev_combo(data):
        rows = []
        for keys, sub in data.groupby(["Region_Limpia", "Canal_Limpio", "Categoria"]):
            k = calcular_kpis(sub)
            rows.append(
                {
                    "Region_Limpia": keys[0],
                    "Canal_Limpio": keys[1],
                    "Categoria": keys[2],
                    "Tasa_Dev": k["tasa_dev"],
                }
            )
        return pd.DataFrame(rows)

    tasa = tasa_dev_combo(df_2026h1)
    opp = opp.merge(tasa, on=["Region_Limpia", "Canal_Limpio", "Categoria"], how="left")

    # Filtro de robustez
    robustas = opp[(opp["Pedidos"] >= 10) & (opp["Ventas"] >= 5000)].copy()
    robustas = robustas.sort_values(["Crecimiento", "Margen"], ascending=[False, False])

    # Priorización
    def decidir(row):
        if row["Margen"] >= 0.31 and row["Tasa_Dev"] <= 0.08 and row["Crecimiento"] > 0.05:
            return "Priorizar"
        elif row["Margen"] >= 0.30 and row["Tasa_Dev"] <= 0.10:
            return "Vigilar"
        else:
            return "No priorizar"

    robustas["Decisión"] = robustas.apply(decidir, axis=1)

    tabla_opp = robustas[
        [
            "Region_Limpia",
            "Canal_Limpio",
            "Categoria",
            "Ventas",
            "Crecimiento",
            "Margen",
            "Tasa_Dev",
            "Pedidos",
            "Decisión",
        ]
    ].rename(
        columns={
            "Region_Limpia": "Región",
            "Canal_Limpio": "Canal",
            "Categoria": "Categoría",
            "Ventas": "Ventas 6m",
            "Crecimiento": "Crecimiento",
            "Margen": "Margen",
            "Tasa_Dev": "Tasa devolución",
            "Pedidos": "Pedidos",
        }
    )

    st.dataframe(
        tabla_opp,
        use_container_width=True,
        height=420,
        column_config={
            "Ventas 6m": st.column_config.NumberColumn(format="%.0f"),
            "Crecimiento": st.column_config.NumberColumn(format="%.1%%"),
            "Margen": st.column_config.NumberColumn(format="%.1%%"),
            "Tasa devolución": st.column_config.NumberColumn(format="%.1%%"),
            "Pedidos": st.column_config.NumberColumn(format="%.0f"),
        },
    )

    # Top priorizar
    priorizar = robustas[robustas["Decisión"] == "Priorizar"].head(5)
    if not priorizar.empty:
        st.subheader("Top combinaciones a priorizar")
        fig_prio = px.scatter(
            priorizar,
            x="Crecimiento",
            y="Margen",
            size="Ventas",
            color="Region_Limpia",
            hover_data=["Canal_Limpio", "Categoria", "Tasa_Dev"],
            labels={"Crecimiento": "Crecimiento", "Margen": "Margen"},
            size_max=50,
        )
        fig_prio.update_layout(height=380)
        fig_prio.update_xaxes(tickformat=".0%")
        fig_prio.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig_prio, use_container_width=True)

# ──────────────────────────────────────────────
# PÁGINA 4 · INSIGHTS Y DECISIONES
# ──────────────────────────────────────────────
elif pagina == "4 · Insights y decisiones":
    st.title("Insights · Hallazgos y decisiones recomendadas")

    wa = canal_whatsapp

    st.markdown(
        f"""
### Insight 1 · La meta no es operativamente alcanzable con el universo observado
**Evidencia:** Bs **{globales['ventas_netas']:,.0f}** de ventas netas frente a Bs **{globales['meta_total']:,.0f}** de meta agregada → **{globales['cumplimiento']:.1%} de cumplimiento**.  
**Significado:** Antes de exigir cierre de brecha a los equipos comerciales, la gerencia debe validar cobertura y calibración de las metas mes–región–canal.  
**Decisión:** Validar si la base representa todo el universo y recalibrar metas (plazo 15 días · BI + Finanzas).
"""
    )

    if wa:
        st.markdown(
            f"""
### Insight 2 · WhatsApp ofrece el mejor margen, pero aún no escala
**Evidencia:** WhatsApp aporta **{wa['pct_ventas']:.1%}** de las ventas con **margen {wa['margen']:.1%}** frente a **{canal_tienda['margen']:.1%}** en Tienda.  
**Significado:** Es el canal de mayor calidad de margen, pero su volumen es todavía pequeño. Una expansión masiva sin control puede deteriorar la ventaja.  
**Decisión:** Piloto controlado de WhatsApp en Electrónica (Cochabamba), con guardarraíles de margen ≥ 35 % y devolución ≤ 8 %.
"""
        )

    st.markdown(
        f"""
### Insight 3 · {globales['cat_top']} es el motor más defendible
**Evidencia:** {globales['cat_top']} concentra **{globales['cat_top_pct']:.1%}** de las ventas con margen **{globales['cat_top_margen']:.1%}**.  
Combinaciones robustas: **Sucre · Tienda · Electrónica**, **Santa Cruz · Tienda · Electrónica** y **Potosí · Tienda · Hogar**.  
**Significado:** Concentrar el crecimiento donde ya existe tracción y rentabilidad evita “comprar ventas” con descuentos.  
**Decisión:** Priorizar tres pilotos de 60 días con metas de +10 % ventas y guardarraíles de margen y devolución.
"""
    )

    st.markdown("---")
    st.subheader("Recomendaciones concretas (mín. 3 · sustentadas en datos)")
    decisiones = pd.DataFrame(
        [
            {
                "Prioridad": 1,
                "Acción": "Validar cobertura y recalibrar metas mes–región–canal",
                "Segmento": "Toda la red",
                "Impacto esperado": "Metas defendibles ante el Directorio",
            },
            {
                "Prioridad": 2,
                "Acción": "Piloto Sucre · Tienda · Electrónica",
                "Segmento": "Sucre / Tienda / Electrónica",
                "Impacto esperado": "+10 % ventas; margen ≥ 34 %; devolución ≤ 8 %",
            },
            {
                "Prioridad": 3,
                "Acción": "Piloto Potosí · Tienda · Hogar",
                "Segmento": "Potosí / Tienda / Hogar",
                "Impacto esperado": "Aprovechar el mayor margen regional",
            },
            {
                "Prioridad": 4,
                "Acción": "Escalar controladamente Santa Cruz · Tienda · Electrónica",
                "Segmento": "Santa Cruz / Tienda / Electrónica",
                "Impacto esperado": "Capturar volumen sin erosionar margen",
            },
            {
                "Prioridad": 5,
                "Acción": "Prueba pequeña WhatsApp · Electrónica (Cochabamba)",
                "Segmento": "WhatsApp / Electrónica",
                "Impacto esperado": "Validar canal de mayor margen antes de escalar",
            },
        ]
    )
    st.dataframe(decisiones, use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────
# PÁGINA 5 · PLAN DE ACCIÓN
# ──────────────────────────────────────────────
elif pagina == "5 · Plan de acción":
    st.title("Plan de acción recomendado")
    st.caption("Cada acción indica responsable, plazo y KPI de verificación (requisito del proyecto final).")

    st.dataframe(plan_accion, use_container_width=True, hide_index=True, height=380)

    st.success(
        "**Principio de gobernanza:** No se escala ningún piloto que incumpla los guardarraíles de margen o devolución. "
        "El crecimiento debe ser defendible ante el Directorio."
    )

# ──────────────────────────────────────────────
# PÁGINA 6 · CALIDAD DE DATOS
# ──────────────────────────────────────────────
elif pagina == "6 · Calidad de datos":
    st.title("Registro de calidad de datos")
    st.caption(f"Base original: `{RUTA_CASO.name}` · Entrega procesada: `{RUTA_ENTREGA.name}`")

    c1, c2, c3 = st.columns(3)
    c1.metric("Filas RAW", f"{len(ventas) + len(ventas_eliminadas):,}")
    c2.metric("Ventas limpias", f"{len(ventas):,}")
    c3.metric("Ventas eliminadas", f"{len(ventas_eliminadas):,}")

    st.markdown("### Auditoría forense de metas (14 combinaciones analizadas)")
    st.markdown(
        "Se revisaron las 13 combinaciones con metas múltiples y la combinación sin meta "
        "(Sep-2025 · Santa Cruz · Marketplace). Cada caso cuenta con identificación de `Venta_ID`, "
        "causa raíz (contaminación por tienda o canal) y justificación metodológica:"
    )
    st.dataframe(auditoria_metas, use_container_width=True, hide_index=True, height=360)

    st.markdown("### Resumen de eliminaciones")
    st.dataframe(_tabla_excel_a_texto(resumen_eliminados), use_container_width=True, hide_index=True, height=260)

    st.markdown("### Detalle de filas eliminadas")
    st.dataframe(ventas_eliminadas, use_container_width=True, hide_index=True, height=320)

    st.markdown("### Registro de calidad y mapeos")
    st.dataframe(_tabla_excel_a_texto(registro_calidad), use_container_width=True, hide_index=True, height=420)

    st.info(
        "Regla aplicada: duplicados, fechas inválidas, pedidos cancelados y campos críticos "
        "en blanco **no permanecen** en Ventas_Limpias. Cada fila eliminada incluye su "
        "`Motivo_Eliminacion` en la hoja `03_Ventas_Eliminadas`."
    )

    with st.expander("Ver KPIs definidos (fórmulas de gestión)"):
        st.dataframe(kpis_tabla, use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────
# FOOTER
# ──────────────────────────────────────────────
st.sidebar.markdown("---")
st.sidebar.caption(
    "Proyecto Final · Módulo 4  \n"
    "Marketing Intelligence, Dashboards & Executive Storytelling  \n"
    "Caso 1 · Retail Omnicanal · Reto A  \n"
    f"Datos: {RUTA_ENTREGA.name}"
)