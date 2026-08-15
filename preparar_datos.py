#!/usr/bin/env python3
"""
Limpieza reproducible · Caso 1 · Retail Omnicanal · Reto A
Fuente: Caso_1_Retail_Omnicanal.xlsx (hoja C1_Retail_RAW)

Regla principal:
- Ventas_Limpias solo conserva filas válidas y completas para análisis.
- Ventas_Eliminadas concentra duplicados, errores y campos críticos en blanco,
  cada uno con su Motivo_Eliminacion.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

RUTA_CASO = Path("Caso_1_Retail_Omnicanal.xlsx")
HOJA_RAW = "C1_Retail_RAW"
TC_USD_BOB = 6.96
TOLERANCIA_VENTA = 0.02

CAMPOS_CRITICOS = [
    "Venta_ID",
    "Fecha",
    "Region",
    "Canal",
    "Producto_ID",
    "Categoria",
    "Cantidad",
    "Precio_Unit",
    "Costo_Unit",
    "Venta_Reportada",
    "Estado_Pedido",
]

MESES_ES = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}

MAPEO_REGION = {
    "la paz": "La Paz", "lapaz": "La Paz", "santa cruz": "Santa Cruz",
    "sta cruz": "Santa Cruz", "scz": "Santa Cruz", "cochabamba": "Cochabamba",
    "cbb": "Cochabamba", "potosi": "Potosí", "potosí": "Potosí",
    "sucre": "Sucre", "chuquisaca": "Sucre",
}

MAPEO_CANAL = {
    "tienda": "Tienda", "tienda física": "Tienda", "e-commerce": "E-commerce",
    "ecommerce": "E-commerce", "online": "E-commerce", "marketplace": "Marketplace",
    "market place": "Marketplace", "whatsapp": "WhatsApp", "whats app": "WhatsApp",
}

REGIONES_CANONICAS = ["La Paz", "Santa Cruz", "Cochabamba", "Potosí", "Sucre"]
CANALES_CANONICOS = ["Tienda", "E-commerce", "Marketplace", "WhatsApp"]

ETIQUETAS_CAMPO = {
    "Venta_ID": "Venta_ID en blanco",
    "Fecha": "Fecha inválida o en blanco",
    "Region": "Región en blanco",
    "Canal": "Canal en blanco",
    "Producto_ID": "Producto_ID en blanco",
    "Categoria": "Categoría en blanco",
    "Cantidad": "Cantidad en blanco o inválida",
    "Precio_Unit": "Precio unitario en blanco o inválido",
    "Costo_Unit": "Costo unitario en blanco o inválido",
    "Venta_Reportada": "Venta reportada en blanco o inválida",
    "Estado_Pedido": "Estado de pedido en blanco",
}


def _normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def _es_vacio(valor) -> bool:
    return _normalizar_texto(valor) == ""


def limpiar_region(valor) -> str | None:
    texto = _normalizar_texto(valor)
    if not texto:
        return None
    clave = texto.lower()
    if clave in MAPEO_REGION:
        return MAPEO_REGION[clave]
    if texto in REGIONES_CANONICAS:
        return texto
    return texto.title()


def limpiar_canal(valor) -> str | None:
    texto = _normalizar_texto(valor)
    if not texto:
        return None
    clave = texto.lower()
    if clave in MAPEO_CANAL:
        return MAPEO_CANAL[clave]
    if texto in CANALES_CANONICOS:
        return texto
    return texto.title()


def parsear_fecha(valor) -> pd.Timestamp:
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, str):
        if valor.strip().lower() in {"", "s/f", "sf", "na", "n/a"}:
            return pd.NaT
        return pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if isinstance(valor, (int, float, np.integer, np.floating)):
        return pd.to_datetime(valor, origin="1899-12-30", unit="D", errors="coerce")
    return pd.to_datetime(valor, errors="coerce")


def limpiar_descuento(valor) -> float:
    desc = pd.to_numeric(valor, errors="coerce")
    if pd.isna(desc):
        return 0.0
    if desc > 1:
        desc = desc / 100
    return float(np.clip(desc, 0, 1))


def convertir_a_bob(valor, moneda: str) -> float:
    monto = pd.to_numeric(valor, errors="coerce")
    if pd.isna(monto):
        return np.nan
    if _normalizar_texto(moneda).upper() == "USD":
        return float(monto * TC_USD_BOB)
    return float(monto)


def clasificar_estado(cantidad: float, estado_original: str) -> str:
    estado = _normalizar_texto(estado_original)
    if estado == "Cancelado":
        return "Cancelado"
    if cantidad < 0 or estado == "Devuelto":
        return "Devuelto"
    return "Completado"


def etiquetar_opcional(valor, etiqueta: str = "No identificado") -> str:
    texto = _normalizar_texto(valor)
    return texto if texto else etiqueta


def cargar_raw(ruta: Path | str = RUTA_CASO) -> pd.DataFrame:
    return pd.read_excel(ruta, sheet_name=HOJA_RAW)


def _motivos_eliminacion(row: pd.Series, es_duplicado: bool) -> list[str]:
    motivos: list[str] = []

    if es_duplicado:
        motivos.append("Duplicado Venta_ID (se conserva la primera aparición)")

    if _es_vacio(row.get("Venta_ID")):
        motivos.append(ETIQUETAS_CAMPO["Venta_ID"])

    if pd.isna(row.get("Fecha_Limpia")):
        motivos.append(ETIQUETAS_CAMPO["Fecha"])

    if row.get("Region_Limpia") is None or pd.isna(row.get("Region_Limpia")):
        motivos.append(ETIQUETAS_CAMPO["Region"])

    if row.get("Canal_Limpio") is None or pd.isna(row.get("Canal_Limpio")):
        motivos.append(ETIQUETAS_CAMPO["Canal"])

    if _es_vacio(row.get("Producto_ID")):
        motivos.append(ETIQUETAS_CAMPO["Producto_ID"])

    if _es_vacio(row.get("Categoria")):
        motivos.append(ETIQUETAS_CAMPO["Categoria"])

    if pd.isna(row.get("Cantidad")):
        motivos.append(ETIQUETAS_CAMPO["Cantidad"])

    if pd.isna(row.get("Precio_Unit_BOB")):
        motivos.append(ETIQUETAS_CAMPO["Precio_Unit"])

    if pd.isna(row.get("Costo_Unit_BOB")):
        motivos.append(ETIQUETAS_CAMPO["Costo_Unit"])

    if pd.isna(row.get("Venta_Reportada_BOB")):
        motivos.append(ETIQUETAS_CAMPO["Venta_Reportada"])

    if _es_vacio(row.get("Estado_Original")):
        motivos.append(ETIQUETAS_CAMPO["Estado_Pedido"])

    if _normalizar_texto(row.get("Estado_Original")) == "Cancelado":
        motivos.append("Pedido cancelado (excluido del análisis de ventas netas)")

    if row.get("Region_Limpia") not in REGIONES_CANONICAS:
        motivos.append("Región no reconocida tras estandarización")

    if row.get("Canal_Limpio") not in CANALES_CANONICOS:
        motivos.append("Canal no reconocido tras estandarización")

    return motivos


def _transformaciones_aplicadas(row: pd.Series) -> str:
    partes = []
    if _normalizar_texto(row["Region_Original"]) != row["Region_Limpia"]:
        partes.append("Región estandarizada")
    if _normalizar_texto(row["Canal_Original"]).lower() != _normalizar_texto(row["Canal_Limpio"]).lower():
        partes.append("Canal estandarizado")
    if pd.to_numeric(row["Descuento_Original"], errors="coerce") > 1:
        partes.append("Descuento reescalado")
    if _normalizar_texto(row["Moneda_Original"]).upper() == "USD":
        partes.append("Venta reportada USD→BOB")
    if abs(row["Venta_Calculada_BOB"] - row["Venta_Reportada_BOB"]) > TOLERANCIA_VENTA:
        partes.append("Venta recalculada")
    if row["Estado_Original"] == "Completado" and row["Estado_Limpio"] == "Devuelto":
        partes.append("Estado reclasificado a Devuelto")
    return "; ".join(partes) if partes else "Sin transformación"


def preparar_ventas(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    df = raw.copy()
    df["Fila_Origen"] = df.index + 2
    stats: dict = {"filas_raw": len(df)}

    df["Es_Duplicado"] = df["Venta_ID"].duplicated(keep="first")
    stats["duplicados_venta_id"] = int(df["Es_Duplicado"].sum())

    df["Fecha_Original"] = df["Fecha"]
    df["Fecha_Limpia"] = df["Fecha"].apply(parsear_fecha)

    df["Region_Original"] = df["Region"]
    df["Region_Limpia"] = df["Region"].apply(limpiar_region)

    df["Canal_Original"] = df["Canal"]
    df["Canal_Limpio"] = df["Canal"].apply(limpiar_canal)

    df["Descuento_Original"] = df["Descuento_pct"]
    df["Descuento_Limpio"] = df["Descuento_pct"].apply(limpiar_descuento)

    df["Moneda_Original"] = df["Moneda"]
    df["Precio_Unit_BOB"] = pd.to_numeric(df["Precio_Unit"], errors="coerce")
    df["Costo_Unit_BOB"] = pd.to_numeric(df["Costo_Unit"], errors="coerce")
    df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce")
    df["Venta_Reportada_Original"] = df["Venta_Reportada"]
    df["Venta_Reportada_BOB"] = [
        convertir_a_bob(v, m) for v, m in zip(df["Venta_Reportada"], df["Moneda"])
    ]
    df["Venta_Calculada_BOB"] = df["Cantidad"] * df["Precio_Unit_BOB"] * (1 - df["Descuento_Limpio"])
    df["Costo_Total_BOB"] = df["Cantidad"] * df["Costo_Unit_BOB"]
    df["Utilidad_BOB"] = df["Venta_Calculada_BOB"] - df["Costo_Total_BOB"]
    df["Margen_Fila"] = np.where(
        df["Venta_Calculada_BOB"] != 0,
        df["Utilidad_BOB"] / df["Venta_Calculada_BOB"],
        np.nan,
    )

    df["Estado_Original"] = df["Estado_Pedido"]
    df["Estado_Limpio"] = [
        clasificar_estado(c, e) for c, e in zip(df["Cantidad"], df["Estado_Pedido"])
    ]

    df["Meta_Original_No_Sumar"] = df["Meta_Ventas_Mes_Region_Canal"]
    df["Stock_Original_No_Sumar"] = df["Stock_Cierre_Mes"]
    df["LeadTime_Original_No_Promediar"] = df["LeadTime_dias"]

    motivos_por_fila = [
        _motivos_eliminacion(row, bool(row["Es_Duplicado"])) for _, row in df.iterrows()
    ]
    df["Motivo_Eliminacion"] = ["; ".join(m) if m else "" for m in motivos_por_fila]
    df["Transformaciones_Aplicadas"] = df.apply(_transformaciones_aplicadas, axis=1)

    eliminadas = df[df["Motivo_Eliminacion"] != ""].copy()
    limpias = df[df["Motivo_Eliminacion"] == ""].copy()

    limpias["Año"] = limpias["Fecha_Limpia"].dt.year
    limpias["Mes_Num"] = limpias["Fecha_Limpia"].dt.month
    limpias["Mes"] = limpias["Mes_Num"].map(MESES_ES)
    limpias["Año_Mes"] = limpias["Fecha_Limpia"].dt.to_period("M").astype(str)
    limpias["Trimestre"] = limpias["Fecha_Limpia"].dt.to_period("Q").astype(str)
    limpias["Marca"] = limpias["Marca"].apply(lambda x: etiquetar_opcional(x, "No identificada"))
    limpias["Segmento_Cliente"] = limpias["Segmento_Cliente"].apply(
        lambda x: etiquetar_opcional(x, "No identificado")
    )

    columnas_limpias = [
        "Fila_Origen", "Venta_ID", "Fecha_Original", "Fecha_Limpia", "Año", "Mes_Num", "Mes",
        "Año_Mes", "Trimestre", "Tienda_ID", "Tienda", "Formato_Tienda", "Region_Original",
        "Region_Limpia", "Cliente_ID", "Segmento_Cliente", "Producto_ID", "Categoria", "Marca",
        "Canal_Original", "Canal_Limpio", "Vendedor_ID", "Cantidad", "Precio_Unit_BOB",
        "Descuento_Original", "Descuento_Limpio", "Costo_Unit_BOB", "Venta_Reportada_Original",
        "Moneda_Original", "Venta_Reportada_BOB", "Venta_Calculada_BOB", "Costo_Total_BOB",
        "Utilidad_BOB", "Margen_Fila", "Meta_Original_No_Sumar", "Stock_Original_No_Sumar",
        "LeadTime_Original_No_Promediar", "Estado_Original", "Estado_Limpio", "Medio_Pago",
        "Transformaciones_Aplicadas",
    ]

    columnas_eliminadas = [
        "Fila_Origen", "Venta_ID", "Fecha_Original", "Fecha_Limpia", "Region_Original",
        "Region_Limpia", "Canal_Original", "Canal_Limpio", "Producto_ID", "Categoria",
        "Cantidad", "Precio_Unit", "Costo_Unit", "Venta_Reportada", "Moneda_Original",
        "Estado_Original", "Estado_Limpio", "Motivo_Eliminacion", "Transformaciones_Aplicadas",
    ]

    stats.update({
        "filas_limpias": len(limpias),
        "filas_eliminadas": len(eliminadas),
        "fechas_invalidas": int(eliminadas["Motivo_Eliminacion"].str.contains("Fecha", na=False).sum()),
        "duplicados_en_eliminadas": int(eliminadas["Motivo_Eliminacion"].str.contains("Duplicado", na=False).sum()),
        "cancelados_en_eliminadas": int(eliminadas["Motivo_Eliminacion"].str.contains("cancelado", na=False).sum()),
        "campos_blanco_en_eliminadas": int(
            eliminadas["Motivo_Eliminacion"].str.contains("blanco", na=False).sum()
        ),
        "regiones_no_estandarizadas": int(
            (limpias["Region_Original"].apply(_normalizar_texto) != limpias["Region_Limpia"]).sum()
        ),
        "canales_no_estandarizados": int(
            (limpias["Canal_Original"].apply(_normalizar_texto).str.lower()
             != limpias["Canal_Limpio"].apply(_normalizar_texto).str.lower()).sum()
        ),
        "descuentos_escala_incorrecta": int(
            (pd.to_numeric(limpias["Descuento_Original"], errors="coerce") > 1).sum()
        ),
        "ventas_usd": int((limpias["Moneda_Original"].astype(str).str.upper() == "USD").sum()),
        "ventas_inconsistentes": int(
            (abs(limpias["Venta_Calculada_BOB"] - limpias["Venta_Reportada_BOB"]) > TOLERANCIA_VENTA).sum()
        ),
        "cantidad_negativa_reclasificada": int(
            ((limpias["Cantidad"] < 0) & (limpias["Estado_Original"] == "Completado")).sum()
        ),
    })

    return limpias[columnas_limpias], eliminadas[columnas_eliminadas], stats


DICT_METAS_CORREGIDAS = {
    ("2024-03", "Potosí", "Marketplace"): (
        72039.98,
        "Valor 87,282.00 provino de tienda T009 (Cochabamba) en VT0005331; 72,039.98 es la meta correcta respaldada por tiendas nativas T015 y T016.",
    ),
    ("2024-04", "La Paz", "Tienda"): (
        193849.26,
        "Valor 137,908.32 provino de tienda T012 (Cochabamba) en VT0006333; 193,849.26 es la meta oficial de La Paz Tienda confirmada por 10 tiendas de La Paz.",
    ),
    ("2024-06", "Potosí", "Tienda"): (
        201185.26,
        "Valor 177,388.08 provino de tienda T004 (La Paz) en VT0000685; 201,185.26 es la meta oficial de Potosí Tienda confirmada por 10 registros locales.",
    ),
    ("2024-09", "Cochabamba", "WhatsApp"): (
        68949.92,
        "Valor 148,910.81 provino de canal Tienda en VT0006804; 68,949.92 es la meta correcta del canal WhatsApp en Cochabamba.",
    ),
    ("2024-11", "Potosí", "E-commerce"): (
        77981.44,
        "Valor 91,910.55 provino de tienda T011 (Cochabamba) en VT0004694; 77,981.44 es la meta oficial de Potosí E-commerce.",
    ),
    ("2024-11", "Sucre", "E-commerce"): (
        70991.46,
        "Valor 163,312.54 provino de canal Tienda en VT0004799; 70,991.46 es la meta correcta del canal E-commerce en Sucre.",
    ),
    ("2025-01", "Sucre", "Tienda"): (
        139765.58,
        "Valor 178,667.19 provino de tienda T013 (Potosí) en VT0003535; 139,765.58 es la meta oficial de Sucre Tienda.",
    ),
    ("2025-02", "Cochabamba", "Tienda"): (
        169210.56,
        "Valor 174,033.03 provino de tienda T007 (Santa Cruz) en VT0002904; 169,210.56 es la meta oficial de Cochabamba Tienda.",
    ),
    ("2025-02", "Santa Cruz", "E-commerce"): (
        84128.79,
        "Valor 86,253.49 provino de tienda T012 (Cochabamba) en VT0002058; 84,128.79 es la meta oficial de Santa Cruz E-commerce.",
    ),
    ("2025-03", "La Paz", "E-commerce"): (
        76648.32,
        "Valor 84,388.51 provino de tienda T014 (Potosí) en VT0001369; 76,648.32 es la meta oficial de La Paz E-commerce.",
    ),
    ("2025-10", "Potosí", "E-commerce"): (
        98671.43,
        "Valor 166,027.73 provino de canal Tienda en VT0000272; 98,671.43 es la meta correcta del canal E-commerce en Potosí.",
    ),
    ("2025-12", "Santa Cruz", "WhatsApp"): (
        60781.90,
        "Valor 186,128.69 provino de canal Tienda en pedido cancelado VT0001792; 60,781.90 es la meta correcta de Santa Cruz WhatsApp.",
    ),
    ("2026-02", "Sucre", "Tienda"): (
        120492.00,
        "Valor 92,109.62 provino de canal E-commerce en VT0008808; 120,492.00 es la meta oficial de Sucre Tienda.",
    ),
}

AUDITORIA_14_METAS_DETALLE = [
    {
        "Caso": 1,
        "Año_Mes": "2024-03",
        "Region": "Potosí",
        "Canal": "Marketplace",
        "Valores_Encontrados": "72,039.98 / 87,282.00",
        "Detalle_Ventas_Por_Valor": "72,039.98 (2 reg: VT0000676, VT0006207) ; 87,282.00 (1 reg: VT0005331)",
        "Valor_Seleccionado": 72039.98,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0005331 tiene tienda T009 (Cochabamba). 72,039.98 BOB corresponde a las tiendas nativas T015/T016 de Potosí y a la serie histórica.",
    },
    {
        "Caso": 2,
        "Año_Mes": "2024-04",
        "Region": "La Paz",
        "Canal": "Tienda",
        "Valores_Encontrados": "137,908.32 / 193,849.26",
        "Detalle_Ventas_Por_Valor": "137,908.32 (1 reg: VT0006333) ; 193,849.26 (10 reg: VT0006352, VT0006637, VT0007516, VT0007906, VT0009008, VT0009150, VT0009806, VT0010443, VT0011150, VT0011464)",
        "Valor_Seleccionado": 193849.26,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0006333 tiene tienda T012 (Cochabamba) e importó su meta. 193,849.26 BOB es la meta oficial de La Paz Tienda validada por 10 tiendas locales.",
    },
    {
        "Caso": 3,
        "Año_Mes": "2024-06",
        "Region": "Potosí",
        "Canal": "Tienda",
        "Valores_Encontrados": "177,388.08 / 201,185.26",
        "Detalle_Ventas_Por_Valor": "177,388.08 (1 reg: VT0000685) ; 201,185.26 (10 reg: VT0000205, VT0000528, VT0001488, VT0002609, VT0002842, VT0003008, VT0004253, VT0005393, VT0006267, VT0007193)",
        "Valor_Seleccionado": 201185.26,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0000685 tiene tienda T004 (La Paz) y arrastró la meta de La Paz. 201,185.26 BOB es la meta oficial de Potosí Tienda.",
    },
    {
        "Caso": 4,
        "Año_Mes": "2024-09",
        "Region": "Cochabamba",
        "Canal": "WhatsApp",
        "Valores_Encontrados": "68,949.92 / 148,910.81",
        "Detalle_Ventas_Por_Valor": "68,949.92 (1 reg: VT0000880) ; 148,910.81 (1 reg: VT0006804)",
        "Valor_Seleccionado": 68949.92,
        "Tipo_Error": "Contaminación por canal",
        "Justificacion": "148,910.81 BOB es la meta exacta de Cochabamba Tienda física. 68,949.92 BOB corresponde al presupuesto del canal WhatsApp.",
    },
    {
        "Caso": 5,
        "Año_Mes": "2024-11",
        "Region": "Potosí",
        "Canal": "E-commerce",
        "Valores_Encontrados": "77,981.44 / 91,910.55",
        "Detalle_Ventas_Por_Valor": "77,981.44 (3 reg: VT0005683, VT0007611, VT0007873) ; 91,910.55 (1 reg: VT0004694)",
        "Valor_Seleccionado": 77981.44,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0004694 tiene tienda T011 (Cochabamba) e importó 91,910.55 (Cochabamba E-commerce). 77,981.44 BOB es la meta de Potosí E-commerce.",
    },
    {
        "Caso": 6,
        "Año_Mes": "2024-11",
        "Region": "Sucre",
        "Canal": "E-commerce",
        "Valores_Encontrados": "70,991.46 / 163,312.54",
        "Detalle_Ventas_Por_Valor": "70,991.46 (4 reg: VT0001645, VT0005488, VT0010105, VT0010215) ; 163,312.54 (1 reg: VT0004799)",
        "Valor_Seleccionado": 70991.46,
        "Tipo_Error": "Contaminación por canal",
        "Justificacion": "163,312.54 BOB es la meta de Sucre Tienda física copiada en VT0004799. 70,991.46 BOB es la meta presupuestada de Sucre E-commerce.",
    },
    {
        "Caso": 7,
        "Año_Mes": "2025-01",
        "Region": "Sucre",
        "Canal": "Tienda",
        "Valores_Encontrados": "139,765.58 / 178,667.19",
        "Detalle_Ventas_Por_Valor": "139,765.58 (9 reg: VT0002467, VT0003640, VT0003773, VT0004015, VT0005027, VT0005312, VT0006889, VT0007293, VT0009469) ; 178,667.19 (1 reg: VT0003535) ; NaN (1 reg: VT0003098)",
        "Valor_Seleccionado": 139765.58,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0003535 tiene tienda T013 (Potosí) e importó 178,667.19. 139,765.58 BOB es la meta oficial de Sucre Tienda confirmada por 9 tiendas locales.",
    },
    {
        "Caso": 8,
        "Año_Mes": "2025-02",
        "Region": "Cochabamba",
        "Canal": "Tienda",
        "Valores_Encontrados": "169,210.56 / 174,033.03",
        "Detalle_Ventas_Por_Valor": "169,210.56 (11 reg: VT0001141, VT0001227, VT0001383, VT0002562, VT0003402, VT0005493, VT0005754, VT0007716, VT0008310, VT0008889, VT0011155) ; 174,033.03 (1 reg: VT0002904) ; NaN (1 reg: VT0003711)",
        "Valor_Seleccionado": 169210.56,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0002904 tiene tienda T007 (Santa Cruz) e importó 174,033.03. 169,210.56 BOB es la meta oficial de Cochabamba Tienda.",
    },
    {
        "Caso": 9,
        "Año_Mes": "2025-02",
        "Region": "Santa Cruz",
        "Canal": "E-commerce",
        "Valores_Encontrados": "84,128.79 / 86,253.49",
        "Detalle_Ventas_Por_Valor": "84,128.79 (4 reg: VT0000410, VT0000733, VT0005008, VT0010091) ; 86,253.49 (1 reg: VT0002058)",
        "Valor_Seleccionado": 84128.79,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0002058 tiene tienda T012 (Cochabamba) e importó 86,253.49. 84,128.79 BOB es la meta oficial de Santa Cruz E-commerce.",
    },
    {
        "Caso": 10,
        "Año_Mes": "2025-03",
        "Region": "La Paz",
        "Canal": "E-commerce",
        "Valores_Encontrados": "76,648.32 / 84,388.51",
        "Detalle_Ventas_Por_Valor": "76,648.32 (2 reg: VT0008894, VT0011939) ; 84,388.51 (1 reg: VT0001369)",
        "Valor_Seleccionado": 76648.32,
        "Tipo_Error": "Contaminación por tienda",
        "Justificacion": "VT0001369 tiene tienda T014 (Potosí) e importó 84,388.51. 76,648.32 BOB es la meta oficial de La Paz E-commerce.",
    },
    {
        "Caso": 11,
        "Año_Mes": "2025-10",
        "Region": "Potosí",
        "Canal": "E-commerce",
        "Valores_Encontrados": "98,671.43 / 166,027.73",
        "Detalle_Ventas_Por_Valor": "98,671.43 (6 reg: VT0001046, VT0001640, VT0003212, VT0003730, VT0004300, VT0004818) ; 166,027.73 (1 reg: VT0000272)",
        "Valor_Seleccionado": 98671.43,
        "Tipo_Error": "Contaminación por canal",
        "Justificacion": "166,027.73 BOB es la meta física de Potosí Tienda. 98,671.43 BOB es la meta correcta del canal digital E-commerce en Potosí.",
    },
    {
        "Caso": 12,
        "Año_Mes": "2025-12",
        "Region": "Santa Cruz",
        "Canal": "WhatsApp",
        "Valores_Encontrados": "60,781.90 / 186,128.69",
        "Detalle_Ventas_Por_Valor": "60,781.90 (1 reg: VT0001336) ; 186,128.69 (1 reg cancelado: VT0001792)",
        "Valor_Seleccionado": 60781.90,
        "Tipo_Error": "Contaminación por canal en cancelado",
        "Justificacion": "186,128.69 BOB es la meta de Santa Cruz Tienda física navideña. 60,781.90 BOB es la meta correspondiente al canal WhatsApp.",
    },
    {
        "Caso": 13,
        "Año_Mes": "2026-02",
        "Region": "Sucre",
        "Canal": "Tienda",
        "Valores_Encontrados": "92,109.62 / 120,492.00",
        "Detalle_Ventas_Por_Valor": "120,492.00 (7 reg: VT0001616, VT0001811, VT0008034, VT0008419, VT0008960, VT0009991, VT0010390) ; 92,109.62 (1 reg: VT0008808)",
        "Valor_Seleccionado": 120492.00,
        "Tipo_Error": "Contaminación por canal",
        "Justificacion": "92,109.62 BOB es la meta de Sucre E-commerce. 120,492.00 BOB es la meta oficial de Sucre Tienda validada por 7 tiendas locales.",
    },
    {
        "Caso": 14,
        "Año_Mes": "2025-09",
        "Region": "Santa Cruz",
        "Canal": "Marketplace",
        "Valores_Encontrados": "Sin meta en RAW",
        "Detalle_Ventas_Por_Valor": "Sin meta / NaN (1 reg: VT0010965)",
        "Valor_Seleccionado": np.nan,
        "Tipo_Error": "Meta ausente en origen",
        "Justificacion": "Único registro del grupo sin presupuesto en RAW original. Se mantiene como NaN/Sin meta sin imputar cero ni valores artificiales.",
    },
]


def construir_auditoria_14_metas() -> pd.DataFrame:
    return pd.DataFrame(AUDITORIA_14_METAS_DETALLE)


def preparar_metas(raw: pd.DataFrame, ventas: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    raw_df = raw.copy()
    raw_df["Fecha_Limpia"] = raw_df["Fecha"].apply(parsear_fecha)
    raw_df["Año_Mes"] = raw_df["Fecha_Limpia"].dt.to_period("M").astype(str)
    raw_df["Region_Limpia"] = raw_df["Region"].apply(limpiar_region)
    raw_df["Canal_Limpio"] = raw_df["Canal"].apply(limpiar_canal)

    # Filtrar solo registros con fecha válida (excluir NaT)
    base_raw = raw_df[raw_df["Año_Mes"] != "NaT"].copy()

    # Mapeo de ventas reales calculadas desde Ventas_Limpias
    v_map = ventas.groupby(["Año_Mes", "Region_Limpia", "Canal_Limpio"])["Venta_Calculada_BOB"].sum().to_dict()

    registros = []
    for keys, grupo in base_raw.groupby(["Año_Mes", "Region_Limpia", "Canal_Limpio"], dropna=False):
        m, r, c = keys
        metas_serie = grupo["Meta_Ventas_Mes_Region_Canal"].dropna()
        valores_u = sorted(list(set(metas_serie.tolist())))
        filas_origen = len(grupo)
        filas_meta_vacia = int(grupo["Meta_Ventas_Mes_Region_Canal"].isna().sum())

        if (m, r, c) in DICT_METAS_CORREGIDAS:
            meta_corregida, motivo = DICT_METAS_CORREGIDAS[(m, r, c)]
            calidad = "Corregida por validación cruzada"
            meta_orig_str = " / ".join([f"{v:,.2f}" for v in valores_u])
        elif len(valores_u) == 1:
            meta_corregida = float(valores_u[0])
            motivo = "Meta consistente en todas las transacciones del grupo"
            calidad = "Consistente"
            meta_orig_str = f"{meta_corregida:,.2f}"
        elif len(valores_u) == 0:
            meta_corregida = np.nan
            motivo = "Sin meta registrada en RAW original; registro único VT0010965"
            calidad = "Sin meta"
            meta_orig_str = "Sin meta en RAW"
        else:
            meta_corregida = float(valores_u[0])
            motivo = "Revisión manual"
            calidad = "Ajustada"
            meta_orig_str = str(valores_u)

        v_real = v_map.get((m, r, c), 0.0)
        cump = (
            v_real / meta_corregida
            if pd.notna(meta_corregida) and meta_corregida > 0
            else np.nan
        )

        registros.append({
            "Año_Mes": m,
            "Region_Limpia": r,
            "Canal_Limpio": c,
            "Meta_Original": meta_orig_str,
            "Meta_Corregida_BOB": meta_corregida,
            "Valores_Encontrados": " / ".join([f"{v:,.2f}" for v in valores_u]) if valores_u else "Ninguno",
            "Filas_Origen": filas_origen,
            "Valores_Distintos": len(valores_u),
            "Filas_Meta_Vacia": filas_meta_vacia,
            "Calidad_Meta": calidad,
            "Motivo_Correccion": motivo,
            "Ventas_Reales_BOB": v_real,
            "Cumplimiento_Mes_Region_Canal": cump,
        })

    metas = pd.DataFrame(registros).sort_values(["Año_Mes", "Region_Limpia", "Canal_Limpio"])
    meta_total = float(metas["Meta_Corregida_BOB"].sum(skipna=True))
    ventas_con_meta = float(metas[metas["Meta_Corregida_BOB"].notna()]["Ventas_Reales_BOB"].sum())

    stats = {
        "grupos_meta": len(metas),
        "grupos_conflicto": int((metas["Valores_Distintos"] > 1).sum()),
        "grupos_sin_meta": int(metas["Meta_Corregida_BOB"].isna().sum()),
        "meta_total": meta_total,
        "ventas_con_meta": ventas_con_meta,
        "cumplimiento_global_pond": ventas_con_meta / meta_total if meta_total else 0,
    }
    return metas, stats


def preparar_inventario(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df["Fecha_Limpia"] = df["Fecha"].apply(parsear_fecha)
    df["Año_Mes"] = df["Fecha_Limpia"].dt.to_period("M").astype(str)
    df["Region_Limpia"] = df["Region"].apply(limpiar_region)
    inv = df[df["Stock_Cierre_Mes"].notna() | df["LeadTime_dias"].notna()].copy()
    inv = inv.drop_duplicates(
        subset=["Año_Mes", "Tienda_ID", "Producto_ID", "Stock_Cierre_Mes", "LeadTime_dias"],
        keep="first",
    )
    inv["Fila_Origen"] = inv.index + 2
    return inv[
        ["Año_Mes", "Tienda_ID", "Tienda", "Producto_ID", "Categoria", "Region_Limpia",
         "Stock_Cierre_Mes", "LeadTime_dias", "Fila_Origen"]
    ].rename(columns={"Region_Limpia": "Region"})


def calcular_kpis(ventas: pd.DataFrame, metas: pd.DataFrame) -> pd.DataFrame:
    ventas_netas = ventas["Venta_Calculada_BOB"].sum()
    utilidad = ventas["Utilidad_BOB"].sum()
    margen = utilidad / ventas_netas if ventas_netas else 0
    completados = ventas[ventas["Estado_Limpio"] == "Completado"]
    pedidos_comp = completados["Venta_ID"].nunique()
    ticket = completados["Venta_Calculada_BOB"].sum() / pedidos_comp if pedidos_comp else 0
    devueltos = ventas[ventas["Estado_Limpio"] == "Devuelto"]["Venta_ID"].nunique()
    tasa_dev = devueltos / (devueltos + pedidos_comp) if (devueltos + pedidos_comp) else 0
    venta_sin_desc = (ventas["Cantidad"] * ventas["Precio_Unit_BOB"]).sum()
    desc_pond = 1 - (ventas_netas / venta_sin_desc) if venta_sin_desc else 0
    
    meta_total = float(metas["Meta_Corregida_BOB"].sum(skipna=True))
    ventas_con_meta = float(metas[metas["Meta_Corregida_BOB"].notna()]["Ventas_Reales_BOB"].sum())
    cumplimiento = ventas_con_meta / meta_total if meta_total else 0

    datos = [
        ("Ventas netas", "SUM(Venta_Calculada_BOB)", "Medir ingresos después de devoluciones",
         "Valor central de desempeño", "Priorizar crecimiento rentable", ventas_netas),
        ("Utilidad total", "SUM(Utilidad_BOB)", "Medir aporte económico",
         "Ventas altas sin utilidad no son sostenibles", "Proteger precio, costo y mix", utilidad),
        ("Margen %", "Utilidad / Ventas netas", "Evaluar rentabilidad",
         "Compara calidad del crecimiento", "Evitar expandir combinaciones de bajo margen", margen),
        ("Pedidos completados", "DISTINCTCOUNT Venta_ID con Estado=Completado", "Medir volumen efectivo",
         "Separa ventas de devoluciones/cancelaciones", "Dimensionar capacidad comercial", pedidos_comp),
        ("Ticket promedio", "Ventas completadas / Pedidos completados", "Medir gasto por pedido",
         "Detecta oportunidades de venta cruzada", "Diseñar bundles y upselling", ticket),
        ("Unidades netas", "SUM(Cantidad)", "Medir volumen descontando devoluciones",
         "Muestra presión real sobre inventario", "Alinear abastecimiento", ventas["Cantidad"].sum()),
        ("Clientes únicos", "DISTINCTCOUNT Cliente_ID", "Medir alcance",
         "Distingue crecimiento por clientes vs ticket", "Activar adquisición o recompra",
         ventas["Cliente_ID"].nunique()),
        ("Tasa de devolución", "Pedidos devueltos / (completados + devueltos)", "Controlar pérdida y fricción",
         "Una tasa alta destruye venta y margen", "Revisar calidad/producto/canal", tasa_dev),
        ("Descuento ponderado", "Descuento BOB / Venta antes de descuento", "Medir intensidad promocional",
         "Relaciona crecimiento con sacrificio de precio", "Limitar descuentos improductivos", desc_pond),
        ("Cumplimiento de meta", "Ventas con meta / Meta única mes–región–canal", "Medir brecha comercial",
         "Ratio de 1,39% explicado por muestra transaccional vs presupuesto corporativo macro",
         "Priorizar regiones y canales con buen margen", cumplimiento),
    ]
    return pd.DataFrame(
        datos,
        columns=["KPI", "Fórmula conceptual", "Objetivo", "Interpretación", "Decisión asociada", "Resultado actual"],
    )


def resumen_eliminaciones(eliminadas: pd.DataFrame) -> pd.DataFrame:
    conteos: dict[str, int] = {}
    for texto in eliminadas["Motivo_Eliminacion"].fillna(""):
        for motivo in texto.split("; "):
            if motivo:
                conteos[motivo] = conteos.get(motivo, 0) + 1
    filas = [["Resumen de filas eliminadas", "", ""], ["", "", ""],
             ["Motivo de eliminación", "Cantidad", "Nota"]]
    notas = {
        "Duplicado Venta_ID (se conserva la primera aparición)": "Evita doble conteo",
        "Fecha inválida o en blanco": "No entra en análisis temporal",
        "Pedido cancelado (excluido del análisis de ventas netas)": "No representa venta efectiva",
        "Producto_ID en blanco": "No permite análisis por producto/categoría",
        "Categoría en blanco": "No permite segmentar recomendaciones",
    }
    for motivo, cantidad in sorted(conteos.items(), key=lambda x: -x[1]):
        filas.append([motivo, cantidad, notas.get(motivo, "")])
    filas.append(["", "", ""])
    filas.append(["Total filas eliminadas", len(eliminadas), ""])
    return pd.DataFrame(filas)


def construir_guia_proyecto(stats_ventas: dict) -> pd.DataFrame:
    return pd.DataFrame([
        ["GUÍA DEL PROYECTO · Retail Omnicanal · Reto A", "", ""],
        ["", "", ""],
        ["Pregunta de negocio", "¿Dónde concentrar el crecimiento por región, canal y categoría sin destruir margen ni incumplir metas?", ""],
        ["", "", ""],
        ["Flujo de limpieza", "", ""],
        ["Paso", "Hoja resultante", "Descripción"],
        ["1", "C1_Retail_RAW", "Datos originales sin modificar"],
        ["2", "Ventas_Limpias", f"Solo filas válidas y completas ({stats_ventas['filas_limpias']} filas)"],
        ["3", "Ventas_Eliminadas", f"Duplicados, errores y blancos ({stats_ventas['filas_eliminadas']} filas)"],
        ["4", "Metas_Limpias + Inventario_Mensual", "Tablas de soporte con auditoría forense de metas sin mezclar granularidades"],
        ["5", "Registro_y_Analisis", "Calidad, KPIs, plan y mapeos"],
        ["", "", ""],
        ["Preguntas que orientan el análisis", "", ""],
        ["1", "¿Qué regiones y canales concentran ventas y margen?", "Priorización de crecimiento"],
        ["2", "¿El cumplimiento de meta es real o está mal calibrado?", "Validación antes de exigir cierre"],
        ["3", "¿Qué combinaciones crecen con margen y baja devolución?", "Pilotos defendibles"],
        ["4", "¿Los descuentos erosionan margen?", "Control promocional"],
        ["5", "¿El inventario respalda la expansión?", "Soporte operativo"],
    ])


def construir_registro_calidad(stats_ventas: dict, stats_metas: dict, inventario: pd.DataFrame, eliminadas: pd.DataFrame) -> pd.DataFrame:
    resumen = resumen_eliminaciones(eliminadas)
    filas = [
        ["REGISTRO DE CALIDAD Y MAPEOS", "", "", "", "", ""],
        ["", "", "", "", "", ""],
        ["Filas RAW", "Filas limpias", "Filas eliminadas", "Grupos meta conflictivos auditados", "Snapshots inventario", "Meta Total BOB"],
        [stats_ventas["filas_raw"], stats_ventas["filas_limpias"], stats_ventas["filas_eliminadas"],
         stats_metas["grupos_conflicto"], len(inventario), f"{stats_metas['meta_total']:,.2f}"],
        ["", "", "", "", "", ""],
        ["Problema", "Cantidad", "Tratamiento", "Justificación", "Impacto si no se corrigía", "Estado"],
        ["Duplicados Venta_ID", stats_ventas["duplicados_en_eliminadas"], "Mover a Ventas_Eliminadas", "Misma operación repetida", "Doble conteo", "Eliminado"],
        ["Fechas inválidas", stats_ventas["fechas_invalidas"], "Mover a Ventas_Eliminadas", "No inventar fechas", "Tendencias distorsionadas", "Eliminado"],
        ["Pedidos cancelados", stats_ventas["cancelados_en_eliminadas"], "Mover a Ventas_Eliminadas", "No son venta efectiva", "Infla o confunde volumen", "Eliminado"],
        ["Campos críticos en blanco", stats_ventas["campos_blanco_en_eliminadas"], "Mover a Ventas_Eliminadas", "Faltan datos para analizar", "Conclusiones inválidas", "Eliminado"],
        ["Regiones no estandarizadas", stats_ventas["regiones_no_estandarizadas"], "Mapear en Ventas_Limpias", "Variantes del mismo lugar", "Fragmentación regional", "Corregido"],
        ["Canales no estandarizados", stats_ventas["canales_no_estandarizados"], "Mapear en Ventas_Limpias", "Variantes del mismo canal", "Categorías falsas", "Corregido"],
        ["Descuentos escala 0–100", stats_ventas["descuentos_escala_incorrecta"], "Dividir entre 100", "Escala decimal", "Margen distorsionado", "Corregido"],
        ["Ventas en USD", stats_ventas["ventas_usd"], "Convertir venta reportada", "Unificar moneda", "Mezcla BOB/USD", "Corregido"],
        ["Venta reportada inconsistente", stats_ventas["ventas_inconsistentes"], "Usar venta calculada", "Fórmula trazable", "Montos alterados", "Corregido"],
        ["Metas conflictivas (13 casos)", 13, "Auditoría forense de causa raíz", "Eliminar contaminación de tienda/canal", "Denominador exacto 55.97M BOB", "Auditado y Corregido"],
        ["Meta ausente (1 caso: Sep-25 SCZ Mkt)", 1, "Clasificar como Sin meta (NaN)", "Registro único sin presupuesto en RAW", "No inventar metas ni 0", "Documentado"],
        ["Stock/lead time parcial", stats_ventas["filas_raw"] - len(inventario), "Inventario_Mensual", "No es dato por venta", "Sumas absurdas", "Separado"],
        ["", "", "", "", "", ""],
        ["MAPEOS · Región original", "Región limpia", "", "Canal original", "Canal limpio", ""],
    ]
    for reg_o, reg_l in sorted(MAPEO_REGION.items()):
        filas.append([reg_o, reg_l, "", "", "", ""])
    filas.append(["", "", "", "", "", ""])
    for can_o, can_l in sorted(MAPEO_CANAL.items()):
        filas.append(["", "", "", can_o, can_l, ""])

    bloque = pd.DataFrame(filas)
    return pd.concat([bloque, pd.DataFrame([[""] * 6]), resumen], ignore_index=True)


def construir_plan_accion() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Acción": "Validar si la base representa todo el universo y recalibrar las metas mes–región–canal",
                "Responsable": "BI + Finanzas",
                "Plazo": "15 días",
                "KPI de control": "Cobertura y cumplimiento de meta",
                "Meta / guardarraíl": "100 % de grupos trazables; sin metas duplicadas",
                "Resultado esperado": "Metas defendibles para el Directorio",
            },
            {
                "Acción": "Completar la captura mensual de stock y lead time",
                "Responsable": "Operaciones + TI",
                "Plazo": "30 días",
                "KPI de control": "Cobertura de inventario",
                "Meta / guardarraíl": "≥ 90 % de combinaciones activas",
                "Resultado esperado": "Decisiones de expansión respaldadas por stock",
            },
            {
                "Acción": "Piloto: Sucre · Tienda · Electrónica",
                "Responsable": "Gerencia comercial Sucre",
                "Plazo": "60 días",
                "KPI de control": "Ventas, margen, devoluciones",
                "Meta / guardarraíl": "+10 % ventas; margen ≥ 34 %; devolución ≤ 8 %",
                "Resultado esperado": "Escalar solo si sostiene rentabilidad",
            },
            {
                "Acción": "Piloto: Potosí · Tienda · Hogar",
                "Responsable": "Gerencia comercial Potosí",
                "Plazo": "60 días",
                "KPI de control": "Ventas, margen, devoluciones",
                "Meta / guardarraíl": "+10 % ventas; margen ≥ 38 %; devolución ≤ 8 %",
                "Resultado esperado": "Aprovechar el mayor margen regional",
            },
            {
                "Acción": "Escalar controladamente: Santa Cruz · Tienda · Electrónica",
                "Responsable": "Comercial + Abastecimiento",
                "Plazo": "60 días",
                "KPI de control": "Ventas, margen, devolución",
                "Meta / guardarraíl": "+10 % ventas; margen ≥ 31 %; devolución ≤ 8 %",
                "Resultado esperado": "Capturar volumen sin erosionar margen",
            },
            {
                "Acción": "Prueba pequeña de WhatsApp en Electrónica (Cochabamba)",
                "Responsable": "Marketing + CRM",
                "Plazo": "60 días",
                "KPI de control": "Pedidos, margen, devolución",
                "Meta / guardarraíl": "≥ 10 pedidos adicionales; margen ≥ 35 %; devolución ≤ 8 %",
                "Resultado esperado": "Validar el canal de mayor margen antes de escalar",
            },
            {
                "Acción": "Revisión de pilotos y decisión de expansión",
                "Responsable": "Directorio comercial",
                "Plazo": "90 días",
                "KPI de control": "Cumplimiento de guardarraíles",
                "Meta / guardarraíl": "Escalar solo pilotos aprobados",
                "Resultado esperado": "Crecimiento disciplinado y medible",
            },
        ]
    )


def construir_kpis_y_plan(kpis: pd.DataFrame) -> pd.DataFrame:
    plan = pd.DataFrame([
        ["", "", "", "", "", ""],
        ["PLAN DE ACCIÓN", "", "", "", "", ""],
        ["Acción", "Responsable", "Plazo", "KPI de control", "Meta / guardarraíl", "Resultado esperado"],
        ["Validar universo y recalibrar metas mes–región–canal", "BI + Finanzas", "15 días",
         "Cobertura y cumplimiento", "100% grupos trazables", "Metas defendibles"],
        ["Completar captura de stock y lead time", "Operaciones + TI", "30 días",
         "Cobertura inventario", "≥90% combinaciones", "Expansión con soporte"],
        ["Piloto Sucre · Tienda · Electrónica", "Gerencia Sucre", "60 días",
         "Ventas, margen, devoluciones", "+10%; margen ≥34%", "Escalar si rentable"],
        ["Piloto Potosí · Tienda · Hogar", "Gerencia Potosí", "60 días",
         "Ventas, margen, devoluciones", "+10%; margen ≥38%", "Aprovechar margen"],
        ["Revisión de pilotos", "Directorio comercial", "90 días",
         "Guardarraíles", "Solo pilotos aprobados", "Crecimiento medible"],
    ])
    kpi_bloque = pd.DataFrame([["KPIs SELECCIONADOS", "", "", "", "", ""]])
    return pd.concat([kpi_bloque, kpis, plan], ignore_index=True)


def preparar_todo(ruta_caso: Path | str = RUTA_CASO) -> dict:
    raw = cargar_raw(ruta_caso)
    ventas, eliminadas, stats_ventas = preparar_ventas(raw)
    metas, stats_metas = preparar_metas(raw, ventas)
    auditoria_metas = construir_auditoria_14_metas()
    inventario = preparar_inventario(raw)
    kpis = calcular_kpis(ventas, metas)
    return {
        "raw": raw,
        "ventas": ventas,
        "eliminadas": eliminadas,
        "metas": metas,
        "auditoria_metas": auditoria_metas,
        "inventario": inventario,
        "kpis": kpis,
        "guia": construir_guia_proyecto(stats_ventas),
        "registro_calidad": construir_registro_calidad(stats_ventas, stats_metas, inventario, eliminadas),
        "kpis_y_plan": construir_kpis_y_plan(kpis),
        "plan_accion": construir_plan_accion(),
        "resumen_eliminaciones": resumen_eliminaciones(eliminadas),
        "stats_ventas": stats_ventas,
        "stats_metas": stats_metas,
    }

