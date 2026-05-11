"""
================================================================================
verificar_usura.py
================================================================================
Propósito
---------
Verificar si las tasas de interés pactadas en una cartera de créditos superan
la tasa de usura certificada por la Superintendencia Financiera de Colombia,
según el tipo y modalidad de cada crédito y la fecha de corte.

El script produce tres salidas:
  1. CSV enriquecido  : cartera original + columnas de resultado por crédito.
  2. Informe HTML     : reporte visual con resumen ejecutivo, gráficas y tabla
                        detallada de alertas (listo para abrir en navegador o
                        compartir).
  3. Consola          : resumen rápido impreso en pantalla al finalizar.

--------------------------------------------------------------------------------
Archivos de entrada requeridos
--------------------------------------------------------------------------------
tasas_usura.csv
  Tabla de tasas de usura certificadas. Columnas esperadas:
    tipo_credito       – Agrupación de modalidad (ej. "Consumo y Ordinario").
    modalidad          – Modalidad específica     (ej. "Consumo").
    tasa_usura_ea      – Tasa máxima legal en Efectivo Anual (%).
    tasa_usura_mensual – Tasa máxima legal mensual (%).
    vigencia_desde     – Inicio de vigencia (YYYY-MM-DD).
    vigencia_hasta     – Fin de vigencia    (YYYY-MM-DD).

cartera.csv
  Cartera de créditos a validar. Columnas esperadas:
    id_credito            – Identificador único del crédito.
    cliente               – Nombre del titular (opcional pero recomendado).
    tipo_credito          – Debe coincidir con la tabla de tasas.
    modalidad             – Debe coincidir con la tabla de tasas.
    fecha_desembolso      – Fecha de desembolso (YYYY-MM-DD).
    saldo_capital         – Saldo de capital vigente (valor numérico).
    tasa_interes_ea       – Tasa pactada EA (%).
    tasa_interes_mensual  – Tasa pactada mensual (%).
    fecha_corte           – Fecha de corte para la verificación (YYYY-MM-DD).

--------------------------------------------------------------------------------
Uso desde línea de comandos
--------------------------------------------------------------------------------
  # Con archivos por defecto en el mismo directorio
  python verificar_usura.py

  # Rutas personalizadas
  python verificar_usura.py --cartera ruta/mi_cartera.csv \\
                             --tasas   ruta/mis_tasas.csv  \\
                             --output  resultado.csv        \\
                             --informe informe_usura.html

  # Solo CSV, sin informe HTML
  python verificar_usura.py --sin-informe

Parámetros
  --cartera   PATH   CSV de cartera          (default: cartera.csv)
  --tasas     PATH   CSV de tasas de usura   (default: tasas_usura.csv)
  --output    PATH   CSV de salida           (default: resultado_verificacion_usura.csv)
  --informe   PATH   Informe HTML de salida  (default: informe_usura.html)
  --sin-informe      Omite la generación del informe HTML

--------------------------------------------------------------------------------
Dependencias
--------------------------------------------------------------------------------
  pandas   >= 1.3
  jinja2   >= 3.0

  Instalar: pip install pandas jinja2

--------------------------------------------------------------------------------
Lógica de verificación
--------------------------------------------------------------------------------
Para cada crédito de la cartera:
  1. Se busca la tasa de usura cuyo tipo_credito y modalidad coincidan
     (comparación insensible a mayúsculas/espacios) Y cuyo rango
     vigencia_desde–vigencia_hasta incluya la fecha_corte.
  2. Si hay varias vigencias solapadas, se toma la más reciente.
  3. Se compara tasa_interes_ea vs tasa_usura_ea
     y  tasa_interes_mensual vs tasa_usura_mensual.
  4. Si cualquiera supera el límite, alerta_usura = True.

Estados posibles:
  OK                  – Tasa dentro del límite legal.
  SUPERA USURA        – Tasa supera el límite certificado.
  SIN_TASA_REFERENCIA – No se encontró tasa aplicable para ese tipo/período.
  SIN_DATO            – El crédito no tiene valor registrado en esa columna.

--------------------------------------------------------------------------------
Columnas adicionales en el CSV de salida
--------------------------------------------------------------------------------
  tasa_usura_ea_referencia      – Límite EA aplicado.
  tasa_usura_mensual_referencia – Límite mensual aplicado.
  diferencia_ea                 – tasa_interes_ea − tasa_usura_ea (positivo = exceso).
  diferencia_mensual            – tasa_interes_mensual − tasa_usura_mensual.
  estado_tasa_ea                – Estado de la validación EA.
  estado_tasa_mensual           – Estado de la validación mensual.
  alerta_usura                  – True si supera el límite en alguna de las dos tasas.

================================================================================
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


# ════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE COLUMNAS
# Si tus CSVs usan nombres distintos, cámbialos aquí únicamente.
# ════════════════════════════════════════════════════════════════════════════

COL_ID               = "id_credito"
COL_CLIENTE          = "cliente"
COL_TIPO             = "tipo_credito"
COL_MODALIDAD        = "modalidad"
COL_TASA_EA          = "tasa_interes_ea"
COL_TASA_MEN         = "tasa_interes_mensual"
COL_SALDO            = "saldo_capital"
COL_FECHA_DESEMBOLSO = "fecha_desembolso"
COL_FECHA_CORTE      = "fecha_corte"
COL_USURA_EA         = "tasa_usura_ea"
COL_USURA_MEN        = "tasa_usura_mensual"
COL_VIGENCIA_DESDE   = "vigencia_desde"
COL_VIGENCIA_HASTA   = "vigencia_hasta"


# ════════════════════════════════════════════════════════════════════════════
# CARGA Y VALIDACIÓN
# ════════════════════════════════════════════════════════════════════════════

def cargar_csv(ruta: str, nombre: str) -> pd.DataFrame:
    """
    Carga un archivo CSV y normaliza los nombres de columna.

    Normalización: strip de espacios, minúsculas, espacios internos → guion bajo.
    El encoding utf-8-sig tolera archivos exportados desde Excel con BOM.

    Parameters
    ----------
    ruta   : Ruta al archivo CSV.
    nombre : Nombre descriptivo para mensajes de error.

    Returns
    -------
    pd.DataFrame con columnas normalizadas.

    Raises
    ------
    SystemExit si el archivo no existe o no puede leerse.
    """
    try:
        df = pd.read_csv(ruta, sep=",", encoding="utf-8-sig")
        df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
        print(f"  ✔ {nombre}: {len(df)} registros cargados desde '{ruta}'")
        return df
    except FileNotFoundError:
        print(f"\n  ✘ ERROR: No se encontró el archivo '{ruta}'")
        sys.exit(1)
    except Exception as e:
        print(f"\n  ✘ ERROR al leer '{ruta}': {e}")
        sys.exit(1)


def validar_columnas(df: pd.DataFrame, columnas: list, nombre: str) -> None:
    """
    Verifica que el DataFrame contenga todas las columnas requeridas.

    Parameters
    ----------
    df       : DataFrame a validar.
    columnas : Lista de nombres de columna requeridos.
    nombre   : Nombre del archivo (para mensajes de error).

    Raises
    ------
    SystemExit si falta alguna columna requerida.
    """
    faltantes = [c for c in columnas if c not in df.columns]
    if faltantes:
        print(f"\n  ✘ ERROR: Al archivo '{nombre}' le faltan columnas: {faltantes}")
        print(f"     Columnas encontradas: {list(df.columns)}")
        sys.exit(1)


def parsear_fechas(df: pd.DataFrame, columnas: list) -> pd.DataFrame:
    """
    Convierte columnas de texto a tipo datetime (formato YYYY-MM-DD).

    Las fechas inválidas o ausentes quedan como NaT sin interrumpir el proceso.

    Parameters
    ----------
    df       : DataFrame con las columnas a convertir.
    columnas : Nombres de las columnas que contienen fechas.

    Returns
    -------
    pd.DataFrame con las columnas convertidas.
    """
    for col in columnas:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


# ════════════════════════════════════════════════════════════════════════════
# LÓGICA DE VERIFICACIÓN
# ════════════════════════════════════════════════════════════════════════════

def obtener_tasa_usura(fila: pd.Series, df_tasas: pd.DataFrame) -> dict | None:
    """
    Busca la tasa de usura aplicable a un crédito específico.

    Criterios de búsqueda:
      - tipo_credito y modalidad coinciden (insensible a mayúsculas/espacios).
      - fecha_corte cae dentro de [vigencia_desde, vigencia_hasta].
    Si hay varias coincidencias, se toma la de mayor vigencia_desde.

    Parameters
    ----------
    fila     : Fila de la cartera (un crédito).
    df_tasas : Tabla de tasas de usura certificadas.

    Returns
    -------
    dict con 'tasa_usura_ea' y 'tasa_usura_mensual', o None si no hay match.
    """
    fecha = fila[COL_FECHA_CORTE]
    mascara = (
        (df_tasas[COL_TIPO].str.strip().str.lower()
         == str(fila[COL_TIPO]).strip().lower())
        & (df_tasas[COL_MODALIDAD].str.strip().str.lower()
           == str(fila[COL_MODALIDAD]).strip().lower())
        & (df_tasas[COL_VIGENCIA_DESDE] <= fecha)
        & (df_tasas[COL_VIGENCIA_HASTA] >= fecha)
    )
    coincidencias = df_tasas[mascara]
    if coincidencias.empty:
        return None
    fila_tasa = coincidencias.sort_values(COL_VIGENCIA_DESDE, ascending=False).iloc[0]
    return {
        "tasa_usura_ea":      fila_tasa[COL_USURA_EA],
        "tasa_usura_mensual": fila_tasa[COL_USURA_MEN],
    }


def verificar_usura(df_cartera: pd.DataFrame, df_tasas: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica la verificación de tasa de usura a toda la cartera.

    Para cada crédito obtiene la tasa de usura aplicable y compara la tasa EA
    y la mensual. Agrega columnas de resultado al DataFrame original.

    Parameters
    ----------
    df_cartera : Cartera de créditos con tasas pactadas.
    df_tasas   : Tabla de tasas de usura certificadas.

    Returns
    -------
    pd.DataFrame con columnas adicionales:
      tasa_usura_ea_referencia, tasa_usura_mensual_referencia,
      diferencia_ea, diferencia_mensual,
      estado_tasa_ea, estado_tasa_mensual, alerta_usura.
    """
    resultados = []

    for _, fila in df_cartera.iterrows():
        tasa_info = obtener_tasa_usura(fila, df_tasas)

        if tasa_info is None:
            estado_ea = estado_men = "SIN_TASA_REFERENCIA"
            excede_ea = excede_men = None
            usura_ea = usura_men = None
        else:
            usura_ea  = tasa_info["tasa_usura_ea"]
            usura_men = tasa_info["tasa_usura_mensual"]

            # Validación EA
            if pd.isna(fila.get(COL_TASA_EA)):
                excede_ea, estado_ea = None, "SIN_DATO"
            else:
                excede_ea = float(fila[COL_TASA_EA]) > float(usura_ea)
                estado_ea = "SUPERA USURA" if excede_ea else "OK"

            # Validación mensual
            if pd.isna(fila.get(COL_TASA_MEN)):
                excede_men, estado_men = None, "SIN_DATO"
            else:
                excede_men = float(fila[COL_TASA_MEN]) > float(usura_men)
                estado_men = "SUPERA USURA" if excede_men else "OK"

        fila_resultado = fila.to_dict()
        fila_resultado.update({
            "tasa_usura_ea_referencia": usura_ea,
            "tasa_usura_mensual_referencia": usura_men,
            "diferencia_ea": (
                round(float(fila[COL_TASA_EA]) - float(usura_ea), 4)
                if (usura_ea is not None and pd.notna(fila.get(COL_TASA_EA)))
                else None
            ),
            "diferencia_mensual": (
                round(float(fila[COL_TASA_MEN]) - float(usura_men), 4)
                if (usura_men is not None and pd.notna(fila.get(COL_TASA_MEN)))
                else None
            ),
            "estado_tasa_ea":      estado_ea,
            "estado_tasa_mensual": estado_men,
            "alerta_usura": (
                (excede_ea or excede_men)
                if (excede_ea is not None or excede_men is not None)
                else None
            ),
        })
        resultados.append(fila_resultado)

    return pd.DataFrame(resultados)


# ════════════════════════════════════════════════════════════════════════════
# INFORME HTML
# ════════════════════════════════════════════════════════════════════════════

PLANTILLA_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Informe de Verificación de Tasa de Usura</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
    :root {
      --bg:      #0d0f14; --surface: #13161e; --border: #1e2330;
      --accent:  #00e5a0; --danger:  #ff4d6d; --warn:   #ffc845;
      --muted:   #4a5068; --text:    #d4d8e8; --dim:    #7a8099;
      --mono: 'IBM Plex Mono', monospace; --sans: 'IBM Plex Sans', sans-serif;
    }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: var(--bg); color: var(--text); font-family: var(--sans);
           font-size: 14px; line-height: 1.6; padding: 40px 28px 80px; }

    /* Header */
    .header { border-left: 3px solid var(--accent); padding-left: 20px; margin-bottom: 48px; }
    .header .label { font-family: var(--mono); font-size: 11px; color: var(--accent);
                     letter-spacing: .15em; text-transform: uppercase; margin-bottom: 6px; }
    .header h1 { font-size: 26px; font-weight: 700; color: #fff; letter-spacing: -.02em; }
    .header .meta { font-family: var(--mono); font-size: 12px; color: var(--dim); margin-top: 6px; }

    /* KPIs */
    .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(155px,1fr));
            gap: 12px; margin-bottom: 40px; }
    .kpi { background: var(--surface); border: 1px solid var(--border);
           border-radius: 8px; padding: 20px 18px 16px; }
    .kpi .kpi-label { font-family: var(--mono); font-size: 10px; color: var(--dim);
                      letter-spacing: .12em; text-transform: uppercase; margin-bottom: 10px; }
    .kpi .kpi-value { font-size: 34px; font-weight: 700; line-height: 1; }
    .kpi.ok      .kpi-value { color: var(--accent); }
    .kpi.alerta  .kpi-value { color: var(--danger); }
    .kpi.warn    .kpi-value { color: var(--warn);   }
    .kpi.neutral .kpi-value { color: #fff; }

    /* Section title */
    .section-title { font-family: var(--mono); font-size: 11px; letter-spacing: .15em;
                     text-transform: uppercase; color: var(--accent); border-bottom: 1px solid var(--border);
                     padding-bottom: 8px; margin-bottom: 20px; margin-top: 48px; }

    /* Bar chart */
    .bars { display: flex; flex-direction: column; gap: 10px; margin-bottom: 8px; }
    .bar-row { display: flex; align-items: center; gap: 12px; }
    .bar-name { font-size: 12px; color: var(--dim); width: 190px; flex-shrink: 0;
                white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .bar-track { flex: 1; height: 24px; background: var(--border); border-radius: 3px; overflow: hidden; }
    .bar-fill { height: 100%; border-radius: 3px; background: var(--danger); display: flex;
                align-items: center; padding-left: 8px; font-family: var(--mono);
                font-size: 11px; color: #fff; min-width: 40px; }
    .bar-fill.safe { background: rgba(0,229,160,.25); color: var(--accent); }

    /* Alert cards */
    .alert-cards { display: flex; flex-direction: column; gap: 12px; }
    .alert-card { background: rgba(255,77,109,.06); border: 1px solid rgba(255,77,109,.25);
                  border-radius: 8px; padding: 16px 20px; display: grid;
                  grid-template-columns: auto 1fr auto; gap: 0 24px; align-items: center; }
    .alert-id { font-family: var(--mono); font-size: 20px; font-weight: 600; color: var(--danger); }
    .alert-info .name { font-weight: 600; color: #fff; }
    .alert-info .detail { font-size: 12px; color: var(--dim); font-family: var(--mono); margin-top: 3px; }
    .alert-rates { text-align: right; }
    .alert-rates .rate-pacted { font-size: 22px; font-weight: 700; color: var(--danger); font-family: var(--mono); }
    .alert-rates .rate-limit  { font-size: 12px; color: var(--dim); font-family: var(--mono); }
    .alert-rates .exceso { font-size: 11px; color: var(--warn); font-family: var(--mono); margin-top: 2px; }

    /* Table */
    .table-wrap { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    thead th { font-family: var(--mono); font-size: 10px; letter-spacing: .1em;
               text-transform: uppercase; color: var(--dim); border-bottom: 1px solid var(--border);
               padding: 10px 12px; text-align: left; white-space: nowrap; }
    tbody tr { border-bottom: 1px solid var(--border); }
    tbody tr:hover { background: rgba(255,255,255,.025); }
    tbody td { padding: 11px 12px; vertical-align: middle; }
    .num { font-family: var(--mono); }
    .pos { color: var(--danger); } .neg { color: var(--accent); }

    .badge { display: inline-block; padding: 2px 9px; border-radius: 99px;
             font-family: var(--mono); font-size: 10px; font-weight: 600; letter-spacing: .05em; }
    .b-ok     { background: rgba(0,229,160,.12); color: var(--accent); }
    .b-alerta { background: rgba(255,77,109,.12); color: var(--danger); }
    .b-warn   { background: rgba(255,200,69,.12); color: var(--warn); }

    /* Footer */
    .footer { margin-top: 64px; padding-top: 20px; border-top: 1px solid var(--border);
              font-family: var(--mono); font-size: 11px; color: var(--muted); }
  </style>
</head>
<body>

<div class="header">
  <div class="label">Informe de Cumplimiento Regulatorio</div>
  <h1>Verificación de Tasa de Usura</h1>
  <div class="meta">
    Generado: {{ fecha_generacion }} &nbsp;│&nbsp;
    Cartera: {{ archivo_cartera }} &nbsp;│&nbsp;
    Tasas: {{ archivo_tasas }}
  </div>
</div>

<div class="kpis">
  <div class="kpi neutral"><div class="kpi-label">Total créditos</div><div class="kpi-value">{{ total }}</div></div>
  <div class="kpi ok">    <div class="kpi-label">Cumplen límite</div><div class="kpi-value">{{ cumplen }}</div></div>
  <div class="kpi alerta"><div class="kpi-label">Superan usura</div><div class="kpi-value">{{ superan }}</div></div>
  <div class="kpi warn">  <div class="kpi-label">Sin referencia</div><div class="kpi-value">{{ sin_ref }}</div></div>
  <div class="kpi alerta"><div class="kpi-label">% en alerta</div>  <div class="kpi-value">{{ pct_alerta }}%</div></div>
</div>

<div class="section-title">Distribución de alertas por modalidad</div>
<div class="bars">
  {% for row in dist_modalidad %}
  <div class="bar-row">
    <div class="bar-name" title="{{ row.modalidad }}">{{ row.modalidad }}</div>
    <div class="bar-track">
      <div class="bar-fill {% if row.superan == 0 %}safe{% endif %}" style="width:{{ row.pct_bar }}%">
        {{ row.superan }}/{{ row.total }}
      </div>
    </div>
  </div>
  {% endfor %}
</div>

{% if alertas %}
<div class="section-title">Créditos que superan la tasa de usura ({{ superan }})</div>
<div class="alert-cards">
  {% for a in alertas %}
  <div class="alert-card">
    <div class="alert-id">#{{ a.id }}</div>
    <div class="alert-info">
      <div class="name">{{ a.cliente }}</div>
      <div class="detail">{{ a.modalidad }} &nbsp;│&nbsp; Corte: {{ a.fecha_corte }}</div>
    </div>
    <div class="alert-rates">
      <div class="rate-pacted">{{ a.tasa_ea }}% EA</div>
      <div class="rate-limit">Límite: {{ a.usura_ea }}% EA</div>
      <div class="exceso">+{{ a.diferencia_ea }}% exceso</div>
    </div>
  </div>
  {% endfor %}
</div>
{% endif %}

<div class="section-title">Detalle completo de la cartera</div>
<div class="table-wrap">
  <table>
    <thead>
      <tr>
        <th>ID</th><th>Cliente</th><th>Modalidad</th><th>Corte</th>
        <th>Tasa EA %</th><th>Usura EA %</th><th>Diferencia</th>
        <th>Tasa Men %</th><th>Usura Men %</th>
        <th>Estado EA</th><th>Estado Men</th>
      </tr>
    </thead>
    <tbody>
      {% for r in filas %}
      <tr>
        <td class="num">{{ r.id }}</td>
        <td>{{ r.cliente }}</td>
        <td>{{ r.modalidad }}</td>
        <td class="num">{{ r.fecha_corte }}</td>
        <td class="num">{{ r.tasa_ea }}</td>
        <td class="num">{{ r.usura_ea }}</td>
        <td class="num {% if r.dif_sign == '+' %}pos{% elif r.dif_sign == '-' %}neg{% endif %}">{{ r.diferencia_ea }}</td>
        <td class="num">{{ r.tasa_men }}</td>
        <td class="num">{{ r.usura_men }}</td>
        <td>
          {% if r.estado_ea == 'OK' %}<span class="badge b-ok">✔ OK</span>
          {% elif r.estado_ea == 'SUPERA USURA' %}<span class="badge b-alerta">⚠ SUPERA</span>
          {% else %}<span class="badge b-warn">{{ r.estado_ea }}</span>{% endif %}
        </td>
        <td>
          {% if r.estado_men == 'OK' %}<span class="badge b-ok">✔ OK</span>
          {% elif r.estado_men == 'SUPERA USURA' %}<span class="badge b-alerta">⚠ SUPERA</span>
          {% else %}<span class="badge b-warn">{{ r.estado_men }}</span>{% endif %}
        </td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
</div>

<div class="footer">
  Generado por verificar_usura.py &nbsp;│&nbsp;
  Tasas certificadas por la Superintendencia Financiera de Colombia &nbsp;│&nbsp;
  {{ fecha_generacion }}
</div>
</body>
</html>
"""


def generar_informe_html(
    df_resultado: pd.DataFrame,
    ruta_informe: str,
    archivo_cartera: str,
    archivo_tasas: str,
) -> None:
    """
    Genera un informe HTML visual a partir del DataFrame de resultados.

    Contenido del informe:
      - Tarjetas KPI: totales de créditos, cumplen, alertas, sin referencia y % en alerta.
      - Gráfico de barras por modalidad mostrando alertas vs total.
      - Tarjetas detalladas de cada crédito que supera el límite.
      - Tabla completa de la cartera con badges de color por estado.

    Parameters
    ----------
    df_resultado    : DataFrame enriquecido por verificar_usura().
    ruta_informe    : Ruta del archivo HTML a generar.
    archivo_cartera : Nombre del CSV de cartera (encabezado del informe).
    archivo_tasas   : Nombre del CSV de tasas   (encabezado del informe).
    """
    from jinja2 import Template

    total    = len(df_resultado)
    superan  = int((df_resultado["estado_tasa_ea"] == "SUPERA USURA").sum())
    sin_ref  = int((df_resultado["estado_tasa_ea"] == "SIN_TASA_REFERENCIA").sum())
    cumplen  = total - superan - sin_ref
    pct_alerta = round(superan / total * 100, 1) if total else 0

    # Distribución por modalidad
    dist = (
        df_resultado.groupby(COL_MODALIDAD)
        .apply(lambda g: pd.Series({
            "total":   len(g),
            "superan": (g["estado_tasa_ea"] == "SUPERA USURA").sum(),
        }))
        .reset_index()
    )
    max_total = dist["total"].max() or 1
    dist_modalidad = [
        {
            "modalidad": row[COL_MODALIDAD],
            "total":     int(row["total"]),
            "superan":   int(row["superan"]),
            "pct_bar":   round(row["total"] / max_total * 100, 1),
        }
        for _, row in dist.iterrows()
    ]

    # Alertas detalladas
    df_alertas = df_resultado[df_resultado["estado_tasa_ea"] == "SUPERA USURA"]
    alertas = [
        {
            "id":            str(r.get(COL_ID, "N/A")),
            "cliente":       str(r.get(COL_CLIENTE, "—")),
            "modalidad":     str(r.get(COL_MODALIDAD, "—")),
            "fecha_corte":   str(r.get(COL_FECHA_CORTE, ""))[:10],
            "tasa_ea":       f"{float(r[COL_TASA_EA]):.2f}",
            "usura_ea":      f"{float(r['tasa_usura_ea_referencia']):.2f}",
            "diferencia_ea": f"{float(r['diferencia_ea']):.2f}",
        }
        for _, r in df_alertas.iterrows()
    ]

    # Filas tabla completa
    def fmt(val, dec=2):
        try:
            return f"{float(val):.{dec}f}" if pd.notna(val) else "—"
        except (TypeError, ValueError):
            return "—"

    filas = []
    for _, r in df_resultado.iterrows():
        dif = r.get("diferencia_ea")
        try:
            dif_sign = "+" if float(dif) > 0 else ("-" if float(dif) < 0 else "")
        except (TypeError, ValueError):
            dif_sign = ""
        filas.append({
            "id":           str(r.get(COL_ID, "N/A")),
            "cliente":      str(r.get(COL_CLIENTE, "—")),
            "modalidad":    str(r.get(COL_MODALIDAD, "—")),
            "fecha_corte":  str(r.get(COL_FECHA_CORTE, ""))[:10],
            "tasa_ea":      fmt(r.get(COL_TASA_EA)),
            "usura_ea":     fmt(r.get("tasa_usura_ea_referencia")),
            "diferencia_ea":fmt(dif),
            "dif_sign":     dif_sign,
            "tasa_men":     fmt(r.get(COL_TASA_MEN)),
            "usura_men":    fmt(r.get("tasa_usura_mensual_referencia")),
            "estado_ea":    str(r.get("estado_tasa_ea", "")),
            "estado_men":   str(r.get("estado_tasa_mensual", "")),
        })

    contexto = dict(
        fecha_generacion=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        archivo_cartera=archivo_cartera,
        archivo_tasas=archivo_tasas,
        total=total, cumplen=cumplen, superan=superan,
        sin_ref=sin_ref, pct_alerta=pct_alerta,
        dist_modalidad=dist_modalidad,
        alertas=alertas, filas=filas,
    )

    Path(ruta_informe).write_text(Template(PLANTILLA_HTML).render(**contexto), encoding="utf-8")
    print(f"  ✔ Informe HTML generado en:     '{ruta_informe}'")


# ════════════════════════════════════════════════════════════════════════════
# RESUMEN EN CONSOLA
# ════════════════════════════════════════════════════════════════════════════

def imprimir_resumen(df_resultado: pd.DataFrame) -> None:
    """
    Imprime un resumen ejecutivo en consola al finalizar el proceso.

    Muestra conteos globales y el detalle de cada crédito con alerta,
    incluyendo el exceso sobre la tasa de usura en puntos porcentuales.

    Parameters
    ----------
    df_resultado : DataFrame enriquecido por verificar_usura().
    """
    total       = len(df_resultado)
    sin_ref     = (df_resultado["estado_tasa_ea"] == "SIN_TASA_REFERENCIA").sum()
    alertas_ea  = (df_resultado["estado_tasa_ea"] == "SUPERA USURA").sum()
    alertas_men = (df_resultado["estado_tasa_mensual"] == "SUPERA USURA").sum()
    cumplen     = total - alertas_ea - sin_ref

    print("\n" + "═" * 58)
    print("  RESUMEN DE VERIFICACIÓN DE TASA DE USURA")
    print("═" * 58)
    print(f"  Total créditos analizados      : {total:>6}")
    print(f"  Sin tasa de referencia         : {sin_ref:>6}")
    print(f"  ✔ Cumplen límite (tasa EA)     : {cumplen:>6}")
    print(f"  ⚠ Superan usura (tasa EA)      : {alertas_ea:>6}")
    print(f"  ⚠ Superan usura (tasa mensual) : {alertas_men:>6}")
    print("═" * 58)

    if alertas_ea > 0:
        print("\n  CRÉDITOS CON ALERTA (superan tasa de usura EA):\n")
        for _, r in df_resultado[df_resultado["estado_tasa_ea"] == "SUPERA USURA"].iterrows():
            print(
                f"    ID {str(r.get(COL_ID,'N/A')):>4} | "
                f"{str(r.get(COL_MODALIDAD,'')):<22} | "
                f"Tasa: {float(r[COL_TASA_EA]):>5.2f}% EA | "
                f"Usura: {float(r['tasa_usura_ea_referencia']):>5.2f}% | "
                f"Exceso: +{float(r['diferencia_ea']):.2f}%"
            )
    print()


# ════════════════════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ════════════════════════════════════════════════════════════════════════════

def main() -> None:
    """
    Punto de entrada principal del script.

    Parsea argumentos CLI, orquesta la carga, validación, verificación
    y generación de salidas (CSV + HTML + consola).
    """
    parser = argparse.ArgumentParser(
        description="Verifica tasa de usura en cartera de créditos y genera informe.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python verificar_usura.py\n"
            "  python verificar_usura.py --cartera data/cartera.csv --tasas data/tasas.csv\n"
            "  python verificar_usura.py --output resultado.csv --informe reporte.html\n"
            "  python verificar_usura.py --sin-informe\n"
        ),
    )
    parser.add_argument("--cartera",     default="cartera.csv",
                        help="CSV de cartera (default: cartera.csv)")
    parser.add_argument("--tasas",       default="tasas_usura.csv",
                        help="CSV de tasas de usura (default: tasas_usura.csv)")
    parser.add_argument("--output",      default="resultado_verificacion_usura.csv",
                        help="CSV de salida (default: resultado_verificacion_usura.csv)")
    parser.add_argument("--informe",     default="informe_usura.html",
                        help="Informe HTML de salida (default: informe_usura.html)")
    parser.add_argument("--sin-informe", action="store_true",
                        help="Omite la generación del informe HTML")
    args = parser.parse_args()

    print("\n" + "═" * 58)
    print("  VERIFICADOR DE TASA DE USURA")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("═" * 58)
    print("\n  Cargando archivos...")

    df_tasas   = cargar_csv(args.tasas,   "Tasas de usura")
    df_cartera = cargar_csv(args.cartera, "Cartera")

    validar_columnas(df_tasas,   [COL_TIPO, COL_MODALIDAD, COL_USURA_EA, COL_USURA_MEN,
                                  COL_VIGENCIA_DESDE, COL_VIGENCIA_HASTA], "tasas_usura.csv")
    validar_columnas(df_cartera, [COL_TIPO, COL_MODALIDAD, COL_TASA_EA,
                                  COL_TASA_MEN, COL_FECHA_CORTE], "cartera.csv")

    df_tasas   = parsear_fechas(df_tasas,   [COL_VIGENCIA_DESDE, COL_VIGENCIA_HASTA])
    df_cartera = parsear_fechas(df_cartera, [COL_FECHA_CORTE])

    print("\n  Verificando tasas de usura...")
    df_resultado = verificar_usura(df_cartera, df_tasas)

    print("\n  Generando salidas...")
    df_resultado.to_csv(args.output, index=False, encoding="utf-8-sig")
    print(f"  ✔ CSV de resultados guardado en: '{args.output}'")

    if not args.sin_informe:
        generar_informe_html(
            df_resultado,
            ruta_informe=args.informe,
            archivo_cartera=args.cartera,
            archivo_tasas=args.tasas,
        )

    imprimir_resumen(df_resultado)


if __name__ == "__main__":
    main()
