# informe_financiero_final_funcional.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.utils import ImageReader
import matplotlib.pyplot as plt
from io import BytesIO
import math
import datetime
import numpy as np

# Definición de colores
# Usamos STRINGS HEX para Matplotlib y para <font color='...'>
HEX_PRINCIPAL = "#003366"  # Azul Marino
HEX_ACENTO = "#4CAF50"     # Verde/Teal
HEX_FONDO_TABLA = "#EEECEC" # Gris Claro
HEX_CAJA = "#DCEFFD"       # Azul muy claro para cajas

# Usamos objetos ReportLab Color para setFillColor o TableStyle BACKGROUND
COLOR_PRINCIPAL = colors.HexColor(HEX_PRINCIPAL)
COLOR_ACENTO = colors.HexColor(HEX_ACENTO)
COLOR_FONDO_TABLA_OBJ = colors.HexColor(HEX_FONDO_TABLA)
COLOR_CAJA_OBJ = colors.HexColor(HEX_CAJA) 

# -----------------------------
# UTILIDADES (Funciones auxiliares, etc.)
# -----------------------------
def safe_div(a, b):
    try:
        return a / b if b != 0 else None
    except:
        return None

def fmt_num(val, digits=2):
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return "N/A"
    try:
        s = f"{val:,.{digits}f}"
    except:
        return str(val)
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

# Funciones de cálculos
def calcular_ratios_from_inputs(d):
    # Normalizar y sacar valores
    AC = d.get("activo_corriente") or 0.0
    ANC = d.get("activo_no_corriente") or 0.0
    PC = d.get("pasivo_corriente") or 0.0
    PNC = d.get("pasivo_no_corriente") or 0.0
    PN = d.get("patrimonio_neto") or 0.0
    Ventas = d.get("ventas") or 0.0
    Costo = d.get("costo_ventas") or 0.0
    BN = d.get("beneficio_neto") or 0.0
    Deudores = d.get("deudores") or 0.0
    Inventario = d.get("inventario") or 0.0
    Caja = d.get("caja") or 0.0
    i_input = d.get("i") if d.get("i") is not None else 0.05
    GFin = d.get("gastos_financieros") or 0.0 # C5
    # Datos CCE (A4)
    dias_inv = d.get("dias_inventario") or 0.0
    dias_clie = d.get("dias_clientes") or 0.0
    dias_prov = d.get("dias_proveedores") or 0.0


    Activo = AC + ANC
    Pasivo = PC + PNC
    DeudaTotal = Pasivo # PC + PNC

    ratios = {}
    
    # C5.a: Calcular costo promedio de deuda (i)
    costo_deuda_i = safe_div(GFin, DeudaTotal) if DeudaTotal != 0 else 0.0
    ratios["Costo Deuda (i)"] = costo_deuda_i

    # Fondo de Maniobra (A1)
    ratios["Fondo Maniobra"] = AC - PC
    ratios["Fondo Maniobra Alternativo"] = PN + PNC - ANC

    # Liquidez (C1)
    ratios["Liquidez General"] = safe_div(AC, PC)
    ratios["Tesorería"] = safe_div((Caja + Deudores), PC)
    ratios["Disponibilidad"] = safe_div(Caja, PC)

    # Solvencia (C2)
    ratios["Garantía"] = safe_div(Activo, Pasivo)
    ratios["Autonomía"] = safe_div(PN, Pasivo)
    ratios["Calidad Deuda"] = safe_div(PC, Pasivo)

    # Rentabilidades (C3)
    BAII = Ventas - Costo
    ratios["RAT"] = safe_div(BAII, Activo)
    ratios["RRP"] = safe_div(BN, PN) # RRP normal, sin apalancamiento

    # C5.c: Apalancamiento Financiero (Usando la 'i' calculada)
    D = DeudaTotal
    if PN != 0:
        RAT = ratios["RAT"] or 0.0
        # RRP = RAT + (D/PN) * (RAT - i)
        apalancamiento_term = safe_div(D, PN) * (RAT - costo_deuda_i)
        ratios["RRP Apalancada"] = RAT + apalancamiento_term
        ratios["Efecto Apalancamiento"] = apalancamiento_term
    else:
        ratios["RRP Apalancada"] = None
        ratios["Efecto Apalancamiento"] = None

    # Store raw items for later use
    ratios["_AC"] = AC
    ratios["_ANC"] = ANC
    ratios["_PC"] = PC
    ratios["_PNC"] = PNC
    ratios["_PN"] = PN
    ratios["_Ventas"] = Ventas
    ratios["_Costo"] = Costo
    ratios["_BAII"] = BAII
    ratios["_BN"] = BN
    ratios["_Deudores"] = Deudores
    ratios["_Inventario"] = Inventario
    ratios["_Caja"] = Caja
    ratios["_i_input"] = i_input
    ratios["_GastosFin"] = GFin
    ratios["_ActivoTotal"] = Activo
    ratios["_PasivoTotal"] = Pasivo
    ratios["_DeudaTotal"] = DeudaTotal
    ratios["_dias_inventario"] = dias_inv # A4
    ratios["_dias_clientes"] = dias_clie # A4
    ratios["_dias_proveedores"] = dias_prov # A4

    return ratios

# Funciones de análisis y diagnóstico
def clasificar_situacion_patrimonial_v2(r):
    AC = r.get("_AC") or 0.0
    PC = r.get("_PC") or 0.0
    PN = r.get("_PN") or 0.0
    Activo = r.get("_ActivoTotal") or 0.0
    fm = r.get("Fondo Maniobra")
    
    justificacion = []

    if PN < 0:
        justificacion.append(f"Patrimonio Neto ({fmt_num(PN)}) es negativo.")
        return "Crisis / Desequilibrio a L/P (Quiebra técnica)", " | ".join(justificacion)
    
    if Activo > 0.0 and safe_div(PN, Activo) > 0.85:
        justificacion.append(f"PN/Activo ({fmt_num(safe_div(PN, Activo)*100)}%) es muy alto.")
        justificacion.append(f"Fondo de Maniobra ({fmt_num(fm)}) es muy grande.")
        return "Equilibrio Total / Estabilidad Máxima", " | ".join(justificacion)

    if fm < 0 and AC < PC:
        justificacion.append(f"Fondo de Maniobra ({fmt_num(fm)}) es negativo.")
        justificacion.append("El Activo Corriente no cubre el Pasivo Corriente.")
        return "Insolvencia / Suspensión de pagos", " | ".join(justificacion)

    if fm < 0:
        justificacion.append(f"Fondo de Maniobra ({fmt_num(fm)}) es negativo.")
        justificacion.append(f"El Pasivo Corriente ({fmt_num(PC)}) supera el Activo Corriente ({fmt_num(AC)}).")
        return "Desequilibrio/Tensión Financiera Normal", " | ".join(justificacion)

    justificacion.append(f"Fondo de Maniobra ({fmt_num(fm)}) es positivo.")
    justificacion.append(f"El AC ({fmt_num(AC)}) financia completamente el PC ({fmt_num(PC)}).")
    return "Equilibrio Normal / Estabilidad Normal", " | ".join(justificacion)

def generar_analisis_vertical(r):
    AT = r.get("_ActivoTotal") or 0.0
    if AT == 0: return "Activo Total es cero. Análisis vertical no es posible.", "N/A", "N/A"
        
    AC = r.get("_AC") or 0.0; ANC = r.get("_ANC") or 0.0
    PC = r.get("_PC") or 0.0; PNC = r.get("_PNC") or 0.0
    PN = r.get("_PN") or 0.0
    
    AC_pct = safe_div(AC, AT) * 100; ANC_pct = safe_div(ANC, AT) * 100
    PC_pct = safe_div(PC, AT) * 100; PNC_pct = safe_div(PNC, AT) * 100
    PN_pct = safe_div(PN, AT) * 100
    
    res = [
        f"**Activo Corriente:** {fmt_num(AC_pct)}%",
        f"**Activo No Corriente:** {fmt_num(ANC_pct)}%",
        f"**Pasivo Corriente:** {fmt_num(PC_pct)}%",
        f"**Pasivo No Corriente:** {fmt_num(PNC_pct)}%",
        f"**Patrimonio Neto:** {fmt_num(PN_pct)}%"
    ]
    
    if AC_pct > ANC_pct:
        econo_str = f"Predomina el **Activo Corriente** ({fmt_num(AC_pct)}%), típica de empresas de ciclo operativo rápido (comercial/servicios)."
    else:
        econo_str = f"Predomina el **Activo No Corriente** ({fmt_num(ANC_pct)}%), típica de empresas con alta inversión en infraestructura (industrial)."

    Deuda_Total = PC + PNC
    if PN > Deuda_Total:
        finan_str = f"Buena solvencia: El **Patrimonio Neto** ({fmt_num(PN_pct)}%) supera la deuda total ({fmt_num(PC_pct + PNC_pct)}%)."
    else:
        finan_str = f"Dependencia de terceros: La deuda total ({fmt_num(PC_pct + PNC_pct)}%) supera o iguala al PN ({fmt_num(PN_pct)}%)."
        
    return "\n".join(res), econo_str, finan_str

def generar_analisis_horizontal(r23, r24):
    AC23 = r23.get("_AC") or 0.0; AC24 = r24.get("_AC") or 0.0
    ANC23 = r23.get("_ANC") or 0.0; ANC24 = r24.get("_ANC") or 0.0
    AT23 = r23.get("_ActivoTotal") or 0.0; AT24 = r24.get("_ActivoTotal") or 0.0
    PC23 = r23.get("_PC") or 0.0; PC24 = r24.get("_PC") or 0.0
    PNC23 = r23.get("_PNC") or 0.0; PNC24 = r24.get("_PNC") or 0.0
    PN23 = r23.get("_PN") or 0.0; PN24 = r24.get("_PN") or 0.0
    
    crec_AC = safe_div(AC24 - AC23, AC23) * 100 if AC23 != 0 else (100 if AC24 > 0 else 0)
    crec_ANC = safe_div(ANC24 - ANC23, ANC23) * 100 if ANC23 != 0 else (100 if ANC24 > 0 else 0)
    
    crecimientos = []
    crecimientos.append(("Activo Corriente", crec_AC, AC24-AC23))
    crecimientos.append(("Activo No Corriente", crec_ANC, ANC24-ANC23))
    crecimientos.sort(key=lambda x: x[2], reverse=True)
    activo_mas_crecido = f"{crecimientos[0][0]} ({fmt_num(crecimientos[0][1])}% de crecimiento)."
    
    var_PC = PC24 - PC23; var_PNC = PNC24 - PNC23; var_PN = PN24 - PN23
    crec_AT = safe_div(AT24 - AT23, AT23) * 100 if AT23 != 0 else (100 if AT24 > 0 else 0)

    contribuciones = [("Pasivo Corriente", var_PC), ("Pasivo No Corriente", var_PNC), ("Patrimonio Neto", var_PN)]
    contribuciones.sort(key=lambda x: x[1], reverse=True)
    
    financiamiento_data = []
    if crec_AT != 0:
        financiamiento_data.append(f"- Variación Pasivo Corriente: {fmt_num(var_PC)}")
        financiamiento_data.append(f"- Variación Pasivo No Corriente: {fmt_num(var_PNC)}")
        financiamiento_data.append(f"- Variación Patrimonio Neto: {fmt_num(var_PN)}")
    
    financiacion_principal = f"El crecimiento se financió principalmente a través de: **{contribuciones[0][0]}** (Aporte absoluto: {fmt_num(contribuciones[0][1])})."

    return {
        "Activo Mas Crecido": activo_mas_crecido,
        "Crecimiento Total Activo": crec_AT,
        "Financiamiento Detalle": "\n".join(financiamiento_data),
        "Financiacion Principal": financiacion_principal
    }

def calcular_cce(dias_inv, dias_clie, dias_prov):
    if None in (dias_inv, dias_clie, dias_prov) or 0.0 in (dias_inv, dias_clie, dias_prov):
        return None, "Datos insuficientes (días = 0 o N/A)."
    
    cce = dias_inv + dias_clie - dias_prov
    
    if cce < 0:
        sostenibilidad = "Ideal (la empresa cobra antes de pagar el inventario)."
    elif cce <= 60:
        sostenibilidad = "Sostenible (aunque positivo, el ciclo de caja es corto)."
    else:
        sostenibilidad = "Insostenible/Riesgoso (el ciclo de caja es muy largo)."
        
    return cce, sostenibilidad

def generar_analisis_financiero(r):
    AT = r.get("_ActivoTotal") or 0.0
    if AT == 0:
        return "Activo Total es cero. Análisis no es posible.", "N/A"
        
    PC = r.get("_PC") or 0.0
    PNC = r.get("_PNC") or 0.0
    PN = r.get("_PN") or 0.0
    
    Total_Financiacion = PC + PNC + PN
    if Total_Financiacion == 0:
        return "Financiación total es cero.", "N/A"
        
    PC_pct = safe_div(PC, Total_Financiacion) * 100
    PNC_pct = safe_div(PNC, Total_Financiacion) * 100
    PN_pct = safe_div(PN, Total_Financiacion) * 100
    
    comentarios = [
        f"- % Deuda a corto plazo (PC/Total): {fmt_num(PC_pct)}%",
        f"- % Deuda a largo plazo (PNC/Total): {fmt_num(PNC_pct)}%",
        f"- % Recursos propios (PN/Total): {fmt_num(PN_pct)}%"
    ]
    
    if PN_pct >= 50 and PC_pct < 30:
        eq_str = "Estructura **equilibrada y sólida**. Los recursos propios (PN) son la principal fuente de financiación, ideal para software."
    elif PC_pct > 50:
        eq_str = "Estructura **desequilibrada** por alta deuda a corto plazo (PC). Esto genera una fuerte dependencia de la liquidez inmediata, riesgosa."
    else:
        eq_str = "Estructura **aceptable**. Se debe buscar aumentar la participación del Patrimonio Neto y reducir la deuda a corto plazo."
        
    return "\n".join(comentarios), eq_str

def generar_estres_financiero(r24, pct_caida_ingreso=0.30):
    V24 = r24.get("_Ventas") or 0.0
    CV24 = r24.get("_Costo") or 0.0
    BN24 = r24.get("_BN") or 0.0
    AC24 = r24.get("_AC") or 0.0
    PC24 = r24.get("_PC") or 0.0
    
    # 1. Escenario Pesimista 2025 (Caída del 30% en Ventas)
    V25 = V24 * (1 - pct_caida_ingreso)
    
    # Suponemos Costos Variables (CVR) = Costo Ventas.
    CVR_V_pct = safe_div(CV24, V24) # % Costo Variable sobre Ventas
    CV25 = V25 * CVR_V_pct
    
    # ESTIMACIÓN GF (B5.c): Fijamos GF = 30% de las ventas de 2024 (proxy de gastos de administración/venta no cubiertos por Costo Ventas)
    GF24 = V24 * 0.30 
    GF25 = GF24 
    
    # B5.a: Impacto en FM
    BN25 = V25 - CV25 - GF25 # Nuevo BAII/BN estimado
    Impacto_Caja = (BN25 - BN24)
    AC25 = AC24 + Impacto_Caja # Si BN baja, la caja (parte de AC) baja.
    FM25 = AC25 - PC24

    # B5.b: Razón de Liquidez
    Liquidez25 = safe_div(AC25, PC24)

    # B5.c: Punto de Quiebra (PQ)
    MC_pct = 1 - CVR_V_pct
    PQ = safe_div(GF24, MC_pct)
    
    return {
        "FM_Impacto": FM25,
        "Liquidez_Impacto": Liquidez25,
        "BN_Impacto": BN25,
        "PQ_Ventas": PQ,
        "V24": V24,
        "V25": V25
    }

def generar_analisis_apalancamiento(r24):
    RAT = r24.get("RAT")
    i = r24.get("Costo Deuda (i)")
    RRP_apalancada = r24.get("RRP Apalancada")
    efecto_apal = r24.get("Efecto Apalancamiento")
    DeudaTotal = r24.get("_DeudaTotal")
    PN = r24.get("_PN")

    # a) Costo promedio de deuda (i)
    if DeudaTotal != 0:
        i_str = f"i (Costo Deuda) = Gastos Fin / Deuda Total = {fmt_num(i*100)}%"
    else:
        i_str = "Deuda Total es cero. i = 0%."
        
    # b) Comparación RAT vs i
    if RAT is None or i is None:
        comp_str = "No es posible comparar RAT vs i (Faltan datos)."
        apal_efecto_str = "N/A"
        recom_deuda = "No se puede recomendar aumentar deuda sin datos de RAT e i."
    else:
        if RAT > i:
            comp_str = f"RAT ({fmt_num(RAT*100)}%) es **mayor** que i ({fmt_num(i*100)}%)."
            apal_efecto_str = "El apalancamiento es **POSITIVO**."
            recom_deuda = "Sí, convendría aumentar deuda (mientras RAT > i) ya que el capital ajeno genera un retorno superior a su costo."
        else:
            comp_str = f"RAT ({fmt_num(RAT*100)}%) es **menor** o igual que i ({fmt_num(i*100)}%)."
            apal_efecto_str = "El apalancamiento es **NEGATIVO** o nulo."
            recom_deuda = "No, no convendría aumentar deuda. El costo de la deuda supera el retorno económico de la empresa (RAT)."

    # c) Efecto Apalancamiento
    if RRP_apalancada is not None:
        RRP_norm = r24.get("RRP")
        rrp_norm_val = RRP_norm or 0.0
        rrp_apal_val = RRP_apalancada or 0.0
        efecto_apal_str = f"RRP Apalancada: {fmt_num(rrp_apal_val*100)}%. (RRP Normal: {fmt_num(rrp_norm_val*100)}%)"
        efecto_apal_detail = f"Efecto puro del apalancamiento: {fmt_num(efecto_apal)}."
    else:
        efecto_apal_str = "No se pudo calcular RRP Apalancada (PN o Deuda=0)."
        efecto_apal_detail = "N/A"

    return {
        "a": i_str,
        "b": comp_str,
        "apal_efecto_str": apal_efecto_str,
        "c_rrp": efecto_apal_str,
        "c_detail": efecto_apal_detail,
        "d": recom_deuda
    }

def generar_fortalezas_debilidades(r23, r24):
    fz = []
    db = []
    # ventas growth
    v23 = r23.get("_Ventas") or 0.0; v24 = r24.get("_Ventas") or 0.0
    if v23 and v24 and v23 != 0:
        pct = (v24 - v23) / abs(v23) * 100
        if pct >= 0: fz.append(f"Crecimiento de ventas: {fmt_num(pct)}% entre 2023-2024.")
        else: db.append(f"Caída de ventas: {fmt_num(pct)}% entre 2023-2024.")
    # patrimonio/activo
    PN = r24.get("_PN") or 0.0; AT = r24.get("_ActivoTotal") or 0.0
    if PN and AT:
        ratio = safe_div(PN, AT) * 100
        if ratio and ratio >= 50: fz.append(f"Patrimonio aporta {fmt_num(ratio)}% del activo (sólido).")
        else: db.append(f"Bajo aporte patrimonial: PN/Activo = {fmt_num(ratio)}%.")
    # margen
    BAII = r24.get("_BAII") or 0.0
    if BAII and v24:
        margen = safe_div(BAII, v24) * 100
        if margen and margen >= 10: fz.append(f"Margen operativo sano: {fmt_num(margen)}%.")
        else: db.append(f"Margen operativo contenido: {fmt_num(margen)}%.")
    # liquidez
    liq = r24.get("Liquidez General")
    if liq is not None:
        if liq >= 1.5: fz.append(f"Liquidez general adecuada: {fmt_num(liq)}.")
        else: db.append(f"Liquidez general baja: {fmt_num(liq)} (óptimo ~1.5-2).")
    # dias clientes
    if v24 and r24.get("_Deudores") is not None:
        dias = safe_div(r24.get("_Deudores") * 365, v24)
        if dias and dias > 60: db.append(f"Ciclo de cobro largo: ~{int(dias)} días.")
        else: fz.append(f"Ciclo de cobro razonable: ~{int(dias)} días.")
    if not db: db.append("Dependencia de contratos/servicios (riesgo sectorial).")
    return fz[:6], db[:6]

def generar_diagnostico(r23, r24):
    lines = []
    lines.append(f"Diagnóstico ejecutivo — {datetime.date.today().isoformat()}")
    v23 = r23.get("_Ventas") or 0.0; v24 = r24.get("_Ventas") or 0.0
    if v23 and v24 and v23 != 0:
        pct = (v24 - v23) / abs(v23) * 100; lines.append(f"- Crecimiento de ventas: {fmt_num(pct)}%")
    else: lines.append("- Datos de ventas insuficientes para comparar crecimiento.")
    fm23 = r23.get("Fondo Maniobra"); fm24 = r24.get("Fondo Maniobra")
    if fm23 is not None and fm24 is not None:
        lines.append(f"- Fondo de Maniobra: 2023 = {fmt_num(fm23)}, 2024 = {fmt_num(fm24)}.")
        if fm24 > fm23: lines.append("  → Mejora del FM respecto a 2023.")
        elif fm24 < fm23: lines.append("  → FM empeora respecto a 2023; vigilar liquidez.")
    rat = r24.get("RAT"); rrp = r24.get("RRP")
    if rat is not None: lines.append(f"- Rentabilidad económica (RAT) 2024: {fmt_num((rat or 0.0)*100)}%")
    if rrp is not None: lines.append(f"- Rentabilidad financiera (RRP) 2024: {fmt_num((rrp or 0.0)*100)}%")
    cal = r24.get("Calidad Deuda")
    if cal is not None and cal > 0.6: lines.append(f"- Riesgo: alta deuda a corto plazo (PC/Pasivo = {fmt_num(cal)}).")
    lines.append(""); lines.append("Conclusión: La empresa muestra resultados que deben complementarse con mejoras en capital de trabajo y gestión de pasivos para sostener el crecimiento.")
    return "\n".join(lines)

def generar_recomendaciones(r23, r24):
    """Genera 3 recomendaciones cuantificadas para D4: Liquidez, Rentabilidad, Eficiencia."""
    texto = {}
    v = r24.get("_Ventas") or 0.0
    
    # a) Recomendación de Liquidez: Refinanciar Pasivo Corriente (PC)
    PC = r24.get("_PC") or 0.0
    monto = PC * 0.30  # Recomendar refinanciar el 30% del PC
    Liquidez = r24.get("Liquidez General") or 0.0
    
    texto["a) Liquidez"] = f"Refinanciar **30% del Pasivo Corriente (PC)** a largo plazo. Esto reduciría la presión de pago a corto plazo y mejoraría la Razón de Liquidez actual ({fmt_num(Liquidez)}) y el Fondo de Maniobra."
    texto["a_cuantif"] = f"Monto estimado a refinanciar: **{fmt_num(monto)}**."

    # b) Recomendación de Rentabilidad: Reducir Gastos Fijos (Impacto en BAII)
    BAII = r24.get("_BAII") or 0.0
    impacto = BAII * 0.10 if BAII else None # Reducción del 10% en BAII
    RAT = r24.get("RAT") or 0.0
    
    texto["b) Rentabilidad"] = f"Implementar un plan de eficiencia para **reducir gastos fijos y/o operativos en un 10%**. Esto mejorará directamente el Beneficio Antes de Intereses e Impuestos (BAII) y la Rentabilidad Económica (RAT)."
    texto["b_cuantif"] = f"Impacto estimado en el BAII: **{fmt_num(impacto)}**."

    # c) Recomendación de Eficiencia Operativa: Reducir Días Clientes
    dias_actuales = r24.get("_dias_clientes")
    if not dias_actuales:
        dias_actuales = safe_div(r24.get("_Deudores") * 365, v) if v and r24.get("_Deudores") is not None else None
    
    dias_actuales = int(dias_actuales) if dias_actuales else 0
    dias_nuevo = (dias_actuales - 15) if dias_actuales > 15 else 45 # Meta de reducción de 15 días
    mejora_cash = safe_div((dias_actuales - dias_nuevo) * v, 365) if dias_actuales and v and dias_actuales > dias_nuevo else None

    texto["c) Eficiencia operativa"] = f"Mejorar la gestión de cobros para **reducir los Días Clientes** de los actuales {dias_actuales} días a una meta de **{dias_nuevo} días**."
    texto["c_cuantif"] = f"Mejora estimada de flujo de caja: **{fmt_num(mejora_cash)}**."
    
    return texto

# -----------------------------
# Funciones de Soporte Gráfico
# -----------------------------

def generar_pie_chart_financiacion(r):
    """Genera un gráfico de pastel para la estructura financiera (PN, PNC, PC)"""
    PC = r.get("_PC") or 0.0
    PNC = r.get("_PNC") or 0.0
    PN = r.get("_PN") or 0.0
    Total_Financiacion = PC + PNC + PN
    
    if Total_Financiacion == 0:
        return None

    labels = ['Patrimonio Neto (PN)', 'Pasivo No C. (PNC)', 'Pasivo C. (PC)']
    sizes = [PN, PNC, PC]
    colors_list = ['#4CAF50', '#FFC107', '#E91E63'] # Verde, Amarillo, Rosa
    
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors_list, wedgeprops={'edgecolor': 'black'})
    ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    ax.set_title('Estructura de Financiación 2024', fontsize=14)
    
    buf = BytesIO()
    plt.savefig(buf, format='PNG', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return ImageReader(buf)

def generar_draw_rat_rrp(r):
    """Genera un gráfico de barras comparando RAT y RRP."""
    RAT = r.get("RAT") or 0.0
    RRP = r.get("RRP") or 0.0
    RAT_apal = r.get("RRP Apalancada") or 0.0

    labels = ['RAT (Económica)', 'RRP (Normal)', 'RRP (Apalancada)']
    # Convertir a porcentajes para la gráfica
    values = [RAT * 100, RRP * 100, RAT_apal * 100]
    colors_list = [HEX_PRINCIPAL, HEX_ACENTO, '#E91E63'] 

    fig, ax = plt.subplots(figsize=(8, 3.5))
    x = np.arange(len(labels))
    ax.bar(x - 0.18, values, width=0.35, color=colors_list)

    # Añadir valores en las barras
    for i, v in enumerate(values):
        ax.text(x[i], v + 0.5, f"{fmt_num(v)}%", ha='center', va='bottom', fontsize=9)

    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_ylabel('Rentabilidad (%)')
    ax.set_title('C4. Comparación de Rentabilidades 2024', fontsize=12)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()

    buf = BytesIO()
    plt.savefig(buf, format='PNG', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return ImageReader(buf)


def draw_section_box(c, x, y_start, content_elements, width, box_color=COLOR_CAJA_OBJ):
    """Dibuja un contenido dentro de una caja de color."""
    p_style = ParagraphStyle(name="BoxStyle", fontSize=9, leading=12, fontName='Helvetica', alignment=TA_LEFT)
    story = []
    
    # Pre-calcular la altura del contenido
    current_y = 0
    for element in content_elements:
        if isinstance(element, str):
            p = Paragraph(element, p_style)
            w, h = p.wrapOn(c, width - 10, 1000)
            story.append(p)
            story.append(Spacer(1, 5))
            current_y += h + 5
        elif isinstance(element, Paragraph):
            w, h = element.wrapOn(c, width - 10, 1000)
            story.append(element)
            story.append(Spacer(1, 5))
            current_y += h + 5
        elif isinstance(element, Table):
            w, h = element.wrapOn(c, width - 10, 1000)
            story.append(element)
            story.append(Spacer(1, 5))
            current_y += h + 5

    box_height = current_y + 10 # 5 arriba + 5 abajo

    # Dibujar la caja
    c.setFillColor(box_color)
    c.rect(x, y_start - box_height, width, box_height, fill=1)
    
    # Dibujar el contenido dentro de la caja
    c.setFillColor(colors.black)
    text_y = y_start - 5 # 5 de margen superior
    
    for element in story:
        if isinstance(element, Paragraph):
            w, h = element.wrapOn(c, width - 10, 1000)
            text_y -= h
            element.drawOn(c, x + 5, text_y)
            text_y -= 5 # Espacio
        elif isinstance(element, Spacer):
            text_y -= 5
        elif isinstance(element, Table):
            w, h = element.wrapOn(c, width - 10, 1000)
            text_y -= h
            element.drawOn(c, x + 5, text_y)
            text_y -= 5
            
    return y_start - box_height - 10 # Nueva posición Y (debajo de la caja + margen)

def generar_table_style(num_filas=None):
    """
    Genera un objeto TableStyle estético para tablas financieras en ReportLab.
    Aplica colores de la plantilla, bordes minimalistas y formato cebra.
    """
    style = TableStyle([
        # Bordes generales de la tabla
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('BOX', (0, 0), (-1, -1), 1, colors.lightgrey),

        # Estilo para la fila de encabezado (asume que la primera fila es el encabezado)
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_PRINCIPAL), # Fondo Azul Marino
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),   # Texto Blanco
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ])

    # Aplicar formato cebra a las filas de datos
    if num_filas and num_filas > 1:
        for i in range(1, num_filas):
            if i % 2 == 1: # Filas impares (empezando por la segunda fila, índice 1)
                style.add('BACKGROUND', (0, i), (-1, i), COLOR_FONDO_TABLA_OBJ) # Fondo Gris Claro

    # Alineación para las columnas de datos (derecha para números, izquierda para texto)
    style.add('ALIGN', (1, 1), (-1, -1), 'RIGHT')
    style.add('ALIGN', (0, 1), (0, -1), 'LEFT') # Primera columna (descripción) a la izquierda
    style.add('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
    style.add('PADDINGLEFT', (0, 0), (0, -1), 12)
    style.add('PADDINGRIGHT', (-1, 0), (-1, -1), 12)
    style.add('LEFTPADDING', (0, 0), (0, -1), 8)
    style.add('RIGHTPADDING', (0, 0), (-1, -1), 8)
    style.add('BOTTOMPADDING', (0, 1), (-1, -1), 6)
    style.add('TOPPADDING', (0, 1), (-1, -1), 6)

    return style
# -----------------------------
# PDF: generar informe completo
# -----------------------------
def generar_pdf_final(r23, r24, filename="Informe_Financiero_Elegante.pdf"):
    width, height = A4
    c = canvas.Canvas(filename, pagesize=A4)
    x_margin = 2*cm
    y = height - 2*cm
    content_width = width - 4*cm
    
    # ESTILOS
    styles = getSampleStyleSheet()
    estilo_titulo_principal = styles['Heading1']
    estilo_titulo_principal.fontSize = 20
    estilo_titulo_principal.textColor = COLOR_PRINCIPAL
    estilo_titulo_principal.alignment = TA_CENTER
    
    estilo_titulo_seccion = styles['Heading2']
    estilo_titulo_seccion.fontSize = 14
    estilo_titulo_seccion.textColor = COLOR_PRINCIPAL
    estilo_titulo_seccion.alignment = TA_LEFT
    
    estilo_titulo = ParagraphStyle(name="titulo", fontSize=11, leading=14, fontName='Helvetica-Bold', textColor=COLOR_PRINCIPAL)
    estilo_contenido = ParagraphStyle(name="contenido", fontSize=9, leading=12, fontName='Helvetica', alignment=TA_LEFT)
    # Usamos HEX_CAJA (string) en backColor para evitar el error de len()
    estilo_key_result = ParagraphStyle(name="key_result", fontSize=10, leading=14, fontName='Helvetica-Bold', alignment=TA_CENTER, backColor=HEX_CAJA) 
    
    # -----------------------------
    # PAGE 1: Portada y Cuestionario A1-A5
    # -----------------------------
    
    # CABECERA PROFESIONAL
    c.setFillColor(COLOR_PRINCIPAL)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(x_margin, y, "INFORME FINANCIERO Y ECONÓMICO")
    y -= 10
    c.setFont("Helvetica", 12)
    c.drawString(x_margin, y, "Análisis Comparativo 2023 - 2024")
    y -= 15
    c.setStrokeColor(COLOR_ACENTO)
    c.setLineWidth(2)
    c.line(x_margin, y, width - x_margin, y)
    y -= 25

    # SECCIÓN A
    p = Paragraph("SECCIÓN A: ANÁLISIS PATRIMONIAL", estilo_titulo_seccion); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 15
    
    # A1. Fondo de Maniobra (DENTRO DE CAJA)
    fm23 = r23.get("Fondo Maniobra"); fm24 = r24.get("Fondo Maniobra")
    equilibrio24, justif_eq = clasificar_situacion_patrimonial_v2(r24)
    
    p_fm = Paragraph(f"<font size=11><b>A1. Fondo de Maniobra (FM) y Equilibrio Patrimonial</b></font>", estilo_contenido)
    
    # Usamos HEX_ACENTO (string) en <font color>
    p_content_fm = Paragraph(f"<b>FM 2023:</b> Bs. {fmt_num(fm23)} | <b>FM 2024:</b> Bs. {fmt_num(fm24)}.<br/><b>Evolución:</b> {'Mejora' if fm24 > fm23 else ('Empeora' if fm24 < fm23 else 'Se mantiene')}.<br/><b>Equilibrio Patrimonial (2024):</b> <font color='{HEX_ACENTO}'><b>{equilibrio24}</b></font>.", estilo_contenido)
    
    y = draw_section_box(c, x_margin, y, [p_fm, p_content_fm], content_width, box_color=COLOR_CAJA_OBJ)

    # A2. Análisis Vertical + Gráfico de Pastel 
    av24, econo_str, finan_str = generar_analisis_vertical(r24)
    
    p = Paragraph("<font size=11><b>A2. Análisis Vertical del Balance 2024</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    
    # Dibujar Pie Chart
    pie_chart = generar_pie_chart_financiacion(r24)
    if pie_chart:
        img_x = x_margin + content_width - 7.5*cm 
        img_y = y - 7*cm
        c.drawImage(pie_chart, img_x, img_y, width=7*cm, height=7*cm)

        # Mover la explicación a la izquierda del gráfico
        y_text_start = y - 5
        text_width = content_width - 8*cm 
        
        p_av = Paragraph(f"**Distribución (% del Activo Total):**<br/>{av24.replace('\n', '<br/>')}", estilo_contenido)
        p_av_econo = Paragraph(f"<b>Estructura Económica (Activo):</b> {econo_str}", estilo_contenido)
        p_av_finan = Paragraph(f"<b>Estructura Financiera (Pasivo + PN):</b> {finan_str}", estilo_contenido)
        
        y_temp = y_text_start
        for item in [p_av, p_av_econo, p_av_finan]:
            w, h = item.wrapOn(c, text_width, y_temp); y_temp -= h + 5
            item.drawOn(c, x_margin, y_temp)

        y = img_y - 10 
    
    if y < 4*cm: c.showPage(); y = height - 2*cm
    
    # A3. Análisis Horizontal
    ah = generar_analisis_horizontal(r23, r24)
    p = Paragraph("<font size=11><b>A3. Análisis Horizontal del Balance</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    
    # Usamos HEX_ACENTO (string) en <font color>
    p_ah = Paragraph(f"<b>Crecimiento Activo Total:</b> <font color='{HEX_ACENTO}'><b>{fmt_num(ah['Crecimiento Total Activo'])}%</b></font>.<br/>Activo que más creció: **{ah['Activo Mas Crecido']}**.<br/><b>Financiamiento:</b> {ah['Financiacion Principal']}", estilo_contenido)
    
    w, h = p_ah.wrapOn(c, content_width, y); p_ah.drawOn(c, x_margin, y - h); y -= h + 10

    # A4. Ciclo de Conversión de Efectivo (DENTRO DE CAJA)
    d_inv = r24.get("_dias_inventario"); d_clie = r24.get("_dias_clientes"); d_prov = r24.get("_dias_proveedores")
    cce_24, sosten_24 = calcular_cce(d_inv, d_clie, d_prov)
    
    p_cce_title = Paragraph(f"<font size=11><b>A4. Ciclo de Conversión de Efectivo (CCE) 2024</b></font>", estilo_contenido)
    # Usamos HEX_ACENTO (string) en <font color>
    p_cce_content = Paragraph(f"CCE = Días Inventario ({fmt_num(d_inv,0)}) + Días Clientes ({fmt_num(d_clie,0)}) - Días Proveedores ({fmt_num(d_prov,0)})<br/>CCE 2024: <font color='{HEX_ACENTO}'><b>{fmt_num(cce_24, 0)} días</b></font>.<br/><b>Sostenibilidad:</b> **{sosten_24}**", estilo_contenido)
    
    y = draw_section_box(c, x_margin, y, [p_cce_title, p_cce_content], content_width, box_color=COLOR_CAJA_OBJ)

    # A5. Diagnóstico Patrimonial
    p = Paragraph("<font size=11><b>A5. Diagnóstico Patrimonial</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    
    # Usamos HEX_PRINCIPAL (string) en <font color>
    p_diag = Paragraph(f"Estado patrimonial: <font color='{HEX_PRINCIPAL}'><b>{equilibrio24}</b></font>.<br/>**Justificación numérica:** {justif_eq}", estilo_contenido)
    
    w, h = p_diag.wrapOn(c, content_width, y); p_diag.drawOn(c, x_margin, y - h); y -= h + 15
    
    # -----------------------------
    # PAGE 2: Cuestionario B y C
    # -----------------------------
    
    # -----------------------------
    # PAGE 2: Liquidez, Solvencia, Rentabilidad y Apalancamiento
    # -----------------------------
    c.showPage(); y = height - 2*cm
    
    p = Paragraph("SECCIÓN B: ANÁLISIS DE RATIOS FINANCIEROS CLAVE", estilo_titulo_seccion); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 15
    
    # B1. Ratios Liquidez (Liquidez, Tesorería, Disponibilidad)
    p = Paragraph("<font size=11><b>B1. Ratios de Liquidez (2024)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    
    table_data_liq = [
        ["Ratio", "Fórmula", "Resultado (2024)", "Interpretación"],
        ["Liquidez General", "AC / PC", fmt_num(r24.get("Liquidez General")), f"Nivel {'Alto' if r24.get('Liquidez General') > 1.5 else ('Bajo' if r24.get('Liquidez General') < 1.0 else 'Aceptable')}"],
        ["Tesorería", "(Caja+Deudores) / PC", fmt_num(r24.get("Tesorería")), f"Capacidad de pago inmediata sin Inventario"],
        ["Disponibilidad", "Caja / PC", fmt_num(r24.get("Disponibilidad")), f"Capacidad de pago con efectivo"],
    ]
    t_data_liq = [[Paragraph(str(k[0]), estilo_contenido), Paragraph(str(k[1]), estilo_contenido), Paragraph(str(k[2]), estilo_contenido), Paragraph(str(k[3]), estilo_contenido)] for k in table_data_liq]
    t_liq = Table(t_data_liq, colWidths=[3.5*cm, 4.5*cm, 3*cm, 3.5*cm])
    t_liq.setStyle(generar_table_style())
    w, t_h = t_liq.wrapOn(c, content_width, y); t_liq.drawOn(c, x_margin, y - t_h); y -= t_h + 10

    # B2. Ratios Solvencia (Garantía, Autonomía, Calidad Deuda)
    p = Paragraph("<font size=11><b>B2. Ratios de Solvencia y Estructura (2024)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5

    table_data_sol = [
        ["Ratio", "Fórmula", "Resultado (2024)", "Interpretación"],
        ["Garantía", "Activo / Pasivo", fmt_num(r24.get("Garantía")), f"Solvencia: El Activo cubre el Pasivo {fmt_num(r24.get('Garantía'), 1)} veces"],
        ["Autonomía", "PN / Pasivo", fmt_num(r24.get("Autonomía")), f"Autofinanciación: Proporción de Recursos Propios"],
        ["Calidad Deuda", "PC / Pasivo", fmt_num(r24.get("Calidad Deuda")), f"Corto Plazo sobre Deuda Total"],
    ]
    t_data_sol = [[Paragraph(str(k[0]), estilo_contenido), Paragraph(str(k[1]), estilo_contenido), Paragraph(str(k[2]), estilo_contenido), Paragraph(str(k[3]), estilo_contenido)] for k in table_data_sol]
    t_sol = Table(t_data_sol, colWidths=[3.5*cm, 4.5*cm, 3*cm, 3.5*cm])
    t_sol.setStyle(generar_table_style())
    w, t_h = t_sol.wrapOn(c, content_width, y); t_sol.drawOn(c, x_margin, y - t_h); y -= t_h + 10
    
    if y < 4*cm: c.showPage(); y = height - 2*cm

    # B3. Ratios Rentabilidad y Apalancamiento (Gráfico y Texto)
    p = Paragraph("<font size=11><b>B3. Análisis de Rentabilidad y Apalancamiento (2024)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    
    img_rat_rrp = generar_draw_rat_rrp(r24)
    c.drawImage(img_rat_rrp, x_margin + 0.5*cm, y - 5*cm, width=7*cm, height=4*cm, mask='auto')
    
    # Análisis de Apalancamiento (Texto)
    apal_res = generar_analisis_apalancamiento(r24)
    p_apal_title = Paragraph(f"<font size=11><b>Efecto Apalancamiento Financiero</b></font>", estilo_contenido)
    p_apal_content = Paragraph(f"<b>Costo Deuda (i):</b> {apal_res['a']}<br/><b>Comparación:</b> {apal_res['b']}<br/><font color='{HEX_ACENTO}'><b>Conclusión:</b> {apal_res['apal_efecto_str']} | **RRP Apalancada:** {apal_res['c_rrp']}</font><br/><b>Recomendación de Deuda:</b> {apal_res['d']}", estilo_contenido)
    
    y = draw_section_box(c, x_margin + 7.5*cm, y, [p_apal_title, p_apal_content], content_width - 7.5*cm, box_color=COLOR_CAJA_OBJ)

    y -= 5*cm # Compensar el espacio del gráfico


    # B4. Análisis de Estructura Financiera 
    comentarios, eq_str = generar_analisis_financiero(r24)
    p = Paragraph("<font size=11><b>B4. Análisis de Estructura Financiera (2024)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    p_b4 = Paragraph(f"{comentarios.replace('\n', '<br/>')}<br/><b>¿Estructura equilibrada para software?:</b> **{eq_str}**", estilo_contenido)
    w, h = p_b4.wrapOn(c, content_width, y); p_b4.drawOn(c, x_margin, y - h); y -= h + 10

    # B5. Estrés Financiero - Escenario Pesimista
    if y < 6*cm: 
        c.showPage()
        y = height - 2*cm
    #
    estres = generar_estres_financiero(r24)
    p = Paragraph("<font size=11><b>B5. Estrés Financiero - Escenario Pesimista (Ventas -30% en 2025)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    
    # Usamos HEX_PRINCIPAL (string) en <font color>
    p_b5 = Paragraph(f"Ingresos 2024: {fmt_num(estres['V24'])} | Ingresos 2025 (proyectado): {fmt_num(estres['V25'])}<br/>"
                     f"a) FM (proyectado): <font color='{HEX_PRINCIPAL}'><b>{fmt_num(estres['FM_Impacto'])}</b></font><br/>"
                     f"b) Razón de Liquidez General (proyectada): **{fmt_num(estres['Liquidez_Impacto'])}**<br/>"
                     f"c) Punto de Quiebra (Ventas mínimas): **{fmt_num(estres['PQ_Ventas'])}**", estilo_contenido)
    
    w, h = p_b5.wrapOn(c, content_width, y); p_b5.drawOn(c, x_margin, y - h); y -= h + 15

    # SECCIÓN C (Completa)
    p = Paragraph("SECCIÓN C: ANÁLISIS DE RATIOS FINANCIEROS Y APALANCAMIENTO", estilo_titulo_seccion); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 15
    
    # C1. Liquidez
    p = Paragraph("<font size=11><b>C1. Ratios de Liquidez (Corto Plazo)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    # Usamos HEX_ACENTO (string) en <font color>
    p_c1 = Paragraph(f"<b>Liquidez General (AC/PC):</b> <font color='{HEX_ACENTO}'><b>{fmt_num(r24.get('Liquidez General'))}</b></font> (óptimo 1.5-2.0)<br/>"
                     f"<b>Razón de Tesorería (C+D/PC):</b> {fmt_num(r24.get('Tesorería'))} (óptimo ~1.0)<br/>"
                     f"<b>Disponibilidad (C/PC):</b> {fmt_num(r24.get('Disponibilidad'))} (óptimo 0.2-0.3)", estilo_contenido)
    w, h = p_c1.wrapOn(c, content_width, y); p_c1.drawOn(c, x_margin, y - h); y -= h + 10
    
    # C2. Solvencia y Endeudamiento
    if y < 4*cm: c.showPage(); y = height - 2*cm
    p = Paragraph("<font size=11><b>C2. Ratios de Solvencia y Endeudamiento (Largo Plazo)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    # Usamos HEX_ACENTO (string) en <font color>
    p_c2 = Paragraph(f"<b>Garantía (Activo/Pasivo):</b> <font color='{HEX_ACENTO}'><b>{fmt_num(r24.get('Garantía'))}</b></font> (óptimo > 1.5)<br/>"
                     f"<b>Autonomía (PN/Pasivo):</b> {fmt_num(r24.get('Autonomía'))} (óptimo > 1.0)<br/>"
                     f"<b>Calidad de la Deuda (PC/Pasivo):</b> {fmt_num(r24.get('Calidad Deuda'))} (vigilancia si es > 0.6)", estilo_contenido)
    w, h = p_c2.wrapOn(c, content_width, y); p_c2.drawOn(c, x_margin, y - h); y -= h + 10

    # C3. Rentabilidad (RAT vs RRP)
    p = Paragraph("<font size=11><b>C3. Ratios de Rentabilidad (RAT vs RRP)</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    RAT_val = (r24.get('RAT') or 0.0) * 100; RRP_val = (r24.get('RRP') or 0.0) * 100
    # Usamos HEX_PRINCIPAL y HEX_ACENTO (strings) en <font color>
    p_c3 = Paragraph(f"<b>Rentabilidad Económica (RAT):</b> <font color='{HEX_PRINCIPAL}'><b>{fmt_num(RAT_val)}%</b></font><br/>"
                     f"<b>Rentabilidad Financiera (RRP):</b> <font color='{HEX_ACENTO}'><b>{fmt_num(RRP_val)}%</b></font><br/>"
                     f"**Análisis:** {'El rendimiento para los accionistas (RRP) es superior al rendimiento de los activos (RAT).' if RRP_val > RAT_val else 'El rendimiento de los activos (RAT) es superior o igual al RRP.'}", estilo_contenido)
    w, h = p_c3.wrapOn(c, content_width, y); p_c3.drawOn(c, x_margin, y - h); y -= h + 10
    
    # C4. Gráfico RAT vs RRP 
    rat_rrp_chart = generar_draw_rat_rrp(r24)
    if rat_rrp_chart:
        c.drawImage(rat_rrp_chart, x_margin, y - 4*cm, width=content_width, height=4*cm)
        y -= 4.5*cm

    # C5. Apalancamiento Financiero
    if y < 6*cm: c.showPage(); y = height - 2*cm
    apal = generar_analisis_apalancamiento(r24)
    p = Paragraph("<font size=11><b>C5. Apalancamiento Financiero</b></font>", estilo_contenido); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 5
    
    c5_color_choice = HEX_ACENTO if apal['apal_efecto_str'].find('POSITIVO')!=-1 else HEX_PRINCIPAL
    
    p_c5 = Paragraph(
        f"a) Costo promedio de deuda (i): **{apal['a']}**<br/>"
        f"b) Comparación RAT vs i: {apal['b']}<br/>"
        f"   -> Efecto apalancamiento: <font color='{c5_color_choice}'><b>{apal['apal_efecto_str']}</b></font><br/>"
        f"c) RRP Apalancada: {apal['c_rrp']}<br/>"
        f"d) ¿Convendría aumentar deuda?: **{apal['d']}**", estilo_contenido)
    w, h = p_c5.wrapOn(c, content_width, y); p_c5.drawOn(c, x_margin, y - h); y -= h + 15

    
    # -----------------------------
    # PAGE 3: Cuestionario D
    # -----------------------------
    c.showPage(); y = height - 2*cm
    
    # SECCIÓN D
    p = Paragraph("SECCIÓN D: ANÁLISIS DE RATIOS Y DIAGNÓSTICO INTEGRAL", estilo_titulo_seccion); w, h = p.wrapOn(c, content_width, y); p.drawOn(c, x_margin, y - h); y -= h + 15

    # D1 Matriz de ratios comparativa 
    c.setFont("Helvetica-Bold", 10); c.drawString(x_margin, y, "D1. Matriz de Ratios Comparativos 2023 vs 2024"); y -= 12
    keys = ["Fondo Maniobra", "Liquidez General", "Tesorería", "Disponibilidad", "Garantía", "Autonomía", "Calidad Deuda", "RAT", "RRP"]
    header = [Paragraph("Ratio", estilo_contenido), Paragraph("2023", estilo_contenido), Paragraph("2024", estilo_contenido), Paragraph("Cambio (abs)", estilo_contenido), Paragraph("Cambio (%)", estilo_contenido)]
    table_data = [header]
    for k in keys:
        v23 = r23.get(k); v24 = r24.get(k); abs_ch = (v24 - v23) if (v23 is not None and v24 is not None) else None
        pct_ch = safe_div(abs_ch, abs(v23)) * 100 if (abs_ch is not None and v23 not in (None,0)) else None
        table_data.append([Paragraph(k, estilo_contenido), Paragraph(fmt_num(v23), estilo_contenido), Paragraph(fmt_num(v24), estilo_contenido), Paragraph(fmt_num(abs_ch), estilo_contenido), Paragraph((fmt_num(pct_ch) + "%" if pct_ch is not None else "N/A"), estilo_contenido)])
    colw = [4.5*cm, 2.8*cm, 2.8*cm, 3.2*cm, 3.2*cm]
    t = Table(table_data, colWidths=colw)
    t.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        # Usamos COLOR_FONDO_TABLA_OBJ (Objeto Color) que sí es aceptado aquí
        ('BACKGROUND',(0,0),(-1,0),COLOR_FONDO_TABLA_OBJ), 
        ('FONT',(0,0),(-1,0),'Helvetica-Bold'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('ALIGN',(1,1),(-1,-1),'CENTER'),
    ])); w,t_h = t.wrapOn(c, content_width, y); t.drawOn(c, x_margin, y - t_h); y -= t_h + 10

    # Gráfico comparativo 
    plot_keys = ["Liquidez General", "Tesorería", "Garantía", "Autonomía", "RAT", "RRP"]
    vals23 = [r23.get(k) or 0 for k in plot_keys]; vals24 = [r24.get(k) or 0 for k in plot_keys]
    fig, ax = plt.subplots(figsize=(8.5,3.5)); x = np.arange(len(plot_keys))
    
    # CORRECCIÓN CLAVE: Usamos las cadenas HEX directamente para Matplotlib
    ax.bar(x - 0.18, vals23, width=0.35, label="2023", color=HEX_PRINCIPAL); 
    ax.bar(x + 0.18, vals24, width=0.35, label="2024", color=HEX_ACENTO)
    
    ax.set_xticks(x); ax.set_xticklabels(plot_keys, rotation=30, ha="right"); 
    ax.legend(); ax.grid(axis='y', linestyle='--', alpha=0.7); 
    ax.set_title('Evolución de Ratios Clave (2023 vs 2024)')
    plt.tight_layout()
    buf = BytesIO(); plt.savefig(buf, format='PNG', dpi=150); plt.close(fig); buf.seek(0); img = ImageReader(buf)
    if y - 7*cm < 2*cm: c.showPage(); y = height - 2*cm
    c.drawImage(img, x_margin, y - 7*cm, width=content_width, height=7*cm); y -= 7.5*cm

    # D2 Fortalezas y Debilidades
    if y < 6*cm: c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 10); c.drawString(x_margin, y, "D2. Fortalezas y Debilidades"); y -= 12
    fz, db = generar_fortalezas_debilidades(r23, r24)
    
    # Dividir el espacio para dos columnas (Fortalezas y Debilidades)
    col1_x = x_margin
    col2_x = x_margin + content_width / 2 + 5
    col_width = content_width / 2 - 5
    y_start_d2 = y
    
    # Fortalezas
    c.setFont("Helvetica-Bold", 9); c.setFillColor(COLOR_ACENTO); c.drawString(col1_x, y_start_d2 - 10, "✅ FORTALEZAS"); y_current = y_start_d2 - 20
    c.setFillColor(colors.black); c.setFont("Helvetica", 9)
    for it in fz:
        p_item = Paragraph("• " + it, estilo_contenido); w, h = p_item.wrapOn(c, col_width, y_current); p_item.drawOn(c, col1_x, y_current - h); y_current -= h + 5
    
    # Debilidades
    c.setFont("Helvetica-Bold", 9); c.setFillColor(colors.red); c.drawString(col2_x, y_start_d2 - 10, "❌ DEBILIDADES"); y_current_db = y_start_d2 - 20
    c.setFillColor(colors.black); c.setFont("Helvetica", 9)
    for it in db:
        p_item = Paragraph("• " + it, estilo_contenido); w, h = p_item.wrapOn(c, col_width, y_current_db - h); p_item.drawOn(c, col2_x, y_current_db - h); y_current_db -= h + 5

    y = min(y_current, y_current_db) - 10

    # D3 Diagnóstico ejecutivo
    if y < 6*cm: c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 10); c.drawString(x_margin, y, "D3. Diagnóstico Ejecutivo Integral"); y -= 12
    c.setFont("Helvetica", 9); diag = generar_diagnostico(r23, r24)
    p_diag_final = Paragraph(diag.replace('\n', '<br/>'), estilo_contenido)
    w, h = p_diag_final.wrapOn(c, content_width, y); p_diag_final.drawOn(c, x_margin, y - h); y -= h + 10

    # D4 Recomendaciones (DENTRO DE CAJA)
    if y < 6*cm: c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 10); c.drawString(x_margin, y, "D4. Recomendaciones Estratégicas (3 medidas cuantificadas)"); y -= 12
    c.setFont("Helvetica", 9)
    recs = generar_recomendaciones(r23, r24)
    
    # Usamos HEX_PRINCIPAL (string) en <font color>
    items_recs = [
        f"**1) Liquidez:** {recs['a) Liquidez']}. Cuantificación: <font color='{HEX_PRINCIPAL}'>{recs['a_cuantif']}</font>",
        f"**2) Rentabilidad:** {recs['b) Rentabilidad']}. Cuantificación: <font color='{HEX_PRINCIPAL}'>{recs['b_cuantif']}</font>",
        f"**3) Eficiencia operativa:** {recs['c) Eficiencia operativa']}. Cuantificación: <font color='{HEX_PRINCIPAL}'>{recs['c_cuantif']}</font>"
    ]
    
    p_recs_title = Paragraph(f"<font size=11><b>Prioridades Estratégicas</b></font>", estilo_contenido)
    p_recs_content = Paragraph("<br/>".join(items_recs), estilo_contenido)
    
    y = draw_section_box(c, x_margin, y, [p_recs_title, p_recs_content], content_width, box_color=COLOR_CAJA_OBJ)

    c.save()

# -----------------------------
# INTERFAZ TKINTER (SIN CAMBIOS)
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Informe Financiero - Cuestionario Completo")
        self.geometry("1000x750")
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Ingrese los datos financieros (valores en la misma unidad)", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=4, pady=6)

        # Definición de campos y valores iniciales basados en el Balance
        datos_iniciales = {
            "activo_corriente": {"2023": "2800", "2024": "3800"},
            "activo_no_corriente": {"2023": "1450", "2024": "1850"},
            "pasivo_corriente": {"2023": "550", "2024": "1000"},
            "pasivo_no_corriente": {"2023": "700", "2024": "1000"},
            "patrimonio_neto": {"2023": "3000", "2024": "3650"},
            
            "ventas": {"2023": "1000", "2024": "1500"}, 
            "costo_ventas": {"2023": "400", "2024": "600"}, 
            "beneficio_neto": {"2023": "300", "2024": "500"}, 
            
            "deudores": {"2023": "1200", "2024": "1600"},
            "inventario": {"2023": "300", "2024": "500"},
            "caja": {"2023": "850", "2024": "1100"},
            
            "i": {"2023": "0.05", "2024": "0.05"},
            "gastos_financieros": {"2023": "50", "2024": "60"}, 
            
            "dias_inventario": {"2023": "45", "2024": "45"}, 
            "dias_clientes": {"2023": "60", "2024": "60"}, 
            "dias_proveedores": {"2023": "30", "2024": "30"} 
        }
        
        self.fields = [
            ("activo_corriente","Activo Corriente (AC)"),
            ("activo_no_corriente","Activo No Corriente (ANC)"),
            ("pasivo_corriente","Pasivo Corriente (PC)"),
            ("pasivo_no_corriente","Pasivo No Corriente (PNC)"),
            ("patrimonio_neto","Patrimonio Neto (PN)"),
            ("ventas","Ventas (E)"),
            ("costo_ventas","Costo de ventas (E)"),
            ("beneficio_neto","Beneficio Neto (E)"),
            ("deudores","Clientes / Cuentas por cobrar"),
            ("inventario","Existencias / Inventario"),
            ("caja","Caja / Efectivo"),
            ("i","Tasa interés (i) (E)"),
            ("gastos_financieros", "Gastos Financieros (GF) (E)"),
            ("dias_inventario","Días de Inventario (A4)"),
            ("dias_clientes","Días de Clientes (A4)"),
            ("dias_proveedores","Días de Proveedores (A4)")
        ]
        
        self.entries = {"2023":{}, "2024":{}}
        ttk.Label(frm, text="Concepto").grid(row=1, column=0, sticky="w")
        ttk.Label(frm, text="2023").grid(row=1, column=1)
        ttk.Label(frm, text="2024").grid(row=1, column=2)

        r=2
        for key,label in self.fields:
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=3)
            e1 = ttk.Entry(frm, width=20); e2 = ttk.Entry(frm, width=20)
            e1.grid(row=r, column=1); e2.grid(row=r, column=2)

            # Insertar valores por defecto
            e1.insert(0, datos_iniciales.get(key, {}).get("2023", "0"))
            e2.insert(0, datos_iniciales.get(key, {}).get("2024", "0"))
            
            self.entries["2023"][key] = e1
            self.entries["2024"][key] = e2
            r += 1

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=r, column=0, columnspan=4, pady=12)
        ttk.Button(btn_frame, text="Calcular y mostrar (pantalla)", command=self.mostrar).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="Generar PDF profesional", command=self.export_pdf).grid(row=0, column=1, padx=6)
        ttk.Button(btn_frame, text="Limpiar", command=self.limpiar).grid(row=0, column=2, padx=6)
        ttk.Button(btn_frame, text="Salir", command=self.destroy).grid(row=0, column=3, padx=6)

        self.output = tk.Text(frm, height=18, width=120, font=("Consolas",10))
        self.output.grid(row=r+1, column=0, columnspan=4, pady=8)

    def leer_inputs(self):
        data = {"2023":{}, "2024":{}}
        for yr in ("2023","2024"):
            for key, ent in self.entries[yr].items():
                txt = ent.get().strip()
                if txt == "":
                    val = None
                else:
                    try:
                        val = float(txt)
                    except:
                        val = None
                data[yr][key] = val
        return data

    def mostrar(self):
        data = self.leer_inputs()
        r23 = calcular_ratios_from_inputs(data["2023"])
        r24 = calcular_ratios_from_inputs(data["2024"])
        out = []

        # --- SECCIÓN A: ANÁLISIS PATRIMONIAL (A1-A5) ---
        out.append("="*80); out.append("=== SECCIÓN A: ANÁLISIS PATRIMONIAL (A1-A5) ==="); out.append("="*80)
        
        # A1. Fondo de Maniobra y Equilibrio
        fm23 = r23.get("Fondo Maniobra"); fm24 = r24.get("Fondo Maniobra")
        equilibrio24, justif_eq = clasificar_situacion_patrimonial_v2(r24)
        out.append("\n=== A1. Fondo de Maniobra (FM) y Equilibrio Patrimonial ===")
        out.append(f"FM 2023: {fmt_num(fm23)} | FM 2024: {fmt_num(fm24)}")
        out.append(f"Tipo de equilibrio patrimonial (2024): **{equilibrio24}**")

        # A2. Análisis Vertical 2024
        av24, econo_str, finan_str = generar_analisis_vertical(r24)
        out.append("\n=== A2. Análisis Vertical del Balance 2024 ===")
        out.append("Cada rubro como % del Activo Total:"); out.append(av24)
        out.append(f"Estructura Económica (Activo): {econo_str}"); out.append(f"Estructura Financiera (Pasivo + PN): {finan_str}")
        
        # A3. Análisis Horizontal 2024 vs 2023
        ah = generar_analisis_horizontal(r23, r24)
        out.append("\n=== A3. Análisis Horizontal 2024 vs 2023 ===")
        out.append(f"Activo que más creció: **{ah['Activo Mas Crecido']}**")
        out.append(f"El Activo Total creció un {fmt_num(ah['Crecimiento Total Activo'])}%.")
        out.append(ah['Financiacion Principal'])

        # A4. Ciclo de Conversión de Efectivo (CCE)
        d_inv = r24.get("_dias_inventario"); d_clie = r24.get("_dias_clientes"); d_prov = r24.get("_dias_proveedores")
        cce_24, sosten_24 = calcular_cce(d_inv, d_clie, d_prov)
        out.append("\n=== A4. Ciclo de Conversión de Efectivo (CCE) 2024 ===")
        out.append(f"CCE 2024: **{fmt_num(cce_24, 0)} días**."); out.append(f"¿Es sostenible?: **{sosten_24}**")

        # A5. Diagnóstico Patrimonial
        out.append("\n=== A5. Diagnóstico Patrimonial ===")
        out.append(f"Estado patrimonial (2024): **{equilibrio24}**."); out.append(f"Justificación con números: {justif_eq}")

        # --- SECCIÓN B: ANÁLISIS COMPLEMENTARIO (B4-B5) ---
        out.append("\n" + "="*80); out.append("=== SECCIÓN B: ANÁLISIS COMPLEMENTARIO (B4-B5) ==="); out.append("="*80)
        
        # B4. Análisis de Estructura Financiera
        comentarios, eq_str = generar_analisis_financiero(r24)
        out.append("\n=== B4. Análisis de Estructura Financiera (2024) ===")
        out.append(comentarios); out.append(f"¿Estructura equilibrada para software?: **{eq_str}**")
        
        # B5. Estrés Financiero - Escenario Pesimista
        estres = generar_estres_financiero(r24)
        out.append("\n=== B5. Estrés Financiero - Escenario Pesimista (Ventas -30% en 2025) ===")
        out.append(f"a) FM (proyectado): **{fmt_num(estres['FM_Impacto'])}**")
        out.append(f"b) Razón de Liquidez General (proyectada): **{fmt_num(estres['Liquidez_Impacto'])}**")
        out.append(f"c) Punto de Quiebra (Ventas mínimas para cubrir costos): **{fmt_num(estres['PQ_Ventas'])}**")

        # --- SECCIÓN C: APALANCAMIENTO (C1-C5) ---
        out.append("\n" + "="*80); out.append("=== SECCIÓN C: ANÁLISIS DE RATIOS FINANCIEROS Y APALANCAMIENTO (C1-C5) ==="); out.append("="*80)
        
        # C1. Liquidez
        out.append("\n=== C1. Ratios de Liquidez (Corto Plazo) ===")
        out.append(f"Liquidez General (AC/PC): **{fmt_num(r24.get('Liquidez General'))}** (óptimo 1.5-2.0)")
        out.append(f"Razón de Tesorería (C+D/PC): **{fmt_num(r24.get('Tesorería'))}** (óptimo ~1.0)")
        out.append(f"Disponibilidad (C/PC): **{fmt_num(r24.get('Disponibilidad'))}** (óptimo 0.2-0.3)")

        # C2. Solvencia
        out.append("\n=== C2. Ratios de Solvencia y Endeudamiento (Largo Plazo) ===")
        out.append(f"Garantía (Activo/Pasivo): **{fmt_num(r24.get('Garantía'))}** (óptimo > 1.5)")
        out.append(f"Autonomía (PN/Pasivo): **{fmt_num(r24.get('Autonomía'))}** (óptimo > 1.0)")
        out.append(f"Calidad de la Deuda (PC/Pasivo): **{fmt_num(r24.get('Calidad Deuda'))}** (vigilancia si es > 0.6)")

        # C3. Rentabilidad (RAT vs RRP)
        RAT_val = (r24.get('RAT') or 0.0) * 100
        RRP_val = (r24.get('RRP') or 0.0) * 100
        out.append("\n=== C3. Ratios de Rentabilidad (RAT vs RRP) ===")
        out.append(f"Rentabilidad Económica (RAT): **{fmt_num(RAT_val)}%**")
        out.append(f"Rentabilidad Financiera (RRP): **{fmt_num(RRP_val)}%**")
        out.append(f"Análisis: {'RRP superior a RAT (Efecto Apalancamiento)' if RRP_val > RAT_val else 'RRP inferior o igual a RAT.'}")
        # C4. (El gráfico es solo para PDF)

        # C5. Apalancamiento Financiero
        apal = generar_analisis_apalancamiento(r24)
        out.append("\n=== C5. Apalancamiento Financiero ===")
        out.append(f"a) Costo promedio de deuda (i): **{apal['a']}**")
        out.append(f"b) Comparación RAT vs i: {apal['b']}")
        out.append(f"   -> Conclusión: **{apal['apal_efecto_str']}**") 
        out.append(f"c) Cálculo del RRP Apalancada: {apal['c_rrp']}")
        out.append(f"d) ¿Convendría aumentar deuda?: **{apal['d']}**")

        # --- SECCIÓN D: ANÁLISIS DE RATIOS Y DIAGNÓSTICO EJECUTIVO ---
        out.append("\n" + "="*80); out.append("=== SECCIÓN D: ANÁLISIS DE RATIOS Y DIAGNÓSTICO EJECUTIVO ==="); out.append("="*80)
        out.append("\n=== D1: Matriz de Ratios 2023 vs 2024 ===")
        keys = ["Fondo Maniobra","Liquidez General","Tesorería","Disponibilidad","Garantía","Autonomía","Calidad Deuda","RAT","RRP"]
        for k in keys: out.append(f"{k}: 2023={fmt_num(r23.get(k))} | 2024={fmt_num(r24.get(k))}")
        out.append("\n=== D2: Fortalezas y Debilidades ==="); fz, db = generar_fortalezas_debilidades(r23, r24)
        out.append("Fortalezas:"); [out.append(f" {i+1}. {it}") for i, it in enumerate(fz)]
        out.append("Debilidades:"); [out.append(f" {i+1}. {it}") for i, it in enumerate(db)]
        out.append("\n=== D3: Diagnóstico ejecutivo ==="); out.append(generar_diagnostico(r23, r24))
        
        # D4: Recomendaciones
        recs = generar_recomendaciones(r23, r24)
        out.append("\n=== D4: Recomendaciones Estratégicas (Cuantificadas) ===")
        out.append(f"1. {recs['a) Liquidez']}. Cuantificación: {recs['a_cuantif']}")
        out.append(f"2. {recs['b) Rentabilidad']}. Cuantificación: {recs['b_cuantif']}")
        out.append(f"3. {recs['c) Eficiencia operativa']}. Cuantificación: {recs['c_cuantif']}")
        
        self.output.delete("1.0", tk.END); self.output.insert(tk.END, "\n".join(out))


    def export_pdf(self):
        data = self.leer_inputs()
        r23 = calcular_ratios_from_inputs(data["2023"])
        r24 = calcular_ratios_from_inputs(data["2024"])
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="Informe_Financiero_Elegante.pdf", filetypes=[("PDF files","*.pdf")])
        if not file_path: return
        try:
            generar_pdf_final(r23, r24, filename=file_path)
            messagebox.showinfo("PDF generado", f"Informe guardado en:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error al generar PDF", f"Ocurrió un error: {e}")

    def limpiar(self):
        # Limpiar y re-insertar los valores por defecto del Balance
        datos_iniciales_reset = { 
            "activo_corriente": {"2023": "2800", "2024": "3800"},
            "activo_no_corriente": {"2023": "1450", "2024": "1850"},
            "pasivo_corriente": {"2023": "550", "2024": "1000"},
            "pasivo_no_corriente": {"2023": "700", "2024": "1000"},
            "patrimonio_neto": {"2023": "3000", "2024": "3650"},
            "ventas": {"2023": "1000", "2024": "1500"},
            "costo_ventas": {"2023": "400", "2024": "600"},
            "beneficio_neto": {"2023": "300", "2024": "500"},
            "deudores": {"2023": "1200", "2024": "1600"},
            "inventario": {"2023": "300", "2024": "500"},
            "caja": {"2023": "850", "2024": "1100"},
            "i": {"2023": "0.05", "2024": "0.05"},
            "gastos_financieros": {"2023": "50", "2024": "60"},
            "dias_inventario": {"2023": "45", "2024": "45"},
            "dias_clientes": {"2023": "60", "2024": "60"},
            "dias_proveedores": {"2023": "30", "2024": "30"}
        }
        
        for yr in ("2023","2024"):
            for key, ent in self.entries[yr].items(): 
                ent.delete(0, tk.END)
                ent.insert(0, datos_iniciales_reset.get(key, {}).get(yr, "0"))
                
        self.output.delete("1.0", tk.END)

# -----------------------------
# EJECUTAR
# -----------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()