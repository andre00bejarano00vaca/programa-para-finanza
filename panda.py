# informe_financiero_final.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.utils import ImageReader
import matplotlib.pyplot as plt
from io import BytesIO
import math
import datetime
import numpy as np

# -----------------------------
# UTILIDADES
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

# -----------------------------
# CÁLCULOS (con solo datos mínimos)
# -----------------------------
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
    i = d.get("i") if d.get("i") is not None else 0.05

    Activo = AC + ANC
    Pasivo = PC + PNC

    ratios = {}
    # Fondo de Maniobra
    ratios["Fondo Maniobra"] = AC - PC
    ratios["Fondo Maniobra Alternativo"] = PN + PNC - ANC

    # Liquidez
    ratios["Liquidez General"] = safe_div(AC, PC)
    ratios["Tesorería"] = safe_div((Caja + Deudores), PC)
    ratios["Disponibilidad"] = safe_div(Caja, PC)

    # Solvencia
    ratios["Garantía"] = safe_div(Activo, Pasivo)
    ratios["Autonomía"] = safe_div(PN, Pasivo)
    ratios["Calidad Deuda"] = safe_div(PC, Pasivo)

    # Rentabilidades
    BAII = None
    # If BAII not provided (we only have ventas y costo), approximate BAII = Ventas - Costo (no incluye admin)
    BAII = Ventas - Costo
    ratios["RAT"] = safe_div(BAII, Activo)  # rentabilidad económica
    ratios["RRP"] = safe_div(BN, PN)  # rentabilidad financiera (ROE)

    # Apalancamiento (using formula RRP = RAT + D*(RAT - i)/PN)
    D = Pasivo - PN
    if PN != 0:
        RAT = ratios["RAT"] or 0.0
        ratios["Apalancamiento Financiero"] = RAT + safe_div(D * (RAT - i), PN)
    else:
        ratios["Apalancamiento Financiero"] = None

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
    ratios["_i"] = i
    ratios["_ActivoTotal"] = Activo
    ratios["_PasivoTotal"] = Pasivo

    return ratios

# -----------------------------
# Clasificación situación patrimonial (regla A)
# -----------------------------
def clasificar_situacion_patrimonial(r):
    AC = r.get("_AC") or 0.0
    PC = r.get("_PC") or 0.0
    PN = r.get("_PN") or 0.0
    BN = r.get("_BN") or 0.0
    Activo = r.get("_ActivoTotal") or 0.0
    fm = r.get("Fondo Maniobra")
    # Rule A: PN<0 or BN<0 => desequilibrio LP
    if PN < 0 or BN < 0:
        return "Desequilibrio financiero a L/P"
    if AC < PC:
        return "Suspensión de pagos"
    if Activo and PN/Activo > 0.9:
        return "Estabilidad financiera total"
    if fm is not None and fm >= 0:
        return "Estabilidad financiera normal"
    return "Estabilidad financiera normal"

# -----------------------------
# Tabla situación patrimonial con Paragraph wrapping
# -----------------------------
def crear_tabla_situacion():
    estilo = ParagraphStyle(name="tabla", fontSize=9, leading=11, alignment=TA_LEFT)
    header = [
        Paragraph("Situación patrimonial", estilo),
        Paragraph("Causas", estilo),
        Paragraph("Consecuencias", estilo),
        Paragraph("Solución", estilo),
    ]
    rows = [
        [
            Paragraph("Estabilidad financiera total", estilo),
            Paragraph("Exceso de financiación propia: la empresa se financia sólo con recursos propios", estilo),
            Paragraph("Total seguridad, pero no se puede beneficiar del efecto positivo del endeudamiento", estilo),
            Paragraph("Endeudarse moderadamente para utilizar capitales ajenos y aumentar su estabilidad", estilo),
        ],
        [
            Paragraph("Estabilidad financiera normal", estilo),
            Paragraph("FM positivo", estilo),
            Paragraph("Es la situación ideal para la empresa", estilo),
            Paragraph("Vigilar que el FM sea el necesario para la actividad", estilo),
        ],
        [
            Paragraph("Suspensión de pagos", estilo),
            Paragraph("AC < PC (FM negativo)", estilo),
            Paragraph("No puede pagar sus deudas a corto plazo", estilo),
            Paragraph("Ampliar plazos de pago y planificar tesorería", estilo),
        ],
        [
            Paragraph("Desequilibrio financiero a L/P", estilo),
            Paragraph("FM negativo y/o acumulación de pérdidas (descapitalización)", estilo),
            Paragraph("Descapitalización; solvencia exigua; riesgo de concurso", estilo),
            Paragraph("Reestructuración financiera o medidas de capitalización", estilo),
        ],
    ]
    data = [header] + rows
    col_widths = [4*cm, 5.5*cm, 5.5*cm, 4.5*cm]
    table = Table(data, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#DCEFFD")),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN',(0,0),(-1,-1),'LEFT'),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('GRID',(0,0),(-1,-1),0.4,colors.black),
        ('LEFTPADDING',(0,0),(-1,-1),4),
        ('RIGHTPADDING',(0,0),(-1,-1),4),
        ('TOPPADDING',(0,0),(-1,-1),4),
        ('BOTTOMPADDING',(0,0),(-1,-1),4),
    ]))
    return table

# -----------------------------
# Fortalezas/Debilidades / Diagnóstico / Recomendaciones
# -----------------------------
def generar_fortalezas_debilidades(r23, r24):
    fz = []
    db = []
    # ventas growth
    v23 = r23.get("_Ventas") or 0.0
    v24 = r24.get("_Ventas") or 0.0
    if v23 and v24 and v23 != 0:
        pct = (v24 - v23) / abs(v23) * 100
        if pct >= 0:
            fz.append(f"Crecimiento de ventas: {fmt_num(pct)}% entre 2023-2024.")
        else:
            db.append(f"Caída de ventas: {fmt_num(pct)}% entre 2023-2024.")
    # patrimonio/activo
    PN = r24.get("_PN") or 0.0
    AT = r24.get("_ActivoTotal") or 0.0
    if PN and AT:
        ratio = safe_div(PN, AT) * 100
        if ratio and ratio >= 50:
            fz.append(f"Patrimonio aporta {fmt_num(ratio)}% del activo (sólido).")
        else:
            db.append(f"Bajo aporte patrimonial: PN/Activo = {fmt_num(ratio)}%.")
    # margen
    BAII = r24.get("_BAII") or 0.0
    if BAII and v24:
        margen = safe_div(BAII, v24) * 100
        if margen and margen >= 10:
            fz.append(f"Margen operativo sano: {fmt_num(margen)}%.")
        else:
            db.append(f"Margen operativo contenido: {fmt_num(margen)}%.")
    # liquidez
    liq = r24.get("Liquidez General")
    if liq is not None:
        if liq >= 1.5:
            fz.append(f"Liquidez general adecuada: {fmt_num(liq)}.")
        else:
            db.append(f"Liquidez general baja: {fmt_num(liq)} (óptimo ~1.5-2).")
    # dias clientes (approx using Deudores/Ventas*365 if no explicit dias)
    if v24 and r24.get("_Deudores") is not None:
        dias = safe_div(r24.get("_Deudores") * 365, v24)
        if dias and dias > 60:
            db.append(f"Ciclo de cobro largo: ~{int(dias)} días.")
        else:
            fz.append(f"Ciclo de cobro razonable: ~{int(dias)} días.")
    if not db:
        db.append("Dependencia de contratos/servicios (riesgo sectorial).")
    return fz[:6], db[:6]

def generar_diagnostico(r23, r24):
    lines = []
    lines.append(f"Diagnóstico ejecutivo — {datetime.date.today().isoformat()}")
    # ventas
    v23 = r23.get("_Ventas") or 0.0
    v24 = r24.get("_Ventas") or 0.0
    if v23 and v24 and v23 != 0:
        pct = (v24 - v23) / abs(v23) * 100
        lines.append(f"- Crecimiento de ventas: {fmt_num(pct)}%")
    else:
        lines.append("- Datos de ventas insuficientes para comparar crecimiento.")
    # FM
    fm23 = r23.get("Fondo Maniobra")
    fm24 = r24.get("Fondo Maniobra")
    if fm23 is not None and fm24 is not None:
        lines.append(f"- Fondo de Maniobra: 2023 = {fmt_num(fm23)}, 2024 = {fmt_num(fm24)}.")
        if fm24 > fm23:
            lines.append("  → Mejora del FM respecto a 2023.")
        elif fm24 < fm23:
            lines.append("  → FM empeora respecto a 2023; vigilar liquidez.")
    # RAT / RRP
    rat = r24.get("RAT")
    rrp = r24.get("RRP")
    if rat is not None:
        lines.append(f"- Rentabilidad económica (RAT) 2024: {fmt_num((rat or 0.0)*100)}%")
    if rrp is not None:
        lines.append(f"- Rentabilidad financiera (RRP) 2024: {fmt_num((rrp or 0.0)*100)}%")
    # risks
    cal = r24.get("Calidad Deuda")
    if cal is not None and cal > 0.6:
        lines.append(f"- Riesgo: alta deuda a corto plazo (PC/Pasivo = {fmt_num(cal)}).")
    # conclusion
    lines.append("")
    lines.append("Conclusión: La empresa muestra resultados que deben complementarse con mejoras en capital de trabajo y gestión de pasivos para sostener el crecimiento.")
    return "\n".join(lines)

def generar_recomendaciones(r23, r24):
    texto = []
    # 1 reducir dias clientes (estimate)
    v = r24.get("_Ventas") or 0.0
    dias_actuales = None
    if v and r24.get("_Deudores") is not None:
        dias_actuales = safe_div(r24.get("_Deudores") * 365, v)
        dias_actuales = int(dias_actuales) if dias_actuales else None
    dias_nuevo = (dias_actuales - 15) if dias_actuales else 45
    mejora_cash = None
    if dias_actuales and v:
        mejora_cash = safe_div((dias_actuales - dias_nuevo) * v, 365)
    texto.append(f"1) Mejorar cobros: reducir Días clientes de {dias_actuales or 'N/D'} a {dias_nuevo} → mejora estimada de caja: {fmt_num(mejora_cash)}")

    # 2 Refinanciar 30% PC
    PC = r24.get("_PC") or 0.0
    monto = PC * 0.30
    texto.append(f"2) Refinanciar 30% del Pasivo Corriente → monto aproximado: {fmt_num(monto)} (alivio en corto plazo)")

    # 3 Control costos 10% sobre gastos admin approximated by BAII? We'll estimate via BAII
    BAII = r24.get("_BAII") or 0.0
    impacto = BAII * 0.10 if BAII else None
    texto.append(f"3) Reducir gastos en 10% → impacto aproximado en BAII: {fmt_num(impacto)}")
    return "\n".join(texto)

# -----------------------------
# PDF: generar informe completo
# -----------------------------
def generar_pdf_final(r23, r24, filename="Informe_Financiero_Final.pdf"):
    width, height = A4
    c = canvas.Canvas(filename, pagesize=A4)
    x_margin = 2*cm
    y = height - 2*cm

    # Portada simple
    c.setFont("Helvetica-Bold", 16)
    c.setFillColor(colors.HexColor("#003366"))
    c.drawString(x_margin, y, "INFORME FINANCIERO Y ECONÓMICO")
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    c.drawString(x_margin, y - 12, f"Generado: {datetime.date.today().isoformat()}")
    y -= 36

    # 10.2: Texto introductorio (resumido)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_margin, y, "10.2 El Fondo de Maniobra y el análisis patrimonial de la empresa")
    y -= 14
    c.setFont("Helvetica", 9)
    intro = [
        "El análisis patrimonial estudia la estructura del activo y pasivo y el equilibrio financiero.",
        "El Fondo de Maniobra (FM = Activo Corriente - Pasivo Corriente) indica la capacidad para cubrir obligaciones a corto plazo.",
        "FM positivo sugiere financiación de circulante con recursos permanentes; FM negativo señala posible tensión de liquidez."
    ]
    for line in intro:
        if y < 4*cm:
            c.showPage(); y = height - 2*cm
        c.drawString(x_margin, y, line)
        y -= 12
    y -= 8

    # Tabla: Situación patrimonial (solo UNA tabla)
    tabla = crear_tabla_situacion()
    w, h = tabla.wrapOn(c, width - 4*cm, y)
    if y - h < 3*cm:
        c.showPage(); y = height - 2*cm
    tabla.drawOn(c, x_margin, y - h)
    y = y - h - 12

    # 10.4 Análisis Económico (RAT, RRP, Apalancamiento)
    if y < 6*cm:
        c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_margin, y, "10.4 Análisis económico")
    y -= 14
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_margin, y, "10.4.1 Rentabilidad económica (RAT)")
    y -= 12
    c.setFont("Helvetica", 9)
    rat = r24.get("RAT")
    c.drawString(x_margin, y, f"RAT (estimado BAII/Activo): {fmt_num((rat or 0.0)*100)}%")
    y -= 12
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_margin, y, "10.4.2 Rentabilidad financiera (RRP)")
    y -= 12
    c.setFont("Helvetica", 9)
    rrp = r24.get("RRP")
    c.drawString(x_margin, y, f"RRP (BN/PN): {fmt_num((rrp or 0.0)*100)}%")
    y -= 12
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_margin, y, "10.4.3 Apalancamiento financiero")
    y -= 12
    c.setFont("Helvetica", 9)
    apal = r24.get("Apalancamiento Financiero")
    c.drawString(x_margin, y, f"Apalancamiento (estimado): {fmt_num((apal or 0.0)*100)}%")
    y -= 18

    # D1 Matriz de ratios comparativa
    if y < 12*cm:
        c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_margin, y, "D1. Matriz de Ratios Comparativos 2023 vs 2024")
    y -= 14

    keys = ["Fondo Maniobra", "Liquidez General", "Tesorería", "Disponibilidad", "Garantía", "Autonomía", "Calidad Deuda", "RAT", "RRP"]
    header = ["Ratio", "2023", "2024", "Cambio (abs)", "Cambio (%)"]
    table_data = [header]
    for k in keys:
        v23 = r23.get(k)
        v24 = r24.get(k)
        abs_ch = (v24 - v23) if (v23 is not None and v24 is not None) else None
        pct_ch = safe_div(abs_ch, abs(v23)) * 100 if (abs_ch is not None and v23 not in (None,0)) else None
        table_data.append([k, fmt_num(v23), fmt_num(v24), fmt_num(abs_ch), (fmt_num(pct_ch) + "%" if pct_ch is not None else "N/A")])
    colw = [5*cm, 2.6*cm, 2.6*cm, 2.6*cm, 3.6*cm]
    t = Table(table_data, colWidths=colw)
    t.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),0.3,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#EEECEC")),
        ('FONT',(0,0),(-1,0),'Helvetica-Bold'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    w,t_h = t.wrapOn(c, width - 4*cm, y)
    t.drawOn(c, x_margin, y - t_h)
    y -= t_h + 12

    # Gráfico comparativo (algunos ratios)
    plot_keys = ["Liquidez General", "Tesorería", "Disponibilidad", "Garantía", "Autonomía", "Calidad Deuda"]
    vals23 = [r23.get(k) or 0 for k in plot_keys]
    vals24 = [r24.get(k) or 0 for k in plot_keys]
    fig, ax = plt.subplots(figsize=(8,3))
    x = np.arange(len(plot_keys))
    ax.bar(x - 0.18, vals23, width=0.35, label="2023", color="#66a3ff")
    ax.bar(x + 0.18, vals24, width=0.35, label="2024", color="#ff9f66")
    ax.set_xticks(x)
    ax.set_xticklabels(plot_keys, rotation=45, ha="right")
    ax.legend()
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='PNG', dpi=150)
    plt.close(fig)
    buf.seek(0)
    img = ImageReader(buf)
    if y - 6*cm < 2*cm:
        c.showPage(); y = height - 2*cm
    c.drawImage(img, x_margin, y - 6*cm, width=16*cm, height=6*cm)
    y -= 6.5*cm

    # D2 Fortalezas y Debilidades
    if y < 6*cm:
        c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_margin, y, "D2. Fortalezas y Debilidades")
    y -= 14
    fz, db = generar_fortalezas_debilidades(r23, r24)
    c.setFont("Helvetica-Bold", 10); c.drawString(x_margin, y, "Fortalezas:")
    y -= 12; c.setFont("Helvetica", 9)
    for it in fz:
        if y < 3*cm:
            c.showPage(); y = height - 2*cm
        c.drawString(x_margin + 10, y, "• " + it)
        y -= 12
    y -= 6
    c.setFont("Helvetica-Bold", 10); c.drawString(x_margin, y, "Debilidades:")
    y -= 12; c.setFont("Helvetica", 9)
    for it in db:
        if y < 3*cm:
            c.showPage(); y = height - 2*cm
        c.drawString(x_margin + 10, y, "• " + it)
        y -= 12
    y -= 10

    # D3 Diagnóstico ejecutivo
    if y < 6*cm:
        c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 12); c.drawString(x_margin, y, "D3. Diagnóstico ejecutivo")
    y -= 14
    c.setFont("Helvetica", 9)
    diag = generar_diagnostico(r23, r24)
    for line in diag.split("\n"):
        if y < 3*cm:
            c.showPage(); y = height - 2*cm
        c.drawString(x_margin, y, line)
        y -= 12
    y -= 8

    # D4 Recomendaciones
    if y < 6*cm:
        c.showPage(); y = height - 2*cm
    c.setFont("Helvetica-Bold", 12); c.drawString(x_margin, y, "D4. Recomendaciones estratégicas (3 medidas cuantificadas)")
    y -= 14
    c.setFont("Helvetica", 9)
    recs = generar_recomendaciones(r23, r24)
    for line in recs.split("\n"):
        if y < 3*cm:
            c.showPage(); y = height - 2*cm
        c.drawString(x_margin, y, line)
        y -= 12

    c.save()

# -----------------------------
# INTERFAZ TKINTER
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Informe Financiero - Versión Final (mínimos)")
        self.geometry("1000x680")
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Ingrese los datos financieros (valores en la misma unidad)", font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=4, pady=6)

        self.fields = [
            ("activo_corriente","Activo Corriente"),
            ("activo_no_corriente","Activo No Corriente"),
            ("pasivo_corriente","Pasivo Corriente"),
            ("pasivo_no_corriente","Pasivo No Corriente"),
            ("patrimonio_neto","Patrimonio Neto"),
            ("ventas","Ventas"),
            ("costo_ventas","Costo de ventas"),
            ("beneficio_neto","Beneficio Neto"),
            ("deudores","Deudores / Cuentas por cobrar"),
            ("inventario","Inventario"),
            ("caja","Caja / Efectivo"),
            ("i","Tasa interés (ej: 0.05)")
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
            # defaults
            if key == "i":
                e1.insert(0,"0.05"); e2.insert(0,"0.05")
            else:
                e1.insert(0,"0"); e2.insert(0,"0")
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
                        if key == "i":
                            val = float(txt)
                        else:
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
        out.append("=== D1: Matriz de Ratios 2023 vs 2024 ===")
        keys = ["Fondo Maniobra","Liquidez General","Tesorería","Disponibilidad","Garantía","Autonomía","Calidad Deuda","RAT","RRP"]
        for k in keys:
            out.append(f"{k}: 2023={fmt_num(r23.get(k))} | 2024={fmt_num(r24.get(k))}")
        out.append("\n=== D2: Fortalezas y Debilidades ===")
        fz, db = generar_fortalezas_debilidades(r23, r24)
        out.append("Fortalezas:")
        for i, it in enumerate(fz): out.append(f" {i+1}. {it}")
        out.append("Debilidades:")
        for i, it in enumerate(db): out.append(f" {i+1}. {it}")
        out.append("\n=== D3: Diagnóstico ejecutivo ===")
        out.append(generar_diagnostico(r23, r24))
        out.append("\n=== D4: Recomendaciones ===")
        out.append(generar_recomendaciones(r23, r24))
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, "\n".join(out))

    def export_pdf(self):
        data = self.leer_inputs()
        r23 = calcular_ratios_from_inputs(data["2023"])
        r24 = calcular_ratios_from_inputs(data["2024"])
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="Informe_Financiero_Final.pdf", filetypes=[("PDF files","*.pdf")])
        if not file_path:
            return
        generar_pdf_final(r23, r24, filename=file_path)
        messagebox.showinfo("PDF generado", f"Informe guardado en:\n{file_path}")

    def limpiar(self):
        for yr in ("2023","2024"):
            for ent in self.entries[yr].values():
                ent.delete(0, tk.END)
        self.output.delete("1.0", tk.END)

# -----------------------------
# EJECUTAR
# -----------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
