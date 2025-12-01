import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle

# -----------------------------
# UTILIDADES FINANCIERAS
# -----------------------------
def safe_div(a, b):
    try:
        return a / b if b != 0 else None
    except:
        return None

def fmt(val, digits=2):
    """Formato boliviano: punto para miles, coma para decimales"""
    if val is None:
        return "N/A"
    s = f"{val:,.{digits}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def pct_change(old, new):
    if old is None or new is None or old == 0:
        return None
    return (new - old) / abs(old) * 100

# -----------------------------
# UMBRALES FINANCIEROS
# -----------------------------
OPT_LIQ_MIN = 1.5
OPT_RAT_MIN = 6
OPT_RRP_MIN = 8

# -----------------------------
# CÁLCULO DE RATIOS
# -----------------------------
def calculate_all(inputs):
    out = {}
    for year in ("2023", "2024"):
        AC = inputs[year].get("activo_corriente")
        PC = inputs[year].get("pasivo_corriente")
        INV = inputs[year].get("inventario")
        EF = inputs[year].get("efectivo")
        BAII = inputs[year].get("BAII")
        BN = inputs[year].get("beneficio_neto")
        AT = inputs[year].get("activo_total")
        PN = inputs[year].get("patrimonio_neto")

        FM = AC - PC if AC is not None and PC is not None else None
        LIQ = safe_div(AC, PC)
        RAT = safe_div(AC - INV, PC)
        RRP = safe_div(BN, PN) * 100 if BN is not None and PN is not None else None
        LA = safe_div(EF, PC)

        out[year] = {"FM": FM, "LIQ": LIQ, "RAT": RAT, "RRP": RRP, "LA": LA}

    changes = {}
    for key in ("FM","LIQ","RAT","RRP","LA"):
        old = out["2023"][key]
        new = out["2024"][key]
        changes[key] = {
            "2023": old,
            "2024": new,
            "abs": None if (old is None or new is None) else (new - old),
            "pct": pct_change(old,new)
        }

    return changes

# -----------------------------
# INTERPRETACIÓN AUTOMÁTICA
# -----------------------------
def interpretar_ratio(nombre, v2023, v2024):
    if v2023 is None or v2024 is None:
        return "Datos insuficientes"
    cambio = v2024 - v2023
    if nombre=="Fondo de Maniobra":
        return "Mejora capacidad de cubrir obligaciones." if cambio>0 else "Disminuye margen operativo."
    if nombre=="Liquidez General":
        if v2024 >= OPT_LIQ_MIN:
            return "Liquidez adecuada."
        elif v2024 >= v2023:
            return "Liquidez estable."
        else:
            return "Liquidez empeorando."
    if nombre=="Razón Ácida":
        return "Mayor solvencia inmediata." if cambio>0 else "Menor solvencia inmediata."
    if nombre=="Liquidez Absoluta":
        return "Más efectivo disponible." if cambio>0 else "Disminuye efectivo disponible."
    if nombre=="ROE":
        return "Rentabilidad del patrimonio mejora." if cambio>0 else "Rentabilidad del patrimonio disminuye."
    return "Interpretación no definida."

# -----------------------------
# MATRIZ D1
# -----------------------------
def matriz_d1(changes):
    filas=[]
    nombres={
        "FM":"Fondo de Maniobra",
        "LIQ":"Liquidez General",
        "RAT":"Razón Ácida",
        "LA":"Liquidez Absoluta",
        "RRP":"ROE"
    }
    for key,nombre in nombres.items():
        d = changes[key]
        cambio = None if (d["2023"] is None or d["2024"] is None) else d["2024"] - d["2023"]
        filas.append({
            "Ratio": nombre,
            "2023": fmt(d["2023"]),
            "2024": fmt(d["2024"]),
            "Cambio": fmt(cambio),
            "Interpretación": interpretar_ratio(nombre,d["2023"],d["2024"])
        })
    return pd.DataFrame(filas)

# -----------------------------
# D2: Fortalezas y Debilidades
# -----------------------------
def fortalezas_debilidades(inputs):
    fz=[]
    db=[]
    # Ventas
    v2023 = inputs["2023"].get("ventas")
    v2024 = inputs["2024"].get("ventas")
    if v2023 and v2024:
        pct = (v2024 - v2023)/v2023*100
        fz.append(f"Crecimiento de ventas: {fmt(pct)}% entre 2023-2024.")
    # Patrimonio / Activo
    PN = inputs["2024"].get("patrimonio_neto")
    AT = inputs["2024"].get("activo_total")
    if PN and AT:
        ratio = PN/AT*100
        fz.append(f"Patrimonio financiando buena parte del crecimiento: PN/Activo 2024 = {fmt(ratio)}%.")
    # Margen bruto
    BAII = inputs["2024"].get("BAII")
    if BAII and v2024:
        margen = BAII/v2024*100
        fz.append(f"Margen bruto sano para el sector: {fmt(margen)}%.")
    # Debilidades
    PC = inputs["2024"].get("pasivo_corriente")
    PT = inputs["2024"].get("pasivo_total")
    if PC and PT:
        deuda_corriente = PC/PT*100
        db.append(f"Elevada deuda corriente relativa: PC/Total Pasivo 2024 = {fmt(deuda_corriente)}%.")
    dias_clientes = inputs["2024"].get("dias_clientes")
    if dias_clientes:
        db.append(f"Necesidad de mejorar cobros (Días clientes {dias_clientes}).")
    db.append("Dependencia de contratos/servicios que pueden ser volátiles.")
    return fz, db

# -----------------------------
# D3: Diagnostico Financiero Integral
# -----------------------------
def diagnostico_integral(inputs, changes):
    ventas_2023 = inputs["2023"].get("ventas",0)
    ventas_2024 = inputs["2024"].get("ventas",0)
    BAII_2024 = inputs["2024"].get("BAII",0)
    BN_2024 = inputs["2024"].get("beneficio_neto",0)
    PN_2024 = inputs["2024"].get("patrimonio_neto",0)
    AC_2024 = inputs["2024"].get("activo_corriente",0)
    PC_2024 = inputs["2024"].get("pasivo_corriente",0)
    AT_2024 = inputs["2024"].get("activo_total",0)

    fm = changes["FM"]["2024"]
    liq = changes["LIQ"]["2024"]
    rat = changes["RAT"]["2024"]
    rrp = changes["RRP"]["2024"]

    pct_ventas = (ventas_2024 - ventas_2023)/ventas_2023*100 if ventas_2023 else None
    ratio_pn_activo = safe_div(PN_2024,AT_2024)*100 if PN_2024 and AT_2024 else None
    margen_bruto = safe_div(BAII_2024,ventas_2024)*100 if BAII_2024 and ventas_2024 else None

    texto = "D3. Diagnóstico Ejecutivo (Resumen Integral)\n\n"
    texto += f"La empresa presenta un crecimiento significativo entre 2023 y 2024, evidenciado por un aumento en ingresos y utilidades"
    if pct_ventas:
        texto += f" ({fmt(pct_ventas)}% de crecimiento en ventas)"
    texto += ", reflejado en la mejora de la rentabilidad financiera y económica.\n"

    if margen_bruto:
        texto += f"El margen bruto se mantiene saludable ({fmt(margen_bruto)}%), indicando eficiencia en costos.\n"
    if ratio_pn_activo:
        texto += f"El patrimonio financia una proporción creciente del activo ({fmt(ratio_pn_activo)}%), reduciendo dependencia del endeudamiento.\n"
    if fm:
        texto += f"El Fondo de Maniobra mejora respecto a 2023 ({fmt(fm)}), mostrando margen de seguridad en obligaciones de corto plazo.\n"
    if liq:
        texto += f"La liquidez general se sitúa en {fmt(liq)}, sugiriendo necesidad de optimizar capital de trabajo.\n"
    if rat:
        texto += f"La rentabilidad sobre activos (RAT) muestra un comportamiento positivo, evidenciando eficiencia en uso de recursos.\n"
    if rrp:
        texto += f"La rentabilidad sobre patrimonio (RRP) es {fmt(rrp)}%, mostrando generación de valor económico para los accionistas.\n"

    dias_clientes = inputs["2024"].get("dias_clientes")
    if dias_clientes:
        texto += f"El ciclo de cobranza es de aproximadamente {dias_clientes} días, lo que puede afectar liquidez y flujo de caja.\n"

    texto += "En conclusión, la empresa se encuentra en una situación financiera saludable, "
    texto += "con buenos niveles de rentabilidad y dinámica de crecimiento sostenida. "
    texto += "Sin embargo, debe mejorar gestión del capital de trabajo, equilibrar la deuda y optimizar eficiencia operativa.\n"
    return texto

# -----------------------------
# D4: Recomendaciones Estrategicas
# -----------------------------
def recomendaciones_estrategicas(inputs, changes):
    texto = "D4. Recomendaciones Estratégicas (3 medidas cuantificadas)\n\n"

    dias_clientes = inputs["2024"].get("dias_clientes",60)
    mejora_dias = max(dias_clientes - 15, 1)
    texto += f"1. Mejorar cobros: reducir Días clientes de {dias_clientes} a {mejora_dias} → estimación mejora de flujo de caja.\n"

    BAII_2024 = inputs["2024"].get("BAII",0)
    recorte_gastos_pct = 10
    aumento_baii = BAII_2024 * recorte_gastos_pct / 100 if BAII_2024 else None
    if aumento_baii:
        texto += f"2. Control de costos: reducir gastos administrativos en {recorte_gastos_pct}% → estimación aumento de BAII en {fmt(aumento_baii)}.\n"
    else:
        texto += f"2. Control de costos: reducir gastos administrativos en {recorte_gastos_pct}% → impacto estimado en BAII.\n"

    PC = inputs["2024"].get("pasivo_corriente",0)
    transfer_pct = 30
    monto_transferir = PC * transfer_pct / 100 if PC else None
    if monto_transferir:
        texto += f"3. Refinanciar parte de la deuda corriente a largo plazo: transferir {transfer_pct}% del PC → mejora FM estimada en {fmt(monto_transferir)}.\n"
    else:
        texto += f"3. Refinanciar parte de la deuda corriente a largo plazo: transferir {transfer_pct}% del PC → mejora FM estimada.\n"

    return texto

# -----------------------------
# GENERAR PDF
# -----------------------------
def generar_pdf(texto, nombre_archivo="Informe_Financiero.pdf"):
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    width, height = A4
    x = 2*cm
    y = height - 2*cm
    c.setFont("Helvetica", 10)
    for line in texto.split("\n"):
        if y < 2*cm:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - 2*cm
        c.drawString(x, y, line)
        y -= 12
    c.save()

# -----------------------------
# INTERFAZ GRÁFICA
# -----------------------------
class App:
    def __init__(self, root):
        self.root=root
        root.title("Análisis Financiero 2023 vs 2024")
        root.geometry("1100x750")
        container=ttk.Frame(root,padding=10)
        container.pack(fill="both",expand=True)

        fields=[
            ("ventas","Ventas"),
            ("activo_corriente","Activo Corriente"),
            ("pasivo_corriente","Pasivo Corriente"),
            ("inventario","Inventario"),
            ("efectivo","Efectivo"),
            ("BAII","BAII"),
            ("beneficio_neto","Beneficio Neto"),
            ("activo_total","Activo Total"),
            ("patrimonio_neto","Patrimonio Neto"),
            ("pasivo_total","Pasivo Total"),
            ("dias_clientes","Días Clientes")
        ]

        self.entries={"2023":{},"2024":{}}

        ttk.Label(container,text="Ingrese los datos financieros",font=("Arial",14,"bold")).grid(row=0,column=0,columnspan=5,pady=5)
        ttk.Label(container,text="Concepto").grid(row=1,column=0)
        ttk.Label(container,text="2023").grid(row=1,column=1)
        ttk.Label(container,text="2024").grid(row=1,column=2)
        r=2
        for key,label in fields:
            ttk.Label(container,text=label).grid(row=r,column=0,sticky="w",pady=2)
            e1=ttk.Entry(container,width=18)
            e2=ttk.Entry(container,width=18)
            e1.grid(row=r,column=1)
            e2.grid(row=r,column=2)
            self.entries["2023"][key]=e1
            self.entries["2024"][key]=e2
            r+=1

        btn_frame=ttk.Frame(container)
        btn_frame.grid(row=r,column=0,columnspan=5,pady=10)
        ttk.Button(btn_frame,text="Generar D1",command=self.gen_d1).grid(row=0,column=0,padx=10)
        ttk.Button(btn_frame,text="Generar D2",command=self.gen_d2).grid(row=0,column=1,padx=10)
        ttk.Button(btn_frame,text="Generar D3",command=self.gen_d3).grid(row=0,column=2,padx=10)
        ttk.Button(btn_frame,text="Generar D4",command=self.gen_d4).grid(row=0,column=3,padx=10)
        ttk.Button(btn_frame,text="Exportar PDF",command=self.export_pdf).grid(row=0,column=4,padx=10)
        ttk.Button(btn_frame,text="Limpiar",command=self.limpiar).grid(row=0,column=5,padx=10)

        self.tree=ttk.Treeview(container,columns=("Ratio","2023","2024","Cambio","Interpretación"),show="headings",height=7)
        self.tree.grid(row=r+1,column=0,columnspan=6,pady=10)
        for col in ["Ratio","2023","2024","Cambio","Interpretación"]:
            self.tree.heading(col,text=col)
            self.tree.column(col,width=200,anchor="center")

        self.output=tk.Text(container,height=15,width=130,font=("Consolas",10))
        self.output.grid(row=r+2,column=0,columnspan=6,pady=5)

    def read_inputs(self):
        data={"2023":{},"2024":{}}
        for yr in ("2023","2024"):
            for key,ent in self.entries[yr].items():
                val=ent.get().strip()
                data[yr][key]=float(val) if val else None
        return data

    def gen_d1(self):
        inputs=self.read_inputs()
        changes=calculate_all(inputs)
        df=matriz_d1(changes)
        self.tree.delete(*self.tree.get_children())
        for _,row in df.iterrows():
            self.tree.insert("", "end", values=(row["Ratio"],row["2023"],row["2024"],row["Cambio"],row["Interpretación"]))
        report=""
        for _,row in df.iterrows():
            report+=f"{row['Ratio']}:\n  2023: {row['2023']}\n  2024: {row['2024']}\n  Cambio: {row['Cambio']}\n  Interpretación: {row['Interpretación']}\n\n"
        self.output.delete("1.0",tk.END)
        self.output.insert(tk.END,report)

    def gen_d2(self):
        inputs=self.read_inputs()
        fz,db=fortalezas_debilidades(inputs)
        texto="=== FORTALEZAS ===\n"
        for item in fz:
            texto+=f"• {item}\n"
        texto+="\n=== DEBILIDADES ===\n"
        for item in db:
            texto+=f"• {item}\n"
        self.output.delete("1.0",tk.END)
        self.output.insert(tk.END,texto)

    def gen_d3(self):
        inputs = self.read_inputs()
        changes = calculate_all(inputs)
        texto = diagnostico_integral(inputs, changes)
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, texto)

    def gen_d4(self):
        inputs = self.read_inputs()
        changes = calculate_all(inputs)
        texto = recomendaciones_estrategicas(inputs, changes)
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, texto)

    def export_pdf(self):
        inputs = self.read_inputs()
        changes = calculate_all(inputs)
        d1 = matriz_d1(changes).to_string(index=False)
        fz, db = fortalezas_debilidades(inputs)
        d2 = "=== FORTALEZAS ===\n" + "\n".join(fz) + "\n\n=== DEBILIDADES ===\n" + "\n".join(db)
        d3 = diagnostico_integral(inputs, changes)
        d4 = recomendaciones_estrategicas(inputs, changes)
        texto_completo = "\n\n".join([d1, d2, d3, d4])
        generar_pdf(texto_completo)
        messagebox.showinfo("PDF generado", "Informe exportado como 'Informe_Financiero.pdf'")

    def limpiar(self):
        for yr in ("2023","2024"):
            for ent in self.entries[yr].values():
                ent.delete(0,tk.END)
        self.tree.delete(*self.tree.get_children())
        self.output.delete("1.0",tk.END)

# -----------------------------
# EJECUCIÓN
# -----------------------------
if __name__=="__main__":
    root=tk.Tk()
    App(root)
    root.mainloop()
