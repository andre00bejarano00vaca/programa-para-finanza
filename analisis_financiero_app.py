import tkinter as tk
from tkinter import messagebox

def calcular_ratios():
    try:
        activos_corrientes = float(entry_activos_corrientes.get())
        pasivos_corrientes = float(entry_pasivos_corrientes.get())
        pasivos_totales = float(entry_pasivos_totales.get())
        patrimonio_neto = float(entry_patrimonio_neto.get())
        
        # Calcular ratio de liquidez
        ratio_liquidez = activos_corrientes / pasivos_corrientes
        
        # Calcular ratio de endeudamiento
        ratio_endeudamiento = pasivos_totales / (patrimonio_neto + pasivos_totales)
        
        resultado = f"Ratio de liquidez: {ratio_liquidez:.2f}\n"
        resultado += f"Ratio de endeudamiento: {ratio_endeudamiento:.2f}"
        
        messagebox.showinfo("Resultados", resultado)
    except ValueError:
        messagebox.showerror("Error", "Por favor ingrese valores numéricos válidos")

# Configurar ventana principal
ventana = tk.Tk()
ventana.title("Calculadora Ratios Financieros")

# Entradas de datos
tk.Label(ventana, text="Activos Corrientes:").grid(row=0, column=0)
entry_activos_corrientes = tk.Entry(ventana)
entry_activos_corrientes.grid(row=0, column=1)

tk.Label(ventana, text="Pasivos Corrientes:").grid(row=1, column=0)
entry_pasivos_corrientes = tk.Entry(ventana)
entry_pasivos_corrientes.grid(row=1, column=1)

tk.Label(ventana, text="Pasivos Totales:").grid(row=2, column=0)
entry_pasivos_totales = tk.Entry(ventana)
entry_pasivos_totales.grid(row=2, column=1)

tk.Label(ventana, text="Patrimonio Neto:").grid(row=3, column=0)
entry_patrimonio_neto = tk.Entry(ventana)
entry_patrimonio_neto.grid(row=3, column=1)

# Botón para calcular
boton_calcular = tk.Button(ventana, text="Calcular Ratios", command=calcular_ratios)
boton_calcular.grid(row=4, columnspan=2)

ventana.mainloop()
