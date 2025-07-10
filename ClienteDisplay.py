import os
import sys
import chardet
import pandas as pd
import win32com.client as win32
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter
from customtkinter import *
import shutil

# Obtiene la ruta correcta para Assets
def resource_path(relative_path):
    try:
        basePath = sys._MEIPASS
    except Exception:
        basePath = os.path.abspath(".")
    return os.path.join(basePath, relative_path)

# Selecciona excel
def seleccionarExcel():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo:
        excelPath.set(archivo)
        cargar_proveedores()

# Carga proveedores + pedidos en el frame
def cargar_proveedores():
    archivoExcel = excelPath.get()
    if not archivoExcel:
        return
    
    try:
        data = pd.read_excel(archivoExcel)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el Excel:\n{str(e)}")
        return

    try:
        dataFilt = (data["Fecha Entrega"].dt.date < datetime.now().date()) & (data["Fecha Entrega"].dt.year == datetime.now().year)
        mailFilt = (data["Email"] != "") & (data["Email"] != ".") & (data["Email"] != ",")
        artFilt = data["Artículo"].str.contains("UN-") == False
        filt = data[(dataFilt) & (mailFilt) & (artFilt)]

        if filt.empty:
            messagebox.showinfo("No hay compras pendientes")
            return

        agrupado = filt.groupby(["Proveedor","Email"])["Pedido"].apply(list).reset_index()
        
        for widget in frameChecks.winfo_children():
            widget.destroy()
        checkboxes.clear()
        
        for _, row in agrupado.iterrows():
            proveedor = row["Proveedor"]
            correo = row["Email"]
            pedidos = row["Pedido"]
            
            proveedorFrame = customtkinter.CTkFrame(frameChecks,width=20,height=20)
            proveedorFrame.pack(anchor="w", padx=5, pady=5, fill="x",expand=True)

            provVar = tk.BooleanVar(value=True)
            provChk = customtkinter.CTkCheckBox(proveedorFrame, text=proveedor, variable=provVar,
                                                checkbox_height=20, checkbox_width=20, corner_radius=50)
            provChk.grid(row=0, column=0, sticky="w",padx=25)

            toggleBtn = customtkinter.CTkButton(proveedorFrame, text=">", width=20, height=20, font=("Arial", 12))
            toggleBtn.grid(row=0, column=0, padx=0,sticky="w")

            pedidosFrame = customtkinter.CTkFrame(proveedorFrame, width=20, height=20)
            pedidosFrame.grid(row=1, column=0, columnspan=2, sticky="w", padx=20, pady=2)
            pedidosFrame.grid_remove()

            pedidoVars = []
            internalChange = {"flag": False}
            
            for pedido in pedidos:
                var = tk.BooleanVar(value=True)
                if not any(p == pedido for p, _ in pedidoVars):
                    chk = customtkinter.CTkCheckBox(pedidosFrame, text=f"Pedido {pedido}", variable=var,
                                                    checkbox_width=18, checkbox_height=18, corner_radius=50)
                    chk.pack(anchor="w", padx=5, pady=1)
                    
                    pedidoVars.append((pedido, var))
                
                def on_pedido_change(var=var, provVar=provVar, pedidoVars=pedidoVars):
                    if internalChange["flag"]:
                        return
                    internalChange["flag"] = True
                    if var.get():
                        provVar.set(True)
                    else:
                        if not any(pVar.get() for _, pVar in pedidoVars):
                            provVar.set(False)
                    internalChange["flag"] = False

                var.trace_add("write", lambda *args, v=var, p=provVar, pv=pedidoVars: on_pedido_change(v, p, pv))
            
            def on_proveedor_change(*args, provVar=provVar, pedidoVars=pedidoVars):
                if internalChange["flag"]:
                    return
                estado = provVar.get()
                internalChange["flag"] = True
                for _, pVar in pedidoVars:
                    pVar.set(estado)
                internalChange["flag"] = False
            provVar.trace_add("write", on_proveedor_change)

            def toggle(frame=pedidosFrame, btn=toggleBtn):
                if frame.winfo_ismapped():
                    frame.grid_remove()
                    btn.configure(text=">")
                else:
                    frame.grid()
                    btn.configure(text="^")
            toggleBtn.configure(command=toggle)

            checkboxes.append((correo, provVar, pedidoVars))
            
            def selectAll():
                for _,provVar,_ in checkboxes:
                    provVar.set(True)
                
            def unselectAll():
                for _,provVar,_ in checkboxes:
                    provVar.set(False)
                
        frameBotones = customtkinter.CTkFrame(ventana, fg_color="transparent")
        frameBotones.grid(row=3, column=0, columnspan=3, pady=(0, 10),sticky="e")

        btnSAll = customtkinter.CTkButton(frameBotones, command=selectAll, text="Sel Todo",
                                        corner_radius=80, fg_color="transparent", border_color="grey",
                                        border_width=2, width=100, height=30)
        btnSAll.grid(padx=5,sticky="w",row=3,column=1)

        btnUAll = customtkinter.CTkButton(frameBotones, command=unselectAll, text="Desel Todo",
                                        corner_radius=80, fg_color="transparent", border_color="grey",
                                        border_width=2, width=100, height=30)
        btnUAll.grid(padx=50,row=3,column=3,columnspan=3)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo procesar proveedores:\n{e}")

# Define path para .exe y script
def get_app_path():
    if hasattr(sys, "_MEIPASS"):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# Selecciona firma
def seleccionarFirma():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo de firma (.txt o .htm)",
        filetypes=[("Archivos de texto o HTML", "*.txt *.htm *.html")]
    )
    if archivo:
        extension = os.path.splitext(archivo)[1].lower()
        if extension in [".htm", ".html"]:
            firmaPath.set(archivo)
            firmaText.delete("1.0", tk.END)
            firmaText.insert(tk.END, f"(Archivo HTML seleccionado: {archivo})")
        else:
            try:
                with open(archivo, "rb") as f:
                    raw = f.read()
                    detected = chardet.detect(raw)
                    encoding = detected['encoding']
                contenido = raw.decode(encoding)
                firmaText.delete("1.0", tk.END)
                firmaText.insert(tk.END, contenido)
                firmaPath.set("")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer la firma:\n{e}")
    try:
        basePath = get_app_path()
        carpetaFirmas = os.path.join(basePath, "Firmas")
        os.makedirs(carpetaFirmas, exist_ok=True)
        
        nombreArchivo = os.path.basename(archivo)
        destino = os.path.join(carpetaFirmas, nombreArchivo)

        print(os.path.abspath(archivo))
        print(os.path.abspath(destino))
        
        if os.path.abspath(archivo) != os.path.abspath(destino):
            if not os.path.exists(destino):
                shutil.copy2(archivo, destino)
                messagebox.showinfo("Firma guardada", f"Copia de la firma guardada en:\n{destino}")
            
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo copiar la firma:\n{e}")


# Traducción
def traduccion(asuntoDef,cuerpo,lang):
    if lang != "FR":
        with open(resource_path("Assets/Mail/CAP_EN.txt"),"r",encoding='utf-8') as arch:
            asuntoDef = arch.read()
        with open(resource_path("Assets/Mail/CUERPO_EN.txt"),'r',encoding='utf-8') as arch2:
            cuerpo = arch2.read()
    else:
        with open(resource_path("Assets/Mail/CAP_FR.txt"),"r",encoding='utf-8') as arch:
            asuntoDef = arch.read()
        with open(resource_path("Assets/Mail/CUERPO_FR.txt"),'r',encoding='utf-8') as arch2:
            cuerpo = arch2.read()
    return asuntoDef, cuerpo

# Lógica para enviar
def enviar_correos():
    if not excelPath.get():
        messagebox.showerror("Error", "No has seleccionado ningún archivo Excel.")
        return
    
    try:
        data = pd.read_excel(excelPath.get())
        dataFilt = (data["Fecha Entrega"].dt.date < datetime.now().date()) & (data["Fecha Entrega"].dt.year == datetime.now().year)
        mailFilt = (data["Email"] != "") & (data["Email"] != ".") & (data["Email"] != ",")
        artFilt = data["Artículo"].str.contains("UN-") == False
        filt = data[(dataFilt) & (mailFilt) & (artFilt)]
        outlook = win32.Dispatch('Outlook.Application')
        with open(resource_path("Assets/Mail/CAP_ES.txt"),"r",encoding='utf-8') as arch:
            asuntoBase = arch.read()
        with open(resource_path("Assets/Mail/CUERPO_ES.txt"),"r",encoding='utf-8') as arch2:
            cuerpoBase = arch2.read()
        firmaArchivo = firmaPath.get()
        if firmaArchivo:
            with open(firmaArchivo, "rb") as f:
                raw = f.read()
                detected = chardet.detect(raw)
                encoding = detected['encoding']
            firmaHtml = raw.decode(encoding)
        else:
            firmaHtml = "<p>" + firmaText.get("1.0", tk.END).replace("\n","<br>") + "</p>"

        for correo, provVar, pedidosList in checkboxes:
            if not provVar.get():
                continue
            pedidos_a_enviar = [
                str(p) for p, var in pedidosList if var.get()
            ]
            if not pedidos_a_enviar:
                continue
            pais = filt[filt["Email"]==correo].iloc[0]["Pais"]
            if pais != "ES":
                asuntoBase,cuerpoBase = traduccion(asuntoBase,cuerpoBase,pais)
            mail = outlook.CreateItem(0)
            mail.To = correo
            asuntoDef = asuntoBase + ", ".join(pedidos_a_enviar)
            cuerpo = cuerpoBase.replace("pedidoN", ", ".join(pedidos_a_enviar))
            htmlDef = cuerpo.replace("\n","<br>") + firmaHtml
            mail.Subject = asuntoDef
            mail.HTMLBody = htmlDef
            if messagebox.askyesno(message="Se enviaran correos a los proveedores selecionados. ¿Desea continuar?"):
                mail.Display()
                messagebox.showinfo("OK", "Correos generados correctamente en Outlook.")
                
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema al enviar correos:\n{str(e)}")

# UI
ventana = customtkinter.CTk()
ventana.title("Enviar Correos a Proveedores")
ventana.geometry("1000x500")
ventana._set_appearance_mode("dark")

excelPath = tk.StringVar()
firmaPath = tk.StringVar()
checkboxes = []

btnExcel = customtkinter.CTkButton(ventana,text="Seleccionar Excel",command=seleccionarExcel,corner_radius=50,fg_color="transparent",border_width=2,border_color="grey")
btnExcel.grid(row=0, column=2, padx=10, pady=20)

labelExcel = customtkinter.CTkLabel(ventana,text="Archivo Excel")
labelExcel.grid(row=0, column=0,padx=10, pady=20)
entryEx = customtkinter.CTkEntry(ventana,textvariable=excelPath,width=400)
entryEx.grid(row=0, column=1,padx=20, pady=20, sticky="ew")

firmaText = customtkinter.CTkTextbox(ventana,width=200,height=80,border_color="grey",border_width=2)
firmaText.grid(row=1, column=1,padx=20, pady=20, sticky="ew")
labelFirma = customtkinter.CTkLabel(ventana,text="Firma del correo")
labelFirma.grid(row=1, column=0,padx=10, pady=20)

btnFirma = customtkinter.CTkButton(ventana,text="Cargar firma desde archivo",command=seleccionarFirma,corner_radius=50,fg_color="transparent",border_width=2,border_color="grey")
btnFirma.grid(row=1, column=2,padx=10, pady=20)

btnEnviar = customtkinter.CTkButton(ventana,text="Enviar Mail",command=enviar_correos,corner_radius=50,fg_color="transparent",border_width=2,border_color="grey")
btnEnviar.grid(row=2, column=1, padx=20, pady=40)

frameChecks = customtkinter.CTkScrollableFrame(ventana, width=300, height=200, border_width=2, border_color="grey")
frameChecks.grid(row=2, column=0, columnspan=3, padx=15, pady=20, sticky="se")

ventana.grid_columnconfigure(0, weight=1)
ventana.grid_columnconfigure(1, weight=1)
ventana.grid_columnconfigure(2, weight=1)

ventana.grid_rowconfigure(0, weight=1)
ventana.grid_rowconfigure(1, weight=2)
ventana.grid_rowconfigure(2, weight=1)
ventana.grid_rowconfigure(3, weight=1)

ventana.mainloop()
