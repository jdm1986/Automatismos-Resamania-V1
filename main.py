import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import os
import sys
import json
import pandas as pd
from logic.wizville import procesar_wizville
from logic.accesos import procesar_accesos_dobles, procesar_accesos_descuadrados, procesar_salidas_pmr_no_autorizadas, \
    procesar_morosos_accediendo
from logic.ultimate import obtener_socios_ultimate
from utils.file_loader import load_data_file

CONFIG_PATH = "config.json"


def get_default_folder():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                data = json.load(f)
                return data.get("carpeta_datos", "")
        except json.JSONDecodeError:
            return ""
    return ""


def set_default_folder(path):
    with open(CONFIG_PATH, "w") as f:
        json.dump({"carpeta_datos": path}, f)


def get_logo_path(filename):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return filename


class ResamaniaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AUTOMATISMOS RESAMANIA - JDM Developer")
        self.state('zoomed')  # Pantalla completa al arrancar
        self.folder_path = get_default_folder()

        self.create_widgets()
        if self.folder_path:
            self.load_data()

    def create_widgets(self):
        top_frame = tk.Frame(self)
        top_frame.pack(pady=10)

        # Logo universal (compatible con .exe y .py)
        logo_path = get_logo_path("logodeveloper.png")
        if os.path.exists(logo_path):
            img = Image.open(logo_path)
            img = img.resize((140, 140))
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(top_frame, image=self.logo_img).pack(side=tk.LEFT, padx=10)

        instrucciones = (
            "INSTRUCCIONES PARA USO CORRECTO:\n\n"
            "- La carpeta seleccionada debe contener los siguientes archivos (exportados de Resamania):\n"
            "   - RESUMEN CLIENTE.csv\n"
            "   - ACCESOS.csv (intervalo de 4 semanas)\n"
            "   - FACTURAS Y VALES.csv (intervalo del mes anterior, del dia 1 al 5)\n"
            "   - IMPAGOS.csv (Incidencias por tipo de pago, ultimos 3 meses)\n\n"
            "- Todos los archivos deben ser del mismo dia de exportacion.\n"
            "- Pulsa el boton 'Seleccionar carpeta' para comenzar la revision.\n\n"
            "NOTA: Una vez seleccionada una carpeta, el programa la mantiene por defecto hasta que elijas otra."
        )
        tk.Label(top_frame, text=instrucciones, justify='left', anchor='w').pack(side=tk.LEFT, padx=10)

        botones_frame = tk.Frame(self)
        botones_frame.pack(pady=5)
        tk.Button(botones_frame, text="Seleccionar carpeta", command=self.select_folder).pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="Actualizar datos", command=self.load_data).pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="Exportar a Excel", command=self.exportar_excel).pack(side=tk.LEFT, padx=10)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=1, fill='both')

        self.tabs = {}
        for tab_name in [
            "Wizville", "Accesos Dobles", "Accesos Descuadrados",
            "Salidas PMR No Autorizadas", "Morosos Accediendo", "Socios Ultimate"
        ]:
            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=tab_name)
            self.tabs[tab_name] = tab
            self.create_table(tab)

    def create_table(self, tab):
        tree = ttk.Treeview(tab, show="headings")
        tree.pack(expand=True, fill='both', side='left')

        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
        scrollbar.pack(side='right', fill='y')
        tree.configure(yscrollcommand=scrollbar.set)

        def sort_column(col, reverse):
            datos = [(tree.set(k, col), k) for k in tree.get_children('')]
            try:
                datos.sort(key=lambda t: float(t[0].replace(",", ".")), reverse=reverse)
            except ValueError:
                datos.sort(key=lambda t: t[0], reverse=reverse)
            for index, (_, k) in enumerate(datos):
                tree.move(k, '', index)
            tree.heading(col, command=lambda: sort_column(col, not reverse))

        tree.bind("<Control-c>", lambda e: self.copiar_celda(e, tree))
        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Copiar", command=lambda: self.copiar_celda(tree.event_context, tree))
        tree.bind("<Button-3>", lambda e: self.mostrar_menu(e, menu, tree))

        tab.tree = tree

    def copiar_celda(self, event, tree):
        seleccion = tree.selection()
        if seleccion:
            item = seleccion[0]
            col = tree.identify_column(event.x)
            col_index = int(col.replace('#', '')) - 1
            valor = tree.item(item)['values'][col_index]
            self.clipboard_clear()
            self.clipboard_append(str(valor))

    def mostrar_menu(self, event, menu, tree):
        tree.event_context = event
        menu.post(event.x_root, event.y_root)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            set_default_folder(folder)
            self.load_data()

    def load_data(self):
        if not self.folder_path:
            messagebox.showerror("Error", "No se ha seleccionado carpeta")
            return

        try:
            resumen = load_data_file(self.folder_path, "RESUMEN CLIENTE")
            accesos = load_data_file(self.folder_path, "ACCESOS")
            incidencias = load_data_file(self.folder_path, "IMPAGOS")

            self.mostrar_en_tabla("Wizville", procesar_wizville(resumen, accesos))
            self.mostrar_en_tabla("Accesos Dobles", procesar_accesos_dobles(resumen, accesos))
            self.mostrar_en_tabla("Accesos Descuadrados", procesar_accesos_descuadrados(resumen, accesos))
            self.mostrar_en_tabla("Salidas PMR No Autorizadas", procesar_salidas_pmr_no_autorizadas(resumen, accesos))
            self.mostrar_en_tabla("Morosos Accediendo", procesar_morosos_accediendo(incidencias, accesos), color="#F4CCCC")
            self.mostrar_en_tabla("Socios Ultimate", obtener_socios_ultimate())

            messagebox.showinfo("Exito", "Datos cargados correctamente.")
        except Exception as e:
            messagebox.showerror("Error al cargar datos", str(e))

    def mostrar_en_tabla(self, tab_name, df, color=None):
        tab = self.tabs[tab_name]
        tree = tab.tree

        tree.delete(*tree.get_children())
        tree["columns"] = list(df.columns)

        for col in df.columns:
            tree.heading(col, text=col, command=lambda c=col: self.sort_column(tree, c, False))
            tree.column(col, anchor="center")

        tree.tag_configure("amarillo", background="#fff3cd")
        tree.tag_configure("rojo", background="#f8d7da")

        for _, row in df.iterrows():
            valores = list(row)
            tag = ""
            if "Días desde alta" in df.columns:
                if row["Días desde alta"] == 16:
                    tag = "amarillo"
                elif row["Días desde alta"] == 180:
                    tag = "rojo"
            tree.insert("", "end", values=valores, tags=(tag,))

    def sort_column(self, tree, col, reverse):
        datos = [(tree.set(k, col), k) for k in tree.get_children('')]
        try:
            datos.sort(key=lambda t: float(t[0].replace(",", ".")), reverse=reverse)
        except ValueError:
            datos.sort(key=lambda t: t[0], reverse=reverse)
        for index, (_, k) in enumerate(datos):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.sort_column(tree, col, not reverse))

    def exportar_excel(self):
        try:
            pestana_activa = self.notebook.select()
            nombre_pestana = self.notebook.tab(pestana_activa, "text")
            tree = self.tabs[nombre_pestana].tree

            if not tree.get_children():
                messagebox.showwarning("Sin datos", f"No hay datos para exportar en la pestana {nombre_pestana}.")
                return

            columnas = tree["columns"]
            datos = [tree.item(i)["values"] for i in tree.get_children()]
            df_exportar = pd.DataFrame(datos, columns=columnas)

            archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"{nombre_pestana.replace(' ', '_')}.xlsx"
            )

            if archivo:
                df_exportar.to_excel(archivo, index=False)
                messagebox.showinfo("Exito", f"Datos exportados correctamente a:\n{archivo}")
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))


if __name__ == "__main__":
    app = ResamaniaApp()
    app.mainloop()
