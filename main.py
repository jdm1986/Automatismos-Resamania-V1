import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import os
import sys
import json
import pandas as pd
import unicodedata
import time
import uuid
import tempfile
from datetime import datetime
import re
import urllib.parse
import webbrowser
import shutil
import random
from logic.wizville import procesar_wizville
from logic.accesos import procesar_accesos_dobles, procesar_accesos_descuadrados, procesar_salidas_pmr_no_autorizadas, \
    procesar_morosos_accediendo
from logic.ultimate import obtener_socios_ultimate, obtener_socios_yanga
from logic.avanza_fit import obtener_avanza_fit
from logic.cumpleanos import obtener_cumpleanos_hoy
from utils.file_loader import load_data_file
from logic.impagos import ImpagosDB
from logic.incidencias import IncidenciasDB

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


def get_security_code() -> str:
    """
    Codigo dinamico: ano + mes - dia (ej. 2025 + 12 - 14 = 2023).
    """
    now = datetime.now()
    return str(now.year + now.month - now.day)


class ResamaniaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AUTOMATISMOS RESAMANIA - JDM Developer")
        self.state('zoomed')  # Pantalla completa al arrancar
        self.folder_path = get_default_folder()
        self.dataframes = {}
        self.raw_accesos = None
        self.resumen_df = None

        # Datos de prestamos
        self.prestamos_file = os.path.join("data", "prestamos.json")
        self.prestamos = []
        self.clientes_ext_file = os.path.join("data", "clientes_ext.json")
        self.clientes_ext = []
        self.prestamos_filtro_activo = False

        # Felicitaciones (persistencia anual)
        self.felicitaciones_file = os.path.join("data", "felicitaciones.json")
        self.felicitaciones_enviadas = {}

        # Impagos (SQLite)
        self.impagos_db = ImpagosDB(os.path.join("data", "impagos.db"))
        self.impagos_last_export = None
        self.impagos_view = tk.StringVar(value="actuales")

        # Incidencias Club
        self.incidencias_db = IncidenciasDB(os.path.join("data", "incidencias.db"))
        self.incidencias_mapas = []
        self.incidencias_map_index = 0
        self.incidencias_img = None
        self.incidencias_canvas = None
        self.incidencias_current_map = None
        self.incidencias_area_items = {}
        self.incidencias_machine_items = {}
        self.incidencias_machine_area = {}
        self.incidencias_selected_area = None
        self.incidencias_selected_machine = None
        self.incidencias_mode = None
        self.incidencias_draw_start = None
        self.incidencias_draw_rect = None
        self.incidencias_panel_mode = None
        self.incidencias_color_used = set()

        self.incidencias_info_filter_area = None
        self.incidencias_area_title = None
        self.incidencias_btn_vista_general = None
        self.incidencias_hover_blink_item = None
        self.incidencias_hover_blink_after = None
        self.incidencias_hover_blink_original = None

        self.create_widgets()
        self.cargar_clientes_ext()
        self.cargar_prestamos_json()
        self.cargar_felicitaciones()
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
            "- La carpeta seleccionada debe contener los siguientes archivos (exportados de Resamania y deben sobreescribir los existentes):\n"
            "   - RESUMEN CLIENTE.csv\n"
            "   - ACCESOS.csv (intervalo de 4 semanas atrás)\n"
            "   - FACTURAS Y VALES.csv (intervalo de 4 semanas atrás)\n"
            "   - IMPAGOS.csv (Exportar el día actual el archivo - Clientes con Incidente de Pago)\n\n"
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
        tk.Button(botones_frame, text="DÍA ASUNTOS PROPIOS", command=self.enviar_asuntos_propios, fg="#0066cc").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="SOLICITUD DE CAMBIO DE TURNO", command=self.enviar_cambio_turno, fg="#0066cc").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="ENVIAR FELICITACIÓN", command=self.enviar_felicitacion, fg="#b30000").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="ENVÍO MARTES AVANZA FIT", command=self.enviar_avanza_fit, fg="#b30000").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="EXTRAER ACCESOS", command=self.extraer_accesos, fg="#0066cc").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="IR A PRÉSTAMOS", command=lambda: self.notebook.select(self.tabs.get("Prestamos")), bg="#ff9800", fg="black").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="IR A IMPAGOS", command=self.ir_a_impagos, bg="#ff6b6b", fg="black").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="INCIDENCIAS CLUB", command=lambda: self.notebook.select(self.tabs.get("Incidencias Club")), bg="#424242", fg="white").pack(side=tk.LEFT, padx=10)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=1, fill='both')
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        tab_colors = {
            "Wizville": "#6fa8dc",
            "Accesos Dobles": "#b4a7d6",
            "Accesos Descuadrados": "#8e7cc3",
            "Salidas PMR No Autorizadas": "#c27ba0",
            "Morosos Accediendo": "#e06666",
            "Socios Ultimate": "#76a5af",
            "Socios Yanga": "#93c47d",
            "Avanza Fit": "#ffd966",
            "Cumpleaños": "#f9cb9c",
            "Accesos Cliente": "#cfe2f3",
            "Prestamos": "#ffb74d",
            "Impagos": "#ff6b6b",
            "Incidencias Club": "#424242",
        }
        self.tab_icons = {}

        self.tabs = {}
        for tab_name in [
            "Wizville", "Accesos Dobles", "Accesos Descuadrados",
            "Salidas PMR No Autorizadas", "Morosos Accediendo", "Socios Ultimate", "Socios Yanga",
            "Avanza Fit", "Cumpleaños", "Accesos Cliente", "Prestamos", "Impagos", "Incidencias Club"
        ]:
            tab = ttk.Frame(self.notebook)
            color = tab_colors.get(tab_name, "#cccccc")
            icon = tk.PhotoImage(width=14, height=14)
            icon.put(color, to=(0, 0, 14, 14))
            self.tab_icons[tab_name] = icon
            self.notebook.add(tab, text=tab_name, image=icon, compound="left")
            self.tabs[tab_name] = tab
            if tab_name == "Prestamos":
                self.create_prestamos_tab(tab)
            elif tab_name == "Impagos":
                self.create_impagos_tab(tab)
            elif tab_name == "Incidencias Club":
                self.create_incidencias_tab(tab)
            else:
                self.create_table(tab)
            if tab_name == "Impagos":
                self.notebook.hide(tab)

    def create_table(self, tab):
        container = tk.Frame(tab)
        container.pack(expand=True, fill='both')
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        tree = ttk.Treeview(container, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        hscrollbar = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        hscrollbar.grid(row=1, column=0, sticky="ew")
        tree.configure(xscrollcommand=hscrollbar.set)

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
            self.resumen_df = resumen.copy()
            accesos = load_data_file(self.folder_path, "ACCESOS")
            self.raw_accesos = accesos.copy()
            incidencias = load_data_file(self.folder_path, "IMPAGOS")
            self.sync_impagos(incidencias)

            self.mostrar_en_tabla("Wizville", procesar_wizville(resumen, accesos))
            self.mostrar_en_tabla("Accesos Dobles", procesar_accesos_dobles(resumen, accesos))
            self.mostrar_en_tabla("Accesos Descuadrados", procesar_accesos_descuadrados(resumen, accesos))
            self.mostrar_en_tabla("Salidas PMR No Autorizadas", procesar_salidas_pmr_no_autorizadas(resumen, accesos))
            self.mostrar_en_tabla("Morosos Accediendo", procesar_morosos_accediendo(incidencias, accesos), color="#F4CCCC")
            self.mostrar_en_tabla("Socios Ultimate", obtener_socios_ultimate())
            self.mostrar_en_tabla("Socios Yanga", obtener_socios_yanga())
            self.mostrar_en_tabla("Avanza Fit", obtener_avanza_fit())
            self.mostrar_en_tabla("Cumpleaños", obtener_cumpleanos_hoy())

            messagebox.showinfo("Exito", "Datos cargados correctamente.")
        except Exception as e:
            messagebox.showerror("Error al cargar datos", str(e))

    def mostrar_en_tabla(self, tab_name, df, color=None):
        # Guarda el ultimo dataframe mostrado para poder reutilizarlo (ej. enviar emails)
        self.dataframes[tab_name] = df.copy()
        tab = self.tabs[tab_name]
        tree = tab.tree

        tree.delete(*tree.get_children())
        tree["columns"] = list(df.columns)

        for col in df.columns:
            tree.heading(col, text=col, command=lambda c=col: self.sort_column(tree, c, False))
            width = max(80, len(str(col)) * 8)
            tree.column(col, anchor="center", width=width, stretch=True)

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

    # -----------------------------
    # PESTAÑA PRÉSTAMOS
    # -----------------------------
    def create_prestamos_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        top = tk.Frame(frm)
        top.pack(fill="x", pady=4)
        tk.Label(top, text="Numero de cliente:").pack(side="left")
        self.prestamo_codigo = tk.Entry(top, width=14)
        self.prestamo_codigo.pack(side="left", padx=5)
        tk.Button(top, text="Buscar", command=self.buscar_cliente_prestamo).pack(side="left", padx=5)
        tk.Button(top, text="Editar cliente manual", command=self.editar_cliente_manual).pack(side="left", padx=5)
        tk.Button(top, text="AGREGAR CLIENTE DE OTRO FITNESS PARK", command=self.agregar_cliente_otro_fp, bg="#ff7043", fg="white").pack(side="left", padx=5)
        self.lbl_info = tk.Label(top, text="", anchor="w")
        self.lbl_info.pack(side="left", padx=10)

        mid = tk.Frame(frm)
        mid.pack(fill="x", pady=6)
        tk.Label(mid, text="Material prestado:").pack(side="left")
        self.prestamo_material = tk.Entry(mid, width=40)
        self.prestamo_material.pack(side="left", padx=5)
        tk.Button(mid, text="Guardar prestamo", command=self.guardar_prestamo, bg="#ffd54f", fg="black").pack(side="left", padx=5)
        tk.Button(mid, text="CHECK DEVUELTO", command=self.marcar_devuelto, bg="#8bc34a", fg="black").pack(side="left", padx=5)
        tk.Button(mid, text="Enviar aviso email", command=self.enviar_aviso_prestamo, bg="#e53935", fg="white").pack(side="left", padx=5)
        tk.Button(mid, text="Enviar aviso Whatsapp", command=self.abrir_whatsapp_prestamo, bg="#e53935", fg="white").pack(side="left", padx=5)
        tk.Button(mid, text="VISTA INDIVIDUAL/COLECTIVA", command=self.toggle_prestamos_vista, bg="#bdbdbd", fg="black").pack(side="left", padx=5)

        self.lbl_perdon = tk.Label(frm, text="", fg="red", font=("", 10, "bold"))
        self.lbl_perdon.pack(fill="x", padx=5, pady=2)

        cols = ["codigo", "nombre", "apellidos", "email", "movil", "material", "fecha", "devuelto", "liberado_pin"]
        table = tk.Frame(frm)
        table.pack(fill="both", expand=True)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=cols, show="headings")
        for c in cols:
            heading = "Liberado por PIN" if c == "liberado_pin" else c.capitalize()
            tree.heading(c, text=heading)
            tree.column(c, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        vscroll = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        tree.tag_configure("naranja", background="#ffe6cc")
        tree.tag_configure("verde", background="#d4edda")
        tree.bind("<Control-c>", lambda e: self.copiar_celda(e, tree))
        tree.bind("<Double-1>", self.on_prestamo_doble_click)
        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Copiar", command=lambda: self.copiar_celda(tree.event_context, tree))
        tree.bind("<Button-3>", lambda e: self.mostrar_menu(e, menu, tree))
        self.tree_prestamos = tree
        tab.tree = tree

    def _norm(self, text):
        t = unicodedata.normalize("NFD", str(text or "")).upper().strip()
        return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")

    def _parse_fecha_prestamo(self, valor):
        try:
            return datetime.strptime(str(valor), "%d/%m/%Y %H:%M")
        except Exception:
            return None

    def _bring_to_front(self):
        try:
            self.lift()
            self.attributes("-topmost", True)
            self.update()
            self.attributes("-topmost", False)
        except Exception:
            pass

    def _normalizar_movil(self, movil):
        if not movil:
            return ""
        mov = "".join(ch for ch in str(movil) if ch.isdigit())
        # Quita prefijo 34 si viene duplicado y re-aplica una sola vez
        if mov.startswith("34"):
            mov = mov[2:]
        mov = mov.lstrip("0")
        if mov:
            mov = "34" + mov
        return mov

    def _pedir_campo(self, titulo, prompt, obligatorio=True, validar_nombre=False):
        self._bring_to_front()
        valor = simpledialog.askstring(titulo, prompt)
        if valor is None:
            return None
        valor = valor.strip()
        if obligatorio and not valor:
            messagebox.showwarning("Dato requerido", prompt)
            return self._pedir_campo(titulo, prompt, obligatorio, validar_nombre)
        if validar_nombre and " " in valor.strip():
            messagebox.showwarning("Formato nombre", "Solo nombre primero (una palabra).")
            return self._pedir_campo(titulo, prompt, obligatorio, validar_nombre)
        return valor

    def buscar_cliente_prestamo(self):
        codigo = self.prestamo_codigo.get().strip()
        if not codigo:
            messagebox.showwarning("Sin numero", "Introduce el numero de cliente.")
            return

        cliente = None
        fila = None

        if self.resumen_df is not None:
            colmap = {self._norm(c): c for c in self.resumen_df.columns}
            col_codigo = colmap.get("NUMERO DE CLIENTE") or colmap.get("NUMERO DE SOCIO")
            col_nom = colmap.get("NOMBRE")
            col_ape = colmap.get("APELLIDOS")
            col_mail = colmap.get("CORREO ELECTRONICO") or colmap.get("EMAIL") or colmap.get("CORREO")
            col_movil = colmap.get("MOVIL") or colmap.get("TELEFONO MOVIL") or colmap.get("NUMERO DE TELEFONO")

            if col_codigo:
                fila = self.resumen_df[self.resumen_df[col_codigo].astype(str).str.strip() == codigo]
                if not fila.empty:
                    row = fila.iloc[0]
                    cliente = {
                        "codigo": codigo,
                        "nombre": row.get(col_nom, ""),
                        "apellidos": row.get(col_ape, ""),
                        "email": row.get(col_mail, ""),
                        "movil": self._normalizar_movil(row.get(col_movil, ""))
                    }
            else:
                messagebox.showwarning("Columna faltante", "No se encontro la columna de numero de cliente.")

        # Si no se encontró en resumen, buscar/crear manual
        if not cliente:
            ext = [c for c in self.clientes_ext if c.get("codigo") == codigo]
            if ext:
                cliente = ext[0].copy()
                cliente["movil"] = self._normalizar_movil(cliente.get("movil", ""))
            else:
                if not messagebox.askyesno("Cliente no encontrado", f"No hay cliente {codigo}. Registrar manualmente?"):
                    return
                nombre = self._pedir_campo("Nombre", "Introduce el nombre:", obligatorio=True, validar_nombre=True)
                if nombre is None:
                    return
                apellidos = self._pedir_campo("Apellidos", "Introduce los apellidos:")
                if apellidos is None:
                    return
                email = self._pedir_campo("Email", "Introduce el email:")
                if email is None:
                    return
                movil = self._pedir_campo("Movil", "Introduce el movil (opcional):", obligatorio=False) or ""
                movil = self._normalizar_movil(movil)
                cliente = {
                    "codigo": codigo,
                    "nombre": nombre.strip(),
                    "apellidos": apellidos.strip(),
                    "email": email.strip(),
                    "movil": movil.strip(),
                }
                self.clientes_ext = [c for c in self.clientes_ext if c.get("codigo") != codigo] + [cliente]
                self.guardar_clientes_ext()

        if not cliente:
            return

        self.prestamo_encontrado = cliente
        info_text = f"{cliente.get('nombre','')} {cliente.get('apellidos','')} | {cliente.get('email','')} | {cliente.get('movil','')}"
        self.lbl_perdon.config(text="")

        pendientes = [p for p in self.prestamos if p.get("codigo") == codigo and not p.get("devuelto")]
        if pendientes:
            materiales = ", ".join(p.get("material", "") for p in pendientes if p.get("material"))
            info_text += f" | Pendiente: {materiales or 'material sin devolver'}"
            messagebox.showwarning("Material pendiente", f"El cliente {codigo} tiene material sin devolver: {materiales}")
            if len(pendientes) >= 2:
                self._resolver_pendientes_multples(pendientes, codigo)

        # Aviso si fue liberado por PIN anteriormente
        perdonado = any(p.get("codigo") == codigo and p.get("liberado_pin") for p in self.prestamos)
        if perdonado:
            self.lbl_perdon.config(text="ATENCIÓN, ESTE CLIENTE HA SIDO PERDONADO UNA VEZ")

        self.lbl_info.config(text=info_text)

    def _resolver_pendientes_multples(self, pendientes, codigo):
        self._bring_to_front()
        # Pregunta si quiere marcar uno como devuelto
        if not messagebox.askyesno(
            "Pendientes multiples",
            f"El cliente {codigo} tiene {len(pendientes)} préstamos sin devolver.\nQuieres marcar uno como devuelto ahora?"
        ):
            return

        # Ventana de selección simple
        win = tk.Toplevel(self)
        win.title("Selecciona préstamo a marcar devuelto")
        tk.Label(win, text="Selecciona el préstamo devuelto:").pack(padx=10, pady=5)
        listbox = tk.Listbox(win, height=min(10, len(pendientes)), exportselection=False)
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        # Mapea índice a id
        ids = []
        for p in pendientes:
            resumen = f"{p.get('fecha','')} | {p.get('material','')}"
            listbox.insert(tk.END, resumen)
            ids.append(p.get("id"))

        def confirmar():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("Sin seleccion", "Elige un préstamo.")
                return
            idx = sel[0]
            prestamo_id = ids[idx]
            for p in self.prestamos:
                if p.get("id") == prestamo_id:
                    p["devuelto"] = True
                    break
            self.guardar_prestamos_json()
            self.refrescar_prestamos_tree()
            win.destroy()

        tk.Button(win, text="Marcar devuelto", command=confirmar).pack(pady=5)
        tk.Button(win, text="Cerrar", command=win.destroy).pack(pady=2)

    def toggle_prestamos_vista(self):
        self.prestamos_filtro_activo = not self.prestamos_filtro_activo
        if self.prestamos_filtro_activo:
            codigo = ""
            if getattr(self, "prestamo_encontrado", None):
                codigo = self.prestamo_encontrado.get("codigo", "")
            else:
                codigo = self.prestamo_codigo.get().strip()
            if not codigo:
                messagebox.showwarning("Sin cliente", "Busca un cliente para ver su vista individual.")
                self.prestamos_filtro_activo = False
                return
        self.refrescar_prestamos_tree()

    def on_prestamo_doble_click(self, event):
        tree = self.tree_prestamos
        item = tree.identify_row(event.y)
        if not item:
            return
        valores = tree.item(item)["values"]
        if not valores:
            return
        codigo = str(valores[0]).strip()
        if not codigo:
            return
        self.prestamo_codigo.delete(0, tk.END)
        self.prestamo_codigo.insert(0, codigo)
        self.buscar_cliente_prestamo()

    def editar_cliente_manual(self):
        """
        Permite editar los datos de un cliente manual (solo los que viven en clientes_ext).
        Requiere tener un cliente buscado y que exista en la base manual.
        """
        if not getattr(self, "prestamo_encontrado", None):
            messagebox.showwarning("Sin cliente", "Busca un cliente primero.")
            return
        codigo = self.prestamo_encontrado.get("codigo")
        ext = [c for c in self.clientes_ext if c.get("codigo") == codigo]
        if not ext:
            messagebox.showwarning("No editable", "Este cliente proviene de RESUMEN CLIENTE y no se puede modificar.")
            return
        cliente = ext[0]
        nombre = self._pedir_campo("Nombre", "Introduce el nombre:", obligatorio=True, validar_nombre=True)
        if nombre is None:
            return
        apellidos = self._pedir_campo("Apellidos", "Introduce los apellidos:")
        if apellidos is None:
            return
        email = self._pedir_campo("Email", "Introduce el email:")
        if email is None:
            return
        movil = self._pedir_campo("Movil", "Introduce el movil (opcional):", obligatorio=False) or ""
        movil = self._normalizar_movil(movil)

        cliente.update({
            "nombre": nombre.strip(),
            "apellidos": apellidos.strip(),
            "email": email.strip(),
            "movil": movil.strip()
        })
        self.guardar_clientes_ext()
        # Actualiza en memoria por si se sigue usando
        self.prestamo_encontrado = cliente.copy()
        self.lbl_info.config(text=f"{cliente.get('nombre','')} {cliente.get('apellidos','')} | {cliente.get('email','')} | {cliente.get('movil','')}")
        messagebox.showinfo("Actualizado", "Datos del cliente manual actualizados.")

    def agregar_cliente_otro_fp(self):
        """
        Alta directa de cliente externo (otro fitness park) sin buscar primero.
        """
        self._bring_to_front()
        codigo = self._pedir_campo("Numero de cliente", "Introduce el numero de cliente:")
        if codigo is None or not codigo.strip():
            return
        codigo = codigo.strip()

        # Si está en resumen, no permitir
        if self.resumen_df is not None:
            colmap = {self._norm(c): c for c in self.resumen_df.columns}
            col_codigo = colmap.get("NUMERO DE CLIENTE") or colmap.get("NUMERO DE SOCIO")
            if col_codigo:
                if any(str(val).strip() == codigo for val in self.resumen_df[col_codigo].fillna("")):
                    messagebox.showwarning("No permitido", "Este cliente ya existe en RESUMEN CLIENTE y no se puede agregar manualmente.")
                    return

        nombre = self._pedir_campo("Nombre", "Introduce el nombre:", obligatorio=True, validar_nombre=True)
        if nombre is None:
            return
        apellidos = self._pedir_campo("Apellidos", "Introduce los apellidos:")
        if apellidos is None:
            return
        email = self._pedir_campo("Email", "Introduce el email:")
        if email is None:
            return
        movil = self._pedir_campo("Movil", "Introduce el movil (opcional):", obligatorio=False) or ""
        movil = self._normalizar_movil(movil)

        cliente = {
            "codigo": codigo,
            "nombre": nombre.strip(),
            "apellidos": apellidos.strip(),
            "email": email.strip(),
            "movil": movil.strip(),
        }
        self.clientes_ext = [c for c in self.clientes_ext if c.get("codigo") != codigo] + [cliente]
        self.guardar_clientes_ext()
        self.prestamo_encontrado = cliente.copy()
        self.lbl_info.config(text=f"{cliente.get('nombre','')} {cliente.get('apellidos','')} | {cliente.get('email','')} | {cliente.get('movil','')}")
        self.lbl_perdon.config(text="")
        messagebox.showinfo("Cliente agregado", f"Cliente {codigo} añadido como externo.")

    def guardar_prestamo(self):
        if not getattr(self, "prestamo_encontrado", None):
            messagebox.showwarning("Sin cliente", "Busca un cliente primero.")
            return
        material = self.prestamo_material.get().strip()
        if not material:
            messagebox.showwarning("Sin material", "Escribe el material prestado.")
            return
        # Evitar que peguen un número de cliente en el campo de material
        if re.fullmatch(r"[cC]u\d+", material):
            messagebox.showwarning(
                "Material inválido",
                "Has puesto un número de cliente en 'Material prestado'. Vuelve a intentarlo."
            )
            return

        codigo = self.prestamo_encontrado.get("codigo")
        pendientes = [p for p in self.prestamos if p.get("codigo") == codigo and not p.get("devuelto")]

        # Bloqueo a partir de 2 pendientes
        if len(pendientes) >= 2:
            resp = messagebox.askyesno(
                "Limite de prestamos",
                f"El cliente {codigo} ya tiene {len(pendientes)} prestamos sin devolver.\n"
                "Introducir codigo de seguridad para liberar y permitir otro prestamo?"
            )
            if not resp:
                return
            self._bring_to_front()
            pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
            if pin is None or pin.strip() != get_security_code():
                messagebox.showerror("Código incorrecto", "No se liberaron los préstamos pendientes.")
                return
            for p in pendientes:
                p["devuelto"] = True
                p["liberado_pin"] = True
            self.guardar_prestamos_json()
            self.refrescar_prestamos_tree()

        prestamo = {
            "id": uuid.uuid4().hex,
            **self.prestamo_encontrado,
            "material": material,
            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "devuelto": False,
            "liberado_pin": False
        }

        self.prestamos.append(prestamo)
        self.guardar_prestamos_json()
        self.refrescar_prestamos_tree()
        self.prestamo_material.delete(0, tk.END)
        self.prestamo_codigo.delete(0, tk.END)
        self.prestamo_encontrado = None
        self.lbl_info.config(text="")
        self.lbl_perdon.config(text="")
        messagebox.showinfo("Guardado", "Prestamo registrado.")

    def marcar_devuelto(self):
        sel = self.tree_prestamos.selection()
        if not sel:
            messagebox.showwarning("Sin seleccion", "Selecciona un prestamo.")
            return
        iid = sel[0]
        seleccionado = None
        for p in self.prestamos:
            if p.get("id") == iid:
                seleccionado = p
                break
        if not seleccionado:
            vals = self.tree_prestamos.item(sel[0])["values"]
            if len(vals) >= 7:
                codigo_sel, material_sel, fecha_sel = vals[0], vals[5], vals[6]
                for p in self.prestamos:
                    if p.get("codigo") == codigo_sel and p.get("material") == material_sel and p.get("fecha") == fecha_sel:
                        seleccionado = p
                        break
        if not seleccionado:
            messagebox.showwarning("No encontrado", "No se pudo identificar el préstamo seleccionado.")
            return
        seleccionado["devuelto"] = not seleccionado.get("devuelto", False)
        # Si se marcó manualmente, no tocar liberado_pin (permite conservar histórico)
        self.guardar_prestamos_json()
        self.refrescar_prestamos_tree()

    def enviar_aviso_prestamo(self):
        sel = self.tree_prestamos.selection()
        if not sel:
            messagebox.showwarning("Sin seleccion", "Selecciona un prestamo.")
            return
        datos = self.tree_prestamos.item(sel[0])["values"]
        email = datos[3]
        nombre = datos[1]
        material = datos[5]

        cuerpo = (
            "Advertencia 1 - Material de prestamo no devuelto\n\n"
            f"Hola {nombre},\n\n"
            f"No has devuelto el material prestado ({material}). "
            "En caso de que vuelva a ocurrir, no podremos ofrecer este servicio. Gracias por colaborar."
        )

        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            messagebox.showwarning("Outlook no disponible", "Cuerpo copiado al portapapeles.")
            return

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "Aviso material prestamo"
            mail.Body = cuerpo
            mail.Display()
        except Exception as e:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            messagebox.showerror("Error con Outlook", f"No se pudo crear el correo.\nDetalle: {e}")

    def abrir_whatsapp_prestamo(self):
        sel = self.tree_prestamos.selection()
        if not sel:
            messagebox.showwarning("Sin seleccion", "Selecciona un prestamo.")
            return
        datos = self.tree_prestamos.item(sel[0])["values"]
        movil = str(datos[4]).replace(" ", "")
        if not movil:
            messagebox.showwarning("Sin movil", "El cliente no tiene movil registrado.")
            return
        nombre = datos[1]
        material = datos[5]
        texto = (
            "Advertencia 1 - Material de prestamo no devuelto\n\n"
            f"Hola {nombre},\n\n"
            f"No has devuelto el material prestado ({material}). "
            "En caso de que vuelva a ocurrir, no podremos ofrecer este servicio. Gracias por colaborar."
        )
        try:
            url = f"https://wa.me/{movil}utext={urllib.parse.quote(texto)}"
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("WhatsApp", f"No se pudo abrir el enlace: {e}")

    def cargar_prestamos_json(self):
        try:
            if os.path.exists(self.prestamos_file):
                with open(self.prestamos_file, "r", encoding="utf-8") as f:
                    self.prestamos = json.load(f)
            # Asegura IDs para prestamos antiguos
            changed = False
            for p in self.prestamos:
                if "id" not in p:
                    p["id"] = uuid.uuid4().hex
                    changed = True
                if "liberado_pin" not in p:
                    p["liberado_pin"] = False
                    changed = True
            if changed:
                self.guardar_prestamos_json()
        except Exception:
            self.prestamos = []
        self.refrescar_prestamos_tree()

    def cargar_clientes_ext(self):
        try:
            if os.path.exists(self.clientes_ext_file):
                with open(self.clientes_ext_file, "r", encoding="utf-8") as f:
                    self.clientes_ext = json.load(f)
        except Exception:
            self.clientes_ext = []

    def cargar_felicitaciones(self):
        try:
            if os.path.exists(self.felicitaciones_file):
                with open(self.felicitaciones_file, "r", encoding="utf-8") as f:
                    self.felicitaciones_enviadas = json.load(f)
        except Exception:
            self.felicitaciones_enviadas = {}

    def guardar_prestamos_json(self):
        try:
            os.makedirs(os.path.dirname(self.prestamos_file), exist_ok=True)
            with open(self.prestamos_file, "w", encoding="utf-8") as f:
                json.dump(self.prestamos, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # silencioso para no romper UI

    def guardar_clientes_ext(self):
        try:
            os.makedirs(os.path.dirname(self.clientes_ext_file), exist_ok=True)
            with open(self.clientes_ext_file, "w", encoding="utf-8") as f:
                json.dump(self.clientes_ext, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def guardar_felicitaciones(self):
        try:
            os.makedirs(os.path.dirname(self.felicitaciones_file), exist_ok=True)
            with open(self.felicitaciones_file, "w", encoding="utf-8") as f:
                json.dump(self.felicitaciones_enviadas, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def refrescar_prestamos_tree(self):
        if not hasattr(self, "tree_prestamos"):
            return
        self.tree_prestamos.delete(*self.tree_prestamos.get_children())
        # Ordenar de más reciente a más antiguo por fecha
        ordenados = sorted(
            self.prestamos,
            key=lambda p: self._parse_fecha_prestamo(p.get("fecha")) or datetime.min,
            reverse=True
        )
        codigo_filtro = None
        if self.prestamos_filtro_activo:
            if getattr(self, "prestamo_encontrado", None):
                codigo_filtro = self.prestamo_encontrado.get("codigo")
            else:
                codigo_filtro = self.prestamo_codigo.get().strip()
        for p in ordenados:
            if codigo_filtro and p.get("codigo") != codigo_filtro:
                continue
            iid = p.get("id") or uuid.uuid4().hex
            p["id"] = iid
            tag = "verde" if p.get("devuelto") else "naranja"
            self.tree_prestamos.insert("", "end", values=[
                p.get("codigo", ""), p.get("nombre", ""), p.get("apellidos", ""),
                p.get("email", ""), p.get("movil", ""), p.get("material", ""),
                p.get("fecha", ""), "SI" if p.get("devuelto") else "NO",
                "SI" if p.get("liberado_pin") else "NO"
            ], tags=(tag,), iid=iid)

    # -----------------------------
    # PESTAÑA IMPAGOS
    # -----------------------------
    def create_impagos_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        top = tk.Frame(frm)
        top.pack(fill="x", pady=4)
        tk.Button(top, text="Actualizar desde CSV", command=self.refrescar_impagos_desde_csv).pack(side="left", padx=5)
        tk.Label(top, text="Vista:").pack(side="left", padx=5)
        tk.Button(top, text="Deudores actuales", command=lambda: self.impagos_set_view("actuales")).pack(side="left", padx=3)
        tk.Button(top, text="1 incidente", command=lambda: self.impagos_set_view("incidentes1")).pack(side="left", padx=3)
        tk.Button(top, text="2+ incidentes", command=lambda: self.impagos_set_view("incidentes2")).pack(side="left", padx=3)
        tk.Button(top, text="Resueltos", command=lambda: self.impagos_set_view("resueltos")).pack(side="left", padx=3)
        self.impagos_status = tk.Label(top, text="", anchor="w")
        self.impagos_status.pack(side="left", padx=10)

        actions = tk.Frame(frm)
        actions.pack(fill="x", pady=6)
        tk.Button(actions, text="Abrir WhatsApp", command=self.abrir_whatsapp_impagos, bg="#8bc34a", fg="black").pack(side="left", padx=5)
        tk.Button(actions, text="Copiar email", command=self.copiar_email_impagos).pack(side="left", padx=5)
        tk.Button(actions, text="Enviar email (1Inc)", command=lambda: self.enviar_email_impagos("1inc"), bg="#ffd54f", fg="black").pack(side="left", padx=5)
        tk.Button(actions, text="Enviar email (2+ inc)", command=lambda: self.enviar_email_impagos("2inc"), bg="#ff8a65", fg="black").pack(side="left", padx=5)
        tk.Button(actions, text="Enviar email resueltos", command=self.enviar_email_resueltos, bg="#81c784", fg="black").pack(side="left", padx=5)

        cols = ["codigo", "nombre", "apellidos", "email", "movil", "incidentes", "email_enviado", "fecha_envio", "historial_email", "reincidente", "fecha_export"]
        table = tk.Frame(frm)
        table.pack(fill="both", expand=True)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=cols, show="headings")
        for c in cols:
            if c == "fecha_export":
                heading = "Fecha export"
            elif c == "email_enviado":
                heading = "Email enviado"
            elif c == "fecha_envio":
                heading = "Fecha envio"
            elif c == "historial_email":
                heading = "Historial emails"
            elif c == "reincidente":
                heading = "Reincidente"
            else:
                heading = c.capitalize()
            tree.heading(c, text=heading, command=lambda col=c: self.sort_column(tree, col, False))
            tree.column(c, anchor="center")
        col_widths = {
            "codigo": 90,
            "nombre": 100,
            "apellidos": 130,
            "email": 190,
            "movil": 90,
            "incidentes": 70,
            "email_enviado": 90,
            "fecha_envio": 90,
            "historial_email": 150,
            "reincidente": 90,
            "fecha_export": 90,
        }
        for c, w in col_widths.items():
            tree.column(c, width=w, stretch=False)
        # Deja la última columna expansible para ocupar el ancho disponible
        tree.column("fecha_export", stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")
        vscroll = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        tree.bind("<Control-c>", lambda e: self.copiar_celda(e, tree))
        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Copiar", command=lambda: self.copiar_celda(tree.event_context, tree))
        tree.bind("<Button-3>", lambda e: self.mostrar_menu(e, menu, tree))
        self.tree_impagos = tree
        tab.tree = tree

    def refrescar_impagos_desde_csv(self):
        if self.folder_path:
            try:
                incidencias = load_data_file(self.folder_path, "IMPAGOS")
                self.sync_impagos(incidencias)
            except Exception as e:
                messagebox.showerror("Impagos", f"No se pudo cargar IMPAGOS.csv: {e}")
        else:
            messagebox.showwarning("Sin carpeta", "Selecciona primero la carpeta de datos.")

    def sync_impagos(self, df):
        try:
            resumen_map = None
            if self.resumen_df is not None:
                cols = {self._norm(c): c for c in self.resumen_df.columns}
                col_codigo = cols.get("NUMERO DE CLIENTE") or cols.get("NUMERO DE SOCIO")
                col_nombre = cols.get("NOMBRE")
                col_apellidos = cols.get("APELLIDOS")
                col_email = cols.get("CORREO ELECTRONICO") or cols.get("EMAIL") or cols.get("CORREO")
                col_movil = cols.get("MOVIL") or cols.get("TELEFONO") or cols.get("TELEFONO MOVIL")
                if col_codigo:
                    resumen_map = {}
                    for _, row in self.resumen_df.iterrows():
                        codigo = str(row.get(col_codigo, "")).strip()
                        if not codigo:
                            continue
                        resumen_map[codigo] = {
                            "nombre": str(row.get(col_nombre, "")).strip(),
                            "apellidos": str(row.get(col_apellidos, "")).strip(),
                            "email": str(row.get(col_email, "")).strip(),
                            "movil": str(row.get(col_movil, "")).strip(),
                        }
            fecha, count = self.impagos_db.sync_from_df(df, resumen_map=resumen_map)
            self.impagos_last_export = fecha
            self.impagos_status.config(text=f"Export: {fecha} | Registros: {count}")
            self.refresh_impagos_view()
        except Exception as e:
            messagebox.showerror("Impagos", f"Error sincronizando impagos: {e}")

    def impagos_set_view(self, view_name):
        self.impagos_view.set(view_name)
        self.refresh_impagos_view()

    def refresh_impagos_view(self):
        if not hasattr(self, "tree_impagos"):
            return
        if not self.impagos_last_export:
            self.impagos_last_export = self.impagos_db.get_last_export()
        rows = self.impagos_db.fetch_view(self.impagos_view.get(), self.impagos_last_export)
        self.tree_impagos.delete(*self.tree_impagos.get_children())
        if self.impagos_view.get() == "resueltos":
            self.tree_impagos["displaycolumns"] = ("codigo", "nombre", "apellidos", "email", "movil")
        elif self.impagos_view.get() == "actuales":
            self.tree_impagos["displaycolumns"] = (
                "codigo", "nombre", "apellidos", "email", "movil",
                "incidentes", "email_enviado", "historial_email"
            )
        else:
            self.tree_impagos["displaycolumns"] = (
                "codigo", "nombre", "apellidos", "email", "movil",
                "incidentes", "email_enviado", "historial_email", "fecha_export"
            )
        for r in rows:
            values = list(r)
            # Normaliza checks para email/reincidente
            if len(values) >= 11:
                values[7] = values[7] or ""
                values[6] = "SI" if values[6] else "NO"
                values[9] = "SI" if values[9] else "NO"
                values[8] = values[8] or ""
            self.tree_impagos.insert("", "end", values=values, iid=str(r[0]))

    def _get_impagos_selected(self):
        sel = self.tree_impagos.selection()
        if not sel:
            messagebox.showwarning("Sin seleccion", "Selecciona un cliente.")
            return None
        values = self.tree_impagos.item(sel[0])["values"]
        return {
            "codigo": values[0],
            "nombre": values[1],
            "apellidos": values[2],
            "email": values[3],
            "movil": values[4],
            "incidentes": values[5],
            "email_enviado": values[6] if len(values) > 6 else "",
            "fecha_envio": values[7] if len(values) > 7 else "",
            "historial_email": values[8] if len(values) > 8 else "",
            "reincidente": values[9] if len(values) > 9 else "",
            "fecha_export": values[10] if len(values) > 10 else "",
        }

    def abrir_whatsapp_impagos(self):
        data = self._get_impagos_selected()
        if not data:
            return
        movil = self._normalizar_movil(data.get("movil", ""))
        if not movil:
            messagebox.showwarning("Sin movil", "El cliente no tiene movil registrado.")
            return
        try:
            url = f"https://wa.me/{movil}"
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("WhatsApp", f"No se pudo abrir el enlace: {e}")

    def copiar_email_impagos(self):
        data = self._get_impagos_selected()
        if not data:
            return
        email = data.get("email", "").strip()
        if not email:
            messagebox.showwarning("Sin email", "El cliente no tiene email registrado.")
            return
        self.clipboard_clear()
        self.clipboard_append(email)
        messagebox.showinfo("Copiado", "Email copiado al portapapeles.")

    def _impagos_registrar(self, accion, plantilla=""):
        data = self._get_impagos_selected()
        if not data:
            return
        cliente_id = self.impagos_db.get_cliente_id(data["codigo"])
        if not cliente_id:
            messagebox.showerror("Impagos", "Cliente no encontrado en la base.")
            return
        self.impagos_db.add_gestion(cliente_id, accion, plantilla, "")
        messagebox.showinfo("Registrado", f"Accion registrada: {accion}.")
        self.refresh_impagos_view()

    def enviar_email_impagos(self, plantilla):
        # Recolecta emails desde la vista actual según incidentes
        rows = [self.tree_impagos.item(i)["values"] for i in self.tree_impagos.get_children()]
        if plantilla == "1inc":
            rows = [r for r in rows if int(r[5]) == 1]
        else:
            rows = [r for r in rows if int(r[5]) >= 2]
        emails = [str(r[3]).strip() for r in rows if str(r[3]).strip()]

        cuerpo_html = self._impagos_email_html(plantilla)
        asunto = "Aviso de impago"
        imagen_path = get_logo_path("PAGADEUDA.png")

        try:
            import win32com.client  # type: ignore
        except ImportError:
            messagebox.showwarning("Outlook no disponible", "No se pudo importar win32com.client.")
            return

        def try_outlook():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            try:
                ns = outlook_app.GetNamespace("MAPI")
                ns.Logon("", "", False, False)
            except Exception:
                pass
            return outlook_app

        try:
            try:
                outlook = try_outlook()
            except Exception:
                os.startfile("outlook.exe")
                time.sleep(3)
                outlook = try_outlook()

            mail = outlook.CreateItem(0)
            if emails:
                mail.BCC = ";".join(sorted(set(emails)))
            mail.Subject = asunto
            cid = uuid.uuid4().hex
            if os.path.exists(imagen_path):
                attachment = mail.Attachments.Add(imagen_path, 1, 0)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid
                )
                mail.HTMLBody = cuerpo_html.replace("{{CID}}", cid)
            else:
                mail.HTMLBody = cuerpo_html.replace('<img src="cid:{{CID}}" alt="Paga deuda" style="max-width:100%;">', "")

            mail.Display()
        except Exception as e:
            messagebox.showerror("Outlook", f"No se pudo crear el correo: {e}", parent=self)
            return

        if not messagebox.askyesno("Confirmacion", "Ha enviado el email?", parent=self):
            return
        if not emails:
            return

        # Registrar gestión para todos los clientes de la lista
        for r in rows:
            codigo = str(r[0]).strip()
            cliente_id = self.impagos_db.get_cliente_id(codigo)
            if cliente_id:
                self.impagos_db.add_gestion(cliente_id, "email", plantilla, "")
        self.refresh_impagos_view()

    def enviar_email_resueltos(self):
        if self.impagos_view.get() != "resueltos":
            messagebox.showwarning("Impagos", "Selecciona la vista Resueltos para enviar este email.")
            return
        rows = [self.tree_impagos.item(i)["values"] for i in self.tree_impagos.get_children()]
        emails = [str(r[3]).strip() for r in rows if str(r[3]).strip()]
        asunto = "TODO EN ORDEN"
        imagen_path = get_logo_path("PAGOHECHO.png")
        html = self._impagos_email_imagen_html()

        try:
            import win32com.client  # type: ignore
        except ImportError:
            messagebox.showwarning("Outlook no disponible", "No se pudo importar win32com.client.")
            return

        def try_outlook():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            try:
                ns = outlook_app.GetNamespace("MAPI")
                ns.Logon("", "", False, False)
            except Exception:
                pass
            return outlook_app

        try:
            try:
                outlook = try_outlook()
            except Exception:
                os.startfile("outlook.exe")
                time.sleep(3)
                outlook = try_outlook()

            mail = outlook.CreateItem(0)
            if emails:
                mail.BCC = ";".join(sorted(set(emails)))
            mail.Subject = asunto
            cid = uuid.uuid4().hex
            if os.path.exists(imagen_path):
                attachment = mail.Attachments.Add(imagen_path, 1, 0)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid
                )
                mail.HTMLBody = html.replace("{{CID}}", cid)
            else:
                mail.HTMLBody = html.replace('<img src="cid:{{CID}}" alt="Pago hecho" style="max-width:100%;">', "")

            mail.Display()
        except Exception as e:
            messagebox.showerror("Outlook", f"No se pudo crear el correo: {e}", parent=self)
            return

        if not messagebox.askyesno("Confirmacion", "Ha enviado el email de resueltos?", parent=self):
            return
        if not emails:
            return
        for r in rows:
            codigo = str(r[0]).strip()
            cliente_id = self.impagos_db.get_cliente_id(codigo)
            if cliente_id:
                self.impagos_db.add_gestion(cliente_id, "resuelto_email", "resuelto", "")
        self.refresh_impagos_view()

    def _impagos_email_html(self, plantilla):
        wa = "https://wa.me/34681872664"
        wa_link = f'<a href="{wa}">{wa}</a>'
        if plantilla == "1inc":
            texto = "<br>".join(
                [
                    "Hola! Tienes un recibo pendiente de pago y el torno no permitira el acceso.",
                    "Si vienes de 6 a 9 am (horario sin atencion comercial), te recomendamos pasar por recepcion a partir de las 9 am para solucionarlo lo antes posible.",
                    "Puedes abonar en efectivo o con tarjeta, incluso abonarlo desde la propia aplicacion. Aqui te mostramos como mas abajo.",
                    "Recuerda: puedes anticipar el pago entre 5 y 15 dias antes de tu siguiente dia de pago que siempre podras consultar en tu APP Fitness Park.",
                    f"Dudasu PREGUNTANOS al whatsapp ({wa_link}) o contestando este email.",
                ]
            )
        else:
            texto = "<br>".join(
                [
                    "Hemos detectado que actualmente tienes 2 recibos o mas pendientes en tu cuenta de socio/a.",
                    "Queremos ayudarte a regularizar tu situacion lo antes posible para que puedas seguir disfrutando de todas las instalaciones sin inconvenientes, evitando que el sistema derive tu caso a la empresa de recobros PayPymes.",
                    "Opciones para abonar:",
                    "Tarjeta o efectivo en recepcion.",
                    "Directamente desde tu APP Fitness Park, en 'Espacio de cliente' -> 'Pagos'. (Te mostramos como mas abajo)",
                    "Aviso importante: En caso de que se genere un tercer recibo impagado, desde el club dejaremos de emitir advertencias y tu expediente el sistema lo derivara directamente a nuestra empresa colaboradora de recobros Paypymes, que se pondra en contacto contigo para gestionar la deuda.",
                    "Recuerda: puedes anticipar el pago entre 5 y 15 dias antes de tu siguiente dia de pago que siempre podras consultar en tu APP Fitness Park, en el icono espacio de cliente y luego en el apartado 'PAGOS'.",
                    "Te animamos a ponerte al dia lo antes posible para evitar cualquier gestion externa y que todo siga con normalidad. Estamos a tu disposicion en recepcion para cualquier duda.",
                    f"Dudasu PREGUNTANOS por whatsapp ({wa_link}) o contestando este email.",
                ]
            )
        return (
            '<div style="font-family:Arial,sans-serif;font-size:14px;">'
            f"{texto}<br><br>"
            '<img src="cid:{{CID}}" alt="Paga deuda" style="max-width:100%;">'
            "</div>"
        )

    def _impagos_email_imagen_html(self):
        return (
            '<div style="font-family:Arial,sans-serif;font-size:14px;">'
            '<img src="cid:{{CID}}" alt="Pago hecho" style="max-width:100%;">'
            '</div>'
        )

    def ir_a_impagos(self):
        pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
        if pin is None:
            return
        if pin.strip() != get_security_code():
            messagebox.showerror("Codigo incorrecto", "El codigo de seguridad no es valido.", parent=self)
            return
        tab = self.tabs.get("Impagos")
        if tab:
            self.notebook.add(tab, text="Impagos", image=self.tab_icons.get("Impagos"), compound="left")
            self.notebook.select(tab)

    def on_tab_changed(self, event):
        tab = self.tabs.get("Impagos")
        if not tab:
            return
        current = self.notebook.select()
        if current != str(tab):
            try:
                self.notebook.hide(tab)
            except Exception:
                pass

    # -----------------------------
    # INCIDENCIAS CLUB
    # -----------------------------
    def create_incidencias_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        botones = tk.Frame(frm)
        botones.pack(fill="x", pady=4)
        tk.Button(botones, text="Cargar Mapas", command=self.incidencias_cargar_mapas).pack(side="left", padx=5)
        tk.Button(botones, text="Borrar Mapa", command=self.incidencias_borrar_mapa).pack(side="left", padx=5)
        self.btn_mapa_anterior = tk.Button(botones, text="Mapa anterior", command=self.incidencias_mapa_anterior)
        self.btn_mapa_anterior.pack(side="left", padx=5)
        self.btn_mapa_siguiente = tk.Button(botones, text="Mapas", command=self.incidencias_siguiente_mapa)
        self.btn_mapa_siguiente.pack(side="left", padx=5)
        tk.Button(botones, text="Asignar Área", command=self.incidencias_asignar_area).pack(side="left", padx=5)
        tk.Button(botones, text="Editar Área", command=self.incidencias_editar_area).pack(side="left", padx=5)
        tk.Button(botones, text="Listar Máquina", command=self.incidencias_listar_maquina).pack(side="left", padx=5)
        tk.Button(botones, text="Editar Máquina", command=self.incidencias_editar_maquina).pack(side="left", padx=5)
        tk.Button(botones, text="Info máquinas", command=self.incidencias_info_maquinas).pack(side="left", padx=5)
        tk.Button(botones, text="Crear Incidencia", command=self.incidencias_crear_incidencia).pack(side="left", padx=5)
        tk.Button(botones, text="Gestión Incidencias", command=self.incidencias_gestion_incidencias).pack(side="left", padx=5)

        body = tk.Frame(frm)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        self.incidencias_panel = tk.Frame(body, bd=1, relief="solid")
        self.incidencias_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        canvas_container = tk.Frame(body)
        canvas_container.grid(row=0, column=1, sticky="nsew")
        canvas_container.rowconfigure(1, weight=1)
        canvas_container.columnconfigure(0, weight=1)

        header = tk.Frame(canvas_container)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        self.incidencias_area_title = tk.Label(header, text="", font=("Arial", 11, "bold"))
        self.incidencias_area_title.pack(side="left", padx=6)
        self.incidencias_btn_vista_general = tk.Button(header, text="VISTA GENERAL", command=self.incidencias_vista_general)
        self.incidencias_btn_vista_general.configure(state="disabled")

        self.incidencias_canvas = tk.Canvas(canvas_container, bg="white")
        self.incidencias_canvas.grid(row=1, column=0, sticky="nsew")
        canvas_y = ttk.Scrollbar(canvas_container, orient="vertical", command=self.incidencias_canvas.yview)
        canvas_y.grid(row=1, column=1, sticky="ns")
        canvas_x = ttk.Scrollbar(canvas_container, orient="horizontal", command=self.incidencias_canvas.xview)
        canvas_x.grid(row=2, column=0, sticky="ew")
        self.incidencias_canvas.configure(yscrollcommand=canvas_y.set, xscrollcommand=canvas_x.set)
        self.incidencias_canvas.bind("<ButtonPress-1>", self.incidencias_canvas_press)
        self.incidencias_canvas.bind("<B1-Motion>", self.incidencias_canvas_drag)
        self.incidencias_canvas.bind("<ButtonRelease-1>", self.incidencias_canvas_release)
        self.incidencias_canvas.bind("<Motion>", self.incidencias_canvas_hover)
        self.incidencias_canvas.bind("<Leave>", self.incidencias_canvas_leave)

        self.incidencias_cargar_listado_mapas()
        self.incidencias_gestion_incidencias()

    def _incidencias_pin_ok(self):
        self._bring_to_front()
        pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
        if pin is None:
            return False
        if pin.strip() != get_security_code():
            messagebox.showerror("Codigo incorrecto", "El codigo de seguridad no es valido.", parent=self)
            return False
        return True

    def _incidencias_color(self):
        palette = [
            "#e57373", "#f06292", "#ba68c8", "#9575cd", "#7986cb", "#64b5f6",
            "#4fc3f7", "#4dd0e1", "#4db6ac", "#81c784", "#aed581", "#dce775",
            "#fff176", "#ffd54f", "#ffb74d", "#ff8a65",
        ]
        for _ in range(100):
            color = random.choice(palette)
            if color not in self.incidencias_color_used:
                self.incidencias_color_used.add(color)
                return color
        # Fallback aleatorio
        while True:
            color = f"#{random.randint(0, 0xFFFFFF):06x}"
            if color not in self.incidencias_color_used:
                self.incidencias_color_used.add(color)
                return color
    def _incidencias_set_area_header(self, area_id):
        if not self.incidencias_area_title or not self.incidencias_btn_vista_general:
            return
        if not self.incidencias_current_map:
            self.incidencias_area_title.configure(text="")
            if self.incidencias_btn_vista_general.winfo_ismapped():
                self.incidencias_btn_vista_general.pack_forget()
            return
        if area_id:
            nombre = None
            for area in self.incidencias_db.list_areas(self.incidencias_current_map):
                if area[0] == area_id:
                    nombre = area[1]
                    break
            self.incidencias_area_title.configure(text=nombre or "")
            self.incidencias_btn_vista_general.configure(state="normal")
            if not self.incidencias_btn_vista_general.winfo_ismapped():
                self.incidencias_btn_vista_general.pack(side="left", padx=6)
        else:
            self.incidencias_area_title.configure(text="")
            self.incidencias_btn_vista_general.configure(state="disabled")
            if self.incidencias_btn_vista_general.winfo_ismapped():
                self.incidencias_btn_vista_general.pack_forget()

    def _incidencias_apply_map_filter(self, area_id):
        if not self.incidencias_current_map:
            return
        for aid, item in self.incidencias_area_items.items():
            if area_id:
                self.incidencias_canvas.itemconfig(item, state="hidden")
            else:
                self.incidencias_canvas.itemconfig(item, state="normal")
        for mid, item in self.incidencias_machine_items.items():
            if area_id and self.incidencias_machine_area.get(mid) != area_id:
                self.incidencias_canvas.itemconfig(item, state="hidden")
            else:
                self.incidencias_canvas.itemconfig(item, state="normal")

    def _incidencias_start_hover_blink(self, mid):
        if not mid or mid not in self.incidencias_machine_items:
            self._incidencias_stop_hover_blink()
            return
        item_id = self.incidencias_machine_items[mid]
        if self.incidencias_canvas.itemcget(item_id, "state") == "hidden":
            return
        if self.incidencias_hover_blink_item == item_id:
            return
        self._incidencias_stop_hover_blink()
        self.incidencias_hover_blink_item = item_id
        outline = self.incidencias_canvas.itemcget(item_id, "outline")
        width = int(float(self.incidencias_canvas.itemcget(item_id, "width") or 2))
        self.incidencias_hover_blink_original = (item_id, outline, width)

        def toggle(show):
            if self.incidencias_hover_blink_item != item_id:
                return
            if show:
                self.incidencias_canvas.itemconfig(item_id, outline=outline, width=max(width + 2, 4))
            else:
                self.incidencias_canvas.itemconfig(item_id, outline="", width=width)
            self.incidencias_hover_blink_after = self.after(350, lambda: toggle(not show))

        toggle(True)

    def _incidencias_stop_hover_blink(self):
        if self.incidencias_hover_blink_after:
            try:
                self.after_cancel(self.incidencias_hover_blink_after)
            except Exception:
                pass
        self.incidencias_hover_blink_after = None
        if self.incidencias_hover_blink_original:
            item_id, outline, width = self.incidencias_hover_blink_original
            if self.incidencias_hover_blink_item == item_id:
                self.incidencias_canvas.itemconfig(item_id, outline=outline, width=width)
        self.incidencias_hover_blink_item = None
        self.incidencias_hover_blink_original = None

    def incidencias_vista_general(self):
        self.incidencias_info_filter_area = None
        self._incidencias_set_area_header(None)
        self._incidencias_apply_map_filter(None)
        self.incidencias_info_maquinas(None)

    def _incidencias_pedir_reporte_visual(self):
        self._bring_to_front()
        if not messagebox.askyesno("Incidencia", "Tiene reporte visual?", parent=self):
            return ""
        ruta = filedialog.askopenfilename(
            parent=self,
            title="Selecciona reporte visual",
            filetypes=[
                ("Imagenes", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
                ("Video", "*.mp4;*.avi;*.mov;*.mkv;*.wmv"),
                ("Todos", "*.*"),
            ],
        )
        if not ruta:
            return ""
        try:
            reports_dir = os.path.join("data", "incidencias_reportes")
            os.makedirs(reports_dir, exist_ok=True)
            ext = os.path.splitext(ruta)[1]
            destino = os.path.join(reports_dir, f"reporte_{uuid.uuid4().hex}{ext}")
            shutil.copy2(ruta, destino)
            return destino
        except Exception:
            return ruta

    def _incidencias_ver_reporte(self, ruta):
        if not ruta or not os.path.exists(ruta):
            messagebox.showwarning("Reporte visual", "No se encontro el archivo del reporte.", parent=self)
            return
        try:
            os.startfile(ruta)
        except Exception as exc:
            messagebox.showerror("Reporte visual", f"No se pudo abrir el reporte.\nDetalle: {exc}", parent=self)

    def _incidencias_guardar_reporte(self, ruta):
        if not ruta or not os.path.exists(ruta):
            messagebox.showwarning("Reporte visual", "No se encontro el archivo del reporte.", parent=self)
            return
        filename = os.path.basename(ruta)
        destino = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar reporte visual",
            initialfile=filename,
            defaultextension=os.path.splitext(filename)[1],
        )
        if not destino:
            return
        try:
            shutil.copy2(ruta, destino)
            messagebox.showinfo("Reporte visual", f"Reporte guardado en: {destino}", parent=self)
        except Exception as exc:
            messagebox.showerror("Reporte visual", f"No se pudo guardar el reporte.\nDetalle: {exc}", parent=self)


    def incidencias_cargar_listado_mapas(self):
        self.incidencias_mapas = self.incidencias_db.list_maps()
        self.incidencias_map_index = 0
        if self.incidencias_mapas:
            self.incidencias_mostrar_mapa(self.incidencias_mapas[0])
        else:
            self.incidencias_canvas.delete("all")
            self.incidencias_current_map = None
        self.incidencias_actualizar_botones_mapa()

    def incidencias_actualizar_botones_mapa(self):
        if not hasattr(self, "btn_mapa_anterior"):
            return
        if len(self.incidencias_mapas) <= 1:
            self.btn_mapa_anterior.configure(state="disabled")
            self.btn_mapa_siguiente.configure(state="disabled")
        else:
            self.btn_mapa_anterior.configure(state="normal")
            self.btn_mapa_siguiente.configure(state="normal")

    def incidencias_cargar_mapas(self):
        if not self._incidencias_pin_ok():
            return
        self._bring_to_front()
        cantidad = simpledialog.askinteger("Mapas", "Cuantos mapas quieres cargar?", minvalue=1, maxvalue=20, parent=self)
        if not cantidad:
            return
        maps_dir = os.path.join("data", "maps")
        os.makedirs(maps_dir, exist_ok=True)
        for _ in range(cantidad):
            ruta = filedialog.askopenfilename(filetypes=[("PNG", "*.png")])
            if not ruta:
                continue
            nombre = os.path.basename(ruta)
            destino = os.path.join(maps_dir, nombre)
            shutil.copy2(ruta, destino)
            img = Image.open(destino)
            self.incidencias_db.add_map(nombre, destino, img.width, img.height)
        self.incidencias_cargar_listado_mapas()

    def incidencias_borrar_mapa(self):
        if not self._incidencias_pin_ok():
            return
        if not self.incidencias_mapas:
            self._bring_to_front()
            messagebox.showwarning("Mapas", "No hay mapas cargados.", parent=self)
            return
        mapa = self.incidencias_mapas[self.incidencias_map_index]
        mapa_id, nombre, ruta, *_rest = mapa
        self._bring_to_front()
        if not messagebox.askyesno("Borrar mapa", f"Borrar el mapa {nombre} y todo lo asociado?", parent=self):
            return
        try:
            if ruta and os.path.exists(ruta):
                os.remove(ruta)
        except Exception:
            pass
        self.incidencias_db.delete_map(mapa_id)
        self.incidencias_cargar_listado_mapas()
        if not self.incidencias_mapas:
            self.incidencias_current_map = None
            self.incidencias_canvas.delete("all")
            for w in self.incidencias_panel.winfo_children():
                w.destroy()

    def incidencias_siguiente_mapa(self):
        if not self.incidencias_mapas:
            self._bring_to_front()
            messagebox.showwarning("Mapas", "No hay mapas cargados.", parent=self)
            return
        self.incidencias_map_index = (self.incidencias_map_index + 1) % len(self.incidencias_mapas)
        self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])

    def incidencias_mapa_anterior(self):
        if not self.incidencias_mapas:
            self._bring_to_front()
            messagebox.showwarning("Mapas", "No hay mapas cargados.", parent=self)
            return
        self.incidencias_map_index = (self.incidencias_map_index - 1) % len(self.incidencias_mapas)
        self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])

    def incidencias_mostrar_mapa(self, mapa_row):
        mapa_id, nombre, ruta, _, _, _ = mapa_row
        self.incidencias_canvas.delete("all")
        self.incidencias_area_items = {}
        self.incidencias_machine_items = {}
        self.incidencias_machine_area = {}
        img = Image.open(ruta)
        self.incidencias_img = ImageTk.PhotoImage(img)
        self.incidencias_canvas.create_image(0, 0, anchor="nw", image=self.incidencias_img)
        self.incidencias_canvas.config(scrollregion=(0, 0, img.width, img.height))
        self.incidencias_current_map = mapa_id
        self.incidencias_color_used = set()
        for area in self.incidencias_db.list_areas(mapa_id):
            area_id, nombre, x1, y1, x2, y2, color = area
            self.incidencias_color_used.add(color)
            item = self.incidencias_canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2, tags=("area", str(area_id)))
            self.incidencias_area_items[area_id] = item
        for m in self.incidencias_db.list_machines(mapa_id):
            mid, area_id, nombre, serie, numero, x1, y1, x2, y2, color, _ = m
            self.incidencias_color_used.add(color)
            item = self.incidencias_canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2, tags=("machine", str(mid)))
            self.incidencias_machine_items[mid] = item
            self.incidencias_machine_area[mid] = area_id
        self._incidencias_apply_map_filter(self.incidencias_info_filter_area)

    def incidencias_asignar_area(self):
        if not self._incidencias_pin_ok():
            return
        self._bring_to_front()
        nombre = simpledialog.askstring("Area", "Como se va a llamar el area?", parent=self)
        if not nombre:
            return
        self.incidencias_mode = ("area", nombre)
        self._bring_to_front()
        messagebox.showinfo("Area", "Dibuja el area en el mapa.", parent=self)

    def incidencias_editar_area(self):
        if not self._incidencias_pin_ok():
            return
        self.incidencias_selected_area = None
        self.incidencias_mode = ("pick_area_edit",)
        self._bring_to_front()
        messagebox.showinfo("Area", "Selecciona el area en el mapa.", parent=self)

    def incidencias_listar_maquina(self):
        if not self._incidencias_pin_ok():
            return
        self.incidencias_selected_area = None
        self.incidencias_mode = ("pick_area_machine",)
        self._bring_to_front()
        messagebox.showinfo("Maquina", "Selecciona el area en el mapa.", parent=self)

    def incidencias_editar_maquina(self):
        if not self._incidencias_pin_ok():
            return
        self.incidencias_selected_machine = None
        self.incidencias_mode = ("pick_machine_edit",)
        self._bring_to_front()
        messagebox.showinfo("Maquina", "Selecciona la maquina en el mapa.", parent=self)

    def incidencias_info_maquinas(self, area_id=None):
        self.incidencias_panel_mode = "machines"
        # Si se llama desde el boton, limpiamos el filtro
        if area_id is None:
            self.incidencias_info_filter_area = None
        else:
            self.incidencias_info_filter_area = area_id

        self._incidencias_set_area_header(self.incidencias_info_filter_area)
        self._incidencias_apply_map_filter(self.incidencias_info_filter_area)

        for w in self.incidencias_panel.winfo_children():
            w.destroy()
        container = tk.Frame(self.incidencias_panel)
        container.pack(fill="both", expand=True)
        tree = ttk.Treeview(container, columns=["area", "nombre", "serie", "numero"], show="headings")
        for c in ["area", "nombre", "serie", "numero"]:
            tree.heading(c, text=c.capitalize())
            tree.column(c, anchor="center")
        tree.column("area", width=120, stretch=True)
        tree.column("nombre", width=140, stretch=True)
        tree.column("serie", width=120, stretch=True)
        tree.column("numero", width=100, stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        vscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vscroll.set)

        for m in self.incidencias_db.list_machines(self.incidencias_current_map):
            mid, area_id_db, nombre, serie, numero, *_rest, area_nombre = m
            if self.incidencias_info_filter_area and area_id_db != self.incidencias_info_filter_area:
                continue
            tree.insert("", "end", values=[area_nombre, nombre, serie, numero], iid=str(mid))

        def on_select(event):
            sel = tree.selection()
            if not sel:
                return
            mid = int(sel[0])
            self._incidencias_resaltar_maquina(mid)

        def on_hover(event):
            row = tree.identify_row(event.y)
            if not row:
                self._incidencias_stop_hover_blink()
                return
            self._incidencias_start_hover_blink(int(row))

        def on_leave(event):
            self._incidencias_stop_hover_blink()

        def copy_cell(event):
            row = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if not row or not col:
                return
            col_index = int(col[1:]) - 1
            values = tree.item(row, "values")
            if col_index >= len(values):
                return
            value = values[col_index]
            self.clipboard_clear()
            self.clipboard_append(str(value))
            self.update()

        def delete_machine(event):
            row = tree.identify_row(event.y)
            if not row:
                return
            mid = int(row)
            if not self._incidencias_pin_ok():
                return
            self._bring_to_front()
            if not messagebox.askyesno("Eliminar", "Eliminar maquina y todos sus datos?", parent=self):
                return
            self.incidencias_db.delete_machine(mid)
            self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
            self.incidencias_info_maquinas(self.incidencias_info_filter_area)

        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Copiar", command=lambda: copy_cell(tree.event_context))
        menu.add_command(label="Eliminar maquina", command=lambda: delete_machine(tree.event_context))
        tree.bind("<<TreeviewSelect>>", on_select)
        tree.bind("<Motion>", on_hover)
        tree.bind("<Leave>", on_leave)
        tree.bind("<Button-3>", lambda e: (setattr(tree, 'event_context', e), menu.post(e.x_root, e.y_root)))
        self.incidencias_machines_tree = tree

    def incidencias_crear_incidencia(self):
        self._bring_to_front()
        respuesta = messagebox.askyesno("Incidencia", "Es sobre una maquina?", parent=self)
        if respuesta:
            self.incidencias_mode = ("inc_maquina",)
            self._bring_to_front()
            messagebox.showinfo("Incidencia", "Haz click sobre la maquina en el mapa.", parent=self)
        else:
            self.incidencias_mode = ("inc_area",)
            self._bring_to_front()
            messagebox.showinfo("Incidencia", "Haz click sobre el area en el mapa.", parent=self)

    def incidencias_gestion_incidencias(self):
        self.incidencias_panel_mode = "incidencias"
        self.incidencias_info_filter_area = None
        self._incidencias_set_area_header(None)
        self._incidencias_apply_map_filter(None)
        for w in self.incidencias_panel.winfo_children():
            w.destroy()
        if not self.incidencias_current_map:
            return
        cols = ["fecha", "estado", "area", "maquina", "serie", "numero", "elemento", "descripcion", "reporte"]
        container = tk.Frame(self.incidencias_panel)
        container.pack(fill="both", expand=True)
        tree = ttk.Treeview(container, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c.capitalize())
            tree.column(c, anchor="center")
        tree.column("fecha", width=120, stretch=True)
        tree.column("estado", width=90, stretch=True)
        tree.column("area", width=120, stretch=True)
        tree.column("maquina", width=140, stretch=True)
        tree.column("serie", width=120, stretch=True)
        tree.column("numero", width=100, stretch=True)
        tree.column("elemento", width=120, stretch=True)
        tree.column("descripcion", width=260, stretch=True)
        tree.column("reporte", width=70, stretch=False)
        tree.grid(row=0, column=0, sticky="nsew")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        vscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=vscroll.set)
        self.incidencias_inc_map = {}
        tree.tag_configure("pendiente", background="#ffe6cc")
        tree.tag_configure("visto", background="#fff3cd")
        tree.tag_configure("reparado", background="#d4edda")
        self.incidencias_area_to_inc = {}
        for inc in self.incidencias_db.list_incidencias(self.incidencias_current_map):
            inc_id, fecha, estado, elemento, descripcion, reporte_path, area, maquina, serie, numero = inc
            tag = ""
            if estado == "PENDIENTE":
                tag = "pendiente"
            elif estado == "VISTO":
                tag = "visto"
            elif estado == "REPARADO":
                tag = "reparado"
            tree.insert(
                "",
                "end",
                values=[fecha, estado, area or "", maquina or "", serie or "", numero or "", elemento or "", descripcion or "", ("[R]" if reporte_path else "")],
                iid=str(inc_id),
                tags=(tag,),
            )
            self.incidencias_inc_map[inc_id] = {
                "area": area,
                "maquina": maquina,
                "serie": serie,
                "numero": numero,
                "descripcion": descripcion,
                "reporte": reporte_path,
            }
            if area:
                self.incidencias_area_to_inc.setdefault(area, []).append(inc_id)

        tooltip = {"win": None, "text": ""}

        def show_tooltip(texto, x, y):
            if not texto:
                return
            if tooltip["win"] is None:
                tip = tk.Toplevel(self)
                tip.wm_overrideredirect(True)
                label = tk.Label(tip, text=texto, justify="left", background="#ffffe0", relief="solid", borderwidth=1, wraplength=380)
                label.pack(ipadx=4, ipady=2)
                tooltip["win"] = tip
            tooltip["win"].wm_geometry(f"+{x}+{y}")
            tooltip["text"] = texto

        def hide_tooltip():
            if tooltip["win"] is not None:
                tooltip["win"].destroy()
                tooltip["win"] = None
                tooltip["text"] = ""

        def copy_cell(event):
            row = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if not row or not col:
                return
            col_index = int(col[1:]) - 1
            values = tree.item(row, "values")
            if col_index >= len(values):
                return
            value = values[col_index]
            self.clipboard_clear()
            self.clipboard_append(str(value))
            self.update()

        def on_right_click(event):
            item = tree.identify_row(event.y)
            if not item:
                return
            menu = tk.Menu(tree, tearoff=0)
            menu.add_command(label="Modificar Incidencia", command=lambda: self._incidencias_editar_incidencia(int(item)))
            menu.add_command(label="Eliminar Incidencia", command=lambda: self._incidencias_eliminar_incidencia(int(item)))
            menu.add_command(label="Modificar Estado", command=lambda: self._incidencias_cambiar_estado(int(item)))
            reporte = self.incidencias_inc_map.get(int(item), {}).get("reporte")
            if reporte:
                menu.add_command(label="Ver reporte visual", command=lambda: self._incidencias_ver_reporte(reporte))
                menu.add_command(label="Guardar reporte visual", command=lambda: self._incidencias_guardar_reporte(reporte))
            menu.add_command(label="Copiar", command=lambda: copy_cell(event))
            menu.post(event.x_root, event.y_root)

        def on_double_click(event):
            item = tree.identify_row(event.y)
            if not item:
                return
            self._incidencias_cambiar_estado(int(item))

        def on_hover(event):
            item = tree.identify_row(event.y)
            if not item:
                self._incidencias_resaltar_maquina(None)
                self._incidencias_resaltar_area(None)
                hide_tooltip()
                return
            inc_id = int(item)
            data = self.incidencias_inc_map.get(inc_id, {})
            maquina_name = data.get("maquina")
            if maquina_name:
                mid = self._incidencias_find_machine_id_by_name(maquina_name)
                self._incidencias_resaltar_maquina(mid)
            else:
                area_name = data.get("area")
                aid = self._incidencias_find_area_id_by_name(area_name) if area_name else None
                self._incidencias_resaltar_area(aid)

            col = tree.identify_column(event.x)
            col_index = int(col[1:]) - 1 if col else -1
            if col_index == cols.index("descripcion"):
                values = tree.item(item, "values")
                desc = values[col_index] if col_index < len(values) else ""
                if desc:
                    show_tooltip(str(desc), event.x_root + 12, event.y_root + 12)
                else:
                    hide_tooltip()
            else:
                hide_tooltip()

        def on_select(event):
            sel = tree.selection()
            if not sel:
                return
            inc_id = int(sel[0])
            data = self.incidencias_inc_map.get(inc_id, {})
            area_name = data.get("area")
            maquina_name = data.get("maquina")
            if area_name:
                aid = self._incidencias_find_area_id_by_name(area_name)
                if aid:
                    self._incidencias_blink_area(aid)
            if maquina_name:
                mid = self._incidencias_find_machine_id_by_name(maquina_name)
                if mid:
                    self._incidencias_blink_machine(mid)

        def on_leave(event):
            self._incidencias_resaltar_maquina(None)
            self._incidencias_resaltar_area(None)
            hide_tooltip()

        tree.bind("<Button-3>", on_right_click)
        tree.bind("<Double-1>", on_double_click)
        tree.bind("<Motion>", on_hover)
        tree.bind("<Leave>", on_leave)
        tree.bind("<<TreeviewSelect>>", on_select)
        self.incidencias_incidencias_tree = tree

    def _incidencias_blink_item(self, item_id):
        if not item_id:
            return
        original_width = self.incidencias_canvas.itemcget(item_id, "width")
        original_outline = self.incidencias_canvas.itemcget(item_id, "outline")

        def toggle(on):
            if on:
                self.incidencias_canvas.itemconfig(item_id, width=4, outline="#ff0000")
            else:
                self.incidencias_canvas.itemconfig(item_id, width=original_width, outline=original_outline)

        toggle(True)
        self.after(150, lambda: toggle(False))
        self.after(300, lambda: toggle(True))
        self.after(450, lambda: toggle(False))

    def _incidencias_blink_area(self, area_id):
        item = self.incidencias_area_items.get(area_id)
        if item:
            self._incidencias_blink_item(item)

    def _incidencias_blink_machine(self, machine_id):
        item = self.incidencias_machine_items.get(machine_id)
        if item:
            self._incidencias_blink_item(item)

    def _incidencias_resaltar_maquina(self, mid):
        for item in self.incidencias_machine_items.values():
            self.incidencias_canvas.itemconfig(item, width=2)
        if mid and mid in self.incidencias_machine_items:
            self.incidencias_canvas.itemconfig(self.incidencias_machine_items[mid], width=4)

    def _incidencias_resaltar_area(self, area_id):
        for item in self.incidencias_area_items.values():
            self.incidencias_canvas.itemconfig(item, width=2)
        if area_id and area_id in self.incidencias_area_items:
            self.incidencias_canvas.itemconfig(self.incidencias_area_items[area_id], width=4)

    def incidencias_canvas_hover(self, event):
        cx = self.incidencias_canvas.canvasx(event.x)
        cy = self.incidencias_canvas.canvasy(event.y)
        item = self.incidencias_canvas.find_closest(cx, cy)
        if not item:
            self.incidencias_canvas_leave(event)
            return
        item_id = item[0]
        tags = self.incidencias_canvas.gettags(item_id)
        if "machine" in tags:
            mid = int(tags[1])
            self._incidencias_resaltar_maquina(mid)
            if self.incidencias_panel_mode == "machines" and hasattr(self, "incidencias_machines_tree"):
                self.incidencias_machines_tree.selection_set(str(mid))
                self.incidencias_machines_tree.see(str(mid))
        elif "area" in tags:
            aid = int(tags[1])
            self._incidencias_resaltar_area(aid)
            if self.incidencias_panel_mode == "incidencias" and hasattr(self, "incidencias_incidencias_tree"):
                # Selecciona la incidencia mas reciente del area
                inc_ids = self.incidencias_area_to_inc.get(self._incidencias_area_name_by_id(aid), [])
                if inc_ids:
                    inc_id = str(inc_ids[0])
                    self.incidencias_incidencias_tree.selection_set(inc_id)
                    self.incidencias_incidencias_tree.see(inc_id)

    def incidencias_canvas_leave(self, event):
        self._incidencias_resaltar_maquina(None)
        self._incidencias_resaltar_area(None)
        if self.incidencias_panel_mode == "machines" and hasattr(self, "incidencias_machines_tree"):
            self.incidencias_machines_tree.selection_remove(self.incidencias_machines_tree.selection())
        if self.incidencias_panel_mode == "incidencias" and hasattr(self, "incidencias_incidencias_tree"):
            self.incidencias_incidencias_tree.selection_remove(self.incidencias_incidencias_tree.selection())

    def _incidencias_area_name_by_id(self, area_id):
        for a in self.incidencias_db.list_areas(self.incidencias_current_map):
            aid, anombre, *_rest = a
            if aid == area_id:
                return anombre
        return None

    def _incidencias_find_machine_id_by_name(self, nombre):
        for mid, item in self.incidencias_machine_items.items():
            # Busca en DB para evitar depender del canvas
            pass
        for m in self.incidencias_db.list_machines(self.incidencias_current_map):
            mid, _area_id, mnombre, _serie, _numero, *_rest = m
            if mnombre == nombre:
                return mid
        return None

    def _incidencias_find_area_id_by_name(self, nombre):
        for a in self.incidencias_db.list_areas(self.incidencias_current_map):
            aid, anombre, *_rest = a
            if anombre == nombre:
                return aid
        return None
    def _incidencias_editar_incidencia(self, inc_id):
        self._bring_to_front()
        elemento = simpledialog.askstring("Incidencia", "Material/elemento:", parent=self)
        descripcion = simpledialog.askstring("Incidencia", "Describe la incidencia:", parent=self)
        estado = self._incidencias_selector_estado()
        if not estado:
            estado = "PENDIENTE"
        self.incidencias_db.update_incidencia(inc_id, elemento or "", descripcion or "", estado)
        self.incidencias_gestion_incidencias()

    def _incidencias_eliminar_incidencia(self, inc_id):
        self._bring_to_front()
        if messagebox.askyesno("Eliminar", "Eliminar incidencia?", parent=self):
            reporte = None
            if hasattr(self, "incidencias_inc_map"):
                reporte = self.incidencias_inc_map.get(inc_id, {}).get("reporte")
            if reporte and os.path.exists(reporte):
                try:
                    os.remove(reporte)
                except Exception:
                    pass
            self.incidencias_db.delete_incidencia(inc_id)
            self.incidencias_gestion_incidencias()

    def _incidencias_cambiar_estado(self, inc_id):
        estado = self._incidencias_selector_estado()
        if not estado:
            return
        self.incidencias_db.update_incidencia_estado(inc_id, estado)
        self.incidencias_gestion_incidencias()

    def _incidencias_selector_estado(self):
        self._bring_to_front()
        win = tk.Toplevel(self)
        win.title("Estado")
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text="Selecciona estado:").pack(padx=10, pady=5)
        opciones = ["PENDIENTE", "VISTO", "REPARADO"]
        var = tk.StringVar(value=opciones[0])
        combo = ttk.Combobox(win, values=opciones, textvariable=var, state="readonly")
        combo.pack(padx=10, pady=5)
        combo.focus_set()
        resultado = {"valor": None}

        def aceptar():
            resultado["valor"] = var.get()
            win.destroy()

        tk.Button(win, text="Aceptar", command=aceptar).pack(pady=5)
        win.grab_set()
        win.wait_window()
        return resultado["valor"]

    def incidencias_canvas_press(self, event):
        cx = self.incidencias_canvas.canvasx(event.x)
        cy = self.incidencias_canvas.canvasy(event.y)
        if not self.incidencias_mode:
            # seleccionar area/maquina
            item = self.incidencias_canvas.find_closest(cx, cy)
            if not item:
                return
            item_id = item[0]
            tags = self.incidencias_canvas.gettags(item_id)
            if "area" in tags:
                self.incidencias_selected_area = int(tags[1])
                if self.incidencias_panel_mode == "machines":
                    self.incidencias_info_maquinas(self.incidencias_selected_area)
            if "machine" in tags:
                self.incidencias_selected_machine = int(tags[1])
            return

        mode = self.incidencias_mode[0]
        if mode in ("pick_area_edit", "pick_area_machine", "pick_machine_edit", "inc_maquina", "inc_area"):
            item = self.incidencias_canvas.find_closest(cx, cy)
            if not item:
                return
            item_id = item[0]
            tags = self.incidencias_canvas.gettags(item_id)

            if mode == "pick_area_edit":
                if "area" not in tags:
                    self._bring_to_front()
                    messagebox.showwarning("Area", "Selecciona un area en el mapa.", parent=self)
                    return
                self.incidencias_selected_area = int(tags[1])
                self._bring_to_front()
                nombre = simpledialog.askstring("Area", "Nuevo nombre del area:", parent=self)
                if not nombre:
                    self.incidencias_mode = None
                    return
                self.incidencias_mode = ("area_edit", nombre, self.incidencias_selected_area)
                self._bring_to_front()
                messagebox.showinfo("Area", "Dibuja el nuevo area en el mapa.", parent=self)
                return

            if mode == "pick_area_machine":
                if "area" not in tags:
                    self._bring_to_front()
                    messagebox.showwarning("Area", "Selecciona un area en el mapa.", parent=self)
                    return
                self.incidencias_selected_area = int(tags[1])
                self.incidencias_mode = ("machine", self.incidencias_selected_area)
                self._bring_to_front()
                messagebox.showinfo("Maquina", "Dibuja la subarea de la maquina en el mapa.", parent=self)
                return

            if mode == "pick_machine_edit":
                if "machine" not in tags:
                    self._bring_to_front()
                    messagebox.showwarning("Maquina", "Selecciona una maquina en el mapa.", parent=self)
                    return
                self.incidencias_selected_machine = int(tags[1])
                self._bring_to_front()
                nombre = simpledialog.askstring("Maquina", "Nuevo nombre de la maquina:", parent=self)
                if not nombre:
                    self.incidencias_mode = None
                    return
                serie = simpledialog.askstring("Maquina", "Numero de serie:", parent=self)
                numero = simpledialog.askstring("Maquina", "Numero asignado:", parent=self)
                self.incidencias_mode = ("machine_edit", self.incidencias_selected_machine, nombre, serie, numero)
                self._bring_to_front()
                messagebox.showinfo("Maquina", "Dibuja la nueva posicion de la maquina.", parent=self)
                return

            if mode == "inc_maquina":
                if "machine" not in tags:
                    self._bring_to_front()
                    messagebox.showwarning("Incidencia", "Selecciona una maquina en el mapa.", parent=self)
                    return
                mid = int(tags[1])
                machines = self.incidencias_db.list_machines(self.incidencias_current_map)
                area_id = None
                for m in machines:
                    if m[0] == mid:
                        area_id = m[1]
                        break
                self._bring_to_front()
                elemento = simpledialog.askstring("Incidencia", "Material/elemento:", parent=self)
                descripcion = simpledialog.askstring("Incidencia", "Describe la incidencia:", parent=self)
                if elemento is None and descripcion is None:
                    self.incidencias_mode = None
                    return
                reporte_path = self._incidencias_pedir_reporte_visual()
                self.incidencias_db.add_incident(
                    self.incidencias_current_map,
                    area_id,
                    mid,
                    elemento or "",
                    descripcion or "",
                    reporte_path=reporte_path,
                )
                self.incidencias_gestion_incidencias()
                self.incidencias_mode = None
                return

            if mode == "inc_area":
                if "area" not in tags:
                    self._bring_to_front()
                    messagebox.showwarning("Incidencia", "Selecciona un area en el mapa.", parent=self)
                    return
                area_id = int(tags[1])
                self._bring_to_front()
                elemento = simpledialog.askstring("Incidencia", "Material/elemento:", parent=self)
                descripcion = simpledialog.askstring("Incidencia", "Describe la incidencia:", parent=self)
                if elemento is None and descripcion is None:
                    self.incidencias_mode = None
                    return
                reporte_path = self._incidencias_pedir_reporte_visual()
                self.incidencias_db.add_incident(
                    self.incidencias_current_map,
                    area_id,
                    None,
                    elemento or "",
                    descripcion or "",
                    reporte_path=reporte_path,
                )
                self.incidencias_gestion_incidencias()
                self.incidencias_mode = None
                return

        self.incidencias_draw_start = (cx, cy)
        self.incidencias_draw_rect = self.incidencias_canvas.create_rectangle(
            cx, cy, cx, cy, outline="red", dash=(2, 2)
        )

    def incidencias_canvas_drag(self, event):
        if not self.incidencias_draw_rect:
            return
        x1, y1 = self.incidencias_draw_start
        cx = self.incidencias_canvas.canvasx(event.x)
        cy = self.incidencias_canvas.canvasy(event.y)
        self.incidencias_canvas.coords(self.incidencias_draw_rect, x1, y1, cx, cy)

    def incidencias_canvas_release(self, event):
        if not self.incidencias_mode:
            return
        if not self.incidencias_draw_rect:
            return
        x1, y1, x2, y2 = self.incidencias_canvas.coords(self.incidencias_draw_rect)
        self.incidencias_canvas.delete(self.incidencias_draw_rect)
        self.incidencias_draw_rect = None
        self._bring_to_front()
        if not messagebox.askyesno("Confirmar", "Estas seguro?", parent=self):
            self.incidencias_mode = None
            return
        if self.incidencias_mode[0] == "area":
            nombre = self.incidencias_mode[1]
            color = self._incidencias_color()
            self.incidencias_db.add_area(self.incidencias_current_map, nombre, int(x1), int(y1), int(x2), int(y2), color)
            self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
        elif self.incidencias_mode[0] == "area_edit":
            nombre, area_id = self.incidencias_mode[1], self.incidencias_mode[2]
            self.incidencias_db.update_area(area_id, nombre, int(x1), int(y1), int(x2), int(y2))
            self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
        elif self.incidencias_mode[0] == "machine":
            area_id = self.incidencias_mode[1]
            self._bring_to_front()
            nombre = simpledialog.askstring("Maquina", "Nombre de la maquina:", parent=self)
            serie = simpledialog.askstring("Maquina", "Numero de serie:", parent=self)
            numero = simpledialog.askstring("Maquina", "Numero asignado:", parent=self)
            color = self._incidencias_color()
            self.incidencias_db.add_machine(area_id, nombre or "", serie or "", numero or "", int(x1), int(y1), int(x2), int(y2), color)
            self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
        elif self.incidencias_mode[0] == "machine_edit":
            mid, nombre, serie, numero = self.incidencias_mode[1], self.incidencias_mode[2], self.incidencias_mode[3], self.incidencias_mode[4]
            self.incidencias_db.update_machine(mid, nombre or "", serie or "", numero or "", int(x1), int(y1), int(x2), int(y2))
            self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
        self.incidencias_vista_general()
        self.incidencias_mode = None

    def exportar_excel(self):
        try:
            pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
            if pin is None:
                return
            if pin.strip() != get_security_code():
                messagebox.showerror("Codigo incorrecto", "El codigo de seguridad no es valido.", parent=self)
                return

            pestana_activa = self.notebook.select()
            nombre_pestana = self.notebook.tab(pestana_activa, "text")
            tree = self.tabs[nombre_pestana].tree

            if not tree.get_children():
                messagebox.showwarning("Sin datos", f"No hay datos para exportar en la pestana {nombre_pestana}.", parent=self)
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

    def enviar_asuntos_propios(self):
        """
        Abre Outlook con un borrador dirigido al manager para solicitar día de asuntos propios.
        """
        destinatario = "manager.sevilla-manueldevillalobos@fitnesspark.es"
        asunto = "Petición de Día de Asuntos Propios"
        cuerpo = (
            "En Sevilla, a __ / __ / 20__\n"
            "Yo, _______________________, con DNI, miembro del equipo\n"
            "del club Fitness Park Sevilla - Villalobos, presento la siguiente solicitud:\n\n"
            "Objeto de la petición\n"
            "Solicito autorización para disfrutar de un día de asuntos propios, conforme a las condiciones establecidas en mi\n"
            "contrato y en el protocolo interno del club.\n"
            "Fecha solicitada: ____ / ____ / 20___\n"
            "Turno habitual en dicha fecha: [ ] Mañana [ ] Tarde [ ] Noche\n"
            "Declaración del trabajador/a\n"
            "Declaro que:\n"
            "- Esta solicitud se realiza con la antelación mínima exigida por la organización.\n"
            "- Entiendo que los días de asuntos propios deben disfrutarse en días laborables y siempre respetando la\n"
            "cobertura del servicio.\n"
            "- Acepto que la concesión del día queda sujeta a aprobación por parte de la Dirección del club,\n"
            "garantizando que no se vea afectada la operativa ni el equilibrio de turnos del equipo."
        )

        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            messagebox.showwarning(
                "Outlook no disponible",
                f"No se pudo importar win32com.client.\n"
                f"Cuerpo copiado al portapapeles. Envía un correo a {destinatario} con asunto:\n{asunto}"
            )
            return

        def try_outlook():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            try:
                ns = outlook_app.GetNamespace("MAPI")
                ns.Logon("", "", False, False)
            except Exception:
                pass
            return outlook_app

        try:
            try:
                outlook = try_outlook()
            except Exception:
                os.startfile("outlook.exe")
                time.sleep(3)
                outlook = try_outlook()

            mail = outlook.CreateItem(0)
            mail.To = destinatario
            mail.Subject = asunto
            mail.Body = cuerpo
            mail.Display()
            messagebox.showinfo(
                "Borrador creado",
                f"Outlook se abrió con el correo dirigido a {destinatario}."
            )
        except Exception as e:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook.\n"
                f"Cuerpo copiado al portapapeles para pegarlo manualmente.\n\nDetalle: {e}"
            )

    def enviar_cambio_turno(self):
        """
        Abre Outlook con un borrador para solicitud de cambio de turno (sin PIN).
        """
        destinatario = "manager.sevilla-manueldevillalobos@fitnesspark.es"
        asunto = "Solicitud Cambio de turno"
        cuerpo = (
            "PETICIÓN VOLUNTARIA DE LOS SOLICITANTES QUE EN TODO MOMENTO Y BAJO SU RESPONSABILIDAD "
            "PROPORCIONARÁ LA COBERTURA NECESARIA PARA QUE LAS NECESIDADES DEL CLUB ESTÉN CUBIERTAS. "
            "REQUIERE DE APROBACIÓN POR PARTE DEL DIRECTOR DEL CLUB.\n\n"
            "STAFF 1 QUE SOLICITA EL CAMBIO*\n"
            "NOMBRE Y APELLIDOS\n\n"
            "STAFF 2 QUE ACEPTA EL CAMBIO*\n"
            "NOMBRE Y APELLIDOS\n\n\n"
            "DESCRIBIR A CONTINUACIÓN CAMBIO ESPECIFICANDO FECHAS Y TURNOS. "
            "RECUERDA QUE DEBEN SER CAMBIOS CERRADOS, SIN QUE QUEDE NADA EN EL AIRE:"
        )

        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            messagebox.showwarning(
                "Outlook no disponible",
                f"No se pudo importar win32com.client.\n"
                f"Cuerpo copiado al portapapeles. Envía un correo a {destinatario} con asunto:\n{asunto}"
            )
            return

        def try_outlook():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            try:
                ns = outlook_app.GetNamespace("MAPI")
                ns.Logon("", "", False, False)
            except Exception:
                pass
            return outlook_app

        try:
            try:
                outlook = try_outlook()
            except Exception:
                os.startfile("outlook.exe")
                time.sleep(3)
                outlook = try_outlook()

            mail = outlook.CreateItem(0)
            mail.To = destinatario
            mail.Subject = asunto
            mail.Body = cuerpo
            mail.Display()
            messagebox.showinfo(
                "Borrador creado",
                f"Outlook se abrió con el correo dirigido a {destinatario}."
            )
        except Exception as e:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook.\n"
                f"Cuerpo copiado al portapapeles para pegarlo manualmente.\n\nDetalle: {e}"
            )

    def enviar_felicitacion(self):
        """
        Pide PIN y abre un borrador de Outlook con los emails de Cumpleaños en BCC.
        """
        df = self.dataframes.get("Cumpleaños")
        if df is None or df.empty:
            messagebox.showwarning("Sin datos", "No hay datos cargados en la pestaña Cumpleaños.")
            return

        pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
        if pin is None:
            return  # cancelado
        if pin.strip() != get_security_code():
            messagebox.showerror("Codigo incorrecto", "El codigo de seguridad no es valido.", parent=self)
            return

        def normalize(text):
            t = unicodedata.normalize("NFD", str(text or "")).upper().strip()
            return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")

        colmap = {normalize(col): col for col in df.columns}
        col_email = colmap.get("CORREO ELECTRONICO") or colmap.get("EMAIL") or colmap.get("CORREO")

        if not col_email:
            messagebox.showerror("Columna faltante", "No se encontró la columna de correo electrónico en Cumpleaños.")
            return

        emails = df[col_email].dropna().astype(str).str.strip()
        emails = [e for e in emails if e]
        emails = list(dict.fromkeys(emails))  # quitar duplicados manteniendo orden

        # Filtra ya enviados este año
        year = str(datetime.now().year)
        enviados = set(self.felicitaciones_enviadas.get(year, []))
        emails_pendientes = [e for e in emails if e.lower() not in enviados]

        if not emails_pendientes:
            messagebox.showwarning("Sin correos", "Todos los cumpleaños de hoy ya fueron felicitados este año.")
            return

        try:
            import win32com.client  # tipo: ignore
        except ImportError:
            # Fallback: copiar correos al portapapeles para pegarlos en BCC manualmente.
            emails_joined = ";".join(emails_pendientes)
            self.clipboard_clear()
            self.clipboard_append(emails_joined)
            messagebox.showwarning(
                "Outlook no disponible",
                "No se pudo importar win32com.client. Se copiaron los correos al portapapeles para pegarlos en BCC.\n"
                "Instala pywin32 si quieres que Outlook se abra automáticamente."
            )
            return

        def try_outlook():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            # Garantiza sesión; si Outlook no está abierto, fuerza logon.
            try:
                ns = outlook_app.GetNamespace("MAPI")
                ns.Logon("", "", False, False)
            except Exception:
                pass
            return outlook_app

        try:
            try:
                outlook = try_outlook()
            except Exception:
                # Intentar arrancar Outlook y reintentar
                try:
                    os.startfile("outlook.exe")
                    time.sleep(3)
                    outlook = try_outlook()
                except Exception:
                    raise

            mail = outlook.CreateItem(0)
            mail.BCC = ";".join(emails_pendientes)
            mail.Subject = "¡Felicidades!"

            # Inserta imagen inline si existe en disco.
            imagen_nombre = "feliz_cumpleanos.png"
            imagen_path = get_logo_path(imagen_nombre)
            cid = uuid.uuid4().hex

            if os.path.exists(imagen_path):
                attachment = mail.Attachments.Add(imagen_path, 1, 0)  # 1 = olByValue
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid
                )
                mail.HTMLBody = (
                    f'<div style="font-family:Arial,sans-serif;font-size:14px;">'
                    f'<img src="cid:{cid}" alt="Feliz cumpleaños" style="max-width:100%;">'
                    f"</div>"
                )
            else:
                mail.HTMLBody = (
                    '<div style="font-family:Arial,sans-serif;font-size:14px;">'
                    "<p>¡Feliz cumpleaños!</p>"
                    "</div>"
                )

            mail.Display()  # abre borrador para revisión
            if not messagebox.askyesno("Confirmación", "Ha enviado la felicitacion?"):
                return
            # Marca como enviados para este año
            enviados.update([e.lower() for e in emails_pendientes])
            self.felicitaciones_enviadas[year] = sorted(enviados)
            self.guardar_felicitaciones()
            messagebox.showinfo("Registrado", "Felicitaciones registradas. No se reenviarán hasta el próximo año.")
        except Exception as e:
            emails_joined = ";".join(emails_pendientes)
            self.clipboard_clear()
            self.clipboard_append(emails_joined)
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook.\nSe copiaron los correos al portapapeles para pegarlos en BCC.\n\nDetalle: {e}"
            )

    def extraer_accesos(self):
        """
        Pide PIN y numero de cliente, filtra ACCESOS.csv y muestra resultado en pestaña 'Accesos Cliente'.
        """
        if self.raw_accesos is None:
            messagebox.showwarning("Sin datos", "Primero carga datos (Actualizar datos) para tener ACCESOS.")
            return

        # Trae la ventana al frente para que los dialogs tomen foco
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass

        pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
        if pin is None:
            return
        if pin.strip() != get_security_code():
            messagebox.showerror("Codigo incorrecto", "El codigo de seguridad no es valido.", parent=self)
            return

        numero = simpledialog.askstring("Número de cliente", "Introduce el número de cliente:")
        if numero is None:
            return
        numero = numero.strip()
        if not numero:
            messagebox.showwarning("Sin número", "No se ingresó un número de cliente.")
            return

        def normalize(text):
            t = unicodedata.normalize("NFD", str(text or "")).upper().strip()
            return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")

        df = self.raw_accesos
        colmap = {normalize(col): col for col in df.columns}
        col_cliente = colmap.get("NUMERO DE CLIENTE") or colmap.get("NÚMERO DE CLIENTE") or colmap.get("NカMERO DE CLIENTE")

        if not col_cliente:
            messagebox.showerror("Columna faltante", "No se encontró la columna 'Número de cliente' en ACCESOS.")
            return

        df_filtrado = df[df[col_cliente].astype(str).str.strip() == numero].copy()

        if df_filtrado.empty:
            messagebox.showinfo("Sin resultados", f"No hay accesos para el número de cliente: {numero}")
        else:
            self.mostrar_en_tabla("Accesos Cliente", df_filtrado)
            # Ofrecer guardado inmediato
            try:
                archivo = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    initialfile=f"accesos_{numero}.xlsx"
                )
                if archivo:
                    df_filtrado.to_excel(archivo, index=False)
                    messagebox.showinfo("Accesos filtrados", f"Se encontraron {len(df_filtrado)} accesos para el cliente {numero}.\nGuardado en:\n{archivo}")
                else:
                    messagebox.showinfo("Accesos filtrados", f"Se encontraron {len(df_filtrado)} accesos para el cliente {numero}.")
            except Exception as e:
                messagebox.showerror("Error al guardar", f"Se filtraron {len(df_filtrado)} accesos pero falló el guardado:\n{e}")

    def enviar_avanza_fit(self):
        """
        Pide PIN y abre borrador Outlook con el Excel adjunto de la pestaña Avanza Fit.
        """
        df = self.dataframes.get("Avanza Fit")
        if df is None or df.empty:
            messagebox.showwarning("Sin datos", "No hay datos cargados en la pestaña Avanza Fit.")
            return

        pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
        if pin is None:
            return
        if pin.strip() != get_security_code():
            messagebox.showerror("Codigo incorrecto", "El codigo de seguridad no es valido.", parent=self)
            return

        # Generar archivo temporal con los datos
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_path = tmp_file.name
        tmp_file.close()
        df.to_excel(tmp_path, index=False)

        try:
            import win32com.client  # tipo: ignore
        except ImportError:
            messagebox.showwarning(
                "Outlook no disponible",
                f"No se pudo importar win32com.client. Adjunta manualmente el archivo:\n{tmp_path}"
            )
            return

        def try_outlook():
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            try:
                ns = outlook_app.GetNamespace("MAPI")
                ns.Logon("", "", False, False)
            except Exception:
                pass
            return outlook_app

        try:
            try:
                outlook = try_outlook()
            except Exception:
                os.startfile("outlook.exe")
                time.sleep(3)
                outlook = try_outlook()

            mail = outlook.CreateItem(0)
            mail.Subject = "Clientes Avanza Fit (martes)"
            mail.Body = "Adjunto listado de clientes Avanza Fit (martes-lunes)."
            mail.Attachments.Add(tmp_path)
            mail.Display()
            messagebox.showinfo("Borrador creado", "Outlook se abrió con el archivo de Avanza Fit adjunto.")
        except Exception as e:
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook. Adjunta manualmente el archivo:\n{tmp_path}\n\nDetalle: {e}"
            )


if __name__ == "__main__":
    app = ResamaniaApp()
    app.mainloop()
