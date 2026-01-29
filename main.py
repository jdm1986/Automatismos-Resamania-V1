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
import io
from datetime import datetime, timedelta
import re
import urllib.parse
import webbrowser
import shutil
import random
import traceback
from logic.wizville import procesar_wizville
from logic.accesos import procesar_salidas_pmr_no_autorizadas, procesar_accesos_dobles_ayer
from logic.avanza_fit import obtener_avanza_fit
from utils.file_loader import load_data_file
from logic.impagos import ImpagosDB
from logic.incidencias import IncidenciasDB
from logic.state_store import AppStateStore


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(get_app_dir(), "config.json")


def _read_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    return {}


def _write_config(data):
    with open(CONFIG_PATH, "w") as f:
        json.dump(data, f)


def _log_timing(label, elapsed):
    try:
        log_path = os.path.join(get_app_dir(), "timings.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()} | {label} | {elapsed:.3f}s\n")
    except Exception:
        pass


def get_default_folder():
    data = _read_config()
    path = data.get("carpeta_datos", "")
    return os.path.normpath(path) if path else ""


def set_default_folder(path):
    data = _read_config()
    data["carpeta_datos"] = os.path.normpath(path)
    _write_config(data)


def get_data_dir():
    data = _read_config()
    path = data.get("data_dir")
    if not path:
        path = os.path.join(get_app_dir(), "data")
    if not os.path.isabs(path):
        path = os.path.join(get_app_dir(), path)
    return os.path.normpath(path)


def set_data_dir(path):
    data = _read_config()
    data["data_dir"] = os.path.normpath(path)
    _write_config(data)


def get_db_config():
    data = _read_config()
    db = data.get("db", {})
    return {
        "host": str(db.get("host", "")).strip(),
        "port": str(db.get("port", "5432")).strip(),
        "name": str(db.get("name", "resamania")).strip(),
        "user": str(db.get("user", "resamania")).strip(),
        "password": str(db.get("password", "")),
    }


def set_db_config(host, port, name, user, password):
    data = _read_config()
    data["db"] = {
        "host": str(host).strip(),
        "port": str(port).strip(),
        "name": str(name).strip(),
        "user": str(user).strip(),
        "password": str(password),
    }
    _write_config(data)


def get_user_role():
    data = _read_config()
    role = str(data.get("user_role", "MANAGER")).upper()
    if role not in ("MANAGER", "STAFF"):
        role = "MANAGER"
    return role


def set_user_role(role):
    data = _read_config()
    data["user_role"] = role
    _write_config(data)


def is_feature_enabled(name: str, default=False):
    data = _read_config()
    flags = data.get("features", {}) if isinstance(data, dict) else {}
    value = flags.get(name, default)
    return bool(value)


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
        self._invalid_default_folder = False
        if self.folder_path and not os.path.exists(self.folder_path):
            self.folder_path = ""
            set_default_folder("")
            self._invalid_default_folder = True
        if not self.folder_path or self._invalid_default_folder:
            app_dir = get_app_dir()
            if self._folder_has_data(app_dir):
                self.folder_path = app_dir
                set_default_folder(app_dir)
                self._invalid_default_folder = False
        self.dataframes = {}
        self.raw_accesos = None
        self.resumen_df = None
        self.data_dir = ""
        self.state_store = None

        # Datos de prestamos
        self.prestamos_file = ""
        self.prestamos = []
        self.clientes_ext_file = ""
        self.clientes_ext = []
        self.prestamos_filtro_activo = False
        self.incidencias_socios_file = ""
        self.incidencias_socios = []
        self.incidencias_socios_filtro = "VISTO_PENDIENTE"
        self.incidencias_socios_encontrado = None
        self.incidencias_socios_filtro_codigo = None
        self.incidencias_filtro_estado = "TODAS"
        self.objetos_taquillas_file = ""
        self.objetos_taquillas = []
        self.objetos_taquillas_blink_job = None
        self.objetos_taquillas_blink_on = False
        self.bajas_file = ""
        self.bajas = []
        self.bajas_view = "TODOS"
        self.bajas_cliente_filter = ""
        self.bajas_impagos_set = set()
        self.suspensiones_file = ""
        self.suspensiones = []
        self.suspensiones_view = "ACTIVAS"
        self.suspensiones_cliente_filter = ""
        self.suspensiones_impagos_set = set()

        # Felicitaciones (persistencia anual)
        self.felicitaciones_file = ""
        self.felicitaciones_enviadas = {}
        self.avanza_fit_envios_file = ""
        self.avanza_fit_envios = {}

        # Parpadeo de botones
        self.blink_states = {}
        self.blink_interval_ms = 650
        self.user_role = get_user_role()
        self.staff_write_areas = {"prestamos", "objetos_taquillas"}
        self._last_allowed_tab = None
        self._suppress_tab_guard = False
        # Alias por compatibilidad: algunos botones usan el nombre antiguo.
        self.nuevo_prestamo = self._prestamos_nuevo_prestamo
        self.prestamos_last_event = None

        # Impagos (SQLite)
        self.impagos_db = None
        self.impagos_last_export = None
        self.impagos_view = tk.StringVar(value="actuales")

        # Staff
        self.staff_file = ""
        self.staff = []
        self.staff_menu = None
        self.staff_tree = None
        self.staff_filter_var = None

        # Incidencias Club
        self.incidencias_db = None
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
        self.incidencias_creador = None
        self.incidencias_pan_active = False

        # Salidas PMR autorizados/advertidos
        self.pmr_autorizados_file = ""
        self.pmr_autorizados = set()
        self.pmr_advertencias_file = ""
        self.pmr_advertencias = {}
        self.pmr_df_raw = None
        self.accesos_grupo_actual = None

        self.auto_refresh_interval_ms = 1800000
        self.auto_refresh_job = None
        self.auto_refresh_last_error = None
        self.auto_refresh_last_error_shown = None
        self.local_cleanup_attempts = 0
        self.db_lock_path = None
        self._prompted_exports = False
        self.auto_refresh_enabled = self._state_get("auto_refresh_enabled", False, None)

        self.create_widgets()
        if self._invalid_default_folder:
            messagebox.showwarning(
                "Carpeta no encontrada",
                "La carpeta guardada no existe en este equipo. Selecciona la carpeta de OneDrive."
            )
        if self.folder_path and os.path.exists(self.folder_path):
            self.data_dir = os.path.join(self.folder_path, "data")
            set_data_dir(self.data_dir)
            self._set_data_dir(self.data_dir, show_message=False)
            self.refresh_persistent_data(show_messages=False)
            if self.load_data():
                self._set_last_refresh()
        else:
            db_cfg = get_db_config()
            if db_cfg.get("host"):
                self.data_dir = get_data_dir()
                self._set_data_dir(self.data_dir, show_message=False)
                self.refresh_persistent_data(show_messages=False)
                if self._db_exports_available():
                    if self.load_data():
                        self._set_last_refresh()
                else:
                    self.after(200, self._prompt_exports_folder)
        self.after(500, self.update_blink_states)
        if self.auto_refresh_enabled:
            self._schedule_auto_refresh()

    def _role_label(self, area):
        labels = {
            "prestamos": "Prestamos",
            "incidencias_club": "Incidencias Club",
            "objetos_taquillas": "Objetos Taquillas",
            "manager": "Gestion",
        }
        return labels.get(area, "esta seccion")

    def _folder_has_data(self, path):
        if not path or not os.path.isdir(path):
            return False
        csvs = [
            "RESUMEN CLIENTE.csv",
            "ACCESOS.csv",
            # "FACTURAS Y VALES.csv",
            "IMPAGOS.csv",
        ]
        if any(os.path.exists(os.path.join(path, name)) for name in csvs):
            return True
        return os.path.isdir(os.path.join(path, "data"))

    def _can_write(self, area):
        if self.user_role == "MANAGER":
            return True
        return area in self.staff_write_areas

    def _require_write(self, area):
        if self._can_write(area):
            return True
        label = self._role_label(area)
        messagebox.showwarning("Permisos", f"Solo el MANAGER puede acceder a {label}.", parent=self)
        return False

    def _require_manager_access(self, feature="esta funcion"):
        if self.user_role == "MANAGER":
            return True
        messagebox.showwarning("Permisos", f"Solo el MANAGER puede acceder a {feature}.", parent=self)
        return False

    def _update_role_button(self):
        if hasattr(self, "role_button") and self.role_button:
            self.role_button.config(text=f"ROL: {self.user_role}")

    def _toggle_role(self):
        if not self._security_pin_ok():
            return
        self.user_role = "STAFF" if self.user_role == "MANAGER" else "MANAGER"
        set_user_role(self.user_role)
        self._update_role_button()
        self._apply_role_ui()
        messagebox.showinfo("Rol", f"Rol actual: {self.user_role}", parent=self)

    def _apply_role_ui(self):
        if self.user_role == "STAFF":
            if self.staff_menu:
                try:
                    self.staff_menu.entryconfig("GESTIONAR STAFF", state="disabled")
                except Exception:
                    pass
            for btn in [
                getattr(self, "btn_inc_cargar_mapas", None),
                getattr(self, "btn_inc_borrar_mapa", None),
                getattr(self, "btn_mapa_anterior", None),
                getattr(self, "btn_mapa_siguiente", None),
                getattr(self, "btn_inc_asignar_area", None),
                getattr(self, "btn_inc_editar_area", None),
                getattr(self, "btn_inc_listar_maquina", None),
                getattr(self, "btn_inc_info_maquinas", None),
                getattr(self, "btn_inc_crear_incidencia", None),
                getattr(self, "btn_inc_gestion_incidencias", None),
                getattr(self, "btn_inc_vista", None),
            ]:
                if btn:
                    btn.configure(state="disabled")
        else:
            if self.staff_menu:
                try:
                    self.staff_menu.entryconfig("GESTIONAR STAFF", state="normal")
                except Exception:
                    pass
            for btn in [
                getattr(self, "btn_inc_cargar_mapas", None),
                getattr(self, "btn_inc_borrar_mapa", None),
                getattr(self, "btn_mapa_anterior", None),
                getattr(self, "btn_mapa_siguiente", None),
                getattr(self, "btn_inc_asignar_area", None),
                getattr(self, "btn_inc_editar_area", None),
                getattr(self, "btn_inc_listar_maquina", None),
                getattr(self, "btn_inc_info_maquinas", None),
                getattr(self, "btn_inc_crear_incidencia", None),
                getattr(self, "btn_inc_gestion_incidencias", None),
                getattr(self, "btn_inc_vista", None),
            ]:
                if btn:
                    btn.configure(state="normal")

    def _state_get(self, key, default, file_path=None):
        store = getattr(self, "state_store", None)
        if store and store.use_postgres:
            try:
                return store.get(key, default)
            except Exception:
                return default
        if not file_path or not os.path.exists(file_path):
            return default
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default

    def _state_set(self, key, value, file_path=None):
        store = getattr(self, "state_store", None)
        if store and store.use_postgres:
            store.set(key, value)
            return
        if not file_path:
            return
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(value, f, ensure_ascii=False, indent=2)

    def _blob_ref(self, blob_id):
        return f"blob:{blob_id}"

    def _is_blob_ref(self, value):
        return isinstance(value, str) and value.startswith("blob:")

    def _blob_id_from_ref(self, value):
        return value.split(":", 1)[1] if self._is_blob_ref(value) else ""

    def _store_image_blob(self, ruta, allow_png=False):
        store = getattr(self, "state_store", None)
        if not store or not store.use_postgres:
            return ""
        ext = os.path.splitext(ruta)[1].lower()
        if ext not in (".jpg", ".jpeg") and not (allow_png and ext == ".png"):
            return ""
        try:
            img = Image.open(ruta)
            if img.mode != "RGB":
                img = img.convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=80, optimize=True)
            blob_id = store.put_blob(buf.getvalue(), content_type="image/jpeg")
            return self._blob_ref(blob_id)
        except Exception:
            return ""

    def create_widgets(self):
        top_frame = tk.Frame(self)
        top_frame.pack(pady=10, fill="x")
        top_frame.columnconfigure(0, weight=1, uniform="top")
        top_frame.columnconfigure(1, weight=1, uniform="top")
        top_frame.columnconfigure(2, weight=1, uniform="top")

        left_frame = tk.Frame(top_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=10)
        center_frame = tk.Frame(top_frame)
        center_frame.grid(row=0, column=1, sticky="nsew", padx=10)
        right_frame = tk.Frame(top_frame)
        right_frame.grid(row=0, column=2, sticky="nsew", padx=10)

        instrucciones = (
            "- Sistema centralizado en base de datos (PostgreSQL).\n"
            "- Solo el PC CLUB debe seleccionar la carpeta y actualizar los CSV una vez al día.\n\n"
            "   Archivos requeridos en la carpeta del PC CLUB:\n"
            "   - RESUMEN CLIENTE.csv\n"
            "   - ACCESOS.csv (intervalo de 4 semanas atrás)\n"
            # "   - FACTURAS Y VALES.csv (intervalo de 4 semanas atrás)\n"
            "   - IMPAGOS.csv (Exportar el día actual - Clientes con Incidente de Pago)\n\n"
            "- Todos los archivos deben ser del mismo día de exportación.\n"
            "- Pulsa el botón 'RECARGAR BD' para subir los CSV a la base de datos.\n"
            "- Los demás PCs no necesitan carpeta ni CSV; leen todo desde la base de datos.\n"
            "- Mapas: se admiten imágenes PNG y JPG/JPEG.\n"
            "- Reportes visuales: solo se admiten imágenes JPG/JPEG.\n"
        )
        self.instrucciones_text = instrucciones
        self.instrucciones_nota = (
            "NOTA: En el PC CLUB la carpeta se guarda por defecto. "
            "En el portátil/manager solo configura la base de datos."
        )

        # Logo Fitness Park (centro)
        logo_fp_path = get_logo_path("LogoFpark.png")
        if os.path.exists(logo_fp_path):
            img_fp = Image.open(logo_fp_path)
            img_fp = img_fp.resize((550, 140))
            self.logo_fp_img = ImageTk.PhotoImage(img_fp)
            tk.Label(center_frame, image=self.logo_fp_img).pack(expand=True)

        botones_frame = tk.Frame(self)
        botones_frame.pack(pady=5)
        tk.Button(botones_frame, text="INSTRUCCIONES", command=self.mostrar_instrucciones, bg="#ffeb3b", fg="black").pack(side=tk.LEFT, padx=10)
        self.btn_staff = tk.Button(botones_frame, text="STAFF", command=self.abrir_staff, bg="#9e9e9e", fg="black")
        self.btn_staff.pack(side=tk.LEFT, padx=10)
        self.role_button = tk.Button(
            botones_frame,
            text=f"ROL: {self.user_role}",
            command=self._toggle_role,
            bg="#e0e0e0",
            fg="black",
        )
        if is_feature_enabled("role_button", default=False):
            self.role_button.pack(side=tk.LEFT, padx=10)
        self.btn_select_folder = tk.Button(botones_frame, text="Seleccionar carpeta", command=self.select_folder)
        if is_feature_enabled("select_folder_button", default=False):
            self.btn_select_folder.pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="RUTAS DATOS", command=self.mostrar_rutas_datos).pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="CONFIG BD", command=self.abrir_config_db, bg="#bbdefb", fg="black").pack(
            side=tk.LEFT, padx=10
        )
        tk.Button(botones_frame, text="RECARGAR BD", command=self.recargar_bd).pack(side=tk.LEFT, padx=10)
        self.btn_auto_refresh = tk.Button(botones_frame, command=self.toggle_auto_refresh)
        self.btn_auto_refresh.pack(side=tk.LEFT, padx=6)
        self._update_auto_refresh_button()
        tk.Button(botones_frame, text="Exportar a Excel", command=self.exportar_excel).pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="INCIDENCIAS CLUB", command=self.ir_a_incidencias_club, bg="#424242", fg="white").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="GESTION CLIENTES", command=self.mostrar_gestion_clientes, bg="#c5e1a5", fg="black").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="PRESTAMOS", command=self.ir_a_prestamos, bg="#ffcc80", fg="black").pack(side=tk.LEFT, padx=10)
        self.lbl_last_refresh = tk.Label(botones_frame, text="Ultima recarga: -", fg="#666666")
        self.lbl_last_refresh.pack(side=tk.RIGHT, padx=10)
        self.staff_menu = tk.Menu(self, tearoff=0)
        self.staff_menu.add_command(label="GESTIONAR STAFF", command=self.abrir_gestion_staff)
        self.staff_menu.add_command(label="DIA DE ASUNTOS PROPIOS", command=self.enviar_asuntos_propios)
        self.staff_menu.add_command(label="SOLICITUD DE CAMBIO DE TURNO", command=self.enviar_cambio_turno)

        self.gestion_clientes_frame = tk.Frame(self)
        self.btn_enviar_felicitacion = tk.Button(
            self.gestion_clientes_frame,
            text="ENVIAR FELICITACION",
            command=self.enviar_felicitacion,
            fg="#b30000",
        )
        if is_feature_enabled("felicitacion_button", default=False):
            self.btn_enviar_felicitacion.pack(side=tk.LEFT, padx=5)
        self.btn_avanza_fit = tk.Button(
            self.gestion_clientes_frame,
            text="ENVIO MARTES AVANZA FIT",
            command=self.enviar_avanza_fit,
            fg="#b30000",
        )
        self.btn_avanza_fit.pack(side=tk.LEFT, padx=5)
        tk.Button(self.gestion_clientes_frame, text="EXTRAER ACCESOS", command=self.extraer_accesos, fg="#0066cc").pack(side=tk.LEFT, padx=5)
        tk.Button(self.gestion_clientes_frame, text="IR A PRESTAMOS", command=lambda: self.notebook.select(self.tabs.get("Prestamos")), bg="#ff9800", fg="black").pack(side=tk.LEFT, padx=5)
        tk.Button(self.gestion_clientes_frame, text="IR A IMPAGOS", command=self.ir_a_impagos, bg="#ff6b6b", fg="black").pack(side=tk.LEFT, padx=5)
        tk.Button(self.gestion_clientes_frame, text="INCIDENCIAS SOCIOS", command=self.ir_a_incidencias_socios, bg="#9e9e9e", fg="black").pack(side=tk.LEFT, padx=5)
        self.btn_objetos_taquillas = tk.Button(
            self.gestion_clientes_frame,
            text="OBJETOS TAQUILLAS",
            command=lambda: self.notebook.select(self.tabs.get("Objetos Taquillas")),
            bg="#ffcc80",
            fg="black",
        )
        self.btn_objetos_taquillas.pack(side=tk.LEFT, padx=5)
        tk.Button(
            self.gestion_clientes_frame,
            text="GESTION BAJAS",
            command=self.ir_a_gestion_bajas,
            bg="#b39ddb",
            fg="black",
        ).pack(side=tk.LEFT, padx=5)
        tk.Button(
            self.gestion_clientes_frame,
            text="GESTION SUSPENSIONES",
            command=self.ir_a_gestion_suspensiones,
            bg="#c8e6c9",
            fg="black",
        ).pack(side=tk.LEFT, padx=5)

        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("Arial", 9, "bold"), padding=[24, 6], anchor="w")
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=1, fill='both')
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        tab_colors = {
            "Wizville": "#6fa8dc",
            "Accesos": "#a4c2f4",
            "Salidas PMR No Autorizadas": "#c27ba0",
            "Accesos Dobles Ayer": "#fce4ec",
            # "Servicios": "#b6d7a8",
            # "Socios Ultimate": "#76a5af",
            # "Socios Yanga": "#93c47d",
            "Avanza Fit": "#ffd966",
            "Accesos Cliente": "#cfe2f3",
            "Prestamos": "#ffb74d",
            "Impagos": "#ff6b6b",
            "Incidencias Club": "#424242",
            "Incidencias Socios": "#9e9e9e",
            "Objetos Taquillas": "#ffe0b2",
            "Gestion Bajas": "#d1c4e9",
            "Gestion Suspensiones": "#c8e6c9",
            "Staff": "#9e9e9e",
        }
        self.tab_icons = {}

        self.tabs = {}
        ocultar_tabs = {
            "Salidas PMR No Autorizadas",
            "Accesos Dobles Ayer",
            # "Socios Ultimate",
            # "Socios Yanga",
            "Staff",
        }
        for tab_name in [
            "Wizville", "Accesos",
            "Salidas PMR No Autorizadas",
            "Accesos Dobles Ayer",
            # "Socios Ultimate", "Socios Yanga",
            "Avanza Fit", "Accesos Cliente", "Prestamos", "Impagos", "Incidencias Club", "Incidencias Socios", "Staff"
        ] + ["Objetos Taquillas", "Gestion Bajas", "Gestion Suspensiones"]:
            tab = ttk.Frame(self.notebook)
            color = tab_colors.get(tab_name, "#cccccc")
            icon = tk.PhotoImage(width=14, height=14)
            icon.put(color, to=(0, 0, 14, 14))
            self.tab_icons[tab_name] = icon
            self.notebook.add(tab, text=tab_name, image=icon, compound="left")
            self.tabs[tab_name] = tab
            if tab_name == "Accesos":
                self.create_accesos_tab(tab)
            # elif tab_name == "Servicios":
            #     self.create_servicios_tab(tab)
            elif tab_name == "Prestamos":
                self.create_prestamos_tab(tab)
            elif tab_name == "Incidencias Socios":
                self.create_incidencias_socios_tab(tab)
            elif tab_name == "Objetos Taquillas":
                self.create_objetos_taquillas_tab(tab)
            elif tab_name == "Gestion Bajas":
                self.create_bajas_tab(tab)
            elif tab_name == "Gestion Suspensiones":
                self.create_suspensiones_tab(tab)
            elif tab_name == "Staff":
                self.create_staff_tab(tab)
            elif tab_name == "Impagos":
                self.create_impagos_tab(tab)
            elif tab_name == "Incidencias Club":
                self.create_incidencias_tab(tab)
            else:
                self.create_table(tab)
            if tab_name in ocultar_tabs:
                self.notebook.hide(tab)
            if tab_name == "Impagos":
                self.notebook.hide(tab)
        self._apply_role_ui()

    def create_table(self, tab, parent=None):
        container_parent = parent or tab
        container = tk.Frame(container_parent)
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
        if not seleccion:
            row = tree.identify_row(event.y)
            if row:
                tree.selection_set(row)
                seleccion = (row,)
        if not seleccion:
            return
        item = seleccion[0]
        col = tree.identify_column(event.x)
        col_index = int(col.replace('#', '')) - 1
        if col_index < 0:
            return
        valores = tree.item(item).get('values', [])
        if col_index >= len(valores):
            return
        valor = valores[col_index]
        self.clipboard_clear()
        self.clipboard_append(str(valor))

    def _prestamos_build_menu(self):
        self.prestamos_menu = tk.Menu(self, tearoff=0)
        self.prestamos_menu.add_command(label="Copiar", command=self._prestamos_copy_selected)
        self.prestamos_menu.add_separator()
        self.prestamos_menu.add_command(label="Editar cliente externo", command=self.editar_cliente_manual)
        self.prestamos_menu.add_command(label="Check devuelto", command=self.marcar_devuelto)
        self.prestamos_menu.add_command(label="Enviar aviso email", command=self.enviar_aviso_prestamo)
        self.prestamos_menu.add_command(label="Enviar aviso Whatsapp", command=self.abrir_whatsapp_prestamo)
        self.prestamos_menu.add_command(label="Vista individual/colectiva", command=self.toggle_prestamos_vista)
        self.prestamos_menu.add_command(label="Eliminar registro", command=self.eliminar_prestamo)

    def mostrar_menu(self, event, menu, tree):
        tree.event_context = event
        menu.delete(0, "end")
        menu.add_command(label="Copiar", command=lambda: self.copiar_celda(tree.event_context, tree))
        accesos_tab = self.tabs.get("Accesos")
        if accesos_tab and tree is accesos_tab.tree and self.accesos_grupo_actual == "Salidas PMR Autorizados":
            menu.add_command(label="Quitar autorizacion", command=self.pmr_quitar_autorizado)
        menu.post(event.x_root, event.y_root)

    def create_accesos_tab(self, tab):
        header = tk.Frame(tab)
        header.pack(fill="x", pady=4)

        botones = [
            ("Salidas PMR No Autorizadas", "Salidas PMR No Autorizadas", "#f8b6c1", "black"),
            ("ACCESOS DOBLES DE AYER", "Accesos Dobles Ayer", "#fce4ec", "black"),
        ]
        for label, fuente, bg, fg in botones:
            tk.Button(
                header,
                text=label,
                command=lambda f=fuente: self._mostrar_grupo("Accesos", f),
                bg=bg,
                fg=fg,
            ).pack(side="left", padx=5)

        self.pmr_tools_frame = tk.Frame(tab)
        self.pmr_tools_visible = False
        tk.Button(
            self.pmr_tools_frame,
            text="Agregar autorizado",
            command=self.pmr_agregar_autorizado,
            bg="#b0bec5",
            fg="black",
        ).pack(side="left", padx=5)
        tk.Button(
            self.pmr_tools_frame,
            text="Advertir WhatsApp",
            command=self.pmr_advertir_whatsapp,
            bg="#81c784",
            fg="black",
        ).pack(side="left", padx=5)
        tk.Button(
            self.pmr_tools_frame,
            text="Enviar aviso email",
            command=self.pmr_enviar_email,
            bg="#e57373",
            fg="black",
        ).pack(side="left", padx=5)
        tk.Button(
            self.pmr_tools_frame,
            text="Ver advertidos",
            command=self.pmr_mostrar_advertidos,
            bg="#64b5f6",
            fg="black",
        ).pack(side="left", padx=5)
        tk.Button(
            self.pmr_tools_frame,
            text="Ver autorizados",
            command=self.pmr_mostrar_autorizados,
            bg="#ffe082",
            fg="black",
        ).pack(side="left", padx=5)

        content = tk.Frame(tab)
        content.pack(expand=True, fill="both")
        self.create_table(tab, parent=content)

    def create_servicios_tab(self, tab):
        header = tk.Frame(tab)
        header.pack(fill="x", pady=4)

        botones = []

        content = tk.Frame(tab)
        content.pack(expand=True, fill="both")
        self.create_table(tab, parent=content)

    def _mostrar_grupo(self, tab_name, fuente):
        df = self.dataframes.get(fuente)
        if df is None:
            messagebox.showwarning("Sin datos", f"No hay datos cargados para {fuente}.")
            return
        if tab_name == "Accesos":
            self.accesos_grupo_actual = fuente
            if fuente == "Salidas PMR No Autorizadas":
                if not self.pmr_tools_visible:
                    self.pmr_tools_frame.pack(fill="x", pady=2)
                    self.pmr_tools_visible = True
            else:
                if self.pmr_tools_visible:
                    self.pmr_tools_frame.pack_forget()
                    self.pmr_tools_visible = False
        self.mostrar_en_tabla(tab_name, df)

    def cargar_pmr_autorizados(self):
        try:
            data = self._state_get("pmr_autorizados", [], self.pmr_autorizados_file)
            if isinstance(data, list):
                self.pmr_autorizados = {str(x).strip() for x in data if str(x).strip()}
        except Exception:
            self.pmr_autorizados = set()

    def guardar_pmr_autorizados(self):
        self._state_set("pmr_autorizados", sorted(self.pmr_autorizados), self.pmr_autorizados_file)

    def cargar_pmr_advertencias(self):
        try:
            data = self._state_get("pmr_advertencias", [], self.pmr_advertencias_file)
            if isinstance(data, list):
                advertencias = {}
                for item in data:
                    codigo = str(item.get("codigo", "")).strip()
                    if codigo:
                        advertencias[codigo] = item
                self.pmr_advertencias = advertencias
        except Exception:
            self.pmr_advertencias = {}

    def guardar_pmr_advertencias(self):
        data = list(self.pmr_advertencias.values())
        self._state_set("pmr_advertencias", data, self.pmr_advertencias_file)

    def _pmr_norm(self, text):
        t = unicodedata.normalize("NFD", str(text or "")).upper().strip()
        return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")

    def _pmr_get_col_index(self, columns, desired):
        desired_norm = self._pmr_norm(desired)
        for idx, col in enumerate(columns):
            if self._pmr_norm(col) == desired_norm:
                return idx
        return None

    def _pmr_is_reincidente(self, codigo):
        item = self.pmr_advertencias.get(str(codigo).strip())
        if not item:
            return False
        fecha = str(item.get("fecha", "")).strip()
        if not fecha:
            return True
        try:
            fecha_dt = datetime.strptime(fecha[:10], "%Y-%m-%d").date()
            return fecha_dt < datetime.now().date()
        except Exception:
            return True

    def _pmr_filtrar_pendientes(self, df):
        if df is None or df.empty:
            return df
        colmap = {self._pmr_norm(c): c for c in df.columns}
        col_codigo = colmap.get("NUMERO DE CLIENTE") or colmap.get("NRO DE CLIENTE")
        if not col_codigo:
            return df
        codigos_aut = {str(x).strip() for x in self.pmr_autorizados}
        codigos_adv_hoy = set()
        hoy = datetime.now().date()
        for codigo, item in self.pmr_advertencias.items():
            fecha = str(item.get("fecha", "")).strip()
            if not fecha:
                continue
            try:
                fecha_dt = datetime.strptime(fecha[:10], "%Y-%m-%d").date()
                if fecha_dt == hoy:
                    codigos_adv_hoy.add(str(codigo).strip())
            except Exception:
                continue
        codigos = codigos_aut | codigos_adv_hoy
        series = df[col_codigo].astype(str).str.strip()
        filtrado = df[~series.isin(codigos)].copy()
        codigos_filtrado = filtrado[col_codigo].astype(str).str.strip()
        filtrado["Reincidente"] = codigos_filtrado.apply(lambda x: "SI" if self._pmr_is_reincidente(x) else "")
        return filtrado

    def _pmr_refrescar_listado(self):
        if self.pmr_df_raw is None:
            return
        pmr_filtrado = self._pmr_filtrar_pendientes(self.pmr_df_raw)
        self.dataframes["Salidas PMR No Autorizadas"] = pmr_filtrado
        if self.accesos_grupo_actual == "Salidas PMR No Autorizadas":
            self.mostrar_en_tabla("Accesos", pmr_filtrado)

    def pmr_agregar_autorizado(self):
        if self.accesos_grupo_actual != "Salidas PMR No Autorizadas":
            messagebox.showwarning("PMR", "Selecciona primero 'Salidas PMR No Autorizadas'.")
            return
        tree = self.tabs["Accesos"].tree
        sel = tree.selection()
        codigo = None
        if sel:
            columns = list(tree["columns"])
            idx_codigo = self._pmr_get_col_index(columns, "Numero de cliente")
            if idx_codigo is not None:
                codigo = str(tree.item(sel[0])["values"][idx_codigo]).strip()
        if not codigo:
            codigo = simpledialog.askstring("PMR", "Numero de cliente autorizado:", parent=self)
            if codigo is None:
                return
            codigo = codigo.strip()
        if not codigo:
            messagebox.showwarning("PMR", "No se ingreso numero de cliente.")
            return
        self.pmr_autorizados.add(codigo)
        self.guardar_pmr_autorizados()
        self._pmr_refrescar_listado()
        messagebox.showinfo("PMR", f"Cliente {codigo} marcado como autorizado.")

    def pmr_advertir_whatsapp(self):
        if self.accesos_grupo_actual != "Salidas PMR No Autorizadas":
            messagebox.showwarning("PMR", "Selecciona primero 'Salidas PMR No Autorizadas'.")
            return
        tree = self.tabs["Accesos"].tree
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("PMR", "Selecciona un cliente.")
            return
        columns = list(tree["columns"])
        idx_movil = self._pmr_get_col_index(columns, "Movil")
        idx_codigo = self._pmr_get_col_index(columns, "Numero de cliente")
        values = tree.item(sel[0])["values"]
        movil = ""
        codigo = ""
        if idx_movil is not None and idx_movil < len(values):
            movil = values[idx_movil]
        if idx_codigo is not None and idx_codigo < len(values):
            codigo = str(values[idx_codigo]).strip()
        movil = self._normalizar_movil(movil)
        if not movil:
            messagebox.showwarning("PMR", "El cliente no tiene movil.")
            return
        mensaje = (
            "\U0001F4E9 INFORMACION IMPORTANTE - FITNESS PARK VILLALOBOS\n\n"
            "Hola \U0001F44B\n\n"
            "LE INFORMAMOS QUE EL ACCESO Y LA SALIDA AL CLUB SE REALIZA POR LOS TORNOS\n\n"
            "\u26A0 EL USO DE LA PUERTA PARA PERSONAS CON MOVILIDAD REDUCIDA NO SE PERMITE A OTRAS PERSONAS "
            "QUE NO TENGAN ESTA CONDICION DE DIVERSIDAD FUNCIONAL \u26A0\n\n"
            "Asi ayudamos a mantener un entorno seguro y respetuoso para tod@s.\n\n"
            "Reglamento: https://www.fitnesspark.es/reglamento-interno.pdf\n\n"
            "Gracias por tu colaboracion.\n"
            "Fitness Park Villalobos - Juntos, mas fuertes."
        )
        url = f"https://wa.me/{movil}?text={urllib.parse.quote(mensaje)}"
        webbrowser.open(url)
        if codigo:
            ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
            self.pmr_advertencias[codigo] = {
                "codigo": codigo,
                "fecha": ahora,
                "canal": "whatsapp",
            }
            self.guardar_pmr_advertencias()
            self._pmr_refrescar_listado()

    def pmr_enviar_email(self):
        if self.accesos_grupo_actual != "Salidas PMR No Autorizadas":
            messagebox.showwarning("PMR", "Selecciona primero 'Salidas PMR No Autorizadas'.")
            return
        tree = self.tabs["Accesos"].tree
        rows = [tree.item(i)["values"] for i in tree.get_children()]
        columns = list(tree["columns"])
        idx_email = self._pmr_get_col_index(columns, "Correo electronico")
        idx_codigo = self._pmr_get_col_index(columns, "Numero de cliente")
        idx_nombre = self._pmr_get_col_index(columns, "Nombre")
        idx_apellidos = self._pmr_get_col_index(columns, "Apellidos")
        idx_movil = self._pmr_get_col_index(columns, "Movil")

        emails = []
        for r in rows:
            if idx_email is not None and idx_email < len(r):
                email = str(r[idx_email]).strip()
                if email:
                    emails.append(email)

        cuerpo = (
            "\U0001F4E9 INFORMACION IMPORTANTE - FITNESS PARK VILLALOBOS\n\n"
            "Hola \U0001F44B\n\n"
            "LE INFORMAMOS QUE EL ACCESO Y LA SALIDA AL CLUB SE REALIZA POR LOS TORNOS\n\n"
            "\u26A0 EL USO DE LA PUERTA PARA PERSONAS CON MOVILIDAD REDUCIDA NO SE PERMITE A OTRAS PERSONAS "
            "QUE NO TENGAN ESTA CONDICION DE DIVERSIDAD FUNCIONAL \u26A0\n\n"
            "Asi ayudamos a mantener un entorno seguro y respetuoso para tod@s.\n\n"
            "Reglamento completo: https://www.fitnesspark.es/reglamento-interno.pdf\n\n"
            "Gracias por tu colaboracion.\n"
            "Fitness Park Villalobos - Juntos, mas fuertes."
        )
        asunto = "Informacion importante - Fitness Park Villalobos"

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
            mail.Body = cuerpo
            mail.Display()
        except Exception as e:
            messagebox.showerror("Outlook", f"No se pudo crear el correo: {e}", parent=self)
            return

        if not messagebox.askyesno("Confirmacion", "Ha enviado el email de advertencia?", parent=self):
            return

        ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
        for r in rows:
            if idx_codigo is None or idx_codigo >= len(r):
                continue
            codigo = str(r[idx_codigo]).strip()
            if not codigo:
                continue
            self.pmr_advertencias[codigo] = {
                "codigo": codigo,
                "fecha": ahora,
                "canal": "email",
                "nombre": r[idx_nombre] if idx_nombre is not None and idx_nombre < len(r) else "",
                "apellidos": r[idx_apellidos] if idx_apellidos is not None and idx_apellidos < len(r) else "",
                "email": r[idx_email] if idx_email is not None and idx_email < len(r) else "",
                "movil": r[idx_movil] if idx_movil is not None and idx_movil < len(r) else "",
            }
        self.guardar_pmr_advertencias()
        self._pmr_refrescar_listado()

    def pmr_mostrar_advertidos(self):
        if not self.pmr_advertencias:
            messagebox.showinfo("PMR", "No hay clientes advertidos.")
            return
        data = []
        for item in self.pmr_advertencias.values():
            data.append(
                {
                    "Numero de cliente": item.get("codigo", ""),
                    "Nombre": item.get("nombre", ""),
                    "Apellidos": item.get("apellidos", ""),
                    "Correo electronico": item.get("email", ""),
                    "Movil": item.get("movil", ""),
                    "Fecha aviso": item.get("fecha", ""),
                    "Canal": item.get("canal", ""),
                }
            )
        df = pd.DataFrame(data)
        self.accesos_grupo_actual = "Salidas PMR Advertidos"
        if self.pmr_tools_visible:
            self.pmr_tools_frame.pack(fill="x", pady=2)
        self.mostrar_en_tabla("Accesos", df)

    def pmr_mostrar_autorizados(self):
        if not self.pmr_autorizados:
            messagebox.showinfo("PMR", "No hay clientes autorizados.")
            return
        data = []
        for codigo in sorted(self.pmr_autorizados):
            row = {
                "Numero de cliente": codigo,
                "Nombre": "",
                "Apellidos": "",
                "Correo electronico": "",
                "Movil": "",
            }
            if self.pmr_df_raw is not None and not self.pmr_df_raw.empty:
                colmap = {self._pmr_norm(c): c for c in self.pmr_df_raw.columns}
                col_codigo = colmap.get("NUMERO DE CLIENTE") or colmap.get("NRO DE CLIENTE")
                if col_codigo:
                    match = self.pmr_df_raw[self.pmr_df_raw[col_codigo].astype(str).str.strip() == codigo]
                    if not match.empty:
                        rec = match.iloc[0]
                        row["Nombre"] = rec.get(colmap.get("NOMBRE", ""), "")
                        row["Apellidos"] = rec.get(colmap.get("APELLIDOS", ""), "")
                        row["Correo electronico"] = rec.get(colmap.get("CORREO ELECTRONICO", ""), "")
                        row["Movil"] = rec.get(colmap.get("MOVIL", ""), "")
            if self.resumen_df is not None and self.resumen_df is not None:
                colmap2 = {self._pmr_norm(c): c for c in self.resumen_df.columns}
                col_codigo2 = colmap2.get("NUMERO DE CLIENTE") or colmap2.get("NRO DE CLIENTE")
                if col_codigo2 and not row["Nombre"]:
                    match2 = self.resumen_df[self.resumen_df[col_codigo2].astype(str).str.strip() == codigo]
                    if not match2.empty:
                        rec2 = match2.iloc[0]
                        row["Nombre"] = rec2.get(colmap2.get("NOMBRE", ""), row["Nombre"])
                        row["Apellidos"] = rec2.get(colmap2.get("APELLIDOS", ""), row["Apellidos"])
                        row["Correo electronico"] = rec2.get(colmap2.get("CORREO ELECTRONICO", ""), row["Correo electronico"])
                        row["Movil"] = rec2.get(colmap2.get("MOVIL", ""), row["Movil"])
            data.append(row)
        df = pd.DataFrame(data)
        self.accesos_grupo_actual = "Salidas PMR Autorizados"
        if not self.pmr_tools_visible:
            self.pmr_tools_frame.pack(fill="x", pady=2)
            self.pmr_tools_visible = True
        self.mostrar_en_tabla("Accesos", df)

    def pmr_quitar_autorizado(self):
        if self.accesos_grupo_actual != "Salidas PMR Autorizados":
            return
        tree = self.tabs["Accesos"].tree
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("PMR", "Selecciona un cliente autorizado.")
            return
        columns = list(tree["columns"])
        idx_codigo = self._pmr_get_col_index(columns, "Numero de cliente")
        if idx_codigo is None:
            messagebox.showwarning("PMR", "No se encontro la columna de numero de cliente.")
            return
        codigo = str(tree.item(sel[0])["values"][idx_codigo]).strip()
        if not codigo:
            return
        if codigo in self.pmr_autorizados:
            self.pmr_autorizados.remove(codigo)
            self.guardar_pmr_autorizados()
            self.pmr_mostrar_autorizados()
            self._pmr_refrescar_listado()

    def _staff_get_value(self, item, *keys):
        for key in keys:
            if key in item and item.get(key) is not None:
                return item.get(key)
        return ""

    def _staff_normalize_item(self, raw_item):
        item = dict(raw_item) if isinstance(raw_item, dict) else {}
        normalized = {
            "id": item.get("id") or uuid.uuid4().hex,
            "nombre": self._staff_get_value(item, "nombre", "Nombre", "NOMBRE"),
            "apellido1": self._staff_get_value(item, "apellido1", "Apellido1", "APELLIDO1", "Apellido 1", "ap1"),
            "apellido2": self._staff_get_value(item, "apellido2", "Apellido2", "APELLIDO2", "Apellido 2", "ap2"),
            "movil": self._staff_get_value(item, "movil", "Movil", "M?vil", "telefono", "Tel?fono", "Telefono", "MOVIL"),
            "email": self._staff_get_value(item, "email", "Email", "correo", "Correo", "correo electronico", "Correo electr?nico"),
        }
        normalized["movil"] = self._normalizar_movil(normalized["movil"])
        return normalized

    def cargar_staff(self):
        self.staff = []
        try:
            data = self._state_get("staff", [], self.staff_file)
            items = []
            if isinstance(data, list):
                items = data
            elif isinstance(data, dict):
                if isinstance(data.get("staff"), list):
                    items = data.get("staff", [])
                elif all(isinstance(v, dict) for v in data.values()):
                    items = list(data.values())
            self.staff = [self._staff_normalize_item(i) for i in items if isinstance(i, dict)]
        except Exception:
            self.staff = []

    def guardar_staff(self):
        self._state_set("staff", self.staff, self.staff_file)

    def _staff_display_name(self, item):
        nombre = str(item.get("nombre", "")).strip()
        ap1 = str(item.get("apellido1", "")).strip()
        return " ".join([p for p in [nombre, ap1] if p]).strip()

    def _staff_label(self, item):
        base = self._staff_display_name(item)
        email = str(item.get("email", "")).strip()
        if email:
            return f"{base} ({email})"
        return base

    def _staff_select(self, title, prompt):
        if not self.staff:
            messagebox.showwarning("Staff", "No hay staff registrado.")
            return None
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text=prompt).pack(padx=10, pady=6)
        opciones = [self._staff_label(s) for s in self.staff]
        var = tk.StringVar(value=opciones[0] if opciones else "")
        combo = ttk.Combobox(win, values=opciones, textvariable=var, state="readonly", width=40)
        combo.pack(padx=10, pady=6)
        combo.focus_set()
        res = {"item": None}

        def aceptar():
            sel = var.get()
            for s in self.staff:
                if self._staff_label(s) == sel:
                    res["item"] = s
                    break
            win.destroy()

        btn = tk.Button(win, text="Aceptar", command=aceptar)
        btn.pack(pady=6)
        win.bind("<Return>", lambda _e: aceptar())
        win.bind("<Escape>", lambda _e: win.destroy())
        self._incidencias_center_window(win)
        win.grab_set()
        win.wait_window()
        return res["item"]

    def mostrar_instrucciones(self):
        if not self._require_manager_access("Instrucciones"):
            return
        texto = self.instrucciones_text if hasattr(self, "instrucciones_text") else ""
        nota = self.instrucciones_nota if hasattr(self, "instrucciones_nota") else ""
        win = tk.Toplevel(self)
        win.title("Instrucciones")
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text="INSTRUCCIONES PARA USO CORRECTO:", font=("Arial", 10, "bold")).pack(padx=12, pady=(10, 4))
        body = tk.Text(win, width=70, height=12, wrap="word")
        body.pack(padx=12, pady=(0, 6))
        body.insert("1.0", texto)
        if nota:
            body.insert("end", "\n" + nota)
        body.configure(state="disabled")
        btn = tk.Button(win, text="Cerrar", command=win.destroy)
        btn.pack(pady=(0, 10))
        self._incidencias_center_window(win)
        win.grab_set()

    def mostrar_gestion_clientes(self):
        if not hasattr(self, "gestion_clientes_frame"):
            return
        if not self.gestion_clientes_frame.winfo_ismapped():
            self.gestion_clientes_frame.pack(pady=4, fill="x", before=self.notebook)
        if self.user_role == "MANAGER":
            tab = self.tabs.get("Incidencias Socios")
            if tab:
                self.notebook.select(tab)
        else:
            tab = self.tabs.get("Prestamos")
            if tab:
                self.notebook.select(tab)

    def ocultar_gestion_clientes(self):
        if hasattr(self, "gestion_clientes_frame") and self.gestion_clientes_frame.winfo_ismapped():
            self.gestion_clientes_frame.pack_forget()

    def ir_a_incidencias_club(self):
        self.ocultar_gestion_clientes()
        tab = self.tabs.get("Incidencias Club")
        if tab:
            self.notebook.select(tab)
            # Vista por defecto: pendientes y vistas.
            self.incidencias_filtro_estado = "VISTO_PENDIENTE"
            if hasattr(self, "incidencias_canvas") and self.incidencias_canvas:
                self.incidencias_cargar_listado_mapas()
            self.after(0, self.incidencias_info_maquinas)

    def ir_a_prestamos(self):
        self.ocultar_gestion_clientes()
        tab = self.tabs.get("Prestamos")
        if tab:
            self.notebook.select(tab)

    def ir_a_incidencias_socios(self):
        if not self._require_manager_access("Incidencias socios"):
            return
        tab = self.tabs.get("Incidencias Socios")
        if tab:
            self.incidencias_socios_filtro = "VISTO_PENDIENTE"
            self.incidencias_socios_filtro_codigo = None
            self.refrescar_incidencias_socios_tree()
            self.notebook.select(tab)

    def ir_a_gestion_bajas(self):
        if not self._security_pin_ok():
            return
        if not self._require_manager_access("Gestion bajas"):
            return
        tab = self.tabs.get("Gestion Bajas")
        if tab:
            self.notebook.select(tab)

    def ir_a_gestion_suspensiones(self):
        if not self._security_pin_ok():
            return
        if not self._require_manager_access("Gestion suspensiones"):
            return
        tab = self.tabs.get("Gestion Suspensiones")
        if tab:
            self.notebook.select(tab)

    def select_folder(self):
        if not self._require_manager_access("Seleccionar carpeta"):
            return
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            set_default_folder(folder)
            data_dir = os.path.join(folder, "data")
            set_data_dir(data_dir)
            self._set_data_dir(data_dir, show_message=True)
            self._cleanup_local_data_dir()
            self.refresh_persistent_data(show_messages=True)
            self.refresh_all_data(show_messages=True)

    def _set_data_dir(self, data_dir, show_message=False):
        self.data_dir = os.path.normpath(data_dir)
        try:
            os.makedirs(self.data_dir, exist_ok=True)
        except Exception as e:
            fallback_dir = os.path.join(get_app_dir(), "data")
            if show_message:
                messagebox.showerror("Carpeta datos", f"No se pudo crear la carpeta de datos:\n{self.data_dir}\n\nSe usara:\n{fallback_dir}\n\nDetalle: {e}", parent=self)
            self.data_dir = fallback_dir
            os.makedirs(self.data_dir, exist_ok=True)

        self.prestamos_file = os.path.join(self.data_dir, "prestamos.json")
        self.clientes_ext_file = os.path.join(self.data_dir, "clientes_ext.json")
        self.incidencias_socios_file = os.path.join(self.data_dir, "incidencias_socios.json")
        self.objetos_taquillas_file = os.path.join(self.data_dir, "objetos_taquillas.json")
        self.bajas_file = os.path.join(self.data_dir, "bajas.json")
        self.suspensiones_file = os.path.join(self.data_dir, "suspensiones.json")
        self.felicitaciones_file = os.path.join(self.data_dir, "felicitaciones.json")
        self.avanza_fit_envios_file = os.path.join(self.data_dir, "avanza_fit_envios.json")
        db_config = get_db_config()
        self.impagos_db = ImpagosDB(os.path.join(self.data_dir, "impagos.db"), db_config=db_config)
        self.staff_file = os.path.join(self.data_dir, "staff.json")
        self.incidencias_db = IncidenciasDB(os.path.join(self.data_dir, "incidencias.db"), db_config=db_config)
        self.state_store = AppStateStore(db_config)
        self.pmr_autorizados_file = os.path.join(self.data_dir, "pmr_autorizados.json")
        self.pmr_advertencias_file = os.path.join(self.data_dir, "pmr_advertencias.json")
        if hasattr(self, "incidencias_canvas") and self.incidencias_canvas:
            self.incidencias_cargar_listado_mapas()

    def _cleanup_local_data_dir(self):
        local_dir = os.path.normpath(os.path.join(get_app_dir(), "data"))
        target_dir = os.path.normpath(self.data_dir or "")
        if not local_dir or not os.path.exists(local_dir):
            return
        if target_dir and os.path.normcase(local_dir) == os.path.normcase(target_dir):
            return
        def onerror(func, path, exc_info):
            try:
                os.chmod(path, 0o700)
                func(path)
            except Exception:
                pass

        try:
            shutil.rmtree(local_dir, onerror=onerror)
        except Exception:
            self.local_cleanup_attempts += 1
            if self.local_cleanup_attempts <= 3:
                self.after(1500, self._cleanup_local_data_dir)

    def mostrar_rutas_datos(self):
        if not self._require_manager_access("Rutas datos"):
            return
        config_path = CONFIG_PATH
        carpeta_csv = self.folder_path or "(no seleccionada)"
        data_dir = self.data_dir or "(no definida)"
        impagos_db_path = getattr(self.impagos_db, "db_path", "") if hasattr(self, "impagos_db") else ""
        incidencias_db_path = getattr(self.incidencias_db, "db_path", "") if hasattr(self, "incidencias_db") else ""
        db_cfg = get_db_config()
        db_host = db_cfg.get("host") or "(no configurado)"
        db_port = db_cfg.get("port") or ""
        db_name = db_cfg.get("name") or ""
        db_user = db_cfg.get("user") or ""
        maps_dir = os.path.join(self.data_dir, "maps") if self.data_dir else ""
        maps_count = None
        map_rows = []
        maps_error = ""
        try:
            map_rows = self.incidencias_db.list_maps() if hasattr(self, "incidencias_db") else []
            maps_count = len(map_rows)
        except Exception as e:
            maps_error = str(e)
        lines = [
            f"CONFIG: {config_path}",
            f"CSV: {carpeta_csv}",
            f"DATA: {data_dir}",
            f"CWD: {os.getcwd()}",
            f"IMPAGOS DB: {impagos_db_path}",
            f"INCIDENCIAS DB: {incidencias_db_path}",
            f"DB host: {db_host}",
            f"DB port: {db_port}",
            f"DB name: {db_name}",
            f"DB user: {db_user}",
            f"MAPS dir: {maps_dir}",
            f"CONFIG existe: {'SI' if os.path.exists(config_path) else 'NO'}",
            f"CSV existe: {'SI' if os.path.exists(carpeta_csv) else 'NO'}",
            f"DATA existe: {'SI' if os.path.exists(data_dir) else 'NO'}",
            f"IMPAGOS DB existe: {'SI' if impagos_db_path and os.path.exists(impagos_db_path) else 'NO'}",
            f"INCIDENCIAS DB existe: {'SI' if incidencias_db_path and os.path.exists(incidencias_db_path) else 'NO'}",
            f"MAPS existe: {'SI' if maps_dir and os.path.exists(maps_dir) else 'NO'}",
        ]
        if maps_error:
            lines.append(f"MAPS DB error: {maps_error}")
        else:
            lines.append(f"MAPS DB count: {maps_count}")
            if map_rows:
                first = map_rows[0]
                try:
                    _id, nombre, ruta, *_rest = first
                except Exception:
                    nombre, ruta = "", ""
                resolved = self._incidencias_resolve_map_path(ruta, nombre)
                lines.append(f"MAPA 1 nombre: {nombre}")
                lines.append(f"MAPA 1 ruta: {ruta}")
                lines.append(f"MAPA 1 resuelta: {resolved}")
                lines.append(f"MAPA 1 existe: {'SI' if resolved and os.path.exists(resolved) else 'NO'}")
        messagebox.showinfo("Rutas de datos", "\n".join(lines), parent=self)

    def abrir_config_db(self):
        if not self._security_pin_ok():
            return
        if not self._require_manager_access("Configurar BD"):
            return
        current = get_db_config()
        win = tk.Toplevel(self)
        win.title("Configurar BD")
        win.transient(self)
        win.resizable(False, False)

        form = tk.Frame(win)
        form.pack(padx=12, pady=10)

        tk.Label(form, text="Host/Servidor:").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        tk.Label(form, text="Puerto:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        tk.Label(form, text="Base de datos:").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        tk.Label(form, text="Usuario:").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        tk.Label(form, text="Contrasena:").grid(row=4, column=0, sticky="e", padx=6, pady=4)

        host_var = tk.StringVar(value=current.get("host", ""))
        port_var = tk.StringVar(value=current.get("port", "5432"))
        name_var = tk.StringVar(value=current.get("name", "resamania"))
        user_var = tk.StringVar(value=current.get("user", "resamania"))
        pass_var = tk.StringVar(value=current.get("password", ""))

        host_entry = tk.Entry(form, textvariable=host_var, width=30)
        port_entry = tk.Entry(form, textvariable=port_var, width=10)
        name_entry = tk.Entry(form, textvariable=name_var, width=30)
        user_entry = tk.Entry(form, textvariable=user_var, width=30)
        pass_entry = tk.Entry(form, textvariable=pass_var, width=30, show="*")

        host_entry.grid(row=0, column=1, padx=6, pady=4)
        port_entry.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        name_entry.grid(row=2, column=1, padx=6, pady=4)
        user_entry.grid(row=3, column=1, padx=6, pady=4)
        pass_entry.grid(row=4, column=1, padx=6, pady=4)

        nota = "Host puede ser el nombre del PC (ej: FPVillalobos)."
        tk.Label(win, text=nota).pack(padx=12, pady=(0, 6))

        def guardar():
            host = host_var.get().strip()
            port = port_var.get().strip() or "5432"
            name = name_var.get().strip()
            user = user_var.get().strip()
            password = pass_var.get()
            if not host:
                messagebox.showwarning("Config BD", "Introduce el host o servidor.", parent=win)
                return
            if not port.isdigit() or not (1 <= int(port) <= 65535):
                messagebox.showwarning("Config BD", "El puerto no es valido.", parent=win)
                return
            if not name:
                messagebox.showwarning("Config BD", "Introduce el nombre de la base de datos.", parent=win)
                return
            if not user:
                messagebox.showwarning("Config BD", "Introduce el usuario.", parent=win)
                return
            try:
                set_db_config(host, port, name, user, password)
                if self.data_dir:
                    db_cfg = get_db_config()
                    self.impagos_db = ImpagosDB(os.path.join(self.data_dir, "impagos.db"), db_config=db_cfg)
                    self.incidencias_db = IncidenciasDB(os.path.join(self.data_dir, "incidencias.db"), db_config=db_cfg)
                    self.state_store = AppStateStore(db_cfg)
                messagebox.showinfo("Config BD", "Configuracion guardada.", parent=win)
                win.destroy()
            except Exception as e:
                messagebox.showerror("Config BD", f"No se pudo guardar.\n\nDetalle: {e}", parent=win)

        def probar():
            host = host_var.get().strip()
            port = port_var.get().strip() or "5432"
            name = name_var.get().strip()
            user = user_var.get().strip()
            password = pass_var.get()
            if not host:
                messagebox.showwarning("Config BD", "Introduce el host o servidor.", parent=win)
                return
            try:
                port_num = int(port)
                if not (1 <= port_num <= 65535):
                    raise ValueError("Puerto fuera de rango")
            except Exception:
                messagebox.showwarning("Config BD", "El puerto no es valido.", parent=win)
                return
            if not name:
                messagebox.showwarning("Config BD", "Introduce el nombre de la base de datos.", parent=win)
                return
            if not user:
                messagebox.showwarning("Config BD", "Introduce el usuario.", parent=win)
                return
            try:
                import psycopg  # type: ignore
            except Exception:
                messagebox.showwarning("Config BD", "psycopg no esta instalado.", parent=win)
                return
            try:
                with psycopg.connect(
                    host=host,
                    port=port,
                    dbname=name,
                    user=user,
                    password=password,
                    connect_timeout=5,
                ) as conn:
                    with conn.cursor() as cur:
                        cur.execute("SELECT 1")
                        cur.fetchone()
                messagebox.showinfo("Config BD", "Conexion OK.", parent=win)
            except Exception as e:
                messagebox.showerror("Config BD", f"No se pudo conectar.\n\nDetalle: {e}", parent=win)

        btns = tk.Frame(win)
        btns.pack(pady=(0, 10))
        tk.Button(btns, text="Probar", width=10, command=probar).pack(side="left", padx=5)
        tk.Button(btns, text="Guardar", width=10, command=guardar).pack(side="left", padx=5)
        tk.Button(btns, text="Cancelar", width=10, command=win.destroy).pack(side="left", padx=5)
        win.bind("<Return>", lambda _e: guardar())
        win.bind("<Escape>", lambda _e: win.destroy())
        self._incidencias_center_window(win)
        win.grab_set()
        win.wait_window()

    def load_data(self, show_messages=True):
        if not self.folder_path:
            if not self._db_exports_available():
                if show_messages:
                    messagebox.showerror("Error", "No se ha seleccionado carpeta")
                return False

        try:
            t0 = time.perf_counter()
            if self.folder_path:
                t_sync_start = time.perf_counter()
                self._sync_exports_to_db()
                _log_timing("sync_exports_to_db", time.perf_counter() - t_sync_start)
            t_resumen_start = time.perf_counter()
            resumen = load_data_file(self.folder_path, "RESUMEN CLIENTE")
            _log_timing("load_csv_resumen", time.perf_counter() - t_resumen_start)
            self.resumen_df = resumen.copy()
            t_accesos_start = time.perf_counter()
            accesos = load_data_file(self.folder_path, "ACCESOS")
            _log_timing("load_csv_accesos", time.perf_counter() - t_accesos_start)
            self.raw_accesos = accesos.copy()
            t_impagos_start = time.perf_counter()
            incidencias = load_data_file(self.folder_path, "IMPAGOS")
            _log_timing("load_csv_impagos", time.perf_counter() - t_impagos_start)
            t_sync_impagos = time.perf_counter()
            self.sync_impagos(incidencias, show_messages=show_messages)
            _log_timing("sync_impagos", time.perf_counter() - t_sync_impagos)

            t_calc_wizville = time.perf_counter()
            self.mostrar_en_tabla("Wizville", procesar_wizville(resumen, accesos))
            _log_timing("calc_wizville_and_render", time.perf_counter() - t_calc_wizville)
            t_calc_pmr = time.perf_counter()
            pmr_df = procesar_salidas_pmr_no_autorizadas(resumen, accesos)
            self.pmr_df_raw = pmr_df.copy()
            pmr_filtrado = self._pmr_filtrar_pendientes(pmr_df)
            self.mostrar_en_tabla("Salidas PMR No Autorizadas", pmr_filtrado)
            _log_timing("calc_pmr_and_render", time.perf_counter() - t_calc_pmr)
            t_calc_dobles_ayer = time.perf_counter()
            dobles_ayer = procesar_accesos_dobles_ayer(resumen, accesos)
            self.mostrar_en_tabla("Accesos Dobles Ayer", dobles_ayer)
            _log_timing("calc_accesos_dobles_ayer_and_render", time.perf_counter() - t_calc_dobles_ayer)
            t_calc_avanza = time.perf_counter()
            self.mostrar_en_tabla("Avanza Fit", obtener_avanza_fit())
            _log_timing("calc_avanza_fit_and_render", time.perf_counter() - t_calc_avanza)

            self._mostrar_grupo("Accesos", "Salidas PMR No Autorizadas")
            self.update_blink_states()
            self._state_set("exports_last_loaded", self._get_exports_mtimes())
            _log_timing("load_data_total", time.perf_counter() - t0)
            if show_messages:
                messagebox.showinfo("Exito", "Datos cargados correctamente.")
            return True
        except Exception as e:
            if show_messages:
                messagebox.showerror("Error al cargar datos", str(e))
            else:
                self.auto_refresh_last_error = str(e)
            return False

    def _db_exports_available(self):
        store = getattr(self, "state_store", None)
        if not store or not store.use_postgres:
            return False
        for base in ("RESUMEN CLIENTE", "ACCESOS", "IMPAGOS"):
            meta = store.get(f"export:{base}", {})
            if not isinstance(meta, dict) or not str(meta.get("blob", "")).strip():
                return False
        return True

    def _get_exports_mtimes(self):
        store = getattr(self, "state_store", None)
        mtimes = {}
        for base in ("RESUMEN CLIENTE", "ACCESOS", "IMPAGOS"):
            mtime = None
            if self.folder_path:
                path = self._find_export_file(base)
                if path and os.path.exists(path):
                    mtime = os.path.getmtime(path)
            elif store and store.use_postgres:
                meta = store.get(f"export:{base}", {})
                if isinstance(meta, dict):
                    mtime = meta.get("mtime")
            mtimes[base] = mtime
        return mtimes

    def _exports_changed(self):
        current = self._get_exports_mtimes()
        if any(v is None for v in current.values()):
            return True
        last = self._state_get("exports_last_loaded", {}, None)
        if not isinstance(last, dict):
            return True
        return current != last

    def _find_export_file(self, base_name):
        if not self.folder_path:
            return ""
        candidates = [f"{base_name}.csv", f"{base_name}.xlsx"]
        for filename in candidates:
            full_path = os.path.join(self.folder_path, filename)
            if os.path.exists(full_path):
                return full_path
        return ""

    def _sync_exports_to_db(self):
        store = getattr(self, "state_store", None)
        if not store or not store.use_postgres:
            return
        for base in ("RESUMEN CLIENTE", "ACCESOS", "IMPAGOS"):
            path = self._find_export_file(base)
            if not path:
                continue
            try:
                mtime = os.path.getmtime(path)
                key = f"export:{base}"
                meta = store.get(key, {}) if hasattr(store, "get") else {}
                if isinstance(meta, dict) and meta.get("mtime") == mtime:
                    continue
                with open(path, "rb") as f:
                    data = f.read()
                blob_id = store.put_blob(data, content_type="application/octet-stream")
                old_blob = meta.get("blob") if isinstance(meta, dict) else ""
                if old_blob:
                    try:
                        store.delete_blob(old_blob)
                    except Exception:
                        pass
                store.set(
                    key,
                    {
                        "blob": blob_id,
                        "filename": os.path.basename(path),
                        "mtime": mtime,
                    },
                )
            except Exception:
                continue

    def _db_lock_file(self):
        if not self.data_dir:
            return ""
        if not self.db_lock_path:
            self.db_lock_path = os.path.join(self.data_dir, "db.lock")
        return self.db_lock_path

    def _acquire_db_lock(self, stale_seconds=120):
        lock_path = self._db_lock_file()
        if not lock_path:
            return False
        now = time.time()
        for _ in range(2):
            try:
                fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, "w") as f:
                    f.write(f"{os.getpid()}|{now}\n")
                return True
            except FileExistsError:
                try:
                    mtime = os.path.getmtime(lock_path)
                    if now - mtime > stale_seconds:
                        os.remove(lock_path)
                        continue
                except Exception:
                    pass
                return False
            except Exception:
                return False
        return False

    def _release_db_lock(self):
        lock_path = self._db_lock_file()
        if not lock_path:
            return
        try:
            os.remove(lock_path)
        except Exception:
            pass

    def refresh_persistent_data(self, show_messages=True):
        if not self.data_dir:
            return False
        try:
            self.cargar_clientes_ext()
            self.cargar_prestamos_json()
            self.cargar_incidencias_socios()
            self.cargar_objetos_taquillas()
            self.cargar_bajas()
            self.cargar_suspensiones()
            self.cargar_felicitaciones()
            self.cargar_avanza_fit_envios()
            self.cargar_staff()
            self.cargar_pmr_autorizados()
            self.cargar_pmr_advertencias()
            if hasattr(self, "incidencias_canvas") and self.incidencias_canvas:
                self.incidencias_cargar_listado_mapas()
                if self.incidencias_panel_mode == "incidencias":
                    self.incidencias_gestion_incidencias()
                elif self.incidencias_panel_mode == "machines":
                    self.incidencias_info_maquinas(self.incidencias_info_filter_area)
            return True
        except Exception as e:
            if show_messages:
                messagebox.showerror("Error al cargar datos", str(e))
            return False

    def refresh_all_data(self, show_messages=True):
        if show_messages and not self._require_manager_access("Actualizar datos"):
            return False
        csv_ok = self.load_data(show_messages=show_messages)
        self.refresh_persistent_data(show_messages=show_messages)
        if csv_ok:
            self._set_last_refresh()
        return csv_ok

    def recargar_bd(self):
        if not self._auto_refresh_allowed():
            messagebox.showwarning(
                "Recargar BD",
                "Cierra las ventanas abiertas o termina la edicion actual antes de recargar.",
                parent=self,
            )
            return False
        self._bring_to_front()
        def run_refresh():
            if self.folder_path or self._db_exports_available():
                if self._exports_changed():
                    return self.refresh_all_data(show_messages=False)
                return self.refresh_persistent_data(show_messages=False)
            return self.refresh_persistent_data(show_messages=False)

        ok = self._with_loading("Recargando datos...", run_refresh)
        if ok is False:
            err = self.auto_refresh_last_error or "No se pudo recargar la base de datos."
            messagebox.showerror("Recargar BD", err, parent=self)
            return False
        self._set_last_refresh()
        messagebox.showinfo("Recargar BD", "Datos recargados.", parent=self)
        return True

    def _set_last_refresh(self):
        if hasattr(self, "lbl_last_refresh") and self.lbl_last_refresh:
            ts = datetime.now().strftime("%d/%m/%Y %H:%M")
            self.lbl_last_refresh.config(text=f"Ultima recarga: {ts}")

    def _update_auto_refresh_button(self):
        if hasattr(self, "btn_auto_refresh") and self.btn_auto_refresh:
            if self.auto_refresh_enabled:
                self.btn_auto_refresh.config(text="AUTO REFRESH: ON", bg="#66bb6a", fg="black")
            else:
                self.btn_auto_refresh.config(text="AUTO REFRESH: OFF", bg="#e0e0e0", fg="black")

    def toggle_auto_refresh(self):
        self.auto_refresh_enabled = not self.auto_refresh_enabled
        self._state_set("auto_refresh_enabled", self.auto_refresh_enabled)
        if self.auto_refresh_enabled:
            self._schedule_auto_refresh()
        else:
            if self.auto_refresh_job is not None:
                try:
                    self.after_cancel(self.auto_refresh_job)
                except Exception:
                    pass
                self.auto_refresh_job = None
        self._update_auto_refresh_button()

    def _with_loading(self, message, func):
        win = tk.Toplevel(self)
        win.title("Recargando")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)
        label = tk.Label(win, text=message, padx=20, pady=15)
        label.pack()
        self._incidencias_center_window(win)
        self.update_idletasks()
        try:
            return func()
        finally:
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()

    def _prompt_exports_folder(self):
        if self._prompted_exports:
            return
        self._prompted_exports = True
        if self.folder_path or self._db_exports_available():
            return
        if not messagebox.askyesno(
            "CSV necesarios",
            "No hay CSV cargados en la base de datos. Selecciona la carpeta de exportaciones para cargarlos.",
            parent=self,
        ):
            return
        self.select_folder()

    def _auto_refresh_allowed(self):
        for w in self.winfo_children():
            if isinstance(w, tk.Toplevel) and w.winfo_viewable():
                return False
        if getattr(self, "incidencias_mode", None):
            return False
        return True

    def _schedule_auto_refresh(self):
        if self.auto_refresh_job is not None:
            try:
                self.after_cancel(self.auto_refresh_job)
            except Exception:
                pass
        self.auto_refresh_job = self.after(self.auto_refresh_interval_ms, self._auto_refresh_tick)

    def _auto_refresh_tick(self):
        if not self._auto_refresh_allowed():
            self._schedule_auto_refresh()
            return
        if not self.folder_path:
            self._with_loading("Auto-recargando datos...", lambda: self.refresh_persistent_data(show_messages=False))
            self._set_last_refresh()
            self._schedule_auto_refresh()
            return
        ok = self._with_loading("Auto-recargando datos...", lambda: self.refresh_all_data(show_messages=False))
        if not ok:
            err = self.auto_refresh_last_error
            if err and err != self.auto_refresh_last_error_shown:
                messagebox.showerror("Error al cargar datos", err, parent=self)
                self.auto_refresh_last_error_shown = err
        else:
            self.auto_refresh_last_error = None
            self.auto_refresh_last_error_shown = None
            self._set_last_refresh()
        self._schedule_auto_refresh()

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
        tree.tag_configure("pmr_reincidente", background="#ffe0b2")

        for _, row in df.iterrows():
            valores = list(row)
            tag = ""
            if "D?as desde alta" in df.columns:
                if row["D?as desde alta"] == 16:
                    tag = "amarillo"
                elif row["D?as desde alta"] == 180:
                    tag = "rojo"
            if "Reincidente" in df.columns:
                try:
                    if bool(row["Reincidente"]):
                        tag = "pmr_reincidente"
                except Exception:
                    pass
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
        tk.Button(top, text="Editar cliente externo", command=self.editar_cliente_manual).pack(side="left", padx=5)
        tk.Button(top, text="AGREGAR CLIENTE EXTERNO", command=self.agregar_cliente_otro_fp, bg="#ff7043", fg="white").pack(side="left", padx=5)
        tk.Button(top, text="NUEVO PRESTAMO", command=self._prestamos_nuevo_prestamo, bg="#ffcc80", fg="black").pack(side="left", padx=5)
        self.lbl_info = tk.Label(top, text="", anchor="w")
        self.lbl_info.pack(side="left", padx=10)

        # Material prestado (entrada oculta, se usa en Nuevo Prestamo)
        self.prestamo_material = tk.Entry(frm)

        self.lbl_perdon = tk.Label(frm, text="", fg="red", font=("", 10, "bold"))
        self.lbl_perdon.pack(fill="x", padx=5, pady=2)

        cols = ["codigo", "nombre", "apellidos", "email", "movil", "material", "fecha", "devuelto", "prestado_por"]
        table = tk.Frame(frm)
        table.pack(fill="both", expand=True)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=cols, show="headings")
        tree["displaycolumns"] = cols
        for c in cols:
            heading = "Prestado por" if c == "prestado_por" else c.capitalize()
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
        tree.bind("<Button-1>", self._prestamos_store_event)
        tree.bind("<Button-3>", self._prestamos_context_menu)
        tree.bind("<ButtonRelease-3>", self._prestamos_context_menu)
        tree.bind("<Button-2>", self._prestamos_context_menu)
        tree.bind("<Shift-Button-1>", self._prestamos_context_menu)
        self.tree_prestamos = tree
        tab.tree = tree
        self._prestamos_build_menu()

    def _norm(self, text):
        t = unicodedata.normalize("NFD", str(text or "")).upper().strip()
        return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")

    def _parse_fecha_prestamo(self, valor):
        try:
            return datetime.strptime(str(valor), "%d/%m/%Y %H:%M")
        except Exception:
            return None

    def on_prestamo_doble_click(self, event):
        tree = getattr(self, "tree_prestamos", None)
        if tree is None:
            tree = event.widget
            self.tree_prestamos = tree
        if not tree.selection():
            return
        self.marcar_devuelto()

    def _prestamos_store_event(self, event):
        self.prestamos_last_event = event

    def _prestamos_copy_selected(self):
        tree = getattr(self, "tree_prestamos", None)
        if tree is None:
            return
        event = self.prestamos_last_event
        if event is None:
            class DummyEvent:
                x = 5
                y = 5
            event = DummyEvent()
        self.copiar_celda(event, tree)

    def _prestamos_context_menu(self, event):
        tree = event.widget
        if tree is None:
            return
        self.prestamos_last_event = event
        row = tree.identify_row(event.y)
        if row:
            tree.selection_set(row)
        if getattr(self, "prestamos_menu", None) is None:
            self._prestamos_build_menu()
        try:
            self.prestamos_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.prestamos_menu.grab_release()
        return "break"

    def _socios_parse_dt(self, valor):
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

    def _prestamos_set_cliente(self, cliente):
        self.prestamo_encontrado = cliente
        info_text = f"{cliente.get('nombre','')} {cliente.get('apellidos','')} | {cliente.get('email','')} | {cliente.get('movil','')}"
        if hasattr(self, "lbl_info"):
            self.lbl_info.config(text=info_text)
        if hasattr(self, "prestamo_codigo") and cliente.get("codigo"):
            self.prestamo_codigo.delete(0, tk.END)
            self.prestamo_codigo.insert(0, str(cliente.get("codigo")))
        if hasattr(self, "lbl_perdon"):
            self.lbl_perdon.config(text="")

    def _prestamos_registrar_cliente_manual(self, codigo=None):
        self._bring_to_front()
        if not codigo:
            codigo = self._incidencias_prompt_text("Cliente manual", "Numero de cliente:")
            if codigo is None or not str(codigo).strip():
                return None
            codigo = str(codigo).strip()

        if self.resumen_df is not None:
            colmap = {self._norm(c): c for c in self.resumen_df.columns}
            col_codigo = colmap.get("NUMERO DE CLIENTE") or colmap.get("NUMERO DE SOCIO")
            if col_codigo:
                if any(str(val).strip() == codigo for val in self.resumen_df[col_codigo].fillna("")):
                    messagebox.showwarning(
                        "No permitido",
                        "Este cliente ya existe en RESUMEN CLIENTE y no se puede agregar manualmente."
                    )
                    return None

        nombre = self._incidencias_prompt_text("Cliente manual", "Nombre:")
        if nombre is None or not nombre.strip():
            messagebox.showwarning("Dato requerido", "El nombre es obligatorio.")
            return None
        apellidos = self._incidencias_prompt_text("Cliente manual", "Apellidos:")
        if apellidos is None:
            return None
        email = self._incidencias_prompt_text("Cliente manual", "Email:")
        if email is None:
            return None
        movil = self._incidencias_prompt_text("Cliente manual", "Movil (opcional):")
        if movil is None:
            return None
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
        self._prestamos_set_cliente(cliente)
        messagebox.showinfo("Cliente agregado", f"Cliente {codigo} agregado correctamente.")
        return cliente

    def agregar_cliente_otro_fp(self):
        self._prestamos_registrar_cliente_manual()

    def editar_cliente_manual(self):
        codigo = self.prestamo_codigo.get().strip() if hasattr(self, "prestamo_codigo") else ""
        if not codigo:
            codigo = self._incidencias_prompt_text("Editar cliente", "Numero de cliente:")
            if codigo is None or not str(codigo).strip():
                return
            codigo = str(codigo).strip()

        if self.resumen_df is not None:
            colmap = {self._norm(c): c for c in self.resumen_df.columns}
            col_codigo = colmap.get("NUMERO DE CLIENTE") or colmap.get("NUMERO DE SOCIO")
            if col_codigo:
                if any(str(val).strip() == codigo for val in self.resumen_df[col_codigo].fillna("")):
                    messagebox.showwarning(
                        "No permitido",
                        "Este cliente existe en RESUMEN CLIENTE y no se puede editar manualmente."
                    )
                    return

        cliente_actual = next((c for c in self.clientes_ext if c.get("codigo") == codigo), None)
        if not cliente_actual:
            messagebox.showwarning("No encontrado", f"No hay cliente manual con codigo {codigo}.")
            return

        nombre = self._incidencias_prompt_text("Editar cliente", "Nombre:", cliente_actual.get("nombre", ""))
        if nombre is None or not nombre.strip():
            messagebox.showwarning("Dato requerido", "El nombre es obligatorio.")
            return
        apellidos = self._incidencias_prompt_text("Editar cliente", "Apellidos:", cliente_actual.get("apellidos", ""))
        if apellidos is None:
            return
        email = self._incidencias_prompt_text("Editar cliente", "Email:", cliente_actual.get("email", ""))
        if email is None:
            return
        movil = self._incidencias_prompt_text("Editar cliente", "Movil (opcional):", cliente_actual.get("movil", ""))
        if movil is None:
            return
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
        self._prestamos_set_cliente(cliente)
        messagebox.showinfo("Cliente actualizado", f"Cliente {codigo} actualizado correctamente.")

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
        self.lbl_info.config(text=info_text)

    def _prestamos_nuevo_prestamo(self):
        if not self.prestamo_encontrado:
            messagebox.showwarning("Sin cliente", "Busca un cliente primero.")
            return
        material = self._incidencias_prompt_text("Prestamo", "Material prestado:")
        if material is None or not material.strip():
            messagebox.showwarning("Dato requerido", "El material es obligatorio.")
            return
        staff = self._staff_select("Prestamo", "Quien presta el material?")
        if not staff:
            return
        prestado_por = self._staff_display_name(staff)
        codigo = str(self.prestamo_encontrado.get("codigo", "")).strip()
        pendientes = [p for p in self.prestamos if p.get("codigo") == codigo and not p.get("devuelto")]
        if len(pendientes) >= 2:
            resp = messagebox.askyesno(
                "Limite de prestamos",
                f"El cliente {codigo} ya tiene {len(pendientes)} prestamos sin devolver.\n"
                "Desea liberar los pendientes para permitir otro prestamo?"
            )
            if not resp:
                return
            for p in pendientes:
                p["devuelto"] = True
                p["liberado_pin"] = True
            self.guardar_prestamos_json()
            self.refrescar_prestamos_tree()
        prestamo = {
            "id": uuid.uuid4().hex,
            **self.prestamo_encontrado,
            "prestado_por": prestado_por or "",
            "material": material.strip(),
            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "devuelto": False,
            "liberado_pin": False,
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

    def nuevo_prestamo(self):
        self._prestamos_nuevo_prestamo()

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

    def eliminar_prestamo(self):
        sel = self.tree_prestamos.selection()
        if not sel:
            messagebox.showwarning("Sin seleccion", "Selecciona un prestamo.")
            return
        if not messagebox.askyesno("Eliminar", "Eliminar este prestamo?", parent=self):
            return
        iid = sel[0]
        eliminado = False
        for p in list(self.prestamos):
            if p.get("id") == iid:
                self.prestamos.remove(p)
                eliminado = True
                break
        if not eliminado:
            vals = self.tree_prestamos.item(iid)["values"]
            if len(vals) >= 7:
                codigo_sel, material_sel, fecha_sel = vals[0], vals[5], vals[6]
                for p in list(self.prestamos):
                    if p.get("codigo") == codigo_sel and p.get("material") == material_sel and p.get("fecha") == fecha_sel:
                        self.prestamos.remove(p)
                        eliminado = True
                        break
        if eliminado:
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
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook.\n"
                f"Cuerpo copiado al portapapeles para pegarlo manualmente.\n\nDetalle: {e}"
            )

    def abrir_whatsapp_prestamo(self):
        sel = self.tree_prestamos.selection()
        if not sel:
            messagebox.showwarning("Sin seleccion", "Selecciona un prestamo.")
            return
        datos = self.tree_prestamos.item(sel[0])["values"]
        movil = self._normalizar_movil(str(datos[4]))
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
            url = f"https://wa.me/{movil}?text={urllib.parse.quote(texto)}"
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("WhatsApp", f"No se pudo abrir el enlace: {e}")

    def cargar_prestamos_json(self):
        try:
            data = self._state_get("prestamos", [], self.prestamos_file)
            self.prestamos = data if isinstance(data, list) else []
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

    def cargar_incidencias_socios(self):
        try:
            data = self._state_get("incidencias_socios", [], self.incidencias_socios_file)
            self.incidencias_socios = data if isinstance(data, list) else []
            changed = False
            for inc in self.incidencias_socios:
                if "id" not in inc:
                    inc["id"] = uuid.uuid4().hex
                    changed = True
                if "gestion" not in inc:
                    inc["gestion"] = ""
                    changed = True
            if changed:
                self.guardar_incidencias_socios()
        except Exception:
            self.incidencias_socios = []
        if hasattr(self, "tree_incidencias_socios"):
            self.refrescar_incidencias_socios_tree()

    def guardar_incidencias_socios(self):
        self._state_set("incidencias_socios", self.incidencias_socios, self.incidencias_socios_file)

    def refrescar_incidencias_socios_tree(self):
        if not hasattr(self, "tree_incidencias_socios"):
            return
        tree = self.tree_incidencias_socios
        tree.delete(*tree.get_children())
        filtro = (self.incidencias_socios_filtro or "TODAS").upper()
        codigo_filtro = (self.incidencias_socios_filtro_codigo or "").strip()
        for inc in sorted(
            self.incidencias_socios,
            key=lambda i: self._socios_parse_dt(i.get("fecha")) or datetime.min,
            reverse=True,
        ):
            if codigo_filtro and str(inc.get("codigo", "")).strip() != codigo_filtro:
                continue
            estado = str(inc.get("estado", "PENDIENTE")).upper()
            if filtro == "VISTO_PENDIENTE":
                if estado not in ("VISTO", "PENDIENTE"):
                    continue
            elif filtro != "TODAS" and estado != filtro:
                continue
            reporte = "R" if inc.get("reporte_path") else ""
            tag = ""
            if estado == "PENDIENTE":
                tag = "pendiente"
            elif estado == "VISTO":
                tag = "visto"
            elif estado == "RESUELTO":
                tag = "resuelto"
            values = [
                inc.get("codigo", ""),
                inc.get("nombre", ""),
                inc.get("apellidos", ""),
                inc.get("email", ""),
                inc.get("movil", ""),
                inc.get("incidencia", ""),
                inc.get("gestion", ""),
                inc.get("fecha", ""),
                estado,
                reporte,
                inc.get("prestado_por", ""),
            ]
            tree.insert("", "end", iid=inc.get("id"), values=values, tags=(tag,))

    def buscar_cliente_incidencia_socio(self):
        codigo = self.incidencia_socios_codigo.get().strip()
        if not codigo:
            messagebox.showwarning("Sin numero", "Introduce el numero de cliente.")
            return
        cliente = None
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
                        "movil": self._normalizar_movil(row.get(col_movil, "")),
                    }
        if not cliente:
            ext = [c for c in self.clientes_ext if c.get("codigo") == codigo]
            if ext:
                cliente = ext[0].copy()
                cliente["movil"] = self._normalizar_movil(cliente.get("movil", ""))
        if not cliente:
            resp = messagebox.askyesno(
                "Sin cliente",
                f"No hay cliente {codigo}. Registrar manualmente?"
            )
            if not resp:
                return
            cliente = self._socios_registrar_cliente_manual(codigo)
            if not cliente:
                return
            self.nueva_incidencia_socio()
            return
        self.incidencias_socios_encontrado = cliente
        self.incidencias_socios_filtro_codigo = None
        info = f"{cliente.get('nombre','')} {cliente.get('apellidos','')} | {cliente.get('email','')} | {cliente.get('movil','')}"
        self.lbl_incidencias_socios_info.config(text=info)
        self._socios_mostrar_boton_historial(codigo)

    def _socios_registrar_cliente_manual(self, codigo=None):
        self._bring_to_front()
        if not codigo:
            codigo = self._pedir_campo("Numero de cliente", "Introduce el numero de cliente:")
            if codigo is None or not codigo.strip():
                return None
            codigo = codigo.strip()

        if self.resumen_df is not None:
            colmap = {self._norm(c): c for c in self.resumen_df.columns}
            col_codigo = colmap.get("NUMERO DE CLIENTE") or colmap.get("NUMERO DE SOCIO")
            if col_codigo:
                if any(str(val).strip() == codigo for val in self.resumen_df[col_codigo].fillna("")):
                    messagebox.showwarning(
                        "No permitido",
                        "Este cliente ya existe en RESUMEN CLIENTE y no se puede agregar manualmente."
                    )
                    return None

        nombre = self._pedir_campo("Nombre", "Introduce el nombre:", obligatorio=True, validar_nombre=True)
        if nombre is None:
            return None
        apellidos = self._pedir_campo("Apellidos", "Introduce los apellidos:")
        if apellidos is None:
            return None
        email = self._pedir_campo("Email", "Introduce el email:")
        if email is None:
            return None
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
        self.incidencias_socios_encontrado = cliente.copy()
        self.incidencias_socios_filtro_codigo = None
        info = f"{cliente.get('nombre','')} {cliente.get('apellidos','')} | {cliente.get('email','')} | {cliente.get('movil','')}"
        self.lbl_incidencias_socios_info.config(text=info)
        self._socios_mostrar_boton_historial(codigo)
        messagebox.showinfo("Cliente agregado", f"Cliente {codigo} anadido como externo.")
        return cliente

    def agregar_cliente_incidencia_socio(self):
        self._socios_registrar_cliente_manual()

    def nueva_incidencia_socio(self):
        if not getattr(self, "incidencias_socios_encontrado", None):
            messagebox.showwarning("Sin cliente", "Busca un cliente primero.")
            return
        codigo = self.incidencias_socios_encontrado.get("codigo")
        if codigo and any(str(i.get("codigo", "")).strip() == str(codigo).strip() for i in self.incidencias_socios):
            messagebox.showwarning("Aviso", "Este cliente ya ha tenido incidencias anteriormente.")
        staff = self._staff_select("Incidencia socio", "Quien registra la incidencia?")
        if not staff:
            return
        prestado_por = self._staff_display_name(staff)
        incidencia = self._incidencias_prompt_text("Incidencia socio", "Cual es la incidencia?")
        if incidencia is None:
            return
        reporte_path = None
        if messagebox.askyesno("Reporte visual", "Desea agregar reporte visual?", parent=self):
            reporte_path = self._incidencias_pedir_reporte_visual()
        cliente = self.incidencias_socios_encontrado
        inc = {
            "id": uuid.uuid4().hex,
            "codigo": cliente.get("codigo", ""),
            "nombre": cliente.get("nombre", ""),
            "apellidos": cliente.get("apellidos", ""),
            "email": cliente.get("email", ""),
            "movil": cliente.get("movil", ""),
            "incidencia": incidencia.strip(),
            "gestion": "",
            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "estado": "PENDIENTE",
            "reporte_path": reporte_path,
            "prestado_por": prestado_por,
        }
        self.incidencias_socios.append(inc)
        self.guardar_incidencias_socios()
        self.refrescar_incidencias_socios_tree()
        self._socios_mostrar_boton_historial(inc.get("codigo"))
        messagebox.showinfo("Guardado", "Incidencia registrada.")

    def _socios_selector_estado(self):
        self._bring_to_front()
        win = tk.Toplevel(self)
        win.title("Estado")
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text="Selecciona estado:").pack(padx=10, pady=5)
        opciones = ["PENDIENTE", "VISTO", "RESUELTO"]
        var = tk.StringVar(value=opciones[0])
        combo = ttk.Combobox(win, values=opciones, textvariable=var, state="readonly", width=20)
        combo.pack(padx=10, pady=5)
        combo.focus_set()
        res = {"val": None}

        def aceptar():
            res["val"] = var.get()
            win.destroy()

        tk.Button(win, text="Aceptar", command=aceptar).pack(pady=6)
        win.bind("<Return>", lambda _e: aceptar())
        win.bind("<Escape>", lambda _e: win.destroy())
        self._incidencias_center_window(win)
        win.grab_set()
        win.wait_window()
        return res["val"]

    def _socios_cambiar_estado(self, inc_id):
        if not inc_id:
            return
        estado = self._socios_selector_estado()
        if not estado:
            return
        for inc in self.incidencias_socios:
            if inc.get("id") == inc_id:
                inc["estado"] = estado
                break
        self.guardar_incidencias_socios()
        self.refrescar_incidencias_socios_tree()

    def _socios_editar_incidencia(self, inc_id):
        inc = next((i for i in self.incidencias_socios if i.get("id") == inc_id), None)
        if not inc:
            return
        self._bring_to_front()
        incidencia = self._incidencias_prompt_text("Incidencia", "Cual es la incidencia?", inc.get("incidencia", ""))
        if incidencia is None:
            return
        inc["incidencia"] = incidencia.strip()
        if messagebox.askyesno("Reporte visual", "Desea cambiar reporte visual?", parent=self):
            nuevo = self._incidencias_pedir_reporte_visual()
            if nuevo:
                inc["reporte_path"] = nuevo
        self.guardar_incidencias_socios()
        self.refrescar_incidencias_socios_tree()

    def _socios_modificar_gestion(self, inc_id):
        inc = next((i for i in self.incidencias_socios if i.get("id") == inc_id), None)
        if not inc:
            return
        self._bring_to_front()
        gestion = self._incidencias_prompt_text("Gestion", "Agregar/editar gestion:", inc.get("gestion", ""))
        if gestion is None:
            return
        inc["gestion"] = gestion.strip()
        self.guardar_incidencias_socios()
        self.refrescar_incidencias_socios_tree()

    def _socios_ver_reporte(self, inc_id):
        inc = next((i for i in self.incidencias_socios if i.get("id") == inc_id), None)
        if not inc:
            return
        reporte = inc.get("reporte_path")
        if not reporte:
            messagebox.showwarning("Reporte", "No hay reporte visual.")
            return
        resolved = self._incidencias_resolve_reporte_path(reporte)
        if resolved:
            stored = self._incidencias_store_reporte_path(resolved)
            if stored and stored != reporte:
                inc["reporte_path"] = stored
                self.guardar_incidencias_socios()
                self.refrescar_incidencias_socios_tree()
        self._incidencias_ver_reporte(resolved or reporte)

    def _socios_modificar_reporte(self, inc_id):
        inc = next((i for i in self.incidencias_socios if i.get("id") == inc_id), None)
        if not inc:
            return
        nuevo = self._incidencias_pedir_reporte_visual()
        if not nuevo:
            return
        inc["reporte_path"] = nuevo
        self.guardar_incidencias_socios()
        self.refrescar_incidencias_socios_tree()

    def _socios_abrir_chat(self, inc_id):
        inc = next((i for i in self.incidencias_socios if i.get("id") == inc_id), None)
        if not inc:
            return
        movil = self._normalizar_movil(inc.get("movil", ""))
        if not movil:
            messagebox.showwarning("Chat", "El cliente no tiene movil registrado.")
            return
        webbrowser.open(f"https://wa.me/{movil}")

    def _socios_enviar_email(self, inc_id):
        inc = next((i for i in self.incidencias_socios if i.get("id") == inc_id), None)
        if not inc:
            return
        email = str(inc.get("email", "")).strip()
        if not email:
            messagebox.showwarning("Email", "El cliente no tiene email registrado.")
            return
        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(email)
            messagebox.showwarning("Outlook no disponible", "Email copiado al portapapeles.")
            return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = ""
            mail.Body = ""
            mail.Display()
        except Exception as e:
            messagebox.showerror("Email", f"No se pudo abrir Outlook.\nDetalle: {e}")

    def _socios_eliminar_incidencia(self, inc_id):
        if not messagebox.askyesno("Eliminar", "Eliminar esta incidencia?", parent=self):
            return
        self.incidencias_socios = [i for i in self.incidencias_socios if i.get("id") != inc_id]
        self.guardar_incidencias_socios()
        self.refrescar_incidencias_socios_tree()

    def _socios_set_filtro(self, estado):
        self.incidencias_socios_filtro = estado
        self.refrescar_incidencias_socios_tree()

    def _socios_mostrar_boton_historial(self, codigo):
        if not hasattr(self, "btn_incidencias_socios_historial"):
            return
        if not codigo:
            self.btn_incidencias_socios_historial.pack_forget()
            return
        hay = any(str(i.get("codigo", "")).strip() == str(codigo).strip() for i in self.incidencias_socios)
        if hay:
            self.btn_incidencias_socios_historial.configure(command=lambda c=codigo: self._socios_filtrar_por_cliente(c))
            if not self.btn_incidencias_socios_historial.winfo_ismapped():
                self.btn_incidencias_socios_historial.pack(side="left", padx=5)
            self._socios_parpadear_boton(self.btn_incidencias_socios_historial)
        else:
            self.btn_incidencias_socios_historial.pack_forget()

    def _socios_parpadear_boton(self, btn):
        try:
            normal = btn.cget("background")
        except Exception:
            normal = None
        highlight = "#ffb74d"

        def toggle(count=0):
            if count >= 6:
                if normal is not None:
                    btn.configure(background=normal)
                return
            btn.configure(background=highlight if count % 2 == 0 else normal)
            self.after(200, lambda: toggle(count + 1))

        toggle()

    def _socios_filtrar_por_cliente(self, codigo):
        self.incidencias_socios_filtro_codigo = str(codigo).strip()
        self.refrescar_incidencias_socios_tree()

    def _socios_limpiar_filtro_cliente(self):
        self.incidencia_socios_codigo.delete(0, tk.END)
        self.incidencias_socios_encontrado = None
        self.incidencias_socios_filtro_codigo = None
        self.lbl_incidencias_socios_info.config(text="")
        if hasattr(self, "btn_incidencias_socios_historial"):
            self.btn_incidencias_socios_historial.pack_forget()
        self.refrescar_incidencias_socios_tree()

    def create_incidencias_socios_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        top = tk.Frame(frm)
        top.pack(fill="x", pady=4)
        tk.Label(top, text="Numero de cliente:").pack(side="left")
        self.incidencia_socios_codigo = tk.Entry(top, width=14)
        self.incidencia_socios_codigo.pack(side="left", padx=5)
        tk.Button(top, text="Buscar", command=self.buscar_cliente_incidencia_socio).pack(side="left", padx=5)
        tk.Button(top, text="AGREGAR NUEVO CLIENTE", command=self.agregar_cliente_incidencia_socio, bg="#ff7043", fg="white").pack(
            side="left", padx=5
        )
        tk.Button(top, text="Nueva incidencia", command=self.nueva_incidencia_socio, bg="#ffcc80", fg="black").pack(side="left", padx=5)
        vista_btn = tk.Menubutton(top, text="VISTA", bg="#bdbdbd", fg="black")
        vista_menu = tk.Menu(vista_btn, tearoff=0)
        vista_menu.add_command(label="TODAS LAS INCIDENCIAS", command=lambda: self._socios_set_filtro("TODAS"))
        vista_menu.add_command(label="PENDIENTES", command=lambda: self._socios_set_filtro("PENDIENTE"))
        vista_menu.add_command(label="VISTAS", command=lambda: self._socios_set_filtro("VISTO"))
        vista_menu.add_command(label="RESUELTAS", command=lambda: self._socios_set_filtro("RESUELTO"))
        vista_menu.add_command(label="VISTAS Y PENDIENTES", command=lambda: self._socios_set_filtro("VISTO_PENDIENTE"))
        vista_btn.configure(menu=vista_menu)
        vista_btn.pack(side="left", padx=5)
        tk.Button(top, text="LIMPIAR", command=self._socios_limpiar_filtro_cliente).pack(side="left", padx=5)
        self.btn_incidencias_socios_historial = tk.Button(top, text="VER INCIDENCIAS ANTERIORES", bg="#ffcc80", fg="black")
        self.lbl_incidencias_socios_info = tk.Label(top, text="", anchor="w")
        self.lbl_incidencias_socios_info.pack(side="left", padx=10)

        cols = ["codigo", "nombre", "apellidos", "email", "movil", "incidencia", "gestion", "fecha", "estado", "reporte", "prestado_por"]
        table = tk.Frame(frm)
        table.pack(fill="both", expand=True)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=cols, show="headings")
        for c in cols:
            heading = "Registrado por" if c == "prestado_por" else c.capitalize()
            tree.heading(c, text=heading)
            width = 120
            if c in ("email",):
                width = 200
            elif c in ("incidencia", "gestion"):
                width = 260
            elif c in ("prestado_por",):
                width = 160
            tree.column(c, anchor="center", width=width, stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")
        vscroll = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        tree.bind("<Control-c>", lambda e: self.copiar_celda(e, tree))
        tree.tag_configure("pendiente", background="#ffe6cc")
        tree.tag_configure("visto", background="#fff3cd")
        tree.tag_configure("resuelto", background="#d4edda")

        def on_right_click(event):
            row = tree.identify_row(event.y)
            if not row:
                return
            tree.selection_set(row)
            menu = tk.Menu(tree, tearoff=0)
            menu.add_command(label="Modificar incidencia", command=lambda: self._socios_editar_incidencia(row))
            menu.add_command(label="Modificar gestion", command=lambda: self._socios_modificar_gestion(row))
            menu.add_command(label="Cambiar estado", command=lambda: self._socios_cambiar_estado(row))
            menu.add_command(label="Ver reporte visual", command=lambda: self._socios_ver_reporte(row))
            menu.add_command(label="Modificar reporte visual", command=lambda: self._socios_modificar_reporte(row))
            menu.add_command(label="Eliminar incidencia", command=lambda: self._socios_eliminar_incidencia(row))
            menu.add_separator()
            menu.add_command(label="Abrir chat", command=lambda: self._socios_abrir_chat(row))
            menu.add_command(label="Enviar email", command=lambda: self._socios_enviar_email(row))
            menu.add_command(label="Copiar", command=lambda: self.copiar_celda(event, tree))
            menu.post(event.x_root, event.y_root)

        tooltip = {"win": None, "label": None, "text": ""}

        def show_tooltip(texto, x, y):
            if not texto:
                return
            texto = str(texto)
            if tooltip["win"] is None:
                tip = tk.Toplevel(self)
                tip.wm_overrideredirect(True)
                label = tk.Label(
                    tip,
                    text=texto,
                    justify="left",
                    background="#ffffe0",
                    relief="solid",
                    borderwidth=1,
                    wraplength=360,
                )
                label.pack(ipadx=4, ipady=2)
                tooltip["win"] = tip
                tooltip["label"] = label
                tooltip["text"] = texto
            elif tooltip["text"] != texto and tooltip["label"] is not None:
                tooltip["label"].config(text=texto)
                tooltip["text"] = texto
            tooltip["win"].wm_geometry(f"+{x}+{y}")

        def hide_tooltip():
            if tooltip["win"] is not None:
                tooltip["win"].destroy()
                tooltip["win"] = None
            tooltip["label"] = None
            tooltip["text"] = ""

        def on_hover(event):
            item = tree.identify_row(event.y)
            if not item:
                hide_tooltip()
                return
            col = tree.identify_column(event.x)
            col_index = int(col[1:]) - 1 if col else -1
            if col_index in (cols.index("incidencia"), cols.index("gestion")):
                values = tree.item(item, "values")
                texto = values[col_index] if col_index < len(values) else ""
                if texto:
                    show_tooltip(str(texto), event.x_root + 12, event.y_root + 12)
                else:
                    hide_tooltip()
            else:
                hide_tooltip()

        def on_double_click(event):
            item = tree.identify_row(event.y)
            if not item:
                return
            self._socios_cambiar_estado(item)

        def on_select(_event):
            sel = tree.selection()
            if not sel:
                self._socios_mostrar_boton_historial(None)
                return
            values = tree.item(sel[0], "values")
            codigo_sel = values[0] if values else ""
            self._socios_mostrar_boton_historial(codigo_sel)

        tree.bind("<Button-3>", on_right_click)
        tree.bind("<Motion>", on_hover)
        tree.bind("<Leave>", lambda _e: hide_tooltip())
        tree.bind("<Double-1>", on_double_click)
        tree.bind("<<TreeviewSelect>>", on_select)
        self.tree_incidencias_socios = tree
        tab.tree = tree
        self.refrescar_incidencias_socios_tree()

    def cargar_objetos_taquillas(self):
        try:
            data = self._state_get("objetos_taquillas", [], self.objetos_taquillas_file)
            self.objetos_taquillas = data if isinstance(data, list) else []
            changed = False
            for item in self.objetos_taquillas:
                if "id" not in item:
                    item["id"] = uuid.uuid4().hex
                    changed = True
                if "reporte_path" not in item:
                    item["reporte_path"] = ""
                    changed = True
                if "retiradas" not in item:
                    item["retiradas"] = "NO"
                    changed = True
                if "fecha_retirada" not in item:
                    item["fecha_retirada"] = ""
                    changed = True
                if "fecha_eliminadas" not in item:
                    item["fecha_eliminadas"] = ""
                    changed = True
                if "staff_elimina" not in item:
                    item["staff_elimina"] = ""
                    changed = True
            if changed:
                self.guardar_objetos_taquillas()
        except Exception:
            self.objetos_taquillas = []
        if hasattr(self, "tree_objetos_taquillas"):
            self.refrescar_objetos_taquillas_tree()
        self._update_taquillas_blink()

    def guardar_objetos_taquillas(self):
        try:
            self._state_set("objetos_taquillas", self.objetos_taquillas, self.objetos_taquillas_file)
        except Exception:
            pass

    def _taquillas_parse_dt(self, value):
        if not value:
            return None
        raw = str(value).strip()
        if not raw:
            return None
        raw = raw.replace("\u00a0", " ")
        raw = " ".join(raw.split())
        candidates = [raw]
        if " " not in raw:
            candidates.append(f"{raw} 00:00")
        formats = [
            "%d/%m/%Y %H:%M",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%y %H:%M",
            "%d/%m/%y %H:%M:%S",
            "%d-%m-%Y %H:%M",
            "%d-%m-%Y %H:%M:%S",
            "%d-%m-%y %H:%M",
            "%d-%m-%y %H:%M:%S",
        ]
        for candidate in candidates:
            for fmt in formats:
                try:
                    return datetime.strptime(candidate, fmt)
                except Exception:
                    pass
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
            try:
                return datetime.strptime(raw, fmt)
            except Exception:
                pass
        return None

    def _taquillas_now_str(self):
        return datetime.now().strftime("%d/%m/%Y %H:%M")

    def _taquillas_due_str(self, base_dt):
        return base_dt.strftime("%d/%m/%Y %H:%M")

    def _taquillas_get_overdue_ids(self):
        now = datetime.now()
        overdue = []
        for item in self.objetos_taquillas:
            if item.get("fecha_retirada") or item.get("fecha_eliminadas"):
                continue
            fin = self._taquillas_parse_dt(item.get("fecha_fin"))
            if fin and fin <= now:
                overdue.append(item.get("id"))
        return overdue

    def _start_taquillas_row_blink(self):
        if self.objetos_taquillas_blink_job:
            return

        def toggle():
            self.objetos_taquillas_blink_on = not self.objetos_taquillas_blink_on
            self.refrescar_objetos_taquillas_tree()
            self.objetos_taquillas_blink_job = self.after(self.blink_interval_ms, toggle)

        toggle()

    def _stop_taquillas_row_blink(self):
        if self.objetos_taquillas_blink_job:
            try:
                self.after_cancel(self.objetos_taquillas_blink_job)
            except Exception:
                pass
        self.objetos_taquillas_blink_job = None
        self.objetos_taquillas_blink_on = False

    def _update_taquillas_blink(self):
        overdue_ids = self._taquillas_get_overdue_ids()
        btn = getattr(self, "btn_objetos_taquillas", None)
        if overdue_ids:
            if btn:
                self._start_blink("objetos_taquillas_btn", btn, on_bg="#ffe082", on_fg="black")
            self._start_taquillas_row_blink()
        else:
            self._stop_blink("objetos_taquillas_btn")
            self._stop_taquillas_row_blink()

    def _taquillas_select_vestuario(self, default=None):
        opciones = ["MASCULINO", "FEMENINO"]
        win = tk.Toplevel(self)
        win.title("Vestuario")
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text="Vestuario:").pack(padx=10, pady=6)
        var = tk.StringVar(value=default or opciones[0])
        combo = ttk.Combobox(win, values=opciones, textvariable=var, state="readonly", width=20)
        combo.pack(padx=10, pady=6)
        combo.focus_set()
        res = {"value": None}

        def aceptar():
            res["value"] = var.get()
            win.destroy()

        btn = tk.Button(win, text="Aceptar", command=aceptar)
        btn.pack(pady=6)
        win.bind("<Return>", lambda _e: aceptar())
        win.bind("<Escape>", lambda _e: win.destroy())
        self._incidencias_center_window(win)
        win.grab_set()
        win.wait_window()
        return res["value"]

    def _taquillas_nuevo_registro(self):
        staff = self._staff_select("Taquillas", "Que staff va a hacer el registro?")
        if not staff:
            return
        vestuario = self._taquillas_select_vestuario()
        if not vestuario:
            return
        taquilla = self._incidencias_prompt_text("Taquillas", "Nº taquilla:")
        if taquilla is None:
            return
        bolsa = self._incidencias_prompt_text("Taquillas", "Nº bolsa asignada:")
        if bolsa is None:
            return
        reporte_path = self._incidencias_pedir_reporte_visual()
        now = datetime.now()
        fecha_extraccion = self._taquillas_due_str(now)
        fecha_fin = self._taquillas_due_str(now + timedelta(days=7))
        item = {
            "id": uuid.uuid4().hex,
            "vestuario": vestuario,
            "taquilla": taquilla,
            "bolsa": bolsa,
            "fecha_extraccion": fecha_extraccion,
            "staff_extrae": self._staff_display_name(staff),
            "retiradas": "NO",
            "socio": "",
            "fecha_retirada": "",
            "fecha_fin": fecha_fin,
            "fecha_eliminadas": "",
            "staff_elimina": "",
            "reporte_path": reporte_path or "",
        }
        self.objetos_taquillas.append(item)
        self.guardar_objetos_taquillas()
        self.refrescar_objetos_taquillas_tree()
        self._update_taquillas_blink()

    def _taquillas_modificar_registro(self, item_id):
        item = next((i for i in self.objetos_taquillas if i.get("id") == item_id), None)
        if not item:
            return
        staff = self._staff_select("Taquillas", "Que staff actualiza el registro?")
        if not staff:
            return
        vestuario = self._taquillas_select_vestuario(item.get("vestuario"))
        if not vestuario:
            return
        taquilla = self._incidencias_prompt_text("Taquillas", "Nº taquilla:", item.get("taquilla", ""))
        if taquilla is None:
            return
        bolsa = self._incidencias_prompt_text("Taquillas", "Nº bolsa asignada:", item.get("bolsa", ""))
        if bolsa is None:
            return
        item["staff_extrae"] = self._staff_display_name(staff)
        item["vestuario"] = vestuario
        item["taquilla"] = taquilla
        item["bolsa"] = bolsa
        if messagebox.askyesno("Reporte visual", "Desea cambiar el reporte visual?", parent=self):
            nuevo = self._incidencias_pedir_reporte_visual()
            if nuevo:
                anterior = item.get("reporte_path")
                if anterior and os.path.exists(anterior):
                    try:
                        os.remove(anterior)
                    except Exception:
                        pass
                item["reporte_path"] = nuevo
        self.guardar_objetos_taquillas()
        self.refrescar_objetos_taquillas_tree()
        self._update_taquillas_blink()

    def _taquillas_eliminar_registro(self, item_id):
        item = next((i for i in self.objetos_taquillas if i.get("id") == item_id), None)
        if not item:
            return
        if not messagebox.askyesno("Eliminar registro", "Estas seguro de eliminar este registro?", parent=self):
            return
        reporte = item.get("reporte_path")
        if reporte and os.path.exists(reporte):
            try:
                os.remove(reporte)
            except Exception:
                pass
        self.objetos_taquillas = [i for i in self.objetos_taquillas if i.get("id") != item_id]
        self.guardar_objetos_taquillas()
        self.refrescar_objetos_taquillas_tree()
        self._update_taquillas_blink()

    def _taquillas_ver_reporte(self, item_id):
        item = next((i for i in self.objetos_taquillas if i.get("id") == item_id), None)
        if not item:
            return
        reporte = item.get("reporte_path", "")
        resolved = self._incidencias_resolve_reporte_path(reporte)
        if resolved:
            stored = self._incidencias_store_reporte_path(resolved)
            if stored and stored != reporte:
                item["reporte_path"] = stored
                self.guardar_objetos_taquillas()
                self.refrescar_objetos_taquillas_tree()
        self._incidencias_ver_reporte(resolved or reporte)

    def _taquillas_eliminar_pertenencias(self, item_id):
        item = next((i for i in self.objetos_taquillas if i.get("id") == item_id), None)
        if not item:
            return
        staff = self._staff_select("Taquillas", "Que staff elimina las pertenencias?")
        if not staff:
            return
        item["staff_elimina"] = self._staff_display_name(staff)
        item["fecha_eliminadas"] = self._taquillas_now_str()
        item["retiradas"] = "NO"
        self.guardar_objetos_taquillas()
        self.refrescar_objetos_taquillas_tree()
        self._update_taquillas_blink()

    def _taquillas_entregar_pertenencias(self, item_id):
        item = next((i for i in self.objetos_taquillas if i.get("id") == item_id), None)
        if not item:
            return
        staff = self._staff_select("Taquillas", "Que staff entrega las pertenencias?")
        if not staff:
            return
        socio = self._incidencias_prompt_text("Taquillas", "Numero de socio:")
        if socio is None:
            return
        item["staff_elimina"] = self._staff_display_name(staff)
        item["socio"] = socio
        item["retiradas"] = "SI"
        item["fecha_retirada"] = self._taquillas_now_str()
        self.guardar_objetos_taquillas()
        self.refrescar_objetos_taquillas_tree()
        self._update_taquillas_blink()

    def refrescar_objetos_taquillas_tree(self):
        if not hasattr(self, "tree_objetos_taquillas"):
            return
        tree = self.tree_objetos_taquillas
        tree.delete(*tree.get_children())
        overdue_ids = set(self._taquillas_get_overdue_ids())
        blink_on = self.objetos_taquillas_blink_on
        for item in sorted(
            self.objetos_taquillas,
            key=lambda i: self._taquillas_parse_dt(i.get("fecha_extraccion")) or datetime.min,
            reverse=True,
        ):
            tag = ""
            if item.get("id") in overdue_ids:
                tag = "overdue_on" if blink_on else "overdue_off"
            values = [
                item.get("vestuario", ""),
                item.get("taquilla", ""),
                item.get("bolsa", ""),
                item.get("fecha_extraccion", ""),
                item.get("staff_extrae", ""),
                item.get("retiradas", ""),
                item.get("socio", ""),
                item.get("fecha_retirada", ""),
                item.get("fecha_fin", ""),
                item.get("fecha_eliminadas", ""),
                item.get("staff_elimina", ""),
            ]
            tree.insert("", "end", iid=item.get("id"), values=values, tags=(tag,) if tag else ())

    def create_objetos_taquillas_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        top = tk.Frame(frm)
        top.pack(fill="x", pady=4)
        tk.Button(top, text="NUEVO REGISTRO", command=self._taquillas_nuevo_registro, bg="#ffcc80", fg="black").pack(side="left", padx=5)

        cols = [
            "vestuario",
            "taquilla",
            "bolsa",
            "fecha_extraccion",
            "staff_extrae",
            "retiradas",
            "socio",
            "fecha_retirada",
            "fecha_fin",
            "fecha_eliminadas",
            "staff_elimina",
        ]
        table = tk.Frame(frm)
        table.pack(fill="both", expand=True)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=cols, show="headings")
        headings = {
            "vestuario": "VESTUARIO",
            "taquilla": "Nº TAQUILLA",
            "bolsa": "Nº BOLSA ASIGNADA",
            "fecha_extraccion": "DÍA DE EXTRACCIÓN",
            "staff_extrae": "¿QUÉ STAFF EXTRAE?",
            "retiradas": "¿PERTENENCIAS RETIRADAS POR SOCIO?",
            "socio": "Nº SOCIO",
            "fecha_retirada": "FECHA RETIRADA POR SOCIO",
            "fecha_fin": "FECHA FIN PARA ELIMINAR",
            "fecha_eliminadas": "FECHA PERTENENCIAS ELIMINADAS",
            "staff_elimina": "STAFF QUE ELIMINA O DEVUELVE",
        }
        for c in cols:
            tree.heading(c, text=headings.get(c, c))
            width = 140
            if c in ("fecha_extraccion", "fecha_retirada", "fecha_fin", "fecha_eliminadas"):
                width = 165
            elif c in ("vestuario", "taquilla", "bolsa", "socio"):
                width = 120
            elif c in ("staff_extrae", "staff_elimina"):
                width = 170
            elif c in ("retiradas",):
                width = 210
            tree.column(c, anchor="center", width=width, stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")
        vscroll = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        tree.bind("<Control-c>", lambda e: self.copiar_celda(e, tree))
        tree.tag_configure("overdue_on", background="#ffd6d6")
        tree.tag_configure("overdue_off", background="#ffecec")

        def on_right_click(event):
            row = tree.identify_row(event.y)
            if not row:
                return
            tree.selection_set(row)
            menu = tk.Menu(tree, tearoff=0)
            menu.add_command(label="Eliminar registro", command=lambda: self._taquillas_eliminar_registro(row))
            menu.add_command(label="Modificar registro", command=lambda: self._taquillas_modificar_registro(row))
            menu.add_command(label="Ver reporte grafico", command=lambda: self._taquillas_ver_reporte(row))
            menu.add_separator()
            menu.add_command(label="Eliminar pertenencias", command=lambda: self._taquillas_eliminar_pertenencias(row))
            menu.add_command(label="Entregar pertenencias", command=lambda: self._taquillas_entregar_pertenencias(row))
            menu.add_separator()
            menu.add_command(label="Copiar", command=lambda: self.copiar_celda(event, tree))
            menu.post(event.x_root, event.y_root)

        tree.bind("<Button-3>", on_right_click)
        self.tree_objetos_taquillas = tree
        tab.tree = tree
        self.refrescar_objetos_taquillas_tree()

    # -----------------------------
    # GESTION BAJAS
    # -----------------------------
    def create_bajas_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        top = tk.Frame(frm)
        top.pack(fill="x", pady=4)
        tk.Button(top, text="NUEVO REGISTRO", command=self._bajas_nuevo_registro, bg="#ffcc80", fg="black").pack(
            side="left", padx=5
        )
        vista_btn = tk.Menubutton(top, text="VISTA", bg="#bdbdbd", fg="black")
        vista_menu = tk.Menu(vista_btn, tearoff=0)
        vista_menu.add_command(label="TODOS LOS REGISTROS", command=lambda: self._bajas_set_view("TODOS"))
        vista_menu.add_command(label="PENDIENTES", command=lambda: self._bajas_set_view("PENDIENTE"))
        vista_menu.add_command(label="TRAMITADAS", command=lambda: self._bajas_set_view("TRAMITADA"))
        vista_menu.add_command(label="RECHAZADAS", command=lambda: self._bajas_set_view("RECHAZADA"))
        vista_menu.add_command(label="RECUPERADAS", command=lambda: self._bajas_set_view("RECUPERADA"))
        vista_btn.configure(menu=vista_menu)
        vista_btn.pack(side="left", padx=5)
        tk.Button(top, text="METRICAS", command=self._bajas_metricas, bg="#90caf9", fg="black").pack(side="left", padx=5)
        tk.Label(top, text="Cliente:").pack(side="left", padx=(15, 4))
        self.bajas_buscar_entry = tk.Entry(top, width=14)
        self.bajas_buscar_entry.pack(side="left", padx=4)
        tk.Button(top, text="BUSCAR", command=self._bajas_buscar_cliente, bg="#e0e0e0", fg="black").pack(side="left", padx=4)
        tk.Button(top, text="LIMPIAR", command=self._bajas_limpiar_cliente_filter, bg="#f5f5f5", fg="black").pack(side="left", padx=4)

        cols = [
            "staff",
            "codigo",
            "email",
            "apellidos",
            "nombre",
            "movil",
            "tipo_baja",
            "motivo",
            "estado",
            "reporte",
            "fecha_registro",
            "fecha_tramitacion",
            "fecha_rechazo",
            "fecha_recuperacion",
            "devolucion_recibo",
            "incidencia",
            "solucion",
        ]
        table = tk.Frame(frm)
        table.pack(fill="both", expand=True)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=cols, show="headings")
        headings = {
            "staff": "STAFF",
            "codigo": "CLIENTE",
            "email": "EMAIL",
            "apellidos": "APELLIDOS",
            "nombre": "NOMBRE",
            "movil": "MOVIL",
            "tipo_baja": "TIPO DE BAJA",
            "motivo": "MOTIVO",
            "estado": "ESTADO",
            "reporte": "REPORTE",
            "fecha_registro": "FECHA REGISTRO",
            "fecha_tramitacion": "FECHA TRAMITACION",
            "fecha_rechazo": "FECHA RECHAZO",
            "fecha_recuperacion": "FECHA RECUPERACION",
            "devolucion_recibo": "DEVOLUCION RECIBO",
            "incidencia": "INCIDENCIA",
            "solucion": "SOLUCION",
        }
        for c in cols:
            tree.heading(c, text=headings.get(c, c))
            width = 110
            if c in ("email",):
                width = 160
            elif c in ("motivo", "incidencia", "solucion"):
                width = 180
            elif c in ("fecha_registro", "fecha_tramitacion", "fecha_rechazo", "fecha_recuperacion"):
                width = 140
            elif c in ("tipo_baja", "devolucion_recibo"):
                width = 140
            elif c in ("reporte",):
                width = 80
            tree.column(c, anchor="center", width=width, stretch=True)
        base_widths = {c: tree.column(c, "width") for c in cols}
        tree.grid(row=0, column=0, sticky="nsew")
        vscroll = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        tree.bind("<Control-c>", lambda e: self.copiar_celda(e, tree))
        tree.tag_configure("pendiente", background="#ffe6cc")
        tree.tag_configure("tramitada", background="#d4edda")
        tree.tag_configure("recuperada", background="#d9edf7")
        tree.tag_configure("rechazada", background="#f8d7da")
        tree.after(200, lambda: self._bajas_autofit_columns(tree, cols, base_widths))

        def on_right_click(event):
            row = tree.identify_row(event.y)
            if not row:
                return
            tree.selection_set(row)
            menu = tk.Menu(tree, tearoff=0)
            menu.add_command(label="Modificar estado", command=lambda: self._bajas_cambiar_estado(row))
            menu.add_command(label="Ver motivo", command=lambda: self._bajas_ver_texto(row, "motivo", "Motivo"))
            menu.add_command(label="Modificar motivo", command=lambda: self._bajas_editar_campo(row, "motivo", "Motivo", "Modificar motivo:"))
            menu.add_command(
                label="Modificar tipo de baja",
                command=lambda: self._bajas_editar_campo(row, "tipo_baja", "Tipo de baja", "Modificar tipo de baja:"),
            )
            menu.add_command(label="Agregar incidencia", command=lambda: self._bajas_editar_campo(row, "incidencia", "Incidencia", "Agregar incidencia:"))
            menu.add_command(label="Modificar incidencia", command=lambda: self._bajas_editar_campo(row, "incidencia", "Incidencia", "Modificar incidencia:"))
            menu.add_command(label="Ver incidencia", command=lambda: self._bajas_ver_texto(row, "incidencia", "Incidencia"))
            menu.add_command(label="Agregar solucion", command=lambda: self._bajas_editar_campo(row, "solucion", "Solucion", "Agregar solucion:"))
            menu.add_command(label="Modificar solucion", command=lambda: self._bajas_editar_campo(row, "solucion", "Solucion", "Modificar solucion:"))
            menu.add_command(label="Ver solucion", command=lambda: self._bajas_ver_texto(row, "solucion", "Solucion"))
            menu.add_command(label="Ver solicitudes individuales", command=lambda: self._bajas_ver_solicitudes_individuales(row))
            menu.add_command(label="Agregar reporte grafico", command=lambda: self._bajas_agregar_reporte(row))
            menu.add_command(label="Eliminar reporte grafico", command=lambda: self._bajas_eliminar_reporte(row))
            menu.add_separator()
            menu.add_command(label="Abrir chat", command=lambda: self._bajas_abrir_chat(row))
            menu.add_command(label="Enviar email", command=lambda: self._bajas_enviar_email(row))
            menu.add_separator()
            menu.add_command(label="Eliminar registro", command=lambda: self._bajas_eliminar_registro(row))
            menu.add_command(label="Copiar", command=lambda: self.copiar_celda(event, tree))
            menu.post(event.x_root, event.y_root)

        def on_double_click(event):
            row = tree.identify_row(event.y)
            if row:
                self._bajas_cambiar_estado(row)

        tooltip = {"win": None, "label": None, "text": ""}

        def show_tooltip(texto, x, y):
            if not texto:
                return
            if tooltip["win"] is None:
                tip = tk.Toplevel(self)
                tip.wm_overrideredirect(True)
                label = tk.Label(
                    tip,
                    text=texto,
                    justify="left",
                    background="#ffffe0",
                    relief="solid",
                    borderwidth=1,
                    wraplength=360,
                )
                label.pack(ipadx=4, ipady=2)
                tooltip["win"] = tip
                tooltip["label"] = label
                tooltip["text"] = texto
            elif tooltip["text"] != texto and tooltip["label"] is not None:
                tooltip["label"].config(text=texto)
                tooltip["text"] = texto
            tooltip["win"].wm_geometry(f"+{x}+{y}")

        def hide_tooltip():
            if tooltip["win"] is not None:
                tooltip["win"].destroy()
                tooltip["win"] = None
            tooltip["label"] = None
            tooltip["text"] = ""

        def on_hover(event):
            item = tree.identify_row(event.y)
            if not item:
                hide_tooltip()
                return
            col = tree.identify_column(event.x)
            col_index = int(col[1:]) - 1 if col else -1
            values = tree.item(item, "values")
            texto = values[col_index] if col_index < len(values) else ""
            if texto:
                show_tooltip(str(texto), event.x_root + 12, event.y_root + 12)
            else:
                hide_tooltip()

        tree.bind("<Button-3>", on_right_click)
        tree.bind("<Double-1>", on_double_click)
        tree.bind("<Motion>", on_hover)
        tree.bind("<Leave>", lambda _e: hide_tooltip())
        self.tree_bajas = tree
        tab.tree = tree
        self.refrescar_bajas_tree()

    def _bajas_autofit_columns(self, tree, cols, base_widths):
        if not tree.winfo_exists():
            return
        available = tree.winfo_width()
        if available <= 1:
            tree.after(100, lambda: self._bajas_autofit_columns(tree, cols, base_widths))
            return
        total = sum(base_widths.get(c, 100) for c in cols)
        if total <= 0:
            return
        scale = min(1.0, available / total)
        min_width = 50
        for c in cols:
            base = base_widths.get(c, 100)
            width = max(min_width, int(base * scale))
            tree.column(c, width=width, stretch=False)

    def _bajas_ver_texto(self, baja_id, field, title):
        item = next((b for b in self.bajas if b.get("id") == baja_id), None)
        if not item:
            return
        texto = (item.get(field) or "").strip()
        if not texto:
            texto = f"Sin {field}."
        messagebox.showinfo(title, texto)

    def cargar_bajas(self):
        try:
            data = self._state_get("bajas", [], self.bajas_file)
            self.bajas = data if isinstance(data, list) else []
            changed = False
            for item in self.bajas:
                if "id" not in item:
                    item["id"] = uuid.uuid4().hex
                    changed = True
                if "estado" not in item:
                    item["estado"] = "PENDIENTE"
                    changed = True
                if "fecha_registro" not in item:
                    item["fecha_registro"] = ""
                    changed = True
                if "fecha_tramitacion" not in item:
                    item["fecha_tramitacion"] = ""
                    changed = True
                if "fecha_rechazo" not in item:
                    item["fecha_rechazo"] = ""
                    changed = True
                if "fecha_recuperacion" not in item:
                    item["fecha_recuperacion"] = ""
                    changed = True
                if "devolucion_recibo" not in item:
                    item["devolucion_recibo"] = ""
                    changed = True
                if "reporte_path" not in item:
                    item["reporte_path"] = ""
                    changed = True
                if "incidencia" not in item:
                    item["incidencia"] = ""
                    changed = True
                if "solucion" not in item:
                    item["solucion"] = ""
                    changed = True
            if changed:
                self.guardar_bajas()
        except Exception:
            self.bajas = []
        self._bajas_actualizar_devolucion()
        if hasattr(self, "tree_bajas"):
            self.refrescar_bajas_tree()

    def guardar_bajas(self):
        try:
            self._state_set("bajas", self.bajas, self.bajas_file)
        except Exception:
            pass

    def _bajas_now_str(self):
        return datetime.now().strftime("%d/%m/%Y %H:%M")

    def _bajas_parse_dt(self, value):
        try:
            return datetime.strptime(str(value or ""), "%d/%m/%Y %H:%M")
        except Exception:
            return None

    def _bajas_get_impagos_set(self):
        fecha = self.impagos_last_export or self.impagos_db.get_last_export()
        if not fecha:
            return set()
        try:
            with self.impagos_db._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    "SELECT DISTINCT c.numero_cliente "
                    "FROM impagos_clientes c "
                    "JOIN impagos_eventos e ON e.cliente_id = c.id "
                    "WHERE e.fecha_export = ?",
                    (fecha,),
                )
                return {str(r[0]).strip() for r in cur.fetchall() if r and r[0]}
        except Exception:
            return set()

    def _bajas_actualizar_devolucion(self):
        impagos_set = self._bajas_get_impagos_set()
        self.bajas_impagos_set = set(impagos_set)
        changed = False
        for item in self.bajas:
            codigo = str(item.get("codigo", "")).strip()
            valor = "SI" if codigo and codigo in impagos_set else "NO"
            if item.get("devolucion_recibo") != valor:
                item["devolucion_recibo"] = valor
                changed = True
        if changed:
            self.guardar_bajas()

    def _bajas_buscar_cliente_info(self, codigo):
        codigo = str(codigo or "").strip()
        if not codigo:
            return {}
        if self.resumen_df is not None:
            cols = {self._norm(c): c for c in self.resumen_df.columns}
            col_codigo = cols.get("NUMERO DE CLIENTE") or cols.get("NUMERO DE SOCIO")
            col_nombre = cols.get("NOMBRE")
            col_apellidos = cols.get("APELLIDOS")
            col_email = cols.get("CORREO ELECTRONICO") or cols.get("EMAIL") or cols.get("CORREO")
            col_movil = cols.get("MOVIL") or cols.get("TELEFONO") or cols.get("TELEFONO MOVIL")
            if col_codigo:
                df = self.resumen_df[self.resumen_df[col_codigo].astype(str).str.strip() == codigo]
                if not df.empty:
                    row = df.iloc[0]
                    return {
                        "nombre": str(row.get(col_nombre, "")).strip(),
                        "apellidos": str(row.get(col_apellidos, "")).strip(),
                        "email": str(row.get(col_email, "")).strip(),
                        "movil": str(row.get(col_movil, "")).strip(),
                    }
        for c in self.clientes_ext:
            if str(c.get("codigo", "")).strip() == codigo:
                return {
                    "nombre": str(c.get("nombre", "")).strip(),
                    "apellidos": str(c.get("apellidos", "")).strip(),
                    "email": str(c.get("email", "")).strip(),
                    "movil": str(c.get("movil", "")).strip(),
                }
        return {}

    def _bajas_select_option(self, title, prompt, opciones, width=28):
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text=prompt).pack(padx=10, pady=6)
        var = tk.StringVar(value=opciones[0] if opciones else "")
        combo = ttk.Combobox(win, values=opciones, textvariable=var, state="readonly", width=width)
        combo.pack(padx=10, pady=6)
        combo.focus_set()
        res = {"value": None}

        def aceptar():
            res["value"] = var.get()
            win.destroy()

        def cancel():
            res["value"] = None
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=6)
        tk.Button(btns, text="Aceptar", command=aceptar).pack(side="left", padx=6)
        tk.Button(btns, text="Cancelar", command=cancel).pack(side="left", padx=6)
        win.bind("<Return>", lambda _e: aceptar())
        win.bind("<Escape>", lambda _e: cancel())
        win.grab_set()
        self._incidencias_center_window(win)
        win.wait_window()
        return res["value"]

    def _bajas_nuevo_registro(self):
        staff = self._staff_select("Staff", "Que staff registra?")
        if not staff:
            return
        codigo = self._incidencias_prompt_text("Cliente", "Numero de cliente:")
        if not codigo:
            return
        codigo = str(codigo).strip()
        ya_existe = any(b.get("codigo", "").strip() == codigo for b in self.bajas)
        if ya_existe:
            messagebox.showwarning("Bajas", f"El cliente {codigo} ya tiene registros previos.")
        tipo = self._bajas_select_option(
            "Tipo de baja",
            "Tipo de baja:",
            ["CON PERMANENCIA", "SIN PERMANENCIA", "DESISTIMIENTO", "DESISTIMIENTO + DEVOLUCION"],
        )
        if not tipo:
            return
        motivo = self._bajas_select_option(
            "Motivo",
            "Motivo:",
            ["TIEMPO", "MOTIVACION", "ECONOMICO", "TRASLADO", "INSATISFACCION", "MEDICO", "SIN MOTIVO", "OTRO"],
        )
        if not motivo:
            return
        if motivo == "OTRO":
            otro = self._incidencias_prompt_text("Motivo", "Describe el motivo:")
            if not otro:
                return
            motivo = f"OTRO: {otro.strip()}"
        incidencia = ""
        if messagebox.askyesno("Incidencia", "Desea agregar una incidencia?", parent=self):
            inc = self._incidencias_prompt_text("Incidencia", "Describe la incidencia:")
            if inc:
                incidencia = inc.strip()
        reporte_path = ""
        if messagebox.askyesno("Reporte grafico", "Desea agregar reporte grafico?", parent=self):
            nuevo = self._incidencias_pedir_reporte_visual()
            if nuevo:
                reporte_path = nuevo

        info = self._bajas_buscar_cliente_info(codigo)
        now_str = self._bajas_now_str()
        registro = {
            "id": uuid.uuid4().hex,
            "staff": self._staff_display_name(staff),
            "codigo": codigo,
            "email": info.get("email", ""),
            "apellidos": info.get("apellidos", ""),
            "nombre": info.get("nombre", ""),
            "movil": info.get("movil", ""),
            "tipo_baja": tipo,
            "motivo": motivo,
            "estado": "PENDIENTE",
            "fecha_registro": now_str,
            "fecha_tramitacion": "",
            "fecha_rechazo": "",
            "fecha_recuperacion": "",
            "devolucion_recibo": "SI" if codigo in self.bajas_impagos_set else "NO",
            "incidencia": incidencia,
            "solucion": "",
            "reporte_path": reporte_path,
        }
        registro["movil"] = self._normalizar_movil(registro.get("movil", ""))
        self.bajas.append(registro)
        self.guardar_bajas()
        self.refrescar_bajas_tree()

    def _bajas_set_view(self, view):
        self.bajas_view = view
        self.refrescar_bajas_tree()

    def _bajas_set_cliente_filter(self, codigo):
        self.bajas_cliente_filter = (codigo or "").strip()
        if hasattr(self, "bajas_buscar_entry"):
            self.bajas_buscar_entry.delete(0, "end")
            self.bajas_buscar_entry.insert(0, self.bajas_cliente_filter)
        self.refrescar_bajas_tree()

    def _bajas_limpiar_cliente_filter(self):
        self.bajas_cliente_filter = ""
        if hasattr(self, "bajas_buscar_entry"):
            self.bajas_buscar_entry.delete(0, "end")
        self.refrescar_bajas_tree()

    def _bajas_buscar_cliente(self):
        if not hasattr(self, "bajas_buscar_entry"):
            return
        codigo = self.bajas_buscar_entry.get().strip()
        if not codigo:
            messagebox.showinfo("Bajas", "Introduce un numero de cliente.")
            return
        self._bajas_set_cliente_filter(codigo)

    def _bajas_ver_solicitudes_individuales(self, baja_id):
        item = next((b for b in self.bajas if b.get("id") == baja_id), None)
        if not item:
            return
        codigo = item.get("codigo", "")
        if not codigo:
            return
        self._bajas_set_cliente_filter(codigo)

    def _bajas_editar_campo(self, baja_id, field, title, prompt):
        item = next((i for i in self.bajas if i.get("id") == baja_id), None)
        if not item:
            return
        nuevo = self._incidencias_prompt_text(title, prompt, item.get(field, ""))
        if nuevo is None:
            return
        item[field] = nuevo.strip()
        self.guardar_bajas()
        self.refrescar_bajas_tree()

    def _bajas_eliminar_registro(self, baja_id):
        if not messagebox.askyesno("Eliminar", "Eliminar este registro?", parent=self):
            return
        self.bajas = [b for b in self.bajas if b.get("id") != baja_id]
        self.guardar_bajas()
        self.refrescar_bajas_tree()

    def _bajas_selector_estado(self):
        opciones = ["PENDIENTE", "TRAMITADA", "RECUPERADA", "RECHAZADA"]
        return self._bajas_select_option("Estado", "Selecciona estado:", opciones, width=18)

    def _bajas_cambiar_estado(self, baja_id):
        item = next((i for i in self.bajas if i.get("id") == baja_id), None)
        if not item:
            return
        estado = self._bajas_selector_estado()
        if not estado:
            return
        item["estado"] = estado
        now_str = self._bajas_now_str()
        if estado == "TRAMITADA":
            item["fecha_tramitacion"] = now_str
        elif estado == "RECUPERADA":
            item["fecha_recuperacion"] = now_str
        elif estado == "RECHAZADA":
            item["fecha_rechazo"] = now_str
        self.guardar_bajas()
        self.refrescar_bajas_tree()

    def _bajas_abrir_chat(self, baja_id):
        item = next((i for i in self.bajas if i.get("id") == baja_id), None)
        if not item:
            return
        movil = self._normalizar_movil(item.get("movil", ""))
        if not movil:
            messagebox.showwarning("Chat", "El cliente no tiene movil registrado.")
            return
        webbrowser.open(f"https://wa.me/{movil}")

    def _bajas_agregar_reporte(self, baja_id):
        item = next((i for i in self.bajas if i.get("id") == baja_id), None)
        if not item:
            return
        nuevo = self._incidencias_pedir_reporte_visual()
        if not nuevo:
            return
        item["reporte_path"] = nuevo
        self.guardar_bajas()
        self.refrescar_bajas_tree()

    def _bajas_eliminar_reporte(self, baja_id):
        item = next((i for i in self.bajas if i.get("id") == baja_id), None)
        if not item:
            return
        if not item.get("reporte_path"):
            messagebox.showwarning("Reporte", "No hay reporte grafico.")
            return
        if not messagebox.askyesno("Reporte", "Eliminar el reporte grafico?", parent=self):
            return
        item["reporte_path"] = ""
        self.guardar_bajas()
        self.refrescar_bajas_tree()

    def _bajas_enviar_email(self, baja_id):
        item = next((i for i in self.bajas if i.get("id") == baja_id), None)
        if not item:
            return
        email = str(item.get("email", "")).strip()
        if not email:
            messagebox.showwarning("Email", "El cliente no tiene email registrado.")
            return
        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(email)
            messagebox.showwarning("Outlook no disponible", "Email copiado al portapapeles.")
            return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = ""
            mail.Body = ""
            mail.Display()
        except Exception as e:
            messagebox.showerror("Email", f"No se pudo abrir Outlook.\nDetalle: {e}")

    def refrescar_bajas_tree(self):
        if not hasattr(self, "tree_bajas"):
            return
        tree = self.tree_bajas
        tree.delete(*tree.get_children())
        view = (self.bajas_view or "TODOS").upper()
        cliente_filter = (self.bajas_cliente_filter or "").strip().upper()
        for item in sorted(self.bajas, key=lambda i: self._bajas_parse_dt(i.get("fecha_registro")) or datetime.min, reverse=True):
            codigo = str(item.get("codigo", "")).strip()
            if cliente_filter and codigo.upper() != cliente_filter:
                continue
            estado = str(item.get("estado", "PENDIENTE")).upper()
            if view != "TODOS" and estado != view:
                continue
            tag = ""
            if estado == "PENDIENTE":
                tag = "pendiente"
            elif estado == "TRAMITADA":
                tag = "tramitada"
            elif estado == "RECUPERADA":
                tag = "recuperada"
            elif estado == "RECHAZADA":
                tag = "rechazada"
            values = [
                item.get("staff", ""),
                item.get("codigo", ""),
                item.get("email", ""),
                item.get("apellidos", ""),
                item.get("nombre", ""),
                item.get("movil", ""),
                item.get("tipo_baja", ""),
                item.get("motivo", ""),
                item.get("estado", ""),
                "R" if item.get("reporte_path") else "",
                item.get("fecha_registro", ""),
                item.get("fecha_tramitacion", ""),
                item.get("fecha_rechazo", ""),
                item.get("fecha_recuperacion", ""),
                item.get("devolucion_recibo", ""),
                item.get("incidencia", ""),
                item.get("solucion", ""),
            ]
            tree.insert("", "end", iid=item.get("id"), values=values, tags=(tag,) if tag else ())

    def _bajas_metricas(self):
        total = len(self.bajas)
        if total == 0:
            messagebox.showinfo("Metricas", "No hay registros.")
            return
        counts_estado = {"PENDIENTE": 0, "TRAMITADA": 0, "RECUPERADA": 0, "RECHAZADA": 0}
        counts_tipo = {}
        counts_motivo = {}
        devolucion_si = 0
        devolucion_no = 0
        for item in self.bajas:
            estado = str(item.get("estado", "PENDIENTE")).upper()
            counts_estado[estado] = counts_estado.get(estado, 0) + 1
            tipo = str(item.get("tipo_baja", "")).strip().upper()
            if tipo:
                counts_tipo[tipo] = counts_tipo.get(tipo, 0) + 1
            motivo = str(item.get("motivo", "")).strip().upper()
            if motivo:
                counts_motivo[motivo] = counts_motivo.get(motivo, 0) + 1
            if str(item.get("devolucion_recibo", "")).strip().upper() == "SI":
                devolucion_si += 1
            else:
                devolucion_no += 1

        win = tk.Toplevel(self)
        win.title("Metricas bajas")
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text="METRICAS - GESTION BAJAS", font=("Arial", 10, "bold")).pack(padx=10, pady=(10, 6))

        resumen = tk.Frame(win)
        resumen.pack(padx=10, pady=6, fill="x")
        tk.Label(resumen, text=f"TOTAL: {total}").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        tk.Label(resumen, text=f"PENDIENTE: {counts_estado['PENDIENTE']}").grid(row=0, column=1, sticky="w", padx=6, pady=2)
        tk.Label(resumen, text=f"TRAMITADA: {counts_estado['TRAMITADA']}").grid(row=0, column=2, sticky="w", padx=6, pady=2)
        tk.Label(resumen, text=f"RECUPERADA: {counts_estado['RECUPERADA']}").grid(row=1, column=1, sticky="w", padx=6, pady=2)
        tk.Label(resumen, text=f"RECHAZADA: {counts_estado['RECHAZADA']}").grid(row=1, column=2, sticky="w", padx=6, pady=2)
        tk.Label(resumen, text=f"DEVOLUCION RECIBO SI: {devolucion_si}").grid(row=2, column=1, sticky="w", padx=6, pady=2)
        tk.Label(resumen, text=f"DEVOLUCION RECIBO NO: {devolucion_no}").grid(row=2, column=2, sticky="w", padx=6, pady=2)

        tk.Label(win, text="MOTIVOS", font=("Arial", 9, "bold")).pack(pady=(8, 4))
        motivos_frame = tk.Frame(win)
        motivos_frame.pack(padx=10, pady=(0, 8), fill="x")
        row = 0
        for motivo, cnt in sorted(counts_motivo.items(), key=lambda x: (-x[1], x[0])):
            pct = int(round((cnt / total) * 100)) if total else 0
            tk.Label(motivos_frame, text=f"{motivo}: {cnt} ({pct}%)").grid(row=row, column=0, sticky="w", padx=6, pady=1)
            row += 1

        tk.Label(win, text="TIPOS DE BAJA", font=("Arial", 9, "bold")).pack(pady=(6, 4))
        tipos_frame = tk.Frame(win)
        tipos_frame.pack(padx=10, pady=(0, 10), fill="x")
        row = 0
        for tipo, cnt in sorted(counts_tipo.items(), key=lambda x: (-x[1], x[0])):
            pct = int(round((cnt / total) * 100)) if total else 0
            tk.Label(tipos_frame, text=f"{tipo}: {cnt} ({pct}%)").grid(row=row, column=0, sticky="w", padx=6, pady=1)
            row += 1

        tk.Button(win, text="Cerrar", command=win.destroy).pack(pady=(0, 10))
        self._incidencias_center_window(win)
        win.grab_set()

    def cargar_clientes_ext(self):
        try:
            data = self._state_get("clientes_ext", [], self.clientes_ext_file)
            self.clientes_ext = data if isinstance(data, list) else []
        except Exception:
            self.clientes_ext = []

    def cargar_felicitaciones(self):
        try:
            data = self._state_get("felicitaciones_enviadas", {}, self.felicitaciones_file)
            self.felicitaciones_enviadas = data if isinstance(data, dict) else {}
        except Exception:
            self.felicitaciones_enviadas = {}

    def guardar_prestamos_json(self):
        try:
            self._state_set("prestamos", self.prestamos, self.prestamos_file)
        except Exception:
            pass  # silencioso para no romper UI

    def guardar_clientes_ext(self):
        try:
            self._state_set("clientes_ext", self.clientes_ext, self.clientes_ext_file)
        except Exception:
            pass

    def guardar_felicitaciones(self):
        try:
            self._state_set("felicitaciones_enviadas", self.felicitaciones_enviadas, self.felicitaciones_file)
        except Exception:
            pass

    def cargar_avanza_fit_envios(self):
        try:
            data = self._state_get("avanza_fit_envios", {}, self.avanza_fit_envios_file)
            self.avanza_fit_envios = data if isinstance(data, dict) else {}
        except Exception:
            self.avanza_fit_envios = {}

    def guardar_avanza_fit_envios(self):
        try:
            self._state_set("avanza_fit_envios", self.avanza_fit_envios, self.avanza_fit_envios_file)
        except Exception:
            pass

    def _start_blink(self, key, button, on_bg="#ffeb3b", on_fg="black"):
        if button is None:
            return
        if key in self.blink_states:
            return
        off_bg = button.cget("bg")
        off_fg = button.cget("fg")
        state = {
            "button": button,
            "on": False,
            "on_bg": on_bg,
            "off_bg": off_bg,
            "on_fg": on_fg,
            "off_fg": off_fg,
            "job": None,
        }
        self.blink_states[key] = state

        def toggle():
            current = self.blink_states.get(key)
            if not current:
                return
            current["on"] = not current["on"]
            try:
                current["button"].config(
                    bg=current["on_bg"] if current["on"] else current["off_bg"],
                    fg=current["on_fg"] if current["on"] else current["off_fg"],
                )
            except tk.TclError:
                self.blink_states.pop(key, None)
                return
            current["job"] = self.after(self.blink_interval_ms, toggle)

        toggle()

    def _stop_blink(self, key):
        current = self.blink_states.pop(key, None)
        if not current:
            return
        if current.get("job"):
            try:
                self.after_cancel(current["job"])
            except Exception:
                pass
        try:
            current["button"].config(bg=current["off_bg"], fg=current["off_fg"])
        except tk.TclError:
            pass

    def update_blink_states(self):
        self._update_felicitacion_blink()
        self._update_avanza_fit_blink()
        self._update_impagos_blinks()
        self._update_taquillas_blink()

    def _hay_cumpleanos_pendientes(self):
        try:
            df = obtener_cumpleanos_hoy()
        except Exception:
            return False
        if df is None or df.empty:
            return False

        def normalize(text):
            t = unicodedata.normalize("NFD", str(text or "")).upper().strip()
            return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")

        colmap = {normalize(c): c for c in df.columns}
        col_email = colmap.get("CORREO ELECTRONICO") or colmap.get("EMAIL") or colmap.get("CORREO")
        if not col_email:
            return False

        current_year = datetime.now().year
        for _, row in df.iterrows():
            email = str(row.get(col_email, "")).strip()
            if not email:
                continue
            sent_year = self.felicitaciones_enviadas.get(email)
            if sent_year != current_year:
                return True
        return False

    def _update_felicitacion_blink(self):
        btn = getattr(self, "btn_enviar_felicitacion", None)
        if not btn:
            return
        if self._hay_cumpleanos_pendientes():
            self._start_blink("felicitacion", btn, on_bg="#ffe082", on_fg="black")
        else:
            self._stop_blink("felicitacion")

    def _update_avanza_fit_blink(self):
        btn = getattr(self, "btn_avanza_fit", None)
        if not btn:
            return
        today = datetime.now().date().isoformat()
        is_tuesday = datetime.now().weekday() == 1
        sent_date = str(self.avanza_fit_envios.get("last_sent_date", ""))
        if is_tuesday and sent_date != today:
            self._start_blink("avanza_fit", btn, on_bg="#ffe082", on_fg="black")
        else:
            self._stop_blink("avanza_fit")

    def _get_impagos_pending_count(self, view):
        if not self.impagos_last_export:
            self.impagos_last_export = self.impagos_db.get_last_export()
        if not self.impagos_last_export:
            return 0
        try:
            rows = self.impagos_db.fetch_view(view, self.impagos_last_export)
        except Exception:
            return 0
        return len(rows) if rows else 0

    def _update_impagos_blinks(self):
        btn1 = getattr(self, "btn_impagos_email_1", None)
        btn2 = getattr(self, "btn_impagos_email_2", None)
        btnr = getattr(self, "btn_impagos_email_resueltos", None)
        if btn1:
            if self._get_impagos_pending_count("incidentes1") > 0:
                self._start_blink("impagos_1inc", btn1, on_bg="#ffe082", on_fg="black")
            else:
                self._stop_blink("impagos_1inc")
        if btn2:
            if self._get_impagos_pending_count("incidentes2") > 0:
                self._start_blink("impagos_2inc", btn2, on_bg="#ffe082", on_fg="black")
            else:
                self._stop_blink("impagos_2inc")
        if btnr:
            if self._get_impagos_pending_count("resueltos") > 0:
                self._start_blink("impagos_resueltos", btnr, on_bg="#c8e6c9", on_fg="black")
            else:
                self._stop_blink("impagos_resueltos")

    def refrescar_prestamos_tree(self):
        if not hasattr(self, "tree_prestamos"):
            return
        self.tree_prestamos.delete(*self.tree_prestamos.get_children())
        self.tree_prestamos["displaycolumns"] = ["codigo", "nombre", "apellidos", "email", "movil", "material", "fecha", "devuelto", "prestado_por"]
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
                p.get("fecha", ""), "SI" if p.get("devuelto") else "NO", p.get("prestado_por", "")
            ], tags=(tag,), iid=iid)

    def toggle_prestamos_vista(self):
        self.prestamos_filtro_activo = not self.prestamos_filtro_activo
        self.refrescar_prestamos_tree()

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
        self.btn_impagos_email_1 = tk.Button(
            actions,
            text="Enviar email (1Inc)",
            command=lambda: self.enviar_email_impagos("1inc"),
            bg="#ffd54f",
            fg="black",
        )
        self.btn_impagos_email_1.pack(side="left", padx=5)
        self.btn_impagos_email_2 = tk.Button(
            actions,
            text="Enviar email (2+ inc)",
            command=lambda: self.enviar_email_impagos("2inc"),
            bg="#ff8a65",
            fg="black",
        )
        self.btn_impagos_email_2.pack(side="left", padx=5)
        self.btn_impagos_email_resueltos = tk.Button(
            actions,
            text="Enviar email resueltos",
            command=self.enviar_email_resueltos,
            bg="#81c784",
            fg="black",
        )
        self.btn_impagos_email_resueltos.pack(side="left", padx=5)

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
                self.sync_impagos(incidencias, show_messages=True)
            except Exception as e:
                messagebox.showerror("Impagos", f"No se pudo cargar IMPAGOS.csv: {e}")
        else:
            messagebox.showwarning("Sin carpeta", "Selecciona primero la carpeta de datos.")

    def sync_impagos(self, df, show_messages=True):
        if not self.impagos_db:
            if show_messages:
                messagebox.showwarning("Impagos", "Base de datos no inicializada.")
            return
        lock_acquired = self._acquire_db_lock()
        if not lock_acquired:
            if show_messages:
                messagebox.showwarning(
                    "Impagos",
                    "Otro equipo está actualizando impagos. Reintenta en unos segundos.",
                )
            self.refresh_impagos_view()
            return
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
            if hasattr(self, "bajas"):
                self._bajas_actualizar_devolucion()
            if hasattr(self, "suspensiones"):
                self._suspensiones_actualizar_devolucion()
        except Exception as e:
            if show_messages:
                messagebox.showerror("Impagos", f"Error sincronizando impagos: {e}")
        finally:
            self._release_db_lock()

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
        self._update_impagos_blinks()

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
            texto = "".join(
                [
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{WAVING HAND SIGN} \u00a1Hola! Tienes un recibo pendiente de pago y el torno no permitir\u00e1 el acceso \N{CRYING FACE}",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{SMALL ORANGE DIAMOND} Si vienes de 6 a 9 am (horario sin atenci\u00f3n comercial), te recomendamos pasar por recepci\u00f3n a partir de las 9 am para solucionarlo lo antes posible \N{SLIGHTLY SMILING FACE}",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{SMALL ORANGE DIAMOND} Puedes abonar en efectivo \N{BANKNOTE WITH EURO SIGN} o con tarjeta \N{CREDIT CARD}, incluso abonarlo desde la propia aplicaci\u00f3n. Aqu\u00ed te mostramos c\u00f3mo m\u00e1s abajo.",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{SPIRAL CALENDAR PAD} Recuerda: puedes anticipar el pago entre 5 y 15 d\u00edas antes de tu siguiente d\u00eda de pago que siempre podr\u00e1s consultar en tu APP Fitness Park.",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    f"\N{MOBILE PHONE} \u00bfDudas? PREG\u00daNTANOS al WhatsApp ({wa_link}) o contestando este email.",
                    "</p>",
                ]
            )
        else:
            texto = "".join(
                [
                    "<p style=\"margin:0 0 10px 0;\">",
                    "Hemos detectado que actualmente tienes 2 recibos o m\u00e1s pendientes en tu cuenta de socio/a.",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "Queremos ayudarte a regularizar tu situaci\u00f3n lo antes posible para que puedas seguir disfrutando de todas las instalaciones sin inconvenientes, evitando que el sistema derive tu caso a la empresa de recobros PayPymes.",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{ELECTRIC LIGHT BULB} Opciones para abonar:",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{CREDIT CARD} Tarjeta o \N{BANKNOTE WITH EURO SIGN} efectivo en recepci\u00f3n.",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{MOBILE PHONE} Directamente desde tu APP Fitness Park, en \"Espacio de cliente\" -> \"Pagos\". (Te mostramos c\u00f3mo m\u00e1s abajo)",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{WARNING SIGN} Aviso importante: En caso de que se genere un tercer recibo impagado, desde el club dejaremos de emitir advertencias y tu expediente el sistema lo derivar\u00e1 directamente a nuestra empresa colaboradora de recobros PayPymes, que se pondr\u00e1 en contacto contigo para gestionar la deuda.",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{SPIRAL CALENDAR PAD} Recuerda: puedes anticipar el pago entre 5 y 15 d\u00edas antes de tu siguiente d\u00eda de pago que siempre podr\u00e1s consultar en tu APP Fitness Park, en el icono espacio de cliente y luego en el apartado \"PAGOS\".",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    "\N{PERSON WITH FOLDED HANDS} Te animamos a ponerte al d\u00eda lo antes posible para evitar cualquier gesti\u00f3n externa y que todo siga con normalidad. Estamos a tu disposici\u00f3n en recepci\u00f3n para cualquier duda.",
                    "</p>",
                    "<p style=\"margin:0 0 10px 0;\">",
                    f"\N{MOBILE PHONE} \u00bfDudas? PREG\u00daNTANOS por WhatsApp ({wa_link}) o contestando este email.",
                    "</p>",
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
        if not self._security_pin_ok():
            return
        if not self._require_manager_access("Impagos"):
            return
        tab = self.tabs.get("Impagos")
        if tab:
            self.notebook.add(tab, text="Impagos", image=self.tab_icons.get("Impagos"), compound="left")
            self.notebook.select(tab)

    def on_tab_changed(self, event):
        current = self.notebook.select()
        if self._suppress_tab_guard:
            self._suppress_tab_guard = False
        else:
            tab_name = self.notebook.tab(current, "text")
            if self.user_role == "STAFF":
                blocked_tabs = {"Incidencias Socios", "Gestion Bajas", "Gestion Suspensiones"}
                if tab_name in blocked_tabs:
                    self._require_manager_access(tab_name)
                    fallback = self._last_allowed_tab
                    if fallback:
                        try:
                            fallback_name = self.notebook.tab(fallback, "text")
                        except Exception:
                            fallback_name = ""
                        if fallback_name in blocked_tabs:
                            fallback = None
                    if not fallback:
                        fallback = self.tabs.get("Prestamos") or self.tabs.get("Wizville")
                    if fallback:
                        self._suppress_tab_guard = True
                        self.notebook.select(fallback)
                    return
        self._last_allowed_tab = current

        tab = self.tabs.get("Impagos")
        if not tab:
            return
        if current != str(tab):
            try:
                self.notebook.hide(tab)
            except Exception:
                pass

    # -----------------------------
    # GESTION SUSPENSIONES
    # -----------------------------
    def create_suspensiones_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        top = tk.Frame(frm)
        top.pack(fill="x", pady=4)
        tk.Button(top, text="NUEVO REGISTRO", command=self._suspensiones_nuevo_registro, bg="#ffcc80", fg="black").pack(
            side="left", padx=5
        )
        vista_btn = tk.Menubutton(top, text="VISTA", bg="#bdbdbd", fg="black")
        vista_menu = tk.Menu(vista_btn, tearoff=0)
        vista_menu.add_command(label="TODOS LOS REGISTROS", command=lambda: self._suspensiones_set_view("TODOS"))
        vista_menu.add_command(label="PENDIENTES", command=lambda: self._suspensiones_set_view("PENDIENTE"))
        vista_menu.add_command(label="TRAMITADAS", command=lambda: self._suspensiones_set_view("TRAMITADA"))
        vista_menu.add_command(label="RECHAZADAS", command=lambda: self._suspensiones_set_view("RECHAZADA"))
        vista_menu.add_command(label="CONCLUIDAS", command=lambda: self._suspensiones_set_view("CONCLUIDA"))
        vista_menu.add_command(label="PENDIENTES Y TRAMITADAS", command=lambda: self._suspensiones_set_view("ACTIVAS"))
        vista_btn.configure(menu=vista_menu)
        vista_btn.pack(side="left", padx=5)
        tk.Label(top, text="Cliente:").pack(side="left", padx=(15, 4))
        self.suspensiones_buscar_entry = tk.Entry(top, width=14)
        self.suspensiones_buscar_entry.pack(side="left", padx=4)
        tk.Button(top, text="BUSCAR", command=self._suspensiones_buscar_cliente, bg="#e0e0e0", fg="black").pack(side="left", padx=4)
        tk.Button(top, text="LIMPIAR", command=self._suspensiones_limpiar_cliente_filter, bg="#f5f5f5", fg="black").pack(side="left", padx=4)

        cols = [
            "staff",
            "codigo",
            "email",
            "apellidos",
            "nombre",
            "movil",
            "motivo",
            "estado",
            "reporte",
            "fecha_registro",
            "fecha_tramitacion",
            "fecha_rechazo",
            "fecha_inicio_suspension",
            "fecha_fin_suspension",
            "devolucion_recibo",
            "incidencia",
            "solucion",
        ]
        table = tk.Frame(frm)
        table.pack(fill="both", expand=True)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        tree = ttk.Treeview(table, columns=cols, show="headings")
        headings = {
            "staff": "STAFF",
            "codigo": "CLIENTE",
            "email": "EMAIL",
            "apellidos": "APELLIDOS",
            "nombre": "NOMBRE",
            "movil": "MOVIL",
            "motivo": "MOTIVO",
            "estado": "ESTADO",
            "reporte": "REPORTE",
            "fecha_registro": "FECHA REGISTRO",
            "fecha_tramitacion": "FECHA TRAMITACION",
            "fecha_rechazo": "FECHA RECHAZO",
            "fecha_inicio_suspension": "FECHA INICIO SUSPENSION",
            "fecha_fin_suspension": "FECHA FIN SUSPENSION",
            "devolucion_recibo": "DEVOLUCION RECIBO",
            "incidencia": "INCIDENCIA",
            "solucion": "SOLUCION",
        }
        for c in cols:
            tree.heading(c, text=headings.get(c, c))
            width = 110
            if c in ("email",):
                width = 160
            elif c in ("motivo", "incidencia", "solucion"):
                width = 180
            elif c in (
                "fecha_registro",
                "fecha_tramitacion",
                "fecha_rechazo",
                "fecha_inicio_suspension",
                "fecha_fin_suspension",
            ):
                width = 140
            elif c in ("devolucion_recibo",):
                width = 140
            elif c in ("reporte",):
                width = 80
            tree.column(c, anchor="center", width=width, stretch=True)
        def sort_column(col, reverse):
            datos = [(tree.set(k, col), k) for k in tree.get_children("")]
            try:
                datos.sort(key=lambda t: float(t[0].replace(",", ".")), reverse=reverse)
            except ValueError:
                datos.sort(key=lambda t: t[0], reverse=reverse)
            for index, (_, k) in enumerate(datos):
                tree.move(k, "", index)
            tree.heading(col, command=lambda: sort_column(col, not reverse))

        for c in cols:
            tree.heading(c, command=lambda col=c: sort_column(col, False))

        base_widths = {c: tree.column(c, "width") for c in cols}
        tree.grid(row=0, column=0, sticky="nsew")
        vscroll = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        hscroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        tree.bind("<Control-c>", lambda e: self.copiar_celda(e, tree))
        tree.tag_configure("pendiente", background="#ffe6cc")
        tree.tag_configure("tramitada", background="#d4edda")
        tree.tag_configure("concluida", background="#d9edf7")
        tree.tag_configure("rechazada", background="#f8d7da")
        tree.after(200, lambda: self._suspensiones_autofit_columns(tree, cols, base_widths))

        def on_right_click(event):
            row = tree.identify_row(event.y)
            if not row:
                return
            tree.selection_set(row)
            menu = tk.Menu(tree, tearoff=0)
            menu.add_command(label="Modificar estado", command=lambda: self._suspensiones_cambiar_estado(row))
            menu.add_command(label="Ver motivo", command=lambda: self._suspensiones_ver_texto(row, "motivo", "Motivo"))
            menu.add_command(label="Modificar motivo", command=lambda: self._suspensiones_editar_campo(row, "motivo", "Motivo", "Modificar motivo:"))
            menu.add_command(label="Agregar incidencia", command=lambda: self._suspensiones_editar_campo(row, "incidencia", "Incidencia", "Agregar incidencia:"))
            menu.add_command(label="Modificar incidencia", command=lambda: self._suspensiones_editar_campo(row, "incidencia", "Incidencia", "Modificar incidencia:"))
            menu.add_command(label="Ver incidencia", command=lambda: self._suspensiones_ver_texto(row, "incidencia", "Incidencia"))
            menu.add_command(label="Agregar solucion", command=lambda: self._suspensiones_editar_campo(row, "solucion", "Solucion", "Agregar solucion:"))
            menu.add_command(label="Modificar solucion", command=lambda: self._suspensiones_editar_campo(row, "solucion", "Solucion", "Modificar solucion:"))
            menu.add_command(label="Ver solucion", command=lambda: self._suspensiones_ver_texto(row, "solucion", "Solucion"))
            menu.add_command(label="Ver solicitudes individuales", command=lambda: self._suspensiones_ver_solicitudes_individuales(row))
            menu.add_command(label="Modificar fecha inicio suspension", command=lambda: self._suspensiones_editar_fecha(row, "fecha_inicio_suspension"))
            menu.add_command(label="Modificar fecha fin suspension", command=lambda: self._suspensiones_editar_fecha(row, "fecha_fin_suspension"))
            menu.add_command(label="Ver reporte grafico", command=lambda: self._suspensiones_ver_reporte(row))
            menu.add_command(label="Agregar reporte grafico", command=lambda: self._suspensiones_agregar_reporte(row))
            menu.add_command(label="Eliminar reporte grafico", command=lambda: self._suspensiones_eliminar_reporte(row))
            menu.add_separator()
            menu.add_command(label="Abrir chat", command=lambda: self._suspensiones_abrir_chat(row))
            menu.add_command(label="Enviar email", command=lambda: self._suspensiones_enviar_email(row))
            menu.add_separator()
            menu.add_command(label="Eliminar registro", command=lambda: self._suspensiones_eliminar_registro(row))
            menu.add_command(label="Copiar", command=lambda: self.copiar_celda(event, tree))
            menu.post(event.x_root, event.y_root)

        def on_double_click(event):
            row = tree.identify_row(event.y)
            if row:
                self._suspensiones_cambiar_estado(row)

        tooltip = {"win": None, "label": None, "text": ""}

        def show_tooltip(texto, x, y):
            if not texto:
                return
            if tooltip["win"] is None:
                tip = tk.Toplevel(self)
                tip.wm_overrideredirect(True)
                label = tk.Label(
                    tip,
                    text=texto,
                    justify="left",
                    background="#ffffe0",
                    relief="solid",
                    borderwidth=1,
                    wraplength=360,
                )
                label.pack(ipadx=4, ipady=2)
                tooltip["win"] = tip
                tooltip["label"] = label
                tooltip["text"] = texto
            elif tooltip["text"] != texto and tooltip["label"] is not None:
                tooltip["label"].config(text=texto)
                tooltip["text"] = texto
            tooltip["win"].wm_geometry(f"+{x}+{y}")

        def hide_tooltip():
            if tooltip["win"] is not None:
                tooltip["win"].destroy()
                tooltip["win"] = None
            tooltip["label"] = None
            tooltip["text"] = ""

        def on_hover(event):
            item = tree.identify_row(event.y)
            if not item:
                hide_tooltip()
                return
            col = tree.identify_column(event.x)
            col_index = int(col[1:]) - 1 if col else -1
            values = tree.item(item, "values")
            texto = values[col_index] if col_index < len(values) else ""
            if texto:
                show_tooltip(str(texto), event.x_root + 12, event.y_root + 12)
            else:
                hide_tooltip()

        tree.bind("<Button-3>", on_right_click)
        tree.bind("<Double-1>", on_double_click)
        tree.bind("<Motion>", on_hover)
        tree.bind("<Leave>", lambda _e: hide_tooltip())
        self.tree_suspensiones = tree
        tab.tree = tree
        self.refrescar_suspensiones_tree()

    def _suspensiones_autofit_columns(self, tree, cols, base_widths):
        if not tree.winfo_exists():
            return
        available = tree.winfo_width()
        if available <= 1:
            tree.after(100, lambda: self._suspensiones_autofit_columns(tree, cols, base_widths))
            return
        total = sum(base_widths.get(c, 100) for c in cols)
        if total <= 0:
            return
        scale = min(1.0, available / total)
        min_width = 50
        for c in cols:
            base = base_widths.get(c, 100)
            width = max(min_width, int(base * scale))
            tree.column(c, width=width, stretch=False)

    def _suspensiones_ver_texto(self, susp_id, field, title):
        item = next((b for b in self.suspensiones if b.get("id") == susp_id), None)
        if not item:
            return
        texto = (item.get(field) or "").strip()
        if not texto:
            texto = f"Sin {field}."
        messagebox.showinfo(title, texto)

    def cargar_suspensiones(self):
        try:
            data = self._state_get("suspensiones", [], self.suspensiones_file)
            self.suspensiones = data if isinstance(data, list) else []
            changed = False
            for item in self.suspensiones:
                if "id" not in item:
                    item["id"] = uuid.uuid4().hex
                    changed = True
                if "estado" not in item:
                    item["estado"] = "PENDIENTE"
                    changed = True
                if "fecha_registro" not in item:
                    item["fecha_registro"] = ""
                    changed = True
                if "fecha_tramitacion" not in item:
                    item["fecha_tramitacion"] = ""
                    changed = True
                if "fecha_rechazo" not in item:
                    item["fecha_rechazo"] = ""
                    changed = True
                if "fecha_inicio_suspension" not in item:
                    item["fecha_inicio_suspension"] = ""
                    changed = True
                if "fecha_fin_suspension" not in item:
                    item["fecha_fin_suspension"] = ""
                    changed = True
                if "fecha_concluida" not in item:
                    item["fecha_concluida"] = ""
                    changed = True
                if "devolucion_recibo" not in item:
                    item["devolucion_recibo"] = ""
                    changed = True
                if "reporte_path" not in item:
                    item["reporte_path"] = ""
                    changed = True
                if "incidencia" not in item:
                    item["incidencia"] = ""
                    changed = True
                if "solucion" not in item:
                    item["solucion"] = ""
                    changed = True
                if "fin_notificado" not in item:
                    item["fin_notificado"] = "NO"
                    changed = True
            if changed:
                self.guardar_suspensiones()
        except Exception:
            self.suspensiones = []
        self._suspensiones_actualizar_devolucion()
        self._suspensiones_actualizar_concluidas()
        if hasattr(self, "tree_suspensiones"):
            self.refrescar_suspensiones_tree()
        self._suspensiones_notificar_fin()

    def guardar_suspensiones(self):
        try:
            self._state_set("suspensiones", self.suspensiones, self.suspensiones_file)
        except Exception:
            pass

    def _suspensiones_now_str(self):
        return datetime.now().strftime("%d/%m/%Y %H:%M")

    def _suspensiones_parse_dt(self, value):
        raw = str(value or "").strip()
        if not raw:
            return None
        for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%y %H:%M", "%d/%m/%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(raw, fmt)
            except Exception:
                continue
        return None

    def _suspensiones_normalize_date(self, value):
        dt = self._suspensiones_parse_dt(value)
        if not dt:
            return None, None
        return dt, dt.strftime("%d/%m/%Y")

    def _suspensiones_prompt_fecha(self, title, label, initial="", min_date=None, max_date=None):
        while True:
            raw = self._incidencias_prompt_text(title, label, initial=initial)
            if raw is None:
                return None, None
            dt, normalized = self._suspensiones_normalize_date(raw)
            if not dt:
                messagebox.showwarning("Fecha", "Formato de fecha invalido. Usa dd/mm/aaaa.")
                continue
            date_value = dt.date()
            if min_date and date_value < min_date:
                messagebox.showwarning("Fecha", "La fecha no puede ser anterior a la fecha minima permitida.")
                continue
            if max_date and date_value > max_date:
                messagebox.showwarning("Fecha", "La fecha no puede ser posterior a la fecha maxima permitida.")
                continue
            return dt, normalized

    def _suspensiones_get_impagos_set(self):
        fecha = self.impagos_last_export or self.impagos_db.get_last_export()
        if not fecha:
            return set()
        try:
            with self.impagos_db._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    "SELECT DISTINCT c.numero_cliente "
                    "FROM impagos_clientes c "
                    "JOIN impagos_eventos e ON e.cliente_id = c.id "
                    "WHERE e.fecha_export = ?",
                    (fecha,),
                )
                return {str(r[0]).strip() for r in cur.fetchall() if r and r[0]}
        except Exception:
            return set()

    def _suspensiones_actualizar_devolucion(self):
        impagos_set = self._suspensiones_get_impagos_set()
        self.suspensiones_impagos_set = set(impagos_set)
        changed = False
        for item in self.suspensiones:
            codigo = str(item.get("codigo", "")).strip()
            valor = "SI" if codigo and codigo in impagos_set else "NO"
            if item.get("devolucion_recibo") != valor:
                item["devolucion_recibo"] = valor
                changed = True
        if changed:
            self.guardar_suspensiones()

    def _suspensiones_actualizar_concluidas(self):
        hoy = datetime.now().date()
        changed = False
        for item in self.suspensiones:
            estado = str(item.get("estado", "")).upper()
            if estado == "CONCLUIDA":
                continue
            dt_fin = self._suspensiones_parse_dt(item.get("fecha_fin_suspension"))
            if not dt_fin:
                continue
            if dt_fin.date() <= hoy:
                item["estado"] = "CONCLUIDA"
                if not item.get("fecha_concluida"):
                    item["fecha_concluida"] = self._suspensiones_now_str()
                changed = True
        if changed:
            self.guardar_suspensiones()

    def _suspensiones_buscar_cliente_info(self, codigo):
        codigo = str(codigo or "").strip()
        if not codigo:
            return {}
        if self.resumen_df is not None:
            cols = {self._norm(c): c for c in self.resumen_df.columns}
            col_codigo = cols.get("NUMERO DE CLIENTE") or cols.get("NUMERO DE SOCIO")
            col_nombre = cols.get("NOMBRE")
            col_apellidos = cols.get("APELLIDOS")
            col_email = cols.get("CORREO ELECTRONICO") or cols.get("EMAIL") or cols.get("CORREO")
            col_movil = cols.get("MOVIL") or cols.get("TELEFONO") or cols.get("TELEFONO MOVIL")
            if col_codigo:
                df = self.resumen_df[self.resumen_df[col_codigo].astype(str).str.strip() == codigo]
                if not df.empty:
                    row = df.iloc[0]
                    return {
                        "nombre": str(row.get(col_nombre, "")).strip(),
                        "apellidos": str(row.get(col_apellidos, "")).strip(),
                        "email": str(row.get(col_email, "")).strip(),
                        "movil": str(row.get(col_movil, "")).strip(),
                    }
        for c in self.clientes_ext:
            if str(c.get("codigo", "")).strip() == codigo:
                return {
                    "nombre": str(c.get("nombre", "")).strip(),
                    "apellidos": str(c.get("apellidos", "")).strip(),
                    "email": str(c.get("email", "")).strip(),
                    "movil": str(c.get("movil", "")).strip(),
                }
        return {}

    def _suspensiones_select_option(self, title, prompt, opciones, width=28):
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text=prompt).pack(padx=10, pady=6)
        var = tk.StringVar(value=opciones[0] if opciones else "")
        combo = ttk.Combobox(win, values=opciones, textvariable=var, state="readonly", width=width)
        combo.pack(padx=10, pady=6)
        combo.focus_set()
        res = {"value": None}

        def aceptar():
            res["value"] = var.get()
            win.destroy()

        def cancel():
            res["value"] = None
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=6)
        tk.Button(btns, text="Aceptar", command=aceptar).pack(side="left", padx=6)
        tk.Button(btns, text="Cancelar", command=cancel).pack(side="left", padx=6)
        win.bind("<Return>", lambda _e: aceptar())
        win.bind("<Escape>", lambda _e: cancel())
        win.grab_set()
        self._incidencias_center_window(win)
        win.wait_window()
        return res["value"]

    def _suspensiones_nuevo_registro(self):
        staff = self._staff_select("Staff", "Que staff registra?")
        if not staff:
            return
        codigo = self._incidencias_prompt_text("Cliente", "Numero de cliente:")
        if not codigo:
            return
        codigo = str(codigo).strip()
        ya_existe = any(b.get("codigo", "").strip() == codigo for b in self.suspensiones)
        if ya_existe:
            messagebox.showwarning("Suspensiones", f"El cliente {codigo} ya tiene registros previos.")
        motivo = self._suspensiones_select_option(
            "Motivo",
            "Motivo:",
            ["MEDICO", "LABORAL", "TRASLADO", "OTRO"],
        )
        if not motivo:
            return
        if motivo == "OTRO":
            otro = self._incidencias_prompt_text("Motivo", "Describe el motivo:")
            if not otro:
                return
            motivo = f"OTRO: {otro.strip()}"
        today = datetime.now().date()
        dt_inicio, fecha_inicio_norm = self._suspensiones_prompt_fecha(
            "Suspension",
            "Fecha inicio suspension (dd/mm/aaaa):",
            min_date=today,
        )
        if not dt_inicio:
            return
        dt_fin, fecha_fin_norm = self._suspensiones_prompt_fecha(
            "Suspension",
            "Fecha fin suspension (dd/mm/aaaa):",
            min_date=dt_inicio.date(),
        )
        if not dt_fin:
            return
        incidencia = ""
        if messagebox.askyesno("Incidencia", "Desea agregar una incidencia?", parent=self):
            inc = self._incidencias_prompt_text("Incidencia", "Describe la incidencia:")
            if inc:
                incidencia = inc.strip()
        reporte_path = ""
        if messagebox.askyesno("Reporte grafico", "Desea agregar reporte grafico?", parent=self):
            nuevo = self._incidencias_pedir_reporte_visual()
            if nuevo:
                reporte_path = nuevo

        info = self._suspensiones_buscar_cliente_info(codigo)
        now_str = self._suspensiones_now_str()
        registro = {
            "id": uuid.uuid4().hex,
            "staff": self._staff_display_name(staff),
            "codigo": codigo,
            "email": info.get("email", ""),
            "apellidos": info.get("apellidos", ""),
            "nombre": info.get("nombre", ""),
            "movil": info.get("movil", ""),
            "motivo": motivo,
            "estado": "PENDIENTE",
            "fecha_registro": now_str,
            "fecha_tramitacion": "",
            "fecha_rechazo": "",
            "fecha_inicio_suspension": fecha_inicio_norm,
            "fecha_fin_suspension": fecha_fin_norm,
            "fecha_concluida": "",
            "devolucion_recibo": "SI" if codigo in self.suspensiones_impagos_set else "NO",
            "incidencia": incidencia,
            "solucion": "",
            "reporte_path": reporte_path,
            "fin_notificado": "NO",
        }
        registro["movil"] = self._normalizar_movil(registro.get("movil", ""))
        self.suspensiones.append(registro)
        self.guardar_suspensiones()
        self.refrescar_suspensiones_tree()

    def _suspensiones_set_view(self, view):
        self.suspensiones_view = view
        self.refrescar_suspensiones_tree()

    def _suspensiones_set_cliente_filter(self, codigo):
        self.suspensiones_cliente_filter = (codigo or "").strip()
        if hasattr(self, "suspensiones_buscar_entry"):
            self.suspensiones_buscar_entry.delete(0, "end")
            self.suspensiones_buscar_entry.insert(0, self.suspensiones_cliente_filter)
        self.refrescar_suspensiones_tree()

    def _suspensiones_limpiar_cliente_filter(self):
        self.suspensiones_cliente_filter = ""
        if hasattr(self, "suspensiones_buscar_entry"):
            self.suspensiones_buscar_entry.delete(0, "end")
        self.refrescar_suspensiones_tree()

    def _suspensiones_buscar_cliente(self):
        if not hasattr(self, "suspensiones_buscar_entry"):
            return
        codigo = self.suspensiones_buscar_entry.get().strip()
        if not codigo:
            messagebox.showinfo("Suspensiones", "Introduce un numero de cliente.")
            return
        self._suspensiones_set_cliente_filter(codigo)

    def _suspensiones_ver_solicitudes_individuales(self, susp_id):
        item = next((b for b in self.suspensiones if b.get("id") == susp_id), None)
        if not item:
            return
        codigo = item.get("codigo", "")
        if not codigo:
            return
        self._suspensiones_set_cliente_filter(codigo)

    def _suspensiones_editar_campo(self, susp_id, field, title, prompt):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        nuevo = self._incidencias_prompt_text(title, prompt, item.get(field, ""))
        if nuevo is None:
            return
        item[field] = str(nuevo).strip()
        self.guardar_suspensiones()
        self.refrescar_suspensiones_tree()

    def _suspensiones_editar_fecha(self, susp_id, field):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        other_field = "fecha_fin_suspension" if field == "fecha_inicio_suspension" else "fecha_inicio_suspension"
        other_dt = self._suspensiones_parse_dt(item.get(other_field))
        today = datetime.now().date()
        if field == "fecha_inicio_suspension":
            min_date = today
            max_date = other_dt.date() if other_dt else None
            title = "Fecha inicio"
            label = "Fecha inicio suspension (dd/mm/aaaa):"
        else:
            min_date = other_dt.date() if other_dt else None
            max_date = None
            title = "Fecha fin"
            label = "Fecha fin suspension (dd/mm/aaaa):"
        dt, normalized = self._suspensiones_prompt_fecha(
            title,
            label,
            initial=item.get(field, ""),
            min_date=min_date,
            max_date=max_date,
        )
        if not dt:
            return
        item[field] = normalized
        self.guardar_suspensiones()
        self.refrescar_suspensiones_tree()

    def _suspensiones_eliminar_registro(self, susp_id):
        if not messagebox.askyesno("Eliminar", "Eliminar este registro?", parent=self):
            return
        self.suspensiones = [b for b in self.suspensiones if b.get("id") != susp_id]
        self.guardar_suspensiones()
        self.refrescar_suspensiones_tree()

    def _suspensiones_selector_estado(self):
        opciones = ["PENDIENTE", "TRAMITADA", "RECHAZADA", "CONCLUIDA"]
        return self._suspensiones_select_option("Estado", "Selecciona estado:", opciones, width=18)

    def _suspensiones_cambiar_estado(self, susp_id):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        estado = self._suspensiones_selector_estado()
        if not estado:
            return
        item["estado"] = estado
        now_str = self._suspensiones_now_str()
        if estado == "TRAMITADA":
            item["fecha_tramitacion"] = now_str
        elif estado == "RECHAZADA":
            item["fecha_rechazo"] = now_str
        elif estado == "CONCLUIDA":
            item["fecha_concluida"] = now_str
        self.guardar_suspensiones()
        self.refrescar_suspensiones_tree()

    def _suspensiones_abrir_chat(self, susp_id):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        movil = self._normalizar_movil(item.get("movil", ""))
        if not movil:
            messagebox.showwarning("Chat", "El cliente no tiene movil registrado.")
            return
        webbrowser.open(f"https://wa.me/{movil}")

    def _suspensiones_ver_reporte(self, susp_id):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        reporte = item.get("reporte_path", "")
        if not reporte:
            messagebox.showwarning("Reporte", "No hay reporte grafico.")
            return
        resolved = self._incidencias_resolve_reporte_path(reporte)
        if resolved:
            stored = self._incidencias_store_reporte_path(resolved)
            if stored and stored != reporte:
                item["reporte_path"] = stored
                self.guardar_suspensiones()
                self.refrescar_suspensiones_tree()
        self._incidencias_ver_reporte(resolved or reporte)

    def _suspensiones_agregar_reporte(self, susp_id):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        nuevo = self._incidencias_pedir_reporte_visual()
        if not nuevo:
            return
        item["reporte_path"] = nuevo
        self.guardar_suspensiones()
        self.refrescar_suspensiones_tree()

    def _suspensiones_eliminar_reporte(self, susp_id):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        if not item.get("reporte_path"):
            messagebox.showwarning("Reporte", "No hay reporte grafico.")
            return
        if not messagebox.askyesno("Reporte", "Eliminar el reporte grafico?", parent=self):
            return
        item["reporte_path"] = ""
        self.guardar_suspensiones()
        self.refrescar_suspensiones_tree()

    def _suspensiones_enviar_email(self, susp_id):
        item = next((i for i in self.suspensiones if i.get("id") == susp_id), None)
        if not item:
            return
        email = str(item.get("email", "")).strip()
        if not email:
            messagebox.showwarning("Email", "El cliente no tiene email registrado.")
            return
        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(email)
            messagebox.showwarning("Outlook no disponible", "Email copiado al portapapeles.")
            return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = ""
            mail.Body = ""
            mail.Display()
        except Exception as e:
            messagebox.showerror("Email", f"No se pudo abrir Outlook.\nDetalle: {e}")

    def _suspensiones_notificar_fin(self):
        if not self.suspensiones:
            return
        hoy = datetime.now().date()
        candidatos = []
        for item in self.suspensiones:
            if str(item.get("fin_notificado", "")).upper() == "SI":
                continue
            dt_fin = self._suspensiones_parse_dt(item.get("fecha_fin_suspension"))
            if not dt_fin:
                continue
            if dt_fin.date() <= hoy:
                candidatos.append(item)
        if not candidatos:
            return
        emails = [str(i.get("email", "")).strip() for i in candidatos if str(i.get("email", "")).strip()]
        if not emails:
            return
        if not messagebox.askyesno(
            "Suspensiones",
            "Hay suspensiones que finalizan hoy. Desea enviar el email?",
            parent=self,
        ):
            return
        subject = "Fin de suspensión"
        body = "\n".join(
            [
                "Tu periodo de suspensión ha finalizado.",
                "",
                "Gracias por confiar en Fitness Park Villalobos.",
                "Estamos deseando volver a verte en nuestras instalaciones desafiándote y superándote cada día.",
                "Nos vemos por el club.",
            ]
        )
        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(";".join(emails))
            messagebox.showwarning("Outlook no disponible", "Emails copiados al portapapeles.")
            return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.BCC = ";".join(emails)
            mail.Body = body
            mail.Display()
        except Exception as e:
            messagebox.showerror("Email", f"No se pudo abrir Outlook.\nDetalle: {e}")
            return
        if not messagebox.askyesno("Suspensiones", "Ha enviado el email?", parent=self):
            return
        for item in candidatos:
            item["fin_notificado"] = "SI"
        self.guardar_suspensiones()

    def _suspensiones_order_key(self, item, view):
        if view in ("TRAMITADA", "PENDIENTE", "ACTIVAS"):
            dt = self._suspensiones_parse_dt(item.get("fecha_fin_suspension"))
            if dt:
                return abs((dt - datetime.now()).total_seconds())
            return float("inf")
        return self._suspensiones_parse_dt(item.get("fecha_registro")) or datetime.min

    def refrescar_suspensiones_tree(self):
        if not hasattr(self, "tree_suspensiones"):
            return
        tree = self.tree_suspensiones
        tree.delete(*tree.get_children())
        self._suspensiones_actualizar_concluidas()
        view = (self.suspensiones_view or "ACTIVAS").upper()
        if view == "PENDIENTES":
            view = "PENDIENTE"
        cliente_filter = (self.suspensiones_cliente_filter or "").strip().upper()
        sorted_items = sorted(
            self.suspensiones,
            key=lambda i: self._suspensiones_order_key(i, view),
            reverse=(view not in ("TRAMITADA", "PENDIENTE", "ACTIVAS")),
        )
        for item in sorted_items:
            codigo = str(item.get("codigo", "")).strip()
            if cliente_filter and codigo.upper() != cliente_filter:
                continue
            estado = str(item.get("estado", "PENDIENTE")).upper()
            if view == "ACTIVAS":
                if estado not in ("PENDIENTE", "TRAMITADA"):
                    continue
            elif view != "TODOS" and estado != view:
                continue
            tag = ""
            if estado == "PENDIENTE":
                tag = "pendiente"
            elif estado == "TRAMITADA":
                tag = "tramitada"
            elif estado == "CONCLUIDA":
                tag = "concluida"
            elif estado == "RECHAZADA":
                tag = "rechazada"
            values = [
                item.get("staff", ""),
                item.get("codigo", ""),
                item.get("email", ""),
                item.get("apellidos", ""),
                item.get("nombre", ""),
                item.get("movil", ""),
                item.get("motivo", ""),
                item.get("estado", ""),
                "R" if item.get("reporte_path") else "",
                item.get("fecha_registro", ""),
                item.get("fecha_tramitacion", ""),
                item.get("fecha_rechazo", ""),
                item.get("fecha_inicio_suspension", ""),
                item.get("fecha_fin_suspension", ""),
                item.get("devolucion_recibo", ""),
                item.get("incidencia", ""),
                item.get("solucion", ""),
            ]
            tree.insert("", "end", iid=item.get("id"), values=values, tags=(tag,) if tag else ())

    # -----------------------------
    # INCIDENCIAS CLUB
    # -----------------------------
    def create_incidencias_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        botones = tk.Frame(frm)
        botones.pack(fill="x", pady=4)
        self.btn_inc_cargar_mapas = tk.Button(botones, text="Cargar Mapas", command=self.incidencias_cargar_mapas)
        self.btn_inc_cargar_mapas.pack(side="left", padx=5)
        self.btn_inc_borrar_mapa = tk.Button(botones, text="Borrar Mapa", command=self.incidencias_borrar_mapa)
        self.btn_inc_borrar_mapa.pack(side="left", padx=5)
        self.btn_mapa_anterior = tk.Button(botones, text="Mapa anterior", command=self.incidencias_mapa_anterior)
        self.btn_mapa_anterior.pack(side="left", padx=5)
        self.btn_mapa_siguiente = tk.Button(botones, text="Mapas", command=self.incidencias_siguiente_mapa)
        self.btn_mapa_siguiente.pack(side="left", padx=5)
        self.btn_inc_asignar_area = tk.Button(botones, text="Asignar Area", command=self.incidencias_asignar_area)
        self.btn_inc_asignar_area.pack(side="left", padx=5)
        self.btn_inc_editar_area = tk.Button(botones, text="Editar Area", command=self.incidencias_editar_area)
        self.btn_inc_editar_area.pack(side="left", padx=5)
        self.btn_inc_listar_maquina = tk.Button(botones, text="Listar Maquina", command=self.incidencias_listar_maquina)
        self.btn_inc_listar_maquina.pack(side="left", padx=5)
        self.btn_inc_info_maquinas = tk.Button(botones, text="Info maquinas", command=self.incidencias_info_maquinas)
        self.btn_inc_info_maquinas.pack(side="left", padx=5)
        self.btn_inc_crear_incidencia = tk.Button(botones, text="Crear Incidencia", command=self.incidencias_crear_incidencia, bg="#ffcc80", fg="black")
        self.btn_inc_crear_incidencia.pack(side="left", padx=5)
        self.btn_inc_gestion_incidencias = tk.Button(botones, text="Gestion Incidencias", command=self.incidencias_gestion_incidencias, bg="#a5d6a7", fg="black")
        self.btn_inc_gestion_incidencias.pack(side="left", padx=5)
        vista_btn = tk.Menubutton(botones, text="VISTA", bg="#bdbdbd", fg="black")
        self.btn_inc_vista = vista_btn
        vista_menu = tk.Menu(vista_btn, tearoff=0)
        vista_menu.add_command(label="TODAS LAS INCIDENCIAS", command=lambda: self._incidencias_set_filtro_estado("TODAS"))
        vista_menu.add_command(label="PENDIENTES", command=lambda: self._incidencias_set_filtro_estado("PENDIENTE"))
        vista_menu.add_command(label="VISTAS", command=lambda: self._incidencias_set_filtro_estado("VISTO"))
        vista_menu.add_command(label="VISTAS Y PENDIENTES", command=lambda: self._incidencias_set_filtro_estado("VISTO_PENDIENTE"))
        vista_menu.add_command(label="REPARADAS", command=lambda: self._incidencias_set_filtro_estado("REPARADO"))
        vista_btn.configure(menu=vista_menu)
        vista_btn.pack(side="left", padx=5)

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
        self.incidencias_canvas.bind("<Button-3>", self.incidencias_canvas_right_click)
        self.incidencias_canvas.bind("<MouseWheel>", self.incidencias_canvas_mousewheel)
        self.incidencias_canvas.bind("<Shift-MouseWheel>", self.incidencias_canvas_mousewheel_x)
        self.incidencias_canvas.bind("<Button-4>", lambda _e: self.incidencias_canvas.yview_scroll(-1, "units"))
        self.incidencias_canvas.bind("<Button-5>", lambda _e: self.incidencias_canvas.yview_scroll(1, "units"))

        self.incidencias_cargar_listado_mapas()
        self.incidencias_gestion_incidencias()

    def _security_pin_ok(self):
        self._bring_to_front()
        pin = simpledialog.askstring("Código de seguridad", "Introduce el código de seguridad:", show="*")
        if pin is None:
            return False
        if pin.strip() != get_security_code():
            messagebox.showerror("Codigo incorrecto", "El codigo de seguridad no es valido.", parent=self)
            return False
        return True

    def _incidencias_pin_ok(self):
        if not self._security_pin_ok():
            return False
        return self._require_manager_access("Incidencias club")

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
    def _incidencias_center_window(self, win):
        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(int((sw - w) / 2), 0)
        y = max(int((sh - h) / 2), 0)
        win.geometry(f"{w}x{h}+{x}+{y}")

    def _incidencias_prompt_text(self, title, label, initial=""):
        self._bring_to_front()
        win = tk.Toplevel(self)
        win.title(title)
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text=label).pack(padx=10, pady=(10, 4))
        entry = tk.Entry(win, width=40)
        entry.pack(padx=10, pady=4)
        if initial:
            entry.insert(0, initial)
        entry.focus_set()
        result = {"value": None}

        def accept():
            result["value"] = entry.get()
            win.destroy()

        def cancel():
            result["value"] = None
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(pady=8)
        tk.Button(btns, text="OK", width=8, command=accept).pack(side="left", padx=5)
        tk.Button(btns, text="Cancel", width=8, command=cancel).pack(side="left", padx=5)
        win.bind("<Return>", lambda _e: accept())
        win.bind("<Escape>", lambda _e: cancel())
        win.grab_set()
        self._incidencias_center_window(win)
        win.wait_window()
        return result["value"]

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
                ("JPEG", "*.jpg;*.jpeg"),
            ],
        )
        if not ruta:
            return ""
        ext = os.path.splitext(ruta)[1].lower()
        if ext not in (".jpg", ".jpeg"):
            messagebox.showwarning("Reporte visual", "Solo se admiten imagenes .jpg o .jpeg.", parent=self)
            return ""
        blob_ref = self._store_image_blob(ruta, allow_png=False)
        if blob_ref:
            return blob_ref
        try:
            reports_dir = os.path.join(self.data_dir, "incidencias_reportes")
            os.makedirs(reports_dir, exist_ok=True)
            destino = os.path.join(reports_dir, f"reporte_{uuid.uuid4().hex}{ext}")
            shutil.copy2(ruta, destino)
            return os.path.basename(destino)
        except Exception:
            return ruta

    def _incidencias_ver_reporte(self, ruta):
        if self._is_blob_ref(ruta):
            store = getattr(self, "state_store", None)
            if not store or not store.use_postgres:
                messagebox.showwarning("Reporte visual", "Reporte en BD, pero no hay conexion.", parent=self)
                return
            _ctype, data = store.get_blob(self._blob_id_from_ref(ruta))
            if not data:
                messagebox.showwarning("Reporte visual", "No se encontro el reporte en la BD.", parent=self)
                return
            try:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
                tmp.write(data)
                tmp.close()
                os.startfile(tmp.name)
            except Exception as exc:
                messagebox.showerror("Reporte visual", f"No se pudo abrir el reporte.\nDetalle: {exc}", parent=self)
            return
        ruta_resuelta = self._incidencias_resolve_reporte_path(ruta)
        if not ruta_resuelta:
            reports_dir = os.path.join(self.data_dir, "incidencias_reportes")
            messagebox.showwarning(
                "Reporte visual",
                f"No se encontro el archivo del reporte.\n\nRuta guardada:\n{ruta}\n\nCarpeta:\n{reports_dir}",
                parent=self,
            )
            return
        try:
            os.startfile(ruta_resuelta)
        except Exception as exc:
            messagebox.showerror("Reporte visual", f"No se pudo abrir el reporte.\nDetalle: {exc}", parent=self)

    def _incidencias_guardar_reporte(self, ruta):
        if self._is_blob_ref(ruta):
            store = getattr(self, "state_store", None)
            if not store or not store.use_postgres:
                messagebox.showwarning("Reporte visual", "Reporte en BD, pero no hay conexion.", parent=self)
                return
            _ctype, data = store.get_blob(self._blob_id_from_ref(ruta))
            if not data:
                messagebox.showwarning("Reporte visual", "No se encontro el reporte en la BD.", parent=self)
                return
            destino = filedialog.asksaveasfilename(
                parent=self,
                title="Guardar reporte visual",
                initialfile="reporte.jpg",
                defaultextension=".jpg",
            )
            if not destino:
                return
            try:
                with open(destino, "wb") as f:
                    f.write(data)
                messagebox.showinfo("Reporte visual", f"Reporte guardado en: {destino}", parent=self)
            except Exception as exc:
                messagebox.showerror("Reporte visual", f"No se pudo guardar el reporte.\nDetalle: {exc}", parent=self)
            return
        ruta_resuelta = self._incidencias_resolve_reporte_path(ruta)
        if not ruta_resuelta:
            messagebox.showwarning("Reporte visual", "No se encontro el archivo del reporte.", parent=self)
            return
        filename = os.path.basename(ruta_resuelta)
        destino = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar reporte visual",
            initialfile=filename,
            defaultextension=os.path.splitext(filename)[1],
        )
        if not destino:
            return
        try:
            shutil.copy2(ruta_resuelta, destino)
            messagebox.showinfo("Reporte visual", f"Reporte guardado en: {destino}", parent=self)
        except Exception as exc:
            messagebox.showerror("Reporte visual", f"No se pudo guardar el reporte.\nDetalle: {exc}", parent=self)

    def _incidencias_chat_trabajador(self, movil):
        movil = self._normalizar_movil(movil)
        if not movil:
            messagebox.showwarning("Chat", "El trabajador no tiene movil.", parent=self)
            return
        url = f"https://wa.me/{movil}"
        webbrowser.open(url)

    def incidencias_cargar_listado_mapas(self):
        if not self.incidencias_db:
            self.incidencias_mapas = []
            return
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
        maps_dir = os.path.join(self.data_dir, "maps")
        os.makedirs(maps_dir, exist_ok=True)
        for _ in range(cantidad):
            ruta = filedialog.askopenfilename(
                filetypes=[("Imagen", "*.png;*.jpg;*.jpeg")]
            )
            if not ruta:
                continue
            nombre = os.path.basename(ruta)
            blob_ref = ""
            store = getattr(self, "state_store", None)
            if store and store.use_postgres:
                blob_ref = self._store_image_blob(ruta, allow_png=True)
            if blob_ref:
                img = Image.open(ruta)
                self.incidencias_db.add_map(nombre, blob_ref, img.width, img.height)
            else:
                destino = os.path.join(maps_dir, nombre)
                shutil.copy2(ruta, destino)
                img = Image.open(destino)
                self.incidencias_db.add_map(nombre, nombre, img.width, img.height)
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
            if self._is_blob_ref(ruta):
                store = getattr(self, "state_store", None)
                if store and store.use_postgres:
                    store.delete_blob(self._blob_id_from_ref(ruta))
            else:
                ruta_resuelta = self._incidencias_resolve_map_path(ruta)
                if ruta_resuelta and os.path.exists(ruta_resuelta):
                    os.remove(ruta_resuelta)
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
        if self._is_blob_ref(ruta):
            store = getattr(self, "state_store", None)
            if not store or not store.use_postgres:
                self._bring_to_front()
                messagebox.showwarning("Mapas", "Mapa en BD, pero no hay conexion.", parent=self)
                return
            _ctype, data = store.get_blob(self._blob_id_from_ref(ruta))
            if not data:
                self._bring_to_front()
                messagebox.showwarning("Mapas", "No se encontro el mapa en la BD.", parent=self)
                return
            try:
                img = Image.open(io.BytesIO(data))
            except Exception as e:
                self._bring_to_front()
                messagebox.showerror("Mapas", f"No se pudo abrir el mapa en BD.\n\nDetalle: {e}", parent=self)
                return
        else:
            ruta_resuelta = self._incidencias_resolve_map_path(ruta, nombre)
            if not ruta_resuelta:
                self._bring_to_front()
                messagebox.showwarning(
                    "Mapas",
                    f"No se encontro el archivo del mapa.\n\nRuta guardada:\n{ruta}\n\nNombre:\n{nombre}",
                    parent=self,
                )
                return
            if ruta_resuelta != ruta:
                self.incidencias_db.update_map_path(mapa_id, ruta_resuelta)
            try:
                img = Image.open(ruta_resuelta)
            except Exception as e:
                self._bring_to_front()
                messagebox.showerror(
                    "Mapas",
                    f"No se pudo abrir el mapa:\n{ruta_resuelta}\n\nDetalle: {e}",
                    parent=self,
                )
                return
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

    def _incidencias_resolve_map_path(self, ruta, nombre=None):
        if not ruta:
            return ""
        if self._is_blob_ref(ruta):
            return ""
        ruta_norm = os.path.normpath(ruta)
        if os.path.isabs(ruta_norm) and os.path.exists(ruta_norm):
            return ruta_norm
        maps_dir = os.path.join(self.data_dir, "maps")
        if not os.path.isabs(ruta_norm):
            candidate = os.path.join(maps_dir, ruta_norm)
            if os.path.exists(candidate):
                return candidate
        base = os.path.basename(ruta_norm)
        fallback = os.path.join(maps_dir, base)
        if os.path.exists(fallback):
            return fallback
        if nombre:
            nombre_norm = os.path.normpath(str(nombre))
            nombre_base = os.path.basename(nombre_norm)
            candidate = os.path.join(maps_dir, nombre_base)
            if os.path.exists(candidate):
                return candidate
            if not os.path.splitext(nombre_base)[1]:
                candidate = os.path.join(maps_dir, f"{nombre_base}.png")
                if os.path.exists(candidate):
                    return candidate
        return ""

    def _incidencias_resolve_reporte_path(self, ruta):
        if not ruta:
            return ""
        if self._is_blob_ref(ruta):
            return ""
        ruta_norm = os.path.normpath(ruta)
        reports_dir = os.path.join(self.data_dir, "incidencias_reportes")
        if os.path.isabs(ruta_norm) and os.path.exists(ruta_norm):
            return ruta_norm
        if not os.path.isabs(ruta_norm):
            candidate = os.path.join(reports_dir, ruta_norm)
            if os.path.exists(candidate):
                return candidate
        base = os.path.basename(ruta_norm)
        candidate = os.path.join(reports_dir, base)
        if os.path.exists(candidate):
            return candidate
        return ""

    def _incidencias_store_reporte_path(self, ruta_resuelta):
        if not ruta_resuelta:
            return ""
        reports_dir = os.path.normpath(os.path.join(self.data_dir, "incidencias_reportes"))
        ruta_norm = os.path.normpath(ruta_resuelta)
        if ruta_norm.startswith(reports_dir + os.sep) or ruta_norm == reports_dir:
            return os.path.basename(ruta_norm)
        return ruta_resuelta

    def incidencias_asignar_area(self):
        if not self._incidencias_pin_ok():
            return
        self._bring_to_front()
        nombre = self._incidencias_prompt_text("Área", "¿Cómo se va a llamar el área?")
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
            tree.heading(c, text=c.capitalize(), command=lambda col=c: self._treeview_sort(tree, col))
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

        def edit_machine(event):
            row = tree.identify_row(event.y)
            if not row:
                return
            mid = int(row)
            if not self._incidencias_pin_ok():
                return
            machine = None
            for m in self.incidencias_db.list_machines(self.incidencias_current_map):
                if int(m[0]) == mid:
                    machine = m
                    break
            if not machine:
                return
            _, _area_id, nombre, serie, numero, x1, y1, x2, y2, _color, _area_nombre = machine
            self._bring_to_front()
            nombre_n = self._incidencias_prompt_text("Maquina", "Nombre de la maquina:", nombre)
            if nombre_n is None:
                return
            serie_n = self._incidencias_prompt_text("Maquina", "Numero de serie:", serie)
            numero_n = self._incidencias_prompt_text("Maquina", "Numero asignado:", numero)
            self.incidencias_db.update_machine(
                mid,
                nombre_n or "",
                serie_n or "",
                numero_n or "",
                int(x1),
                int(y1),
                int(x2),
                int(y2),
            )
            self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
            self.incidencias_info_maquinas(self.incidencias_info_filter_area)

        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label="Copiar", command=lambda: copy_cell(tree.event_context))
        if self._can_write("incidencias_club"):
            menu.add_command(label="Editar maquina", command=lambda: edit_machine(tree.event_context))
            menu.add_command(label="Eliminar maquina", command=lambda: delete_machine(tree.event_context))
        tree.bind("<<TreeviewSelect>>", on_select)
        tree.bind("<Motion>", on_hover)
        tree.bind("<Leave>", on_leave)
        tree.bind("<Button-3>", lambda e: (setattr(tree, 'event_context', e), menu.post(e.x_root, e.y_root)))
        self.incidencias_machines_tree = tree

    def create_staff_tab(self, tab):
        frm = tk.Frame(tab)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        header = tk.Frame(frm)
        header.pack(fill="x", pady=4)
        tk.Button(header, text="Agregar Staff", command=self._staff_agregar).pack(side="left", padx=5)
        tk.Label(header, text="Buscar:").pack(side="left", padx=(10, 4))
        filtro_var = tk.StringVar()
        filtro_entry = tk.Entry(header, textvariable=filtro_var, width=24)
        filtro_entry.pack(side="left", padx=5)

        cols = ["nombre", "apellido1", "apellido2", "movil", "email"]
        tree = ttk.Treeview(frm, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c.capitalize())
            tree.column(c, anchor="center")
        tree.column("nombre", width=140, stretch=True)
        tree.column("apellido1", width=140, stretch=True)
        tree.column("apellido2", width=140, stretch=True)
        tree.column("movil", width=140, stretch=True)
        tree.column("email", width=220, stretch=True)
        tree.pack(fill="both", expand=True)

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
            values = tree.item(item, "values")
            movil = values[3] if len(values) > 3 else ""
            menu = tk.Menu(tree, tearoff=0)
            menu.add_command(label="Editar", command=lambda: self._staff_editar(item))
            menu.add_command(label="Eliminar", command=lambda: self._staff_eliminar(item))
            if movil:
                menu.add_command(label="Abrir chat", command=lambda: self._staff_open_chat(movil))
            menu.add_command(label="Copiar", command=lambda: copy_cell(event))
            menu.post(event.x_root, event.y_root)

        tree.bind("<Button-3>", on_right_click)
        filtro_entry.bind("<KeyRelease>", lambda _e: self._staff_populate_tree(tree, filtro_var.get()))
        filtro_entry.focus_set()
        self.staff_tree = tree
        self.staff_filter_var = filtro_var
        self._staff_populate_tree(tree, "")

    def _staff_agregar(self):
        self._bring_to_front()
        nombre = self._incidencias_prompt_text("Staff", "Nombre:")
        if nombre is None:
            return
        apellido1 = self._incidencias_prompt_text("Staff", "Apellido 1:")
        apellido2 = self._incidencias_prompt_text("Staff", "Apellido 2:")
        movil = self._incidencias_prompt_text("Staff", "Movil:")
        email = self._incidencias_prompt_text("Staff", "Email:")
        movil = self._normalizar_movil(movil)
        item = {
            "id": uuid.uuid4().hex,
            "nombre": nombre or "",
            "apellido1": apellido1 or "",
            "apellido2": apellido2 or "",
            "movil": movil or "",
            "email": (email or "").strip(),
        }
        self.staff.append(item)
        self.guardar_staff()
        self._staff_refresh_views()

    def _staff_editar(self, staff_id):
        item = next((s for s in self.staff if s.get("id") == staff_id), None)
        if not item:
            return
        self._bring_to_front()
        nombre = self._incidencias_prompt_text("Staff", "Nombre:", item.get("nombre", ""))
        apellido1 = self._incidencias_prompt_text("Staff", "Apellido 1:", item.get("apellido1", ""))
        apellido2 = self._incidencias_prompt_text("Staff", "Apellido 2:", item.get("apellido2", ""))
        movil = self._incidencias_prompt_text("Staff", "Movil:", item.get("movil", ""))
        email = self._incidencias_prompt_text("Staff", "Email:", item.get("email", ""))
        if nombre is None:
            return
        item["nombre"] = nombre or ""
        item["apellido1"] = apellido1 or ""
        item["apellido2"] = apellido2 or ""
        item["movil"] = self._normalizar_movil(movil)
        item["email"] = (email or "").strip()
        self.guardar_staff()
        self._staff_refresh_views()

    def _staff_eliminar(self, staff_id):
        if not messagebox.askyesno("Staff", "Eliminar este staff?", parent=self):
            return
        self.staff = [s for s in self.staff if s.get("id") != staff_id]
        self.guardar_staff()
        self._staff_refresh_views()

    def _staff_open_chat(self, movil):
        movil = self._normalizar_movil(movil)
        if not movil:
            messagebox.showwarning("Chat", "El trabajador no tiene movil.", parent=self)
            return
        webbrowser.open(f"https://wa.me/{movil}")

    def abrir_gestion_staff(self):
        if not self._require_manager_access("Gestionar staff"):
            return
        if not self._security_pin_ok():
            return
        self.cargar_staff()
        tab = self.tabs.get("Staff")
        if tab:
            self.notebook.add(tab, text="Staff", image=self.tab_icons.get("Staff"), compound="left")
            self.notebook.select(tab)
        self._staff_refresh_views()

    def _staff_populate_tree(self, tree, filtro_text=""):
        tree.delete(*tree.get_children())
        needle = filtro_text.strip().lower()
        for item in self.staff:
            values = [
                item.get("nombre", ""),
                item.get("apellido1", ""),
                item.get("apellido2", ""),
                item.get("movil", ""),
                item.get("email", ""),
            ]
            haystack = " ".join(str(v) for v in values).lower()
            if needle and needle not in haystack:
                continue
            tree.insert("", "end", iid=item.get("id"), values=values)

    def _staff_refresh_views(self):
        if self.staff_tree is not None and self.staff_filter_var is not None:
            self._staff_populate_tree(self.staff_tree, self.staff_filter_var.get())

    def incidencias_crear_incidencia(self):
        self._bring_to_front()
        creador = self._staff_select("Incidencia", "Quien crea la incidencia?")
        if not creador:
            return
        self.incidencias_creador = creador
        respuesta = messagebox.askyesno("Incidencia", "Es sobre una maquina?", parent=self)
        if respuesta:
            self.incidencias_mode = ("inc_maquina",)
            self._bring_to_front()
            messagebox.showinfo("Incidencia", "Haz click sobre la maquina en el mapa.", parent=self)
        else:
            self.incidencias_mode = ("inc_area",)
            self._bring_to_front()
            messagebox.showinfo("Incidencia", "Haz click sobre el area en el mapa.", parent=self)

    def _incidencias_set_filtro_estado(self, estado):
        self.incidencias_filtro_estado = estado
        self.incidencias_gestion_incidencias()

    def incidencias_gestion_incidencias(self):
        self.incidencias_panel_mode = "incidencias"
        self.incidencias_info_filter_area = None
        self._incidencias_set_area_header(None)
        self._incidencias_apply_map_filter(None)
        for w in self.incidencias_panel.winfo_children():
            w.destroy()
        if not self.incidencias_current_map:
            return
        cols = ["fecha", "creador", "estado", "area", "maquina", "serie", "numero", "elemento", "descripcion", "reporte"]
        container = tk.Frame(self.incidencias_panel)
        container.pack(fill="both", expand=True)
        tree = ttk.Treeview(container, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c.capitalize())
            tree.column(c, anchor="center")
        tree.column("fecha", width=120, stretch=True)
        tree.column("creador", width=140, stretch=True)
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
        filtro_estado = (self.incidencias_filtro_estado or "TODAS").upper()
        for inc in self.incidencias_db.list_incidencias(self.incidencias_current_map):
            (
                inc_id,
                fecha,
                estado,
                elemento,
                descripcion,
                reporte_path,
                creador_nombre,
                creador_apellido1,
                creador_apellido2,
                creador_movil,
                creador_email,
                area,
                maquina,
                serie,
                numero,
            ) = inc
            if reporte_path:
                resolved = self._incidencias_resolve_reporte_path(reporte_path)
                if resolved:
                    stored = self._incidencias_store_reporte_path(resolved)
                    if stored and stored != reporte_path:
                        self.incidencias_db.update_incidencia_reporte(inc_id, stored)
                    reporte_path = stored
            if filtro_estado == "VISTO_PENDIENTE":
                if estado not in ("VISTO", "PENDIENTE"):
                    continue
            elif filtro_estado != "TODAS" and estado != filtro_estado:
                continue
            creador = " ".join([p for p in [creador_nombre, creador_apellido1] if p]).strip()
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
                values=[
                    fecha,
                    creador,
                    estado,
                    area or "",
                    maquina or "",
                    serie or "",
                    numero or "",
                    elemento or "",
                    descripcion or "",
                    ("[R]" if reporte_path else ""),
                ],
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
                "creador_movil": creador_movil,
                "creador_email": creador_email,
                "creador_nombre": creador,
            }
            if area:
                self.incidencias_area_to_inc.setdefault(area, []).append(inc_id)

        tooltip = {"win": None, "label": None, "text": ""}

        def show_tooltip(texto, x, y):
            if not texto:
                return
            if tooltip["win"] is None:
                tip = tk.Toplevel(self)
                tip.wm_overrideredirect(True)
                label = tk.Label(tip, text=texto, justify="left", background="#ffffe0", relief="solid", borderwidth=1, wraplength=380)
                label.pack(ipadx=4, ipady=2)
                tooltip["win"] = tip
                tooltip["label"] = label
                tooltip["text"] = texto
            elif tooltip["text"] != texto and tooltip["label"] is not None:
                tooltip["label"].config(text=texto)
                tooltip["text"] = texto
            tooltip["win"].wm_geometry(f"+{x}+{y}")

        def hide_tooltip():
            if tooltip["win"] is not None:
                tooltip["win"].destroy()
                tooltip["win"] = None
            tooltip["label"] = None
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
            if self._can_write("incidencias_club"):
                menu.add_command(label="Modificar Incidencia", command=lambda: self._incidencias_editar_incidencia(int(item)))
                menu.add_command(label="Eliminar Incidencia", command=lambda: self._incidencias_eliminar_incidencia(int(item)))
                menu.add_command(label="Modificar Estado", command=lambda: self._incidencias_cambiar_estado(int(item)))
                menu.add_command(label="Modificar reporte visual", command=lambda: self._incidencias_cambiar_reporte(int(item)))
                reporte = self.incidencias_inc_map.get(int(item), {}).get("reporte")
                if reporte:
                    menu.add_command(label="Ver reporte visual", command=lambda: self._incidencias_ver_reporte(reporte))
                    menu.add_command(label="Guardar reporte visual", command=lambda: self._incidencias_guardar_reporte(reporte))
                creador_movil = self.incidencias_inc_map.get(int(item), {}).get("creador_movil", "")
                if creador_movil:
                    menu.add_command(label="Abrir chat con trabajador", command=lambda: self._incidencias_chat_trabajador(creador_movil))
            menu.add_command(label="Copiar", command=lambda: copy_cell(event))
            menu.post(event.x_root, event.y_root)

        def on_double_click(event):
            item = tree.identify_row(event.y)
            if not item:
                return
            if not self._can_write("incidencias_club"):
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
        if not self._require_write("incidencias_club"):
            return
        self._bring_to_front()
        elemento = self._incidencias_prompt_text("Incidencia", "Material/elemento:")
        descripcion = self._incidencias_prompt_text("Incidencia", "Describe la incidencia:")
        estado = self._incidencias_selector_estado()
        if not estado:
            estado = "PENDIENTE"
        reporte_actual = None
        if hasattr(self, "incidencias_inc_map"):
            reporte_actual = self.incidencias_inc_map.get(inc_id, {}).get("reporte")
        if messagebox.askyesno("Incidencia", "Desea cambiar reporte visual?", parent=self):
            reporte_nuevo = self._incidencias_pedir_reporte_visual()
            if reporte_nuevo:
                if reporte_actual and reporte_actual != reporte_nuevo and os.path.exists(reporte_actual):
                    try:
                        os.remove(reporte_actual)
                    except Exception:
                        pass
                self.incidencias_db.update_incidencia_reporte(inc_id, reporte_nuevo)
        self.incidencias_db.update_incidencia(inc_id, elemento or "", descripcion or "", estado)
        self.incidencias_gestion_incidencias()

    def _incidencias_cambiar_reporte(self, inc_id):
        if not self._require_write("incidencias_club"):
            return
        self._bring_to_front()
        reporte_actual = None
        if hasattr(self, "incidencias_inc_map"):
            reporte_actual = self.incidencias_inc_map.get(inc_id, {}).get("reporte")
        reporte_nuevo = self._incidencias_pedir_reporte_visual()
        if not reporte_nuevo:
            return
        if reporte_actual and reporte_actual != reporte_nuevo and os.path.exists(reporte_actual):
            try:
                os.remove(reporte_actual)
            except Exception:
                pass
        self.incidencias_db.update_incidencia_reporte(inc_id, reporte_nuevo)
        self.incidencias_gestion_incidencias()

    def _incidencias_eliminar_incidencia(self, inc_id):
        if not self._require_write("incidencias_club"):
            return
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
        if not self._require_write("incidencias_club"):
            return
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

        btn = tk.Button(win, text="Aceptar", command=aceptar)
        btn.pack(pady=5)
        win.bind("<Return>", lambda _e: aceptar())
        win.bind("<Escape>", lambda _e: win.destroy())
        self._incidencias_center_window(win)
        win.grab_set()
        win.wait_window()
        return resultado["valor"]

    def incidencias_canvas_mousewheel(self, event):
        delta = int(-1 * (event.delta / 120))
        self.incidencias_canvas.yview_scroll(delta, "units")

    def incidencias_canvas_mousewheel_x(self, event):
        delta = int(-1 * (event.delta / 120))
        self.incidencias_canvas.xview_scroll(delta, "units")

    def _treeview_sort(self, tree, col):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        if not hasattr(tree, "_sort_reverse"):
            tree._sort_reverse = {}
        reverse = tree._sort_reverse.get(col, False)

        def to_number(val):
            try:
                return float(val)
            except Exception:
                return val

        if all(str(v).strip().replace(".", "", 1).isdigit() for v, _k in items if str(v).strip()):
            items.sort(key=lambda x: to_number(x[0]), reverse=reverse)
        else:
            items.sort(key=lambda x: str(x[0]).lower(), reverse=reverse)

        for index, (_val, k) in enumerate(items):
            tree.move(k, "", index)
        tree._sort_reverse[col] = not reverse

    def _incidencias_get_machine_by_id(self, mid):
        machines = self.incidencias_db.list_machines(self.incidencias_current_map)
        for m in machines:
            if int(m[0]) == int(mid):
                return m
        return None

    def _incidencias_crear_incidencia_maquina(self, mid):
        machine = self._incidencias_get_machine_by_id(mid)
        if not machine:
            return
        _mid, area_id, _nombre, _serie, _numero, *_rest = machine
        self._bring_to_front()
        elemento = self._incidencias_prompt_text("Incidencia", "Material/elemento:")
        descripcion = self._incidencias_prompt_text("Incidencia", "Describe la incidencia:")
        if elemento is None and descripcion is None:
            return
        reporte_path = self._incidencias_pedir_reporte_visual()
        creador = self.incidencias_creador or {}
        self.incidencias_db.add_incident(
            self.incidencias_current_map,
            area_id,
            mid,
            elemento or "",
            descripcion or "",
            reporte_path=reporte_path,
            creador_nombre=creador.get("nombre", ""),
            creador_apellido1=creador.get("apellido1", ""),
            creador_apellido2=creador.get("apellido2", ""),
            creador_movil=creador.get("movil", ""),
            creador_email=creador.get("email", ""),
        )
        self.incidencias_gestion_incidencias()

    def _incidencias_editar_maquina_directa(self, mid):
        machine = self._incidencias_get_machine_by_id(mid)
        if not machine:
            return
        _, _area_id, nombre, serie, numero, x1, y1, x2, y2, _color, _area_nombre = machine
        self._bring_to_front()
        nombre_n = self._incidencias_prompt_text("Maquina", "Nombre de la maquina:", nombre)
        if nombre_n is None:
            return
        serie_n = self._incidencias_prompt_text("Maquina", "Numero de serie:", serie)
        numero_n = self._incidencias_prompt_text("Maquina", "Numero asignado:", numero)
        self.incidencias_db.update_machine(
            mid,
            nombre_n or "",
            serie_n or "",
            numero_n or "",
            int(x1),
            int(y1),
            int(x2),
            int(y2),
        )
        self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
        self.incidencias_info_maquinas(self.incidencias_info_filter_area)

    def _incidencias_eliminar_maquina_directa(self, mid):
        self._bring_to_front()
        if not messagebox.askyesno("Eliminar", "Eliminar maquina y todos sus datos?", parent=self):
            return
        self.incidencias_db.delete_machine(mid)
        self.incidencias_mostrar_mapa(self.incidencias_mapas[self.incidencias_map_index])
        self.incidencias_info_maquinas(self.incidencias_info_filter_area)

    def incidencias_canvas_right_click(self, event):
        cx = self.incidencias_canvas.canvasx(event.x)
        cy = self.incidencias_canvas.canvasy(event.y)
        items = self.incidencias_canvas.find_overlapping(cx, cy, cx, cy)
        target = None
        for item_id in items:
            tags = self.incidencias_canvas.gettags(item_id)
            if "machine" in tags:
                target = item_id
                break
        if target is None:
            return
        tags = self.incidencias_canvas.gettags(target)
        mid = int(tags[1])
        self.incidencias_selected_machine = mid
        if self.incidencias_panel_mode == "machines" and self.incidencias_machines_tree is not None:
            self.incidencias_machines_tree.selection_set(str(mid))
            self.incidencias_machines_tree.see(str(mid))
        self._incidencias_resaltar_maquina(mid)

        machine = self._incidencias_get_machine_by_id(mid)
        serie = ""
        if machine:
            serie = str(machine[3]).strip()
        menu = tk.Menu(self.incidencias_canvas, tearoff=0)
        menu.add_command(
            label="Copiar N Serie",
            command=lambda: (self.clipboard_clear(), self.clipboard_append(serie), self.update()),
        )
        if self._can_write("incidencias_club"):
            menu.add_command(label="Crear incidencia", command=lambda: self._incidencias_crear_incidencia_maquina(mid))
            menu.add_command(
                label="Editar maquina",
                command=lambda: (self._incidencias_pin_ok() and self._incidencias_editar_maquina_directa(mid)),
            )
            menu.add_command(
                label="Eliminar maquina",
                command=lambda: (self._incidencias_pin_ok() and self._incidencias_eliminar_maquina_directa(mid)),
            )
        menu.post(event.x_root, event.y_root)

    def incidencias_canvas_press(self, event):
        cx = self.incidencias_canvas.canvasx(event.x)
        cy = self.incidencias_canvas.canvasy(event.y)
        if not self.incidencias_mode:
            # seleccionar area/maquina o iniciar paneo
            items = self.incidencias_canvas.find_overlapping(cx, cy, cx, cy)
            target = None
            for item_id in items:
                tags = self.incidencias_canvas.gettags(item_id)
                if "area" in tags or "machine" in tags:
                    target = item_id
                    break
            if target is None:
                self.incidencias_pan_active = True
                self.incidencias_canvas.scan_mark(event.x, event.y)
                return
            tags = self.incidencias_canvas.gettags(target)
            if "area" in tags:
                self.incidencias_selected_area = int(tags[1])
                if self.incidencias_panel_mode == "machines":
                    self.incidencias_info_maquinas(self.incidencias_selected_area)
            if "machine" in tags:
                self.incidencias_selected_machine = int(tags[1])
                if self.incidencias_panel_mode == "machines" and self.incidencias_machines_tree is not None:
                    machine_id = str(self.incidencias_selected_machine)
                    self.incidencias_machines_tree.selection_set(machine_id)
                    self.incidencias_machines_tree.see(machine_id)
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
                nombre = self._incidencias_prompt_text("Area", "Nuevo nombre del area:")
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
                nombre = self._incidencias_prompt_text("Maquina", "Nuevo nombre de la maquina:")
                if not nombre:
                    self.incidencias_mode = None
                    return
                serie = self._incidencias_prompt_text("Maquina", "Numero de serie:")
                numero = self._incidencias_prompt_text("Maquina", "Numero asignado:")
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
                elemento = self._incidencias_prompt_text("Incidencia", "Material/elemento:")
                descripcion = self._incidencias_prompt_text("Incidencia", "Describe la incidencia:")
                if elemento is None and descripcion is None:
                    self.incidencias_mode = None
                    return
                reporte_path = self._incidencias_pedir_reporte_visual()
                creador = self.incidencias_creador or {}
                self.incidencias_db.add_incident(
                    self.incidencias_current_map,
                    area_id,
                    mid,
                    elemento or "",
                    descripcion or "",
                    reporte_path=reporte_path,
                    creador_nombre=creador.get("nombre", ""),
                    creador_apellido1=creador.get("apellido1", ""),
                    creador_apellido2=creador.get("apellido2", ""),
                    creador_movil=creador.get("movil", ""),
                    creador_email=creador.get("email", ""),
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
                elemento = self._incidencias_prompt_text("Incidencia", "Material/elemento:")
                descripcion = self._incidencias_prompt_text("Incidencia", "Describe la incidencia:")
                if elemento is None and descripcion is None:
                    self.incidencias_mode = None
                    return
                reporte_path = self._incidencias_pedir_reporte_visual()
                creador = self.incidencias_creador or {}
                self.incidencias_db.add_incident(
                    self.incidencias_current_map,
                    area_id,
                    None,
                    elemento or "",
                    descripcion or "",
                    reporte_path=reporte_path,
                    creador_nombre=creador.get("nombre", ""),
                    creador_apellido1=creador.get("apellido1", ""),
                    creador_apellido2=creador.get("apellido2", ""),
                    creador_movil=creador.get("movil", ""),
                    creador_email=creador.get("email", ""),
                )
                self.incidencias_gestion_incidencias()
                self.incidencias_mode = None
                return

        self.incidencias_draw_start = (cx, cy)
        self.incidencias_draw_rect = self.incidencias_canvas.create_rectangle(
            cx, cy, cx, cy, outline="red", dash=(2, 2)
        )

    def incidencias_canvas_drag(self, event):
        if self.incidencias_pan_active and not self.incidencias_mode:
            self.incidencias_canvas.scan_dragto(event.x, event.y, gain=1)
            return
        if not self.incidencias_draw_rect:
            return
        x1, y1 = self.incidencias_draw_start
        cx = self.incidencias_canvas.canvasx(event.x)
        cy = self.incidencias_canvas.canvasy(event.y)
        self.incidencias_canvas.coords(self.incidencias_draw_rect, x1, y1, cx, cy)

    def incidencias_canvas_release(self, event):
        if self.incidencias_pan_active and not self.incidencias_mode:
            self.incidencias_pan_active = False
            return
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
            nombre = self._incidencias_prompt_text("Maquina", "Nombre de la maquina:")
            serie = self._incidencias_prompt_text("Maquina", "Numero de serie:")
            numero = self._incidencias_prompt_text("Maquina", "Numero asignado:")
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
        if not self._require_manager_access("Exportar a Excel"):
            return
        try:
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
        Abre Outlook con un borrador dirigido al manager para solicitar dia de asuntos propios.
        """
        solicitante = self._staff_select("Asuntos propios", "Quien solicita el dia?")
        if not solicitante:
            return
        destinatario = "manager.sevilla-manueldevillalobos@fitnesspark.es"
        asunto = "Peticion de Dia de Asuntos Propios"
        nombre = self._staff_display_name(solicitante)
        apellido2 = str(solicitante.get("apellido2", "")).strip()
        nombre_completo = " ".join([p for p in [nombre, apellido2] if p]).strip()
        cuerpo = f"""En Sevilla, a __ / __ / 20__
Yo, {nombre_completo}, con DNI, miembro del equipo
 del club Fitness Park Sevilla - Villalobos, presento la siguiente solicitud:

Objeto de la peticion
Solicito autorizacion para disfrutar de un dia de asuntos propios, conforme a las condiciones establecidas en mi
contrato y en el protocolo interno del club.
Fecha solicitada: ____ / ____ / 20___
Turno habitual en dicha fecha: [ ] Manana [ ] Tarde [ ] Noche
Declaracion del trabajador/a
Declaro que:
- Esta solicitud se realiza con la antelacion minima exigida por la organizacion.
- Entiendo que los dias de asuntos propios deben disfrutarse en dias laborables y siempre respetando la
cobertura del servicio.
- Acepto que la concesion del dia queda sujeta a aprobacion por parte de la Direccion del club,
 garantizando que no se vea afectada la operativa ni el equilibrio de turnos del equipo."""

        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            warn = f"""No se pudo importar win32com.client.
Cuerpo copiado al portapapeles. Envia un correo a {destinatario} con asunto:
{asunto}"""
            messagebox.showwarning("Outlook no disponible", warn)
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
            if solicitante.get("email"):
                mail.CC = solicitante.get("email")
            mail.Subject = asunto
            mail.Body = cuerpo
            mail.Display()
            messagebox.showinfo(
                "Borrador creado",
                f"Outlook se abrio con el correo dirigido a {destinatario}."
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
        Abre Outlook con un borrador para felicitar cumpleanos (1 vez por año).
        """
        if not self._require_manager_access("Enviar felicitacion"):
            return
        self._bring_to_front()
        try:
            df = obtener_cumpleanos_hoy()
        except Exception:
            df = None

        if df is None or df.empty:
            messagebox.showinfo("Cumplea\u00f1os", "No hay cumplea\u00f1os para hoy.")
            return

        def normalize(text):
            t = unicodedata.normalize("NFD", str(text or "")).upper().strip()
            return "".join(ch for ch in t if unicodedata.category(ch) != "Mn")

        colmap = {normalize(c): c for c in df.columns}
        col_email = colmap.get("CORREO ELECTRONICO") or colmap.get("EMAIL") or colmap.get("CORREO")

        if not col_email:
            messagebox.showwarning("Cumplea\u00f1os", "No se encontro columna de email.")
            return

        current_year = datetime.now().year
        emails = []
        for _, row in df.iterrows():
            email = str(row.get(col_email, "")).strip()
            if not email:
                continue
            sent_year = self.felicitaciones_enviadas.get(email)
            if sent_year == current_year:
                continue
            emails.append(email)

        if not emails:
            messagebox.showinfo("Cumplea\u00f1os", "No hay cumplea\u00f1os pendientes de enviar.")
            return

        asunto = "\u00a1FELICIDADES!! EL EQUIPO HUMANO DE FITNESS PARK VILLALOBOS TE DESEA LO MEJOR!!"
        cuerpo_html = (
            '<div style="font-family:Arial,sans-serif;font-size:14px;">'
            '<img src="cid:{{CID}}" alt="Feliz cumple\u00f1os" style="max-width:100%;">'
            "</div>"
        )
        imagen_path = get_logo_path("feliz_cumpleanos.png")

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
                mail.HTMLBody = cuerpo_html.replace('<img src="cid:{{CID}}" alt="Feliz cumpleanos" style="max-width:100%;">', "")

            mail.Display()
        except Exception as e:
            messagebox.showerror("Outlook", f"No se pudo crear el correo: {e}", parent=self)
            return

        if not messagebox.askyesno("Confirmacion", "Ha enviado la felicitacion?", parent=self):
            return

        for email in set(emails):
            self.felicitaciones_enviadas[email] = current_year
        self.guardar_felicitaciones()
        self.update_blink_states()

    def enviar_cambio_turno(self):
        """
        Abre Outlook con un borrador para solicitud de cambio de turno.
        """
        solicitante = self._staff_select("Cambio de turno", "Quien solicita el cambio?")
        if not solicitante:
            return
        aceptante = self._staff_select("Cambio de turno", "Quien acepta el cambio?")
        if not aceptante:
            return

        destinatario = "manager.sevilla-manueldevillalobos@fitnesspark.es"
        asunto = "Solicitud Cambio de turno"

        nombre1 = self._staff_display_name(solicitante)
        apellido1b = str(solicitante.get("apellido2", "")).strip()
        nombre1c = " ".join([p for p in [nombre1, apellido1b] if p]).strip()

        nombre2 = self._staff_display_name(aceptante)
        apellido2b = str(aceptante.get("apellido2", "")).strip()
        nombre2c = " ".join([p for p in [nombre2, apellido2b] if p]).strip()

        cuerpo = f"""PETICION VOLUNTARIA DE LOS SOLICITANTES QUE EN TODO MOMENTO Y BAJO SU RESPONSABILIDAD PROPORCIONARA LA COBERTURA NECESARIA PARA QUE LAS NECESIDADES DEL CLUB ESTEN CUBIERTAS. REQUIERE DE APROBACION POR PARTE DEL DIRECTOR DEL CLUB.

STAFF 1 QUE SOLICITA EL CAMBIO*
{nombre1c}

STAFF 2 QUE ACEPTA EL CAMBIO*
{nombre2c}

DESCRIBIR A CONTINUACION CAMBIO ESPECIFICANDO FECHAS Y TURNOS. RECUERDA QUE DEBEN SER CAMBIOS CERRADOS, SIN QUE QUEDE NADA EN EL AIRE:"""

        try:
            import win32com.client  # type: ignore
        except ImportError:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            warn = f"""No se pudo importar win32com.client.
Cuerpo copiado al portapapeles. Envia un correo a {destinatario} con asunto:
{asunto}"""
            messagebox.showwarning("Outlook no disponible", warn)
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
            cc_emails = []
            for persona in (solicitante, aceptante):
                email = str(persona.get("email", "")).strip()
                if email and email not in cc_emails:
                    cc_emails.append(email)
            if cc_emails:
                mail.CC = ";".join(cc_emails)
            mail.Subject = asunto
            mail.Body = cuerpo
            mail.Display()
            messagebox.showinfo(
                "Borrador creado",
                f"Outlook se abrio con el correo dirigido a {destinatario}."
            )
        except Exception as e:
            self.clipboard_clear()
            self.clipboard_append(cuerpo)
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook.\n"
                f"Cuerpo copiado al portapapeles para pegarlo manualmente.\n\nDetalle: {e}"
            )

    def abrir_staff(self):
        if not self.staff_menu:
            return
        self.update_idletasks()
        x = self.btn_staff.winfo_rootx()
        y = self.btn_staff.winfo_rooty() + self.btn_staff.winfo_height()
        self.staff_menu.post(x, y)

    def extraer_accesos(self):
        """
        Pide numero de cliente, filtra ACCESOS.csv y muestra resultado en pestaña 'Accesos Cliente'.
        """
        if not self._require_manager_access("Extraer accesos"):
            return
        if self.raw_accesos is None:
            messagebox.showwarning("Sin datos", "Primero carga datos (Actualizar datos) para tener ACCESOS.")
            return

        self._bring_to_front()
        numero = simpledialog.askstring("Número de cliente", "Introduce el número de cliente:", parent=self)
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
        Abre borrador Outlook con el Excel adjunto de la pestaña Avanza Fit.
        """
        if not self._require_manager_access("Envio Avanza Fit"):
            return
        df = self.dataframes.get("Avanza Fit")
        if df is None or df.empty:
            messagebox.showwarning("Sin datos", "No hay datos cargados en la pestaña Avanza Fit.")
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
            if not messagebox.askyesno("Confirmacion", "Ha enviado el email?", parent=self):
                return
            self.avanza_fit_envios["last_sent_date"] = datetime.now().date().isoformat()
            self.guardar_avanza_fit_envios()
            self.update_blink_states()
            messagebox.showinfo("Borrador creado", "Outlook se abrió con el archivo de Avanza Fit adjunto.")
        except Exception as e:
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook.\n\nDetalle: {e}"
            )


if __name__ == "__main__":
    try:
        app = ResamaniaApp()
        app.mainloop()
    except Exception:
        try:
            log_path = os.path.join(get_app_dir(), "app.log")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("\n=== FATAL ERROR ===\n")
                f.write(datetime.now().isoformat())
                f.write("\n")
                f.write(traceback.format_exc())
                f.write("\n")
        except Exception:
            pass
        raise
