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
from logic.wizville import procesar_wizville
from logic.accesos import procesar_accesos_dobles, procesar_accesos_descuadrados, procesar_salidas_pmr_no_autorizadas, \
    procesar_morosos_accediendo
from logic.ultimate import obtener_socios_ultimate
from logic.avanza_fit import obtener_avanza_fit
from logic.cumpleanos import obtener_cumpleanos_hoy
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
        self.dataframes = {}

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
        tk.Button(botones_frame, text="ENVIAR FELICITACIÓN", command=self.enviar_felicitacion, fg="#b30000").pack(side=tk.LEFT, padx=10)
        tk.Button(botones_frame, text="ENVÍO MARTES AVANZA FIT", command=self.enviar_avanza_fit, fg="#b30000").pack(side=tk.LEFT, padx=10)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=1, fill='both')

        self.tabs = {}
        for tab_name in [
            "Wizville", "Accesos Dobles", "Accesos Descuadrados",
            "Salidas PMR No Autorizadas", "Morosos Accediendo", "Socios Ultimate",
            "Avanza Fit", "Cumpleaños"
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
        if pin.strip() != "123":
            messagebox.showerror("Código incorrecto", "El código de seguridad no es válido.")
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

        if not emails:
            messagebox.showwarning("Sin correos", "No hay correos electrónicos en la pestaña Cumpleaños.")
            return

        try:
            import win32com.client  # tipo: ignore
        except ImportError:
            # Fallback: copiar correos al portapapeles para pegarlos en BCC manualmente.
            emails_joined = ";".join(emails)
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
            mail.BCC = ";".join(emails)
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
            messagebox.showinfo("Borrador creado", "Outlook se abrió con los correos en copia oculta.")
        except Exception as e:
            emails_joined = ";".join(emails)
            self.clipboard_clear()
            self.clipboard_append(emails_joined)
            messagebox.showerror(
                "Error con Outlook",
                f"No se pudo crear el correo en Outlook.\nSe copiaron los correos al portapapeles para pegarlos en BCC.\n\nDetalle: {e}"
            )

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
        if pin.strip() != "123":
            messagebox.showerror("Código incorrecto", "El código de seguridad no es válido.")
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
