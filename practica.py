import tkinter as tk
from tkinter import messagebox, ttk
import json
import csv
import xml.etree.ElementTree as ET
import xml.dom.minidom
import os
import subprocess
from pymongo import MongoClient as MC

# ─────────────────────────────────────────────
#  CONEXIÓN MONGODB
# ─────────────────────────────────────────────
try:
    conexion = MC("mongodb://admin:root@localhost:27017/")
    BD = conexion["BD_GrupoAlumno"]
    ColGrupo = BD["Grupo"]
    ColAlumno = BD["Alumno"]
except Exception as e:
    import sys
    print(f"Error de conexión: {e}")
    sys.exit(1)

BACKUP_DIR = r"C:\Backup_Mongo"
os.makedirs(BACKUP_DIR, exist_ok=True)

PAGE_SIZE = 10  # registros por página

# ─────────────────────────────────────────────
#  HELPER: localizar binarios de MongoDB Tools
# ─────────────────────────────────────────────
def _find_mongo_tool(name):
    """Busca mongoexport/mongoimport en rutas comunes de Windows y en el PATH."""
    import shutil
    # 1. En el PATH del sistema
    found = shutil.which(name)
    if found:
        return found
    # 2. Rutas típicas de instalación en Windows
    rutas = [
        rf"C:\Program Files\MongoDB\Tools\100\bin\{name}.exe",
        rf"C:\Program Files\MongoDB\Tools\bin\{name}.exe",
        rf"C:\Program Files\MongoDB\Server\6.0\bin\{name}.exe",
        rf"C:\Program Files\MongoDB\Server\7.0\bin\{name}.exe",
        rf"C:\Program Files\MongoDB\Server\8.0\bin\{name}.exe",
    ]
    for r in rutas:
        if os.path.exists(r):
            return r
    return None



# ══════════════════════════════════════════════
#  VENTANA PRINCIPAL  –  MENÚ
# ══════════════════════════════════════════════
class VentanaPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema BD GrupoAlumno")
        self.geometry("360x220")
        self.resizable(False, False)
        self._aplicar_estilo()
        self._construir_ui()

    def _aplicar_estilo(self):
        self.configure(bg="#0f172a")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TButton",
            background="#1e3a5f", foreground="#e2e8f0",
            font=("Consolas", 10, "bold"), relief="flat",
            borderwidth=0, padding=8)
        style.map("TButton",
            background=[("active", "#2563eb")],
            foreground=[("active", "#ffffff")])

    def _construir_ui(self):
        tk.Label(self, text="▸ SISTEMA BD GRUPOALUMNO",
                 bg="#0f172a", fg="#38bdf8",
                 font=("Consolas", 13, "bold")).pack(pady=(22, 4))
        tk.Label(self, text="MongoDB  ·  Tkinter",
                 bg="#0f172a", fg="#475569",
                 font=("Consolas", 9)).pack(pady=(0, 18))
        frame = tk.Frame(self, bg="#0f172a")
        frame.pack()
        ttk.Button(frame, text="⊞  Administrar Grupos",
                   command=lambda: VentanaGrupo(self), width=28).grid(row=0, column=0, padx=8, pady=6)
        ttk.Button(frame, text="⊞  Administrar Alumnos",
                   command=lambda: VentanaAlumno(self), width=28).grid(row=1, column=0, padx=8, pady=6)


# ══════════════════════════════════════════════
#  BASE PARA VENTANAS CRUD
# ══════════════════════════════════════════════
class VentanaBase(tk.Toplevel):
    COLOR_BG      = "#0f172a"
    COLOR_PANEL   = "#1e293b"
    COLOR_BORDE   = "#334155"
    COLOR_ACENTO  = "#2563eb"
    COLOR_TEXTO   = "#e2e8f0"
    COLOR_MUTED   = "#94a3b8"
    COLOR_ENTRY   = "#0f172a"
    COLOR_PELIGRO = "#dc2626"
    COLOR_OK      = "#16a34a"
    FUENTE_LABEL  = ("Consolas", 10)
    FUENTE_TITULO = ("Consolas", 13, "bold")
    FUENTE_BTN    = ("Consolas", 9, "bold")
    FUENTE_TABLA  = ("Consolas", 9)

    def __init__(self, parent, titulo, ancho, alto):
        super().__init__(parent)
        self.title(titulo)
        self.geometry(f"{ancho}x{alto}")
        self.resizable(True, True)
        self.configure(bg=self.COLOR_BG)
        self._entries = {}
        self._page = 0
        self._datos_filtrados = []
        self._construir()

    def _label(self, parent, texto, **kw):
        return tk.Label(parent, text=texto,
                        bg=kw.pop("bg", self.COLOR_PANEL),
                        fg=kw.pop("fg", self.COLOR_MUTED),
                        font=self.FUENTE_LABEL, **kw)

    def _entry(self, parent, key, **kw):
        e = tk.Entry(parent,
                     bg=self.COLOR_ENTRY, fg=self.COLOR_TEXTO,
                     insertbackground=self.COLOR_TEXTO,
                     relief="flat", font=self.FUENTE_LABEL,
                     highlightthickness=1,
                     highlightcolor=self.COLOR_ACENTO,
                     highlightbackground=self.COLOR_BORDE, **kw)
        self._entries[key] = e
        return e

    def _boton(self, parent, texto, comando, color=None, **kw):
        c = color or self.COLOR_PANEL
        b = tk.Button(parent, text=texto, command=comando,
                      bg=c, fg=self.COLOR_TEXTO,
                      font=self.FUENTE_BTN,
                      relief="flat", cursor="hand2",
                      activebackground=self.COLOR_ACENTO,
                      activeforeground="#fff",
                      pady=5, **kw)
        b.bind("<Enter>", lambda e: b.config(bg=self.COLOR_ACENTO))
        b.bind("<Leave>", lambda e: b.config(bg=c))
        return b

    def _get(self, key):
        return self._entries[key].get().strip()

    def _set(self, key, valor):
        self._entries[key].delete(0, tk.END)
        self._entries[key].insert(0, valor)

    def _limpiar_entries(self):
        for e in self._entries.values():
            e.delete(0, tk.END)

    def ok(self, msg):   messagebox.showinfo("✔ Éxito", msg, parent=self)
    def err(self, msg):  messagebox.showerror("✘ Error", msg, parent=self)
    def conf(self, msg): return messagebox.askyesno("¿Confirmar?", msg, parent=self)

    # ── Tabla con scroll ──────────────────────
    def _crear_tabla(self, parent, columnas, row, col, rowspan=1, colspan=3):
        frame = tk.Frame(parent, bg=self.COLOR_PANEL,
                         highlightthickness=1,
                         highlightbackground=self.COLOR_BORDE)
        frame.grid(row=row, column=col, columnspan=colspan,
                   rowspan=rowspan, padx=16, pady=(4,0), sticky="nsew")
        parent.grid_rowconfigure(row, weight=1)

        style = ttk.Style()
        style.configure("Dark.Treeview",
                        background=self.COLOR_PANEL,
                        foreground=self.COLOR_TEXTO,
                        fieldbackground=self.COLOR_PANEL,
                        rowheight=22,
                        font=self.FUENTE_TABLA)
        style.configure("Dark.Treeview.Heading",
                        background=self.COLOR_BORDE,
                        foreground="#38bdf8",
                        font=("Consolas", 9, "bold"))
        style.map("Dark.Treeview",
                  background=[("selected", self.COLOR_ACENTO)],
                  foreground=[("selected", "#fff")])

        tree = ttk.Treeview(frame, columns=columnas, show="headings",
                            style="Dark.Treeview", selectmode="browse")
        for col_name in columnas:
            tree.heading(col_name, text=col_name)
            tree.column(col_name, width=120, anchor="w")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        return tree

    def _crear_paginacion(self, parent, row, col, colspan=3):
        pf = tk.Frame(parent, bg=self.COLOR_BG)
        pf.grid(row=row, column=col, columnspan=colspan, pady=(2,4), sticky="ew")
        for i in range(5): pf.grid_columnconfigure(i, weight=1)

        self._lbl_pagina = tk.Label(pf, text="Página 0 / 0",
                                    bg=self.COLOR_BG, fg=self.COLOR_MUTED,
                                    font=("Consolas", 8))
        self._lbl_pagina.grid(row=0, column=2)

        self._btn_ant = self._boton(pf, "◀ Ant", self._pag_anterior, "#1e3a5f")
        self._btn_ant.grid(row=0, column=1, padx=4, sticky="ew")
        self._btn_sig = self._boton(pf, "Sig ▶", self._pag_siguiente, "#1e3a5f")
        self._btn_sig.grid(row=0, column=3, padx=4, sticky="ew")

        self._lbl_total = tk.Label(pf, text="Total: 0",
                                   bg=self.COLOR_BG, fg=self.COLOR_MUTED,
                                   font=("Consolas", 8))
        self._lbl_total.grid(row=0, column=4, sticky="e", padx=8)

    def _pag_anterior(self):
        if self._page > 0:
            self._page -= 1
            self._actualizar_tabla()

    def _pag_siguiente(self):
        total_pages = max(1, (len(self._datos_filtrados) + PAGE_SIZE - 1) // PAGE_SIZE)
        if self._page < total_pages - 1:
            self._page += 1
            self._actualizar_tabla()

    def _actualizar_paginacion_labels(self):
        total = len(self._datos_filtrados)
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        self._lbl_pagina.config(text=f"Página {self._page + 1} / {total_pages}")
        self._lbl_total.config(text=f"Total: {total}")


# ══════════════════════════════════════════════
#  POPUP AGREGAR / EDITAR  (base)
# ══════════════════════════════════════════════
class PopupBase(tk.Toplevel):
    COLOR_BG    = "#0f172a"
    COLOR_PANEL = "#1e293b"
    COLOR_BORDE = "#334155"
    COLOR_ACENTO= "#2563eb"
    COLOR_TEXTO = "#e2e8f0"
    COLOR_MUTED = "#94a3b8"
    COLOR_ENTRY = "#0f172a"
    FUENTE      = ("Consolas", 10)
    FUENTE_BOLD = ("Consolas", 11, "bold")

    def __init__(self, parent, titulo):
        super().__init__(parent)
        self.title(titulo)
        self.resizable(False, False)
        self.configure(bg=self.COLOR_BG)
        self.grab_set()
        self._entries = {}
        self.resultado = None
        self._construir()
        self.transient(parent)

    def _label(self, parent, texto):
        return tk.Label(parent, text=texto,
                        bg=self.COLOR_PANEL, fg=self.COLOR_MUTED,
                        font=self.FUENTE)

    def _entry(self, parent, key, **kw):
        e = tk.Entry(parent,
                     bg=self.COLOR_ENTRY, fg=self.COLOR_TEXTO,
                     insertbackground=self.COLOR_TEXTO,
                     relief="flat", font=self.FUENTE,
                     highlightthickness=1,
                     highlightcolor=self.COLOR_ACENTO,
                     highlightbackground=self.COLOR_BORDE, **kw)
        self._entries[key] = e
        return e

    def _get(self, key): return self._entries[key].get().strip()
    def _set(self, key, v):
        widget = self._entries[key]
        if isinstance(widget, ttk.Combobox):
            widget.set(v)
        else:
            widget.delete(0, tk.END)
            widget.insert(0, v)

    def err(self, msg): messagebox.showerror("✘ Error", msg, parent=self)


# ══════════════════════════════════════════════
#  POPUP GRUPO
# ══════════════════════════════════════════════
class PopupGrupo(PopupBase):
    def __init__(self, parent, datos=None):
        self._datos_iniciales = datos
        super().__init__(parent, "Nuevo Grupo" if datos is None else "Editar Grupo")

    def _construir(self):
        panel = tk.Frame(self, bg=self.COLOR_PANEL,
                         highlightthickness=1,
                         highlightbackground=self.COLOR_BORDE)
        panel.pack(padx=20, pady=16, fill="both")
        panel.grid_columnconfigure(1, weight=1)

        tk.Label(panel, text="▸ DATOS DEL GRUPO",
                 bg=self.COLOR_PANEL, fg="#38bdf8",
                 font=self.FUENTE_BOLD).grid(
                 row=0, column=0, columnspan=2, padx=12, pady=(12,8), sticky="w")

        self._label(panel, "Clave Grupo:").grid(row=1, column=0, padx=12, pady=8, sticky="w")
        self._entry(panel, "cve", width=25).grid(row=1, column=1, padx=12, pady=8, sticky="ew")

        self._label(panel, "Nombre Grupo:").grid(row=2, column=0, padx=12, pady=8, sticky="w")
        self._entry(panel, "nom", width=30).grid(row=2, column=1, padx=12, pady=8, sticky="ew")

        if self._datos_iniciales:
            self._set("cve", self._datos_iniciales.get("cveGru", ""))
            self._set("nom", self._datos_iniciales.get("nomGru", ""))
            self._entries["cve"].config(
                state="disabled",
                disabledbackground="#0f172a",
                disabledforeground="#e2e8f0"
            )

        bf = tk.Frame(self, bg=self.COLOR_BG)
        bf.pack(padx=20, pady=(0, 16), fill="x")
        bf.grid_columnconfigure(0, weight=1)
        bf.grid_columnconfigure(1, weight=1)

        tk.Button(bf, text="✔ Guardar", command=self._guardar,
                  bg="#15803d", fg=self.COLOR_TEXTO, font=("Consolas", 9, "bold"),
                  relief="flat", cursor="hand2", pady=6).grid(row=0, column=0, padx=4, sticky="ew")
        tk.Button(bf, text="✕ Cancelar", command=self.destroy,
                  bg="#7f1d1d", fg=self.COLOR_TEXTO, font=("Consolas", 9, "bold"),
                  relief="flat", cursor="hand2", pady=6).grid(row=0, column=1, padx=4, sticky="ew")

    def _guardar(self):
        # En modo edición la clave está disabled, leerla desde datos_iniciales
        if self._datos_iniciales is not None:
            cve = str(self._datos_iniciales.get("cveGru", ""))
        else:
            cve = self._get("cve")
        nom = self._get("nom")
        if not cve or not nom:
            return self.err("Clave y Nombre son obligatorios.")

        if self._datos_iniciales is None:
            # Agregar: validar duplicado
            if ColGrupo.find_one({"cveGru": cve}):
                return self.err(f"Ya existe un grupo con clave '{cve}'.")
            ColGrupo.insert_one({"cveGru": cve, "nomGru": nom})
            messagebox.showinfo("✔ Éxito", "Grupo agregado correctamente.", parent=self)
        else:
            # Editar: actualizar nombre
            ColGrupo.update_one({"cveGru": cve}, {"$set": {"nomGru": nom}})
            messagebox.showinfo("✔ Éxito", "Grupo modificado correctamente.", parent=self)

        self.resultado = True
        self.destroy()


# ══════════════════════════════════════════════
#  POPUP ALUMNO
# ══════════════════════════════════════════════
class PopupAlumno(PopupBase):
    def __init__(self, parent, datos=None):
        self._datos_iniciales = datos
        super().__init__(parent, "Nuevo Alumno" if datos is None else "Editar Alumno")

    def _construir(self):
        panel = tk.Frame(self, bg=self.COLOR_PANEL,
                         highlightthickness=1,
                         highlightbackground=self.COLOR_BORDE)
        panel.pack(padx=20, pady=16, fill="both")
        panel.grid_columnconfigure(1, weight=1)

        tk.Label(panel, text="▸ DATOS DEL ALUMNO",
                 bg=self.COLOR_PANEL, fg="#38bdf8",
                 font=self.FUENTE_BOLD).grid(
                 row=0, column=0, columnspan=2, padx=12, pady=(12,8), sticky="w")

        # Campos texto simples
        campos_texto = [("Clave Alumno:", "cve"), ("Nombre Alumno:", "nom"), ("Edad:", "eda")]
        for i, (lbl, key) in enumerate(campos_texto, start=1):
            self._label(panel, lbl).grid(row=i, column=0, padx=12, pady=8, sticky="w")
            self._entry(panel, key, width=28).grid(row=i, column=1, padx=12, pady=8, sticky="ew")

        # Campo Grupo como Combobox
        self._label(panel, "Grupo:").grid(row=4, column=0, padx=12, pady=8, sticky="w")
        grupos = [d["cveGru"] for d in ColGrupo.find().sort("cveGru", 1)]
        self._var_cveGru = tk.StringVar()
        self._combo_gru = ttk.Combobox(panel, textvariable=self._var_cveGru,
                                        values=grupos, state="readonly", width=26,
                                        font=self.FUENTE)
        self._combo_gru.grid(row=4, column=1, padx=12, pady=8, sticky="ew")
        # Registrar en _entries para que _get("cveGru") funcione
        self._entries["cveGru"] = self._combo_gru

        if self._datos_iniciales:
            self._cve_readonly = self._datos_iniciales.get("cveAlu", "")
            self._set("cve",   self._cve_readonly)
            self._set("nom",   self._datos_iniciales.get("nomAlu", ""))
            self._set("eda",   str(self._datos_iniciales.get("edaAlu", "")))
            self._var_cveGru.set(self._datos_iniciales.get("cveGru", ""))
            self._entries["cve"].config(state="readonly")
        else:
            self._cve_readonly = ""
            if grupos:
                self._combo_gru.current(0)

        bf = tk.Frame(self, bg=self.COLOR_BG)
        bf.pack(padx=20, pady=(0, 16), fill="x")
        bf.grid_columnconfigure(0, weight=1)
        bf.grid_columnconfigure(1, weight=1)

        tk.Button(bf, text="✔ Guardar", command=self._guardar,
                  bg="#15803d", fg=self.COLOR_TEXTO, font=("Consolas", 9, "bold"),
                  relief="flat", cursor="hand2", pady=6).grid(row=0, column=0, padx=4, sticky="ew")
        tk.Button(bf, text="✕ Cancelar", command=self.destroy,
                  bg="#7f1d1d", fg=self.COLOR_TEXTO, font=("Consolas", 9, "bold"),
                  relief="flat", cursor="hand2", pady=6).grid(row=0, column=1, padx=4, sticky="ew")

    def _guardar(self):
        # Si estamos editando, la clave viene de _cve_readonly (Entry disabled devuelve "")
        if self._datos_iniciales is not None:
            cve = self._cve_readonly
        else:
            cve = self._get("cve")
        nom    = self._get("nom")
        eda    = self._get("eda")
        cveGru = self._var_cveGru.get().strip()

        if not all([cve, nom, eda, cveGru]):
            return self.err("Todos los campos son obligatorios.")
        if not eda.isdigit():
            return self.err("La edad debe ser un número entero.")
        if not ColGrupo.find_one({"cveGru": cveGru}):
            return self.err(f"El grupo '{cveGru}' no existe.")

        if self._datos_iniciales is None:
            if ColAlumno.find_one({"cveAlu": cve}):
                return self.err(f"Ya existe un alumno con clave '{cve}'.")
            ColAlumno.insert_one({"cveAlu": cve, "nomAlu": nom,
                                   "edaAlu": int(eda), "cveGru": cveGru})
            messagebox.showinfo("✔ Éxito", "Alumno agregado correctamente.", parent=self)
        else:
            ColAlumno.update_one({"cveAlu": cve},
                {"$set": {"nomAlu": nom, "edaAlu": int(eda), "cveGru": cveGru}})
            messagebox.showinfo("✔ Éxito", "Alumno modificado correctamente.", parent=self)

        self.resultado = True
        self.destroy()


# ══════════════════════════════════════════════
#  VENTANA GRUPOS
# ══════════════════════════════════════════════
class VentanaGrupo(VentanaBase):
    def __init__(self, parent):
        super().__init__(parent, "Administración de Grupos", 680, 620)

    def _construir(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Título
        tk.Label(self, text="▸ ADMINISTRACIÓN DE GRUPOS",
                 bg=self.COLOR_BG, fg="#38bdf8",
                 font=self.FUENTE_TITULO).grid(
                 row=0, column=0, pady=(16,4), padx=16, sticky="w")

        # ── Barra de búsqueda / filtro ────────
        bf = tk.Frame(self, bg=self.COLOR_PANEL,
                      highlightthickness=1,
                      highlightbackground=self.COLOR_BORDE)
        bf.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
        bf.grid_columnconfigure(1, weight=1)

        tk.Label(bf, text="🔍 Filtrar:", bg=self.COLOR_PANEL,
                 fg=self.COLOR_MUTED, font=self.FUENTE_LABEL).grid(
                 row=0, column=0, padx=10, pady=8, sticky="w")

        self._var_filtro = tk.StringVar()
        self._var_filtro.trace_add("write", lambda *a: self._aplicar_filtro())
        tk.Entry(bf, textvariable=self._var_filtro,
                 bg=self.COLOR_ENTRY, fg=self.COLOR_TEXTO,
                 insertbackground=self.COLOR_TEXTO,
                 relief="flat", font=self.FUENTE_LABEL,
                 highlightthickness=1,
                 highlightcolor=self.COLOR_ACENTO,
                 highlightbackground=self.COLOR_BORDE).grid(
                 row=0, column=1, padx=8, pady=8, sticky="ew")

        tk.Label(bf, text="(clave o nombre)", bg=self.COLOR_PANEL,
                 fg=self.COLOR_MUTED, font=("Consolas", 8)).grid(
                 row=0, column=2, padx=8)

        # ── Botones de acción ─────────────────
        af = tk.Frame(self, bg=self.COLOR_BG)
        af.grid(row=2, column=0, padx=16, pady=4, sticky="ew")
        for i in range(5): af.grid_columnconfigure(i, weight=1)

        self._boton(af, "＋ Agregar",  self._agregar,  "#15803d").grid(row=0, column=0, padx=3, pady=3, sticky="ew")
        self._boton(af, "✎ Editar",    self._editar,   "#1d4ed8").grid(row=0, column=1, padx=3, pady=3, sticky="ew")
        self._boton(af, "✕ Eliminar",  self._eliminar, "#b91c1c").grid(row=0, column=2, padx=3, pady=3, sticky="ew")
        self._boton(af, "↺ Refrescar", self._refrescar,"#0e7490").grid(row=0, column=3, padx=3, pady=3, sticky="ew")

        # ── Tabla ─────────────────────────────
        self._tree = self._crear_tabla(self, ["Clave", "Nombre Grupo"], row=3, col=0, colspan=1)
        self._tree.bind("<Double-1>", lambda e: self._editar())
        self._crear_paginacion(self, row=4, col=0, colspan=1)

        # ── Exportar / Importar ───────────────
        ef = tk.Frame(self, bg=self.COLOR_BG)
        ef.grid(row=5, column=0, padx=16, pady=4, sticky="ew")
        for i in range(6): ef.grid_columnconfigure(i, weight=1)

        tk.Label(ef, text="Exportar:", bg=self.COLOR_BG, fg=self.COLOR_MUTED,
                 font=("Consolas", 8)).grid(row=0, column=0, sticky="w")
        self._boton(ef, "CSV",  self._exp_csv ).grid(row=0, column=1, padx=2, sticky="ew")
        self._boton(ef, "JSON", self._exp_json).grid(row=0, column=2, padx=2, sticky="ew")
        self._boton(ef, "XML",  self._exp_xml ).grid(row=0, column=3, padx=2, sticky="ew")

        tk.Label(ef, text="Importar:", bg=self.COLOR_BG, fg=self.COLOR_MUTED,
                 font=("Consolas", 8)).grid(row=1, column=0, sticky="w")
        self._boton(ef, "CSV",  self._imp_csv ).grid(row=1, column=1, padx=2, sticky="ew")
        self._boton(ef, "JSON", self._imp_json).grid(row=1, column=2, padx=2, sticky="ew")
        self._boton(ef, "XML",  self._imp_xml ).grid(row=1, column=3, padx=2, sticky="ew")

        # ── Backup / Peligro ──────────────────
        df = tk.Frame(self, bg=self.COLOR_BG)
        df.grid(row=6, column=0, padx=16, pady=(4,12), sticky="ew")
        for i in range(3): df.grid_columnconfigure(i, weight=1)

        self._boton(df, "⬡ Backup",              self._backup,          "#7c3aed").grid(row=0, column=0, padx=3, pady=2, sticky="ew")
        self._boton(df, "⚠ Eliminar todos",       self._eliminar_todos,  self.COLOR_PELIGRO).grid(row=0, column=1, padx=3, pady=2, sticky="ew")
        self._boton(df, "⟳ Restaurar todos",      self._restaurar_todos, "#b45309").grid(row=0, column=2, padx=3, pady=2, sticky="ew")

        self._refrescar()

    # ── Datos y tabla ─────────────────────────
    def _cargar_datos(self):
        return [{"cveGru": d["cveGru"], "nomGru": d["nomGru"]}
                for d in ColGrupo.find().sort("cveGru", 1)]

    def _refrescar(self):
        self._todos = self._cargar_datos()
        self._aplicar_filtro()

    def _aplicar_filtro(self):
        filtro = self._var_filtro.get().strip().lower()
        if filtro:
            self._datos_filtrados = [
                d for d in self._todos
                if filtro in d["cveGru"].lower() or filtro in d["nomGru"].lower()
            ]
        else:
            self._datos_filtrados = list(self._todos)
        self._page = 0
        self._actualizar_tabla()

    def _actualizar_tabla(self):
        for row in self._tree.get_children():
            self._tree.delete(row)
        inicio = self._page * PAGE_SIZE
        fin    = inicio + PAGE_SIZE
        for d in self._datos_filtrados[inicio:fin]:
            self._tree.insert("", "end", values=(d["cveGru"], d["nomGru"]))
        self._actualizar_paginacion_labels()

    def _fila_seleccionada(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Seleccione un grupo de la tabla.", parent=self)
            return None
        return self._tree.item(sel[0])["values"]

    # ── CRUD ──────────────────────────────────
    def _agregar(self):
        p = PopupGrupo(self)
        self.wait_window(p)
        if p.resultado:
            self._refrescar()

    def _editar(self):
        vals = self._fila_seleccionada()
        if not vals: return
        cve = str(vals[0])
        doc = ColGrupo.find_one({"cveGru": cve})
        if not doc:
            try: doc = ColGrupo.find_one({"cveGru": int(cve)})
            except Exception: pass
        if not doc: return self.err("No se encontró el documento.")
        p = PopupGrupo(self, datos=doc)
        self.wait_window(p)
        if p.resultado:
            self._refrescar()

    def _eliminar(self):
        vals = self._fila_seleccionada()
        if not vals: return
        cve = str(vals[0])
        n_alumnos = ColAlumno.count_documents({"cveGru": cve})
        if n_alumnos > 0:
            aviso = (f"El grupo '{cve}' tiene {n_alumnos} alumno(s) asignado(s).\n"
                     f"¿Eliminar el grupo Y todos sus alumnos? (eliminación en cascada)")
        else:
            aviso = f"¿Eliminar el grupo '{cve}'?"
        if self.conf(aviso):
            ColAlumno.delete_many({"cveGru": cve})
            res = ColGrupo.delete_one({"cveGru": cve})
            if res.deleted_count == 0:
                try: ColGrupo.delete_one({"cveGru": int(cve)})
                except Exception: pass
            detalle = f"Grupo '{cve}' eliminado." + (f"\n{n_alumnos} alumno(s) eliminado(s) en cascada." if n_alumnos else "")
            self.ok(detalle)
            self._refrescar()

    def _eliminar_todos(self):
        n_grupos  = ColGrupo.count_documents({})
        n_alumnos = ColAlumno.count_documents({})
        if n_grupos == 0:
            return self.err("No hay grupos registrados.")
        aviso = (f"¿Eliminar TODOS los grupos ({n_grupos}) "
                 f"y TODOS los alumnos ({n_alumnos}) en cascada?\n"
                 "Esta acción NO se puede deshacer.")
        if self.conf(aviso):
            ColAlumno.delete_many({})
            ColGrupo.delete_many({})
            self.ok(f"Eliminados {n_grupos} grupo(s) y {n_alumnos} alumno(s) en cascada.")
            self._refrescar()

    # ── Exportar ──────────────────────────────
    def _datos(self):
        return self._cargar_datos()

    def _exp_csv(self):
        datos = self._datos()
        if not datos: return self.err("No hay grupos para exportar.")
        ruta = os.path.join(BACKUP_DIR, "Grupo.csv")
        with open(ruta, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["cveGru","nomGru"])
            w.writeheader(); w.writerows(datos)
        self.ok(f"CSV exportado:\n{ruta}")

    def _exp_json(self):
        datos = self._datos()
        if not datos: return self.err("No hay grupos para exportar.")
        ruta = os.path.join(BACKUP_DIR, "Grupo.json")
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
        self.ok(f"JSON exportado:\n{ruta}")

    def _exp_xml(self):
        datos = self._datos()
        if not datos: return self.err("No hay grupos para exportar.")
        root = ET.Element("Grupos")
        for d in datos:
            g = ET.SubElement(root, "Grupo")
            ET.SubElement(g, "cveGru").text = d["cveGru"]
            ET.SubElement(g, "nomGru").text = d["nomGru"]
        ruta = os.path.join(BACKUP_DIR, "Grupo.xml")
        dom = xml.dom.minidom.parseString(ET.tostring(root, encoding="unicode"))
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(dom.toprettyxml(indent="  "))
        self.ok(f"XML exportado:\n{ruta}")

    # ── Importar ──────────────────────────────
    def _imp_csv(self):
        ruta = os.path.join(BACKUP_DIR, "Grupo.csv")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el archivo:\n{ruta}\n\nPrimero exporte los datos para generar el archivo.")
        insertados = omitidos = 0
        try:
            with open(ruta, encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    if not ColGrupo.find_one({"cveGru": row["cveGru"]}):
                        ColGrupo.insert_one({"cveGru": row["cveGru"], "nomGru": row["nomGru"]})
                        insertados += 1
                    else:
                        omitidos += 1
        except Exception as e:
            return self.err(f"Error al leer el archivo CSV:\n{e}")
        msg = f"CSV importado.\n✔ Insertados: {insertados}"
        if omitidos: msg += f"\n⚠ Omitidos (ya existían): {omitidos}"
        self.ok(msg)
        self._refrescar()

    def _imp_json(self):
        ruta = os.path.join(BACKUP_DIR, "Grupo.json")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el archivo:\n{ruta}\n\nPrimero exporte los datos para generar el archivo.")
        insertados = omitidos = 0
        try:
            with open(ruta, encoding="utf-8") as f:
                datos = json.load(f)
        except Exception as e:
            return self.err(f"Error al leer el archivo JSON:\n{e}")
        for d in datos:
            if not ColGrupo.find_one({"cveGru": d["cveGru"]}):
                ColGrupo.insert_one({"cveGru": d["cveGru"], "nomGru": d["nomGru"]})
                insertados += 1
            else:
                omitidos += 1
        msg = f"JSON importado.\n✔ Insertados: {insertados}"
        if omitidos: msg += f"\n⚠ Omitidos (ya existían): {omitidos}"
        self.ok(msg)
        self._refrescar()

    def _imp_xml(self):
        ruta = os.path.join(BACKUP_DIR, "Grupo.xml")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el archivo:\n{ruta}\n\nPrimero exporte los datos para generar el archivo.")
        insertados = omitidos = 0
        try:
            tree = ET.parse(ruta)
        except Exception as e:
            return self.err(f"Error al leer el archivo XML:\n{e}")
        for g in tree.getroot().findall("Grupo"):
            cve = g.find("cveGru").text
            nom = g.find("nomGru").text
            if not ColGrupo.find_one({"cveGru": cve}):
                ColGrupo.insert_one({"cveGru": cve, "nomGru": nom})
                insertados += 1
            else:
                omitidos += 1
        msg = f"XML importado.\n✔ Insertados: {insertados}"
        if omitidos: msg += f"\n⚠ Omitidos (ya existían): {omitidos}"
        self.ok(msg)
        self._refrescar()

    # ── Backup / Restaurar ────────────────────
    def _backup(self):
        exe = _find_mongo_tool("mongoexport")
        if not exe:
            return self.err(
                "No se encontró 'mongoexport'.\n"
                "Instala MongoDB Database Tools:\n"
                "https://www.mongodb.com/try/download/database-tools\n"
                "y agrega la carpeta bin al PATH del sistema."
            )
        try:
            salida = os.path.join(BACKUP_DIR, "Grupo_backup.json")
            cmd = [exe,
                   "--host=localhost", "--port=27017",
                   "--username=admin", "--password=root",
                   "--authenticationDatabase=admin",
                   "--db=BD_GrupoAlumno", "--collection=Grupo",
                   f"--out={salida}"]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok(f"Backup ejecutado.\n{salida}")
            else:
                self.err(f"Error en backup:\n{result.stderr}")
        except Exception as e:
            self.err(str(e))

    def _restaurar_todos(self):
        ruta = os.path.join(BACKUP_DIR, "Grupo_backup.json")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el backup:\n{ruta}")
        exe = _find_mongo_tool("mongoimport")
        if not exe:
            return self.err("No se encontró 'mongoimport'. Instala MongoDB Database Tools.")
        if not self.conf("¿Restaurar TODOS los grupos desde el backup?"):
            return
        try:
            cmd = [exe,
                   "--host=localhost", "--port=27017",
                   "--username=admin", "--password=root",
                   "--authenticationDatabase=admin",
                   "--db=BD_GrupoAlumno", "--collection=Grupo",
                   f"--file={ruta}"]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok("Grupos restaurados correctamente.")
                self._refrescar()
            else:
                self.err(f"Error al restaurar:\n{result.stderr}")
        except Exception as e:
            self.err(str(e))


# ══════════════════════════════════════════════
#  VENTANA ALUMNOS
# ══════════════════════════════════════════════
class VentanaAlumno(VentanaBase):
    def __init__(self, parent):
        super().__init__(parent, "Administración de Alumnos", 760, 640)

    def _construir(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        tk.Label(self, text="▸ ADMINISTRACIÓN DE ALUMNOS",
                 bg=self.COLOR_BG, fg="#38bdf8",
                 font=self.FUENTE_TITULO).grid(
                 row=0, column=0, pady=(16,4), padx=16, sticky="w")

        # ── Barra de búsqueda / filtro ────────
        bf = tk.Frame(self, bg=self.COLOR_PANEL,
                      highlightthickness=1,
                      highlightbackground=self.COLOR_BORDE)
        bf.grid(row=1, column=0, padx=16, pady=6, sticky="ew")
        bf.grid_columnconfigure(1, weight=1)
        bf.grid_columnconfigure(3, weight=1)

        tk.Label(bf, text="🔍 Filtrar:", bg=self.COLOR_PANEL,
                 fg=self.COLOR_MUTED, font=self.FUENTE_LABEL).grid(
                 row=0, column=0, padx=10, pady=8, sticky="w")

        # Inicializar _todos antes de registrar los traces para evitar
        # AttributeError si el trace se dispara antes de _refrescar()
        self._todos = []

        self._var_filtro = tk.StringVar()
        self._var_filtro.trace_add("write", lambda *a: self._aplicar_filtro())
        tk.Entry(bf, textvariable=self._var_filtro,
                 bg=self.COLOR_ENTRY, fg=self.COLOR_TEXTO,
                 insertbackground=self.COLOR_TEXTO,
                 relief="flat", font=self.FUENTE_LABEL,
                 highlightthickness=1,
                 highlightcolor=self.COLOR_ACENTO,
                 highlightbackground=self.COLOR_BORDE).grid(
                 row=0, column=1, padx=8, pady=8, sticky="ew")

        tk.Label(bf, text="Grupo:", bg=self.COLOR_PANEL,
                 fg=self.COLOR_MUTED, font=self.FUENTE_LABEL).grid(
                 row=0, column=2, padx=(12,4), pady=8, sticky="w")

        self._var_grupo = tk.StringVar()
        self._var_grupo.trace_add("write", lambda *a: self._aplicar_filtro())
        grupos = ["(todos)"] + [d["cveGru"] for d in ColGrupo.find().sort("cveGru", 1)]
        self._combo_grupo = ttk.Combobox(bf, textvariable=self._var_grupo,
                                          values=grupos, state="readonly", width=14)
        self._combo_grupo.current(0)
        self._combo_grupo.grid(row=0, column=3, padx=8, pady=8, sticky="ew")

        # ── Botones de acción ─────────────────
        af = tk.Frame(self, bg=self.COLOR_BG)
        af.grid(row=2, column=0, padx=16, pady=4, sticky="ew")
        for i in range(5): af.grid_columnconfigure(i, weight=1)

        self._boton(af, "＋ Agregar",  self._agregar,  "#15803d").grid(row=0, column=0, padx=3, pady=3, sticky="ew")
        self._boton(af, "✎ Editar",    self._editar,   "#1d4ed8").grid(row=0, column=1, padx=3, pady=3, sticky="ew")
        self._boton(af, "✕ Eliminar",  self._eliminar, "#b91c1c").grid(row=0, column=2, padx=3, pady=3, sticky="ew")
        self._boton(af, "↺ Refrescar", self._refrescar,"#0e7490").grid(row=0, column=3, padx=3, pady=3, sticky="ew")

        # ── Tabla ─────────────────────────────
        self._tree = self._crear_tabla(
            self, ["Clave", "Nombre Alumno", "Edad", "Clave Grupo"],
            row=3, col=0, colspan=1)
        self._tree.bind("<Double-1>", lambda e: self._editar())
        self._crear_paginacion(self, row=4, col=0, colspan=1)

        # ── Exportar / Importar ───────────────
        ef = tk.Frame(self, bg=self.COLOR_BG)
        ef.grid(row=5, column=0, padx=16, pady=4, sticky="ew")
        for i in range(6): ef.grid_columnconfigure(i, weight=1)

        tk.Label(ef, text="Exportar:", bg=self.COLOR_BG, fg=self.COLOR_MUTED,
                 font=("Consolas", 8)).grid(row=0, column=0, sticky="w")
        self._boton(ef, "CSV",  self._exp_csv ).grid(row=0, column=1, padx=2, sticky="ew")
        self._boton(ef, "JSON", self._exp_json).grid(row=0, column=2, padx=2, sticky="ew")
        self._boton(ef, "XML",  self._exp_xml ).grid(row=0, column=3, padx=2, sticky="ew")

        tk.Label(ef, text="Importar:", bg=self.COLOR_BG, fg=self.COLOR_MUTED,
                 font=("Consolas", 8)).grid(row=1, column=0, sticky="w")
        self._boton(ef, "CSV",  self._imp_csv ).grid(row=1, column=1, padx=2, sticky="ew")
        self._boton(ef, "JSON", self._imp_json).grid(row=1, column=2, padx=2, sticky="ew")
        self._boton(ef, "XML",  self._imp_xml ).grid(row=1, column=3, padx=2, sticky="ew")

        # ── Backup / Peligro ──────────────────
        df = tk.Frame(self, bg=self.COLOR_BG)
        df.grid(row=6, column=0, padx=16, pady=(4,12), sticky="ew")
        for i in range(3): df.grid_columnconfigure(i, weight=1)

        self._boton(df, "⬡ Backup",              self._backup,         "#7c3aed").grid(row=0, column=0, padx=3, pady=2, sticky="ew")
        self._boton(df, "⚠ Eliminar todos",       self._eliminar_todos, self.COLOR_PELIGRO).grid(row=0, column=1, padx=3, pady=2, sticky="ew")
        self._boton(df, "⟳ Restaurar todos",      self._restaurar_todos,"#b45309").grid(row=0, column=2, padx=3, pady=2, sticky="ew")

        self._refrescar()

    # ── Datos y tabla ─────────────────────────
    def _cargar_datos(self):
        return [{"cveAlu": d["cveAlu"], "nomAlu": d["nomAlu"],
                 "edaAlu": d["edaAlu"], "cveGru": d["cveGru"]}
                for d in ColAlumno.find().sort("cveAlu", 1)]

    def _refrescar(self):
        self._todos = self._cargar_datos()
        # Refrescar también opciones del combo
        grupos = ["(todos)"] + [d["cveGru"] for d in ColGrupo.find().sort("cveGru", 1)]
        self._combo_grupo["values"] = grupos
        self._aplicar_filtro()

    def _aplicar_filtro(self):
        # Evitar que el trace dispare antes de que _tree esté construido
        if not hasattr(self, "_tree") or not hasattr(self, "_todos"):
            return
        filtro = self._var_filtro.get().strip().lower()
        grupo  = self._var_grupo.get()

        resultado = self._todos
        if filtro:
            resultado = [
                d for d in resultado
                if filtro in d["cveAlu"].lower() or filtro in d["nomAlu"].lower()
            ]
        if grupo and grupo != "(todos)":
            resultado = [d for d in resultado if d["cveGru"] == grupo]

        self._datos_filtrados = resultado
        self._page = 0
        self._actualizar_tabla()

    def _actualizar_tabla(self):
        for row in self._tree.get_children():
            self._tree.delete(row)
        inicio = self._page * PAGE_SIZE
        fin    = inicio + PAGE_SIZE
        for d in self._datos_filtrados[inicio:fin]:
            self._tree.insert("", "end", values=(
                d["cveAlu"], d["nomAlu"], d["edaAlu"], d["cveGru"]))
        self._actualizar_paginacion_labels()

    def _fila_seleccionada(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Seleccione un alumno de la tabla.", parent=self)
            return None
        return self._tree.item(sel[0])["values"]

    # ── CRUD ──────────────────────────────────
    def _agregar(self):
        p = PopupAlumno(self)
        self.wait_window(p)
        if p.resultado:
            self._refrescar()

    def _editar(self):
        vals = self._fila_seleccionada()
        if not vals: return
        cve = str(vals[0])   # Treeview puede devolver int si la clave es numérica
        doc = ColAlumno.find_one({"cveAlu": cve})
        if not doc:
            # Intentar también como int por si fue guardado así
            try:
                doc = ColAlumno.find_one({"cveAlu": int(cve)})
            except Exception:
                pass
        if not doc: return self.err("No se encontró el documento.")
        p = PopupAlumno(self, datos=doc)
        self.wait_window(p)
        if p.resultado:
            self._refrescar()

    def _eliminar(self):
        vals = self._fila_seleccionada()
        if not vals: return
        cve = str(vals[0])
        if self.conf(f"¿Eliminar al alumno '{cve}'?"):
            res = ColAlumno.delete_one({"cveAlu": cve})
            if res.deleted_count == 0:
                # Intentar como int
                try:
                    ColAlumno.delete_one({"cveAlu": int(cve)})
                except Exception:
                    pass
            self.ok("Alumno eliminado.")
            self._refrescar()

    def _eliminar_todos(self):
        if not self.conf("¿Eliminar TODOS los alumnos? Esta acción no se puede deshacer."):
            return
        ColAlumno.delete_many({})
        self.ok("Todos los alumnos han sido eliminados.")
        self._refrescar()

    # ── Exportar ──────────────────────────────
    def _datos(self):
        return self._cargar_datos()

    def _exp_csv(self):
        datos = self._datos()
        if not datos: return self.err("No hay alumnos para exportar.")
        ruta = os.path.join(BACKUP_DIR, "Alumno.csv")
        with open(ruta, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["cveAlu","nomAlu","edaAlu","cveGru"])
            w.writeheader(); w.writerows(datos)
        self.ok(f"CSV exportado:\n{ruta}")

    def _exp_json(self):
        datos = self._datos()
        if not datos: return self.err("No hay alumnos para exportar.")
        ruta = os.path.join(BACKUP_DIR, "Alumno.json")
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
        self.ok(f"JSON exportado:\n{ruta}")

    def _exp_xml(self):
        datos = self._datos()
        if not datos: return self.err("No hay alumnos para exportar.")
        root = ET.Element("Alumnos")
        for d in datos:
            a = ET.SubElement(root, "Alumno")
            for k, v in d.items():
                ET.SubElement(a, k).text = str(v)
        ruta = os.path.join(BACKUP_DIR, "Alumno.xml")
        dom = xml.dom.minidom.parseString(ET.tostring(root, encoding="unicode"))
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(dom.toprettyxml(indent="  "))
        self.ok(f"XML exportado:\n{ruta}")

    # ── Importar ──────────────────────────────
    def _imp_csv(self):
        ruta = os.path.join(BACKUP_DIR, "Alumno.csv")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el archivo:\n{ruta}\n\nPrimero exporte los datos para generar el archivo.")
        insertados = omitidos_dup = omitidos_gru = 0
        try:
            with open(ruta, encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    if ColAlumno.find_one({"cveAlu": row["cveAlu"]}):
                        omitidos_dup += 1
                        continue
                    if not ColGrupo.find_one({"cveGru": row["cveGru"]}):
                        omitidos_gru += 1
                        continue
                    ColAlumno.insert_one({
                        "cveAlu": row["cveAlu"], "nomAlu": row["nomAlu"],
                        "edaAlu": int(row["edaAlu"]), "cveGru": row["cveGru"]
                    })
                    insertados += 1
        except Exception as e:
            return self.err(f"Error al leer el archivo CSV:\n{e}")
        msg = f"CSV importado.\n✔ Insertados: {insertados}"
        if omitidos_dup: msg += f"\n⚠ Omitidos (duplicados): {omitidos_dup}"
        if omitidos_gru: msg += f"\n⚠ Omitidos (grupo no existe): {omitidos_gru}"
        self.ok(msg)
        self._refrescar()

    def _imp_json(self):
        ruta = os.path.join(BACKUP_DIR, "Alumno.json")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el archivo:\n{ruta}\n\nPrimero exporte los datos para generar el archivo.")
        insertados = omitidos_dup = omitidos_gru = 0
        try:
            with open(ruta, encoding="utf-8") as f:
                datos = json.load(f)
        except Exception as e:
            return self.err(f"Error al leer el archivo JSON:\n{e}")
        for d in datos:
            if ColAlumno.find_one({"cveAlu": d["cveAlu"]}):
                omitidos_dup += 1
                continue
            if not ColGrupo.find_one({"cveGru": d["cveGru"]}):
                omitidos_gru += 1
                continue
            ColAlumno.insert_one({
                "cveAlu": d["cveAlu"], "nomAlu": d["nomAlu"],
                "edaAlu": int(d["edaAlu"]), "cveGru": d["cveGru"]
            })
            insertados += 1
        msg = f"JSON importado.\n✔ Insertados: {insertados}"
        if omitidos_dup: msg += f"\n⚠ Omitidos (duplicados): {omitidos_dup}"
        if omitidos_gru: msg += f"\n⚠ Omitidos (grupo no existe): {omitidos_gru}"
        self.ok(msg)
        self._refrescar()

    def _imp_xml(self):
        ruta = os.path.join(BACKUP_DIR, "Alumno.xml")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el archivo:\n{ruta}\n\nPrimero exporte los datos para generar el archivo.")
        insertados = omitidos_dup = omitidos_gru = 0
        try:
            tree = ET.parse(ruta)
        except Exception as e:
            return self.err(f"Error al leer el archivo XML:\n{e}")
        for a in tree.getroot().findall("Alumno"):
            cve    = a.find("cveAlu").text
            cveGru = a.find("cveGru").text
            if ColAlumno.find_one({"cveAlu": cve}):
                omitidos_dup += 1
                continue
            if not ColGrupo.find_one({"cveGru": cveGru}):
                omitidos_gru += 1
                continue
            ColAlumno.insert_one({
                "cveAlu": cve,
                "nomAlu": a.find("nomAlu").text,
                "edaAlu": int(a.find("edaAlu").text),
                "cveGru": cveGru
            })
            insertados += 1
        msg = f"XML importado.\n✔ Insertados: {insertados}"
        if omitidos_dup: msg += f"\n⚠ Omitidos (duplicados): {omitidos_dup}"
        if omitidos_gru: msg += f"\n⚠ Omitidos (grupo no existe): {omitidos_gru}"
        self.ok(msg)
        self._refrescar()

    # ── Backup / Restaurar ────────────────────
    def _backup(self):
        exe = _find_mongo_tool("mongoexport")
        if not exe:
            return self.err(
                "No se encontró 'mongoexport'.\n"
                "Instala MongoDB Database Tools:\n"
                "https://www.mongodb.com/try/download/database-tools\n"
                "y agrega la carpeta bin al PATH del sistema."
            )
        try:
            salida = os.path.join(BACKUP_DIR, "Alumno_backup.json")
            cmd = [exe,
                   "--host=localhost", "--port=27017",
                   "--username=admin", "--password=root",
                   "--authenticationDatabase=admin",
                   "--db=BD_GrupoAlumno", "--collection=Alumno",
                   f"--out={salida}"]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok(f"Backup ejecutado.\n{salida}")
            else:
                self.err(f"Error en backup:\n{result.stderr}")
        except Exception as e:
            self.err(str(e))

    def _restaurar_todos(self):
        ruta = os.path.join(BACKUP_DIR, "Alumno_backup.json")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el backup:\n{ruta}")
        exe = _find_mongo_tool("mongoimport")
        if not exe:
            return self.err("No se encontró 'mongoimport'. Instala MongoDB Database Tools.")
        if not self.conf("¿Restaurar TODOS los alumnos desde el backup?"):
            return
        try:
            cmd = [exe,
                   "--host=localhost", "--port=27017",
                   "--username=admin", "--password=root",
                   "--authenticationDatabase=admin",
                   "--db=BD_GrupoAlumno", "--collection=Alumno",
                   f"--file={ruta}"]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok("Alumnos restaurados correctamente.")
                self._refrescar()
            else:
                self.err(f"Error al restaurar:\n{result.stderr}")
        except Exception as e:
            self.err(str(e))


# ══════════════════════════════════════════════
#  ARRANQUE
# ══════════════════════════════════════════════
if __name__ == "__main__":
    app = VentanaPrincipal()
    app.mainloop()