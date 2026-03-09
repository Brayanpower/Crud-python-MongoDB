import tkinter as tk
from tkinter import messagebox, ttk
import json
import csv
import xml.etree.ElementTree as ET
import xml.dom.minidom
import os
import subprocess
import shutil
from datetime import datetime
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

    # ── Estilo global ──────────────────────────
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
                   command=self._abrir_grupos, width=28).grid(row=0, column=0, padx=8, pady=6)
        ttk.Button(frame, text="⊞  Administrar Alumnos",
                   command=self._abrir_alumnos, width=28).grid(row=1, column=0, padx=8, pady=6)

    def _abrir_grupos(self):
        VentanaGrupo(self)

    def _abrir_alumnos(self):
        VentanaAlumno(self)


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

    def __init__(self, parent, titulo, ancho, alto):
        super().__init__(parent)
        self.title(titulo)
        self.geometry(f"{ancho}x{alto}")
        self.resizable(False, False)
        self.configure(bg=self.COLOR_BG)
        self._entries = {}
        self._construir()

    # ── helpers ───────────────────────────────
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

    def _separador(self, parent, row, cols=3):
        tk.Frame(parent, bg=self.COLOR_BORDE, height=1).grid(
            row=row, column=0, columnspan=cols, sticky="ew",
            padx=10, pady=4)

    def _get(self, key):
        return self._entries[key].get().strip()

    def _set(self, key, valor):
        self._entries[key].delete(0, tk.END)
        self._entries[key].insert(0, valor)

    def _limpiar(self):
        for e in self._entries.values():
            e.delete(0, tk.END)

    def ok(self, msg):   messagebox.showinfo("✔ Éxito", msg, parent=self)
    def err(self, msg):  messagebox.showerror("✘ Error", msg, parent=self)
    def conf(self, msg): return messagebox.askyesno("¿Confirmar?", msg, parent=self)


# ══════════════════════════════════════════════
#  VENTANA GRUPOS
# ══════════════════════════════════════════════
class VentanaGrupo(VentanaBase):
    def __init__(self, parent):
        super().__init__(parent, "Admon Grupo", 540, 440)

    def _construir(self):
        self.grid_columnconfigure(0, weight=1)

        # ── Título ────────────────────────────
        tk.Label(self, text="▸ ADMINISTRACIÓN DE GRUPOS",
                 bg=self.COLOR_BG, fg="#38bdf8",
                 font=self.FUENTE_TITULO).grid(
                 row=0, column=0, pady=(16, 4), padx=16, sticky="w")

        # ── Panel de campos ───────────────────
        panel = tk.Frame(self, bg=self.COLOR_PANEL,
                         highlightthickness=1,
                         highlightbackground=self.COLOR_BORDE)
        panel.grid(row=1, column=0, padx=16, pady=8, sticky="ew")
        panel.grid_columnconfigure(1, weight=1)

        self._label(panel, "Clave:").grid(row=0, column=0, padx=12, pady=8, sticky="w")
        self._entry(panel, "cve", width=20).grid(row=0, column=1, padx=8, pady=8, sticky="ew")

        self._label(panel, "Nombre:").grid(row=1, column=0, padx=12, pady=8, sticky="w")
        self._entry(panel, "nom", width=30).grid(row=1, column=1, padx=8, pady=8, sticky="ew")

        # ── Botones CRUD ──────────────────────
        bf = tk.Frame(self, bg=self.COLOR_BG)
        bf.grid(row=2, column=0, padx=16, pady=4, sticky="ew")
        for i in range(3): bf.grid_columnconfigure(i, weight=1)

        self._boton(bf, "＋ Agregar",   self._agregar,   "#15803d").grid(row=0, column=0, padx=4, pady=4, sticky="ew")
        self._boton(bf, "✎ Modificar",  self._modificar, "#1d4ed8").grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        self._boton(bf, "✕ Eliminar",   self._eliminar,  "#b91c1c").grid(row=0, column=2, padx=4, pady=4, sticky="ew")
        self._boton(bf, "⌕ Buscar",     self._buscar,    "#0e7490").grid(row=1, column=0, padx=4, pady=4, sticky="ew")
        self._boton(bf, "↺ Limpiar",    self._limpiar                ).grid(row=1, column=1, padx=4, pady=4, sticky="ew")

        # ── Exportar / Importar ───────────────
        ef = tk.Frame(self, bg=self.COLOR_BG)
        ef.grid(row=3, column=0, padx=16, pady=4, sticky="ew")
        for i in range(3): ef.grid_columnconfigure(i, weight=1)

        tk.Label(ef, text="── Exportar", bg=self.COLOR_BG,
                 fg=self.COLOR_MUTED, font=("Consolas", 8)).grid(
                 row=0, column=0, columnspan=3, sticky="w")
        self._boton(ef, "CSV",  self._exp_csv ).grid(row=1, column=0, padx=3, pady=2, sticky="ew")
        self._boton(ef, "JSON", self._exp_json).grid(row=1, column=1, padx=3, pady=2, sticky="ew")
        self._boton(ef, "XML",  self._exp_xml ).grid(row=1, column=2, padx=3, pady=2, sticky="ew")

        tk.Label(ef, text="── Importar", bg=self.COLOR_BG,
                 fg=self.COLOR_MUTED, font=("Consolas", 8)).grid(
                 row=2, column=0, columnspan=3, sticky="w", pady=(6,0))
        self._boton(ef, "CSV",  self._imp_csv ).grid(row=3, column=0, padx=3, pady=2, sticky="ew")
        self._boton(ef, "JSON", self._imp_json).grid(row=3, column=1, padx=3, pady=2, sticky="ew")
        self._boton(ef, "XML",  self._imp_xml ).grid(row=3, column=2, padx=3, pady=2, sticky="ew")

        # ── Backup / Peligro ──────────────────
        df = tk.Frame(self, bg=self.COLOR_BG)
        df.grid(row=4, column=0, padx=16, pady=(6,12), sticky="ew")
        df.grid_columnconfigure(0, weight=1)

        self._boton(df, "⬡ Ejecutar Backup",          self._backup,          "#7c3aed").grid(row=0, column=0, pady=3, sticky="ew")
        self._boton(df, "⚠ Eliminar todos los Grupos", self._eliminar_todos,   self.COLOR_PELIGRO).grid(row=1, column=0, pady=3, sticky="ew")
        self._boton(df, "⟳ Restaurar todos los Grupos",self._restaurar_todos,  "#b45309").grid(row=2, column=0, pady=3, sticky="ew")

    # ── CRUD ──────────────────────────────────
    def _agregar(self):
        cve, nom = self._get("cve"), self._get("nom")
        if not cve or not nom:
            return self.err("Clave y Nombre son obligatorios.")
        if ColGrupo.find_one({"cveGru": cve}):
            return self.err(f"Ya existe un grupo con clave '{cve}'.")
        ColGrupo.insert_one({"cveGru": cve, "nomGru": nom})
        self.ok("Grupo agregado correctamente.")
        self._limpiar()

    def _buscar(self):
        cve = self._get("cve")
        if not cve:
            return self.err("Ingrese la clave para buscar.")
        doc = ColGrupo.find_one({"cveGru": cve})
        if not doc:
            return self.err(f"No se encontró el grupo '{cve}'.")
        self._set("nom", doc["nomGru"])

    def _modificar(self):
        cve, nom = self._get("cve"), self._get("nom")
        if not cve or not nom:
            return self.err("Busque primero el grupo (Clave + Buscar).")
        if not ColGrupo.find_one({"cveGru": cve}):
            return self.err(f"No existe el grupo '{cve}'.")
        ColGrupo.update_one({"cveGru": cve}, {"$set": {"nomGru": nom}})
        self.ok("Grupo modificado correctamente.")

    def _eliminar(self):
        cve = self._get("cve")
        if not cve:
            return self.err("Ingrese la clave del grupo a eliminar.")
        if not ColGrupo.find_one({"cveGru": cve}):
            return self.err(f"No existe el grupo '{cve}'.")
        if ColAlumno.find_one({"cveGru": cve}):
            return self.err(f"No puede eliminar: hay alumnos asignados al grupo '{cve}'.")
        if self.conf(f"¿Eliminar el grupo '{cve}'?"):
            ColGrupo.delete_one({"cveGru": cve})
            self.ok("Grupo eliminado.")
            self._limpiar()

    def _eliminar_todos(self):
        if ColAlumno.count_documents({}) > 0:
            return self.err("No puede eliminar todos los grupos: existen alumnos registrados.")
        if self.conf("¿Eliminar TODOS los grupos? Esta acción no se puede deshacer."):
            ColGrupo.delete_many({})
            self.ok("Todos los grupos han sido eliminados.")

    # ── Exportar ──────────────────────────────
    def _datos(self):
        return [{"cveGru": d["cveGru"], "nomGru": d["nomGru"]}
                for d in ColGrupo.find()]

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
        if not os.path.exists(ruta): return self.err(f"No se encontró:\n{ruta}")
        with open(ruta, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if not ColGrupo.find_one({"cveGru": row["cveGru"]}):
                    ColGrupo.insert_one({"cveGru": row["cveGru"], "nomGru": row["nomGru"]})
        self.ok("CSV importado correctamente.")

    def _imp_json(self):
        ruta = os.path.join(BACKUP_DIR, "Grupo.json")
        if not os.path.exists(ruta): return self.err(f"No se encontró:\n{ruta}")
        with open(ruta, encoding="utf-8") as f:
            for d in json.load(f):
                if not ColGrupo.find_one({"cveGru": d["cveGru"]}):
                    ColGrupo.insert_one({"cveGru": d["cveGru"], "nomGru": d["nomGru"]})
        self.ok("JSON importado correctamente.")

    def _imp_xml(self):
        ruta = os.path.join(BACKUP_DIR, "Grupo.xml")
        if not os.path.exists(ruta): return self.err(f"No se encontró:\n{ruta}")
        tree = ET.parse(ruta)
        for g in tree.getroot().findall("Grupo"):
            cve = g.find("cveGru").text
            nom = g.find("nomGru").text
            if not ColGrupo.find_one({"cveGru": cve}):
                ColGrupo.insert_one({"cveGru": cve, "nomGru": nom})
        self.ok("XML importado correctamente.")

    # ── Backup / Restaurar ────────────────────
    def _backup(self):
        try:
            cmd = (
                f'mongoexport --host=localhost --port=27017 '
                f'--username=admin --password=root '
                f'--authenticationDatabase=admin '
                f'--db=BD_GrupoAlumno --collection=Grupo '
                f'--out="{os.path.join(BACKUP_DIR, "Grupo_backup.json")}"'
            )
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok(f"Backup ejecutado.\n{BACKUP_DIR}\\Grupo_backup.json")
            else:
                self.err(f"Error en backup:\n{result.stderr}")
        except Exception as e:
            self.err(str(e))

    def _restaurar_todos(self):
        ruta = os.path.join(BACKUP_DIR, "Grupo_backup.json")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el backup:\n{ruta}")
        if not self.conf("¿Restaurar TODOS los grupos desde el backup?"):
            return
        try:
            cmd = (
                f'mongoimport --host=localhost --port=27017 '
                f'--username=admin --password=root '
                f'--authenticationDatabase=admin '
                f'--db=BD_GrupoAlumno --collection=Grupo '
                f'--file="{ruta}"'
            )
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok("Grupos restaurados correctamente.")
            else:
                self.err(f"Error al restaurar:\n{result.stderr}")
        except Exception as e:
            self.err(str(e))


# ══════════════════════════════════════════════
#  VENTANA ALUMNOS
# ══════════════════════════════════════════════
class VentanaAlumno(VentanaBase):
    def __init__(self, parent):
        super().__init__(parent, "Admon Alumno", 540, 520)

    def _construir(self):
        self.grid_columnconfigure(0, weight=1)

        tk.Label(self, text="▸ ADMINISTRACIÓN DE ALUMNOS",
                 bg=self.COLOR_BG, fg="#38bdf8",
                 font=self.FUENTE_TITULO).grid(
                 row=0, column=0, pady=(16,4), padx=16, sticky="w")

        panel = tk.Frame(self, bg=self.COLOR_PANEL,
                         highlightthickness=1,
                         highlightbackground=self.COLOR_BORDE)
        panel.grid(row=1, column=0, padx=16, pady=8, sticky="ew")
        panel.grid_columnconfigure(1, weight=1)

        campos = [("Clave Alumno:", "cve"), ("Nombre Alumno:", "nom"),
                  ("Edad:", "eda"), ("Clave Grupo:", "cveGru")]
        for i, (lbl, key) in enumerate(campos):
            self._label(panel, lbl).grid(row=i, column=0, padx=12, pady=7, sticky="w")
            self._entry(panel, key, width=28).grid(row=i, column=1, padx=8, pady=7, sticky="ew")

        # CRUD
        bf = tk.Frame(self, bg=self.COLOR_BG)
        bf.grid(row=2, column=0, padx=16, pady=4, sticky="ew")
        for i in range(3): bf.grid_columnconfigure(i, weight=1)

        self._boton(bf, "＋ Agregar",  self._agregar,  "#15803d").grid(row=0, column=0, padx=4, pady=4, sticky="ew")
        self._boton(bf, "✎ Modificar", self._modificar,"#1d4ed8").grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        self._boton(bf, "✕ Eliminar",  self._eliminar, "#b91c1c").grid(row=0, column=2, padx=4, pady=4, sticky="ew")
        self._boton(bf, "⌕ Buscar",    self._buscar,   "#0e7490").grid(row=1, column=0, padx=4, pady=4, sticky="ew")
        self._boton(bf, "↺ Limpiar",   self._limpiar              ).grid(row=1, column=1, padx=4, pady=4, sticky="ew")

        # Exportar / Importar
        ef = tk.Frame(self, bg=self.COLOR_BG)
        ef.grid(row=3, column=0, padx=16, pady=4, sticky="ew")
        for i in range(3): ef.grid_columnconfigure(i, weight=1)

        tk.Label(ef, text="── Exportar", bg=self.COLOR_BG,
                 fg=self.COLOR_MUTED, font=("Consolas", 8)).grid(
                 row=0, column=0, columnspan=3, sticky="w")
        self._boton(ef, "CSV",  self._exp_csv ).grid(row=1, column=0, padx=3, pady=2, sticky="ew")
        self._boton(ef, "JSON", self._exp_json).grid(row=1, column=1, padx=3, pady=2, sticky="ew")
        self._boton(ef, "XML",  self._exp_xml ).grid(row=1, column=2, padx=3, pady=2, sticky="ew")

        tk.Label(ef, text="── Importar", bg=self.COLOR_BG,
                 fg=self.COLOR_MUTED, font=("Consolas", 8)).grid(
                 row=2, column=0, columnspan=3, sticky="w", pady=(6,0))
        self._boton(ef, "CSV",  self._imp_csv ).grid(row=3, column=0, padx=3, pady=2, sticky="ew")
        self._boton(ef, "JSON", self._imp_json).grid(row=3, column=1, padx=3, pady=2, sticky="ew")
        self._boton(ef, "XML",  self._imp_xml ).grid(row=3, column=2, padx=3, pady=2, sticky="ew")

        # Backup / Peligro
        df = tk.Frame(self, bg=self.COLOR_BG)
        df.grid(row=4, column=0, padx=16, pady=(6,12), sticky="ew")
        df.grid_columnconfigure(0, weight=1)

        self._boton(df, "⬡ Ejecutar Backup",           self._backup,         "#7c3aed").grid(row=0, column=0, pady=3, sticky="ew")
        self._boton(df, "⚠ Eliminar todos los Alumnos", self._eliminar_todos, self.COLOR_PELIGRO).grid(row=1, column=0, pady=3, sticky="ew")
        self._boton(df, "⟳ Restaurar todos los Alumnos",self._restaurar_todos,"#b45309").grid(row=2, column=0, pady=3, sticky="ew")

    # ── CRUD ──────────────────────────────────
    def _agregar(self):
        cve    = self._get("cve")
        nom    = self._get("nom")
        eda    = self._get("eda")
        cveGru = self._get("cveGru")
        if not all([cve, nom, eda, cveGru]):
            return self.err("Todos los campos son obligatorios.")
        if not eda.isdigit():
            return self.err("La edad debe ser un número entero.")
        if not ColGrupo.find_one({"cveGru": cveGru}):
            return self.err(f"El grupo '{cveGru}' no existe.")
        if ColAlumno.find_one({"cveAlu": cve}):
            return self.err(f"Ya existe un alumno con clave '{cve}'.")
        ColAlumno.insert_one({"cveAlu": cve, "nomAlu": nom,
                               "edaAlu": int(eda), "cveGru": cveGru})
        self.ok("Alumno agregado correctamente.")
        self._limpiar()

    def _buscar(self):
        cve = self._get("cve")
        if not cve: return self.err("Ingrese la clave para buscar.")
        doc = ColAlumno.find_one({"cveAlu": cve})
        if not doc: return self.err(f"No se encontró el alumno '{cve}'.")
        self._set("nom", doc["nomAlu"])
        self._set("eda", str(doc["edaAlu"]))
        self._set("cveGru", doc["cveGru"])

    def _modificar(self):
        cve    = self._get("cve")
        nom    = self._get("nom")
        eda    = self._get("eda")
        cveGru = self._get("cveGru")
        if not all([cve, nom, eda, cveGru]):
            return self.err("Busque primero el alumno y complete los campos.")
        if not eda.isdigit():
            return self.err("La edad debe ser un número entero.")
        if not ColAlumno.find_one({"cveAlu": cve}):
            return self.err(f"No existe el alumno '{cve}'.")
        if not ColGrupo.find_one({"cveGru": cveGru}):
            return self.err(f"El grupo '{cveGru}' no existe.")
        ColAlumno.update_one({"cveAlu": cve},
            {"$set": {"nomAlu": nom, "edaAlu": int(eda), "cveGru": cveGru}})
        self.ok("Alumno modificado correctamente.")

    def _eliminar(self):
        cve = self._get("cve")
        if not cve: return self.err("Ingrese la clave del alumno a eliminar.")
        if not ColAlumno.find_one({"cveAlu": cve}):
            return self.err(f"No existe el alumno '{cve}'.")
        if self.conf(f"¿Eliminar al alumno '{cve}'?"):
            ColAlumno.delete_one({"cveAlu": cve})
            self.ok("Alumno eliminado.")
            self._limpiar()

    def _eliminar_todos(self):
        if not self.conf("¿Eliminar TODOS los alumnos? Esta acción no se puede deshacer."):
            return
        ColAlumno.delete_many({})
        self.ok("Todos los alumnos han sido eliminados.")

    # ── Exportar ──────────────────────────────
    def _datos(self):
        return [{"cveAlu": d["cveAlu"], "nomAlu": d["nomAlu"],
                 "edaAlu": d["edaAlu"], "cveGru": d["cveGru"]}
                for d in ColAlumno.find()]

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
        if not os.path.exists(ruta): return self.err(f"No se encontró:\n{ruta}")
        with open(ruta, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if not ColAlumno.find_one({"cveAlu": row["cveAlu"]}):
                    ColAlumno.insert_one({
                        "cveAlu": row["cveAlu"], "nomAlu": row["nomAlu"],
                        "edaAlu": int(row["edaAlu"]), "cveGru": row["cveGru"]
                    })
        self.ok("CSV importado correctamente.")

    def _imp_json(self):
        ruta = os.path.join(BACKUP_DIR, "Alumno.json")
        if not os.path.exists(ruta): return self.err(f"No se encontró:\n{ruta}")
        with open(ruta, encoding="utf-8") as f:
            for d in json.load(f):
                if not ColAlumno.find_one({"cveAlu": d["cveAlu"]}):
                    ColAlumno.insert_one({
                        "cveAlu": d["cveAlu"], "nomAlu": d["nomAlu"],
                        "edaAlu": int(d["edaAlu"]), "cveGru": d["cveGru"]
                    })
        self.ok("JSON importado correctamente.")

    def _imp_xml(self):
        ruta = os.path.join(BACKUP_DIR, "Alumno.xml")
        if not os.path.exists(ruta): return self.err(f"No se encontró:\n{ruta}")
        tree = ET.parse(ruta)
        for a in tree.getroot().findall("Alumno"):
            cve = a.find("cveAlu").text
            if not ColAlumno.find_one({"cveAlu": cve}):
                ColAlumno.insert_one({
                    "cveAlu": cve,
                    "nomAlu": a.find("nomAlu").text,
                    "edaAlu": int(a.find("edaAlu").text),
                    "cveGru": a.find("cveGru").text
                })
        self.ok("XML importado correctamente.")

    # ── Backup / Restaurar ────────────────────
    def _backup(self):
        try:
            cmd = (
                f'mongoexport --host=localhost --port=27017 '
                f'--username=admin --password=root '
                f'--authenticationDatabase=admin '
                f'--db=BD_GrupoAlumno --collection=Alumno '
                f'--out="{os.path.join(BACKUP_DIR, "Alumno_backup.json")}"'
            )
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok(f"Backup ejecutado.\n{BACKUP_DIR}\\Alumno_backup.json")
            else:
                self.err(f"Error en backup:\n{result.stderr}")
        except Exception as e:
            self.err(str(e))

    def _restaurar_todos(self):
        ruta = os.path.join(BACKUP_DIR, "Alumno_backup.json")
        if not os.path.exists(ruta):
            return self.err(f"No se encontró el backup:\n{ruta}")
        if not self.conf("¿Restaurar TODOS los alumnos desde el backup?"):
            return
        try:
            cmd = (
                f'mongoimport --host=localhost --port=27017 '
                f'--username=admin --password=root '
                f'--authenticationDatabase=admin '
                f'--db=BD_GrupoAlumno --collection=Alumno '
                f'--file="{ruta}"'
            )
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                self.ok("Alumnos restaurados correctamente.")
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