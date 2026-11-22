"""Parole: calculadora de fechas de condena y generador de documentos Word.

Interfaz gráfica (Tkinter). La lógica de cálculo está en calculos.py.
"""

import json
import os
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt
except ImportError:
    # Aviso claro en vez de un error técnico si falta la librería
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "Falta una librería",
        "Para generar documentos Word hace falta instalar python-docx.\n\n"
        "Abre una terminal y escribe:\n\n"
        "    py -m pip install -r requirements.txt\n\n"
        "(o: py -m pip install python-docx)",
    )
    sys.exit(1)

from calculos import (calcular_condena, convertir_dias, edad_en, fecha_larga,
                      nacimiento_desde_ci, parsear_fecha, rellenar_plantilla)

# Carpeta de la app: la plantilla y el escudo se buscan aquí aunque se abra
# el programa desde otra carpeta (antes se perdía la plantilla en ese caso).
if getattr(sys, "frozen", False):  # ejecutable creado con PyInstaller
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Marcadores que se escriben en negrita en el documento
MARCADORES_NEGRITA = ["(nombre)", "(delito)", "(tribunal)", "(causa)", "(favorable)"]

class PrisionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Calculadora de Tiempo en Prisión")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        self.set_style()
        
        self.plantilla_file = os.path.join(BASE_DIR, "plantilla.json")
        
        # Variables de datos del recluso
        self.nombre = tk.StringVar()
        self.establecimiento = tk.StringVar()
        self.carnet = tk.StringVar()
        self.edad_actual = tk.StringVar()
        self.edad_entrada = tk.StringVar()
        self.direccion = tk.StringVar()
        self.delito = tk.StringVar()
        self.tribunal = tk.StringVar()
        self.causa = tk.StringVar()
        self.inciso = tk.StringVar()
        self.escolaridad = tk.StringVar()
        self.natural_de = tk.StringVar()
        self.estado_civil = tk.StringVar()
        self.regimen = tk.StringVar()
        self.posh = tk.StringVar()
        self.razon = tk.StringVar()
        self.clasificacion = tk.StringVar()
        self.favorable = tk.StringVar()
        self.trabajo_social = tk.StringVar()
        self.oficios = tk.StringVar()
        self.responsabilidad_civil = tk.StringVar()
        self.indisciplinas = tk.StringVar()
        
        self.entrada = tk.StringVar()
        self.salida = tk.StringVar()
        self.rebaja = tk.StringVar()
        self.prev_inicio = tk.StringVar()
        self.prev_fin = tk.StringVar()
        self.prev_dias = tk.StringVar()
        
        self.fecha_entrada_dt = None
        self.fecha_salida_dt = None
        self.hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        self.condena_original_dias = 0
        self.condena_original_str = ""
        self.fecha_mitad_original = None
        self.fecha_tercio_original = None
        self.fecha_dos_tercios_original = None
        self.fecha_salida_original_con_preventiva = None
        self.fecha_salida_con_rebaja = None
        self.fecha_mitad_con_rebaja = None
        self.fecha_tercio_con_rebaja = None
        self.fecha_dos_tercios_con_rebaja = None
        self.dias_cumplidos = 0
        self.dias_restantes = 0
        self.cumplido_str = ""
        self.restante_str = ""
        self.rebaja_valor = 0
        self.prev_dias_valor = 0
        
        self.ruta_escudo = tk.StringVar(value=os.path.join(BASE_DIR, "escudo.png"))
        
        self.crear_widgets()
        self.cargar_plantilla()
        self.root.bind('<Return>', self.enter_calcular)
    
    def set_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        bg_color = "#f5f5f5"
        fg_color = "#2c3e50"
        select_bg = "#3498db"
        select_fg = "white"
        field_bg = "#ffffff"
        field_border = "#bdc3c7"
        button_bg = "#ecf0f1"
        button_fg = "#2c3e50"
        button_active = "#d0d3d4"
        
        style.configure('TNotebook', background=bg_color, borderwidth=0)
        style.configure('TNotebook.Tab', background=button_bg, foreground=fg_color, padding=[10, 5], font=('Segoe UI', 10))
        style.map('TNotebook.Tab', background=[('selected', '#ffffff'), ('active', '#d0d3d4')])
        
        style.configure('TLabel', background=bg_color, foreground=fg_color, font=('Segoe UI', 10))
        style.configure('TEntry', fieldbackground=field_bg, foreground=fg_color, borderwidth=1, relief='solid')
        style.configure('TButton', background=button_bg, foreground=button_fg, font=('Segoe UI', 10, 'bold'), borderwidth=1, focusthickness=3)
        style.map('TButton', background=[('active', button_active), ('pressed', '#b0b3b5')])
        
        style.configure('TFrame', background=bg_color)
        style.configure('TLabelframe', background=bg_color, foreground=fg_color)
        style.configure('TLabelframe.Label', background=bg_color, foreground=fg_color)
        
        self.root.configure(bg=bg_color)
    
    def enter_calcular(self, event):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.calcular()
    
    def crear_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tab1 = self.crear_pestana_con_scroll(self.notebook, "Cálculos")
        self.crear_pestana_calculos(tab1)
        
        tab2 = self.crear_pestana_con_scroll(self.notebook, "Datos del recluso")
        self.crear_pestana_datos(tab2)
        
        tab3 = ttk.Frame(self.notebook)
        self.notebook.add(tab3, text="Plantilla")
        self.crear_pestana_plantilla_mejorada(tab3)  # Versión mejorada con botones
        
        # Pie de página
        pie_frame = ttk.Frame(self.root)
        pie_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0,5))
        pie_label = ttk.Label(pie_frame, text="Alex Alvarez Morales 2026", font=('Segoe UI', 9, 'italic'), foreground='#7f8c8d')
        pie_label.pack()
    
    def crear_pestana_con_scroll(self, notebook, titulo):
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=titulo)
        canvas = tk.Canvas(frame, borderwidth=0, highlightthickness=0, bg='#f5f5f5')
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return scrollable_frame
    
    def crear_pestana_calculos(self, parent):
        frame = ttk.Frame(parent, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        row = 0
        ttk.Label(frame, text="Fecha de entrada (DD/MM/YYYY):", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=(5,5))
        ttk.Entry(frame, textvariable=self.entrada, width=25, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=(5,5))
        row+=1
        
        ttk.Label(frame, text="Fecha de salida (DD/MM/YYYY):", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.salida, width=25, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row+=1
        
        ttk.Label(frame, text="Días de rebaja de sanción:", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.rebaja, width=25, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row+=1
        
        ttk.Label(frame, text="Preventiva (solo para calcular condena original):", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(15,5))
        row+=1
        
        ttk.Label(frame, text="Inicio (DD/MM/YYYY):").grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.prev_inicio, width=25, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row+=1
        
        ttk.Label(frame, text="Fin (DD/MM/YYYY):").grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.prev_fin, width=25, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row+=1
        
        ttk.Label(frame, text="O solo días de preventiva:").grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.prev_dias, width=25, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        row+=1
        
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=25)
        self.btn_calcular = tk.Button(btn_frame, text="Calcular", command=self.calcular,
                                      bg='#3498db', fg='white', font=('Segoe UI', 11, 'bold'),
                                      bd=0, padx=25, pady=8, cursor='hand2',
                                      activebackground='#2980b9', activeforeground='white')
        self.btn_calcular.pack()
        row+=1
        
        ttk.Label(frame, text="Resultados:", font=('Segoe UI', 11, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(10,5))
        row+=1
        self.resultado_calculos = tk.Text(frame, height=15, width=80, font=('Consolas', 9), wrap=tk.WORD,
                                          bg='#ffffff', fg='#2c3e50', bd=1, relief='solid')
        self.resultado_calculos.grid(row=row, column=0, columnspan=2, pady=5)
        self.resultado_calculos.config(state=tk.DISABLED)
    
    def crear_pestana_datos(self, parent):
        frame = ttk.Frame(parent, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        row = 0
        ttk.Label(frame, text="Ruta del escudo (PNG/JPG):", font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.ruta_escudo, width=45, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
        ttk.Button(frame, text="Examinar...", command=self.seleccionar_escudo).grid(row=row, column=2, padx=5)
        row+=1
        
        campos = [
            ("Nombre completo:", self.nombre),
            ("Establecimiento penitenciario:", self.establecimiento),
            ("Carnet de identidad:", self.carnet),
            ("Edad actual (calculada):", self.edad_actual),
            ("Edad al momento de entrar (calculada):", self.edad_entrada),
            ("Dirección:", self.direccion),
            ("Delito:", self.delito),
            ("Tribunal:", self.tribunal),
            ("Causa:", self.causa),
            ("Inciso:", self.inciso),
            ("Grado de escolaridad:", self.escolaridad),
            ("Natural de:", self.natural_de),
            ("Estado civil:", self.estado_civil),
            ("Régimen:", self.regimen),
            ("POSH:", self.posh),
            ("Razón (motivo):", self.razon),
            ("Clasificación:", self.clasificacion),
            ("Favorable/Desfavorable:", self.favorable),
            ("Trabajo socialmente útil (dónde trabaja):", self.trabajo_social),
            ("Oficios en los que se ha capacitado:", self.oficios),
            ("Responsabilidad civil (texto adicional):", self.responsabilidad_civil),
            ("Indisciplinas:", self.indisciplinas)
        ]
        
        self.carnet.trace_add("write", self.calcular_edades)
        self.entrada.trace_add("write", self.calcular_edades)
        
        for i, (label, var) in enumerate(campos):
            ttk.Label(frame, text=label, font=('Segoe UI', 10)).grid(row=row, column=0, sticky=tk.W, pady=4)
            ttk.Entry(frame, textvariable=var, width=50, font=('Segoe UI', 10)).grid(row=row, column=1, sticky=tk.W, padx=10, pady=4)
            row+=1
    
    def seleccionar_escudo(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar imagen del escudo",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.gif"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.ruta_escudo.set(filename)
    
    def calcular_edades(self, *args):
        """Rellena la edad actual y la edad al entrar a partir del carnet."""
        nacimiento = nacimiento_desde_ci(self.carnet.get(), self.hoy)
        if nacimiento is None:
            return
        self.edad_actual.set(str(edad_en(nacimiento, self.hoy)))
        try:
            entrada = parsear_fecha(self.entrada.get())
            self.edad_entrada.set(str(edad_en(nacimiento, entrada)))
        except ValueError:
            self.edad_entrada.set("")

    def crear_pestana_plantilla_mejorada(self, parent):
        """Versión con botones para insertar marcadores"""
        # Panel principal con dos columnas
        main_panel = ttk.Frame(parent, padding="10")
        main_panel.pack(fill=tk.BOTH, expand=True)
        
        # Columna izquierda: botones de marcadores
        left_frame = ttk.LabelFrame(main_panel, text="Marcadores", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0,10))
        
        # Crear un canvas con scroll para los botones (por si hay muchos)
        canvas_left = tk.Canvas(left_frame, borderwidth=0, highlightthickness=0, bg='#f5f5f5', width=200)
        scroll_left = ttk.Scrollbar(left_frame, orient="vertical", command=canvas_left.yview)
        scrollable_left = ttk.Frame(canvas_left)
        scrollable_left.bind("<Configure>", lambda e: canvas_left.configure(scrollregion=canvas_left.bbox("all")))
        canvas_left.create_window((0, 0), window=scrollable_left, anchor="nw")
        canvas_left.configure(yscrollcommand=scroll_left.set)
        canvas_left.pack(side="left", fill="both", expand=True)
        scroll_left.pack(side="right", fill="y")
        
        # Organizar marcadores por categorías
        categorias = {
            "Datos personales": ["(nombre)", "(establecimiento)", "(ci)", "(edadactual)", "(edaddentrada)", "(direccion)"],
            "Delito y tribunal": ["(delito)", "(tribunal)", "(causa)", "(inciso)", "(posh)", "(razon)", "(clasif)", "(favorable)"],
            "Información adicional": ["(escolaridad)", "(naturalde)", "(estadocivil)", "(regimen)", "(trabajosocial)", "(oficios)", "(respcivil)", "(indisciplinas)"],
            "Fechas originales": ["(FechaEntrada)", "(FechaSalida)", "(CondenaOriginalInfo)", "(MitadOriginal)", "(TercioOriginal)", "(DosTerciosOriginal)"],
            "Fechas con rebaja": ["(FechaSalidaConRebaja)", "(MitadConRebaja)", "(TercioConRebaja)", "(DosTerciosConRebaja)", "(DiasRebaja)"],
            "Preventiva": ["(DiasPreventiva)", "(FechaSalidaOriginalConPreventiva)", "(CondenaOriginalTotal)"],
            "Otros": ["(FechaActual)"]
        }
        
        for cat, marcadores in categorias.items():
            ttk.Label(scrollable_left, text=cat, font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W, pady=(10,2))
            for marc in marcadores:
                btn = tk.Button(scrollable_left, text=marc, font=('Consolas', 9),
                                bg='#ecf0f1', fg='#2c3e50', bd=1, relief='solid',
                                cursor='hand2', padx=5, pady=2,
                                command=lambda m=marc: self.insertar_marcador(m))
                btn.pack(anchor=tk.W, pady=1, fill=tk.X)
        
        # Columna derecha: área de texto de la plantilla
        right_frame = ttk.Frame(main_panel)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        ttk.Label(right_frame, text="Plantilla (escribe o haz clic en los marcadores)", font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        
        self.plantilla_text = tk.Text(right_frame, height=20, width=70, font=('Consolas', 10), wrap=tk.WORD,
                                      bg='#ffffff', fg='#2c3e50', bd=1, relief='solid', undo=True)
        self.plantilla_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Barra de herramientas inferior
        tool_frame = ttk.Frame(right_frame)
        tool_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(tool_frame, text="Guardar plantilla", command=self.guardar_plantilla).pack(side=tk.LEFT, padx=5)
        ttk.Button(tool_frame, text="Generar documento", command=self.generar_documento).pack(side=tk.LEFT, padx=5)
        ttk.Button(tool_frame, text="Limpiar plantilla", command=self.limpiar_plantilla).pack(side=tk.LEFT, padx=5)
    
    def insertar_marcador(self, marcador):
        """Inserta el marcador en la posición actual del cursor en el texto de la plantilla"""
        self.plantilla_text.insert(tk.INSERT, marcador)
        self.plantilla_text.focus_set()
    
    def cargar_plantilla(self):
        if os.path.exists(self.plantilla_file):
            try:
                with open(self.plantilla_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    plantilla = data.get("plantilla", "")
                    self.plantilla_text.insert("1.0", plantilla)
            except (OSError, json.JSONDecodeError) as e:
                messagebox.showwarning("Plantilla", f"No se pudo cargar la plantilla guardada:\n{e}")
    
    def guardar_plantilla(self):
        plantilla = self.plantilla_text.get("1.0", tk.END).strip()
        try:
            with open(self.plantilla_file, "w", encoding="utf-8") as f:
                json.dump({"plantilla": plantilla}, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Éxito", "Plantilla guardada correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la plantilla:\n{e}")
    
    def limpiar_plantilla(self):
        self.plantilla_text.delete("1.0", tk.END)
    
    def leer_entero(self, texto, nombre_campo):
        """Convierte un campo numérico opcional en entero (vacío = 0)."""
        texto = texto.strip()
        if not texto:
            return 0
        try:
            valor = int(texto)
        except ValueError:
            raise ValueError(f"{nombre_campo} debe ser un número entero")
        if valor < 0:
            raise ValueError(f"{nombre_campo} no puede ser negativo")
        return valor

    def calcular(self):
        try:
            fecha_entrada = parsear_fecha(self.entrada.get())
        except ValueError:
            messagebox.showerror("Error", "Fecha de entrada inválida (use DD/MM/YYYY)")
            return
        try:
            fecha_salida = parsear_fecha(self.salida.get())
        except ValueError:
            messagebox.showerror("Error", "Fecha de salida inválida (use DD/MM/YYYY)")
            return

        try:
            rebaja = self.leer_entero(self.rebaja.get(), "La rebaja")
            if self.prev_inicio.get().strip() and self.prev_fin.get().strip():
                try:
                    p_ini = parsear_fecha(self.prev_inicio.get())
                    p_fin = parsear_fecha(self.prev_fin.get())
                except ValueError:
                    raise ValueError("Fechas de preventiva inválidas (use DD/MM/YYYY)")
                if p_fin < p_ini:
                    raise ValueError("El fin de la preventiva debe ser posterior al inicio")
                prev_dias = (p_fin - p_ini).days
            else:
                prev_dias = self.leer_entero(self.prev_dias.get(), "Los días de preventiva")

            r = calcular_condena(fecha_entrada, fecha_salida, rebaja, prev_dias, self.hoy)
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return

        # Guardamos los resultados para la plantilla
        self.rebaja_valor = r.rebaja
        self.prev_dias_valor = r.prev_dias
        self.fecha_entrada_dt = r.fecha_entrada
        self.fecha_salida_dt = r.fecha_salida
        self.condena_original_dias = r.condena_original_dias
        self.condena_original_str = convertir_dias(r.condena_original_dias)
        self.fecha_salida_original_con_preventiva = r.salida_original_con_preventiva
        self.fecha_mitad_original = r.mitad
        self.fecha_tercio_original = r.tercio
        self.fecha_dos_tercios_original = r.dos_tercios
        self.fecha_salida_con_rebaja = r.salida_con_rebaja
        self.fecha_mitad_con_rebaja = r.mitad_con_rebaja
        self.fecha_tercio_con_rebaja = r.tercio_con_rebaja
        self.fecha_dos_tercios_con_rebaja = r.dos_tercios_con_rebaja
        self.dias_cumplidos = r.dias_cumplidos
        self.dias_restantes = r.dias_restantes
        self.cumplido_str = convertir_dias(r.dias_cumplidos)
        self.restante_str = convertir_dias(r.dias_restantes)

        d = lambda f: f.strftime('%d/%m/%Y')
        resultado = f"""
=== DATOS DE LA CONDENA ===
Fecha de entrada: {d(r.fecha_entrada)}
Fecha de salida (con preventiva ya descontada): {d(r.fecha_salida)}
Condena efectiva (entrada a salida): {r.condena_efectiva_dias} días -> {convertir_dias(r.condena_efectiva_dias)}

=== PREVENTIVA ===
Días de preventiva: {r.prev_dias}
Condena original (con preventiva): {r.condena_original_dias} días -> {self.condena_original_str}
Fecha de salida si no se descontara preventiva: {d(r.salida_original_con_preventiva)}

=== FECHAS ORIGINALES (sin rebaja) ===
Mitad: {d(r.mitad)}
Tercio: {d(r.tercio)}
Dos tercios: {d(r.dos_tercios)}

=== FECHAS CON REBAJA (restamos {r.rebaja} días) ===
Fecha de salida con rebaja: {d(r.salida_con_rebaja)}
Mitad con rebaja: {d(r.mitad_con_rebaja)}
Tercio con rebaja: {d(r.tercio_con_rebaja)}
Dos tercios con rebaja: {d(r.dos_tercios_con_rebaja)}

=== TIEMPO CUMPLIDO HASTA HOY ({d(self.hoy)}) ===
Días desde entrada hasta hoy: {r.dias_transcurridos}
+ días de rebaja: {r.rebaja} = {r.dias_cumplidos} días cumplidos -> {self.cumplido_str}
Días restantes hasta nueva salida: {r.dias_restantes} -> {self.restante_str}
"""
        self.resultado_calculos.config(state=tk.NORMAL)
        self.resultado_calculos.delete(1.0, tk.END)
        self.resultado_calculos.insert(tk.END, resultado)
        self.resultado_calculos.config(state=tk.DISABLED)

        messagebox.showinfo("Cálculo completado", "Datos actualizados. Puede generar el documento.")

    def generar_documento(self):
        if self.fecha_entrada_dt is None:
            messagebox.showerror("Error", "Primero debe realizar el cálculo.")
            return
        
        plantilla = self.plantilla_text.get("1.0", tk.END).strip()
        if not plantilla:
            messagebox.showerror("Error", "La plantilla está vacía.")
            return
        
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(12)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        if os.path.exists(self.ruta_escudo.get()):
            try:
                doc.add_picture(self.ruta_escudo.get(), width=Inches(2.0))
                doc.add_paragraph()
            except Exception as e:  # imagen dañada o formato no soportado
                messagebox.showwarning("Escudo", f"No se pudo añadir el escudo:\n{e}")
        
        def f(fecha):
            return fecha_larga(fecha) if fecha else ""
        
        valores = {
            "(nombre)": self.nombre.get().strip(),
            "(establecimiento)": self.establecimiento.get().strip(),
            "(ci)": self.carnet.get().strip(),
            "(edadactual)": self.edad_actual.get().strip(),
            "(edaddentrada)": self.edad_entrada.get().strip(),
            "(direccion)": self.direccion.get().strip(),
            "(delito)": self.delito.get().strip(),
            "(tribunal)": self.tribunal.get().strip(),
            "(causa)": self.causa.get().strip(),
            "(inciso)": self.inciso.get().strip(),
            "(escolaridad)": self.escolaridad.get().strip(),
            "(naturalde)": self.natural_de.get().strip(),
            "(estadocivil)": self.estado_civil.get().strip(),
            "(regimen)": self.regimen.get().strip(),
            "(posh)": self.posh.get().strip(),
            "(razon)": self.razon.get().strip(),
            "(clasif)": self.clasificacion.get().strip(),
            "(favorable)": self.favorable.get().strip() or "No especificado",
            "(trabajosocial)": self.trabajo_social.get().strip() or "No especificado",
            "(oficios)": self.oficios.get().strip() or "Ninguno",
            "(respcivil)": self.responsabilidad_civil.get().strip() or "No especificada",
            "(indisciplinas)": self.indisciplinas.get().strip() or "Ninguna",
            "(FechaEntrada)": f(self.fecha_entrada_dt),
            "(FechaSalida)": f(self.fecha_salida_dt),
            "(CondenaOriginalInfo)": self.condena_original_str,
            "(MitadOriginal)": f(self.fecha_mitad_original),
            "(TercioOriginal)": f(self.fecha_tercio_original),
            "(DosTerciosOriginal)": f(self.fecha_dos_tercios_original),
            "(FechaSalidaOriginalConPreventiva)": f(self.fecha_salida_original_con_preventiva),
            "(CondenaOriginalTotal)": self.condena_original_str,
            "(FechaSalidaConRebaja)": f(self.fecha_salida_con_rebaja),
            "(MitadConRebaja)": f(self.fecha_mitad_con_rebaja),
            "(TercioConRebaja)": f(self.fecha_tercio_con_rebaja),
            "(DosTerciosConRebaja)": f(self.fecha_dos_tercios_con_rebaja),
            "(DiasRebaja)": str(self.rebaja_valor),
            "(DiasPreventiva)": str(self.prev_dias_valor),
            "(FechaActual)": f(self.hoy),
        }
        
        rellenar_plantilla(doc, plantilla, valores, MARCADORES_NEGRITA)

        nombre_limpio = self.nombre.get().strip().replace(" ", "_") if self.nombre.get().strip() else "documento"
        sugerencia = f"{nombre_limpio}.docx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx"), ("Todos los archivos", "*.*")],
            title="Guardar documento como",
            initialfile=sugerencia
        )
        if file_path:
            try:
                doc.save(file_path)
                messagebox.showinfo("Éxito", f"Documento guardado en:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PrisionApp(root)
    root.mainloop()