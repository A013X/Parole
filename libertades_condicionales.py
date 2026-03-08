import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime, timedelta
import json
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

meses_espanol = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

def fecha_larga(fecha_dt):
    if not fecha_dt:
        return ""
    dia = fecha_dt.day
    mes = meses_espanol[fecha_dt.month - 1]
    año = fecha_dt.year
    return f"{dia} de {mes} de {año}"

def convertir_dias_mejorado(dias):
    if dias == 0:
        return "0 días"
    signo = '-' if dias < 0 else ''
    dias_abs = abs(dias)
    años = dias_abs // 365
    resto = dias_abs % 365
    meses = resto // 30
    dias_resto = resto % 30
    if meses >= 12:
        años += meses // 12
        meses = meses % 12
    partes = []
    if años > 0:
        partes.append(f"{años} año{'s' if años != 1 else ''}")
    if meses > 0:
        partes.append(f"{meses} mes{'es' if meses != 1 else ''}")
    if dias_resto > 0:
        partes.append(f"{dias_resto} día{'s' if dias_resto != 1 else ''}")
    resultado = ', '.join(partes) if partes else "0 días"
    return signo + resultado

class PrisionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Calculadora de Tiempo en Prisión")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        self.set_style()
        
        self.plantilla_file = "plantilla.json"
        
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
        
        self.ruta_escudo = tk.StringVar(value="escudo.png")
        
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
        
        self.carnet.trace("w", self.calcular_edades)
        self.entrada.trace("w", self.calcular_edades)
        
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
        carnet = self.carnet.get().strip()
        if not (carnet and len(carnet) >= 2 and carnet[:2].isdigit()):
            return
        try:
            dos_digitos = int(carnet[:2])
            año_nacimiento = 2000 + dos_digitos if dos_digitos <= datetime.now().year % 100 else 1900 + dos_digitos
            edad_actual = self.hoy.year - año_nacimiento
            if (self.hoy.month, self.hoy.day) < (1, 1):
                edad_actual -= 1
            self.edad_actual.set(str(edad_actual))
            try:
                fecha_entrada_dt = datetime.strptime(self.entrada.get().strip(), "%d/%m/%Y")
                edad_entrada = fecha_entrada_dt.year - año_nacimiento
                if (fecha_entrada_dt.month, fecha_entrada_dt.day) < (1, 1):
                    edad_entrada -= 1
                self.edad_entrada.set(str(edad_entrada))
            except:
                self.edad_entrada.set("")
        except:
            pass
    
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
            except:
                pass
    
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
    
    def calcular(self):
        try:
            fecha_entrada = datetime.strptime(self.entrada.get().strip(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Fecha de entrada inválida (use DD/MM/YYYY)")
            return
        try:
            fecha_salida = datetime.strptime(self.salida.get().strip(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Fecha de salida inválida (use DD/MM/YYYY)")
            return
        if fecha_salida <= fecha_entrada:
            messagebox.showerror("Error", "La fecha de salida debe ser posterior a la de entrada")
            return
        
        try:
            rebaja = int(self.rebaja.get().strip()) if self.rebaja.get().strip() else 0
        except ValueError:
            messagebox.showerror("Error", "Rebaja debe ser número entero")
            return
        self.rebaja_valor = rebaja
        
        prev_dias = 0
        if self.prev_inicio.get().strip() and self.prev_fin.get().strip():
            try:
                p_ini = datetime.strptime(self.prev_inicio.get().strip(), "%d/%m/%Y")
                p_fin = datetime.strptime(self.prev_fin.get().strip(), "%d/%m/%Y")
                if p_fin >= p_ini:
                    prev_dias = (p_fin - p_ini).days
                else:
                    messagebox.showerror("Error", "Fin de preventiva debe ser posterior a inicio")
                    return
            except ValueError:
                messagebox.showerror("Error", "Fechas de preventiva inválidas")
                return
        elif self.prev_dias.get().strip():
            try:
                prev_dias = int(self.prev_dias.get().strip())
            except ValueError:
                messagebox.showerror("Error", "Días de preventiva deben ser número entero")
                return
        self.prev_dias_valor = prev_dias
        
        condena_efectiva_dias = (fecha_salida - fecha_entrada).days
        mitad_orig = condena_efectiva_dias // 2
        tercio_orig = condena_efectiva_dias // 3
        dos_tercios_orig = condena_efectiva_dias * 2 // 3
        
        self.fecha_mitad_original = fecha_entrada + timedelta(days=mitad_orig)
        self.fecha_tercio_original = fecha_entrada + timedelta(days=tercio_orig)
        self.fecha_dos_tercios_original = fecha_entrada + timedelta(days=dos_tercios_orig)
        
        self.fecha_salida_con_rebaja = fecha_salida - timedelta(days=rebaja) if rebaja > 0 else fecha_salida
        self.fecha_mitad_con_rebaja = self.fecha_mitad_original - timedelta(days=rebaja) if rebaja > 0 else self.fecha_mitad_original
        self.fecha_tercio_con_rebaja = self.fecha_tercio_original - timedelta(days=rebaja) if rebaja > 0 else self.fecha_tercio_original
        self.fecha_dos_tercios_con_rebaja = self.fecha_dos_tercios_original - timedelta(days=rebaja) if rebaja > 0 else self.fecha_dos_tercios_original
        
        self.condena_original_dias = condena_efectiva_dias + prev_dias
        self.condena_original_str = convertir_dias_mejorado(self.condena_original_dias)
        self.fecha_salida_original_con_preventiva = fecha_salida + timedelta(days=prev_dias)
        
        dias_transcurridos = (self.hoy - fecha_entrada).days
        if dias_transcurridos < 0:
            dias_transcurridos = 0
        self.dias_cumplidos = dias_transcurridos + rebaja
        self.cumplido_str = convertir_dias_mejorado(self.dias_cumplidos)
        
        dias_hasta_salida_con_rebaja = (self.fecha_salida_con_rebaja - self.hoy).days
        if dias_hasta_salida_con_rebaja < 0:
            dias_hasta_salida_con_rebaja = 0
        self.dias_restantes = dias_hasta_salida_con_rebaja
        self.restante_str = convertir_dias_mejorado(self.dias_restantes)
        
        self.fecha_entrada_dt = fecha_entrada
        self.fecha_salida_dt = fecha_salida
        
        resultado = f"""
=== DATOS DE LA CONDENA ===
Fecha de entrada: {fecha_entrada.strftime('%d/%m/%Y')}
Fecha de salida (con preventiva ya descontada): {fecha_salida.strftime('%d/%m/%Y')}
Condena efectiva (entrada a salida): {condena_efectiva_dias} días -> {convertir_dias_mejorado(condena_efectiva_dias)}

=== PREVENTIVA ===
Días de preventiva: {prev_dias}
Condena original (con preventiva): {self.condena_original_dias} días -> {self.condena_original_str}
Fecha de salida si no se descontara preventiva: {self.fecha_salida_original_con_preventiva.strftime('%d/%m/%Y')}

=== FECHAS ORIGINALES (sin rebaja) ===
Mitad: {self.fecha_mitad_original.strftime('%d/%m/%Y')}
Tercio: {self.fecha_tercio_original.strftime('%d/%m/%Y')}
Dos tercios: {self.fecha_dos_tercios_original.strftime('%d/%m/%Y')}

=== FECHAS CON REBAJA (restamos {rebaja} días) ===
Fecha de salida con rebaja: {self.fecha_salida_con_rebaja.strftime('%d/%m/%Y')}
Mitad con rebaja: {self.fecha_mitad_con_rebaja.strftime('%d/%m/%Y')}
Tercio con rebaja: {self.fecha_tercio_con_rebaja.strftime('%d/%m/%Y')}
Dos tercios con rebaja: {self.fecha_dos_tercios_con_rebaja.strftime('%d/%m/%Y')}

=== TIEMPO CUMPLIDO HASTA HOY ({self.hoy.strftime('%d/%m/%Y')}) ===
Días desde entrada hasta hoy: {dias_transcurridos}
+ días de rebaja: {rebaja} = {self.dias_cumplidos} días cumplidos -> {self.cumplido_str}
Días restantes hasta nueva salida: {self.dias_restantes} -> {self.restante_str}
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
            except:
                pass
        
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
        
        negrita = ["(nombre)", "(delito)", "(tribunal)", "(causa)", "(favorable)"]
        
        lineas = plantilla.split('\n')
        for i, linea in enumerate(lineas):
            if not linea.strip():
                if i < len(lineas)-1 and lineas[i+1].strip() == "":
                    continue
                doc.add_paragraph()
                continue
            p = doc.add_paragraph()
            resto = linea
            while resto:
                inicio = resto.find('(')
                if inicio == -1:
                    if resto:
                        p.add_run(resto)
                    break
                if inicio > 0:
                    p.add_run(resto[:inicio])
                fin = resto.find(')', inicio)
                if fin == -1:
                    p.add_run(resto[inicio:])
                    break
                marcador = resto[inicio:fin+1]
                valor = valores.get(marcador, marcador)
                run = p.add_run(valor)
                if marcador in negrita:
                    run.bold = True
                resto = resto[fin+1:]
        
        nombre_limpio = self.nombre.get().strip().replace(" ", "_") if self.nombre.get().strip() else "documento"
        sugerencia = f"{nombre_limpio}.doc"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".doc",
            filetypes=[("Documento Word 97-2003", "*.doc"), ("Todos los archivos", "*.*")],
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