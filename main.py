import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
import win32com.client
from typing import Optional, Dict
from dataclasses import dataclass
import re

TEMPLATE_PATH = "plantilla.xlsx"

@dataclass
class UserData:
    """Estructura de datos del usuario"""
    alias: str = ""
    full_name: str = ""
    first_name: str = ""
    second_name: str = ""
    first_surname: str = ""
    second_surname: str = ""
    correo: str = ""
    puesto: str = ""
    empresa: str = ""
    depart: str = ""
    
    # Datos adicionales del formulario
    ticket: str = ""
    pc_type: str = ""
    st: str = ""
    model: str = ""
    form_types: list = None  # Cambiado a lista para múltiples selecciones
    marca: str = ""  # Agregado campo marca
    estado: str = ""
    more_equipment: list = None
    reported_problem: str = ""
    diagnosis: str = ""
    documentation: str = ""
    
    # Nuevos campos
    disco_duro: str = ""
    memoria_ram: str = ""
    hostname: str = ""
    
    def __post_init__(self):
        if self.more_equipment is None:
            self.more_equipment = []
        if self.form_types is None:
            self.form_types = []

class OutlookSearcher:
    """Maneja la búsqueda en Outlook/GAL"""
    
    @staticmethod
    def buscar_usuario(alias: str) -> Optional[Dict[str, str]]:
        """Busca usuario en la Lista Global de Direcciones de Outlook"""
        alias = alias.lower().split("@")[0].strip()
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            gal = ns.AddressLists["Lista global de direcciones"]
            
            for i in range(gal.AddressEntries.Count):
                entry = gal.AddressEntries.Item(i + 1)
                try:
                    ex = entry.GetExchangeUser()
                    correo = ex.PrimarySmtpAddress.lower()
                    
                    if correo.startswith(alias + "@"):
                        name_parts = ex.Name.split()
                        
                        # Separar nombre completo correctamente
                        first_name = ""
                        second_name = ""
                        first_surname = ""
                        second_surname = ""
                        
                        if len(name_parts) >= 1:
                            first_name = name_parts[0]
                        if len(name_parts) >= 3:
                            # Si hay 3 o más partes, los últimos 2 son apellidos
                            first_surname = name_parts[-2]
                            second_surname = name_parts[-1]
                            # Todo lo del medio son segundos nombres
                            if len(name_parts) > 3:
                                second_name = " ".join(name_parts[1:-2])
                        elif len(name_parts) == 2:
                            # Solo 2 partes: nombre y apellido
                            first_surname = name_parts[1]
                        
                        return {
                            "alias": alias,
                            "full_name": ex.Name,
                            "first_name": first_name,
                            "second_name": second_name,
                            "first_surname": first_surname,
                            "second_surname": second_surname,
                            "correo": correo,
                            "puesto": ex.JobTitle or "",
                            "empresa": ex.CompanyName or "",
                            "depart": ex.Department or ""
                        }
                except AttributeError:
                    continue
                    
            return None
            
        except Exception as e:
            messagebox.showerror("Error COM", f"Error al conectar con Outlook:\n{str(e)}")
            return None

class ExcelGenerator:
    """Genera la plantilla de Excel con los datos usando etiquetas"""
    
    FORM_TYPES = [
        "FECHA DE DEVOLUCIÓN DE EQUIPO",
        "HOJA DE RESPONSABILIDAD",
        "SOPORTE TÉCNICO",
        "PASE DE SALIDA",
        "DICTAMEN TÉCNICO",
        "RECEPCIÓN DE EQUIPO",
        "PRÉSTAMO DE EQUIPO"
    ]
    
    def __init__(self, template_path: str = TEMPLATE_PATH):
        self.template_path = template_path
    
    def generar(self, data: UserData) -> bool:
        """Genera la plantilla Excel con los datos proporcionados"""
        
        # Validar datos mínimos
        if not data.ticket or not data.alias:
            messagebox.showwarning(
                "Faltan datos",
                "Debes buscar un usuario y completar al menos el número de ticket."
            )
            return False
        
        try:
            wb = load_workbook(self.template_path)
            ws = wb.active
            
            # Preparar datos para reemplazo
            replacements = self._preparar_datos(data)
            
            # Reemplazar etiquetas en toda la hoja
            self._reemplazar_etiquetas(ws, replacements)
            
            # Marcar checkboxes de los tipos de formulario seleccionados
            if data.form_types:
                for form_type in data.form_types:
                    self._marcar_checkbox(ws, form_type)
            
            # Llenar tabla de equipos adicionales
            if data.more_equipment:
                self._llenar_tabla_equipos(ws, data.more_equipment, data.estado)
            
            # Guardar archivo
            output_file = f"{data.ticket}_{data.alias}.xlsx"
            wb.save(output_file)
            
            messagebox.showinfo(
                "Guardado exitoso",
                f"Plantilla exportada como:\n{output_file}"
            )
            return True
            
        except FileNotFoundError:
            messagebox.showerror(
                "Error",
                f"No se encontró la plantilla: {self.template_path}"
            )
            return False
        except Exception as e:
            messagebox.showerror(
                "Error al generar",
                f"Error al crear la plantilla:\n{str(e)}"
            )
            return False
    
    def _preparar_datos(self, data: UserData) -> Dict[str, str]:
        """Prepara el diccionario de reemplazos según las etiquetas de tu plantilla"""
        return {
            # Etiquetas exactas de tu plantilla
            "ticket": data.ticket,
            "alias": data.alias,
            "first_name": data.first_name,  # Corregido: ahora usa first_name
            "second_name": data.second_name,
            "first_surname": data.first_surname,
            "second_surname": data.second_surname,  # Agregado
            "puesto": data.puesto,
            "depart": data.depart,
            "empresa": data.empresa,
            "pc_type": data.pc_type,
            "st": data.st,
            "model": data.model,
            "marca": data.marca,  # Agregado campo marca
            "disco_duro": data.disco_duro,  # Agregado
            "memoria_ram": data.memoria_ram,  # Agregado
            "hostname": data.hostname,  # Agregado
            "estado": data.estado,  # Agregado: Estado del equipo
            "reported_problem": data.reported_problem,  # No será reemplazado si está vacío
            "diagnosis": data.diagnosis,
            "documentation": data.documentation if data.documentation else "",
        }
    
    def _reemplazar_etiquetas(self, ws, replacements: Dict[str, str]):
        """
        Busca y reemplaza todas las etiquetas {{tag}} en la hoja de Excel.
        Maneja correctamente celdas combinadas escribiendo en la celda principal.
        """
        # Patrón para detectar etiquetas
        pattern = re.compile(r'\{\{(\w+)\}\}')
        
        # Obtener dimensiones de la hoja
        max_row = ws.max_row
        max_col = ws.max_column
        
        # Recorrer todas las celdas por coordenadas
        for row_idx in range(1, max_row + 1):
            for col_idx in range(1, max_col + 1):
                try:
                    cell = ws.cell(row=row_idx, column=col_idx)
                    
                    # Saltar celdas combinadas
                    if isinstance(cell, MergedCell):
                        continue
                    
                    # Solo procesar celdas con contenido de texto
                    if cell.value and isinstance(cell.value, str):
                        original_value = cell.value
                        new_value = original_value
                        
                        # Buscar todas las etiquetas en el contenido de la celda
                        matches = pattern.findall(original_value)
                        
                        for tag in matches:
                            if tag in replacements:
                                # Reemplazar la etiqueta
                                placeholder = f"{{{{{tag}}}}}"
                                replacement_value = str(replacements[tag])
                                
                                # No reemplazar si el valor está vacío y es un campo opcional
                                if replacement_value or tag in ["ticket", "alias", "st"]:
                                    new_value = new_value.replace(placeholder, replacement_value)
                        
                        # Si hubo cambios, escribir de forma segura
                        if new_value != original_value:
                            self._escribir_celda_segura(ws, row_idx, col_idx, new_value)
                
                except Exception as e:
                    # Si hay error en una celda específica, continuar con las demás
                    continue
    
    def _escribir_celda_segura(self, ws, row: int, col: int, value: str):
        """
        Escribe en una celda manejando correctamente celdas combinadas.
        Si la celda está combinada, escribe en la celda superior izquierda del rango.
        """
        try:
            celda = ws.cell(row=row, column=col)
            
            # Si es una celda combinada, buscar la celda principal
            if isinstance(celda, MergedCell):
                # Buscar el rango combinado que contiene esta celda
                for merged_range in ws.merged_cells.ranges:
                    if (merged_range.min_row <= row <= merged_range.max_row and
                        merged_range.min_col <= col <= merged_range.max_col):
                        # Escribir en la celda superior izquierda del rango
                        ws.cell(
                            row=merged_range.min_row,
                            column=merged_range.min_col,
                            value=value
                        )
                        return
            else:
                # Celda normal, escribir directamente
                celda.value = value
        except Exception as e:
            # Si falla, intentar escribir directamente
            try:
                ws.cell(row=row, column=col).value = value
            except:
                pass
    
    def _marcar_checkbox(self, ws, form_type: str):
        """Marca con 'X' el checkbox del tipo de formulario seleccionado"""
        if not form_type:
            return
        
        # Buscar el texto del formulario en las primeras filas de forma segura
        for r in range(5, 12):
            for c in range(1, 15):
                try:
                    cell = ws.cell(row=r, column=c)
                    
                    # Saltar celdas combinadas
                    if isinstance(cell, MergedCell):
                        continue
                    
                    if cell.value and isinstance(cell.value, str):
                        cell_text = cell.value.upper()
                        form_text = form_type.upper()
                        
                        # Verificar coincidencia exacta o parcial
                        if form_text in cell_text or cell_text in form_text:
                            # Intentar marcar en celdas cercanas (izquierda, arriba, o misma celda)
                            # Opción 1: Celda a la izquierda
                            try:
                                checkbox_cell = ws.cell(row=r, column=c-1)
                                if not isinstance(checkbox_cell, MergedCell):
                                    if not checkbox_cell.value or str(checkbox_cell.value).strip() == "":
                                        self._escribir_celda_segura(ws, r, c-1, "X")
                                        return
                            except:
                                pass
                            
                            # Opción 2: Celda dos columnas a la izquierda
                            try:
                                if c >= 2:
                                    checkbox_cell = ws.cell(row=r, column=c-2)
                                    if not isinstance(checkbox_cell, MergedCell):
                                        if not checkbox_cell.value or str(checkbox_cell.value).strip() == "":
                                            self._escribir_celda_segura(ws, r, c-2, "X")
                                            return
                            except:
                                pass
                except Exception as e:
                    continue
    
    def _llenar_tabla_equipos(self, ws, equipos: list, estado: str):
        """
        Llena la tabla de equipos adicionales (sección C del formulario).
        Busca la tabla por sus encabezados y llena las filas de forma segura.
        """
        if not equipos:
            return
            
        # Buscar la fila de encabezados de la tabla de equipos
        tabla_inicio = None
        
        try:
            for row_idx in range(30, 50):
                for col_idx in range(1, 10):
                    try:
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if isinstance(cell, MergedCell):
                            continue
                        if cell.value and isinstance(cell.value, str):
                            if "EQUIPO" in cell.value.upper():
                                # Verificar que sea la fila de encabezados
                                next_cell = ws.cell(row=row_idx, column=col_idx+1)
                                if not isinstance(next_cell, MergedCell) and next_cell.value:
                                    if "MARCA" in str(next_cell.value).upper():
                                        tabla_inicio = row_idx + 1
                                        break
                    except:
                        continue
                if tabla_inicio:
                    break
        except:
            pass
        
        if not tabla_inicio:
            # Usar posición por defecto si no se encuentra
            tabla_inicio = 42
        
        # Llenar equipos de forma segura
        for i, equipo_info in enumerate(equipos):
            row_num = tabla_inicio + i
            try:
                # No (A/1), EQUIPO (B/2), MARCA (C/3), MODELO (D/4), SERIE (E/5), ESTADO (F/6)
                self._escribir_celda_segura(ws, row_num, 1, str(i+1))
                self._escribir_celda_segura(ws, row_num, 2, equipo_info.get("equipo", ""))
                self._escribir_celda_segura(ws, row_num, 3, equipo_info.get("marca", ""))
                self._escribir_celda_segura(ws, row_num, 4, equipo_info.get("modelo", ""))
                self._escribir_celda_segura(ws, row_num, 5, equipo_info.get("serie", ""))
                self._escribir_celda_segura(ws, row_num, 6, estado)
            except Exception as e:
                continue

class FormularioApp:
    """Aplicación principal con interfaz gráfica"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Formulario DISAGRO - Departamento de Sistemas")
        self.root.geometry("900x650")
        self.data = UserData()
        self.excel_gen = ExcelGenerator()
        self.widgets = {}
        
        self._crear_interfaz()
    
    def _crear_interfaz(self):
        """Crea todos los elementos de la interfaz"""
        
        # Frame principal con título
        title_frame = tk.Frame(self.root, bg="#2c3e50", padx=10, pady=10)
        title_frame.grid(row=0, column=0, sticky="ew")
        
        tk.Label(
            title_frame,
            text="SISTEMA DE FORMULARIOS - DEPARTAMENTO DE SISTEMAS",
            font=("Arial", 14, "bold"),
            bg="#2c3e50",
            fg="white"
        ).pack()
        
        # Frame de búsqueda
        search_frame = tk.LabelFrame(self.root, text="1. Buscar Usuario", padx=10, pady=10)
        search_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        
        tk.Label(search_frame, text="Alias o correo:").grid(row=0, column=0, sticky="w")
        self.widgets['alias_entry'] = tk.Entry(search_frame, width=40)
        self.widgets['alias_entry'].grid(row=0, column=1, padx=5)
        
        tk.Button(
            search_frame,
            text="🔍 Buscar en Outlook",
            command=self._buscar_usuario,
            bg="#3498db",
            fg="white",
            padx=10
        ).grid(row=0, column=2)
        
        # TreeView de resultados
        self._crear_treeview()
        
        # Botones de acción
        button_frame = tk.Frame(self.root, padx=10, pady=10)
        button_frame.grid(row=3, column=0)
        
        self.widgets['btn_rellenar'] = tk.Button(
            button_frame,
            text="📝 Rellenar Datos del Formulario",
            command=self._abrir_formulario_datos,
            state="disabled",
            bg="#e67e22",
            fg="white",
            padx=20,
            pady=10,
            font=("Arial", 10, "bold")
        )
        self.widgets['btn_rellenar'].grid(row=0, column=0, padx=5)
        
        tk.Button(
            button_frame,
            text="✅ Generar Plantilla Excel",
            command=self._generar_plantilla,
            bg="#27ae60",
            fg="white",
            padx=20,
            pady=10,
            font=("Arial", 10, "bold")
        ).grid(row=0, column=1, padx=5)
    
    def _crear_treeview(self):
        """Crea el TreeView para mostrar resultados"""
        tree_frame = tk.LabelFrame(self.root, text="2. Usuario Encontrado", padx=10, pady=5)
        tree_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        
        columns = ("Nombre", "Correo", "Puesto", "Empresa", "Departamento")
        self.widgets['tree'] = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            height=3
        )
        
        for col in columns:
            self.widgets['tree'].heading(col, text=col)
            self.widgets['tree'].column(col, width=160)
        
        self.widgets['tree'].pack(fill="both", expand=True)
    
    def _buscar_usuario(self):
        """Busca usuario en Outlook"""
        alias = self.widgets['alias_entry'].get().strip()
        
        if not alias:
            messagebox.showwarning("Campo vacío", "Ingresa un alias o correo.")
            return
        
        # Limpiar resultados anteriores
        tree = self.widgets['tree']
        tree.delete(*tree.get_children())
        
        # Buscar usuario
        resultado = OutlookSearcher.buscar_usuario(alias)
        
        if resultado:
            # Actualizar datos
            for key, value in resultado.items():
                setattr(self.data, key, value)
            
            # Mostrar en TreeView
            tree.insert("", "end", values=(
                resultado["full_name"],
                resultado["correo"],
                resultado["puesto"],
                resultado["empresa"],
                resultado["depart"]
            ))
            
            # Habilitar botón de rellenar
            self.widgets['btn_rellenar'].config(state="normal")
            messagebox.showinfo("✓ Usuario encontrado", f"Usuario {resultado['full_name']} cargado correctamente.")
        else:
            messagebox.showwarning(
                "No encontrado",
                f"No se encontró el usuario: {alias}\nVerifica el alias o correo."
            )
    
    def _abrir_formulario_datos(self):
        """Abre ventana para capturar datos adicionales"""
        
        if not self.data.alias:
            messagebox.showwarning(
                "Busca primero",
                "Primero debes buscar y encontrar un usuario."
            )
            return
        
        win = tk.Toplevel(self.root)
        win.title(f"Datos del Formulario - {self.data.full_name}")
        win.geometry("700x800")
        
        # Contenedor con scroll
        canvas = tk.Canvas(win)
        scrollbar = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # ===== SECCIÓN A: DATOS BÁSICOS =====
        tk.Label(scrollable_frame, text="A. DATOS BÁSICOS", font=("Arial", 11, "bold"), bg="#34495e", fg="white").grid(
            row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        entries = {}
        row = 1
        
        # Campos simples
        simple_fields = [
            ("No. de Ticket *", "ticket"),
            ("Tipo de Computadora", "pc_type"),
            ("Service Tag (ST) *", "st"),
            ("Modelo", "model"),
            ("Marca *", "marca"),  # AGREGADO: Campo Marca
            ("Disco Duro", "disco_duro"),  # AGREGADO
            ("Memoria RAM", "memoria_ram"),  # AGREGADO
            ("Hostname", "hostname"),  # AGREGADO
        ]
        
        for label, key in simple_fields:
            tk.Label(scrollable_frame, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            entry = tk.Entry(scrollable_frame, width=50)
            entry.grid(row=row, column=1, sticky="w", padx=5, pady=5)
            entry.insert(0, getattr(self.data, key, ""))
            entries[key] = entry
            row += 1
        
        # MODIFICADO: Tipo de formulario con selección múltiple usando Checkbuttons
        tk.Label(scrollable_frame, text="Tipo de formulario *").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        
        form_frame = tk.Frame(scrollable_frame)
        form_frame.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        
        form_vars = {}
        for i, form_type in enumerate(ExcelGenerator.FORM_TYPES):
            var = tk.BooleanVar()
            if form_type in self.data.form_types:
                var.set(True)
            cb = tk.Checkbutton(form_frame, text=form_type, variable=var)
            cb.grid(row=i, column=0, sticky="w")
            form_vars[form_type] = var
        
        entries['form_types'] = form_vars
        row += 1
        
        # Estado
        tk.Label(scrollable_frame, text="Estado del equipo *").grid(row=row, column=0, sticky="e", padx=5, pady=5)
        combo_estado = ttk.Combobox(
            scrollable_frame,
            values=["OK", "MALO"],
            state="readonly",
            width=47
        )
        combo_estado.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        if self.data.estado:
            combo_estado.set(self.data.estado)
        entries['estado'] = combo_estado
        row += 1
        
        # ===== SECCIÓN C: EQUIPOS ADICIONALES =====
        tk.Label(scrollable_frame, text="C. HARDWARE ADICIONAL", font=("Arial", 11, "bold"), bg="#34495e", fg="white").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(15, 10))
        row += 1
        
        # Frame para lista de equipos
        equipos_frame = tk.Frame(scrollable_frame, relief="sunken", bd=1)
        equipos_frame.grid(row=row, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        
        equipos_list = tk.Listbox(equipos_frame, height=5, width=70)
        equipos_list.pack(side="left", fill="both", expand=True)
        
        equipos_scroll = tk.Scrollbar(equipos_frame, command=equipos_list.yview)
        equipos_scroll.pack(side="right", fill="y")
        equipos_list.config(yscrollcommand=equipos_scroll.set)
        
        # Cargar equipos existentes
        if self.data.more_equipment:
            for eq in self.data.more_equipment:
                equipos_list.insert("end", f"{eq.get('equipo', '')} - {eq.get('marca', '')} {eq.get('modelo', '')}")
        
        # Botones para gestionar equipos
        btn_frame = tk.Frame(scrollable_frame)
        btn_frame.grid(row=row+1, column=0, columnspan=2)
        
        def agregar_equipo():
            eq_win = tk.Toplevel(win)
            eq_win.title("Agregar Equipo")
            eq_win.geometry("400x250")
            
            eq_entries = {}
            labels = ["Equipo:", "Marca:", "Modelo:", "Serie:"]
            keys = ["equipo", "marca", "modelo", "serie"]
            
            for i, (lbl, key) in enumerate(zip(labels, keys)):
                tk.Label(eq_win, text=lbl).grid(row=i, column=0, sticky="e", padx=5, pady=5)
                e = tk.Entry(eq_win, width=30)
                e.grid(row=i, column=1, padx=5, pady=5)
                eq_entries[key] = e
            
            def guardar_equipo():
                equipo_data = {k: v.get().strip() for k, v in eq_entries.items()}
                if equipo_data["equipo"]:
                    equipos_list.insert("end", f"{equipo_data['equipo']} - {equipo_data['marca']} {equipo_data['modelo']}")
                    if not hasattr(self.data, 'temp_equipos'):
                        self.data.temp_equipos = []
                    self.data.temp_equipos.append(equipo_data)
                    eq_win.destroy()
            
            tk.Button(eq_win, text="Guardar", command=guardar_equipo, bg="#27ae60", fg="white").grid(
                row=len(labels), column=0, columnspan=2, pady=10)
        
        def eliminar_equipo():
            sel = equipos_list.curselection()
            if sel:
                idx = sel[0]
                equipos_list.delete(idx)
                if hasattr(self.data, 'temp_equipos'):
                    self.data.temp_equipos.pop(idx)
        
        tk.Button(btn_frame, text="➕ Agregar Equipo", command=agregar_equipo, bg="#3498db", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="❌ Eliminar", command=eliminar_equipo, bg="#e74c3c", fg="white").pack(side="left", padx=5)
        
        row += 2
        
        # ===== SECCIÓN E: PROBLEMA Y DIAGNÓSTICO =====
        tk.Label(scrollable_frame, text="E. PROBLEMA REPORTADO", font=("Arial", 11, "bold"), bg="#34495e", fg="white").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(15, 5))
        row += 1
        
        tk.Label(scrollable_frame, text="Problema reportado:").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        txt_problem = tk.Text(scrollable_frame, width=50, height=4)
        txt_problem.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        if self.data.reported_problem:
            txt_problem.insert("1.0", self.data.reported_problem)
        entries['reported_problem'] = txt_problem
        row += 1
        
        tk.Label(scrollable_frame, text="F. DIAGNÓSTICO", font=("Arial", 11, "bold"), bg="#34495e", fg="white").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(15, 5))
        row += 1
        
        tk.Label(scrollable_frame, text="Diagnóstico:").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        txt_diag = tk.Text(scrollable_frame, width=50, height=4)
        txt_diag.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        if self.data.diagnosis:
            txt_diag.insert("1.0", self.data.diagnosis)
        entries['diagnosis'] = txt_diag
        row += 1
        
        tk.Label(scrollable_frame, text="Documentación/Repuestos:").grid(row=row, column=0, sticky="ne", padx=5, pady=5)
        txt_doc = tk.Text(scrollable_frame, width=50, height=3)
        txt_doc.grid(row=row, column=1, sticky="w", padx=5, pady=5)
        if self.data.documentation:
            txt_doc.insert("1.0", self.data.documentation)
        entries['documentation'] = txt_doc
        row += 1
        
        # Botón de guardar
        def guardar_datos():
            # Guardar campos de texto y Entry normales
            for key, widget in entries.items():
                if key == 'form_types':
                    # Procesar checkboxes de tipo de formulario
                    selected_forms = [form_type for form_type, var in widget.items() if var.get()]
                    self.data.form_types = selected_forms
                elif isinstance(widget, tk.Text):
                    value = widget.get("1.0", "end").strip()
                    setattr(self.data, key, value)
                else:
                    value = widget.get().strip()
                    setattr(self.data, key, value)
            
            # Guardar equipos
            if hasattr(self.data, 'temp_equipos'):
                self.data.more_equipment = self.data.temp_equipos
            
            win.destroy()
            
            messagebox.showinfo(
                "✓ Datos guardados",
                "Datos capturados correctamente.\nYa puedes generar la plantilla Excel."
            )
        
        tk.Button(
            scrollable_frame,
            text="💾 Guardar Todos los Datos",
            command=guardar_datos,
            bg="#27ae60",
            fg="white",
            padx=30,
            pady=15,
            font=("Arial", 11, "bold")
        ).grid(row=row, column=0, columnspan=2, pady=20)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _generar_plantilla(self):
        """Genera la plantilla Excel"""
        if not self.data.ticket:
            messagebox.showwarning(
                "Faltan datos",
                "Debes completar los datos del formulario primero."
            )
            return
        
        self.excel_gen.generar(self.data)

def main():
    root = tk.Tk()
    app = FormularioApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()