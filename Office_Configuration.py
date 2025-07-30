
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import subprocess
import sys
import ctypes
import tempfile
import threading
import time
import math
from tkinter import font as tkfont
import itertools



def resource_path(relative_path):
    """Obtiene la ruta al recurso, compatible con PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class LoadingWindow:
    def __init__(self, parent, title="Instalando Office"):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("450x280")
        self.top.configure(bg="#0d1117")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()
        
        # Centrar la ventana
        self.center_window()

        # Marco principal con bordes redondeados simulados
        main_frame = tk.Frame(self.top, bg="#0d1117", padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título con mejor tipografía
        self.label_var = tk.StringVar(value=title)
        self.label = tk.Label(
            main_frame, 
            textvariable=self.label_var, 
            bg="#0d1117", 
            fg="#58a6ff",
            font=("Segoe UI", 16, "bold")
        )
        self.label.pack(pady=(0, 25))

        # Canvas de la onda con mejor diseño
        self.canvas = tk.Canvas(
            main_frame, 
            width=380, 
            height=120, 
            bg="#0d1117", 
            highlightthickness=0
        )
        self.canvas.pack()
        
        # Crear gradiente de fondo
        self.create_gradient_bg()
        
        self.wave_line = self.canvas.create_line(
            0, 60, 380, 60, 
            fill="#58a6ff", 
            width=4, 
            smooth=True
        )

        # Variables de animación mejoradas
        self.frame = 0
        self.running = True
        self.color_cycle = itertools.cycle([
            "#58a6ff", "#7c3aed", "#06d6a0", "#ffd60a", "#ff006e"
        ])
        self.current_color = next(self.color_cycle)
        self.dots = 0

        # Iniciar animaciones
        self.animate_wave()
        self.animate_color()
        self.animate_text()

    def center_window(self):
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.top.winfo_screenheight() // 2) - (280 // 2)
        self.top.geometry(f"450x280+{x}+{y}")

    def create_gradient_bg(self):
        # Simular gradiente con rectángulos
        for i in range(120):
            color_intensity = int(13 + (i / 120) * 10)
            color = f"#{color_intensity:02x}{color_intensity+2:02x}{color_intensity+5:02x}"
            self.canvas.create_line(0, i, 380, i, fill=color, width=1)

    def animate_wave(self):
        if not self.running:
            return
        points = []
        for x in range(0, 381, 8):
            wave1 = 25 * math.sin((x + self.frame) * 0.04)
            wave2 = 15 * math.sin((x + self.frame) * 0.06 + math.pi/3)
            y = 60 + wave1 + wave2
            points.extend([x, y])
        
        if len(points) >= 4:
            self.canvas.coords(self.wave_line, *points)
            self.canvas.itemconfig(self.wave_line, fill=self.current_color)
        
        self.frame += 6
        self.top.after(25, self.animate_wave)

    def animate_color(self):
        if not self.running:
            return
        self.current_color = next(self.color_cycle)
        self.top.after(1200, self.animate_color)

    def animate_text(self):
        if not self.running:
            return
        dots = "●" * (self.dots % 4)
        spaces = "  " * (3 - (self.dots % 4))
        self.label_var.set(f"Instalando Office{dots}{spaces}")
        self.dots += 1
        self.top.after(400, self.animate_text)

    def close(self):
        self.running = False
        self.top.destroy()



class OfficeInstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Instalador Office LTSC Pro")
        
        # Configurar ventana responsiva
        self.setup_responsive_window()
        
        # Configurar estilo moderno
        self.setup_modern_styles()
        
        # Variables para almacenar las selecciones
        self.version_var = tk.StringVar(value="2021")
        self.architecture_var = tk.StringVar(value="64")
        self.visio_var = tk.BooleanVar(value=False)
        self.project_var = tk.BooleanVar(value=False)
        self.exclude_lync_var = tk.BooleanVar(value=True)
        self.language_var = tk.StringVar(value="en-us")
        
        self.is_loading = False
        self.loading_label = None
        self.loading_thread = None
        
        self.create_responsive_widgets()
    
    def setup_responsive_window(self):
        # Obtener dimensiones de la pantalla
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Calcular tamaño de ventana basado en la pantalla
        if screen_width < 1024 or screen_height < 768:
            # Pantallas pequeñas
            window_width = min(800, screen_width - 100)
            window_height = min(700, screen_height - 100)
        else:
            # Pantallas normales/grandes
            window_width = 900
            window_height = 800
        
        # Centrar la ventana
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(650, 500)
        self.root.configure(bg="#0d1117")
    
    def setup_modern_styles(self):
        style = ttk.Style()
        
        # Tema moderno GitHub-inspired - cargar solo si existe
        try:
            if os.path.exists('azure.tcl'):
                self.root.tk.call('source', 'azure.tcl')
                self.root.tk.call('set_theme', 'dark')
        except Exception:
            # Si falla, usar tema por defecto con colores personalizados
            pass
        
        # Fuentes con fallback
        try:
            self.title_font = tkfont.Font(family="Segoe UI", size=18, weight="bold")
            self.section_font = tkfont.Font(family="Segoe UI", size=11, weight="bold")
            self.normal_font = tkfont.Font(family="Segoe UI", size=10)
            self.small_font = tkfont.Font(family="Segoe UI", size=9)
        except:
            # Fallback para sistemas sin Segoe UI
            self.title_font = tkfont.Font(family="Arial", size=18, weight="bold")
            self.section_font = tkfont.Font(family="Arial", size=11, weight="bold")
            self.normal_font = tkfont.Font(family="Arial", size=10)
            self.small_font = tkfont.Font(family="Arial", size=9)
        
        # Colores modernos
        self.colors = {
            'bg': '#0d1117',
            'secondary_bg': '#161b22',
            'card_bg': '#21262d',
            'border': '#30363d',
            'text': '#f0f6fc',
            'secondary_text': '#8b949e',
            'accent': '#58a6ff',
            'success': '#3fb950',
            'warning': '#d29922',
            'danger': '#f85149'
        }
        
        # Configurar estilos TTK básicos y compatibles
        try:
            style.configure('Custom.TFrame', background=self.colors['bg'])
            style.configure('Card.TFrame', background=self.colors['card_bg'], relief='solid', borderwidth=1)
            style.configure('Custom.TLabel', background=self.colors['bg'], foreground=self.colors['text'])
            style.configure('Title.TLabel', background=self.colors['bg'], foreground=self.colors['accent'])
            style.configure('Custom.TButton', font=self.normal_font)
            style.configure('Accent.TButton', font=self.section_font)
            style.configure('Custom.TRadiobutton', background=self.colors['card_bg'], foreground=self.colors['text'])
            style.configure('Custom.TCheckbutton', background=self.colors['card_bg'], foreground=self.colors['text'])
            
            # Configurar LabelFrame más simple y compatible
            style.configure('Custom.TLabelFrame', 
                          background=self.colors['card_bg'], 
                          foreground=self.colors['text'],
                          borderwidth=2, 
                          relief='solid')
            style.configure('Custom.TLabelFrame.Label', 
                          background=self.colors['card_bg'], 
                          foreground=self.colors['accent'])
        except Exception as e:
            print(f"Error configurando estilos: {e}")
            # Si falla, continuamos sin estilos personalizados
    
    def create_responsive_widgets(self):
        # Contenedor principal con scroll
        self.main_canvas = tk.Canvas(self.root, bg=self.colors['bg'], highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.main_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.main_canvas, style='Custom.TFrame')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        )
        
        self.canvas_window = self.main_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Bind para redimensionar
        self.main_canvas.bind('<Configure>', self.on_canvas_configure)
        self.root.bind_all("<MouseWheel>", self.on_mousewheel)
        
        # Empaquetar canvas y scrollbar
        self.main_canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Contenido dentro del frame scrollable
        content_frame = ttk.Frame(self.scrollable_frame, style='Custom.TFrame', padding=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header moderno
        self.create_header(content_frame)
        
        # Configuración en cards
        self.create_config_cards(content_frame)
        
        # Log area mejorada
        self.create_log_area(content_frame)
        
        # Botones modernos
        self.create_action_buttons(content_frame)
    
    def on_canvas_configure(self, event):
        canvas_width = event.width
        self.main_canvas.itemconfig(self.canvas_window, width=canvas_width)
    
    def on_mousewheel(self, event):
        self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def create_header(self, parent):
        header_frame = ttk.Frame(parent, style='Custom.TFrame', padding=(0, 0, 0, 30))
        header_frame.pack(fill=tk.X)
        
        # Título principal
        title_label = ttk.Label(
            header_frame,
            text="🏢 Office LTSC Professional",
            style='Title.TLabel',
            font=self.title_font
        )
        title_label.pack()
        
        # Subtítulo
        subtitle_label = ttk.Label(
            header_frame,
            text="Instalador y activador completo para Microsoft Office",
            style='Custom.TLabel',
            font=self.small_font
        )
        subtitle_label.pack(pady=(5, 0))
    
    def create_config_cards(self, parent):
        # Contenedor de cards en grid responsivo
        cards_container = ttk.Frame(parent, style='Custom.TFrame')
        cards_container.pack(fill=tk.X, pady=(0, 20))
        
        # Configurar grid responsivo
        cards_container.columnconfigure(0, weight=1)
        cards_container.columnconfigure(1, weight=1)
        
        # Card 1: Versión y Arquitectura
        version_card = self.create_card(cards_container, "⚙️ Configuración Principal")
        version_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=5)
        
        # Versiones en grid compacto
        versions_frame = ttk.Frame(version_card, style='Card.TFrame')
        versions_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(versions_frame, text="Versión:", style='Custom.TLabel', font=self.section_font).grid(row=0, column=0, sticky="w", padx=5)
        
        versions = [("2019", "2019"), ("2021", "2021"), ("365", "365")]
        for i, (text, val) in enumerate(versions):
            ttk.Radiobutton(
                versions_frame, text=text, variable=self.version_var, value=val, style='Custom.TRadiobutton'
            ).grid(row=0, column=i+1, padx=10, pady=2)
        
        # Arquitectura
        arch_frame = ttk.Frame(version_card, style='Card.TFrame')
        arch_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(arch_frame, text="Arquitectura:", style='Custom.TLabel', font=self.section_font).grid(row=0, column=0, sticky="w", padx=5)
        ttk.Radiobutton(
            arch_frame, text="64-bit", variable=self.architecture_var, value="64", style='Custom.TRadiobutton'
        ).grid(row=0, column=1, padx=10)
        ttk.Radiobutton(
            arch_frame, text="32-bit", variable=self.architecture_var, value="32", style='Custom.TRadiobutton'
        ).grid(row=0, column=2, padx=10)
        
        # Card 2: Componentes
        components_card = self.create_card(cards_container, "📦 Componentes Adicionales")
        components_card.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=5)
        
        components_list = [
            ("✏️ Visio Professional", self.visio_var),
            ("📊 Project Professional", self.project_var),
            ("🚫 Excluir Skype for Business", self.exclude_lync_var)
        ]
        
        for i, (text, var) in enumerate(components_list):
            ttk.Checkbutton(
                components_card, text=text, variable=var, style='Custom.TCheckbutton'
            ).pack(anchor="w", padx=10, pady=5)
        
        # Card 3: Idioma (span completo)
        language_card = self.create_card(cards_container, "🌐 Idioma de Instalación")
        language_card.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(15, 5))
        
        languages_frame = ttk.Frame(language_card, style='Card.TFrame')
        languages_frame.pack(fill=tk.X, pady=5)
        
        languages = [
            ("🇺🇸 English", "en-us"), ("🇪🇸 Español", "es-es"), ("🇫🇷 Français", "fr-fr"),
            ("🇩🇪 Deutsch", "de-de"), ("🇧🇷 Português", "pt-br")
        ]
        
        for i, (text, lang) in enumerate(languages):
            row = i // 3
            col = i % 3
            ttk.Radiobutton(
                languages_frame, text=text, variable=self.language_var, value=lang, style='Custom.TRadiobutton'
            ).grid(row=row, column=col, sticky="w", padx=15, pady=3)
    
    def create_card(self, parent, title):
        # Frame de la card con estilo compatible
        try:
            card_frame = ttk.LabelFrame(parent, text=title, style='Custom.TLabelFrame', padding=15)
        except:
            # Fallback sin estilo personalizado si hay problemas
            card_frame = ttk.LabelFrame(parent, text=title, padding=15)
            try:
                card_frame.configure(background=self.colors['card_bg'], foreground=self.colors['text'])
            except:
                pass
        return card_frame
    
    def create_log_area(self, parent):
        log_card = self.create_card(parent, "📋 Registro de Actividad")
        log_card.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        # Área de texto con mejor diseño
        try:
            log_font = ('JetBrains Mono', 9)
        except:
            log_font = ('Courier New', 9)
            
        self.log_text = scrolledtext.ScrolledText(
            log_card,
            width=70,
            height=10,
            state='disabled',
            bg=self.colors['bg'],
            fg=self.colors['text'],
            insertbackground=self.colors['accent'],
            selectbackground='#1f6feb',
            font=log_font,
            wrap=tk.WORD,
            borderwidth=0,
            highlightthickness=1,
            highlightcolor=self.colors['border']
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Barra de estado moderna
        status_frame = ttk.Frame(log_card, style='Card.TFrame')
        status_frame.pack(fill=tk.X, pady=(5, 0))
        
        # Estado principal
        self.status_label = ttk.Label(
            status_frame,
            text="💡 Listo para instalar Office LTSC",
            style='Custom.TLabel',
            font=self.small_font
        )
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # Indicador de permisos de administrador
        admin_status = "🔐 Admin" if self.is_admin() else "👤 Usuario"
        admin_color = self.colors['success'] if self.is_admin() else self.colors['warning']
        
        self.admin_label = ttk.Label(
            status_frame,
            text=admin_status,
            style='Custom.TLabel',
            font=self.small_font
        )
        self.admin_label.pack(side=tk.RIGHT, padx=5)
        try:
            self.admin_label.configure(foreground=admin_color)
        except:
            pass
    
    def create_action_buttons(self, parent):
        # Contenedor de botones centrado
        button_container = ttk.Frame(parent, style='Custom.TFrame')
        button_container.pack(pady=30)
        
        # Frame interno para centrar
        button_frame = ttk.Frame(button_container, style='Custom.TFrame')
        button_frame.pack()
        
        # Botón de instalación principal
        self.install_button = ttk.Button(
            button_frame,
            text="🚀 Instalar Office",
            command=self.start_installation,
            style='Accent.TButton'
        )
        self.install_button.pack(side=tk.LEFT, padx=10)
        
        # Botón de activación
        self.activate_button = ttk.Button(
            button_frame,
            text="🔐 Activar Office",
            command=self.start_activation,
            style='Accent.TButton'
        )
        self.activate_button.pack(side=tk.LEFT, padx=10)
        
        # Botón de limpiar log
        clear_button = ttk.Button(
            button_frame,
            text="🧹 Limpiar Log",
            command=self.clear_log,
            style='Custom.TButton'
        )
        clear_button.pack(side=tk.LEFT, padx=10)
    
    def log_message(self, message):
        self.log_text.configure(state='normal')
        # Agregar timestamp y formateo
        timestamp = time.strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text.insert(tk.END, formatted_message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.update_status(message)
        self.root.update()
    
    def update_status(self, message):
        # Limpiar mensaje para estado
        clean_message = message.replace("✓", "").replace("✗", "").replace("▶", "").strip()
        display_message = f"💭 {clean_message[:60]}..." if len(clean_message) > 60 else f"💭 {clean_message}"
        self.status_label.config(text=display_message)
    
    def clear_log(self):
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        self.update_status("Listo para instalar Office LTSC")
    
    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def run_as_admin(self):
        if not self.is_admin():
            try:
                # Usar pythonw.exe para ejecutar sin mostrar terminal
                python_exe = sys.executable
                if python_exe.endswith('python.exe'):
                    pythonw_exe = python_exe.replace('python.exe', 'pythonw.exe')
                    if os.path.exists(pythonw_exe):
                        python_exe = pythonw_exe
                
                # Crear argumentos para el reinicio
                args = ' '.join([f'"{arg}"' for arg in sys.argv])
                
                # Ejecutar directamente sin script temporal
                result = ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", python_exe, args, None, 1
                )
                
                if result > 32:  # Éxito
                    # Cerrar el programa actual de manera ordenada
                    self.root.after(500, self._close_program)
                    return True
                else:
                    # Error al iniciar como admin
                    messagebox.showerror(
                        "Error", 
                        "No se pudo obtener permisos de administrador.\n"
                        "El programa continuará sin permisos elevados."
                    )
                    return False
                    
            except Exception as e:
                messagebox.showerror(
                    "Error", 
                    f"Error al solicitar permisos de administrador: {str(e)}"
                )
                return False
        return True
    
    def _close_program(self):
        """Cerrar el programa de manera ordenada"""
        try:
            self.root.quit()
            self.root.destroy()
        except:
            pass
        finally:
            os._exit(0)
    
    def generate_config_xml(self):
        version = self.version_var.get()
        architecture = self.architecture_var.get()
        language = self.language_var.get()
        
        config_file = "configuration.xml"
        
        xml_content = f"""<Configuration>
    <Add OfficeClientEdition="{architecture}" Channel="PerpetualVL{version if version in ['2019', '2021'] else ''}">
        <Product ID="ProPlus{version if version in ['2019', '2021'] else ''}Volume">
            <Language ID="{language}"/>"""
        
        if self.exclude_lync_var.get():
            xml_content += "\n            <ExcludeApp ID=\"Lync\"/>"
        
        xml_content += "\n        </Product>"
        
        if self.visio_var.get():
            xml_content += f"""
        <Product ID="VisioPro{version if version in ['2019', '2021'] else ''}Volume">
            <Language ID="{language}"/>
        </Product>"""
        
        if self.project_var.get():
            xml_content += f"""
        <Product ID="ProjectPro{version if version in ['2019', '2021'] else ''}Volume">
            <Language ID="{language}"/>
        </Product>"""
        
        xml_content += """
    </Add>
    <Remove All="True"/>
    <Display Level="None" AcceptEULA="TRUE" />
    <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>"""
        
        return config_file, xml_content
    
    def start_loading_animation(self, action):
        self.is_loading = True
        dots = 0
        base_text = f"{action}..."
        while self.is_loading:
            text = base_text + "." * dots
            self.update_status(text)
            dots = (dots + 1) % 4
            time.sleep(0.5)
    
    def stop_loading_animation(self):
        self.is_loading = False
        self.update_status("Operación completada")
    
    def toggle_buttons(self, state):
        state = "normal" if state else "disabled"
        self.install_button.config(state=state)
        self.activate_button.config(state=state)
    
    def start_installation(self):
        if not self.is_admin():
            result = messagebox.askyesno(
                "🔐 Permisos de Administrador", 
                "Este programa necesita permisos de administrador para instalar Office.\n\n"
                "¿Deseas reiniciar el programa con permisos de administrador?\n\n"
                "NOTA: El programa se cerrará y se abrirá una nueva ventana con permisos de administrador."
            )
            if result:
                self.log_message("🔄 Reiniciando con permisos de administrador...")
                self.log_message("⏳ Cerrando programa actual...")
                
                # Deshabilitar botones y mostrar estado
                self.toggle_buttons(False)
                self.update_status("Reiniciando con permisos de administrador...")
                
                # Ejecutar después de un pequeño delay para que el usuario vea el mensaje
                self.root.after(1000, lambda: self.run_as_admin())
                return
            else:
                self.log_message("❌ Instalación cancelada - Se requieren permisos de administrador")
                return
        
        # Esta parte solo se ejecuta si ya tenemos permisos de admin
        self.clear_log()
        self.log_message("🔧 Preparando instalación de Office LTSC...")
        self.toggle_buttons(False)

        # Mostrar ventana de carga mejorada
        self.loading_popup = LoadingWindow(self.root, title="Instalando Office LTSC")
        
        threading.Thread(target=self.install_office, daemon=True).start()
    
    def install_office(self):
        try:
            config_file, xml_content = self.generate_config_xml()
            
            with open(config_file, "w", encoding="utf-8") as f:
                f.write(xml_content)
            
            self.log_message(f"✅ Archivo de configuración generado: {config_file}")
            
            setup_path = os.path.join(os.path.dirname(sys.argv[0]), "setup.exe")
            if not os.path.exists(setup_path):
                setup_path = "setup.exe"
                if not os.path.exists(setup_path):
                    self.log_message("❌ Error: No se encontró setup.exe")
                    self.stop_loading_animation()
                    self.toggle_buttons(True)
                    return
            
            self.log_message("🚀 Iniciando instalación de Office LTSC...")
            
            process = subprocess.Popen(
                [setup_path, "/configure", config_file],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    self.log_message(f"📄 {output.strip()}")
            
            return_code = process.poll()
            self.stop_loading_animation()
            self.toggle_buttons(True)
            
            if return_code == 0:
                self.log_message("🎉 ¡Instalación completada con éxito!")
                messagebox.showinfo("✅ Éxito", "Office LTSC se instaló correctamente.")
            else:
                error_msg = process.stderr.read()
                self.log_message(f"❌ Error durante la instalación (Código {return_code}): {error_msg}")
                messagebox.showerror("❌ Error", f"Error durante la instalación:\n{error_msg}")
                
        except Exception as e:
            self.stop_loading_animation()
            self.toggle_buttons(True)
            self.log_message(f"💥 Error inesperado: {str(e)}")
            messagebox.showerror("💥 Error", f"Error inesperado:\n{str(e)}")

        finally:
            if hasattr(self, "loading_popup"):
                self.loading_popup.close()
    
    def start_activation(self):
        if not self.is_admin():
            result = messagebox.askyesno(
                "🔐 Permisos de Administrador", 
                "Este programa necesita permisos de administrador para activar Office.\n\n"
                "¿Deseas reiniciar el programa con permisos de administrador?\n\n"
                "NOTA: El programa se cerrará y se abrirá una nueva ventana con permisos de administrador."
            )
            if result:
                self.log_message("🔄 Reiniciando con permisos de administrador...")
                self.log_message("⏳ Cerrando programa actual...")
                
                # Deshabilitar botones y mostrar estado
                self.toggle_buttons(False)
                self.update_status("Reiniciando con permisos de administrador...")
                
                # Ejecutar después de un pequeño delay para que el usuario vea el mensaje
                self.root.after(1000, lambda: self.run_as_admin())
                return
            else:
                self.log_message("❌ Activación cancelada - Se requieren permisos de administrador")
                return
        
        # Esta parte solo se ejecuta si ya tenemos permisos de admin
        self.clear_log()
        self.log_message("🔐 Preparando activación de Office...")
        self.toggle_buttons(False)
        self.loading_thread = threading.Thread(target=self.start_loading_animation, args=("Activando Office",), daemon=True)
        self.loading_thread.start()
        threading.Thread(target=self.activate_office, daemon=True).start()
    
    def activate_office(self):
        try:
            version = self.version_var.get()
            
            kms_keys = {
                "2019": "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP",
                "2021": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH",
                "365": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
            }
            
            kms_key = kms_keys.get(version, kms_keys["2021"])
            
            batch_script = f"""@echo off
:: Buscar la ruta de Office
for /f "tokens=*" %%a in ('dir /b /s "%ProgramFiles%\\Microsoft Office\\Office16\\ospp.vbs" 2^>nul') do (
    set "office_path=%%~dpa"
    goto :found_office
)
for /f "tokens=*" %%a in ('dir /b /s "%ProgramFiles(x86)%\\Microsoft Office\\Office16\\ospp.vbs" 2^>nul') do (
    set "office_path=%%~dpa"
    goto :found_office
)
echo No se encontró la instalación de Office
exit /b 1

:found_office
echo Ruta de Office encontrada: %office_path%
cd /d "%office_path%"

:: Instalar licencia
for /f %%x in ('dir /b "..\\root\\Licenses16\\ProPlus{version}VL_KMS*.xrm-ms"') do (
    echo Instalando licencia: %%x
    cscript ospp.vbs /inslic:"..\\root\\Licenses16\\%%x"
)

:: Configurar KMS
echo Configurando puerto KMS...
cscript ospp.vbs /setprt:1688

echo Eliminando clave anterior...
cscript ospp.vbs /unpkey:6F7TH >nul

echo Instalando nueva clave...
cscript ospp.vbs /inpkey:{kms_key}

echo Configurando servidor KMS...
cscript ospp.vbs /sethst:e8.us.to

echo Intentando activación...
cscript ospp.vbs /act

echo Proceso de activación completado
"""
            temp_dir = tempfile.gettempdir()
            batch_path = os.path.join(temp_dir, "activate_office.bat")
            
            with open(batch_path, "w", encoding="utf-8") as f:
                f.write(batch_script)
            
            self.log_message("🔧 Ejecutando script de activación...")
            
            process = subprocess.Popen(
                ['cmd.exe', '/c', batch_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    self.log_message(f"📄 {output.strip()}")
            
            return_code = process.poll()
            self.stop_loading_animation()
            self.toggle_buttons(True)
            
            if return_code == 0:
                self.log_message("🎉 ¡Activación completada con éxito!")
                messagebox.showinfo("✅ Éxito", "Office se activó correctamente.")
            else:
                error_msg = process.stderr.read()
                self.log_message(f"❌ Error durante la activación (Código {return_code}): {error_msg}")
                messagebox.showerror("❌ Error", f"Error durante la activación:\n{error_msg}")
            
        except Exception as e:
            self.stop_loading_animation()
            self.toggle_buttons(True)
            self.log_message(f"💥 Error inesperado: {str(e)}")
            messagebox.showerror("💥 Error", f"Error inesperado:\n{str(e)}")

def main():
    root = tk.Tk()
    
    try:
        root.iconbitmap(resource_path("icono.ico"))
    except:
        pass
    
    app = OfficeInstallerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
