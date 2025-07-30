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
        self.top.geometry("400x240")
        self.top.configure(bg="#1e1e1e")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()

        self.label_var = tk.StringVar(value=title)
        self.label = ttk.Label(self.top, textvariable=self.label_var, style="Title.TLabel")
        self.label.pack(pady=20)

        # Canvas de la onda
        self.canvas = tk.Canvas(self.top, width=360, height=100, bg="#1e1e1e", highlightthickness=0)
        self.canvas.pack()
        self.wave_line = self.canvas.create_line(0, 50, 360, 50, fill="#00bfff", width=3, smooth=True)

        # Variables de animación
        self.frame = 0
        self.running = True
        self.color_cycle = itertools.cycle(["#00bfff", "#1abc9c", "#9b59b6", "#f39c12", "#e74c3c"])
        self.current_color = next(self.color_cycle)

        # Contador de puntos
        self.dots = 0

        # Iniciar animaciones
        self.animate_wave()
        self.animate_color()
        self.animate_text()

    def animate_wave(self):
        if not self.running:
            return
        points = []
        for x in range(0, 361, 10):
            y = 50 + 20 * math.sin((x + self.frame) * 0.05) * math.cos((x + self.frame) * 0.03)
            points.extend([x, y])
        self.canvas.coords(self.wave_line, *points)
        self.canvas.itemconfig(self.wave_line, fill=self.current_color)
        self.frame += 5
        self.top.after(30, self.animate_wave)

    def animate_color(self):
        if not self.running:
            return
        self.current_color = next(self.color_cycle)
        self.top.after(1000, self.animate_color)

    def animate_text(self):
        if not self.running:
            return
        dots = "." * (self.dots % 4)
        self.label_var.set(f"Instalando Office{dots}")
        self.dots += 1
        self.top.after(500, self.animate_text)

    def close(self):
        self.running = False
        self.top.destroy()



class OfficeInstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Instalador Office LTSC")
        self.root.geometry("750x770")
        self.root.resizable(False, False)
        
        # Configurar estilo moderno
        self.setup_styles()
        
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
        
        self.create_widgets()
    
    def setup_styles(self):
        style = ttk.Style()
        
        # Tema azure oscuro (asegúrate de tener azure.tcl en el directorio)
        try:
            self.root.tk.call('source', resource_path('azure.tcl'))
            self.root.tk.call('set_theme', 'dark')
        except Exception:
            pass  # En caso de error con el tema, usar default
        
        # Fuentes
        self.title_font = tkfont.Font(family="Segoe UI", size=12, weight="bold")
        self.section_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        self.normal_font = tkfont.Font(family="Segoe UI", size=9)
        
        # Estilos
        style.configure('TFrame', background='#333333')
        style.configure('TLabel', background='#333333', foreground='white', font=self.normal_font)
        style.configure('Title.TLabel', font=self.title_font)
        style.configure('Section.TLabel', font=self.section_font)
        style.configure('TButton', font=self.normal_font, padding=6)
        style.configure('Accent.TButton', font=("Segoe UI", 10, "bold"), padding=8)
        style.configure('TRadiobutton', background='#333333', foreground='white', font=self.normal_font)
        style.configure('TCheckbutton', background='#333333', foreground='white', font=self.normal_font)
        style.configure('TLabelFrame', background='#333333', foreground='white', font=self.section_font)
        style.configure('TLabelFrame.Label', background='#333333', foreground='white')
        style.configure('TNotebook', background='#333333', borderwidth=0)
        style.configure('TNotebook.Tab', background='#444444', foreground='white', padding=[10, 5], font=self.normal_font)
        style.map('TNotebook.Tab', background=[('selected', '#0078D7')])
        style.configure('Vertical.TScrollbar', arrowsize=12, troughcolor='#444444')
    
    def create_widgets(self):
        # Frame principal con pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña de instalación
        self.install_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.install_tab, text="Instalación")
        
        # Frame de contenido
        main_frame = ttk.Frame(self.install_tab, padding=(15, 10))
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(
            main_frame, 
            text="Instalador de Office LTSC Profesional",
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        # Frame de configuración
        config_frame = ttk.LabelFrame(main_frame, text="Configuración de Instalación", padding=10)
        config_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=5)
        config_frame.columnconfigure(0, weight=1)
        config_frame.columnconfigure(1, weight=1)
        
        # Versión de Office
        version_frame = ttk.LabelFrame(config_frame, text="Versión de Office", padding=10)
        version_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        versions = [
            ("Office 2019 LTSC", "2019"),
            ("Office 2021 LTSC", "2021"),
            ("Office 365", "365")
        ]
        
        for i, (text, val) in enumerate(versions):
            ttk.Radiobutton(
                version_frame, text=text, 
                variable=self.version_var, value=val
            ).grid(row=i, column=0, sticky="w", padx=5, pady=2)
        
        # Arquitectura
        arch_frame = ttk.LabelFrame(config_frame, text="Arquitectura", padding=10)
        arch_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        ttk.Radiobutton(
            arch_frame, text="64 bits (Recomendado)", 
            variable=self.architecture_var, value="64"
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        
        ttk.Radiobutton(
            arch_frame, text="32 bits", 
            variable=self.architecture_var, value="32"
        ).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        
        # Componentes adicionales
        components_frame = ttk.LabelFrame(config_frame, text="Componentes Adicionales", padding=10)
        components_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
        ttk.Checkbutton(
            components_frame, text="Instalar Visio Professional",
            variable=self.visio_var
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        
        ttk.Checkbutton(
            components_frame, text="Instalar Project Professional",
            variable=self.project_var
        ).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        
        ttk.Checkbutton(
            components_frame, text="Excluir Skype for Business (Lync)",
            variable=self.exclude_lync_var
        ).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        # Idioma
        language_frame = ttk.LabelFrame(config_frame, text="Idioma de Instalación", padding=10)
        language_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
        languages = [
            ("Inglés (en-us)", "en-us"),
            ("Español (es-es)", "es-es"),
            ("Francés (fr-fr)", "fr-fr"),
            ("Alemán (de-de)", "de-de"),
            ("Portugués (pt-br)", "pt-br")
        ]
        
        for i, (text, lang) in enumerate(languages):
            ttk.Radiobutton(
                language_frame, text=text,
                variable=self.language_var, value=lang
            ).grid(row=i//3, column=i%3, sticky="w", padx=5, pady=2)
        
        # Área de registro
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Actividad", padding=10)
        log_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=10)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            width=80, 
            height=12, 
            state='disabled',
            bg='#252525',
            fg='white',
            insertbackground='white',
            selectbackground='#0078D7',
            font=('Consolas', 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Barra de estado
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5)
        
        self.status_label = ttk.Label(
            status_frame, 
            text="Listo para instalar", 
            style='Section.TLabel',
            foreground='lightgray'
        )
        self.status_label.pack(side=tk.LEFT)
        
        # Frame para botones
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        # Botón de instalación
        self.install_button = ttk.Button(
            button_frame, 
            text="Instalar Office", 
            command=self.start_installation,
            style='Accent.TButton'
        )
        self.install_button.pack(side=tk.LEFT, padx=10)
        
        # Botón de activación
        self.activate_button = ttk.Button(
            button_frame, 
            text="Activar Office", 
            command=self.start_activation,
            style='Accent.TButton'
        )
        self.activate_button.pack(side=tk.LEFT, padx=10)
        
        # Configurar pesos de filas/columnas
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
    
    def log_message(self, message):
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.update_status(message)
        self.root.update()
    
    def update_status(self, message):
        self.status_label.config(text=message[:80] + "..." if len(message) > 80 else message)
    
    def clear_log(self):
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        self.update_status("Listo para instalar")
    
    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def run_as_admin(self):
        if not self.is_admin():
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            sys.exit()
    
    def generate_config_xml(self):
        version = self.version_var.get()
        architecture = self.architecture_var.get()
        language = self.language_var.get()
        
        # Determinar el nombre del archivo de configuración
        config_file = "configuration.xml"
        
        # Generar el contenido XML
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
        self.clear_log()
        self.log_message("Preparando instalación de Office...")
        self.toggle_buttons(False)
        self.run_as_admin()

        # Mostrar ventana de carga
        self.loading_popup = LoadingWindow(self.root, title="Instalando Office...")

        threading.Thread(target=self.install_office, daemon=True).start()

    
    def install_office(self):
        try:
            # Generar el archivo de configuración
            config_file, xml_content = self.generate_config_xml()
            
            with open(config_file, "w", encoding="utf-8") as f:
                f.write(xml_content)
            
            self.log_message(f"✓ Archivo de configuración generado: {config_file}")
            
            # Ejecutar el instalador
            setup_path = os.path.join(os.path.dirname(sys.argv[0]), "setup.exe")
            if not os.path.exists(setup_path):
                setup_path = "setup.exe"
                if not os.path.exists(setup_path):
                    self.log_message("✗ Error: No se encontró setup.exe")
                    self.stop_loading_animation()
                    self.toggle_buttons(True)
                    return
            
            self.log_message("▶ Iniciando instalación de Office...")
            
            process = subprocess.Popen(
                [setup_path, "/configure", config_file],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            # Mostrar salida en tiempo real
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    self.log_message(f"  {output.strip()}")
            
            return_code = process.poll()
            self.stop_loading_animation()
            self.toggle_buttons(True)
            
            if return_code == 0:
                self.log_message("✓ Instalación completada con éxito")
                messagebox.showinfo("Éxito", "Office se instaló correctamente.")
            else:
                error_msg = process.stderr.read()
                self.log_message(f"✗ Error durante la instalación (Código {return_code}): {error_msg}")
                messagebox.showerror("Error", f"Error durante la instalación:\n{error_msg}")
                
        except Exception as e:
            self.stop_loading_animation()
            self.toggle_buttons(True)
            self.log_message(f"✗ Error inesperado: {str(e)}")
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")

        finally:
            if hasattr(self, "loading_popup"):
                self.loading_popup.close()

    
    def start_activation(self):
        self.clear_log()
        self.log_message("Preparando activación de Office...")
        self.toggle_buttons(False)
        self.run_as_admin()
        self.loading_thread = threading.Thread(target=self.start_loading_animation, args=("Activando Office",), daemon=True)
        self.loading_thread.start()
        threading.Thread(target=self.activate_office, daemon=True).start()
    
    def activate_office(self):
        try:
            version = self.version_var.get()
            
            # Claves KMS para diferentes versiones
            kms_keys = {
                "2019": "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP",
                "2021": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH",
                "365": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"  # Misma clave que 2021
            }
            
            kms_key = kms_keys.get(version, kms_keys["2021"])
            
            # Crear un script batch temporal para la activación
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
            # Guardar el script batch temporal
            temp_dir = tempfile.gettempdir()
            batch_path = os.path.join(temp_dir, "activate_office.bat")
            
            with open(batch_path, "w", encoding="utf-8") as f:
                f.write(batch_script)
            
            self.log_message("▶ Ejecutando script de activación...")
            
            # Ejecutar el script y capturar la salida
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
                    self.log_message(f"  {output.strip()}")
            
            return_code = process.poll()
            self.stop_loading_animation()
            self.toggle_buttons(True)
            
            if return_code == 0:
                self.log_message("✓ Activación completada con éxito")
                messagebox.showinfo("Éxito", "Office se activó correctamente.")
            else:
                error_msg = process.stderr.read()
                self.log_message(f"✗ Error durante la activación (Código {return_code}): {error_msg}")
                messagebox.showerror("Error", f"Error durante la activación:\n{error_msg}")
            
        except Exception as e:
            self.stop_loading_animation()
            self.toggle_buttons(True)
            self.log_message(f"✗ Error inesperado: {str(e)}")
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")

def main():
    root = tk.Tk()
    
    # Configurar icono (opcional)
    try:
        root.iconbitmap(resource_path("icono.ico"))

    except:
        pass
    
    app = OfficeInstallerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()