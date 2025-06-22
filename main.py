from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import os
import subprocess
import threading
import time
import signal
import sys
import tkinter as tk
from tkinter import ttk
import random
import string
from datetime import datetime, timezone, timedelta
import queue
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import psutil
import gc

class ResourceProfiler:
    """Track and display resource usage metrics."""
    
    def __init__(self):
        self.process = psutil.Process()
        self.start_time = time.time()
        self.metrics = {
            'cpu_percent': [],
            'memory_mb': [],
            'thread_count': [],
            'handle_count': [],
            'io_counters': [],
            'timestamps': []
        }
        self.profiling_active = False
        self.profile_thread = None
        
    def start_profiling(self):
        """Start resource profiling in a separate thread."""
        if not self.profiling_active:
            self.profiling_active = True
            self.profile_thread = threading.Thread(target=self._profile_loop, daemon=True)
            self.profile_thread.start()
            print("Resource profiling started")
    
    def stop_profiling(self):
        """Stop resource profiling."""
        self.profiling_active = False
        if self.profile_thread:
            self.profile_thread.join(timeout=1)
        print("Resource profiling stopped")
    
    def _profile_loop(self):
        """Main profiling loop."""
        while self.profiling_active:
            try:
                # Get current metrics
                cpu_percent = self.process.cpu_percent()
                memory_info = self.process.memory_info()
                memory_mb = memory_info.rss / 1024 / 1024  # Convert to MB
                thread_count = self.process.num_threads()
                handle_count = self.process.num_handles()
                
                # Get IO counters if available
                try:
                    io_counters = self.process.io_counters()
                    io_read_mb = io_counters.read_bytes / 1024 / 1024
                    io_write_mb = io_counters.write_bytes / 1024 / 1024
                except:
                    io_read_mb = 0
                    io_write_mb = 0
                
                # Store metrics
                current_time = time.time()
                self.metrics['cpu_percent'].append(cpu_percent)
                self.metrics['memory_mb'].append(memory_mb)
                self.metrics['thread_count'].append(thread_count)
                self.metrics['handle_count'].append(handle_count)
                self.metrics['io_counters'].append((io_read_mb, io_write_mb))
                self.metrics['timestamps'].append(current_time)
                
                # Keep only last 1000 data points to prevent memory bloat
                max_points = 1000
                if len(self.metrics['timestamps']) > max_points:
                    for key in self.metrics:
                        self.metrics[key] = self.metrics[key][-max_points:]
                
                time.sleep(1)  # Sample every second
                
            except Exception as e:
                print(f"Profiling error: {e}")
                time.sleep(1)
    
    def get_current_stats(self):
        """Get current resource statistics."""
        if not self.metrics['timestamps']:
            return None
            
        cpu_avg = sum(self.metrics['cpu_percent'][-10:]) / min(10, len(self.metrics['cpu_percent']))
        memory_current = self.metrics['memory_mb'][-1] if self.metrics['memory_mb'] else 0
        memory_avg = sum(self.metrics['memory_mb'][-10:]) / min(10, len(self.metrics['memory_mb']))
        thread_current = self.metrics['thread_count'][-1] if self.metrics['thread_count'] else 0
        handle_current = self.metrics['handle_count'][-1] if self.metrics['handle_count'] else 0
        
        # Calculate IO rates
        if len(self.metrics['io_counters']) >= 2:
            recent_io = self.metrics['io_counters'][-10:]
            io_read_total = sum(read for read, write in recent_io)
            io_write_total = sum(write for read, write in recent_io)
        else:
            io_read_total = 0
            io_write_total = 0
        
        uptime = time.time() - self.start_time
        
        return {
            'cpu_avg': cpu_avg,
            'memory_current': memory_current,
            'memory_avg': memory_avg,
            'thread_count': thread_current,
            'handle_count': handle_current,
            'io_read_mb': io_read_total,
            'io_write_mb': io_write_total,
            'uptime': uptime,
            'data_points': len(self.metrics['timestamps'])
        }
    
    def get_peak_stats(self):
        """Get peak resource usage."""
        if not self.metrics['timestamps']:
            return None
            
        return {
            'cpu_peak': max(self.metrics['cpu_percent']),
            'memory_peak': max(self.metrics['memory_mb']),
            'thread_peak': max(self.metrics['thread_count']),
            'handle_peak': max(self.metrics['handle_count'])
        }
    
    def export_profile_data(self, filename="resource_profile.json"):
        """Export profiling data to JSON file."""
        try:
            export_data = {
                'start_time': self.start_time,
                'end_time': time.time(),
                'metrics': self.metrics,
                'peak_stats': self.get_peak_stats(),
                'current_stats': self.get_current_stats()
            }
            
            with open(filename, 'w') as f:
                json.dump(export_data, f, indent=2, default=str)
            
            print(f"Profile data exported to {filename}")
            return True
        except Exception as e:
            print(f"Failed to export profile data: {e}")
            return False

# Global profiler instance
resource_profiler = ResourceProfiler()

class SelectionActions:
    START = "launch"
    STOP = "shutdown"

global_should_stop = False
is_cycling = False

def signal_handler(sig, frame):
    """Handle Ctrl+C and other termination signals."""
    global global_should_stop
    print("\nRecebido Ctrl+C, encerrando graciosamente...")
    global_should_stop = True

def get_vm_info(mumu_base_path):
    """Get VM information by querying MuMuManager info for indices 0 to 20 in parallel."""
    start_time = time.time()
    mumu_manager = os.path.join(mumu_base_path, "shell", "MuMuManager.exe")
    vm_info = []
    indices = range(21)

    def query_vm(index):
        try:
            result = subprocess.run(
                [mumu_manager, "info", "-v", str(index)],
                capture_output=True,
                text=True,
                check=True,
                timeout=2  # Add timeout to prevent hanging
            )
            vm_data = json.loads(result.stdout)
            if vm_data.get("error_code", -1) == 0:
                is_android_started = vm_data.get("is_android_started", False)
                is_process_started = vm_data.get("is_process_started", False)
                launch_err_code = vm_data.get("launch_err_code", 0)
                launch_err_msg = vm_data.get("launch_err_msg", "")
                has_error = launch_err_code != 0 or bool(launch_err_msg)

                # Status reflects the process state, which is what 'stop' commands target.
                process_status = "running" if is_process_started else "stopped"
                
                return {
                    "index": vm_data["index"],
                    "name": vm_data["name"],
                    "status": process_status,
                    "is_main": vm_data.get("is_main", False),
                    "is_android_started": is_android_started,
                    "has_error": has_error,
                    "error_message": launch_err_msg,
                }
            return None
        except (subprocess.CalledProcessError, json.JSONDecodeError, subprocess.TimeoutExpired):
            return None

    # Use fewer workers when not cycling to reduce resource usage
    max_workers = 5 if not is_cycling else 10
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_index = {executor.submit(query_vm, i): i for i in indices}
        for future in as_completed(future_to_index):
            result = future.result()
            if result:
                vm_info.append(result)

    vm_info.sort(key=lambda x: x['index'])  # Sort by index for consistent order
    elapsed_time = time.time() - start_time
    if elapsed_time > 1.0:  # Only print if it takes more than 1 second
        print(f"Tempo decorrido: {elapsed_time:.2f} segundos")
    return vm_info if vm_info else []

def padronize_vm_names(vm_name_prefix, mumu_base_path):
    """Standardize VM names, skipping main VM (index 0)."""
    mumu_manager = os.path.join(mumu_base_path, "shell", "MuMuManager.exe")
    vm_names = []
    used_names = set()

    def generate_unique_name():
        while True:
            suffix = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
            name = f"{vm_name_prefix}{suffix}"
            if name not in used_names:
                used_names.add(name)
                return name

    vm_info = get_vm_info(mumu_base_path)
    if not vm_info:
        print("Aviso: Nenhuma VM válida encontrada. Crie VMs no Gerenciador de Instâncias Múltiplas do MuMu.")
        return vm_names

    for vm in vm_info:
        current_name = vm["name"]
        if vm["is_main"]:
            print(f"Ignorando renomeação para VM principal {vm['index']}: {current_name}")
            vm_names.append(current_name)
            used_names.add(current_name)
            continue
        if not current_name.startswith(vm_name_prefix):
            new_name = generate_unique_name()
            try:
                subprocess.run(
                    [mumu_manager, "rename", "-v", vm["index"], new_name],
                    capture_output=True,
                    text=True,
                    check=True
                )
                vm_names.append(new_name)
                print(f"Renomeada VM {vm['index']} para {new_name}")
            except subprocess.CalledProcessError as e:
                print(f"Erro ao renomear VM {vm['index']}: {e}")
        else:
            vm_names.append(current_name)
            used_names.add(current_name)

    return sorted(vm_names)

def control_instances(mumu_manager, vm_indices, action):
    """
    Control VM instances one at a time using MuMuManager.
    Returns a tuple of (successful_indices, failed_indices).
    """
    if not vm_indices:
        print("Nenhum índice de VM fornecido para controle.")
        return [], []
    
    successful_indices = []
    failed_indices = []
    
    for idx in vm_indices:
        try:
            # Add a timeout to prevent hanging
            subprocess.run(
                [mumu_manager, "control", action, "-v", str(idx)],
                capture_output=True,
                text=True,
                check=True,
                timeout=30
            )
            print(f"Executado com sucesso {action} para VM {idx}")
            successful_indices.append(idx)
            # wait 1 second
            time.sleep(1)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            print(f"Erro ao executar {action} para VM {idx}: {e}")
            failed_indices.append(idx)
            
    return successful_indices, failed_indices

def load_settings():
    """Load settings from JSON file."""
    settings_file = "mumu_settings.json"
    default_settings = {
        "batch_size": 1,
        "cycle_interval": 60
    }
    
    try:
        with open(settings_file, "r", encoding="utf-8") as f:
            settings = json.load(f)
            # Ensure all required keys exist
            for key, default_value in default_settings.items():
                if key not in settings:
                    settings[key] = default_value
            return settings
    except (FileNotFoundError, json.JSONDecodeError):
        return default_settings

def save_settings(batch_size, cycle_interval):
    """Save settings to JSON file."""
    settings_file = "mumu_settings.json"
    settings = {
        "batch_size": batch_size,
        "cycle_interval": cycle_interval
    }
    
    try:
        with open(settings_file, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
        print(f"Configurações salvas em {settings_file}")
    except Exception as e:
        print(f"Erro ao salvar configurações: {e}")

def create_path_config_ui():
    """Create a Tkinter UI for configuring the MuMu installation path."""
    path_window = tk.Tk()
    path_window.title("Configurar Caminho do MuMu")
    path_window.geometry("600x200")
    path_window.configure(bg="#F5F5F5")
    path_window.resizable(False, False)

    style = ttk.Style()
    style.theme_use('default')
    style.configure("TButton", font=("Arial", 10))
    style.configure("TLabel", background="#F5F5F5", foreground="#212121", font=("Arial", 10))

    last_path_file = "last_mumu_path.json"
    prev_path = ""
    try:
        with open(last_path_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        prev_path = data.get("mumu_base_path", "").strip()
        if not (prev_path and os.path.isdir(prev_path) and os.path.exists(os.path.join(prev_path, "shell", "MuMuManager.exe"))):
            prev_path = ""
    except Exception:
        prev_path = ""

    main_frame = tk.Frame(path_window, bg="#F5F5F5")
    main_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

    path_var = tk.StringVar(value=prev_path)
    result = {"mumu_base_path": None}

    def validate_path(path):
        if not path or not os.path.isdir(path):
            return False
        mumu_manager = os.path.join(path, "shell", "MuMuManager.exe")
        return os.path.exists(mumu_manager)

    def reuse_path():
        if validate_path(prev_path):
            result["mumu_base_path"] = prev_path
            try:
                with open(last_path_file, "w", encoding="utf-8") as f:
                    json.dump({"mumu_base_path": prev_path}, f)
                print(f"Caminho salvo em {last_path_file}")
            except Exception as e:
                print(f"Não foi possível salvar o caminho: {e}")
            path_window.destroy()
        else:
            error_label.config(text="Caminho anterior inválido ou MuMuManager.exe não encontrado")

    def confirm_path():
        path = path_entry.get().strip('"').strip()
        if validate_path(path):
            result["mumu_base_path"] = path
            try:
                with open(last_path_file, "w", encoding="utf-8") as f:
                    json.dump({"mumu_base_path": path}, f)
                print(f"Caminho salvo em {last_path_file}")
            except Exception as e:
                print(f"Não foi possível salvar o caminho: {e}")
            path_window.destroy()
        else:
            error_label.config(text="Caminho inválido ou MuMuManager.exe não encontrado")

    if prev_path:
        tk.Label(main_frame, text=f"Caminho anterior: {prev_path}", bg="#F5F5F5", fg="#212121", font=("Arial", 10)).pack(anchor="w", pady=5)
        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(anchor="w", pady=5)
        ttk.Button(button_frame, text="Reutilizar", command=reuse_path).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Mudar Caminho", command=lambda: [prev_path_label.pack_forget(), button_frame.pack_forget(), new_path_frame.pack(anchor="w", pady=5)]).pack(side=tk.LEFT, padx=5)
    else:
        tk.Label(main_frame, text="Digite o caminho de instalação do MuMu (ex.: D:\\Program Files\\Netease\\MuMuPlayerGlobal-12.0):", bg="#F5F5F5", fg="#212121", font=("Arial", 10)).pack(anchor="w", pady=5)

    prev_path_label = tk.Label(main_frame, text="Digite o caminho de instalação do MuMu (ex.: D:\\Program Files\\Netease\\MuMuPlayerGlobal-12.0):", bg="#F5F5F5", fg="#212121", font=("Arial", 10))
    new_path_frame = tk.Frame(main_frame, bg="#F5F5F5")
    path_entry = tk.Entry(new_path_frame, textvariable=path_var, width=50, font=("Arial", 10))
    path_entry.pack(side=tk.LEFT, padx=5)
    ttk.Button(new_path_frame, text="Confirmar", command=confirm_path).pack(side=tk.LEFT, padx=5)
    error_label = tk.Label(main_frame, text="", bg="#F5F5F5", fg="#D32F2F", font=("Arial", 10))
    error_label.pack(anchor="w", pady=5)

    if not prev_path:
        prev_path_label.pack(anchor="w", pady=5)
        new_path_frame.pack(anchor="w", pady=5)

    def on_closing():
        sys.exit(0)

    path_window.protocol("WM_DELETE_WINDOW", on_closing)
    path_window.mainloop()
    return result["mumu_base_path"]

def create_ui(vm_names, cycle_interval, update_queue, mumu_manager, management_thread_container, total_instances):
    """Create a Tkinter UI with a light mode design, chart, and cycle controls."""
    root = tk.Tk()
    root.title("Gerenciador de Instâncias MuMu")
    root.geometry("800x600")
    root.configure(bg="#F5F5F5")  # Light background

    style = ttk.Style()
    style.theme_use('default')
    style.configure("TProgressbar", troughcolor="#E0E0E0", background="#2196F3")
    style.configure("TLabel", background="#F5F5F5", foreground="#212121", font=("Arial", 10))
    style.configure("TButton", font=("Arial", 10))

    # Load saved settings
    settings = load_settings()
    saved_batch_size = settings.get("batch_size", 1)
    saved_cycle_interval = settings.get("cycle_interval", 60)

    top_frame = tk.Frame(root, bg="#F5F5F5")
    top_frame.pack(pady=10, fill=tk.X, padx=10)

    control_frame = tk.Frame(top_frame, bg="#F5F5F5")
    control_frame.pack(anchor="w", pady=5)
    
    # Batch size controls
    tk.Label(control_frame, text="Tamanho do Lote:", bg="#F5F5F5", fg="#212121", font=("Arial", 10)).pack(side=tk.LEFT)
    batch_size_entry = tk.Entry(control_frame, width=5, font=("Arial", 10))
    batch_size_entry.insert(0, str(saved_batch_size))
    batch_size_entry.pack(side=tk.LEFT, padx=5)
    
    # Cycle interval controls
    tk.Label(control_frame, text="Intervalo (seg):", bg="#F5F5F5", fg="#212121", font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 0))
    cycle_interval_entry = tk.Entry(control_frame, width=5, font=("Arial", 10))
    cycle_interval_entry.insert(0, str(saved_cycle_interval))
    cycle_interval_entry.pack(side=tk.LEFT, padx=5)
    
    start_button = ttk.Button(control_frame, text="Iniciar Ciclo")
    start_button.pack(side=tk.LEFT, padx=5)
    stop_button = ttk.Button(control_frame, text="Parar Ciclo", state="disabled")
    stop_button.pack(side=tk.LEFT, padx=5)
    error_label = tk.Label(control_frame, text="", bg="#F5F5F5", fg="#D32F2F", font=("Arial", 10))
    error_label.pack(side=tk.LEFT, padx=5)

    progress_label = tk.Label(top_frame, text="Progresso do Ciclo:", bg="#F5F5F5", fg="#212121", font=("Arial", 12))
    progress_label.pack(anchor="w")
    progress_bar = ttk.Progressbar(top_frame, length=400, mode='determinate')
    progress_bar.pack(fill=tk.X, pady=5)

    status_label = tk.Label(top_frame, text="Status: Inativo", bg="#F5F5F5", fg="#212121", font=("Arial", 10))
    status_label.pack(anchor="w")
    current_batch_label = tk.Label(top_frame, text="Lote Atual: Nenhum", bg="#F5F5F5", fg="#212121", font=("Arial", 10))
    current_batch_label.pack(anchor="w")
    last_cycle_label = tk.Label(top_frame, text="Último Ciclo: Nenhum", bg="#F5F5F5", fg="#212121", font=("Arial", 10))
    last_cycle_label.pack(anchor="w")
    cycle_count_label = tk.Label(top_frame, text="Ciclos: 0 / 0", bg="#F5F5F5", fg="#212121", font=("Arial", 10))
    cycle_count_label.pack(anchor="w")
    time_remaining_label = tk.Label(top_frame, text="Tempo Restante: N/A", bg="#F5F5F5", fg="#212121", font=("Arial", 10))
    time_remaining_label.pack(anchor="w")

    # Profiling section
    profiling_frame = tk.Frame(top_frame, bg="#F5F5F5")
    profiling_frame.pack(anchor="w", pady=(10, 0))
    
    profiling_header = tk.Label(profiling_frame, text="📊 Monitoramento de Recursos", bg="#F5F5F5", fg="#212121", font=("Arial", 12, "bold"))
    profiling_header.pack(anchor="w")
    
    # Resource metrics display
    metrics_frame = tk.Frame(profiling_frame, bg="#F5F5F5")
    metrics_frame.pack(anchor="w", pady=5)
    
    # CPU and Memory
    cpu_label = tk.Label(metrics_frame, text="CPU: 0.0%", bg="#F5F5F5", fg="#212121", font=("Arial", 9))
    cpu_label.pack(side=tk.LEFT, padx=(0, 10))
    
    memory_label = tk.Label(metrics_frame, text="RAM: 0.0 MB", bg="#F5F5F5", fg="#212121", font=("Arial", 9))
    memory_label.pack(side=tk.LEFT, padx=(0, 10))
    
    thread_label = tk.Label(metrics_frame, text="Threads: 0", bg="#F5F5F5", fg="#212121", font=("Arial", 9))
    thread_label.pack(side=tk.LEFT, padx=(0, 10))
    
    handle_label = tk.Label(metrics_frame, text="Handles: 0", bg="#F5F5F5", fg="#212121", font=("Arial", 9))
    handle_label.pack(side=tk.LEFT, padx=(0, 10))
    
    uptime_label = tk.Label(metrics_frame, text="Uptime: 0s", bg="#F5F5F5", fg="#212121", font=("Arial", 9))
    uptime_label.pack(side=tk.LEFT, padx=(0, 10))
    
    # Profiling controls
    profiling_controls = tk.Frame(profiling_frame, bg="#F5F5F5")
    profiling_controls.pack(anchor="w", pady=5)
    
    start_profiling_btn = ttk.Button(profiling_controls, text="Iniciar Profiling")
    start_profiling_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    stop_profiling_btn = ttk.Button(profiling_controls, text="Parar Profiling", state="disabled")
    stop_profiling_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    export_profile_btn = ttk.Button(profiling_controls, text="Exportar Dados")
    export_profile_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    clear_profile_btn = ttk.Button(profiling_controls, text="Limpar Dados")
    clear_profile_btn.pack(side=tk.LEFT, padx=(0, 5))

    bottom_frame = tk.Frame(root, bg="#F5F5F5")
    bottom_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=10)

    instance_frame = tk.Frame(bottom_frame, bg="#F5F5F5")
    instance_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    instance_listbox = tk.Listbox(instance_frame, height=10, width=30, bg="#FFFFFF", fg="#212121", font=("Arial", 10))
    instance_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar = tk.Scrollbar(instance_frame, orient=tk.VERTICAL, command=instance_listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    instance_listbox.config(yscrollcommand=scrollbar.set)

    chart_frame = tk.Frame(bottom_frame, bg="#F5F5F5")
    chart_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
    fig, ax = plt.subplots(figsize=(4, 3), facecolor="#F5F5F5")
    ax.set_facecolor("#FFFFFF")
    canvas = FigureCanvasTkAgg(fig, master=chart_frame)
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    times = []
    running_counts = []
    last_vm_info = None
    stop_ui_update = False
    after_id = None
    batch_size = tk.IntVar(value=0)  # Create batch_size here
    cycle_count = 0  # Track current session cycles
    total_cycles_needed = 0  # Track total cycles needed to complete all instances
    is_retry_attempt = False # New state variable for retry logic

    def calculate_total_cycles():
        """Calculate total cycles needed based on current batch size."""
        try:
            batch_size_val = int(batch_size_entry.get())
            if batch_size_val > 0 and batch_size_val <= total_instances:
                return (total_instances + batch_size_val - 1) // batch_size_val
        except ValueError:
            pass
        return 0

    def update_cycle_display():
        """Update the cycle count display."""
        total_needed = calculate_total_cycles()
        cycle_count_label.config(text=f"Ciclos: {cycle_count} / {total_needed}")

    def validate_batch_size():
        try:
            size = int(batch_size_entry.get())
            
            if size <= 0:
                error_label.config(text="Tamanho do lote deve ser um número positivo")
                return False
            if size > total_instances:
                error_label.config(text=f"Tamanho do lote deve ser ≤ {total_instances}")
                return False
            error_label.config(text="")
            # Update cycle display when batch size changes
            update_cycle_display()
            return size
        except ValueError:
            error_label.config(text="Tamanho do lote deve ser um número válido")
            return False

    def validate_cycle_interval():
        try:
            interval = int(cycle_interval_entry.get())
            
            if interval <= 0:
                error_label.config(text="Intervalo deve ser um número positivo")
                return False
            if interval < 10:
                error_label.config(text="Intervalo mínimo é 10 segundos")
                return False
            error_label.config(text="")
            return interval
        except ValueError:
            error_label.config(text="Intervalo deve ser um número válido")
            return False

    def start_cycle():
        global is_cycling
        batch_size_val = validate_batch_size()
        cycle_interval_val = validate_cycle_interval()
        
        if batch_size_val and cycle_interval_val:
            # Stop all running instances first to ensure clean start
            running_vms = [int(vm["index"]) for vm in get_vm_info(os.path.dirname(os.path.dirname(mumu_manager))) if vm["status"] == "running" and not vm["is_main"]]
            if running_vms:
                print("Parando todas as instâncias para iniciar novo ciclo...")
                control_instances(mumu_manager, running_vms, SelectionActions.STOP)
                time.sleep(3)  # Wait for instances to stop
            
            batch_size.set(batch_size_val)
            is_cycling = True
            start_button.config(state="disabled")
            stop_button.config(state="normal")
            error_label.config(text="")
            status_label.config(text="Status: Em Execução")
            
            # Reset cycle count when starting and calculate total cycles needed
            nonlocal cycle_count, total_cycles_needed
            cycle_count = 0
            total_cycles_needed = calculate_total_cycles()
            
            # Save settings
            save_settings(batch_size_val, cycle_interval_val)
            
            # Clear queue to avoid stale data
            while not update_queue.empty():
                update_queue.get()
            # Always send the latest interval and batch size
            update_queue.put({
                "reset_index": True, 
                "batch_size": batch_size_val,
                "cycle_interval": cycle_interval_val,
                "reset_cycle": True  # Signal to reset cycle state
            })
            print(f"Ciclo iniciado com reset completo. Intervalo: {cycle_interval_val}s")
        else:
            print("Falha ao iniciar ciclo: valores inválidos")

    def stop_cycle():
        global is_cycling
        is_cycling = False
        start_button.config(state="normal")
        stop_button.config(state="disabled")
        status_label.config(text="Status: Parado")
        progress_bar['value'] = 0
        time_remaining_label.config(text="Tempo Restante: N/A")
        current_batch_label.config(text="Lote Atual: Nenhum")
        last_cycle_label.config(text="Último Ciclo: Nenhum")
        
        # Reset cycle count
        nonlocal cycle_count
        cycle_count = 0
        update_cycle_display()
        
        # Try to save settings if validation passes, but don't require it
        try:
            batch_size_val = validate_batch_size()
            cycle_interval_val = validate_cycle_interval()
            if batch_size_val and cycle_interval_val:
                save_settings(batch_size_val, cycle_interval_val)
        except:
            pass  # Don't fail if validation doesn't pass
        
        # Stop all running instances
        running_vms = [int(vm["index"]) for vm in get_vm_info(os.path.dirname(os.path.dirname(mumu_manager))) if vm["status"] == "running" and not vm["is_main"]]
        if running_vms:
            print("Parando todas as instâncias...")
            control_instances(mumu_manager, running_vms, SelectionActions.STOP)
            print("Todas as instâncias paradas.")
        
        # Clear queue and reset state
        while not update_queue.empty():
            update_queue.get()
        update_queue.put({
            "last_routine_run": None, 
            "last_routine_range": None,
            "reset_cycle": True
        })
        print("Ciclo parado e resetado completamente.")

    start_button.config(command=start_cycle)
    stop_button.config(command=stop_cycle)

    # Initialize cycle display with loaded settings
    update_cycle_display()

    # Profiling control functions
    def start_profiling():
        global resource_profiler
        resource_profiler.start_profiling()
        start_profiling_btn.config(state="disabled")
        stop_profiling_btn.config(state="normal")
        print("Resource profiling iniciado")

    def stop_profiling():
        global resource_profiler
        resource_profiler.stop_profiling()
        start_profiling_btn.config(state="normal")
        stop_profiling_btn.config(state="disabled")
        print("Resource profiling parado")

    def export_profile():
        global resource_profiler
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"mumu_profile_{timestamp}.json"
        if resource_profiler.export_profile_data(filename):
            print(f"Perfil exportado: {filename}")
        else:
            print("Erro ao exportar perfil")

    def clear_profile():
        global resource_profiler
        resource_profiler = ResourceProfiler()
        print("Dados de perfil limpos")

    def show_peak_stats():
        """Display peak resource usage statistics."""
        global resource_profiler
        peak_stats = resource_profiler.get_peak_stats()
        if peak_stats:
            stats_window = tk.Toplevel(root)
            stats_window.title("Estatísticas de Pico")
            stats_window.geometry("400x300")
            stats_window.configure(bg="#F5F5F5")
            
            tk.Label(stats_window, text="📈 Estatísticas de Pico de Recursos", 
                    bg="#F5F5F5", fg="#212121", font=("Arial", 14, "bold")).pack(pady=10)
            
            stats_text = f"""
CPU Máximo: {peak_stats['cpu_peak']:.1f}%
Memória Máxima: {peak_stats['memory_peak']:.1f} MB
Threads Máximos: {peak_stats['thread_peak']}
Handles Máximos: {peak_stats['handle_peak']}
            """
            
            tk.Label(stats_window, text=stats_text, 
                    bg="#F5F5F5", fg="#212121", font=("Arial", 10), justify=tk.LEFT).pack(pady=10)
            
            ttk.Button(stats_window, text="Fechar", command=stats_window.destroy).pack(pady=10)
        else:
            print("Nenhum dado de pico disponível")

    # Configure profiling buttons
    start_profiling_btn.config(command=start_profiling)
    stop_profiling_btn.config(command=stop_profiling)
    export_profile_btn.config(command=export_profile)
    clear_profile_btn.config(command=clear_profile)
    
    # Add peak stats button
    peak_stats_btn = ttk.Button(profiling_controls, text="Ver Picos", command=show_peak_stats)
    peak_stats_btn.pack(side=tk.LEFT, padx=(0, 5))

    # Ensure clean startup state
    def ensure_clean_startup():
        """Ensure no instances are running when the program starts."""
        running_vms = [int(vm["index"]) for vm in get_vm_info(os.path.dirname(os.path.dirname(mumu_manager))) if vm["status"] == "running" and not vm["is_main"]]
        if running_vms:
            print("Limpando instâncias de sessões anteriores...")
            control_instances(mumu_manager, running_vms, SelectionActions.STOP)
            time.sleep(2)
            print("Estado limpo - pronto para iniciar.")

    # Call clean startup
    ensure_clean_startup()

    def close_window():
        """Handle window close event with a confirmation dialog."""
        from tkinter import messagebox

        title = "Aviso: Ação Crítica"
        message = (
            "Fechar esta janela irá interromper e reiniciar todo o ciclo de automação.\n\n"
            "Esta é uma ação crítica que não deve ser executada durante a operação normal.\n\n"
            "Tem certeza que deseja continuar e encerrar o programa?"
        )

        if messagebox.askyesno(title, message, icon='warning'):
            global global_should_stop, is_cycling, resource_profiler
            nonlocal stop_ui_update, after_id
            print("Confirmado. Fechando janela e encerrando programa...")

            # Stop profiling and export final data
            if resource_profiler.profiling_active:
                print("Parando profiling e exportando dados finais...")
                resource_profiler.stop_profiling()
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                final_filename = f"mumu_profile_final_{timestamp}.json"
                resource_profiler.export_profile_data(final_filename)
            
            global_should_stop = True
            is_cycling = False
            stop_ui_update = True
            if after_id:
                root.after_cancel(after_id)
            if management_thread_container and management_thread_container[0].is_alive():
                management_thread_container[0].join(timeout=10.0)  # Wait for cleanup
            root.destroy()
            print("Janela fechada, processo encerrado.")
            sys.exit(0)
        else:
            print("Encerramento da aplicação cancelado pelo usuário.")

    root.protocol("WM_DELETE_WINDOW", close_window)

    def update_ui():
        nonlocal last_vm_info, stop_ui_update, after_id, cycle_count, total_cycles_needed
        if stop_ui_update or global_should_stop:
            return
        try:
            # Adaptive UI update frequency: more frequent when cycling, less when idle
            update_interval = 100 if is_cycling else 500  # 100ms when cycling, 500ms when idle
            
            while True:
                data = update_queue.get_nowait()
                last_routine_run = data.get("last_routine_run")
                last_routine_range = data.get("last_routine_range")
                current_time = data.get("current_time")
                status = data.get("status")
                vm_info = data.get("vm_info", [])
                cycle_completed = data.get("cycle_completed", False)
                reset_cycle = data.get("reset_cycle", False)
                
                # Handle cycle reset
                if reset_cycle:
                    cycle_count = 0
                    update_cycle_display()
                
                # Update cycle count if a cycle was completed
                if cycle_completed:
                    cycle_count += 1
                
                # Get current cycle interval from UI or use default
                try:
                    current_cycle_interval = int(cycle_interval_entry.get())
                except ValueError:
                    current_cycle_interval = 60

                if status:
                    status_label.config(text=f"Status: {status}")

                if last_routine_run is None or not is_cycling:
                    progress = 0
                    time_remaining = "N/A"
                else:
                    elapsed = current_time - last_routine_run
                    progress = (elapsed / current_cycle_interval) * 100
                    time_remaining = f"{int(current_cycle_interval - elapsed)} segundos" if elapsed < current_cycle_interval else "0 segundos"
                
                progress_bar['value'] = min(progress, 100)
                time_remaining_label.config(text=f"Tempo Restante: {time_remaining}")
                
                # Update cycle count display
                update_cycle_display()
                
                # Update resource metrics display
                stats = resource_profiler.get_current_stats()
                if stats:
                    cpu_label.config(text=f"CPU: {stats['cpu_avg']:.1f}%")
                    memory_label.config(text=f"RAM: {stats['memory_current']:.1f} MB")
                    thread_label.config(text=f"Threads: {stats['thread_count']}")
                    handle_label.config(text=f"Handles: {stats['handle_count']}")
                    
                    # Format uptime
                    uptime_seconds = int(stats['uptime'])
                    uptime_str = f"{uptime_seconds//3600}h {(uptime_seconds%3600)//60}m {uptime_seconds%60}s"
                    uptime_label.config(text=f"Uptime: {uptime_str}")
                
                if last_routine_range and is_cycling:
                    start, end = last_routine_range
                    if start > end:
                        current_batch_label.config(text=f"Lote Atual: Instâncias {start} a {total_instances}, 1 a {end}")
                    else:
                        current_batch_label.config(text=f"Lote Atual: Instâncias {start} a {end}")
                else:
                    current_batch_label.config(text="Lote Atual: Nenhum")
                
                if last_routine_run and is_cycling:
                    last_cycle_label.config(text=f"Último Ciclo: {datetime.fromtimestamp(last_routine_run, tz=timezone(timedelta(hours=-3))).strftime('%H:%M:%S')}")
                
                instance_listbox.delete(0, tk.END)
                if vm_info:  # Check if vm_info is not empty
                    for vm in vm_info:
                        status_pt = "Parado"  # Default
                        if vm.get('is_android_started'):
                            status_pt = "Em Execução"
                        elif vm.get('has_error'):
                            error_msg = vm.get('error_message', 'Erro')
                            status_pt = f"Erro ({error_msg})"
                        elif vm.get('status') == 'running':
                            status_pt = "Iniciando"
                        
                        instance_listbox.insert(tk.END, f"VM {vm['index']}: {vm['name']} ({status_pt})")
                else:
                    instance_listbox.insert(tk.END, "Nenhuma VM disponível")

                # Optimize chart updates: only update when there are significant changes
                if vm_info and vm_info != last_vm_info:
                    running_count = sum(1 for vm in vm_info if vm["status"] == "running")
                    times.append(current_time)
                    running_counts.append(running_count)
                    
                    # Limit chart data points to prevent memory bloat
                    if len(times) > 100:
                        times.pop(0)
                        running_counts.pop(0)
                    
                    # Only update chart when cycling or when there are significant changes
                    if is_cycling or abs(running_count - (running_counts[-2] if len(running_counts) > 1 else 0)) > 0:
                        ax.clear()
                        ax.plot(times, running_counts, label="VMs em Execução", color="#2196F3")
                        ax.set_xlabel("Tempo (s)", color="#212121")
                        ax.set_ylabel("Número de VMs em Execução", color="#212121")
                        ax.tick_params(colors="#212121")
                        ax.grid(color=(0, 0, 0, 0.1))
                        ax.set_facecolor("#FFFFFF")
                        fig.patch.set_facecolor("#F5F5F5")
                        ax.legend(facecolor="#FFFFFF", edgecolor="#212121", labelcolor="#212121")
                        canvas.draw()
                    last_vm_info = vm_info

        except queue.Empty:
            pass
        except Exception as e:
            print(f"Erro na atualização da UI: {e}")
        if not stop_ui_update and not global_should_stop:
            after_id = root.after(update_interval, update_ui)

    return root, update_ui, status_label, batch_size, cycle_interval_entry

def main():
    global global_should_stop, total_instances
    
    # Load saved settings
    settings = load_settings()
    cycle_interval = settings.get("cycle_interval", 60)
    
    update_queue = queue.Queue()

    # Get path configuration via UI
    mumu_base_path = create_path_config_ui()
    if not mumu_base_path:
        print("Nenhum caminho fornecido. Encerrando.")
        sys.exit(1)

    mumu_manager = os.path.join(mumu_base_path, "shell", "MuMuManager.exe")
    vm_base_name = "ROM_"
    print("Verificando nomes das VMs...")
    vm_info = get_vm_info(mumu_base_path)
    active_vm_info = [vm for vm in vm_info if not vm["is_main"]]
    active_vm_indices = sorted([int(vm["index"]) for vm in active_vm_info])

    print("Parando todas as VMs não principais antes de renomear...")
    running_vm_indices = [int(vm["index"]) for vm in active_vm_info if vm["status"] == "running"]
    if running_vm_indices:
        control_instances(mumu_manager, running_vm_indices, SelectionActions.STOP)
        time.sleep(5)
    else:
        print("Nenhuma VM não principal em execução encontrada.")

    vm_names = padronize_vm_names(vm_base_name, mumu_base_path)
    print(f"Encontradas {len(vm_names)} VMs: {vm_names}")
    
    if not active_vm_indices:
        print("Erro: Nenhuma VM não principal encontrada. Crie VMs no Gerenciador de Instâncias Múltiplas do MuMu.")
        sys.exit(1)

    time.sleep(5)

    total_instances = len(active_vm_indices)
    signal.signal(signal.SIGINT, signal_handler)

    management_thread_container = []
    root, update_ui, status_label, batch_size, cycle_interval_entry = create_ui(vm_names, cycle_interval, update_queue, mumu_manager, management_thread_container, total_instances)
    
    management_thread = threading.Thread(target=run_management, args=(update_queue, mumu_manager, active_vm_indices, cycle_interval, batch_size))
    management_thread_container.append(management_thread)

    management_thread.start()
    update_ui()

    root.mainloop()
    global_should_stop = True
    if management_thread_container and management_thread_container[0].is_alive():
        management_thread_container[0].join(timeout=10.0)
    print("Programa encerrado.")
    sys.exit(0)

def run_management(update_queue, mumu_manager, active_vm_indices, cycle_interval, batch_size_var=None):
    global global_should_stop, is_cycling
    last_routine_run = None
    last_routine_range = None
    current_cycle_interval = cycle_interval  # Track current cycle interval
    vm_info = get_vm_info(os.path.dirname(os.path.dirname(mumu_manager)))
    last_vm_info_update = 0
    current_index = 0  # Track position in active_vm_indices
    is_retry_attempt = False # New state variable for retry logic

    print("Iniciando gerenciamento de instâncias...")
    while not global_should_stop:
        try:
            current_time = time.time()
            
            # Adaptive VM info update frequency: more frequent when cycling, less when idle
            update_interval = 2 if is_cycling else 5  # 2s when cycling, 5s when idle
            if current_time - last_vm_info_update >= update_interval:
                vm_info = get_vm_info(os.path.dirname(os.path.dirname(mumu_manager)))
                last_vm_info_update = current_time

            # Process queue for control messages
            while not update_queue.empty():
                data = update_queue.get_nowait()
                if data.get("reset_index"):
                    current_index = 0
                if "last_routine_run" in data and data["last_routine_run"] is None:
                    last_routine_run = None
                    last_routine_range = None
                # Update cycle interval if provided
                if "cycle_interval" in data:
                    current_cycle_interval = data["cycle_interval"]
                    print(f"Cycle interval updated to: {current_cycle_interval} seconds")
                # Handle complete reset
                if data.get("reset_cycle"):
                    last_routine_run = None
                    last_routine_range = None
                    current_index = 0
                    print("Cycle state reset complete")

            # Only send updates when cycling or when there are changes
            if is_cycling or current_time - last_vm_info_update < 1:
                update_queue.put({
                    "last_routine_run": last_routine_run,
                    "last_routine_range": last_routine_range,
                    "current_time": current_time,
                    "status": "Em Execução" if is_cycling else "Parado",
                    "vm_info": vm_info
                })

            if is_cycling and (last_routine_run is None or (current_time - last_routine_run) >= current_cycle_interval):
                print(f"Cycle time reached. Current time: {current_time}, Last run: {last_routine_run}, Interval: {current_cycle_interval}")
                
                # 1. Stop all currently running instances
                all_vms_info = get_vm_info(os.path.dirname(os.path.dirname(mumu_manager)))
                running_vms = [int(vm["index"]) for vm in all_vms_info if vm["status"] == "running" and not vm["is_main"]]
                if running_vms:
                    print(f"Stopping running VMs: {running_vms}")
                    control_instances(mumu_manager, running_vms, SelectionActions.STOP)
                    time.sleep(5) # Give time for VMs to shut down

                # 2. Determine the next batch to start
                try:
                    batch_size = batch_size_var.get() if batch_size_var else 1
                except Exception:
                    batch_size = 1
                
                if not isinstance(batch_size, int) or batch_size <= 0:
                    batch_size = 1
                
                vm_indices_to_start = []
                for i in range(batch_size):
                    idx = active_vm_indices[(current_index + i) % len(active_vm_indices)]
                    vm_indices_to_start.append(idx)

                # 3. Start the new batch
                print(f"Attempting to start new batch: VMs {vm_indices_to_start}")
                control_instances(mumu_manager, vm_indices_to_start, SelectionActions.START)
                
                # 4. Verify the new batch started correctly
                print("Verifying startup status...")
                start_verification_time = time.time()
                batch_started_successfully = False
                
                while time.time() - start_verification_time < 90: # 90-second verification window
                    vm_info = get_vm_info(os.path.dirname(os.path.dirname(mumu_manager)))
                    
                    # Check for successfully started VMs
                    running_in_batch = [
                        int(vm['index']) for vm in vm_info 
                        if vm.get('is_android_started', False) and int(vm['index']) in vm_indices_to_start
                    ]
                    if len(running_in_batch) > 0:
                        print(f"Successfully started VMs: {running_in_batch}. Batch confirmed.")
                        batch_started_successfully = True
                        break # Exit verification loop on success

                    # Check if all VMs in the batch have errored out to fail fast
                    errored_vms = [
                        vm for vm in vm_info
                        if int(vm['index']) in vm_indices_to_start and vm.get('has_error')
                    ]
                    if len(errored_vms) == len(vm_indices_to_start):
                        print("CRITICAL: All VMs in the batch have reported a startup error. Failing fast.")
                        break # Exit verification loop on definitive failure
                    
                    print(f"Waiting for VMs to start... ({int(time.time() - start_verification_time)}s)")
                    time.sleep(5)

                # 5. Update state only if batch start was successful
                if batch_started_successfully:
                    print("Batch start confirmed. Resetting cycle timer.")
                    last_routine_run = time.time() # Use current time after verification
                    current_index = (current_index + batch_size) % len(active_vm_indices)
                    is_retry_attempt = False # Reset retry flag on success
                    
                    min_idx = min([active_vm_indices.index(idx) for idx in vm_indices_to_start]) + 1
                    max_idx = max([active_vm_indices.index(idx) for idx in vm_indices_to_start]) + 1
                    last_routine_range = (min_idx, max_idx)
                    
                    update_queue.put({
                        "cycle_completed": True,
                        "current_time": last_routine_run
                    })
                    print(f"Cycle completed. Next cycle in {current_cycle_interval} seconds")
                else:
                    # The batch failed to start. As a safety measure, ensure all non-main VMs are stopped.
                    print("Ensuring all non-main VMs are stopped after a batch start failure.")
                    all_vms_info = get_vm_info(os.path.dirname(os.path.dirname(mumu_manager)))
                    vms_to_stop = [int(vm["index"]) for vm in all_vms_info if vm["status"] == "running" and not vm["is_main"]]
                    if vms_to_stop:
                        print(f"Stopping lingering VMs: {vms_to_stop}")
                        control_instances(mumu_manager, vms_to_stop, SelectionActions.STOP)

                    if is_retry_attempt:
                        # The retry attempt also failed. Give up and move immediately to the next batch.
                        print("CRITICAL: Retry failed. Moving immediately to the next batch.")
                        current_index = (current_index + batch_size) % len(active_vm_indices)
                        is_retry_attempt = False # Reset for the next batch
                        # By leaving last_routine_run unchanged, the main loop will immediately attempt the next batch.
                        print(f"State after failed retry: is_retry_attempt={is_retry_attempt}, next_index will be {current_index}")
                    else:
                        # First failure. Set flag to retry once.
                        print("CRITICAL: No VMs in the batch started or all errored. Will retry same batch once.")
                        is_retry_attempt = True
            
            # Adaptive sleep: shorter when cycling, longer when idle
            sleep_time = 0.5 if is_cycling else 1.0
            time.sleep(sleep_time)
        except Exception as e:
            print(f"Erro no loop principal: {e}")
            update_queue.put({
                "last_routine_run": last_routine_run,
                "last_routine_range": last_routine_range,
                "current_time": current_time,
                "status": f"Erro - {str(e)}",
                "vm_info": vm_info
            })
            time.sleep(2)

    print("Limpando...")
    update_queue.put({
        "last_routine_run": last_routine_run,
        "last_routine_range": last_routine_range,
        "current_time": time.time(),
        "status": "Desligando",
        "vm_info": vm_info
    })
    try:
        running_vms = [int(vm["index"]) for vm in vm_info if vm["status"] == "running" and not vm["is_main"]]
        if running_vms:
            control_instances(mumu_manager, running_vms, SelectionActions.STOP)
            print("Todas as instâncias paradas.")
    except Exception as e:
        print(f"Erro na limpeza: {e}")

if __name__ == "__main__":
    main()