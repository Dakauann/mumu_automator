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
                check=True
            )
            vm_data = json.loads(result.stdout)
            if vm_data.get("error_code", -1) == 0:
                status = "running" if vm_data.get("is_android_started", False) else "stopped"
                return {
                    "index": vm_data["index"],
                    "name": vm_data["name"],
                    "status": status,
                    "is_main": vm_data.get("is_main", False)
                }
            return None
        except (subprocess.CalledProcessError, json.JSONDecodeError):
            return None

    with ThreadPoolExecutor(max_workers=10) as executor:
        future_to_index = {executor.submit(query_vm, i): i for i in indices}
        for future in as_completed(future_to_index):
            result = future.result()
            if result:
                vm_info.append(result)

    vm_info.sort(key=lambda x: x['index'])  # Sort by index for consistent order
    print(f"Tempo decorrido: {time.time() - start_time} segundos")
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
    """Control VM instances one at a time using MuMuManager."""
    if not vm_indices:
        print("Nenhum índice de VM fornecido para controle.")
        return False
    success = True
    for idx in vm_indices:
        try:
            subprocess.run(
                [mumu_manager, "control", action, "-v", str(idx)],
                capture_output=True,
                text=True,
                check=True
            )
            print(f"Executado com sucesso {action} para VM {idx}")
        except subprocess.CalledProcessError as e:
            print(f"Erro ao executar {action} para VM {idx}: {e}")
            success = False
    return success

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
        """Handle window close event."""
        global global_should_stop, is_cycling
        nonlocal stop_ui_update, after_id
        print("Fechando janela, encerrando programa...")
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

    root.protocol("WM_DELETE_WINDOW", close_window)

    def update_ui():
        nonlocal last_vm_info, stop_ui_update, after_id, cycle_count, total_cycles_needed
        if stop_ui_update or global_should_stop:
            return
        try:
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
                        status_pt = "Em Execução" if vm["status"] == "running" else "Parado"
                        instance_listbox.insert(tk.END, f"VM {vm['index']}: {vm['name']} ({status_pt})")
                else:
                    instance_listbox.insert(tk.END, "Nenhuma VM disponível")

                if vm_info and vm_info != last_vm_info:
                    running_count = sum(1 for vm in vm_info if vm["status"] == "running")
                    times.append(current_time)
                    running_counts.append(running_count)
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
            after_id = root.after(50, update_ui)

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

    print("Iniciando gerenciamento de instâncias...")
    while not global_should_stop:
        try:
            current_time = time.time()
            if current_time - last_vm_info_update >= 1:
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

            update_queue.put({
                "last_routine_run": last_routine_run,
                "last_routine_range": last_routine_range,
                "current_time": current_time,
                "status": "Em Execução" if is_cycling else "Parado",
                "vm_info": vm_info
            })

            if is_cycling and (last_routine_run is None or (current_time - last_routine_run) >= current_cycle_interval):
                print(f"Cycle time reached. Current time: {current_time}, Last run: {last_routine_run}, Interval: {current_cycle_interval}")
                
                # Stop all running instances before starting new batch
                running_vms = [int(vm["index"]) for vm in vm_info if vm["status"] == "running" and not vm["is_main"]]
                if running_vms:
                    print(f"Stopping running VMs: {running_vms}")
                    update_queue.put({
                        "last_routine_run": last_routine_run,
                        "last_routine_range": last_routine_range,
                        "current_time": current_time,
                        "status": f"Parando todas as VMs em execução: {running_vms}",
                        "vm_info": vm_info
                    })
                    control_instances(mumu_manager, running_vms, SelectionActions.STOP)
                    time.sleep(5)
                else:
                    print("No running VMs to stop")

                try:
                    batch_size = batch_size_var.get() if batch_size_var else 1
                except Exception:
                    batch_size = 1
                print(f"Batch size from UI: {batch_size}")
                
                if not isinstance(batch_size, int) or batch_size <= 0:
                    batch_size = 1  # Fallback to avoid zero/negative batch size
                vm_indices = []
                start_point = current_index
                for i in range(batch_size):
                    idx = active_vm_indices[(current_index + i) % len(active_vm_indices)]
                    vm_indices.append(idx)
                current_index = (current_index + batch_size) % len(active_vm_indices)

                if vm_indices:
                    min_idx = min([active_vm_indices.index(idx) for idx in vm_indices]) + 1
                    max_idx = max([active_vm_indices.index(idx) for idx in vm_indices]) + 1
                    end_point = max_idx if max_idx >= min_idx else max_idx + len(active_vm_indices)
                else:
                    min_idx = start_point + 1
                    end_point = start_point + 1

                print(f"Starting new batch: VMs {vm_indices} (batch {min_idx}-{end_point})")
                update_queue.put({
                    "last_routine_run": last_routine_run,
                    "last_routine_range": last_routine_range,
                    "current_time": current_time,
                    "status": f"Iniciando lote {min_idx}-{end_point}",
                    "batch_size": batch_size,
                    "vm_info": vm_info
                })
                print(f"Starting VMs: {vm_indices}")  # Debug print
                control_instances(mumu_manager, vm_indices, SelectionActions.START)
                last_routine_run = current_time
                last_routine_range = (min_idx, end_point)
                
                # Send cycle completion signal
                update_queue.put({
                    "cycle_completed": True,
                    "current_time": current_time
                })
                print(f"Cycle completed. Next cycle in {current_cycle_interval} seconds")

            time.sleep(0.1)
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