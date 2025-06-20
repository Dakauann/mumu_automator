import ctypes
from ctypes import wintypes
import json
import os
import threading
import time
import signal
import sys
from pywinauto import Application
from pywinauto.mouse import click, move
from pywinauto.timings import Timings
import random
import string
import enum
import win32con

# Increase default timeout
Timings.window_find_timeout = 10.0  # Increased for reliability

class SelectionActions(enum.Enum):
    START = 1
    STOP = 2

class SelectionPoints(enum.Enum):
    FROM_START = 1
    TO_END = 2
    FROM_MID = 3
    TO_START = 4
    TO_MID = 5

global_should_stop = False

def signal_handler(sig, frame):
    """Handle Ctrl+C and other termination signals."""
    global global_should_stop
    print("\nReceived Ctrl+C, stopping gracefully...")
    global_should_stop = True

def main():
    global global_should_stop

    last_path_file = "last_mumu_path.json"
    mumu_base_path = None
    cycle_interval = 480
    
    try:
        with open(last_path_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        prev_path = data.get("mumu_base_path", "").strip()
        if prev_path and os.path.isdir(prev_path):
            print(f"Previously used MuMu installation path: {prev_path}")
            choice = input("Use path? (0: reuse, 1: change path):").strip()
            if choice == "0":
                mumu_base_path = prev_path
                mumu_path = os.path.join(prev_path, "shell", "MuMuMultiPlayer.exe")
    except Exception as e:
        print(f"Could not load previous path: {e}")

    while not mumu_base_path:
        mumu_base_path = input(
            "Enter MuMu installation path (e.g., D:\\Program Files\\Netease\\MuMuPlayerGlobal-12.0):\n> "
        ).strip('"').strip()
        if not mumu_base_path:
            print("Path cannot be empty.")
            continue
        if not os.path.isdir(mumu_base_path):
            print(f"Directory not found: {mumu_base_path}")
            mumu_base_path = None
            continue
        mumu_path = os.path.join(mumu_base_path, "shell", "MuMuMultiPlayer.exe")
        if not os.path.exists(mumu_path):
            print(f"MuMuMultiPlayer.exe not found: {mumu_path}")
            mumu_base_path = None
            continue
        try:
            with open(last_path_file, "w", encoding="utf-8") as f:
                json.dump({"mumu_base_path": mumu_base_path}, f)
            print(f"Saved path to {last_path_file}")
        except Exception as e:
            print(f"Could not save path: {e}")
        break

    print(f"Using MuMu executable: {mumu_path}")

    vm_base_name = "ROM_"
    print("Scanning for VM names...")
    vm_names = padronize_vm_names(vm_base_name, mumu_base_path)
    print(f"Found {len(vm_names)} VM(s): {vm_names}")
    last_routine_run = None
    last_routine_range = None
    batch_size = int(input("Enter batch size: "))
    total_instances = len(vm_names)
    
    if not vm_names:
        print("Error: No VM names found. Please create VMs in MuMu Multi-Instance Manager.")
        return

    signal.signal(signal.SIGINT, signal_handler)

    try:
        app = Application(backend="uia").connect(title="MuMu Multi-instance 12", timeout=5)
        print("MuMu Multi-Instance running. Closing...")
        app.window(title="MuMu Multi-instance 12").close()
        time.sleep(2)
    except Exception:
        print("MuMu Multi-Instance not running.")

    print("Launching program...")
    try:
        os.startfile(mumu_path)
        print("Program launched. Waiting 10 seconds...")
        time.sleep(10)
        app = Application(backend="uia").connect(title="MuMu Multi-instance 12", timeout=10)
        print("MuMu Multi-Instance window open.")
    except FileNotFoundError:
        print(f"Error: Executable not found at {mumu_path}.")
        return
    except Exception as e:
        print(f"Error launching: {e}")
        return

    main_window = app.window(title="MuMu Multi-instance 12")
    main_window.set_focus()

    main_window_rect = main_window.rectangle()
    center_x = main_window_rect.left + main_window_rect.width() // 2
    center_y = main_window_rect.top + main_window_rect.height() // 2
    move(coords=(center_x, center_y))

    print("Starting instance management...")
    
    while not global_should_stop:
        try:
            current_time = time.time()
            print(f"\nCurrent time: {time.ctime(current_time)}")
            print(f"Last run: {last_routine_run}, Range: {last_routine_range}")
            print(f"Time remaining: {cycle_interval - (current_time - last_routine_run) if last_routine_run else 'N/A'} seconds")
            print("-" * 60)
            if last_routine_run is None or (current_time - last_routine_run) >= cycle_interval:
                if last_routine_range:
                    print(f"Stopping batch: {last_routine_range}")
                    start_instances_routine(main_window, vm_names, 
                                           last_routine_range[0], 
                                           last_routine_range[1], 
                                           action=SelectionActions.STOP)
                    time.sleep(10)

                if last_routine_range is None:
                    start_point = 1
                    end_point = min(batch_size, total_instances)
                else:
                    start_point = last_routine_range[1] + 1
                    end_point = min(start_point + batch_size - 1, total_instances)

                if start_point > total_instances:
                    print("All instances processed. Restarting.")
                    start_point = 1
                    end_point = min(batch_size, total_instances)

                start_instances_routine(main_window, vm_names, start_point, end_point, 
                                       action=SelectionActions.START)
                last_routine_run = current_time
                last_routine_range = (start_point, end_point)

            time.sleep(1)
        except Exception as e:
            print(f"Main loop error: {e}")
            time.sleep(2)

    print("Cleaning up...")
    try:
        if app.is_process_running():
            app.window(title="MuMu Multi-instance 12").close()
            print("Closed MuMu window.")
    except Exception as e:
        print(f"Cleanup error: {e}")
    
    print("Program terminated.")
    sys.exit(0)

def start_instances_routine(main_window, vm_names, start_point=1, end_point=None, action=SelectionActions.START):
    global global_should_stop
    if global_should_stop:
        print("Termination signal received, stopping routine.")
        return

    if not vm_names:
        print("Error: No VM names provided.")
        return

    if end_point is None:
        end_point = len(vm_names)

    print(f"Processing instances {start_point} to {end_point} with action {action.name}")

    if not main_window.exists(timeout=5):
        print("Main window not found. Exiting routine.")
        return
    
    unselect_all_instances(main_window)

    main_window.set_focus()
    time.sleep(0.5)

    instance_list = main_window.child_window(class_name="PlayerListWidget", control_type="List")
    if not instance_list.exists(timeout=5):
        print("Instance list not found.")
        return

    for idx in range(start_point - 1, end_point):
        if global_should_stop:
            print("Termination signal received, stopping routine.")
            return

        if idx >= len(vm_names):
            print(f"Reached end of VM names at index {idx}. Stopping.")
            break

        vm_name = vm_names[idx]
        print(f"Processing instance {idx + 1}: {vm_name}")

        search_bar = main_window.child_window(control_type="Edit", class_name="SearchEdit")
        if not search_bar.exists(timeout=5):
            print("Search bar not found.")
            continue

        try:
            search_bar.set_focus()
            time.sleep(0.2)
            search_bar.set_edit_text("")
            time.sleep(0.2)
            search_bar.set_text(vm_name.upper())
            time.sleep(1)

            instances = instance_list.children(control_type="ListItem")
            if instances:
                first_instance = instances[0]
                row_rect = first_instance.rectangle()
                if row_rect:
                    checkbox_x = row_rect.left + 36 + 15
                    checkbox_y = row_rect.top + 28
                    click(coords=(checkbox_x, checkbox_y))
                    print(f"Clicked checkbox for instance: {vm_name}")
                else:
                    print(f"Could not get rectangle for instance: {vm_name}")
            else:
                print(f"No instances found for VM: {vm_name}")
        except Exception as e:
            print(f"Error processing {vm_name} via search bar: {str(e)}")
            continue

        time.sleep(2)

    try:
        search_bar.set_edit_text("")
        print("Cleared search bar after processing instances.")
    except Exception as e:
        print(f"Error clearing search bar: {str(e)}")

    time.sleep(1)
    find_and_inspect_toolbar(main_window, action=action)
    
def click_toolbar_action(main_window, action: SelectionActions):
    global global_should_stop
    if global_should_stop:
        print("Termination signal received, skipping toolbar action.")
        return

    toolbar = main_window.child_window(control_type="Group", class_name="QWidget")
    if not toolbar.exists(timeout=5):
        print("Toolbar not found.")
        return

    buttons = toolbar.children(control_type="Button")
    if not buttons or len(buttons) < 2:
        print("Not enough buttons found in toolbar.")
        return

    try:
        if action == SelectionActions.START:
            buttons[0].click()
            print("Clicked 'Start' button in toolbar.")
        elif action == SelectionActions.STOP:
            buttons[1].click()
            print("Clicked 'Stop' button in toolbar. Waiting for confirmation dialog...")
            time.sleep(2)

            app = Application(backend="uia").connect(title_re=".*", timeout=5)
            dialog = app.window(class_name="NemuMessageBox")
            max_attempts = 3
            attempt = 1

            while attempt <= max_attempts:
                if dialog.exists(timeout=5):
                    print("Confirmation dialog detected.")
                    dialog_buttons = dialog.children(control_type="Button")
                    if dialog_buttons:
                        confirm_button = dialog_buttons[0]
                        try:
                            confirm_button.click()
                            print("Clicked 'Confirm' button in dialog.")
                            time.sleep(1)
                            return
                        except Exception as e:
                            print(f"Error clicking Confirm button: {str(e)}")
                    else:
                        print("No buttons found in dialog.")
                else:
                    print(f"Confirmation dialog not found on attempt {attempt}.")
                
                time.sleep(1)
                attempt += 1

            print(f"Failed to find or click 'Confirm' button after {max_attempts} attempts.")
    except Exception as e:
        print(f"Error clicking toolbar button for {action.name}: {str(e)}")

def unselect_all_instances(main_window):
    print("Attempting to unselect all instances...")
    select_all_checkbox = main_window.child_window(control_type="CheckBox", title="Select All")
    max_attempts = 3
    attempt = 1

    while attempt <= max_attempts:
        try:
            if select_all_checkbox.exists(timeout=5):
                print("Found 'Select All' checkbox.")
                # First click to unselect all
                select_all_checkbox.click()
                print("Clicked 'Select All' checkbox to unselect all instances.")
                time.sleep(1)
                # Second click to select all
                select_all_checkbox.click()
                print("Clicked 'Select All' checkbox to select all instances.")
                time.sleep(1)
                return True
            else:
                print(f"'Select All' checkbox not found on attempt {attempt}.")
                if attempt == max_attempts:
                    print("Debugging UI hierarchy...")
                    list_elements_on_window(main_window)
        except Exception as e:
            print(f"Error interacting with 'Select All' checkbox on attempt {attempt}: {str(e)}")
        
        time.sleep(2)
        attempt += 1

    print("Failed to find or click 'Select All' checkbox after all attempts.")
    return False

def manage_instances_searchbar(vm_names, main_window, start_point=1, end_point=None):
    """
    Start searching for instances in the search bar and select them.
    """
    search_bar = main_window.child_window(control_type="Edit", class_name="SearchEdit")
    if not search_bar.exists(timeout=5):
        print("Search bar not found.")
        return

    # make sure the window is focused
    main_window.set_focus()

    print("Starting instance selection process...")
    for idx, vm_name in enumerate(vm_names):
        if end_point is not None and idx >= end_point:
            print(f"Reached end point at index {idx}. Stopping further searches.")
            break

        if start_point is not None and idx < start_point - 1:
            print(f"Skipping index {idx} as it is below the start point {start_point}.")
            continue

        print(f"Searching for VM {idx + 1}: {vm_name.upper()}")
        try:
            search_bar.set_focus()
            time.sleep(0.2)
            search_bar.set_edit_text("")  # Clear existing text
            time.sleep(0.2)
            search_bar.set_text(vm_name.upper())
            time.sleep(1)  # Allow time for search results to update

            instance_list = main_window.child_window(class_name="PlayerListWidget", control_type="List")
            if instance_list.exists():
                instances = instance_list.children(control_type="ListItem")
                if instances:
                    first_instance = instances[0]
                    row_rect = first_instance.rectangle()
                    if row_rect:
                        checkbox_x = row_rect.left + 36 + 15
                        checkbox_y = row_rect.top + 28
                        click(coords=(checkbox_x, checkbox_y))
                        print(f"Clicked checkbox for instance: {vm_name}")
                    else:
                        print(f"Could not get rectangle for first instance of VM: {vm_name}")
                else:
                    print(f"No instances found for VM: {vm_name}")
            else:
                print("Instance list not found.")
        except Exception as e:
            print(f"Error focusing or typing in search bar: {e}")
            continue

        time.sleep(1)

def padronize_vm_names(vm_name_prefix, mumu_base_path) -> list[str]:
    import os
    import json
    vm_dir = os.path.join(mumu_base_path, "vms")
    vm_names = []
    used_names = set()

    def generate_unique_name():
        while True:
            suffix = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
            name = f"{vm_name_prefix}{suffix}"
            if name not in used_names:
                used_names.add(name)
                return name

    if not os.path.isdir(vm_dir):
        print(f"Error: VM directory not found at {vm_dir}. Please check MuMu installation.")
        return vm_names

    found_valid_dir = False
    for dir_name in os.listdir(vm_dir):
        dir_path = os.path.join(vm_dir, dir_name)
        # Skip non-VM directories (e.g., MuMuPlayerGlobal-12.0-base)
        if not os.path.isdir(dir_path) or not dir_name.startswith("MuMuPlayerGlobal-12.0-") or "base" in dir_name.lower():
            print(f"Skipping directory: {dir_path} (not a valid VM directory)")
            continue

        config_path = os.path.join(dir_path, "configs")
        config_file = os.path.join(config_path, "extra_config.json")
        if not os.path.exists(config_file):
            print(f"No config file found: {config_file}")
            continue

        found_valid_dir = True
        print(f"Found config file: {config_file}")
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                new_player_name = generate_unique_name()
                config_data["playerName"] = new_player_name
                with open(config_file, 'w', encoding='utf-8') as f2:
                    json.dump(config_data, f2, indent=2, ensure_ascii=False)
                vm_names.append(new_player_name)
                print(f"Updated and set unique VM name: {new_player_name}")
        except json.JSONDecodeError as e:
            print(f"Error reading JSON from {config_file}: {str(e)}")
        except Exception as e:
            print(f"Error updating {config_file}: {str(e)}")

    if not found_valid_dir:
        print("Warning: No valid VM directories found. Please create VMs in MuMu Multi-Instance Manager.")
        # Optional: Prompt user to continue with a default VM name for testing
        create_default = input("No VMs found. Use a default VM name for testing? (y/n): ").strip().lower()
        if create_default == 'y':
            vm_names.append(generate_unique_name())
            print(f"Added default VM name: {vm_names[0]}")

    return vm_names

def get_detailed_scroll_info(control):
    """
    Get detailed scrolling information for a control
    """
    scroll_info = {
        'is_scrollable': False,
        'scroll_patterns': [],
        'scroll_methods_available': [],
        'scroll_bars': {'horizontal': False, 'vertical': False},
        'error': None
    }
    
    try:
        # Check basic scrollable property
        scroll_info['is_scrollable'] = control.is_scrollable()
        
        # Try to get UIA patterns for scrolling
        try:
            element = control.element_info
            # Check for scroll pattern
            if hasattr(element, 'get_current_pattern'):
                try:
                    scroll_pattern = element.get_current_pattern(10016)  # ScrollPattern ID
                    if scroll_pattern:
                        scroll_info['scroll_patterns'].append('ScrollPattern')
                        # Get scroll properties
                        try:
                            h_scrollable = scroll_pattern.current_horizontally_scrollable
                            v_scrollable = scroll_pattern.current_vertically_scrollable
                            scroll_info['scroll_bars']['horizontal'] = h_scrollable
                            scroll_info['scroll_bars']['vertical'] = v_scrollable
                        except:
                            pass
                except:
                    pass
        except Exception as e:
            scroll_info['error'] = f"Pattern check error: {str(e)}"
        
        # Check for scroll methods availability
        scroll_methods = ['scroll', 'scroll_mouse', 'wheel_mouse_input']
        for method in scroll_methods:
            if hasattr(control, method):
                scroll_info['scroll_methods_available'].append(method)
        
        # Look for scroll bar controls as children
        try:
            scrollbars = control.descendants(control_type="ScrollBar")
            if scrollbars:
                for sb in scrollbars:
                    sb_name = sb.window_text() or sb.element_info.automation_id or "Unknown"
                    orientation = "horizontal" if "horizontal" in sb_name.lower() else "vertical"
                    scroll_info['scroll_bars'][orientation] = True
        except:
            pass
            
    except Exception as e:
        scroll_info['error'] = f"General error: {str(e)}"
    
    return scroll_info

def list_elements_on_window(window):
    def print_detailed_tree(control, depth=0):
        indent = "  " * depth
        # Gather detailed properties
        control_type = control.element_info.control_type or "[No ControlType]"
        control_name = control.window_text() or "[No Name]"
        automation_id = control.element_info.automation_id or "[No AutomationId]"
        class_name = control.element_info.class_name or "[No ClassName]"
        rectangle = control.rectangle() if control.rectangle() else "[No Rectangle]"
        enabled = "Enabled" if control.is_enabled() else "Disabled"
        visible = "Visible" if control.is_visible() else "Hidden"

        # Print inline summary
        print(f"{indent}- {control_type} | Name: '{control_name}' | Class: {class_name} | AutomationId: {automation_id} | Rect: {rectangle} | {enabled} | {visible}")

        # Get detailed scroll information
        scroll_info = get_detailed_scroll_info(control)
        if scroll_info['is_scrollable']:
            print(f"{indent}  [Scrollable] Patterns: {', '.join(scroll_info['scroll_patterns'])} | Methods: {', '.join(scroll_info['scroll_methods_available'])} | ScrollBars: H={scroll_info['scroll_bars']['horizontal']} V={scroll_info['scroll_bars']['vertical']}")
        if scroll_info['error']:
            print(f"{indent}  [Scroll Debug Error] {scroll_info['error']}")

        # Special handling for List controls
        if control_type == "List":
            print(f"{indent}  *** LIST CONTROL DETECTED ***")
            print(f"{indent}  List Items Count: {len(control.children())}")
            try:
                list_rect = control.rectangle()
                if list_rect:
                    print(f"{indent}  List Dimensions: {list_rect.width()}x{list_rect.height()}")
                scrollbars = control.descendants(control_type="ScrollBar")
                if scrollbars:
                    print(f"{indent}  Found {len(scrollbars)} scroll bars in list")
                    for i, sb in enumerate(scrollbars):
                        sb_rect = sb.rectangle()
                        sb_orientation = "Horizontal" if sb_rect and sb_rect.width() > sb_rect.height() else "Vertical"
                        print(f"{indent}    ScrollBar {i+1}: {sb_orientation}, Rect: {sb_rect}")
                else:
                    print(f"{indent}  No scroll bars found in list")
            except Exception as e:
                print(f"{indent}  List analysis error: {str(e)}")

        # Recursively process children
        for child in control.children():
            print_detailed_tree(child, depth + 1)

    print("\nListing all controls in the MuMu Multi-Instance window with detailed properties:")
    print_detailed_tree(window)

def find_and_inspect_toolbar(window, action=SelectionActions.START):
    print("\nSearching for toolbar...")
    all_groups = window.descendants(control_type="Group")
    toolbar = None
    
    window.set_focus()

    for group in all_groups:
        rect = group.rectangle()
        if not rect:
            continue
        if rect.top >= 800 and rect.bottom <= 900:
            buttons = group.children(control_type="Button")
            if len(buttons) >= 5:
                toolbar = group
                print(f"Found toolbar at rectangle: {rect}")
                print(f"Toolbar contains {len(buttons)} buttons")
                break

    if not toolbar:
        print("Toolbar not found. Searching more broadly...")
        for group in all_groups:
            buttons = group.children(control_type="Button")
            if len(buttons) >= 6:
                toolbar = group
                rect = group.rectangle()
                print(f"Found toolbar candidate at rectangle: {rect}")
                print(f"Toolbar contains {len(buttons)} buttons")
                break

    if not toolbar:
        print("Toolbar still not found. Let me try a different approach...")
        try:
            main_container = window.child_window(class_name="QWidget")
            if main_container.exists():
                groups = main_container.descendants(control_type="Group")
                for group in reversed(groups):
                    buttons = group.children(control_type="Button")
                    if len(buttons) >= 5:
                        toolbar = group
                        rect = group.rectangle() if group.rectangle() else "Unknown"
                        print(f"Found toolbar using alternative method at: {rect}")
                        break
        except Exception as e:
            print(f"Error in alternative search: {e}")

    if not toolbar:
        print("Toolbar could not be found using any method.")
        return

    print(f"\nInspecting toolbar buttons...")
    buttons = toolbar.children(control_type="Button")

    if not buttons:
        print("No buttons found in toolbar.")
        return

    print(f"Found {len(buttons)} buttons in toolbar:")
    print("-" * 80)

    for idx, button in enumerate(buttons):
        try:
            button_rect = button.rectangle()
            button_name = button.window_text() or f"Button_{idx}"
            button_class = button.element_info.class_name or "[No ClassName]"
            automation_id = button.element_info.automation_id or "[No AutomationId]"
            is_enabled = button.is_enabled()
            is_visible = button.is_visible()

            if idx == 0:
                logical_name = "START (first button)"
                if action == SelectionActions.START:
                    button.click()
            elif idx == 1:
                logical_name = "STOP (second button)"
                if action == SelectionActions.STOP:
                    button.click()
                    print("Clicked 'Stop' button. Waiting for confirmation dialog...")
                    time.sleep(6)

                    list_elements_on_window(window)
                    print("Searching for dialog in current window hierarchy...")
                    dialogs = window.descendants(class_name="NemuMessageBox")
                    if dialogs:
                        dialog = dialogs[0]

                        def find_and_click_pushbutton7(control):
                            if control.element_info.class_name == "NemuUiLib::NemuPushButton7":
                                try:
                                    control.click()
                                    print("Clicked the first NemuPushButton7 in the dialog.")
                                    return True
                                except Exception as e:
                                    print(f"Error clicking NemuPushButton7: {e}")
                                    return False
                            for child in control.children():
                                if find_and_click_pushbutton7(child):
                                    return True
                            return False

                        if not find_and_click_pushbutton7(dialog):
                            print("NemuPushButton7 not found or could not be clicked.")
                    else:
                        print("No NemuMessageBox found in current window hierarchy.")
                        return

                    max_attempts = 5
                    attempt = 1

                    while attempt <= max_attempts:
                        if dialog.exists(timeout=5):
                            print("Confirmation dialog detected.")
                            dialog.set_focus()
                            dialog_buttons = dialog.children(control_type="Button")
                            if dialog_buttons:
                                confirm_button = dialog_buttons[0]
                                if confirm_button.is_enabled() and confirm_button.is_visible():
                                    try:
                                        confirm_button.click()
                                        print("Clicked 'Confirm' button in dialog.")
                                        time.sleep(1)
                                        return
                                    except Exception as e:
                                        print(f"Error clicking Confirm button: {str(e)}. Retrying...")
                                else:
                                    print("Confirm button not enabled or visible.")
                            else:
                                print("No buttons found in dialog.")
                        else:
                            print(f"Confirmation dialog not found on attempt {attempt}.")
                        
                        time.sleep(1)
                        attempt += 1

                    print(f"Failed to find or click 'Confirm' button after {max_attempts} attempts.")
            else:
                logical_name = ""

            print(f"Button {idx + 1}: {logical_name}")
            print(f"  Name: {button_name}")
            print(f"  Class: {button_class}")
            print(f"  AutomationId: {automation_id}")
            print(f"  Rectangle: {button_rect}")
            print(f"  Enabled: {'Yes' if is_enabled else 'No'}")
            print(f"  Visible: {'Yes' if is_visible else 'No'}")
            print(f"  Position: ({button_rect.left}, {button_rect.top})" if button_rect else "Unknown")
            print(f"  Size: {button_rect.width()}x{button_rect.height()}" if button_rect else "Unknown")
            print("-" * 40)
        except Exception as e:
            print(f"Error inspecting button {idx + 1}: {str(e)}")
            continue

def analyze_instance_state(instance_row, index, main_window):
    """
    Analyze the state of a MuMu instance to determine if it's running and if checkbox is checked.
    
    Args:
        instance_row: The ListItem control representing the instance
        index: The index of the instance
        main_window: The main window to search for additional controls
    
    Returns:
        dict: Contains 'name', 'running_state', and 'checkbox_state'
    """
    instance_info = {
        'name': f'Instance_{index}',
        'running_state': 'Unknown',
        'checkbox_state': 'Unknown'
    }
    
    try:
        # Get all text elements within this instance row
        text_elements = instance_row.descendants(control_type="Text")
        
        # Look for status text
        status_found = False
        for text_elem in text_elements:
            text_content = text_elem.window_text()
            if text_content:
                # Check for running indicators
                if "running" in text_content.lower():
                    instance_info['running_state'] = 'Running'
                    status_found = True
                elif "not started" in text_content.lower():
                    instance_info['running_state'] = 'Not Started'
                    status_found = True
                elif "stopped" in text_content.lower():
                    instance_info['running_state'] = 'Stopped'
                    status_found = True
                
                # Try to extract instance name
                if any(keyword in text_content.upper() for keyword in ['ROM', 'MUMU', 'PLAYER', 'CLONE']):
                    instance_info['name'] = text_content
        
        # Alternative method: Look for buttons within the instance row
        if not status_found:
            buttons = instance_row.descendants(control_type="Button")
            for button in buttons:
                button_rect = button.rectangle()
                if button_rect:
                    # Play buttons are typically on the right side of the row
                    row_rect = instance_row.rectangle()
                    if row_rect and button_rect.left > (row_rect.left + row_rect.width() * 0.7):
                        # Check if button is enabled/disabled or has specific properties
                        if button.is_enabled():
                            # Try to determine button type by position or class
                            button_class = button.element_info.class_name or ""
                            automation_id = button.element_info.automation_id or ""
                            
                            # This is a heuristic - you might need to adjust based on actual button properties
                            if "play" in button_class.lower() or "start" in button_class.lower():
                                instance_info['running_state'] = 'Not Started'
                            elif "stop" in button_class.lower() or "power" in button_class.lower():
                                instance_info['running_state'] = 'Running'
        
        # Check checkbox state
        # Look for checkbox controls within the main window that correspond to this instance
        row_rect = instance_row.rectangle()
        if row_rect:
            # Calculate expected checkbox position
            checkbox_x = row_rect.left + 36 + 15
            checkbox_y = row_rect.top + 28
            
            # Try to find checkbox controls near this position
            all_checkboxes = main_window.descendants(control_type="CheckBox")
            for checkbox in all_checkboxes:
                cb_rect = checkbox.rectangle()
                if cb_rect:
                    # Check if checkbox is in similar Y position (same row)
                    if abs(cb_rect.top - checkbox_y) < 20:  # Within 20 pixels
                        try:
                            # Check if checkbox is checked
                            if hasattr(checkbox, 'get_toggle_state'):
                                toggle_state = checkbox.get_toggle_state()
                                instance_info['checkbox_state'] = 'Checked' if toggle_state == 1 else 'Unchecked'
                            else:
                                # Alternative method using selection state
                                try:
                                    is_selected = checkbox.is_selected()
                                    instance_info['checkbox_state'] = 'Checked' if is_selected else 'Unchecked'
                                except:
                                    # Try using window text or other properties
                                    checkbox_text = checkbox.window_text()
                                    if checkbox_text and "checked" in checkbox_text.lower():
                                        instance_info['checkbox_state'] = 'Checked'
                                    else:
                                        instance_info['checkbox_state'] = 'Unchecked'
                        except Exception as e:
                            instance_info['checkbox_state'] = f'Error: {str(e)}'
                        break
        
        # Fallback: Try to detect state from visual elements positioning
        if instance_info['running_state'] == 'Unknown':
            # Look for specific UI elements that indicate running state
            all_elements = instance_row.descendants()
            for elem in all_elements:
                elem_class = elem.element_info.class_name or ""
                elem_text = elem.window_text() or ""
                
                # Look for specific MuMu UI indicators
                if "running" in elem_text.lower() or "run" in elem_class.lower():
                    instance_info['running_state'] = 'Running'
                    break
                elif "start" in elem_text.lower() or "play" in elem_class.lower():
                    instance_info['running_state'] = 'Not Started'
                    break
    
    except Exception as e:
        instance_info['running_state'] = f'Error: {str(e)}'
        instance_info['checkbox_state'] = f'Error: {str(e)}'
    
    return instance_info

if __name__ == "__main__":
    main()