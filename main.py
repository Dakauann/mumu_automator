import json
import os
import time
import signal
import sys
from pywinauto import Application
from pywinauto.mouse import click, move
from pywinauto.timings import Timings
from pywinauto.keyboard import send_keys
import random
import string
import enum

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

# Global flag to signal termination
global_should_stop = False

def signal_handler(sig, frame):
    """Handle Ctrl+C and other termination signals."""
    global global_should_stop
    print("\nReceived Ctrl+C, stopping gracefully...")
    global_should_stop = True

def main():
    global global_should_stop

    # Persistent storage for last used path in a JSON file
    last_path_file = "last_mumu_path.json"
    mumu_base_path = None
    # 40 minutes cycle interval: 2400 seconds
    cycle_interval = 2400  # seconds

    # Try to load last used path from JSON
    if os.path.isfile(last_path_file):
        try:
            with open(last_path_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            prev_path = data.get("mumu_base_path", "").strip()
            if prev_path and os.path.isdir(prev_path):
                print(f"Previously used MuMu installation path: {prev_path}")
                choice = input("Use this path? (0: reuse, 1: change path)\n> ").strip()
                if choice == "0":
                    mumu_base_path = prev_path
                    mumu_path = os.path.join(mumu_base_path, "shell", "MuMuMultiPlayer.exe")
        except Exception as e:
            print(f"Could not load previous path from JSON: {e}")

    # Prompt user for MuMu Multi-Instance executable path if not reusing
    while not mumu_base_path:
        mumu_base_path = input(
            "Enter the base path to MuMu installation (e.g., D:\\Program Files\\Netease\\MuMuPlayerGlobal-12.0):\n> "
        ).strip('"').strip()
        if not mumu_base_path:
            print("Path cannot be empty. Please try again.")
            continue
        if not os.path.isdir(mumu_base_path):
            print(f"Directory not found at: {mumu_base_path}")
            mumu_base_path = None
            continue
        mumu_path = os.path.join(mumu_base_path, "shell", "MuMuMultiPlayer.exe")
        if not os.path.isfile(mumu_path):
            print(f"MuMuMultiPlayer.exe not found at: {mumu_path}")
            mumu_base_path = None
            continue
        # Save path for next time in JSON
        try:
            with open(last_path_file, "w", encoding="utf-8") as f:
                json.dump({"mumu_base_path": mumu_base_path}, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Could not save path to JSON: {e}")
        break

    print(f"Using MuMu Multi-Instance executable at: {mumu_path}")

    vm_base_name = "ROM_"
    vm_names = padronize_vm_names(vm_base_name, mumu_base_path)
    if not vm_names:
        print("Error: No VM names found. Exiting.")
        return

    # Register signal handler for Ctrl+C
    signal.signal(signal.SIGINT, signal_handler)

    # Try to close MuMu Multi-Instance if already running
    try:
        app = Application(backend="uia").connect(title="MuMu Multi-instance 12", timeout=5)
        print("MuMu Multi-Instance is already running. Closing it first...")
        app.window(title="MuMu Multi-instance 12").close()
        time.sleep(2)
    except Exception:
        print("MuMu Multi-Instance is not running or could not connect.")

    # Start MuMu Multi-Instance
    print("Attempting to launch the program...")
    try:
        os.startfile(mumu_path)
        print("Program launched successfully. Waiting 10 seconds for it to start...")
        time.sleep(10)
        app = Application(backend="uia").connect(title="MuMu Multi-instance 12", timeout=10)
        print("MuMu Multi-Instance window is now open.")
    except FileNotFoundError:
        print(f"Error: The MuMu Multi-Instance executable was not found at {mumu_path}.")
        return
    except Exception as e:
        print(f"Error launching MuMu Multi-Instance: {str(e)}")
        return

    main_window = app.window(title="MuMu Multi-instance 12")
    main_window.set_focus()

    main_window_rect = main_window.rectangle()
    center_x = main_window_rect.left + main_window_rect.width() // 2
    center_y = main_window_rect.top + main_window_rect.height() // 2
    move(coords=(center_x, center_y))

    # Uncomment to debug UI hierarchy
    # list_elements_on_window(main_window)

    # Start the routine to start/stop instances
    print("Starting instance management routine...")
    last_routine_run = None
    last_routine_range = None
    batch_size = int(input("Enter batch size: "))
    total_instances = len(vm_names)

    # Main loop replacing threading.Timer
    while not global_should_stop:
        try:
            current_time = time.time()
            print(f"\nCurrent time: {time.ctime(current_time)}")
            print(f"Last routine run: {last_routine_run}, Last routine range: {last_routine_range}")
            # tome for next routine run
            print(f"Time remaining for next routine run: {cycle_interval - (current_time - last_routine_run) if last_routine_run else 'N/A'} seconds")
            print("--------------------------------------------------------")
            if last_routine_run is None or (current_time - last_routine_run) >= cycle_interval:  # 10-minute interval
                # Stop previous batch if it exists
                if last_routine_range:
                    print(f"Stopping previous batch: {last_routine_range}")
                    start_instances_routine(main_window, vm_names, 
                                           last_routine_range[0], 
                                           last_routine_range[1], 
                                           action=SelectionActions.STOP)
                    time.sleep(10)  # Wait for stop to complete

                # Calculate new batch
                if last_routine_range is None:
                    start_point = 1
                    end_point = min(batch_size, total_instances)
                else:
                    start_point = last_routine_range[1] + 1
                    end_point = min(start_point + batch_size - 1, total_instances)

                if start_point > total_instances:
                    print("All instances processed. Restarting from beginning.")
                    start_point = 1
                    end_point = min(batch_size, total_instances)

                # Start the new batch
                start_instances_routine(main_window, vm_names, start_point, end_point, 
                                       action=SelectionActions.START)
                # find_and_inspect_toolbar(main_window, action=SelectionActions.START)

                last_routine_run = current_time
                last_routine_range = (start_point, end_point)

            # Short sleep to prevent CPU overuse and allow Ctrl+C to be processed
            time.sleep(1)
        except Exception as e:
            print(f"Error in main loop: {str(e)}")
            time.sleep(2)  # Prevent rapid error looping

    # Cleanup before exiting
    print("Performing cleanup...")
    try:
        if app.is_process_running():
            app.window(title="MuMu Multi-instance 12").close()
            print("Closed MuMu Multi-Instance window.")
    except Exception as e:
        print(f"Error during cleanup: {str(e)}")
    
    print("Program terminated.")
    sys.exit(0)

def start_instances_routine(main_window, vm_names, start_point=1, end_point=None, action=SelectionActions.START):
    """
    Routine to start or stop instances in batches.
    """
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

    # wait for the main window to be ready
    if not main_window.exists(timeout=5):
        print("Main window not found. Exiting routine.")
        return
    
    unselect_all_instances(main_window)

    # Ensure the main window is focused
    main_window.set_focus()
    time.sleep(0.5)  # Allow time for focus to settle

    # Ensure the instance list is accessible
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

        # Use search bar to locate the instance
        search_bar = main_window.child_window(control_type="Edit", class_name="SearchEdit")
        if not search_bar.exists(timeout=5):
            print("Search bar not found.")
            continue

        try:
            search_bar.set_focus()
            time.sleep(0.2)
            search_bar.set_edit_text("")  # Clear existing text
            time.sleep(0.2)
            send_keys(vm_name.upper(), with_spaces=True)
            time.sleep(1)  # Allow search results to update

            # Click the checkbox of the first result
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

        time.sleep(2)  # Wait before next instance

    # in the end will clear the search bar
    try:
        search_bar.set_edit_text("")  # Clear the search bar after processing
        print("Cleared search bar after processing instances.")
    except Exception as e:
        print(f"Error clearing search bar: {str(e)}")

    # wait 1 second before clicking the toolbar action
    time.sleep(1)
    # Click the toolbar action (Start or Stop)
    # click_toolbar_action(main_window, action)
    # print(f"Finished {action.name} routine for instances {start_point} to {end_point}")
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
            buttons[0].click_input()
            print("Clicked 'Start' button in toolbar.")
        elif action == SelectionActions.STOP:
            buttons[1].click_input()
            print("Clicked 'Stop' button in toolbar. Waiting for confirmation dialog...")
            time.sleep(2)  # Increased wait for dialog to appear

            # Connect to the application and find the confirmation dialog
            app = Application(backend="uia").connect(title_re=".*", timeout=5)
            dialog = app.window(class_name="NemuMessageBox")
            max_attempts = 3
            attempt = 1

            while attempt <= max_attempts:
                if dialog.exists(timeout=5):
                    print("Confirmation dialog detected.")
                    # Find all buttons in the dialog
                    dialog_buttons = dialog.children(control_type="Button")
                    if dialog_buttons:
                        confirm_button = dialog_buttons[0]  # First button is "Confirm"
                        try:
                            confirm_button.click_input()
                            print("Clicked 'Confirm' button in dialog.")
                            time.sleep(1)  # Wait for dialog to close
                            return
                        except Exception as e:
                            print(f"Error clicking Confirm button: {str(e)}")
                    else:
                        print("No buttons found in dialog.")
                else:
                    print(f"Confirmation dialog not found on attempt {attempt}.")
                
                time.sleep(1)  # Wait before retrying
                attempt += 1

            print(f"Failed to find or click 'Confirm' button after {max_attempts} attempts.")
    except Exception as e:
        print(f"Error clicking toolbar button for {action.name}: {str(e)}")

def unselect_all_instances(main_window):
    select_all_checkbox = main_window.child_window(control_type="CheckBox", title="Select All")
    if select_all_checkbox.exists(timeout=2):
        select_all_checkbox.click_input()
        print("Clicked 'Select All' checkbox to unselect all instances.")
        time.sleep(1)
        
        select_all_checkbox.click_input()
        print("Clicked 'Select All' checkbox to select all instances.")
        # wait for the UI to update
        time.sleep(1)
    else:
        print("'Select All' checkbox not found.")

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
            send_keys(vm_name.upper(), with_spaces=True)
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
    # vm_dir = r"D:\Program Files\Netease\MuMuPlayerGlobal-12.0\vms"
    # D:\Program Files\Netease\MuMuPlayerGlobal-12.0\vms will extract this part from the
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

    for dir_name in os.listdir(vm_dir):
        dir_path = os.path.join(vm_dir, dir_name)
        if not os.path.isdir(dir_path) or not dir_name.startswith("MuMuPlayerGlobal-12.0-"):
            continue

        config_path = os.path.join(dir_path, "configs")
        config_file = os.path.join(config_path, "extra_config.json")
        if not os.path.exists(config_file):
            print(f"No config file found: {config_file}")
            continue

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

def scroll_instance_list(instance_list, direction="down", amount=2):
    """
    Scroll the instance list by sending Page Up/Down keys after focusing the list.
    Handles focus issues by attempting multiple focus methods and verifying focus.
    Args:
        instance_list: The list control to scroll
        direction: "up" or "down"
        amount: Number of Page Up/Down key presses (default 2 as per user observation)
    """
    print(f"Attempting to scroll instance list {direction}...")

    # Helper function to verify if the control is focused
    def is_control_focused(control):
        try:
            return control.has_focus() or control.is_active()
        except Exception:
            return False

    # Attempt to focus the instance list
    focused = False
    try:
        instance_list.set_focus()
        time.sleep(0.2)  # Small delay to ensure focus is set
        if is_control_focused(instance_list):
            focused = True
            print("Successfully focused instance list using set_focus()")
        else:
            print("set_focus() executed but control not focused")
    except Exception as e:
        print(f"set_focus() failed: {str(e)}")

    # Fallback: Click the list to focus it
    if not focused:
        print("Attempting to focus by clicking the list...")
        try:
            list_rect = instance_list.rectangle()
            if list_rect:
                center_x = list_rect.left + list_rect.width() // 2
                center_y = list_rect.top + list_rect.height() // 2
                

                click(coords=(center_x, center_y))
                time.sleep(0.2)
                if is_control_focused(instance_list):
                    focused = True
                    print(f"Successfully focused instance list by clicking at ({center_x}, {center_y})")
                else:
                    print("Click executed but control not focused")
            else:
                print("Could not get list rectangle for clicking")
        except Exception as e:
            print(f"Click to focus failed: {str(e)}")

    # Final fallback: Try focusing the parent window
    if not focused:
        print("Attempting to focus parent window...")
        try:
            parent = instance_list.parent()
            if parent and parent.is_enabled() and parent.is_visible():
                parent.set_focus()
                time.sleep(0.2)
                if is_control_focused(parent) or is_control_focused(instance_list):
                    focused = True
                    print("Successfully focused parent window")
                else:
                    print("Parent window focus executed but control not focused")
            else:
                print("Parent window not valid for focusing")
        except Exception as e:
            print(f"Parent window focus failed: {str(e)}")

    # Proceed with scrolling if focused, or attempt anyway with warning
    if not focused:
        print("Warning: Could not verify focus, attempting scroll anyway...")

    try:
        # Send Page Down or Page Up keys based on direction
        key = "{PGDN}" if direction == "down" else "{PGUP}"
        for _ in range(amount):
            send_keys(key)
            time.sleep(0.1)  # Small delay between key presses
        print(f"Successfully scrolled using {key} key ({direction}, {amount} steps)")
        time.sleep(0.5)  # Allow time for UI to update
        return True
    except Exception as e:
        print(f"Page {direction} key scroll method failed: {str(e)}")

    # Fallback: Try UIA Scroll Pattern
    print("Falling back to UIA Scroll Pattern...")
    try:
        element = instance_list.element_info
        scroll_pattern = element.get_current_pattern(10016)  # ScrollPattern ID
        if scroll_pattern:
            scroll_amount = 1 if direction == "down" else -1
            for _ in range(amount):
                scroll_pattern.scroll(0, scroll_amount)
                time.sleep(0.1)
            print(f"Successfully scrolled using UIA Scroll Pattern ({direction}, {amount} steps)")
            return True
        else:
            print("UIA Scroll Pattern not available")
    except Exception as e:
        print(f"UIA Scroll Pattern failed: {str(e)}")

    print("All scroll methods failed")
    return False

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
                    button.click_input()
            elif idx == 1:
                logical_name = "STOP (second button)"
                if action == SelectionActions.STOP:
                    button.click_input()
                    print("Clicked 'Stop' button. Waiting for confirmation dialog...")
                    time.sleep(6)  # Increased wait time

                    # Debug: List all elements to confirm dialog presence
                    list_elements_on_window(window)
                    print("Searching for dialog in current window hierarchy...")
                    dialogs = window.descendants(class_name="NemuMessageBox")
                    if dialogs:
                        dialog = dialogs[0]

                        # Recursively search for the first NemuPushButton7 and click it
                        def find_and_click_pushbutton7(control):
                            if control.element_info.class_name == "NemuUiLib::NemuPushButton7":
                                try:
                                    control.click_input()
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
                            dialog.set_focus()  # Ensure dialog is focused
                            dialog_buttons = dialog.children(control_type="Button")
                            if dialog_buttons:
                                confirm_button = dialog_buttons[0]  # First button is "Confirm"
                                if confirm_button.is_enabled() and confirm_button.is_visible():
                                    try:
                                        confirm_button.click_input()
                                        print("Clicked 'Confirm' button in dialog.")
                                        time.sleep(1)  # Wait for dialog to close
                                        return
                                    except Exception as e:
                                        print(f"Error clicking Confirm button: {str(e)}. Retrying...")
                                else:
                                    print("Confirm button not enabled or visible.")
                            else:
                                print("No buttons found in dialog.")
                        else:
                            print(f"Confirmation dialog not found on attempt {attempt}.")
                        
                        time.sleep(1)  # Wait before retrying
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
            continue  # Continue to the next button despite the error

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