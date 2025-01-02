import win32gui
import win32con
import time
import ctypes
import json
import os
import sys
import re
import keyboard
import win32process

WINDOW_NAMES_FILE = "window_names.json"
FROZEN_WINDOWS = set()
THREAD_ALL_ACCESS = 0x1F03FF


def get_window_titles():
    """Returns a list of titles and HWNDs of all visible open windows."""
    window_list = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                window_list.append((title, hwnd))
        return True

    win32gui.EnumWindows(callback, None)
    return window_list


def find_window_by_hwnd(hwnd):
    """Finds window data by its HWND."""
    if win32gui.IsWindow(hwnd):
        title = win32gui.GetWindowText(hwnd)
        return title, hwnd
    return None, None


def move_window(hwnd, x, y, width=None, height=None):
    """Moves and resizes the window."""
    if not win32gui.IsWindow(hwnd):
        print("Invalid window handle.")
        return

    if width is None or height is None:
        # Get the current sizes if not specified
        rect = win32gui.GetWindowRect(hwnd)
        current_width = rect[2] - rect[0]
        current_height = rect[3] - rect[1]
        width = current_width
        height = current_height

    # Support DPI awareness
    dpi_awareness = ctypes.windll.user32.SetProcessDPIAware()
    win32gui.MoveWindow(hwnd, x, y, width, height, True)


def get_window_position(hwnd):
    """Gets the current window position and dimensions."""
    if not win32gui.IsWindow(hwnd):
        print("Invalid window handle.")
        return None
    rect = win32gui.GetWindowRect(hwnd)
    x = rect[0]
    y = rect[1]
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    return x, y, width, height


def load_window_names():
    """Loads window names from file, removes invalid handles."""
    window_names = {}
    if os.path.exists(WINDOW_NAMES_FILE):
        with open(WINDOW_NAMES_FILE, "r") as f:
            try:
                window_names = json.load(f)
            except json.JSONDecodeError:
                print("Error loading window names. File might be corrupted.")
                return {}

    # Remove handles of windows that no longer exist
    valid_window_names = {}
    for hwnd_str, name in window_names.items():
        try:
            hwnd = int(hwnd_str)
            if win32gui.IsWindow(hwnd):
                valid_window_names[hwnd_str] = name
        except ValueError:
            print(f"Invalid hwnd {hwnd_str}")

    if len(valid_window_names) < len(window_names):
        save_window_names(valid_window_names)
        print("Removed invalid window handles from window names")
    return valid_window_names


def save_window_names(window_names):
    """Saves window names to file."""
    with open(WINDOW_NAMES_FILE, "w") as f:
        json.dump(window_names, f, indent=4)


def display_window_list(window_names):
    """Displays available windows with their indices and custom names if available."""
    print("Available windows:")
    titles_and_hwnds = get_window_titles()
    for i, (title, hwnd) in enumerate(titles_and_hwnds):
        custom_name = window_names.get(str(hwnd), "N/A")
        print(f"{i + 1}. {title} (HWND: {hwnd}, Name: {custom_name})")
    return titles_and_hwnds


def handle_move_action(hwnd):
    """Handles the move window action."""
    try:
        x = int(input("Enter X coordinate: "))
        y = int(input("Enter Y coordinate: "))
        width = input("Enter width (leave empty to keep current): ")
        height = input("Enter height (leave empty to keep current): ")

        if width.strip() != '':
            width = int(width)
        else:
            width = None

        if height.strip() != '':
            height = int(height)
        else:
            height = None

        move_window(hwnd, x, y, width, height)
        print("Window moved.")
    except ValueError:
        print("Error: Invalid coordinates or sizes. Please provide valid numbers.")


def handle_get_position_action(hwnd):
    """Handles the get window position action."""
    position = get_window_position(hwnd)
    if position:
        x, y, width, height = position
        print(f"Current window position: X={x}, Y={y}, Width={width}, Height={height}")


def handle_rename_action(hwnd, window_names):
    """Handles the rename window action."""
    title, hwnd = find_window_by_hwnd(hwnd)
    if not hwnd:
        print("Invalid window handle.")
        return
    new_name = input("Enter a new name for the window: ")
    window_names[str(hwnd)] = new_name
    save_window_names(window_names)
    print(f"Window renamed to: {new_name}")


def handle_entr_kord_action(hwnd):
    """Handles the entr_kord action."""
    try:
        kord_str = input("Enter coordinates (e.g., X=-7, Y=0, Width=466, Height=359): ")
        match = re.match(r"X=(-?\d+), Y=(-?\d+), Width=(-?\d+), Height=(-?\d+)", kord_str)
        if match:
            x, y, width, height = map(int, match.groups())
            move_window(hwnd, x, y, width, height)
            print("Window moved to specified coordinates.")
        else:
            print("Error: Invalid format for coordinates.")
    except ValueError:
        print("Error: Could not parse coordinates, please provide valid numbers")


def freeze_window(hwnd):
    """Freezes the specified window by suspending its threads."""
    if hwnd in FROZEN_WINDOWS:
        print("Window is already frozen.")
        return

    h_thread = None  # Initialize h_thread
    try:
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        if thread_id == 0:
            print("Error getting window thread ID.")
            return

        h_thread = ctypes.windll.kernel32.OpenThread(win32con.THREAD_SUSPEND_RESUME, False, thread_id)
        if not h_thread:
            print("Error opening window thread.")
            return

        ctypes.windll.kernel32.SuspendThread(h_thread)
        FROZEN_WINDOWS.add(hwnd)
        print("Window frozen.")
    except Exception as e:
        print(f"Error freezing window: {e}")
    finally:
        if h_thread:
            ctypes.windll.kernel32.CloseHandle(h_thread)


def unfreeze_window(hwnd):
    """Unfreezes the specified window by resuming its threads."""
    if hwnd not in FROZEN_WINDOWS:
        print("Window is not frozen.")
        return

    h_thread = None  # Initialize h_thread
    try:
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        if thread_id == 0:
            print("Error getting window thread ID.")
            return

        h_thread = ctypes.windll.kernel32.OpenThread(THREAD_ALL_ACCESS, False, thread_id)
        if not h_thread:
            print("Error opening window thread.")
            return

        ctypes.windll.kernel32.ResumeThread(h_thread)
        FROZEN_WINDOWS.remove(hwnd)
        print("Window unfrozen.")
    except Exception as e:
        print(f"Error unfreezing window: {e}")
    finally:
        if h_thread:
            ctypes.windll.kernel32.CloseHandle(h_thread)


def unfreeze_all_windows():
    """Unfreezes all frozen windows."""
    for hwnd in list(FROZEN_WINDOWS):
        unfreeze_window(hwnd)


def handle_window_selection(window_names):
    """Handles the window selection process."""
    selected_hwnd = 0
    titles_and_hwnds = display_window_list(window_names)

    while True:
        try:
            choice = input(
                "Enter the number of the window to manage, a name to search for (or 'q' to quit, 'r' to restart): ")
            if choice.lower() == 'q':
                return 0, None, False  # return (hwnd, None, quit_flag)
            if choice.lower() == 'r':
                return 0, None, True  # return (hwnd, None, restart_flag)

            # Attempt to match by name in window_names
            matches = []
            for hwnd_str, name in window_names.items():
                if choice in name.lower():
                    try:
                        matches.append((name, int(hwnd_str)))
                    except ValueError:
                        pass

            if matches:
                if len(matches) == 1:
                    selected_title, selected_hwnd = matches[0]
                    if win32gui.IsWindow(selected_hwnd):
                        print(f"Selected window by custom name: {selected_title}")
                        return selected_hwnd, selected_title, False
                    else:
                        print("Window is no longer available.")

                else:
                    print("Multiple windows found, select by index")
                    for i, (title, hwnd) in enumerate(matches):
                        print(f"{i + 1}. {title} (HWND: {hwnd}, Name: {title})")

                    while True:
                        sub_choice = input("Select window to manage by index (or c to cancel window selection): ")
                        if sub_choice.lower() == 'c':
                            break

                        try:
                            sub_choice_index = int(sub_choice) - 1
                            if 0 <= sub_choice_index < len(matches):
                                selected_title, selected_hwnd = matches[sub_choice_index]
                                if win32gui.IsWindow(selected_hwnd):
                                    return selected_hwnd, selected_title, False
                                else:
                                    print("Window is no longer available.")
                                    break
                            else:
                                print("Invalid window choice.")
                        except ValueError:
                            print("Error: Please enter a valid number.")

            try:
                choice_int = int(choice) - 1
                if 0 <= choice_int < len(titles_and_hwnds):
                    selected_title, selected_hwnd = titles_and_hwnds[choice_int]
                    if win32gui.IsWindow(selected_hwnd):
                        return selected_hwnd, selected_title, False
                    else:
                        print("Window is no longer available.")
                else:
                    print("Invalid window choice.")
            except ValueError:
                print("Error: Please enter a valid number or name.")

        except ValueError:
            print("Error: Please enter a valid number or name.")


def main():
    """Main function to manage window actions."""
    window_names = load_window_names()
    selected_hwnd = 0
    selected_title = None
    restart = False

    keyboard.add_hotkey('f', unfreeze_all_windows)  # hotkey to unfreeze all windows

    while True:
        if restart:
            selected_hwnd = 0
            selected_title = None
            restart = False

        if selected_hwnd:
            title, hwnd = find_window_by_hwnd(selected_hwnd)

            if not hwnd or not win32gui.IsWindow(hwnd):
                print("\nSelected window is not available")
                selected_hwnd = 0
                selected_title = None
                continue

            print(f"\nCurrently selected window: {title}")
            custom_name = window_names.get(str(hwnd), "N/A")
            print(f"Current name: {custom_name}")

        selected_hwnd, selected_title, restart = handle_window_selection(window_names)

        if not selected_hwnd and not restart:
            break

        if restart:
            continue  # Skip action loop and restart selection

        if win32gui.IsWindow(selected_hwnd):
            while True:
                action = input(
                    "Choose an action (move/get_pos/rename/entr_kord/freeze/unfreeze/r to change window/q to quit): ").lower()
                if action == 'q':
                    break
                elif action == 'r':
                    restart = True
                    break  # restart
                elif action == 'move':
                    handle_move_action(selected_hwnd)
                elif action == 'get_pos':
                    handle_get_position_action(selected_hwnd)
                elif action == 'rename':
                    handle_rename_action(selected_hwnd, window_names)
                elif action == 'entr_kord':
                    handle_entr_kord_action(selected_hwnd)
                elif action == 'freeze':
                    freeze_window(selected_hwnd)
                elif action == 'unfreeze':
                    unfreeze_window(selected_hwnd)
                else:
                    print("Unknown action.")
        time.sleep(1)  # Pause to not overstress CPU


if __name__ == "__main__":
    main()