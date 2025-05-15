import glob
import logging
import os
import time
from pathlib import Path
import ctypes

from pywinauto.application import Application
from pywinauto.controls.uia_controls import ButtonWrapper, TreeItemWrapper
from pywinauto.controls.common_controls import TreeViewWrapper
from pywinauto.base_wrapper import BaseWrapper


root_dir = os.getenv("LOCALAPPDATA")
assert root_dir, "local appdata"
exe_files = glob.glob(
    "Programs/Mixed In Key/Mixed In Key*/MixedInKey.exe",
    root_dir=root_dir,
    recursive=True,
)
assert len(exe_files) != 0, "mik .exe not found"
assert len(exe_files) == 1, "more than 1 mik .exe found"
exe_file = os.path.join(root_dir, exe_files[0])


def split_dir(dir: str):
    return [i for i in dir.split(os.sep) if i]


def get_screen_resolution() -> tuple[int, int]:
    user32 = ctypes.windll.user32
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


resolution = get_screen_resolution()


def scale_coords(coords: tuple[int, int], base_resolution=tuple[int, int]):
    """
    base_resolution: the resolution that we found the coords on
    """
    return int(coords[0] * resolution[0] / base_resolution[0]), int(
        coords[1] * resolution[1] / base_resolution[1]
    )


# Create window (or connect to existing)
def get_mik_window(create_new_if_not_open: bool):
    app = Application(backend="uia")

    for i in range(3):
        try:
            try:
                app.connect(title_re="Mixed In Key.*")
            except:  # pylint: disable=bare-except
                if create_new_if_not_open:
                    app.start(exe_file)

            window = app.top_window()
            window.wait("ready", timeout=20)
            window.wait("visible", timeout=20)
            return window, app
        except:  # pylint: disable=bare-except
            if i == 2:
                logging.error("Could not open MIK")
                raise
        time.sleep(3)
    raise Exception("Could not open MIK")


def run(dir: str):
    USER_DIR = os.path.expanduser("~")
    if not dir.startswith(USER_DIR):
        raise NotImplementedError(
            f"incorrect dir format, must start with user dir ({USER_DIR})"
        )

    window, app = get_mik_window(create_new_if_not_open=True)
    window_wrapper: BaseWrapper = window
    # window.print_control_identifiers()

    # Maximize window to make coordinate clicking consistent.
    try:
        maximize_btn: ButtonWrapper = window.child_window(auto_id="MaximizeButton")
        maximize_btn.click()
    except:
        pass  # already maximized

    window.child_window(control_type="ProgressBar").wait_not("visible", timeout=60 * 5)

    # Add Files button-- cannot find an identifier, have to just click coords
    window_wrapper.click_input(
        coords=scale_coords(coords=(111, 325), base_resolution=(1920, 1080))
    )
    # Add Folder button in dialog
    add_folder_btn: ButtonWrapper = window.child_window(best_match="ADD FOLDERButton")
    add_folder_btn.click()

    # Folder tree
    tree: TreeViewWrapper = window.Pane.TreeView

    def expand_nodes(originally_visible_nodes: list[str], remaining_nodes: list[str]):
        node: TreeItemWrapper = tree.get_item(originally_visible_nodes, False)
        node.ensure_visible()
        node.click_input()
        time.sleep(1)
        for node_str in remaining_nodes[:-1]:
            node = node.get_child(node_str)
            node.ensure_visible()
            node.click_input()
            time.sleep(1)
        node = node.get_child(remaining_nodes[-1])
        node.select()
        time.sleep(2)

    is_in_user_dir = False
    for user_dir in ["Music", "Downloads", "Documents"]:
        user_dir_long = f"{USER_DIR}\\{user_dir}"
        if user_dir_long in dir:
            is_in_user_dir = True
            expand_nodes(
                originally_visible_nodes=["Desktop", user_dir],
                remaining_nodes=split_dir(dir.split(user_dir_long, 1)[1]),
            )
            break
    if not is_in_user_dir:
        expand_nodes(
            originally_visible_nodes=["Desktop", "This PC"],
            remaining_nodes=split_dir(os.path.abspath(dir)),
        )

    # OK after selected folder
    folder_browser_popup = window.child_window(best_match="Browse For Folder")
    ok_btn: ButtonWrapper = folder_browser_popup.child_window(best_match="OK Button")
    ok_btn.click_input()

    # Wait for files to be added
    ADDING_FILES = "Adding files to queue"
    window.child_window(best_match=ADDING_FILES).wait("visible", timeout=60 * 5)
    window.child_window(best_match=ADDING_FILES).wait_not("visible", timeout=60 * 5)
    window.child_window(control_type="ProgressBar").wait_not("visible", timeout=60 * 5)

    # Attempt to close. If window closes, analysis done. If confirmation popup appears, is not done.
    while True:
        close_button: ButtonWrapper = window.child_window(auto_id="CloseButton")
        close_button.click()
        try:
            # Window item gets all weird here, sometimes can't find the NO button. Re-connect to the window.
            del app
            window, app = get_mik_window(create_new_if_not_open=False)
            no_button: ButtonWrapper = window.child_window(best_match="NOButton")
            no_button.wait("visible", timeout=5)
        except KeyboardInterrupt:
            input("press any key to continue checking mik")
            break
        except:
            break
        no_button.click()
        time.sleep(10)


if __name__ == "__main__":
    dir = os.path.join(Path.home(), "Music\\DJ Tracks")
    print("adding dir: ", dir)
    run(dir)
