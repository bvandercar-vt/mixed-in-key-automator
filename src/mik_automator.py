import glob
import logging
import os
import time
from pathlib import Path

from pywinauto import application

root_dir = os.getenv("LOCALAPPDATA")
assert root_dir, "local appdata"
exe_files = glob.glob(
    f"Programs/Mixed In Key/Mixed In Key*/MixedInKey.exe",
    root_dir=root_dir,
    recursive=True,
)
assert len(exe_files) != 0, "mik .exe not found"
assert len(exe_files) == 1, "more than 1 mik .exe found"
exe_file = os.path.join(root_dir, exe_files[0])


# Create window (or connect to existing)
def get_mik_window(create_new_if_not_open: bool):
    app = application.Application(backend="uia")

    for i in range(3):
        try:
            try:
                app.connect(title_re="Mixed In Key.*")
            except:
                if create_new_if_not_open:
                    app.start(exe_file)

            window = app.top_window()
            window.wait("ready", timeout=20)
            window.wait("visible", timeout=20)
            return window, app
        except:
            if i == 2:
                logging.error("Could not open MIK")
                raise
        time.sleep(3)
    raise Exception("Could not open MIK")


def run(folder: str):
    USER_FOLDER = os.path.expanduser("~")
    if not folder.startswith(USER_FOLDER):
        raise NotImplementedError(
            f"incorrect folder format, must start with user folder ({USER_FOLDER})"
        )

    window, app = get_mik_window(create_new_if_not_open=True)

    # window.print_control_identifiers()

    # Maximize window to make coordinate clicking consistent.
    try:
        window.child_window(auto_id="MaximizeButton").click()
    except:
        pass  # already maximized

    window.child_window(control_type="ProgressBar").wait_not("visible", timeout=60 * 5)

    # Add Files button
    window.click_input(coords=(111, 325))
    # Add Folder button in dialog
    window.child_window(best_match="ADD FOLDERButton").click()

    # Folder tree
    node = window.Pane.TreeView.get_item(["Desktop", "This PC"], False)
    node.ensure_visible()
    node.click_input()
    time.sleep(1)
    subfolder = folder.split(f"{USER_FOLDER}\\", 1)[1]
    subfolders = subfolder.split("\\")
    for f in subfolders[:-1]:
        node = node.get_child(f)
        node.ensure_visible()
        node.click_input()
        time.sleep(1)
    node = node.get_child(subfolders[-1])
    node.select()
    time.sleep(2)

    # OK after selected folder
    folder_browser_popup = window.child_window(best_match="Browse For Folder")
    ok_button = folder_browser_popup.child_window(best_match="OK Button")
    ok_button.click_input()

    # Wait for files to be added
    window.child_window(best_match="Adding files to queue").wait(
        "visible", timeout=60 * 5
    )
    window.child_window(best_match="Adding files to queue").wait_not(
        "visible", timeout=60 * 5
    )
    window.child_window(control_type="ProgressBar").wait_not("visible", timeout=60 * 5)

    # Attempt to close. If window closes, analysis done. If confirmation popup appears, is not done.
    while True:
        window.child_window(auto_id="CloseButton").click()
        try:
            # Window item gets all weird here, sometimes can't find the NO button. Re-connect to the window.
            del app
            window, app = get_mik_window(create_new_if_not_open=False)
            no_button = window.child_window(best_match="NOButton").wait(
                "visible", timeout=5
            )
        except KeyboardInterrupt:
            input("press any key to continue checking mik")
            break
        except:
            break
        no_button.click()
        time.sleep(10)


if __name__ == "__main__":
    folder = os.path.join(Path.home(), "Music\\downloaded")
    print("downloading from: ", folder)
    run(folder)
