'''
This script will have you choose a folder then what the prefix will be. Then it will reneame all images in order of oldest to newest labeling them with the prefix then _T1 through _B5.
i.e.  if you input "C1" all images be renamed C1_T1, C1_T2, C1_T3, ... C1_B5
'''
import tkinter as tk
from tkinter import filedialog
import os
import json

CONFIG_FILE = "last_folder.json"  # Where we store the last folder path

def main():
    folder = choose_folder()
    if folder:
        images = get_images_from_folder(folder)
        prefix = input("Enter the prefix for renaming: ").strip()
        rename_images_with_pattern(folder, prefix)
    else:
        print("No folder selected")
        
def choose_folder():
    #Open folder picker dialog, starting from last folder if possible.
    root = tk.Tk()
    root.withdraw()  # Hide root window

    last_folder = load_last_folder()
    folder_path = filedialog.askdirectory(
        title="Select a Folder",
        initialdir=last_folder if os.path.exists(last_folder) else "/"
    )

    if folder_path:
        save_last_folder(folder_path)
    return folder_path

def save_last_folder(path):
    #Save last folder path to config file.
    with open(CONFIG_FILE, "w") as f:
        json.dump({"last_folder": path}, f)

def load_last_folder():
    #Load last folder path from config file.
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                return data.get("last_folder", "")
        except (json.JSONDecodeError, OSError):
            pass
    return ""

def get_images_from_folder(folder_path):
    #Return a list of image file paths from a folder.
    # Allowed image file extensions
    image_extensions = (".tif", ".tiff", ".jpg", ".jpeg", ".png")

    # Create a list of all files that match image extensions
    images = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(image_extensions)
    ]

    return images

def rename_images_with_pattern(folder_path, prefix):
    images = get_images_from_folder(folder_path)
    if not images:
        print("No images found in that folder.")
        return

    invalid = '<>:"/\\|?*'
    if any(char in prefix for char in invalid):
        print(f"Prefix cannot contain any of these characters: {invalid}")
        return

    images.sort(key=lambda x: os.path.getctime(x))  # Makes sure ordered by the time the file was created
    if len(images) % 2 != 0:
        leftover = os.path.basename(images[-1])
        print(f'Odd number of images ({len(images)}). "{leftover}" was not renamed.')

    max_num = int(len(images)/2)
    if max_num == 0:
        print("Need at least two images (one T and one B). Nothing was renamed.")
        return

    # Create naming sequence: T1..Tn + B1..Bn
    labels = [f"T{i}" for i in range(1, max_num+1)] + [f"B{i}" for i in range(1, max_num+1)]
    planned = []
    for img_path, label in zip(images, labels):
        folder, old_name = os.path.split(img_path)
        ext = os.path.splitext(old_name)[1]
        new_name = f"{prefix}_{label}{ext}"
        new_path = os.path.join(folder, new_name)
        planned.append((img_path, new_path, old_name, new_name))

    def same_path(left, right):
        return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))

    sources = {os.path.normcase(os.path.abspath(src)) for src, _, _, _ in planned}
    for src, dst, old_name, new_name in planned:
        if same_path(src, dst):
            continue
        if os.path.exists(dst) and os.path.normcase(os.path.abspath(dst)) not in sources:
            print(f"Cannot rename {old_name}: {new_name} already exists.")
            return

    # Move everything aside first so names inside the batch can swap without a collision
    staging = []
    finished = []
    try:
        for index, (src, dst, old_name, new_name) in enumerate(planned):
            if same_path(src, dst):
                print(f"Already named: {old_name}")
                continue
            temporary = os.path.join(os.path.dirname(src), f".rename_tmp_{index}_{old_name}")
            os.rename(src, temporary)
            staging.append((src, temporary, dst, old_name, new_name))

        for src, temporary, dst, old_name, new_name in staging:
            os.rename(temporary, dst)
            finished.append((temporary, dst))
            print(f"Renamed: {old_name} → {new_name}")
    except OSError as exc:
        print(f"Rename failed ({exc}). Restoring original file names.")
        for temporary, dst in reversed(finished):
            if os.path.exists(dst) and not os.path.exists(temporary):
                os.rename(dst, temporary)
        for src, temporary, _, _, _ in reversed(staging):
            if os.path.exists(temporary) and not os.path.exists(src):
                os.rename(temporary, src)

if __name__ == "__main__":
    main()