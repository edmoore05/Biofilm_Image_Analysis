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
        except:
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
    images.sort(key=lambda x: os.path.getctime(x))  # Makes sure ordered by the time the file was created
    max_num = int(len(images)/2)
    # Create naming sequence: T1..Tn + B1..Bn
    labels = [f"T{i}" for i in range(1, max_num+1)] + [f"B{i}" for i in range(1, max_num+1)]

    for img_path, label in zip(images, labels):
        folder, old_name = os.path.split(img_path)
        ext = os.path.splitext(old_name)[1]
        new_name = f"{prefix}_{label}{ext}"
        new_path = os.path.join(folder, new_name)

        os.rename(img_path, new_path)
        print(f"Renamed: {old_name} → {new_name}")

if __name__ == "__main__":
    main()