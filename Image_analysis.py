import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

#Filtering so that pixels that are outside the range of mean +/- 10*std are set to 0
def filter_image(img):
    mean = mean_intensity(img)
    std = np.std(img)
    low_cutoff = mean - std*10
    high_cutoff = mean + std*10
    gray_image_filtered = np.copy(img)
    gray_image_filtered[(img < low_cutoff) | (img > high_cutoff)] = 0
    return gray_image_filtered

#Blurs image as a another way to decrease noise: by increasing kernel size you increase the amount of blurring
#Sigma is the variance within the kernel(Sigma of 1 keeps blurring localized)
def blur_image(img, kernel_size=(5,5), sigma=1):
    blur = cv2.GaussianBlur(img, kernel_size, sigma)
    return blur

#Figure out %Area of white pixels
def white_area_percentage(img):
    white_pixels = np.sum(img >= 125)
    total_pixels = img.size
    percentage = (white_pixels / total_pixels) * 100
    return percentage

#Median intensity of pixels
def median_intensity(img):
    median_val = np.median(img)
    return median_val

#Mean intensity of pixels
def mean_intensity(img):
    mean_val = np.mean(img)
    return mean_val

#Microscopy TIFFs may already be grayscale, or have an alpha channel
def to_grayscale(img):
    if img.ndim == 2:
        return img
    channels = img.shape[2]
    if channels == 1:
        return img[:, :, 0]
    if channels == 4:
        return cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
    if channels == 3:
        return cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    return cv2.cvtColor(img[:, :, :3], cv2.COLOR_BGR2GRAY)

#adaptiveThreshold only accepts 8-bit images. Stats stay on the original scale.
def to_uint8(img):
    if img.dtype == np.uint8:
        return img
    if np.issubdtype(img.dtype, np.integer):
        max_val = np.iinfo(img.dtype).max
        scaled = img.astype(np.float32) * (255.0 / max_val)
    else:
        arr = img.astype(np.float32)
        peak = float(np.max(arr)) if arr.size else 0.0
        if peak <= 1.0:
            scaled = arr * 255.0
        elif peak > 0:
            scaled = arr * (255.0 / peak)
        else:
            scaled = arr
    return np.clip(scaled, 0, 255).astype(np.uint8)

'''
Start of program execution
'''
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    #Folder that user selects containing the images to analyze
    folder_path = filedialog.askdirectory(title='Select an image folder')
    if not folder_path:
        print("No folder selected. Exiting.")
        return

    #Ask the user where they want the results saved, defaulting to a .csv
    #(CSV avoids the OneDrive save-lag issue that .xlsx files can run into)
    save_path = filedialog.asksaveasfilename(
        title="Choose where to save the results",
        defaultextension=".csv",
        filetypes=[("CSV file", "*.csv")],
        initialfile="Image_analysis.csv"
    )
    if not save_path:
        print("No save location selected. Exiting.")
        return

    #list used for holding results until for loop is finished
    results = []

    for file in os.listdir(folder_path):
        #Check if file is a .tif, .tiff
        ext = os.path.splitext(file)[1].lower()  # get file extension
        valid_exts = ['.tif', '.tiff']
        if ext not in valid_exts:
            print(f'Skipping "{file}"(not an image)')
            continue

        file_path = os.path.join(folder_path, file)
        img = cv2.imread(file_path, cv2.IMREAD_UNCHANGED)
        if img is None:
            print(f'Skipping "{file}" (could not read image)')
            continue

        #Sets image colors to B&W
        try:
            gray_image = to_grayscale(img)
        except cv2.error as exc:
            print(f'Skipping "{file}" (could not convert to grayscale: {exc})')
            continue
        gray_image_filtered = filter_image(gray_image)

        blur = blur_image(to_uint8(gray_image_filtered), (5,5), 0)

        #Guassian thresholding method
        thresh_guass = cv2.adaptiveThreshold(
            blur, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,  #You can also try 'cv2.ADAPTIVE_THRESH_MEAN_C' or 'cv2.THRESH_OTSU'
            cv2.THRESH_BINARY,
            725, 0
        )

        percentage = white_area_percentage(thresh_guass)
        median_val = median_intensity(gray_image_filtered)
        mean_val = mean_intensity(gray_image_filtered)

        results.append({
            "Image_Name": file,
            "White_Area_Percentage": round(percentage, 2),
            "Median_Intensity": round(median_val, 2),
            "Mean_Intensity": round(mean_val, 2),
            "Std_Dev": round(np.std(gray_image_filtered), 2)
        })

        #Prints results to console
        print(f'Image: {os.path.basename(file)}\nPercent coverage: {percentage:.2f}% \nMedian: {median_val:.2f} \nMean: {mean_val:.2f}\nStd: {np.std(gray_image_filtered):.2f}\n-------------------------')

    #Saves results to the file the user chose
    if results:
        df = pd.DataFrame(results)
        df.to_csv(save_path, index=False)
        print(f"Saved {len(results)} results to {save_path}")
    else:
        print("No results to save.")

if __name__ == "__main__":
    main()