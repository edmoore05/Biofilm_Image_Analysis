import cv2
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import pandas as pd
import os

#This is your excel file location
excel_file_location = "C:\\Users\\Ethan\\OneDrive - Westminster College\\Lignin_Biofilm_Project\\python_image_analysis.xlsx"

#Filtering so that pixels below 5 and above 84 are set to 0
low_cutoff = 5
hight_cutoff = 84
def filter_image(img, low_cutoff=5, high_cutoff=84):
    gray_image_filtered = np.copy(gray_image)
    gray_image_filtered[gray_image < low_cutoff] = 0
    gray_image_filtered[gray_image > hight_cutoff] = 0
    return gray_image_filtered

#Blur
def blur_image(img, kernel_size=(5,5), sigma=0):
    blur = cv2.GaussianBlur(img, (5,5), 0)
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

#Append results to excel sheet
def append_results(path, row_dict, sheet_name="Sheet1"):
    new_df = pd.DataFrame([row_dict])

    try:
        # Load existing workbook
        book = load_workbook(path)

        # Find the last row in the existing sheet
        startrow = book[sheet_name].max_row

        # Append new row without header
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            new_df.to_excel(writer, sheet_name=sheet_name,
                            startrow=startrow, index=False, header=False)
        file_name = os.path.basename(file)
        print(f"Appended {file_name} to {path}")

    except PermissionError:
        print(f"PermissionError: {path} is locked (Excel or OneDrive may have it open). Close it and try again.")


'''
Start of program execution
'''
#Open file dialog to select image
root = tk.Tk()
root.withdraw()  # Hide the root window
#img that user selects
file = filedialog.askopenfilename(title='Select an image file')
if not file:
    print("No file selected. Exiting.")
    exit()
img = cv2.imread(f'{file}', cv2.IMREAD_UNCHANGED)

#Sets image colors to B&W
gray_image = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
#Filters the image so that pixels below 5 and above 84 are set to 0
gray_image_filtered = filter_image(gray_image, low_cutoff, hight_cutoff)

#Blurs image so that thresholding is cleaner
blur = blur_image(gray_image_filtered, (5,5), 0)
#Guassian thresholding method
thresh_guass = cv2.adaptiveThreshold(
    blur, 255,
    cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
    cv2.THRESH_BINARY,
    725, 0
)

#Calculates the values of %Area, Median, and Mean
percentage = white_area_percentage(thresh_guass)
median_val = median_intensity(gray_image_filtered)
mean_val = mean_intensity(gray_image_filtered)
print(f'Image: {os.path.basename(file)}\npercent coverage: {percentage:.2f}% \nMedian: {median_val} \nMean: {mean_val}')

# #Check you have an excel file location then adds the data
# if excel_file_location is not None:
#     new_row = {'Image_Name': os.path.basename(file), 
#                'White_Area_Percentage': percentage, 
#                'Median_Intensity': median_val, 
#                'Mean_Intensity': mean_val
#     }
#     append_results(excel_file_location, new_row, sheet_name='Sheet1')

