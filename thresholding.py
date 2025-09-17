import cv2
import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import pandas as pd
import os

#This is your excel file location
#If you are changing it paste it below to make sure the syntax is correct first
excel_file_location = "C:\\Users\\Ethan\\OneDrive - Westminster College\\Lignin_Biofilm_Project\\python_image_analysis.xlsx"

#Filtering so that pixels that are outside the range of mean +/- 10*std are set to 0
def filter_image(img):
    mean = mean_intensity(img)
    std = np.std(img)
    low_cutoff = mean - std*10
    high_cutoff = mean + std*10
    gray_image_filtered = np.copy(img)
    gray_image_filtered[(img < low_cutoff) | (img > high_cutoff)] = 0
    return gray_image_filtered

#Blurs image as a another way to decrease noise
def blur_image(img, kernel_size=(5,5), sigma=0):
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

'''
Start of program execution
'''
#Open file dialog to select a folder
root = tk.Tk()
root.withdraw()  # Hide the root window
#folder that user selects
folder_path = filedialog.askdirectory(title='Select an image folder')
if not folder_path:
    print("No file selected. Exiting.")
    exit()

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

    #Sets image colors to B&W
    gray_image = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray_image_filtered = filter_image(gray_image)

    blur = blur_image(gray_image_filtered, (5,5), 0)

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

    '''
    Uncomment to see the images
    '''
    #save_path = 'C:/Users/Ethan/Downloads/gray_image_filtered_1.png'
    #cv2.imwrite(save_path, gray_image_filtered)

    #Check you have an excel file location then adds the data
    if excel_file_location is not None:
        results.append({
          "Image_Name": file,
          "White_Area_Percentage": round(percentage, 2),
          "Median_Intensity": round(median_val, 2),
          "Mean_Intensity": round(mean_val, 2),
          "Std_Dev": round(np.std(gray_image_filtered), 2)
        })
    else:
        print("No excel file location specified. Skipping saving results.")
        break
    
    #Prints results to console
    print(f'Image: {os.path.basename(file)}\nPercent coverage: {percentage:.2f}% \nMedian: {median_val:.2f} \nMean: {mean_val:.2f}\nStd: {np.std(gray_image_filtered):.2f}\n-------------------------')

#Saves results to excel file
if results:
    df = pd.DataFrame(results)
#     df.to_excel(excel_file_location, sheet_name="Sheet1", index=False)
#     print(f"Saved {len(results)} results to {excel_file_location}")
else:
    print("No results to save.")