# Performance-Metrics-with-Excel
Calculates different performance metrics from NIFTI images in Excel and then returns the results back to Excel

# modules to import
import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
from sklearn import metrics
import SimpleITK as sitk
from sklearn.metrics import confusion_matrix
import openpyxl


# function that gives the metrics: recall, specificity, precision, accuracy, dice coefficient, jaccard respectively when given two nifti format images
# y_true is the ground truth image
# y_pred is the image that you want to compare with the ground truth image
def Full_Metrics(y_true, y_pred, true1, pred1):
# where y_pred has the pixel value of pred1, change it all to the pixel value true1
    y_pred[np.where(y_pred==pred1)] = true1
# tn = true negative, fp = false positive, fn = false negative, tp = true positive 
# confusion_matrix is a function of sklearn.matrix and search it up for how it specifically works
    tn, fp, fn, tp = confusion_matrix(y_true, y_pred, labels=[0,1]).ravel() 
# formulas for how to calculate each indicated metrics
    recall = tp/(tp+fn)
    specificity = tn/(tn+fp)
    precision = tp/(tp+fp)
    accuracy = (tp+tn)/(tp+tn+fp+fn)
    dice_coefficient = (2*precision*recall)/(precision+recall)
    jaccard = (precision*recall)/(precision+recall-(precision*recall))
    return recall, specificity, precision, accuracy, dice_coefficient, jaccard


# function that takes in the nifti images in an excel file and then calculates and returns the metrics
# can take in and calculate multiple images at once
# excel_file_initial is the excel file that holds the files that you want to calculate the metrics of
# pred_column is the column name that holds the prediction nifti images
# true_column is the column name that holds the ground truth nifti images
# val_true is the column name that holds the selected pixel value of the ground truth images
# val_pred is the column name that holds the pixel value of the prediction images that needs to be changed to the indicated val_true
def Excel_To_Metrics(excel_file_initial, pred_column, true_column, val_true, val_pred):
# reads the excel file 
    df = pd.read_excel(excel_file_initial)
# lists that take in the nifti files under the ground truth column and prediction column, respectively
    Truths = list(df[true_column])
    Predictions = list(df[pred_column])
# lists that hold the val_true and val_pred values, respectively
    True_Values_List = list(df[val_true])
    Pred_Values_List = list(df[val_pred])
# list that stores the metrics for the files
    Result = []
# loops through the number of pairs of files given and returns the metrics
    for i in range(len(Truths)):
# reads the images given in the excel file
        true = sitk.ReadImage(Truths[i])
        pred = sitk.ReadImage(Predictions[i])
# gets the array of the images
        true_array = sitk.GetArrayFromImage(true).flatten() 
        pred_array = sitk.GetArrayFromImage(pred).flatten()
# gets the pixel values that needs to be converted
        val_true = True_Values_List[i]
        val_pred = Pred_Values_List[i]
# uses the Full_Metrics function and puts all the calculated metrics in a list
        metrics_result = list(Full_Metrics(true_array, pred_array, val_true, val_pred))
# appends the list into the Result list and loops again until all pairs of images are looped 
        Result.append(metrics_result)
    return Result


# excel_file_initial is the excel file that holds the files that you want to calculate the metrics of and the file that the results calculated are put in
# excel_sheet is the excel sheet that you want to put the results in 
def Metrics_Back_To_Excel(excel_file_initial, excel_sheet, pred_column, true_column, val_true, val_pred):
    metrics_type = ["Recall", "Specificity", "Precision", "Accuracy", "Dice Coefficient", "Jaccard Index"]
# Excel_To_Metrics function
    calculations = Excel_To_Metrics(excel_file_initial, pred_column, true_column, val_true, val_pred)
    
    df = pd.read_excel(excel_file_initial)
    Truths = list(df[true_column])
    Predictions = list(df[pred_column])
    True_Values_List = list(df[val_true])
    Pred_Values_List = list(df[val_pred])
# opens the excel file that is inputted
    wb = openpyxl.load_workbook(excel_file_initial)
# opens the excel sheet that is inputted
    sheet = wb[excel_sheet]
# adds all the metric types from the metrics_type list to the excel file
    for i in range(0, len(metrics_type)):
        sheet.cell(row=1, column=i+5).value = metrics_type[i]
# adds all the metric results under the given metric type column
    for x in range(3, 9):
        for i in range(0, len(Truths)):
            sheet.cell(row=Truths.index(Truths[i])+i+2, column=x+2).value = calculations[i][x-3]
# saves the excel file
    wb.save(excel_file_initial)
