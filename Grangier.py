import glob
import os
import re

from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def find(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)


wb = Workbook()

ws = wb.active

path = 'C:\\Users\\tarun\\Downloads\\'

folderName = "runs"

fullPath = path + folderName

runs = os.listdir(fullPath)

row = 2

for run in runs:
    col = 1
    configFile = find("config.txt", fullPath + "\\" + run)
    detFile = find("det.txt", fullPath + "\\" + run)
    with open(configFile, "r") as file:
        text = file.read()
        regex = re.compile("strength=.+?(?=,)")
        result = regex.search(text)
        strength = text[result.span()[0]+9: result.span()[1]]

    with open(detFile, "r") as file:
        detections = {}
        lines = file.readlines()
        for line in lines:
            data = line.split(":")
            numDetections = data[1].strip()
            detections[data[0]] = numDetections

        ws[get_column_letter(col) + str(row)] = float(strength)
    for detector in detections:
        col += 1
        ws[get_column_letter(col) + str(row)] = int(detections[detector])

    row += 1

ws['A1'] = "Strength"
col = 1
row = 1
for detector in detections:
    col += 1
    ws[get_column_letter(col) + str(row)] = detector

wb.save(filename="test.xlsx")
wb.close()
