import os
import pandas
import shutil
from datetime import datetime
from collections import OrderedDict
from functools import reduce

__dir__ = os.path.abspath(os.path.dirname(__file__))
data_dir = os.path.join(__dir__, "data")
output_dir = os.path.join(__dir__, "output")
os.makedirs(output_dir, exist_ok=True)
all_files = os.scandir(data_dir)
files = [f for f in all_files if f.name.endswith(".xlsx")]

# Clean up messy file names
for f in files:
    path, name = os.path.split(f.path)
    name = name.replace("DLR ","").replace(" Khawar", "").replace(" ","").replace("-revised", "")
    if name[1] == "-":
        name = "0" + name
    if name != f.name:
        new_path = os.path.join(path, name)
        shutil.move(f.path, new_path)

# 
all_files = os.scandir(data_dir)
files = {}
for f in all_files:
    if f.name.endswith(".xlsx"):
        try:
            key = datetime.strptime(f.name.replace(".xlsx", ""), "%d-%m-%Y")
            files[key] = f.path
        except Exception:
            print("File that has wrong name: {}".format(f.name))

# Sort files based on month
files = OrderedDict(sorted(files.items()))
months = {month: [] for month in range(1, 13)}
for key, val in files.items():
    months[key.month].append(val)

cols = ["HYDEL."] + [str(num).zfill(2) + "00" for num in range(1, 25)]
index_col = None

for month, files in months.items():
    if len(files) == 0:
        continue
    ind = None
    wbs = []

    # Read each file and export interested columns to csv
    for f in files:
        if ind is None:
            ind = pandas.read_excel(f, sheet_name="MW", header=None, usecols="A", skiprows=3).squeeze(0)
            f_csv = f.replace(".xlsx", ".csv")
            ind.to_csv(f_csv, index=False)
            with open(f_csv) as fcsv:
                index_col = fcsv.readlines()
            index_col = [item.strip() for item in index_col]
            # index_col = list(ind)
        try:
            wb = pandas.read_excel(f, sheet_name="MW", header=None, usecols="B:Y", skiprows=3)
        except ValueError:
            print("Error reading file {}".format(f))
        # wbs.append(wb)
        f_csv = f.replace(".xlsx", ".csv")
        wb.to_csv(f_csv, index=False)
        with open(f_csv) as fcsv:
            wbs.append(fcsv.readlines())
        os.remove(f_csv)

    # Merge the output csv files
    csv_file = os.path.join(output_dir, "{}.csv".format(month))
    lines = []
    for idx in range(1, 307):
        line = index_col[idx]
        for wb in wbs:
            line += ",{}".format(wb[idx].strip())
        lines.append(line)
    with open(csv_file, "w") as f:
        f.write("\n".join(lines))
    merged_wb = pandas.read_csv(csv_file, header=None)
    os.remove(csv_file)
    merged_wb.to_excel(os.path.join(output_dir, "{}.xlsx".format(month)), index=False, header=None)
