# Made by Jack M, contact jjrryyaa@student.ubc.ca if you got issues
# I know this is scuffed so feel free to improve on it
# You need to install pandas, xlrd, xlwt, XlsxWriter

import pandas as pd
import os

dir_path = os.path.dirname(os.path.realpath(__file__))
files = "/files/"
directory = dir_path + files
lab_name = ""
start = 0
end = 0


# Place all lab marks, along with the class list and prelab marks in a folder called "files" in the directory
# labnumber is the sheet number of the lab (first sheet is 0, L20), must be xlsx
# Must be formatted in the way where the TA marksheets start with lab (case sensitive), must be xlsx
# Prelabs must start with Lab, must be csv
# Class list must start with cpsc_121, must be csv
def mergeLabs(labnumber: int = 5) -> pd.DataFrame:
    first = 1
    global start
    global end
    cleanup = 0
    out = pd.DataFrame
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx") and filename.startswith("lab"):
            if first:
                first = 0
                out = pd.read_excel(directory + filename, sheet_name=labnumber)
                for idx, val in enumerate(out.columns.array):
                    if "Pre-lab".lower() in val.lower():
                        start = idx
                for idx, val in enumerate(out.columns.array):
                    if "Total".lower() in val.lower():
                        end = idx
                for idx, val in enumerate(out.columns.array):
                    if "Clean".lower() in val.lower():
                        cleanup = idx
            else:
                curr = pd.read_excel(directory + filename, sheet_name=labnumber)
                out = pd.concat([out, curr]).groupby(by=out.columns[0], sort=False).max().reset_index()
    out.loc[1:, out.columns[cleanup]] = 1
    return out


def attach_prelabs(input : pd.DataFrame) -> pd.DataFrame:
    for filename in os.listdir(directory):
        if filename.endswith(".csv") and filename.startswith("cpsc_121"):
            classList = pd.read_csv(directory + filename)
        if filename.endswith(".csv") and filename.startswith("Lab"):
            prelab = pd.read_csv(directory + filename)
    classList = classList.filter(items=["CSID","Secondary Section","Preferred Name","Surname"])
    classList = classList.loc[classList["Secondary Section"] == lab_name].reset_index(drop = True)
    temp = classList.loc[:,"CSID"].array
    prelab = prelab.loc[prelab["Name"].isin(temp)].filter(items = ["Name", "Total Score"])
    prelab.rename(columns={"Name":"CSID"},inplace = True)
    prelab = prelab.merge(classList,on = "CSID")
    prelab = prelab.filter(items = ["Preferred Name", "Surname", "Total Score"])
    input.rename(columns={input.columns[0] : "Preferred Name", input.columns[1] : "Surname"}, inplace = True)
    input = input.merge(prelab,on = [input.columns[0],input.columns[1]])
    input[input.columns[2]] = input[input.columns[-1]]
    input = input.drop(columns=input.columns[-1])
    input[input.columns[end]] = input.sum(axis=1)
    return input



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    frame = mergeLabs()
    lab_name = frame.columns[0]
    frame = attach_prelabs(frame)
    frame.to_excel(dir_path + "/out.xlsx")

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
