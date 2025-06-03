import pandas as pd
import openpyxl
import glob
import os
os.environ["TF_CPP_MIN_LOG_LEVEL"]="0"
# df = pd.read_excel(file_path, engine='openpyxl',usecols=list(range(0,7)),sheet_name="Sheet1",dtype=column_dtype)

def find_state_name_in_sheet(path,defective_state_name):
    state_name_in_sheet=""
# Create an ExcelFile object first
    excel_file = pd.ExcelFile(path)
    
    # Get all sheet names
    sheet_names = excel_file.sheet_names

    for i in sheet_names:
        if defective_state_name in i:
            state_name_in_sheet=i
            break
    return (state_name_in_sheet)


current_dir=os.path.dirname(os.path.abspath(__file__))
sub_dir=current_dir+"\\مستندات"
kk=0
for i in glob.glob(sub_dir+"\\*"):
    sho=0
    state_name=(i.split("\\"))[-1]
    state_path_xls=sub_dir+"\\"+state_name+"\\*.xls"
    state_path_xlsx=sub_dir+"\\"+state_name+"\\*.xlsx"

    for xls_path in glob.glob(state_path_xls)+glob.glob(state_path_xlsx):
        sho+=1
        kk+=1
        if sho==1:
            name_exist_in_sheet=find_state_name_in_sheet(xls_path,state_name)
            if name_exist_in_sheet!="":
                print(name_exist_in_sheet,"fffffffffffffffff")
                input()
                # df = pd.read_excel(xls_path, engine='openpyxl',usecols=list(range(0,10)),sheet_name=name_exist_in_sheet,dtype=column_dtype)
                print(name_exist_in_sheet,"dadasdsa")
                print(xls_path)
                df = pd.read_excel(xls_path,usecols=list(range(0,10)),sheet_name=name_exist_in_sheet)
                print (df.iloc[0:20,])
                input()
            break
    


