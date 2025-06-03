# import pandas as pd
# path1=r"C:\Users\VAIO\Desktop\خسروی\مستندات\بوشهر\عملکرد 12ماهه 1403.xls"
# defective_state_name="ووجنوبي"
def find_sname_in_sheet(path,defective_state_name):
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

import pandas as pd
import openpyxl
xls_path=r"C:\Users\VAIO\Desktop\خسروی\مستندات\اصفهان\noName.xls"
path_2=r"C:\Users\VAIO\Desktop\خسروی\zir.xlsx"
path_3=r"C:\Users\VAIO\Desktop\خسروی\zir1.xlsx"
name_exist_in_sheet="اصفهان"
kk=find_sname_in_sheet(xls_path,name_exist_in_sheet)
kk1=str(kk).strip()
df = pd.read_excel(xls_path,sheet_name=kk,dtype=str,usecols=[8],skiprows=3,nrows=75)
df1 = pd.read_excel(path_2,usecols=[1],sheet_name=0,dtype=str)

vorudi=df.fillna("").values.flatten()
print(vorudi)

input("GIGILIIIII")
print (kk,"this is kk")
# print(df1[df1.iloc[:, 1] == kk])
# print(kk)
condition = df1.iloc[:, 0].str.strip() == kk1

if condition.any():
    row_index = condition.idxmax()
else:
    row_index = None
    print("No match found.")
print(row_index,"ROW_INDEXXXXX")
# ss=df1[df1.iloc[:, 1].str.strip() ==kk1]
datas=df.fillna("").values.flatten()

wb = openpyxl.load_workbook(path_2)
ws = wb.worksheets[0]

for i in range(75):
    ws.cell(row=row_index+2,column=3+i,value=datas[i])
wb.save(path_3)

# if ss.empty:
#     print("EMPTY")
# else:
#     print(ss)
# print(ss,"********************************************************************************")
# gigili=df1.iloc[:, 1].values


# for i in range(20):
#      for j in range(10):
#         print(df.iat[i,j])
#         if df.iat[i,j]=="برنامه/nعملکرد":
#             print(i,j,"*************************************************")



