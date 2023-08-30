import pandas as pd
import numpy as np
import datetime as dt
import glob
import xlsxwriter

na_vals = ["NA", "Missing"]
benefitCodes = ['ADDL1', 'ADDL2', 'ADDl3', 'ADDl4', 'ADDLE', 'LIFLC', 'LIFLE', 'LIFLS', 'LIFL1', 'LIFL2', 'LIFL3', 'LIFL4', 'LTDL', 'ADDLS']
#glEEcode = ""
#glERCobra = ""
#glERcode = ""
#glCompany = ""
benefit = 'Lincoln'
monthOVRD = False
monthYearOVRD = '08-2023'
monthFile = '082023'

## Used to get current Month+Year
if monthOVRD == False:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)
    strMonthYear = str(dtToday.month).rjust(2, "0") + "-" + str(dtToday.year)
else:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)    
    strMonthYear = monthYearOVRD


dfEmpBasicInfo = pd.DataFrame()
for file in glob.glob("_py_Basic_Emp_Info_KVG*.xlsx"):
    df = pd.read_excel(
        file,
        dtype={
            "Employee_Number": str,
            "SSN": str,
            "Job_Code": str,
            "Org_Level_1_Code": str,
            "Org_Level_2_Code": str,
        },
        na_values=na_vals, engine="openpyxl"
    )
    if dfEmpBasicInfo.empty:
        dfEmpBasicInfo = df
    else:
        dfEmpBasicInfo = pd.merge(dfEmpBasicInfo, df, how="outer")

dfEmpBasicInfo.loc[:, "SSN"] = dfEmpBasicInfo.loc[:, "SSN"].str.replace(" ", "", regex=True)
dfEmpBasicInfo.loc[:, "Hourly_Pay_Rate"] = dfEmpBasicInfo.loc[
    :, "Hourly_Pay_Rate"
].round(2)
dfEmpBasicInfo['Employee_Number'] = dfEmpBasicInfo['Employee_Number'].str.strip()
dfEmpBasicInfo['Org_Level_1_Code'] = dfEmpBasicInfo['Org_Level_1_Code'].str.strip()
dfEmpBasicInfo['Org_Level_2_Code'] = dfEmpBasicInfo['Org_Level_2_Code'].str.strip()




dfPayrollData = pd.DataFrame()
for file in glob.glob("_py_Payroll_Benefits_Deductions_KVG*.xlsx"):
    df = pd.read_excel(
        file, dtype={"Employee_Number": str, "SSN": str}, na_values=na_vals, engine="openpyxl"
    )
    if dfPayrollData.empty:
        dfPayrollData = df
    else:
        dfPayrollData = pd.merge(dfPayrollData, df, how="outer")

dfPayrollData = dfPayrollData.loc[
    dfPayrollData.loc[:, "Deduction/Benefit_Code"].isin(benefitCodes)
].copy()
dfPayrollData.loc[:, "SSN"] = dfPayrollData.loc[:, "SSN"].str.replace(" ", "", regex=True)
dfPayrollData['Employee_Number'] = dfPayrollData['Employee_Number'].str.strip()





conditions = [
    (dfPayrollData['Deduction/Benefit_Code'] == 'ADDL1'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'ADDL2'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'ADDL3'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'ADDL4'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'ADDLE'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LIFLC'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LIFLE'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LIFLS'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LIFL1'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LIFL2'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LIFL3'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LIFL4'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'LTDL'),
    (dfPayrollData['Deduction/Benefit_Code'] == 'ADDLS')
]

values = ['AD+D', 'AD+D', 'AD+D', 'AD+D', 'V AD+D', 'VC LIFE', 'V LIFE', 'VS LIFE', 'LIFE', 'LIFE', 'LIFE', 'LIFE', 'LTD', 'VS AD+D']

dfPayrollData['Coverage'] = np.select(conditions, values)



#dfPayrollData



dfBenefitBill = pd.DataFrame()
#for file in glob.glob(f"{benefit}/Bill/*_LincolnBill.xlsx"):
#name of the file is odd, the month goes before "LincolnBill" making it hard to find, look into some reroutes 
for file in glob.glob(f"{benefit}/Bill/082023_LincolnBill.xlsx"):
    df = pd.read_excel(
        file,
        na_values=na_vals, engine="openpyxl"
    )
    if dfBenefitBill.empty:
        dfBenefitBill = df
    else:
        dfBenefitBill = pd.merge(dfBenefitBill, df, how="outer")



header = dfBenefitBill[dfBenefitBill['Unnamed: 0'] == 'Current Premium'].index[0] + 1
headerAdj = dfBenefitBill[dfBenefitBill['Unnamed: 0'] == 'Adjustments'].index[0] + 1
dfBenefitBill.columns = dfBenefitBill.iloc[header]
currentDF = dfBenefitBill[header+1:headerAdj+1]
adjDF = dfBenefitBill[headerAdj+1:]
dfBill = currentDF[currentDF['CERT NO.'].astype(str).str.isdigit()]
adjBill = adjDF[adjDF['CERT NO.'].astype(str).str.isdigit()]




dfBill = dfBill.rename(columns = {'CERT NO.':'SSN'})




EN = dfEmpBasicInfo[['Employee_Number', 'SSN']]




dfBill2 = pd.merge(EN, dfBill).dropna(axis=1,how='all')





pivotBill = pd.melt(dfBill2, 
                    id_vars=['Employee_Number', 'SSN', 'NAME'], 
                    value_vars=['LIFE', 'AD+D', 'LTD', 'V LIFE', 'V AD+D', 'VS LIFE', 'VS AD+D', 'VC LIFE'],
                    var_name='Coverage', 
                    value_name='Invoice_Sum',
                    ignore_index = False
                   ).sort_index().dropna()
#pivotBill







pivotBill.to_excel('Lincoln_Python_Data_8_2023.xlsx')







final = pd.merge(dfPayrollData, pivotBill, on = ['Employee_Number', 'Coverage', 'SSN'], how = 'outer')
final = final[['NAME', 'Employee_Number', 'SSN', 'Deduction/Benefit_Code', 'Coverage', 'Employee_Amount', 'Employer_Amount', 'Payroll_Total', 'Invoice_Sum']]





def audit(df):
    payroll = df['Payroll_Total']
    bill = df['Invoice_Sum']
    temp = round(abs(bill - payroll),2)
    if bill == payroll or temp == 0.01:
        return "Good"
    elif bill > payroll:
        x = bill - payroll
        x = round(x, 2)
        return f"Issue - Billed ${x} MORE than Payroll"
    elif bill < payroll:
        x = bill - payroll
        x = round(x, 2)
        return f"Issue - Bill ${x} LESS than Payroll"
    else:
        return "Issue - No Payroll Deduction"

final['Audit'] = final.apply(audit, axis = 1)






#final['Difference'] = abs(final['Payroll_Total'] - final['Invoice_Sum'])
#final.apply(lambda x: 'Good' if x['Payroll_Total'] == x['Invoice_Sum'] else ('Good' if x['Difference'] == 0.01 else  x['Payroll_Total'] - x['Invoice_Sum']))
#final


final.to_excel('Lincoln_Bill_Audit_8_2023.xlsx')