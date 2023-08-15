import pandas as pd
import numpy as np
import datetime as dt
import glob
import xlsxwriter
import sys


## Make sure to remove Password Protection on "Invoice*" that is downloaded from Anthem Site.
## Password to Excel Invoice is 281301
## To remove password, open the sheet, enter the password, then go into the sheet and save it with a blank password

### If you run the scrip outside the month your in put a True and replace OVRD with MM-YYYY
monthOVRD = True
monthYearOVRD = '06-2023'
###

# import time
# start_time = time.time()

# pd.set_option("display.max_columns", 85)
# pd.set_option("display.max_rows", 500)

## used throughout script
na_vals = ['NA', 'Missing']
benefitCode = 'VIS01'
glEEcode = '2411'
glERcode = '7107'
glERCobra = '1408'
glCompany = '2003'
benefit = 'Anthem'



## used to get current Month+Year
if monthOVRD == False:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)
    strMonthYear = str(dtToday.month).rjust(2, "0") + "-" + str(dtToday.year)
else:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)
    strMonthYear = monthYearOVRD


## to be updated when benefit cost changes
BenefitCost = {
    'Benefit_Option' : [
        "Employee Only",
        "Employee + Child(ren)",
        "Employee + Spouse/DP",
        "Employee + Family",
    ],
    "Benefit_Cost": [5.92, 10.66, 10.08, 16.0],
    "EE_Cost": [2.96, 5.33, 5.04, 8.0],
    "ER_Cost": [2.96, 5.33, 5.04, 8.0],
}
dfBenefitCost = pd.DataFrame(BenefitCost)
dfBenefitCost.set_index(['Benefit_Option'], inplace=True)
dfBenefitCostRetro = dfBenefitCost.multiply(-1).copy()


## df (DataFrame/Table) import files
dfCoreBenefits = pd.DataFrame()
for file in glob.glob("_py_Active_Benefits_Census_KVG*.xlsx"):
    df = pd.read_excel(
        file, dtype={"Employee_Number": str, "SSN": str}, na_values=na_vals, engine="openpyxl"
    )
    if dfCoreBenefits.empty:
        dfCoreBenefits = df
    else:
        dfCoreBenefits = pd.merge(dfCoreBenefits, df, how="outer")

dfCoreBenefits = dfCoreBenefits.loc[
    dfCoreBenefits["Deduction/Benefit_Code"] == benefitCode
].copy()
dfCoreBenefits.loc[:, "SSN"] = dfCoreBenefits.loc[:, "SSN"].str.replace(" ", "", regex=True)
dfCoreBenefits['Employee_Number'] = dfCoreBenefits['Employee_Number'].str.strip()


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
    dfPayrollData.loc[:, "Deduction/Benefit_Code"] == benefitCode
].copy()
dfPayrollData.loc[:, "SSN"] = dfPayrollData.loc[:, "SSN"].str.replace(" ", "", regex=True)
dfPayrollData['Employee_Number'] = dfPayrollData['Employee_Number'].str.strip()


dfBenefitBill = pd.DataFrame()
for file in glob.glob(f"{benefit}/Bill/Invoice*.xlsx"):
    df = pd.read_excel(
        file, sheet_name='Membership Details', header=1,
        dtype={
            "Member ID No.": str
        },
        na_values=na_vals, engine="openpyxl"
    )
    if dfBenefitBill.empty:
        dfBenefitBill = df
    else:
        dfBenefitBill = pd.merge(dfBenefitBill, df, how="outer")

dfBenefitBill.columns = dfBenefitBill.columns.str.replace(" ", "_", regex=True)

dfBenefitBill.drop(columns='_Employee_Number', inplace=True)
dfBenefitBill.dropna(axis='index', subset=['Product_Type'], inplace=True)
dfBenefitBill.rename(columns={'Member_ID_No.': 'Member_ID', '_Number_Covered':'Number_Covered', "_Premium_Amount": "Cost"}, inplace=True)
dfBenefitBill.drop(dfBenefitBill.loc[:, 'Group_No./_Suffix':'Volume'].columns, axis=1, inplace=True)
dfBenefitBill.loc[:, 'Cost'] = dfBenefitBill.loc[:, 'Cost'].str.replace('$','', regex=True)
dfBenefitBill.loc[:, 'Cost'] = dfBenefitBill.loc[:, 'Cost'].astype(float)
dfBenefitBill["Contract_Type"] = dfBenefitBill["Contract_Type"].map(
    {
        "S": "Employee Only",
        "S+DEP": "Employee + Child(ren)",
        "2P": "Employee + Spouse/DP",
        "FAM": "Employee + Family",
    }
)


dfBillAdj = pd.DataFrame()
for file in glob.glob(f"{benefit}/Bill/Invoice*.xlsx"):
    df = pd.read_excel(
        file, 
        sheet_name='Eligibility Changes', header=3,
        dtype={
            "Member ID No.": str
        },
        na_values=na_vals, engine="openpyxl"
    )
    if dfBillAdj.empty:
        dfBillAdj = df
    else:
        dfBillAdj = pd.merge(dfBillAdj, df, how="outer")
dfBillAdj.columns = dfBillAdj.columns.str.replace(" ", "_", regex=True)
dfBillAdj.dropna(axis='index', subset=['Product_Type'], inplace=True)
dfBillAdj.rename(columns={'Member_ID_Number': 'Member_ID', "Prem._Adj": "Cost"}, inplace=True)
dfBillAdj.loc[:, 'Cost'] = dfBillAdj.loc[:, 'Cost'].str.replace('$','', regex=True)
dfBillAdj.loc[:, 'Cost'] = dfBillAdj.loc[:, 'Cost'].astype(float)
dfBillAdj = dfBillAdj.loc[dfBillAdj.loc[:, 'Cost'] != 0]
dfBillAdj = dfBillAdj[
    [
        'Member_ID',
        'Subscriber_Name',
        'Cost',
        'Reason_Code'
]].copy()


dfMergeBillPayroll = pd.concat([dfBenefitBill, dfBillAdj])
dfMergeBillPayroll['Subscriber_Name'] = dfMergeBillPayroll['Subscriber_Name'].str.strip()


dfMergeBillPayroll["GL_Employee"] = glEEcode
def Gl_ER(cobraDate):
    if pd.notnull(cobraDate):
        return glERCobra
    else:
        return glERcode
dfMergeBillPayroll["GL_Employer"] = dfMergeBillPayroll.apply(
    lambda x: Gl_ER(x["COBRA_End_Date"]), axis=1
)


# Bring in Manual Subscriber List
dfSubList = pd.DataFrame()
for file in glob.glob(f"{benefit}/_Subscriber_Name.xlsx"):
    df = pd.read_excel(
        file,
        dtype={
            "Employee_Number": str
        },
        na_values=na_vals, engine="openpyxl"
    )
    if dfSubList.empty:
        dfSubList = df
    else:
        dfSubList = pd.merge(dfSubList, df, how="outer")
dfSubList['Subscriber_Name'] = dfSubList['Subscriber_Name'].str.strip()


## test if Employee Number is missing
dfSubTest = pd.merge(
    dfMergeBillPayroll, dfSubList, left_on='Subscriber_Name', right_on='Subscriber_Name', how='left'
)

is_NaN = pd.DataFrame()
is_NaN = dfSubTest.loc[dfSubTest.loc[:, 'Employee_Number'].isnull()]
is_NaN = is_NaN.drop_duplicates(subset = ['Subscriber_Name'])
is_NaN = is_NaN[['Employee_Number','Subscriber_Name']]

if is_NaN.empty:
    print('Pass - No NaN')
else:
    is_NaN.to_excel(f'{benefit}/_{benefit}_NaN_list_DELETE_WHEN_DONE.xlsx', engine="xlsxwriter")
    sys.exit('Fail - Has NaN')


## Merge data
dfMergeBillPayroll = dfSubTest.copy()

dfBenefitBillActive = dfMergeBillPayroll.loc[dfMergeBillPayroll.loc[:, 'Reason_Code'].isnull()].copy()
dfBenefitBillActive['Activity'] = 'Activity'
dfBenefitBillReActive = dfMergeBillPayroll.loc[dfMergeBillPayroll.loc[:, 'Reason_Code'].notnull()].copy()
dfBenefitBillReActive['Activity'] = 'Retroactivity'

dfMerge1 = pd.merge(
    dfBenefitBillActive, dfPayrollData, left_on="Employee_Number", right_on="Employee_Number", how="left"
)
dfMergeBillPayroll = pd.merge(
    dfMerge1, dfCoreBenefits, left_on="Employee_Number", right_on="Employee_Number", how="left"
)
dfMergeBillPayroll = pd.concat([dfMergeBillPayroll, dfBenefitBillReActive])
dfMergeBillPayroll = pd.merge(
    dfMergeBillPayroll, dfEmpBasicInfo, left_on="Employee_Number", right_on="Employee_Number", how="left"
)

dfCoreBenefitsAudit = pd.merge(
    dfCoreBenefits, dfBenefitBillActive, left_on="Employee_Number", right_on="Employee_Number", how="left"
)

dfCoreBenefitsAudit = pd.merge(
    dfCoreBenefitsAudit, dfEmpBasicInfo, left_on="Employee_Number", right_on="Employee_Number", how="left"
)

## Functions
def EE_Cost(bill):
    SCost = dfBenefitCost["Benefit_Cost"].values.tolist()
    PSCost = dfBenefitCostRetro["Benefit_Cost"].values.tolist()
    if bill in SCost:
        x = dfBenefitCost.loc[dfBenefitCost["Benefit_Cost"] == bill, "EE_Cost"].to_list()
        x = x[0]
        return x
    elif bill in PSCost:
        y = dfBenefitCostRetro.loc[
            dfBenefitCostRetro["Benefit_Cost"] == bill, "EE_Cost"
        ].to_list()
        y = y[0]
        return y
    else:
        return 0


dfMergeBillPayroll["EE_Cost"] = dfMergeBillPayroll.apply(
    lambda x: EE_Cost(x["Cost"]), axis=1
)



def ER_Cost(bill):
    SCost = dfBenefitCost["Benefit_Cost"].values.tolist()
    PSCost = dfBenefitCostRetro["Benefit_Cost"].values.tolist()
    if bill in SCost:
        x = dfBenefitCost.loc[dfBenefitCost["Benefit_Cost"] == bill, "ER_Cost"].to_list()
        x = x[0]
        return x
    elif bill in PSCost:
        y = dfBenefitCostRetro.loc[
            dfBenefitCostRetro["Benefit_Cost"] == bill, "ER_Cost"
        ].to_list()
        y = y[0]
        return y
    else:
        return bill


dfMergeBillPayroll["ER_Cost"] = dfMergeBillPayroll.apply(
    lambda x: ER_Cost(x["Cost"]), axis=1
)



def Audit(cobra, activity, bill, payroll, empStatus, termDate):
    termStr = str(termDate.month).rjust(2, "0") + "-" + str(termDate.year)
    if activity == "Retroactivity":
        return "Retro - Okay"
    elif pd.notnull(cobra):
        return f"COBRA - Emp Termed in {termStr}"
    elif empStatus == "Terminated":
        return f"Termed in {termStr}"
    else:
        if bill == payroll:
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


dfMergeBillPayroll["Audit"] = dfMergeBillPayroll.apply(
    lambda x: Audit(
        x["COBRA_End_Date"],
        x["Activity"],
        x["Cost"],
        x["Payroll_Total"],
        x["Employment_Status"],
        x["Termination_Date"],
    ),
    axis=1,
)


def CoreAudit(BillOption, CoreOption):
    if BillOption == CoreOption:
        return "Good"
    else:
        return "Issue - Bill Option does not match Core"


dfCoreBenefitsAudit["Audit"] = dfCoreBenefitsAudit.apply(
    lambda x: CoreAudit(x["Benefit_Option"], x["Contract_Type"]), axis=1
)


def Comments(audit):
    if "Issue" in audit:
        return "Replace this text with Comments on issue"
    else:
        return ""


dfMergeBillPayroll["Comments"] = dfMergeBillPayroll.apply(
    lambda x: Comments(x["Audit"]), axis=1
)
dfCoreBenefitsAudit["Comments"] = dfCoreBenefitsAudit.apply(
    lambda x: Comments(x["Audit"]), axis=1
)


def CreditDebit(cost):
    if cost > 0:
        return "D"
    else:
        return "C"


# Checks and Balance
total_Cost = dfMergeBillPayroll["Cost"].sum()
total_Cost = round(total_Cost, 2)
total_Cost = "{:.2f}".format(total_Cost)
total_EE_ER_Cost = (
    dfMergeBillPayroll["ER_Cost"].sum() + dfMergeBillPayroll["EE_Cost"].sum()
)
total_EE_ER_Cost = round(total_EE_ER_Cost, 2)
total_EE_ER_Cost = "{:.2f}".format(total_EE_ER_Cost)


if total_Cost == total_EE_ER_Cost:
    P = f"In Balance, Bill = ${total_Cost}"
    print(P)
else:
    x = total_Cost - total_EE_ER_Cost
    y = f"Out of balance by ${x}"
    print(y)


## final tables
dfBill = dfMergeBillPayroll[
    [
        "Employee_Number",
        "Employee_Name",
        "Contract_Type",
        "Deduction/Benefit_Code_x",
        "Activity",
        "Org_Level_1_Code",
        "Org_Level_2_Code",
        "GL_Employee",
        "GL_Employer",
        "EE_Cost",
        "ER_Cost",
    ]
].copy()


dfBill.loc[:, "Deduction/Benefit_Code_x"] = dfBill.loc[
    :, "Deduction/Benefit_Code_x"
].fillna(benefitCode)
dfBill = dfBill.fillna(0)

dfBillEE = dfBill[
    [
        "Org_Level_1_Code",
        "Org_Level_2_Code",
        "GL_Employee",
        "EE_Cost",
    ]
].copy()
dfBillEE["Org_Level_1_Code"] = "01"
dfBillEE["Org_Level_2_Code"] = "01"
dfBillEE["Account"] = (
    "01-"
    + dfBillEE["GL_Employee"]
    + "-"
    + dfBillEE["Org_Level_1_Code"]
    + "-"
    + dfBillEE["Org_Level_2_Code"]
)
dfBillEE.rename(columns={"EE_Cost": "Amount"}, inplace=True)
dfBillEE["Type"] = dfBillEE.apply(lambda x: CreditDebit(x["Amount"]), axis=1)
dfBillEE.loc[:, ["Amount"]] = dfBillEE.loc[:, ["Amount"]].abs()
dfBillEE.drop(
    dfBillEE.loc[:, "Org_Level_1_Code":"GL_Employee"].columns, axis=1, inplace=True
)


dfBillER = dfBill[
    [
        "Org_Level_1_Code",
        "Org_Level_2_Code",
        "GL_Employer",
        "ER_Cost",
    ]
].copy()
dfBillER["Account"] = (
    "01-"
    + dfBillER["GL_Employer"]
    + "-"
    + dfBillER["Org_Level_1_Code"]
    + "-"
    + dfBillER["Org_Level_2_Code"]
)
dfBillER.rename(columns={"ER_Cost": "Amount"}, inplace=True)
dfBillER["Type"] = dfBillER.apply(lambda x: CreditDebit(x["Amount"]), axis=1)
dfBillER.loc[:, ["Amount"]] = dfBillER.loc[:, ["Amount"]].abs()
dfBillER.drop(
    dfBillER.loc[:, "Org_Level_1_Code":"GL_Employer"].columns, axis=1, inplace=True
)


dfGL = pd.concat([dfBillEE, dfBillER]).copy()
dfGL = dfGL.groupby(["Account", "Type"], as_index=False).sum("Cost")
totalGL = {"Account": f"01-{glCompany}-01-01", "Type": "C", "Amount": total_Cost}
dfGL = dfGL.append(totalGL, ignore_index=True)
dfGL["Reference"] = f"{benefit} {strMonthYear}"
dfGL["Journal"] = "JE"
dfGL["Date"] = strToday
dfGL = dfGL[
    [
        "Account",
        "Date",
        "Type",
        "Amount",
        "Journal",
        "Reference",
    ]
].copy()
dfGL.set_index("Account", inplace=True)
dfGL["Amount"] = dfGL["Amount"].astype(float).round(2)

totalBill = {
    "Audit": f"Cost(Bill) Total = ${total_Cost}",
    "Comments": f"GL Total = ${total_EE_ER_Cost}",
}
dfMergeBillPayroll = dfMergeBillPayroll.append(totalBill, ignore_index=True)

dfGlDetail = dfMergeBillPayroll[
    [
        "Employee_Number",
        "SSN",
        "Subscriber_Name",
        "Contract_Type",
        "Cost",
        "Activity",
        "Benefit_Option",
        "Employee_Amount",
        "Employer_Amount",
        "Payroll_Total",
        "Employment_Status",
        "Employee_Type",
        "Salary_or_Hourly",
        "Full/Part_Time",
        "Deduction/Benefit_Group",
        "Scheduled_Work_Hours",
        "Job_Code",
        "Job_Title",
        "Pay_Group",
        "Org_Level_1_Code",
        "Org_Level_1",
        "Org_Level_2_Code",
        "Org_Level_2",
        "Last_Hire_Date",
        "Termination_Date",
        "EE_Cost",
        "ER_Cost",
        "Audit",
        "Comments",
    ]
].copy()


dfCoreBenefitsAudit = dfCoreBenefitsAudit[
    [
        "Employee_Number",
        "Employee_Name",
        "Deduction/Benefit_Long",
        "Benefit_Option",
        "Contract_Type",
        "Audit",
        "Comments",
    ]
].copy()

dfCoreBenefitsAudit.sort_values(
    by=["Audit"], ascending=False, na_position="first", inplace=True
)

## Make csv files
dfGL.to_csv(f"{benefit}/{benefit}_GL_{strMonthYear}.csv")


## Make xlsx files
Invoice = pd.DataFrame()
Confirmation = pd.DataFrame()

writer = pd.ExcelWriter(
    f"{benefit}/{benefit}_Bill_Audit_{strMonthYear}.xlsx", engine="xlsxwriter"
)

Invoice.to_excel(writer, sheet_name="Invoice")
Confirmation.to_excel(writer, sheet_name="Confirmation")
dfGlDetail.to_excel(writer, sheet_name="Bill_Audit")
dfCoreBenefitsAudit.to_excel(writer, sheet_name="Core_Audit")
dfGL.to_excel(writer, sheet_name=f"GL_{strMonthYear}")

worksheet1 = writer.sheets["Invoice"]
worksheet2 = writer.sheets["Confirmation"]
worksheet3 = writer.sheets["Bill_Audit"]
worksheet4 = writer.sheets["Core_Audit"]
worksheet5 = writer.sheets[f"GL_{strMonthYear}"]

worksheet1.set_tab_color("#0bd422")  # dark green
worksheet2.set_tab_color("#0dff00")  # light green
worksheet3.set_tab_color("#00fbff")  # light blue
worksheet4.set_tab_color("#ffbf00")  # light orange
worksheet5.set_tab_color("yellow")

writer.save()


writerInvoice = pd.ExcelWriter(
    f"{benefit}/{benefit}_Bill_Invoice_GL_{strMonthYear}.xlsx", engine="xlsxwriter"
)

Invoice.to_excel(writerInvoice, sheet_name="Invoice")
Confirmation.to_excel(writerInvoice, sheet_name="Confirmation")
dfGL.to_excel(writerInvoice, sheet_name=f"GL_{strMonthYear}")

worksheetA = writerInvoice.sheets["Invoice"]
worksheetB = writerInvoice.sheets["Confirmation"]
worksheetC = writerInvoice.sheets[f"GL_{strMonthYear}"]

worksheetA.set_tab_color("#0bd422")  # dark green
worksheetB.set_tab_color("#0dff00")  # light green
worksheetC.set_tab_color("yellow")

writerInvoice.save()




# list(dfMemberList.columns.values)
# dfMemberList.loc[dfMemberList['SSN'] == '625403557']



# list(dfSubList.values)
# print(dfSubList)

# dfMergeBillPayroll.to_excel('Z_TestAnthem.xlsx')
# dfCoreBenefitsAudit.to_excel('Z_TestCoreAnthem.xlsx')

dfBill

# dfMergeBillPayroll
# list(dfMergeBillPayroll['Subscriber_Name'].unique())
