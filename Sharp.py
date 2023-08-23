import pandas as pd
import numpy as np
import datetime as dt
import glob
import xlsxwriter

# import time
# start_time = time.time()

## Used throughout script
na_vals = ["NA", "Missing"]
benefitCode = "MED02"
glEEcode = "2404"
glERCobra = "1405"
glERcode = "7101"
glCompany = "2003"
benefit = 'Sharp'
monthOVRD = False
monthYearOVRD = '07-2023'

## Used to get current Month+Year
if monthOVRD == False:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)
    strMonthYear = str(dtToday.month).rjust(2, "0") + "-" + str(dtToday.year)
else:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)    
    strMonthYear = monthYearOVRD




## To be updated when benefit cost change
BenefitCost = {
    "Benefit_Option": [
        "Employee Only",
        "Employee + Child(ren)",
        "Employee + Spouse/DP",
        "Employee + Family",
    ],
    "Benefit_Cost": [794.64, 1430.35, 1748.21, 2463.39],
    "EE_Cost": [103.30, 354.41, 479.96, 762.46],
    "ER_Cost": [691.34, 1075.94, 1268.25, 1700.93],
}
dfBenefitCost = pd.DataFrame(BenefitCost)
dfBenefitCost.set_index(["Benefit_Option"], inplace=True)
dfBenefitCostRetro = dfBenefitCost.multiply(-1).copy()


## df (DataFrame/Table) import files
dfCoreBenefits = pd.DataFrame()
for file in glob.glob("_py_Active_Benefits_Census_KVG*.xlsx"):
    df = pd.read_excel(
        file, dtype={"Employee_Number": str, "SSN": str}, na_values=na_vals
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
        na_values=na_vals,
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
        file, dtype={"Employee_Number": str, "SSN": str}, na_values=na_vals
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
## note the below is only to override the payroll deductions by 1 cent because of rounding
dfPayrollData['Employer_Amount'] = dfPayrollData['Employer_Amount'].replace([593.84], 593.85)
dfPayrollData['Payroll_Total'] = dfPayrollData['Payroll_Total'].replace([733.12], 733.13)
dfPayrollData


dfBenefitBill = pd.DataFrame()
for file in glob.glob(f"{benefit}/Bill/THE ARC OF*.xlsx"):
    df = pd.read_excel(
        file,
        dtype={
            "Bill Number": str,
            "Account HCC ID": str,
            "Plan HCC ID": str,
            "Member HCC ID": str,
            "Member SSN": str,
            "Subscription HCC ID": str,
        },
        na_values=na_vals,
    )
    if dfBenefitBill.empty:
        dfBenefitBill = df
    else:
        dfBenefitBill = pd.merge(dfBenefitBill, df, how="outer")

dfBenefitBill.columns = dfBenefitBill.columns.str.replace(" ", "_", regex=True)
dfBenefitBill.rename(columns={"Member_SSN": "SSN", "Total_Premium": "Cost"}, inplace=True)
dfBenefitBill.loc[:, "SSN"] = dfBenefitBill.loc[:, "SSN"].str.replace("-", "", regex=True)
dfBenefitBill.dropna(axis="index", subset=["Bill_Number"], inplace=True)
dfBenefitBill.drop(
    dfBenefitBill.loc[:, "Bill_Number":"Account_HCC_ID_Detail"].columns,
    axis=1,
    inplace=True
)
dfBenefitBill.drop(
    dfBenefitBill.loc[:, "Plan_HCC_ID":"Member_DOB"].columns, axis=1, inplace=True
)
def EE_GL(company):
    if company == 'The Arc of San Diego':
        return glEEcode
    else:
        return glERCobra
dfBenefitBill["GL_Employee"] = dfBenefitBill.apply(
    lambda x: EE_GL(x['Account_Name_Detail']), axis=1
)

def ER_GL(company):
    if company == 'The Arc of San Diego':
        return glERcode
    else:
        return glERCobra
dfBenefitBill["GL_Employer"] = dfBenefitBill.apply(
    lambda x: ER_GL(x['Account_Name_Detail']), axis=1
)

dfBenefitBill["Contract_Type"] = dfBenefitBill["Contract_Type"].map(
    {
        "EE Only": "Employee Only",
        "EE Plus Child": "Employee + Child(ren)",
        "EE Plus Spouse": "Employee + Spouse/DP",
        "Family": "Employee + Family",
    }
)


## Merge data
dfBenefitBillActive = dfBenefitBill.loc[dfBenefitBill.loc[:, "Activity"] != "Retroactivity"]
dfBenefitBillReActive = dfBenefitBill.loc[dfBenefitBill.loc[:, "Activity"] == "Retroactivity"]

dfMerge1 = pd.merge(
    dfBenefitBillActive, dfPayrollData, left_on="SSN", right_on="SSN", how="left"
)
dfMergeBillPayroll = pd.merge(
    dfMerge1, dfCoreBenefits, left_on="SSN", right_on="SSN", how="left"
)
dfMergeBillPayroll = pd.concat([dfMergeBillPayroll, dfBenefitBillReActive])
dfMergeBillPayroll = pd.merge(
    dfMergeBillPayroll, dfEmpBasicInfo, left_on="SSN", right_on="SSN", how="left"
)

dfCoreBenefitsAudit = pd.merge(
    dfCoreBenefits, dfBenefitBillActive, left_on="SSN", right_on="SSN", how="left"
)

dfCoreBenefitsAudit = pd.merge(
    dfCoreBenefitsAudit, dfEmpBasicInfo, left_on="SSN", right_on="SSN", how="left"
)


## Functions
def EE_Cost(company, bill):
    SCost = dfBenefitCost["Benefit_Cost"].values.tolist()
    PSCost = dfBenefitCostRetro["Benefit_Cost"].values.tolist()
    if company == 'The Arc of San Diego COBRA':
        return bill
    elif bill in SCost:
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
    lambda x: EE_Cost(x["Account_Name_Detail"], x["Cost"]), axis=1
)


def ER_Cost(company, bill):
    SCost = dfBenefitCost["Benefit_Cost"].values.tolist()
    PSCost = dfBenefitCostRetro["Benefit_Cost"].values.tolist()
    if company == 'The Arc of San Diego COBRA':
        return 0
    elif bill in SCost:
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
    lambda x: ER_Cost(x["Account_Name_Detail"], x["Cost"]), axis=1
)


def Audit(cobra, activity, bill, payroll, empStatus, termDate):
    termStr = str(termDate.month).rjust(2, "0") + "-" + str(termDate.year)
    if activity == "Retroactivity":
        return "Retro - Okay"
    elif "COBRA" in cobra:
        return f"COBRA - Emp Termed in {termStr}"
    elif empStatus == "Terminated":
        return f"Termed in {termStr}"
    else:
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


dfMergeBillPayroll["Audit"] = dfMergeBillPayroll.apply(
    lambda x: Audit(
        x["Account_Name_Detail"],
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


## Checks and Balance
total_Cost = dfMergeBillPayroll["Cost"].sum()
total_Cost = round(total_Cost, 2)
int_Total_Cost = round(total_Cost, 2)
total_Cost = "{:.2f}".format(total_Cost)
total_EE_ER_Cost = (
    dfMergeBillPayroll["ER_Cost"].sum() + dfMergeBillPayroll["EE_Cost"].sum()
)
total_EE_ER_Cost = round(total_EE_ER_Cost, 2)
int_Total_EE_ER_Cost = round(total_EE_ER_Cost, 2)
total_EE_ER_Cost = "{:.2f}".format(total_EE_ER_Cost)


if total_Cost == total_EE_ER_Cost:
    P = f"In Balance, Bill = ${total_Cost}"
    print(P)
else:
    y = int_Total_Cost - int_Total_EE_ER_Cost
    y = "{:.2f}".format(y)
    y = f"Out of balance by ${y}"
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
dfGL = dfGL[dfGL['Amount'] !=0].copy()

totalBill = {
    "Audit": f"Cost(Bill) Total = ${total_Cost}",
    "Comments": f"GL Total = ${total_EE_ER_Cost}",
}
dfMergeBillPayroll = dfMergeBillPayroll.append(totalBill, ignore_index=True)

dfGlDetail = dfMergeBillPayroll[
    [
        "Coverage_Month",
        "Employee_Number",
        "SSN",
        "Subscriber_Full_Name",
        "Account_Name_Detail",
        "Contract_Type",
        "Contract_Size",
        "Cost",
        "Activity",
        "Deduction/Benefit_Long",
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

dfCoreBenefitsAudit.drop(columns='Employee_Number_y', inplace=True)
dfCoreBenefitsAudit.rename(columns={"Employee_Number_x": "Employee_Number"}, inplace=True)
dfCoreBenefitsAudit = dfCoreBenefitsAudit[
    [
        "Employee_Number",
        "Employee_Name",
        "Deduction/Benefit_Long",
        "Benefit_Option",
        "Contract_Type",
        "Contract_Size",
        "Coverage_Month",
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



## For testing

# dfCoreBenefitsAudit.to_excel(f"Test_{strMonthYear}.xlsx", sheet_name="BIll")

# dfCoreBenefitsAudit
# list(dfCoreBenefitsAudit.columns.values)
# list(dfGL.columns.values)

dfBill

# end_time = time.time() - start_time
# print(end_time)

print(f'{benefit}.py ran successfully')