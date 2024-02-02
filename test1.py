import openpyxl as xl

# excel file used here gfg.xlsx 
excel_file = "C:\\Users\\indronil\\Downloads\\Projected FY24 Budget 05 Nov.xlsm"
  
# load the workbook 
wb = xl.load_workbook(excel_file, keep_vba=True) 
  
sheets = wb.sheetnames
accounts = []

#-----------------------------------------------------------------------------STEP_1---------------------------------------------------------------------------------
def copy_monthly_sheet_data():
    # Clear all data from the destination sheet
    destination_sheet = wb["FI-U227 Statement of Revenu (2"]
    destination_sheet.delete_rows(1, destination_sheet.max_row)

    # read monthly file
    monthly_file_name = "C:\\Users\\indronil\\Downloads\\FI-U227 OCT Statement of Revenue and Expense Detail.xlsx"
    monthly_file = xl.load_workbook(monthly_file_name) 
    source_sheet = monthly_file.worksheets[0]

    accounts_monthly = []
    # Copy data from the source sheet to the destination sheet
    for row in source_sheet.iter_rows(min_row=1, max_row=source_sheet.max_row, values_only=True):
        account_number = row[5][0:6]
        try:
            accounts_monthly.append(int(account_number))
        except:
            account_number = "ACCT"

        row = list(row)
        row.insert(5,account_number)
        destination_sheet.append(row)
    
    accounts_monthly = list(set(accounts_monthly))
    wb.save("modified.xlsm")
    return accounts_monthly

#-----------------------------------------------------------------------------STEP_2---------------------------------------------------------------------------------
def compare_account_numbers():
    existing_accounts = []
    for cell in wb["SUMMARY - FS (000000)"]['B']:
        try:
            account_number = int(cell.value)
            if(len(str(account_number))):
                existing_accounts.append(account_number)
        except:
            account_number=0

    new_account = []

    for acct in accounts:
        if acct not in existing_accounts:
            new_account.append(acct)
    
    if len(new_account) > 0:
        print("New accounts to add ", new_account)
    else:
        print("No new account to add!")
