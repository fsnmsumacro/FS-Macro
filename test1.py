import openpyxl as xl
import bisect

# excel file used here
excel_file = ""
monthly_file_name = ""
wb = None
accounts = []
new_acc = []

#-----------------------------------------------------------------------------STEP_1---------------------------------------------------------------------------------
def copy_monthly_sheet_data():
    # load the workbook
    global wb
    wb = xl.load_workbook(excel_file, keep_vba=True) 
    sheets = wb.sheetnames

    # Clear all data from the destination sheet
    destination_sheet = wb["FI-U227 Statement of Revenu (2"]
    destination_sheet.delete_rows(1, destination_sheet.max_row)

    # read monthly file
    #monthly_file_name = "C:\\Users\\indronil\\Downloads\\FI-U227 OCT Statement of Revenue and Expense Detail.xlsx"
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
    wb.save("New_modified.xlsm")
    return accounts_monthly

#-----------------------------------------------------------------------------STEP_2---------------------------------------------------------------------------------
def compare_account_numbers():
    global new_acc
    existing_accounts = []
    for cell in wb["SUMMARY - FS (000000)"]['B']:
        try:
            account_number = int(cell.value)
            if(len(str(account_number))):
                existing_accounts.append(account_number)
        except:
            account_number=0

    for acct in accounts:
        if acct not in existing_accounts:
            new_acc.append(acct)
    
    if len(new_acc) > 0:
        msg = "\nPart 1 : New accounts to add for new month" + str(new_acc)

    else:
        msg = "\nPart 1 : No new account to add for new month!"

    return msg

def compare_summary_and_others():
    summary_check = [] #data of the rows in summary to check
    check = 0 # whether to check the sheet
    msg = ""
    for sheet in wb.sheetnames:
        if sheet == "SUMMARY - FS (000000)":
            for i in range (10, len(wb[sheet]['B'])+1):
                cell = wb[sheet]['B'][i]
                if cell.value is not None:
                    summary_check.append(str(cell.value).replace(" ",""))
                elif cell.value is None and i>=10:
                    if wb[sheet]['B'][i+1].value is None and wb[sheet]['B'][i+2].value:
                        check=1
                        break
                    else:
                        summary_check.append("")
        elif sheet == "Mapping":
            check=0
            
        elif check==1:
            for i in range (9, len(wb[sheet]['B'])+1):
                cell = wb[sheet]['B'][i]
                if cell.value is not None:
                    if str(cell.value).replace(" ","") != summary_check[i-10]:
                        print("Mismatch in line ", i+1, "with summary and ", sheet, cell.value,"!=", summary_check[i-10], "!\n")
                        msg = "\nPart 2 : Mismatch found in "+sheet
                        return msg
                elif cell.value is None and i>=10:
                    if wb[sheet]['B'][i+1].value is None and wb[sheet]['B'][i+2].value:
                        check=1
                        break

            print(sheet, "checking completed!\n")
    msg = "\nPart 2 : Individual Org checking completed succesfully!"
    return msg

#-----------------------------------------------------------------------------STEP_3---------------------------------------------------------------------------------
def add_account(account_number, account_name, account_type):
    global new_acc
    print("Start")
    accounts_index = {"Personnel Services":0,"Fringe Benefits":1,"Travel and Training":2,"Other Expenses":3,"Recovery":4}
    starting_row = [11]
    all_accounts = []
    sub = []
    row=10
    break_both = False
    while True:
        if break_both:
            break
        try:
            cell = wb["SUMMARY - FS (000000)"]['B'][row]
            number = int(cell.value)
            if(len(str(number))):
                sub.append(number)
        except:
            all_accounts.append(sub)
            sub = []
            if wb["SUMMARY - FS (000000)"]['B'][row+1].value==None:
                break_both = True
                break
            else:
                starting_row.append(row+2)
        row += 1
    print(starting_row)

    if account_type=="Personnel Services":
        insert_index = bisect.bisect_left(all_accounts[0], int(account_number))
        print("Insert at - " + str(insert_index + starting_row[0]))
        all_accounts[0].insert(insert_index,int(account_number))
        print(all_accounts[0])
    elif account_type=="Fringe Benefits":
        account_type="Fringe Benefits"
    elif account_type=="Travel and Training":
        account_type="Travel and Training"
    elif account_type=="Other Expenses":
        account_type="Other Expenses"
    elif account_type=="Recovery":
        account_type=="Recovery"
    
    print("New account "+str(account_number)+" - '"+str(account_name)+"' added successfully!") 
    try:
        new_acc.remove(int(str(account_number).replace(" ",'')))
    except:
        new_acc
    print(new_acc)

#-----------------------------------------------------------------------------STEP_4---------------------------------------------------------------------------------
