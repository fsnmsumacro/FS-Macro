import openpyxl as xl
import bisect
import string

# excel file used here
excel_file = ""
monthly_file_name = ""
wb = None
accounts = []
new_acc = []
current_month = ""
cols = {"Jul":['R',1,1], "Aug":['S',3,2], "Sep":['T',5,3], "Oct":['U',7,4], "Nov":['V',9,5], "Dec":['W',11,6],
        "Jan":['X',13,7], "Feb":['Y',15,8], "Mar":['Z',17,9], "Apr":['AA',19,10], "May":['AB',21,11], "Jun":['AC',23,12],
        "1st Close":['AD',23,12], "2nd Close":['AE',23,12], "3rd Close":['AF',23,12], "Final Close":['AG',23,12]} 

#-----------------------------------------------------------------------------STEP_1---------------------------------------------------------------------------------
def copy_monthly_sheet_data():
    # load the workbook
    global wb, current_month
    wb = xl.load_workbook(excel_file, keep_vba=True) 
    sheets = wb.sheetnames

    # Clear all data from the destination sheet
    destination_sheet = wb["FI-U227 Statement of Revenu (2"]
    destination_sheet.delete_rows(1, destination_sheet.max_row)

    # read monthly file
    #monthly_file_name = "C:\\Users\\indronil\\Downloads\\FI-U227 OCT Statement of Revenue and Expense Detail.xlsx"
    monthly_file = xl.load_workbook(monthly_file_name)
    current_month = monthly_file_name.split("\\")[-1][8:11].capitalize()
    print(current_month)
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
    
    
    accounting_style = xl.styles.NamedStyle(name='accounting', number_format='_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)')

    print("New month's data added to 'FI-U227 Statement of Revenu (2' sheet!")
    accounts_monthly = list(set(accounts_monthly))
    wb.save("New_modified.xlsm")
    print(f"---------------------------------------")
    return accounts_monthly

#-----------------------------------------------------------------------------STEP_2---------------------------------------------------------------------------------
def compare_account_numbers():
    global new_acc, wb
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
                        msg = "\nPart 2 : Checking stopped! Mismatch found in "+sheet
                        return msg
                elif cell.value is None and i>=10:
                    if wb[sheet]['B'][i+1].value is None and wb[sheet]['B'][i+2].value:
                        check=1
                        break

            print(sheet, "checking completed!")
    msg = "\nPart 2 : Individual Org checking completed succesfully!"
    print(f"---------------------------------------")
    return msg

#-----------------------------------------------------------------------------STEP_3---------------------------------------------------------------------------------
def add_account(account_number, account_name, account_type):
    global new_acc
    accounts_index = {"Personnel Services":0,"Fringe Benefits":1,"Travel and Training":2,"Other Expenses":3,"Recovery":4}
    starting_row = [11]
    closing_row = []
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
                closing_row.append(row)
                break_both = True
                break
            else:
                closing_row.append(row+1)
                starting_row.append(row+2)
        row += 1
    print("Start: ", starting_row, " Close: ", closing_row)

    def change_other_account_type_starting_closing(starting_row, index): # Change account type starting and closing after an insertion
        for i in range(index+1,len(starting_row)):
            starting_row[i]+=1
        for i in range(index,len(starting_row)):
            closing_row[i]+=1
        return starting_row, closing_row
    
    def search_insert_position_and_insert(all_accounts, starting_row, index):
        accounting_style = xl.styles.NamedStyle(name='accounting', number_format='_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)')
        sheet_to_insert = wb["SUMMARY - FS (000000)"]
        insert_index = bisect.bisect_left(all_accounts[index], int(account_number))
        print("Insert at - " + str(insert_index + starting_row[index]))
        num_row = insert_index + starting_row[index]
        all_accounts[index].insert(insert_index,int(account_number))
        starting_row, closing_row = change_other_account_type_starting_closing(starting_row, index)
        insert = 0 # variable to check which sheets to insert
        for sheet_to_insert in wb.sheetnames: #inserting to all the pages
            if sheet_to_insert == "SUMMARY - FS (000000)":
                insert = 1
            elif sheet_to_insert == "Mapping":
                insert = 0
            if insert==1:
                wb[sheet_to_insert].insert_rows(num_row)
                row = str(num_row)
                wb[sheet_to_insert][num_row][0].value = account_name
                wb[sheet_to_insert][num_row][1].value = account_number
                wb[sheet_to_insert][num_row][2].value = 0
                wb[sheet_to_insert][num_row][3].value = "=C"+ row
                wb[sheet_to_insert][num_row][4].value = "=D"+ row
                wb[sheet_to_insert][num_row][5].value = "=SUM(R"+row+":AF"+row+")"
                wb[sheet_to_insert][num_row][7].value = "=+D"+row+"-F"+row+"-G"+row
                wb[sheet_to_insert][num_row][10].value = "=(((+F"+row+"-J"+row+")/$A$5)*12)+J"+row
                for r in range(11,32):
                    if r!=16:
                        wb[sheet_to_insert][num_row][r].value = 0
                for r in range(0,32):
                    if r!=16:
                        if num_row in starting_row:
                            wb[sheet_to_insert][num_row][r].style = wb[sheet_to_insert][num_row+1][r].style
                        else:
                            wb[sheet_to_insert][num_row][r].style = wb[sheet_to_insert][num_row-1][r].style
        return all_accounts, starting_row

    if account_type=="Personnel Services": # calling insert function according to account type
        all_accounts, starting_row = search_insert_position_and_insert(all_accounts, starting_row, 0)
    elif account_type=="Fringe Benefits":
        all_accounts, starting_row = search_insert_position_and_insert(all_accounts, starting_row, 1)
    elif account_type=="Travel and Training":
        all_accounts, starting_row = search_insert_position_and_insert(all_accounts, starting_row, 2)
    elif account_type=="Other Expenses":
        all_accounts, starting_row = search_insert_position_and_insert(all_accounts, starting_row, 3)
    elif account_type=="Recovery":
        all_accounts, starting_row = search_insert_position_and_insert(all_accounts, starting_row, 4)
    
    print("Start: ", starting_row, " Close: ", closing_row)
    print("New account "+str(account_number)+" - '"+str(account_name)+"' added successfully!") 
    try:
        new_acc.remove(int(str(account_number).replace(" ",'')))
    except:
        new_acc

    print("Account to be added: ", new_acc, "\n")
    wb.save("New_modified.xlsm")
#-----------------------------------------------------------------------------STEP_4---------------------------------------------------------------------------------

#def update_formula():
    sheets = wb.get_sheet_names()
    summary = wb["SUMMARY - FS (000000)"]
    for row in range(starting_row[0], closing_row[4]+1):
        for col in range(2,32): #Number of columns=32
            val = ""
            if row not in closing_row:
                    val += "="
                    start = 0
                    for sheet in sheets:
                        if sheet == "Mapping":
                            start = 0
                        if start == 1:
                            if(col<=25):
                                col_letter = str(string.ascii_uppercase[col])
                            else:
                                col_letter = str(string.ascii_uppercase[int((col/25))-1]) + str(string.ascii_uppercase[int(col%25)-1])
                            val += "+'"+sheet+"'!"+col_letter+str(row)
                        if sheet == "SUMMARY - FS (000000)":
                            start = 1
            summary[str(col_letter)+str(row)].value=val
    wb.save("New_modified.xlsm")
#-----------------------------------------------------------------------------STEP_5---------------------------------------------------------------------------------
def update_monthly_expenses_into_organizations():
    global wb, cols
    sheets = wb.get_sheet_names()
    account_dict = {}
    for cell in wb["SUMMARY - FS (000000)"]['B']:
        try:
            account_number = str(cell.value)
            if(len(str(account_number))):
                account_dict.update({account_number: int(cell.coordinate.replace('B',''))})
        except:
            account_number=0
    #print(account_dict)
    
    for row in wb['FI-U227 Statement of Revenu (2'].iter_rows(min_row=2):
        for org in sheets:
            if (row[3].value.split('-')[0]) in str(org):
                ins_org = org # get the organization to insert
                ins_amount = float(row[8].value) # get the value to insert
                break
        ins_cell = cols[current_month][0] + str(account_dict[row[5].value])  # get the cell to insert
        wb[ins_org][ins_cell].value = ins_amount
    print("Accounts updated with the monthly expenses\n-------------------------------------")
    msg = "\nAccounts updated with the monthly expenses"
    wb.save("New_modified.xlsm")
    return msg

#-----------------------------------------------------------------------------STEP_6---------------------------------------------------------------------------------
