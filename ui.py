import tkinter as tk
from tkinter import PhotoImage, filedialog
import warnings
import time
warnings.filterwarnings('ignore')
import test1

class StatusMsg(tk.Frame):
    def __init__(self, master, **kwargs):
        tk.Frame.__init__(self, master, **kwargs)
        self.status_label = tk.Label(self, text="Ready to go!", bd=1, relief=tk.SUNKEN, anchor=tk.CENTER, height=4, justify=tk.CENTER)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def progress(self, message):
        self.status_label.config(text=f"{message}")
        

class StatusBar(tk.Frame):
    def __init__(self, master, **kwargs):
        tk.Frame.__init__(self, master, **kwargs)
        self.canvas = tk.Canvas(self, height=20, width=1200, bg='white', borderwidth=2, relief=tk.SUNKEN)
        self.canvas.pack(fill=tk.X)

    def ongoingbar(self, button_number):
        part_width = 1200 // 9
        start_x = (button_number - 1) * part_width
        end_x = button_number * part_width
        self.canvas.create_rectangle(start_x, 0, end_x, 20, fill='yellow')

    def progressbar(self, button_number):
        part_width = 1200 // 9
        start_x = (button_number - 1) * part_width
        end_x = button_number * part_width
        self.canvas.create_rectangle(start_x, 0, end_x, 20, fill='green')

def button_click(button_number): # Button click triggering
    button_start(button_number)
    button_function(button_number)

def button_function(button_number): # Button functions on click
    msg = ""
    if button_number==1:
        file = upload_file()
        test1.excel_file = file.replace("/", "\\")
        monthly_file = upload_file_monthly()
        test1.monthly_file_name = monthly_file.replace("/", "\\")
        test1.accounts = test1.copy_monthly_sheet_data()
        window.after(500, button_done, msg, button_number)

    elif button_number==2:
        msg = test1.compare_account_numbers()
        msg += test1.compare_summary_and_others()
        window.after(500, button_done, msg, button_number)

    elif button_number==3:
        button3_input_window()
        if len(test1.new_acc)!=0:
            status_msg.progress(f"Step {button_number} going on!\nNew accounts to add for new month {str(test1.new_acc)}")
            status_bar.ongoingbar(button_number)
        else:
            msg = "\nNo more new accounts to add!"
            window.after(500, button_done, msg, button_number)
    
    elif button_number==4:
        msg = test1.update_monthly_expenses_into_organizations()
        window.after(500, button_done, msg, button_number)
        

def button_start(button_number): # Button click initially shows ongoing in the progress bar
    status_msg.progress(f"Step {button_number} going on!")
    status_bar.ongoingbar(button_number)

def button_done(msg, button_number): # When the button function is done, it shows completed in the progress bar
    status_bar.progressbar(button_number)
    status_msg.progress(f"Step {button_number} completed!{msg}")

def upload_file():
    file_path = filedialog.askopenfilename(title='Select the latest "Projected FY Budget" file')
    if file_path:
        return file_path
def upload_file_monthly():
    file_path = filedialog.askopenfilename(title='Select the monthly cognos "Statement of Revenue and Expense Detail" file')
    if file_path:
        return file_path

def on_enter(event):
    event.widget.config(bg='lightblue')  # Change color on hover
    window.config(cursor='hand2')  # Change cursor to hand

def on_leave(event):
    event.widget.config(bg='SystemButtonFace')  # Change back to the default color
    window.config(cursor='')  # Reset cursor to default

def button3_input_window():
    input_window = tk.Toplevel(window)
    input_window.title("Enter Account Details")

    # Account to add
    tk.Label(input_window, text="Account to add:").grid(row=0, column=0, padx=5, pady=5)
    tk.Label(input_window, text=str(test1.new_acc)).grid(row=0, column=1, padx=5, pady=5)
    
    # Account Number
    tk.Label(input_window, text="Account Number:").grid(row=1, column=0, padx=5, pady=5)
    account_number_entry = tk.Entry(input_window)
    account_number_entry.grid(row=1, column=1, padx=20, pady=10)

    # Account Name
    tk.Label(input_window, text="Account Name:").grid(row=2, column=0, padx=5, pady=5)
    account_name_entry = tk.Entry(input_window)
    account_name_entry.grid(row=2, column=1, padx=20, pady=10)

    # Account Type
    tk.Label(input_window, text="Account Type:").grid(row=3, column=0, padx=5, pady=5)
    account_types = ["Select Account Type", "Personnel Services", "Fringe Benefits", "Travel and Training", "Other Expenses", "Recovery"]
    account_type_var = tk.StringVar(input_window)
    account_type_var.set(account_types[0])  # Default value
    account_type_dropdown = tk.OptionMenu(input_window, account_type_var, *account_types)
    account_type_dropdown.grid(row=3, column=1, padx=20, pady=10)

    def submit_account_details(account_number, account_name, account_type):
        test1.add_account(account_number, account_name, account_type)
        print("New Account Added!:")
        print(f"Account Number: {account_number}")
        print(f"Account Name: {account_name}")
        print(f"Account Type: {account_type}")
        print(f"---------------------------------------")
        input_window.destroy()

    # Submit Button
    submit_button = tk.Button(input_window, text="Submit", command=lambda:submit_account_details( 
        account_number_entry.get(), account_name_entry.get(), account_type_var.get()))
    submit_button.grid(row=4, column=0, columnspan=2, padx=5, pady=10)

# Create the main window
window = tk.Tk()
window.title("NMSU Facilities and Services")

# Set the size of the window
window.geometry("1200x750") 

# Load a background image
background_image = PhotoImage(file="ui_bg.png")

# Create a frame for the background label using pack
background_frame = tk.Frame(window)
background_frame.place(relwidth=1, relheight=1)

# Create and place 9 buttons with some padding using grid
buttons_frame = tk.Label(window, image=background_image, anchor=tk.CENTER, justify=tk.CENTER)
buttons_frame.place(relwidth=1, relheight=1)

buttons = []
button_description = ["", "\nSelect most recent budget file and\nSelect MONTHLY COGNOS Input File",
                      "\nCurrent and Previous Month Accounts Comparison",
                      "\nInsert an account number to the organizational sheets\n(If required)",
                      "",
                      "",
                      "",
                      "",
                      "",
                      ""]
for i in range(1, 10):
    button = tk.Button(buttons_frame, text=f"STEP {i}{button_description[i]}", command=lambda i=i: button_click(i), padx=50, pady=20, width=40)
    button.grid(row=(i-1)//2, column=(i-1)%2, padx=10, pady=10)
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)
    buttons.append(button)


# Create a status bar
status_msg = StatusMsg(window, bg="lightgray")
status_msg.pack(side=tk.BOTTOM, fill=tk.X)
status_bar = StatusBar(window, bg="lightgray")
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Run the main loop
window.mainloop()


