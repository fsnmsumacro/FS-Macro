import tkinter as tk
from tkinter import PhotoImage
import test1

def button_click(button_number):
    print(f"Button {button_number} clicked!")
    if button_number==1:
        test1.accounts = test1.copy_monthly_sheet_data()
    elif button_number==2:
        test1.compare_account_numbers()

def on_enter(event):
    event.widget.config(bg='lightblue')  # Change color on hover
    window.config(cursor='hand2')  # Change cursor to hand

def on_leave(event):
    event.widget.config(bg='SystemButtonFace')  # Change back to the default color
    window.config(cursor='')  # Reset cursor to default

# Create the main window
window = tk.Tk()
window.title("Button Example")

# Set the size of the window
window.geometry("1200x750") 

# Load a background image
background_image = PhotoImage(file="ui_bg.png")  # Replace with your image file path

# Create a label to hold the background image
background_label = tk.Label(window, image=background_image)
background_label.place(relwidth=1, relheight=1)

# Create and place 9 buttons with some padding
buttons = []
for i in range(1, 10):
    button = tk.Button(window, text=f"Step {i}", command=lambda i=i: button_click(i), padx=80, pady=10, border=5)
    button.grid(row=(i-1)//3, column=(i-1)%3, padx=60, pady=10)
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)
    buttons.append(button)

# Run the main loop
window.mainloop()

