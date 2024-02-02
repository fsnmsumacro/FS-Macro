import tkinter as tk
from tkinter import PhotoImage

class StatusBar(tk.Frame):
    def __init__(self, master, **kwargs):
        tk.Frame.__init__(self, master, **kwargs)
        self.status_label = tk.Label(self, text="Status: Ready", bd=1, relief=tk.SUNKEN, anchor=tk.CENTER, height=4, justify=tk.CENTER)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def progress(self, message):
        self.status_label.config(text=f"Status: {message}")

def button_click(button_number):
    status_bar.progress(f"Button {button_number} execution done!")
    # Add your button execution logic here

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
window.geometry("1200x600")  # Set your desired width and height

# Load a background image
background_image = PhotoImage(file="ui_bg.png")  # Replace with your image file path

# Create a frame for the background label using pack
background_frame = tk.Frame(window)
background_frame.place(relwidth=1, relheight=1)

# Create and place 9 buttons with some padding using grid
buttons_frame = tk.Label(window, image=background_image)
buttons_frame.place(relwidth=1, relheight=1)

buttons = []
button_description = ["", "\nSelect MONTHLY COGNOS Input File",
                      "\nCurrent and Previous Month Accounts Comparison",
                      "",
                      "",
                      "",
                      "",
                      "",
                      "",
                      ""]
for i in range(1, 10):
    button = tk.Button(buttons_frame, text=f"Button {i}{button_description[i]}", command=lambda i=i: button_click(i), padx=50, pady=20)
    button.grid(row=(i-1)//2, column=(i-1)%2, padx=10, pady=10)
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)
    buttons.append(button)

# Create a status bar
status_bar = StatusBar(window, bg="lightgray")
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Run the main loop
window.mainloop()
