from tkinter import *
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import os
import excel_manager
 
# important variables
height = 600 #height of the window
width = 900 # width the window
windowBg = "#0d0d0d" #window background color
fontcolor = "#eeeeee" # windor fontcolor
bordercolor = "#888888" # entry border color
bgColor = "#1a1a1a" # entry background color

# Functions 

# Location function
def browseLocation():
    location = filedialog.askopenfilename(title="Select an Excel File", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    if location:
        locationBox.delete(0, END)  # Clear the current text in the entry box
        locationBox.insert(0, location)  # Insert the selected file path into the entry box

# Generate data
def generatData():
    input_value = rangebox.get() # get input value of rangebox
    location = locationBox.get() # get input value of locationBox
    if input_value == "": # check if input value is empty
        setError.configure(text="Please enter a range") 
        rangebox.focus()
    elif not input_value.isdigit(): # check if input value is a digit
        setError.configure(text="Only numbers are allowed")
        print(type(input_value))
        rangebox.focus()
    elif int(input_value) <= 0: # check if input value is less than or equal to 0
        setError.configure(text="Please enter a positive number")
        rangebox.focus()
    elif location == "": # check if location is empty
        setError.configure(text="Please enter a location")
        locationBox.focus()
    elif not location.endswith(".xlsx"): # check if location ends with .xlsx
        setError.configure(text="Please enter a valid Excel file location")
        locationBox.focus()
    else:
        setError.configure(text="")  # Clear the error message
        print(f"Entered Range: {input_value}")
        print(f"Location: {location}")  # Simulate data generation
        data = excel_manager.generate_data(int(input_value), location)  # Call the data generation function
        if data: # Check if data generation was successful
            messagebox.showinfo("Success", f"Data generated successfully in {location}")
        else: # If data generation failed, show an error message
            messagebox.showerror("Error", "Failed to generate data. Please check the file path and try again.")

# Main window
window = tk.Tk()
window.title("Pyxcel")
window.geometry(f"{width}x{height}+230+50")
window.resizable(width=FALSE, height=FALSE)
window.configure(bg=windowBg)

menu_bar = tk.Menu(window, bg=bgColor, fg=fontcolor, activebackground=bordercolor, activeforeground=fontcolor)
file_menu = tk.Menu(menu_bar, tearoff=0) 
menu_bar.add_cascade(label="Home", menu=file_menu)  
theme_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Theme", menu=theme_menu)
theme_menu.add_command(label="Dark", command=lambda: None)  # Placeholder for theme change
theme_menu.add_command(label="Light", command=lambda: None)  # Placeholder for theme change
help_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Help")
window.config(menu=menu_bar, background=bgColor)

# Title label
label = tk.Label(window, text="Excel Data Generator", font=("Monospace", 20), fg=fontcolor, background=bgColor)
label.pack(pady=30)

# Input label and entry
tk.Label(window, font=("Arial", 13), text="Enter data range you want (e.g. 100)", fg=fontcolor, background=bgColor).pack(pady=4)
rangebox = tk.Entry(window, font=("Monospace", 13), width=35, insertbackground="#cccccc", background=bgColor, fg="#cccccc", justify=CENTER, highlightthickness=1, highlightbackground=bordercolor, relief=FLAT)
rangebox.pack(ipady=5, pady=5)


tk.Label(window, font=("Arial", 13), text="Enter Excel file location", fg=fontcolor, background=bgColor).pack(pady=4)
locationBox = tk.Entry(window, font=("Monospace", 13), width=35, insertbackground="#cccccc", bg=bgColor, fg="#cccccc", justify=CENTER, highlightthickness=1, highlightbackground=bordercolor, relief=FLAT)
locationBox.pack(ipady=5, pady=5)

browseBtn = tk.Button(window, text="Browse", font=("Arial", 10), command=browseLocation, cursor="hand2")
browseBtn.place(x=552, y=217)
browseBtn.config(background="#000000", fg="#FFFFFF", relief=FLAT)
# Buttons
Submit = tk.Button(window, text="Generate", font=("Arial", 14), command=generatData, cursor="hand2")
Submit.pack(pady=10)
Submit.config(background="#fa0000", fg="#FFFFFF", width=28, relief=FLAT)           

# Error label
setError = tk.Label(window, font=("Monospace", 10), background=bgColor, fg="#ff0000")
setError.pack()

# Run the application
window.mainloop()