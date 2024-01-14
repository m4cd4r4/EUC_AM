import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook
from openpyxl import Workbook
import tkinter.font as tkFont
from datetime import datetime
import os

# Initialize Tkinter root
root = tk.Tk()
root.title("Store-Room 4.2")
root.geometry("800x600")

# Define a large font for buttons and treeview
button_font = tkFont.Font(size=16)
treeview_font = tkFont.Font(size=12)  # Adjust the size as needed for tablet readability

# Function to update the Treeview widget with the spreadsheet data
# ... [rest of the update_treeview function]

# Function to log the changes to the second sheet and update the log view
# ... [rest of the log_change function]

# Function to update the log view with the 5 most recent changes
# ... [rest of the update_log_view function]

# Function to update counts in the spreadsheet
# ... [rest of the update_count function]

# Create a frame to hold the widgets
frame = tk.Frame(root)
frame.pack(padx=10, pady=10, fill='both', expand=True)

# Entry field and buttons layout
entry_frame = tk.Frame(frame)
entry_frame.pack(pady=10)

# Subtract Button
button_subtract = tk.Button(entry_frame, text="-", command=lambda: update_count('subtract'), font=button_font)
button_subtract.pack(side=tk.LEFT, padx=5)

# Entry Value
entry_value = tk.Entry(entry_frame, width=10, font=button_font)
entry_value.pack(side=tk.LEFT)

# Add Button
button_add = tk.Button(entry_frame, text="+", command=lambda: update_count('add'), font=button_font)
button_add.pack(side=tk.LEFT, padx=5)

# Treeview for displaying spreadsheet data
columns = ("Item", "LastCount", "NewCount")
tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode='browse', height=10)
tree.heading("Item", text="Item", anchor=tk.W)  # Align "Item" text to the left
tree.column("Item", anchor=tk.W, stretch=False)
tree.column("LastCount", anchor=tk.CENTER, stretch=False)
tree.column("NewCount", anchor=tk.CENTER, stretch=False)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=200, minwidth=50, stretch=False)  # Set minimum width and disable stretching
tree.pack(expand=True, fill="both", padx=10, pady=10)
tree.bind('<ButtonRelease-1>', lambda e: entry_value.focus())  # Focus on entry field when an item is selected
tree.tag_configure('treeview', font=treeview_font)

# Log view for displaying the 5 most recent changes
# ... [rest of the log_view setup]

# Load the logo image and place it at the top right corner
logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '1280px-Santos_limited_corporate_logo.svg.png')
logo_image = tk.PhotoImage(file=logo_path)
# Calculate an appropriate width for the logo (20% of the window's width)
logo_width = root.winfo_width() * 0.2
logo_label = tk.Label(root, image=logo_image, width=int(logo_width))
logo_label.pack(side=tk.TOP, anchor='ne', padx=20, pady=20)

# Initially populate the Treeview and the Log View
update_treeview()
update_log_view()

# Start the GUI event loop
root.mainloop()
