import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from openpyxl import load_workbook

# Load Excel workbook and worksheet
workbook = load_workbook('assets.xlsx')
sheet = workbook.active

# ... [Functions remain unchanged]

# Initialize Tkinter root
root = tk.Tk()
root.title("EUC Asset Management GUI")

# Set the initial size of the window
root.geometry("800x600")

# Create a frame to hold the widgets
frame = tk.Frame(root)
frame.place(relx=0.1, rely=0.1, relwidth=0.8, relheight=0.8)

# Banner label at the top
banner = tk.Label(frame, text="EUC Asset Management GUI", font=("Arial", 16))
banner.grid(row=0, columnspan=2, pady=20)

# Dropdown for assets
combobox_assets = ttk.Combobox(frame, values=get_asset_types(), width=40)
combobox_assets.grid(row=1, columnspan=2, padx=20, pady=10)

# Entry field for new asset type
entry_new_asset = tk.Entry(frame, width=40)
entry_new_asset.grid(row=2, columnspan=2, padx=20, pady=10)

# Button for adding new asset type
button_add_asset = tk.Button(frame, text="Add Asset Type", command=add_new_asset_type, width=20, height=2)
button_add_asset.grid(row=3, column=0, padx=20, pady=10)

# Button for adding model
button_add_model = tk.Button(frame, text="Add Model", command=add_model, width=20, height=2)
button_add_model.grid(row=3, column=1, padx=20, pady=10)

# Entry field for quantity
entry_quantity = tk.Entry(frame, width=40)
entry_quantity.grid(row=4, columnspan=2, padx=20, pady=10)

# Buttons for updating quantity
button_add = tk.Button(frame, text="Add Quantity", command=lambda: update_quantity(True), width=20, height=2)
button_add.grid(row=5, column=0, padx=20, pady=10)

button_subtract = tk.Button(frame, text="Subtract Quantity", command=lambda: update_quantity(False), width=20, height=2)
button_subtract.grid(row=5, column=1, padx=20, pady=10)

# Start the GUI event loop
root.mainloop()
