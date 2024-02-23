import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Check if the application is "frozen"
if getattr(sys, 'frozen', False):
    # If it's frozen, use the path relative to the executable
    application_path = sys._MEIPASS
else:
    # If it's not frozen, use the path relative to the script file
    application_path = os.path.dirname(__file__)

# Construct the path to the file
file_path = os.path.join(application_path, 'EUC_Perth_Assets.xlsx')

# Load the spreadsheet
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('4.2 Items')

# Replace NaN values with 0 in 'NewCount' column
df_items['NewCount'].fillna(0, inplace=True)

# Create a horizontal bar chart for the current inventory levels
plt.figure(figsize=(14 * 0.60, 10 * 0.60))
bars = plt.barh(df_items['Item'], df_items['NewCount'], color='#006aff', label='Volume')

# Add the text with the count at the end of each bar
for bar in bars:
    width = bar.get_width()
    width_int = int(width)
    plt.text(width + 1, bar.get_y() + bar.get_height()/2, width_int, ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=12)
plt.xlabel('Volume', fontsize=12)
plt.xlim(0, 120)
current_date = datetime.now().strftime('%d-%m-%Y')
plt.title(f'Basement - 4.2 - Inventory Levels (Perth) - {current_date}', fontsize=14)
plt.tight_layout()

# Ensure 'Plots' folder exists
plots_folder = os.path.join(application_path, 'Plots')
if not os.path.exists(plots_folder):
    os.makedirs(plots_folder)

# Get current date and time for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%H.%M.%S')

# Construct the full file path for saving the plot
file_name = os.path.join(plots_folder, f'4.2_Inventory_Levels_{current_datetime}.png')

# Save and show the plot
plt.savefig(file_name)
plt.show()
