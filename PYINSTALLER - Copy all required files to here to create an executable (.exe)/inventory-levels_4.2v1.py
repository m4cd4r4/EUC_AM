# Local file_path to spreadsheet

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
file_path = os.path.join(application_path, './EUC_Perth_Assets.xlsx')

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
    # Convert width to integer
    width_int = int(width)
    plt.text(width + 1, bar.get_y() + bar.get_height()/2, width_int, ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=12)
plt.xlabel('Volume', fontsize=12)

# Set the range of the x-axis
plt.xlim(0, 120)

# Get current date in the format dd-mm-yyyy
current_date = datetime.now().strftime('%d-%m-%Y')

# Update the title to include current date
plt.title(f'Basement - 4.2 - Inventory Levels (Perth) - {current_date}', fontsize=14)

plt.tight_layout()

# Get current date and time in the format dd.mm.yy-hh.mm[am/pm] for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%H.%M.%S')

# Save the plot to a file with timestamp in the label
file_name = f'./Plots/4.2_inventory_levels_{current_datetime}.png'
plt.savefig(file_name)
plt.show()