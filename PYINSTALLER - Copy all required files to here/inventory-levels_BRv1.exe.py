# File-path is to the folder being shared with the executable
# User manually puts the spreadsheet in the same folder as the .exe

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import os
import sys

# Determine the directory where the EXE is located
exe_dir = os.path.dirname(sys.executable)

# Construct the path to the spreadsheet in the same directory
file_path = os.path.join(exe_dir, 'EUC_Perth_Assets.xlsx')

# # Load the spreadsheet
# file_path = 'C:/Users/Administrator/Documents/Github/EUC_AM/EUC_Perth_Assets.xlsx'
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('4.2 Items')

# Replace NaN values with 0 in 'NewCount' column
df_items['NewCount'].fillna(0, inplace=True)

# Create a horizontal bar chart for the current inventory levels
plt.figure(figsize=(14, 10))
bars = plt.barh(df_items['Item'], df_items['NewCount'], color='#006aff', label='Volume')

# Add the text with the count at the end of each bar
for bar in bars:
    width = bar.get_width()
    # Convert width to integer
    width_int = int(width)
    plt.text(width + 1, bar.get_y() + bar.get_height()/2, width_int, ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=14)
plt.xlabel('Volume', fontsize=14)

# Get current date in the format dd-mm-yyyy
current_date = datetime.now().strftime('%d-%m-%Y')

# Update the title to include current date
plt.title(f'Basement 4.2 - Inventory Levels (Perth) - {current_date}', fontsize=16)

plt.legend()
plt.tight_layout()

# Get current date and time in the format dd.mm.yy-hh.mm[am/pm] for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%I.%M%p')

# Save the plot to a file with timestamp in the label
file_name = f'C:/Users/Administrator/Documents/Github/EUC_AM/Plots/inventory_levels_4.2_{current_datetime}.png'
plt.savefig(file_name)
plt.show()
