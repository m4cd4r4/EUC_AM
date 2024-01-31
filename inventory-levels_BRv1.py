# Local file_path to spreadsheet

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Load the spreadsheet
file_path = 'EUC_Perth_Assets.xlsx'
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('BR Items')

# Replace NaN values with 0 in 'NewCount' column
df_items['NewCount'].fillna(0, inplace=True)

# Create a horizontal bar chart for the current inventory levels
plt.figure(figsize=(14, 10))
bars = plt.barh(df_items['Item'], df_items['NewCount'], color='#006aff', label='Volume')

# Define the spacing for the text
spacing = 0.6  # Adjust this value for more or less spacing

# Add the text with the summed count at the end of each bar
for bar in bars:
    width = bar.get_width()
    plt.text(width + spacing, bar.get_y() + bar.get_height()/2,
             f'{int(width)}', ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=14)
plt.xlabel('Volume', fontsize=14)

# Set the range of the x-axis
plt.xlim(0, 120)

# Get current date in the format dd-mm-yyyy
current_date = datetime.now().strftime('%d-%m-%Y')

# Update the title to include current date
plt.title(f'Level 6 - Build Room - Inventory Levels (Perth) - {current_date}', fontsize=16)

plt.legend()
plt.tight_layout()

# Get current date and time in the format dd.mm.yy-hh.mm[am/pm] for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%H.%M.%S')

# Save the plot to a file with timestamp in the label
file_name = f'Plots/BR_inventory_levels_{current_datetime}.png'
plt.savefig(file_name)
plt.show()
