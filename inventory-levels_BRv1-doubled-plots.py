import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Load the spreadsheet
file_path = './EUC_Perth_Assets.xlsx'
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('BR Items')

# Replace NaN values with 0 in 'NewCount' column
df_items['NewCount'].fillna(0, inplace=True)

# Create a horizontal bar chart for the current inventory levels
plt.figure(figsize=(14 * 0.60, 10 * 0.60))
bars = plt.barh(df_items['Item'], df_items['NewCount'], color='#006aff', label='Volume')

# Define the spacing for the text
spacing = 0.6  # Adjust this value for more or less spacing

# Add the text with the summed count at the end of each bar
for bar in bars:
    width = bar.get_width()
    plt.text(width + spacing, bar.get_y() + bar.get_height()/2,
             f'{int(width)}', ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=12)
plt.xlabel('Volume', fontsize=12)
plt.xlim(0, 120)
current_date = datetime.now().strftime('%d-%m-%Y')
plt.title(f'Level 6 - Build Room - Inventory Levels (Perth) - {current_date}', fontsize=14)
plt.tight_layout()

# Ensure 'Plots' folder exists
plots_folder = 'Plots'
if not os.path.exists(plots_folder):
    os.makedirs(plots_folder)

# Get current date and time for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%H.%M.%S')

# Construct the full file path for saving the plot
file_name = os.path.join(plots_folder, f'BR_inventory_levels_{current_datetime}.png')

# Save and show the plot
plt.savefig(file_name)
plt.show()
import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Function to get the directory of the executable or script
def get_base_path():
    if getattr(sys, 'frozen', False):
        # If the application is frozen (executable)
        return os.path.dirname(sys.executable)
    else:
        # If it's run as a script
        return os.path.dirname(os.path.abspath(__file__))

base_path = get_base_path()

# Load the spreadsheet
file_path = os.path.join(base_path, 'EUC_Perth_Assets.xlsx')
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('BR Items')

# Replace NaN values with 0 in 'NewCount' column
df_items['NewCount'].fillna(0, inplace=True)

# Create a horizontal bar chart for the current inventory levels
plt.figure(figsize=(14 * 0.60, 10 * 0.60))
bars = plt.barh(df_items['Item'], df_items['NewCount'], color='#006aff', label='Volume')

# Define the spacing for the text
spacing = 0.6  # Adjust this value for more or less spacing

# Add the text with the summed count at the end of each bar
for bar in bars:
    width = bar.get_width()
    plt.text(width + spacing, bar.get_y() + bar.get_height()/2,
             f'{int(width)}', ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=12)
plt.xlabel('Volume', fontsize=12)
plt.xlim(0, 120)
current_date = datetime.now().strftime('%d-%m-%Y')
plt.title(f'Level 6 - Build Room - Inventory Levels (Perth) - {current_date}', fontsize=14)
plt.tight_layout()

# Ensure 'Plots' folder exists in the same directory as the executable/script
plots_folder = os.path.join(base_path, 'Plots')
if not os.path.exists(plots_folder):
    os.makedirs(plots_folder)

# Get current date and time for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%H.%M.%S')

# Construct the full file path for saving the plot
file_name = os.path.join(plots_folder, f'BR_inventory_levels_{current_datetime}.png')

# Save and show the plot
plt.savefig(file_name)
plt.show()
