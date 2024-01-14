import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Load the spreadsheet
file_path = 'C:/Users/Administrator/Documents/Github/EUC_AM/EUC_Perth_Assets.xlsx'
xl = pd.ExcelFile(file_path)

# Load sheets into DataFrames by name
df_42_items = xl.parse('4.2 Items')
df_br_items = xl.parse('BR Items')

# Replace NaN values with 0 in 'NewCount' column for both dataframes
df_42_items['NewCount'].fillna(0, inplace=True)
df_br_items['NewCount'].fillna(0, inplace=True)

# Combine the two dataframes
combined_df = pd.concat([df_42_items, df_br_items])

# Group by 'Item' and sum the 'NewCount' values
grouped_df = combined_df.groupby('Item')['NewCount'].sum().reset_index()

# Create a horizontal bar chart for the summed inventory levels
plt.figure(figsize=(14, 10))
bars = plt.barh(grouped_df['Item'], grouped_df['NewCount'], color='skyblue')

# Define the spacing for the text
spacing = 1  # Adjust this value for more or less spacing

# Add the text with the summed count at the end of each bar
for bar in bars:
    width = bar.get_width()
    plt.text(width + spacing, bar.get_y() + bar.get_height()/2,
             f'{int(width)}', ha='left', va='center', color='black')

plt.ylabel('Item', fontsize=14)
plt.xlabel('Volume', fontsize=14)

# Get current date in the format dd-mm-yyyy
current_date = datetime.now().strftime('%d-%m-%Y')

# Update the title to include current date
plt.title(f'4.2 & BR Combined - Total Inventory Levels (Perth) - {current_date}', fontsize=16)

plt.legend()
plt.tight_layout()

# Get current date and time in the format dd.mm.yy-hh.mm[am/pm] for file name
current_datetime = datetime.now().strftime('%d.%m.%y-%I.%M%p')

# Save the plot to a file with timestamp in the label
file_name = f'C:/Users/Administrator/Documents/Github/EUC_AM/Plots/combined_inventory_levels_{current_datetime}.png'
plt.savefig(file_name)
plt.show()
