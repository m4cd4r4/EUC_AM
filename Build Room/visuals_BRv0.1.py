import pandas as pd
import matplotlib.pyplot as plt

# Load the spreadsheet
file_path = 'C:/Users/Madhous/Documents/GitHub/EUD_AM/Build Room/EUC_Perth_Assets.xlsx'
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('4.2 Items')

# Create a bar chart for the current inventory levels
plt.figure(figsize=(10, 8))
plt.bar(df_items['Item'], df_items['NewCount'], color='skyblue')
plt.xlabel('Item', fontsize=14)
plt.ylabel('New Count', fontsize=14)
plt.title('Current Inventory Levels in the Store-room', fontsize=16)
plt.xticks(rotation=90)
plt.tight_layout()

# Save the plot to a file
plt.savefig('C:/Users/Madhous/Documents/GitHub/EUD_AM/Build Room/inventory_levels.png')
plt.show()