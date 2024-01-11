import pandas as pd
import matplotlib.pyplot as plt

# Load the spreadsheet
file_path = 'C:/Users/Madhous/Documents/GitHub/EUD_AM/Build Room/EUC_Perth_Assets.xlsx'
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('4.2 Items')

# Create a bar chart for the current inventory levels
plt.figure(figsize=(14, 10))
bars = plt.bar(df_items['Item'], df_items['NewCount'], color='skyblue', label='Volume')

# Add the text with the count on each bar
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval - 5, yval, ha='center', va='bottom', color='black')

plt.xlabel('Item', fontsize=14)
plt.ylabel('Volume', fontsize=14)
plt.title('Current Inventory Levels in the Store-room', fontsize=16)
plt.xticks(rotation=90)
plt.legend()
plt.tight_layout()

# Save the plot to a file
plt.savefig('C:/Users/Madhous/Documents/GitHub/EUD_AM/Build Room/inventory_levels.png')
plt.show()
