import pandas as pd
import matplotlib.pyplot as plt
import datetime

# Define the file path (change this to your actual file path)
file_path = 'C:/Users/Madhous/Documents/GitHub/EUD_AM/Build Room/EUC_Perth_Assets.xlsx'

# Create an instance of ExcelFile to work with
xl = pd.ExcelFile(file_path)

# Load a sheet into a DataFrame by name: df_items
df_items = xl.parse('4.2 Items')

# Create a box plot for the 'NewCount' of each item
plt.figure(figsize=(14, 10))
df_items.boxplot(column='NewCount', by='Item', vert=False)

# Add labels and title
plt.xlabel('Volume', fontsize=14)
plt.title('Box Plot of Inventory Volume by Item', fontsize=16)
plt.suptitle('')  # Suppress the default title to only show our custom title

# Get the current date and time
current_time = datetime.datetime.now()

# Format the timestamp
timestamp_str = current_time.strftime("%H%M%S")

# Define the path for the boxplot image with the timestamp
boxplot_path = f'C:/Users/Madhous/Documents/GitHub/EUD_AM/Asset Management/Visuals - Build Room/Screenshots/boxplot_inventory_volume_{timestamp_str}.png'

# Save the plot to a file
plt.savefig(boxplot_path)
plt.close()  # Close the plot to avoid displaying it inline in this notebook
