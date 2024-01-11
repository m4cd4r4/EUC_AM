# This code performs the following steps:

# It reads the '4.2 Timestamps' sheet from the Excel file into a DataFrame.
# It converts the 'Timestamp' column into a datetime object to enable time series analysis.
# It groups the data by date and 'Action' and counts the occurrences of each action type per day.
# It plots this aggregated data, using different markers for 'add' and 'subtract' actions.
# It saves the plot as a PNG file to the specified path.
# This plot is useful for visualizing the flow of inventory actions over time and identifying any trends or patterns in the data.

import pandas as pd
import matplotlib.pyplot as plt

# Define the file path (change this to your actual file path)
file_path = 'C:/Users/Madhous/Documents/GitHub/EUD_AM/Build Room/EUC_Perth_Assets.xlsx'

# Create an instance of ExcelFile to work with
xl = pd.ExcelFile(file_path)

# Now you can use xl to parse sheets
df_timestamps = xl.parse('4.2 Timestamps')

# Load the "Timestamps" sheet into a DataFrame by name: df_timestamps
df_timestamps = xl.parse('4.2 Timestamps')

# Convert 'Timestamp' column to datetime type
df_timestamps['Timestamp'] = pd.to_datetime(df_timestamps['Timestamp'])

# Create a new DataFrame that counts the number of 'add' and 'subtract' actions per day
df_actions_per_day = df_timestamps.groupby([df_timestamps['Timestamp'].dt.date, 'Action']).size().unstack(fill_value=0)

# Plot this data on a time series chart
plt.figure(figsize=(14, 7))
plt.plot(df_actions_per_day.index, df_actions_per_day['add'], label='Add', marker='o')
plt.plot(df_actions_per_day.index, df_actions_per_day['subtract'], label='Subtract', marker='x')

plt.xlabel('Date', fontsize=14)
plt.ylabel('Number of Actions', fontsize=14)
plt.title('Inventory Actions Over Time', fontsize=16)
plt.legend()
plt.grid(True)
plt.tight_layout()

# Save the plot to a file
actions_plot_path = 'C:/Users/Madhous/Documents/GitHub/EUD_AM/Asset Management/Visuals - Build Room/Screenshots/inventory_actions_over_time.png'
plt.savefig(actions_plot_path)
plt.close()  # Close the plot to avoid displaying it inline in this notebook
