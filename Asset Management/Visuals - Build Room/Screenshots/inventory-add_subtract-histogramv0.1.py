# # NOT WORKING YET


import pandas as pd
import matplotlib.pyplot as plt
import datetime

# Define the file path for the Excel file
file_path = 'C:/Users/Madhous/Documents/GitHub/EUD_AM/Build Room/EUC_Perth_Assets.xlsx'

# Create an instance of ExcelFile to work with
xl = pd.ExcelFile(file_path)

# Load the "Timestamps" sheet into a DataFrame
df_timestamps = xl.parse('4.2 Timestamps')

# Convert 'Timestamp' column to datetime type and drop rows with NaT values
df_timestamps['Timestamp'] = pd.to_datetime(df_timestamps['Timestamp'], errors='coerce')
df_timestamps.dropna(subset=['Timestamp'], inplace=True)

# Filter 'add' and 'subtract' actions for SANs only and drop rows with missing timestamps
add_sans = df_timestamps[(df_timestamps['Action'] == 'add') & (df_timestamps['Item'].str.contains('SAN'))].dropna(subset=['Timestamp'])
subtract_sans = df_timestamps[(df_timestamps['Action'] == 'subtract') & (df_timestamps['Item'].str.contains('SAN'))].dropna(subset=['Timestamp'])

# Check if the dataframes are not empty before proceeding
if not add_sans.empty and not subtract_sans.empty:
    # Create a combined histogram for 'add' and 'subtract' actions
    plt.figure(figsize=(14, 7))

    # Determine the number of bins based on the earliest and latest dates in both datasets
    min_date = min(add_sans['Timestamp'].dt.date.min(), subtract_sans['Timestamp'].dt.date.min())
    max_date = max(add_sans['Timestamp'].dt.date.max(), subtract_sans['Timestamp'].dt.date.max())
    bins = pd.date_range(min_date, max_date, freq='D')

    # Overlay histograms with equal binning, opacity, and solid borders
    plt.hist(add_sans['Timestamp'].dt.date, bins=bins, color='green', alpha=0.5, edgecolor='black', label='Add SANs')
    plt.hist(subtract_sans['Timestamp'].dt.date, bins=bins, color='red', alpha=0.5, edgecolor='black', label='Subtract SANs')

    # Add titles and labels
    plt.title('Equal Width Frequency Histogram for SAN Actions', fontsize=16)
    plt.xlabel('Date', fontsize=14)
    plt.ylabel('Count', fontsize=14)
    plt.legend()

    plt.tight_layout()

    # Save the corrected plot with a timestamp
    timestamp_str = datetime.datetime.now().strftime("%H%M%S")
    san_histogram_equal_path = f'C:/Users/Madhous/Documents/GitHub/EUD_AM/Asset Management/Visuals - Build Room/Screenshots/san_actions_frequency_histogram_equal_{timestamp_str}.png'
    plt.savefig(san_histogram_equal_path)
    plt.close()  # Close the plot to avoid displaying it inline in this notebook

    print(san_histogram_equal_path)
else:
    print("No data available for 'add' or 'subtract' actions for SANs.")
