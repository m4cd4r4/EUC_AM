<img width="345" alt="python_S3sUP6W9qQ" src="https://github.com/m4cd4r4/EUD_AM/assets/47749761/7dceb424-4359-456f-bf10-23802ebeebac">

# Inventory Management Script

## Key Components of the Script

### Configuration and Setup
- Imports necessary modules.
- Configures logging.
- Sets up the main Tkinter window (root).

### Workbook Initialization
- Checks for the existence of an Excel workbook and initializes it with specific sheets and headers if it doesn't exist.

### GUI Elements
- `SANInputDialog` class: A dialog for SAN number input.
- Various Tkinter widgets for the main application window, like buttons, entry fields, frames, and Treeview for item display.

### Functionality
- `is_san_unique`: Checks if a SAN number is unique.
- `show_san_input`: Displays the SAN input dialog.
- `open_spreadsheet`: Opens the Excel file.
- `update_treeview` and `update_log_view`: Updates the Treeview widgets with data from Excel sheets.
- `log_change`: Logs changes to an Excel sheet.
- `switch_sheets`: Switches between different sheets in the Excel workbook.
- `update_count`: Updates the count of items, includes handling for SAN numbers.

### Event Handling
- Assigns functions to buttons for various operations like adding or subtracting counts, switching sheets, etc.

### Main Loop
- Starts the Tkinter event loop.

## High-Level Flow

### Start and Initialization
- Configure logging.
- Initialize the main window and workbook.

### User Interaction
- Users interact with the GUI to perform various tasks like adding/subtracting item counts, entering SAN numbers, switching between sheets, and opening the spreadsheet.

### Data Processing and Logging
- The script processes user inputs, updates the item counts, checks for SAN uniqueness, and logs changes to the Excel workbook.

### Continuous Update
- The Treeview widgets are continuously updated to reflect changes in the workbook.

---

This script is a comprehensive tool for inventory management, particularly focused on handling items and SAN numbers, with an emphasis on logging and data presentation via a graphical interface.
