# Excel Data Retrieval Scripts

## Script 1(row): Retrieve Row Data by ID from Excel Sheet

This script utilizes the `openpyxl` library to load an Excel file, allowing users to retrieve data based on a specified ID. The user is prompted to input an ID, and the script searches for the corresponding row containing that ID in a designated column. If found, it prints the entire row. The program continues to prompt for IDs until the user opts to exit.

### Description of Functionality:
1. **Loading Excel Data**: The script loads data from an Excel file using `openpyxl`.
2. **User Input**: The user is prompted to input an ID or 'exit'/'q' to terminate the program.
3. **ID Lookup**: It searches for the entered ID in a specific column of the Excel sheet.
4. **Display Results**: If the ID is found, it prints the entire row corresponding to that ID.
5. **Loop**: The script continues to prompt for IDs until the user chooses to exit.

## Script 2(column): Retrieve Column Data by ID from Excel Sheet

This script also utilizes the `openpyxl` library to load an Excel file. It allows users to retrieve data based on a specified ID, but in this case, it fetches the values from a designated column instead of a row. If the provided ID is found within the specified column, it prints the corresponding values found in subsequent rows. Similar to the first script, the program continues to prompt for IDs until the user opts to exit.

### Description of Functionality:
1. **Loading Excel Data**: The script loads data from an Excel file using `openpyxl`.
2. **User Input**: The user is prompted to input an ID or 'exit'/'q' to terminate the program.
3. **ID Conversion**: The entered ID is attempted to be converted into an integer.
4. **ID Lookup**: It searches for the entered ID within a specific row (assuming row 5 in this case).
5. **Display Results**: If the ID is found, it prints the corresponding values found in the column.
6. **Loop**: The script continues to prompt for IDs until the user chooses to exit.

These scripts provide efficient means to interact with Excel data, allowing users to retrieve specific information based on provided IDs while maintaining a user-friendly interface.
