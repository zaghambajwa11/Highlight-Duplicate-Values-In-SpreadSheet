# Highlight-Duplicate-Values-In-SpreadSheet
This code is a Python function that takes an Excel file and a column name as input, finds duplicate values in the specified column, and returns their addresses.

Here's what the function does step-by-step:

The function first reads the Excel file into a Pandas dataframe using the pd.read_excel() function.

It then identifies the duplicate values in the specified column using the duplicated() function in Pandas. The keep=False argument is used to mark all duplicates as True.

The function converts the column to a string using the astype() function in Pandas. This is done to avoid issues with formatting when the duplicates are highlighted later.

The function creates an empty dictionary colors and a Pandas Styler object style. The colors dictionary will be used to store the unique colors assigned to each duplicate value. The style object will be used to highlight the duplicate values.

The function iterates over the column_name column of the dataframe using the apply() function in Pandas. For each value, it checks if the value is already in the colors dictionary. If it is, it uses the corresponding color value. If it is not, it generates a random color value using the random.randint() function, adds the value-color pair to the colors dictionary, and uses the new color value.

The function creates an Excel writer object using the pd.ExcelWriter() function and loads the original Excel file using openpyxl.load_workbook(). The writer object is used to create a new sheet named 'Duplicates' in the workbook.

The function applies the background color to the style object using the apply() function. For each value in the specified column, it checks if the value is already in the colors dictionary. If it is, it applies the corresponding color to the background of that cell.

The function saves the style object to the 'Duplicates' sheet using the to_excel() function.

The function creates an empty list addresses and iterates over the duplicates series using the items() function. For each row where duplicates is True, it gets the cell address of that row using openpyxl.utils.get_column_letter() and adds the cell address and the previous occurrence of the duplicate value to the addresses list.

Finally, the function returns the addresses list.

The example usage at the end of the code demonstrates how to call the function with an Excel file named data.xlsx and a column named Driver. If there are any duplicate values in the Driver column, the function prints their addresses. If there are no duplicates, it prints a message indicating that there are no duplicates.
