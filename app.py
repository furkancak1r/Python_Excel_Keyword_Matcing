import os
import win32com.client
import json

# Specify the file path
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Python_Excel_Data_Matching_Copy\\old.xlsx"

# Load keywords from keywords.json
with open('keywords.json', 'r', encoding="utf-8") as file:
    data = json.load(file)
    keywords = data["keywords"]

# Check if the file exists
if os.path.exists(file_path):
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.Sheets("Sayfa1")
        excel.Visible = False

        # Get the last row of column B
        last_row = worksheet.Cells(worksheet.Rows.Count, "B").End(-4162).Row

        # Create a batch update object
        batch_update = []

        # Iterate through rows
        for i in range(1, last_row + 1):
            # Get the value of the cell in column B
            cell_value = worksheet.Cells(i, "B").Value

            # Check if the cell value is not None
            if cell_value is not None:
                # Convert to lowercase
                cell_value_lower = str(cell_value).lower()

                matching_keyword = None

                # Iterate through keywords
                for keyword in keywords:
                    keyword_lower = keyword.lower()

                    # Check if the keyword is in the cell value
                    if keyword_lower in cell_value_lower:
                        matching_keyword = keyword
                        break

                # If no match in the entire cell value, check between the first two '-'
                if matching_keyword is None:
                    # Split the cell value by '-'
                    cell_parts = cell_value.split('-')

                    # Check if there are at least two parts separated by '-'
                    if len(cell_parts) > 1:
                        # Get the part between the first two '-'
                        between_dashes = cell_parts[1].strip()

                        # Convert to lowercase
                        between_dashes_lower = between_dashes.lower()

                        # Iterate through keywords again
                        for keyword in keywords:
                            keyword_lower = keyword.lower()

                            # Check if the keyword is in the part between the first two '-'
                            if keyword_lower in between_dashes_lower:
                                matching_keyword = keyword
                                break

                        # If matching keyword found between dashes, break the loop
                        if matching_keyword:
                            break

                # If no match found, set "Tanımlanmayan Arıza"
                if not matching_keyword:
                    batch_update.append((i, "Tanımlanmayan Arıza"))
                else:
                    batch_update.append((i, matching_keyword))

            else:
                # If cell is empty, set "Bilinmeyen Arıza"
                batch_update.append((i, "Bilinmeyen Arıza"))

        # Perform the batch update
        for row, value in batch_update:
            worksheet.Cells(row, "A").Value = value

        # Save and close the workbook
        workbook.Save()
        workbook.Close()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the Excel application
        excel.Quit()
