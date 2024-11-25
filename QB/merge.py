import xlwings as xw
import os

def merge_excels_with_xlwings(input_folder, output_file):
    # Ensure output_file has a valid .xlsx extension
    if not output_file.endswith('.xlsx'):
        raise ValueError("Output file must have a '.xlsx' extension.")
    
    # Create a new Excel workbook using xlwings
    app = xw.App(visible=False)
    merged_wb = xw.Book()  # New workbook
    
    # Store the reference to the default sheet
    default_sheet = merged_wb.sheets[0]
    
    # Loop through all Excel files in the input folder
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
            file_path = os.path.join(input_folder, file_name)
            
            # Open each source workbook
            source_wb = xw.Book(file_path)
            
            # Copy each sheet from the source workbook into the merged workbook
            for sheet in source_wb.sheets:
                sheet.api.Copy(Before=merged_wb.sheets[0].api)
                # Create a valid sheet name (remove the '.xlsx' extension and limit to 31 characters)
                sheet_name = file_name[:-5][:31]
                merged_wb.sheets[0].name = sheet_name  # Rename the copied sheet
            
            source_wb.close()  # Close the source workbook after copying
    
    # Delete the default sheet only at the end (if it exists)
    if len(merged_wb.sheets) > 1:
        default_sheet.delete()  # Delete only if other sheets are present

    # Save the merged workbook
    merged_wb.save(output_file)
    merged_wb.close()
    app.quit()  # Close the Excel app

# Usage example
input_folder = r'C:\\Users\\jawad.khan\\Desktop'
output_file = r'C:\\Users\\jawad.khan\\Desktop\\Work\\Automation\\QB\\final_report\\merged_file.xlsx'  # Ensure this ends with .xlsx
merge_excels_with_xlwings(input_folder, output_file)
