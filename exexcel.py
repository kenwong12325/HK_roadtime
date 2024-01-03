import os
import win32com.client as win32

def autovba():
    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # Construct the full file path
    xlsm_file_path = os.path.join(current_dir, "traffic.xlsm")

    # Open the xlsm file
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = False  # Run Excel in the background without displaying the window

    try:
        workbook = excel_app.Workbooks.Open(xlsm_file_path)

        # Perform operations on the workbook or sheets here
        
        #Save and close the workbook
        workbook.Save()
        workbook.Close(SaveChanges=False)  # Don't save changes made to the workbook

        # Quit Excel application
        excel_app.Quit()
    except Exception as e:
        print(f"Error: {e}")