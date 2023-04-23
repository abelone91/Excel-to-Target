import openpyxl
import time
import win32clipboard
import win32con
import win32gui
import win32api

# Prompt user to select an Excel file
file_path = input("Enter the file path of the Excel file: ")

# Load the workbook
workbook = openpyxl.load_workbook(file_path)

# Prompt user to select a worksheet
worksheet_name = input("Enter the name of the worksheet: ")
worksheet = workbook[worksheet_name]

# Prompt user to select a column to copy
column_letter = input("Enter the column letter to copy from (e.g. A): ")
column_values = []
for cell in worksheet[column_letter]:
    column_values.append(cell.value)

# Prompt user to select the number of cells to paste
num_cells = int(input("Enter the number of cells to paste: "))

# Wait for 2 seconds before starting to paste the first cell
time.sleep(2)

# Initialize the previous value to an empty string
previous_value = ""

# Loop through the column values and paste them sequentially
for i in range(num_cells):
    value = column_values[i]

    # Check if the current value is the same as the previous value
    if str(value) == previous_value:
        continue

    # Copy the value to the clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(str(value))
    win32clipboard.CloseClipboard()

    # Set the active window as the foreground window
    win32gui.SetForegroundWindow(win32gui.GetForegroundWindow())

    # Send the paste command to the active window
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0) # Press the Control key
    win32api.keybd_event(ord('V'), 0, 0, 0) # Press the V key
    win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0) # Release the V key
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0) # Release the Control key

    # Wait for the paste operation to complete
    time.sleep(2)

    # Update the previous value
    previous_value = str(value)

print("Done pasting values.")
