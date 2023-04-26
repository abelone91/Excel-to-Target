import openpyxl
import time
import win32clipboard
import win32con
import win32gui
import win32api

# Prompt user to select an Excel file
while True:
    try:
        file_path = input("Enter the file path of the Excel file: ")
        workbook = openpyxl.load_workbook(file_path)
        break
    except:
        print("Invalid file path. Please try again.")

# Prompt user to select a worksheet
while True:
    try:
        worksheet_name = input("Enter the name of the worksheet: ")
        worksheet = workbook[worksheet_name]
        break
    except:
        print("Invalid worksheet name. Please try again.")

# Prompt user to select a column to copy
while True:
    try:
        column_letter = input("Enter the column letter to copy from (e.g. A): ")
        column_values = []
        for cell in worksheet[column_letter]:
            column_values.append(cell.value)
        break
    except:
        print("Invalid column letter. Please try again.")

# Prompt user to select the number of cells to paste
while True:
    try:
        num_cells = int(input("Enter the number of cells to paste: "))
        break
    except:
        print("Invalid number of cells. Please try again.")

# Prompt user to specify the number of seconds to wait before the first paste
while True:
    try:
        wait_first = int(input("Enter the number of seconds to wait before the first paste: "))
        break
    except:
        print("Invalid input. Please enter an integer.")

# Prompt user to specify the number of seconds to wait between pastes
while True:
    try:
        wait_between = int(input("Enter the number of seconds to wait between pastes: "))
        break
    except:
        print("Invalid input. Please enter an integer.")

# Wait for the specified number of seconds before starting to paste the first cell
time.sleep(wait_first)

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

    # Wait for the specified number of seconds before pasting the next value
    time.sleep(wait_between)

    # Update the previous value
    previous_value = str(value)

print("Done pasting values.")
