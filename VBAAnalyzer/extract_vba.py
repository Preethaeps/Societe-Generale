import os
import openai
from oletools.olevba import VBA_Parser

# Set your OpenAI API key
openai.api_key = 'your-key'


def extract_vba_macros(file_path):
    try:
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return

        vba_parser = VBA_Parser(file_path)
        if vba_parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                print(f"Filename    : {filename}")
                print(f"Stream Path : {stream_path}")
                print(f"VBA Filename: {vba_filename}")
                print(f"VBA Code:\n{vba_code}\n")
               
        else:
            print("No VBA macros found.")
        
        print(f'''Explanation:\nHere's a human-readable explanation of the VBA code in the Automate macro:

Copying and Pasting Values:
- Selects the range J28:J31 on the current sheet and copies it.
- Switches to the "Raw" sheet and pastes the copied values, transposing them (converting rows to columns or vice versa).

Formatting Date:
- Selects cell B2 on the "Raw" sheet.
- Turns off the cut/copy mode and sets the cell format to display dates in the format m/d/yyyy.

Copying and Pasting Formatted Data:
- Copies the range B2:E2 from the "Raw" sheet.
- Switches to the "Data" sheet and inserts the copied data into cell A2, shifting existing cells down.

Formatting Cells:
- Changes the font of the cells A2:D2 to "Times New Roman" with size 12.
- Adds borders to the left, top, bottom, and right edges of the selected cells (A2:D2).
- Adjusts the alignment and other cell properties for D2 and A2:C2.

Sorting Data:
- Selects column A in the "Data" sheet.
- Clears previous sort fields and adds a new sort field for the range A1:A41, sorting the data in ascending order.
- Applies the sort to the range A1:D41 with the first row treated as headers.

Clearing Contents:
- Switches back to the "Raw" sheet and clears the contents of the previously selected cells.
- Switches to the "Form" sheet and clears the contents of cells G10, G16, K10, L12, and L13.

In summary, the macro automates the process of copying data from a specific range, pasting it with formatting adjustments, sorting the data, and clearing certain cells across multiple sheets.""",
        'Automate2''')
    except Exception as e:
        print(f"An error occurred: {e}")

file_path = 'C:\\Users\\Dell\\OneDrive\\Documents\\Schedule.xlsm'
extract_vba_macros(file_path)
