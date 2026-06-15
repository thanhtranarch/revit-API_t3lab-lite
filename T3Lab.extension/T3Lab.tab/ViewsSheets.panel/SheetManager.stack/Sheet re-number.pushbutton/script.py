# Import necessary libraries
from pyrevit import revit, DB, forms

# Get the current document
doc = revit.doc

# Collect all ViewSheet elements
sheets = list(DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements())

# Check if there are any sheets
if not sheets:
    forms.alert("No sheets found in the document.")
    raise SystemExit

# Prepare a list for displaying sheets
sheet_display_list = ["{} - {}".format(sheet.SheetNumber, sheet.Name) for sheet in sheets]

# Prompt user to select sheets to renumber
selected_sheet_numbers = forms.SelectFromList.show(
    sheet_display_list,
    title='Select Sheets to Renumber',
    multiselect=True,
    show_images=False
)

# Check if any sheets were selected
if not selected_sheet_numbers:
    forms.alert("No sheets selected. Exiting script.")
    raise SystemExit

# Ask for starting number
starting_number_input = forms.ask_for_string(
    prompt='Enter the starting number for renumbering:',
    title='Starting Number',
    default='1'
)

# Ask for prefix
prefix_input = forms.ask_for_string(
    prompt='Enter a prefix for the sheet numbers (leave blank for none):',
    title='Sheet Prefix',
    default=''
)

# Convert input to integer
try:
    starting_number = int(starting_number_input)
except ValueError:
    forms.alert("Invalid starting number. Please enter a valid integer.")
    raise SystemExit

# Start a transaction to modify the sheets
with revit.Transaction('Renumber Sheets'):
    for index, sheet_display in enumerate(selected_sheet_numbers):
        # Extract the sheet number from the display string
        sheet_number = sheet_display.split(" - ")[0]
        sheet = next(sheet for sheet in sheets if "{} - {}".format(sheet.SheetNumber, sheet.Name) == sheet_display)

        # Create the new sheet number with the optional prefix
        new_sheet_number = "{}{}".format(prefix_input, starting_number + index)

        # Try to assign the new sheet number
        try:
            sheet.SheetNumber = new_sheet_number
            print("Renumbered sheet '{}' to '{}'".format(sheet.Name, new_sheet_number))
        except Exception as e:
            # Skip any errors and continue with the next sheet
            print("Failed to renumber sheet '{}': {}".format(sheet.Name, str(e)))

print("Sheet renumbering completed.")