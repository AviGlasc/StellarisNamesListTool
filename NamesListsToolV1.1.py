# Avi's Stellaris Names List Tool V1.1; 2/3/2025
# This tool is designed to help Stellaris modders create names lists for their mods.
# This tool is built around the google sheets template available here: https://docs.google.com/spreadsheets/d/1GpY67q_PQ5TbjV4x9hqWE4KWJxhgVF0coUvGhxl-C4o/edit?usp=sharing

# To get started, copy the above file to your own Google Drive and set it to public access.
# The first page of the document is the tutorial page, which explains how to use the template, and any subsequent pages are the actual name lists themselves.
# Send the document to all of your collaborators and have them fill in the name lists as needed.
# Each new name list should be on a new page, and the name of the page will be the name of the list in the game, page name is also automatically updated from the 'NamesList Title' cell.
# Once all of the name lists are filled in, the tool user needs to go through each sheet and set the 'Enable List Processing:' ell to TRUE (checked), without this checked the page will be skipped.
# These will be put into a folder called "ASNLT_output" in the same directory as the tool. Make sure to take the TXT files out of this folder before placing them in the name_lists folder.

# Currently, the tool only supports the google sheets format. Future versions may support other formats.
# There is no automated support for sequential names at this time, but this will be added in the future, as well as added manually with a little bit of research (if you're interested, just look at the localization files in the game and compare them with the nameslist).
# The tool requires the pandas library to be installed. This can be done by running "pip install pandas" in the command line.

# If you notice issues or have suggestions for the tool, please let me know via my email: avilynnglasc@gmail.com
# If you're more savvy, feel free to fork the repository and make your own changes. I'd love to see what you come up with!
# For those wanting to learn about the Stellaris side of things, check out the Stellaris wiki page: https://stellaris.paradoxwikis.com/Empire_modding#Name_lists or the README file in the Stellaris/common/name_lists folder.

# Importing the required libraries
import pandas as pd
import re
import string
import openpyxl
import requests
from string import Template 
import os

NONALLOWEDCHARS = ['$', '%', '/', '{', '}', ',']
REPLACEMENTCHARS = ['S', 'P', ' ', '', '', ' ']

#df.iloc uses [row, column] 0 indexed, very first row is reserved for df.columns entries, so "row 0" in code is actually row 2 in the sheet

def check_and_replace(array_of_strings, non_allowed_chars, replacement_chars):
    # Create a dictionary for quick lookup of replacements
    replacement_dict = dict(zip(non_allowed_chars, replacement_chars))
    
    # Process each string in the array
    cleaned_array = []
    for s in array_of_strings:
        s = str(s)  # Ensure the entry is treated as a string, allowing numbers in names
        if pd.isna(s) or s == 'nan':
            continue  # Skip "nan" values
        for char in non_allowed_chars:
            if char in s:
                s = s.replace(char, replacement_dict[char])
        cleaned_array.append(s)
    
    return cleaned_array

def convert_google_sheet_url(url):
    # Regular expression to match and capture the necessary part of the URL
    pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)(/edit#gid=(\d+)|/edit.*)?'

    # Replace function to construct the new URL for CSV export
    # If gid is present in the URL, it includes it in the export URL, otherwise, it's omitted
    replacement = lambda m: f'https://docs.google.com/spreadsheets/d/{m.group(1)}/export?' + (f'gid={m.group(3)}&' if m.group(3) else '') + 'format=xlsx'

    # Replace using regex
    new_url = re.sub(pattern, replacement, url)

    return new_url

def concatenate_column_data(df, columnGrab, startValue=3):
    # Grabs column data from column =columnGrab from a data frame =df starting at the row=startValue and continuing downward

    # Grab the column data from the dataframe and convert to array
    columnData = df.iloc[startValue:, columnGrab]
    columnList = columnData.to_numpy()
    
    # Check for and replace any non-allowed characters in the strings
    cleaned_columnList = check_and_replace(columnList, NONALLOWEDCHARS, REPLACEMENTCHARS)

    # Concatenate all strings together, separated by spaces and include quotations around each
    concatenatedString = ' '.join([f'"{s}"' for s in cleaned_columnList])

    return concatenatedString

def col2num(col):
    # Converts a column letter(s) to a number, since this is plugged into the pandas dataframe, which is zero-indexed, we subtract 1 from the result
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1           
    return num - 1

def process_sheets(sheets):
    for sheet_name, df in sheets.items():
        if sheet_name == "Tutorial":
            continue # Skip the tutorial sheet
        elif df.iloc[(14), col2num('L')] == False: # The 16th row, L column is the "Enable List Processing:" row, if this is set to FALSE, the sheet will be skipped
            print(f"Skipping sheet: {sheet_name}, update the 'Enable List Processing:' cell to TRUE (checked) to process this sheet.")
            continue
        else:
            print(f"Processing sheet: {sheet_name}")
        
            # Create a template object from the template.txt file
            with open(r'StellarisNamesListTool\template.txt', 'r') as file:
                template_string = file.read()
            template = string.Template(template_string)

            # Create a dictionary of values to fill into the template
            values = {
                "listname": sheet_name, # The name of the list is the name of the sheet, this is what is used in the game, so it should be unique
                "categoryname": df.iloc[7, col2num('L')], # The category name is the type of names list (toxoids, humanoids, etc.), I'm not sure where this is used in the game, but it's here for completeness
                "shipgeneral": concatenate_column_data(df, col2num('AU')),
                "shipcorvette": concatenate_column_data(df, col2num('BC')),
                "shipdestroyer": concatenate_column_data(df, col2num('BD')),
                "shipcruiser": concatenate_column_data(df, col2num('BE')),
                "shipbattleship": concatenate_column_data(df, col2num('BF')),
                "shiptitan": concatenate_column_data(df, col2num('BG')),
                "shipcolossus": concatenate_column_data(df, col2num('BH')),
                "shipjuggernaut": concatenate_column_data(df, col2num('BI')),
                "shipscience": concatenate_column_data(df, col2num('AY')),
                "shipcolonizer": concatenate_column_data(df, col2num('AZ')),
                "shipconstructor": concatenate_column_data(df, col2num('AX')),
                "shiptransport": concatenate_column_data(df, col2num('BL')),
                "shipstarbase": concatenate_column_data(df, col2num('BN')),
                "shipioncannon": concatenate_column_data(df, col2num('BM')),
                "fleetgeneral": concatenate_column_data(df, col2num('BS')),
                "armygeneral": concatenate_column_data(df, col2num('BV')),
                "armydefense": concatenate_column_data(df, col2num('BW')),
                "armyassault": concatenate_column_data(df, col2num('BX')),
                "armyslave": concatenate_column_data(df, col2num('BY')),
                "armyundead": concatenate_column_data(df, col2num('BZ')),
                "armyclone": concatenate_column_data(df, col2num('CA')),
                "armymachinedefence": concatenate_column_data(df, col2num('CB')),
                "armyrobotic": concatenate_column_data(df, col2num('CC')),
                "armyroboticdefense": concatenate_column_data(df, col2num('CD')),
                "armypsionic": concatenate_column_data(df, col2num('CE')),
                "armyxenomorph": concatenate_column_data(df, col2num('CF')),
                "armygenewarrior": concatenate_column_data(df, col2num('CG')),
                "armyoccupation": concatenate_column_data(df, col2num('CH')),
                "armyindividualmachineoccupation": concatenate_column_data(df, col2num('CI')),
                "armyroboticoccupation": concatenate_column_data(df, col2num('CJ')),
                "armyprimitive": concatenate_column_data(df, col2num('CK')),
                "armyindustrial": concatenate_column_data(df, col2num('CL')),
                "armypostatomic": concatenate_column_data(df, col2num('CM')),
                "armymachineassault1": concatenate_column_data(df, col2num('CN')),
                "armymachineassault2": concatenate_column_data(df, col2num('CO')),
                "armymachineassault3": concatenate_column_data(df, col2num('CP')),
                "armywarpling": concatenate_column_data(df, col2num('CQ')),
                "planetgeneral": concatenate_column_data(df, col2num('CU')),
                "planetdesert": concatenate_column_data(df, col2num('CV')),
                "planettropical": concatenate_column_data(df, col2num('DA')),
                "planetarid": concatenate_column_data(df, col2num('CW')),
                "planetcontinental": concatenate_column_data(df, col2num('CZ')),
                "planetocean": concatenate_column_data(df, col2num('CY')),
                "planettundra": concatenate_column_data(df, col2num('DD')),
                "planetarctic": concatenate_column_data(df, col2num('DB')),
                "planetsavannah": concatenate_column_data(df, col2num('CX')),
                "planetalpine": concatenate_column_data(df, col2num('DC')),
                "characterfullgeneral": concatenate_column_data(df, col2num('T')),
                "characterfullfemale": concatenate_column_data(df, col2num('U')),
                "characterfullmale": concatenate_column_data(df, col2num('V')),
                "characterfirstgeneral": concatenate_column_data(df, col2num('X')),
                "characterfirstfemale": concatenate_column_data(df, col2num('Y')),
                "characterfirstmale": concatenate_column_data(df, col2num('Z')),
                "charactersecondgeneral": concatenate_column_data(df, col2num('AB')),
                "charactersecondfemale": concatenate_column_data(df, col2num('AC')),
                "charactersecondmale": concatenate_column_data(df, col2num('AD')),
                "characterregnalfullgeneral": concatenate_column_data(df, col2num('AG')),
                "characterregnalfullfemale": concatenate_column_data(df, col2num('AH')),
                "characterregnalfullmale": concatenate_column_data(df, col2num('AI')),
                "characterregnalfirstgeneral": concatenate_column_data(df, col2num('AK')),
                "characterregnalfirstfemale": concatenate_column_data(df, col2num('AL')),
                "characterregnalfirstmale": concatenate_column_data(df, col2num('AM')),
                "characterregnalsecondgeneral": concatenate_column_data(df, col2num('AO')),
                "characterregnalsecondfemale": concatenate_column_data(df, col2num('AP')),
                "characterregnalsecondmale": concatenate_column_data(df, col2num('AQ')),
            }

            finalNamesList = template.substitute(values)  
            #print(finalNamesList)
            with open(f'StellarisNamesListTool/ASNLT_output/{sheet_name}.txt', 'w') as file:
                file.write(finalNamesList)
        
    

# Create the directory new names lists will be saved to
directory_name = "ASNLT_output"
try:
    os.mkdir('StellarisNamesListTool/' + directory_name)
    print(f"Directory '{directory_name}' created successfully.")
except FileExistsError:
    print(f"Directory '{directory_name}' already exists, continuing.")

# Get URLs from user input and split them into a list for processing
inputString = input("Enter the URL of the Google Sheet, multiple sheets can be given seperated by commas, multiple tabs in the same sheet will be processed: ")
urls = inputString.split(',')

# Loop through the URLs provided by user
for url in urls:
    print(f"Processing document: {url}")

    # Convert the Google Sheet URL to a CSV export URL
    new_url = convert_google_sheet_url(url.strip())
    sheets = pd.read_excel(new_url, sheet_name=None)
    process_sheets(sheets)


