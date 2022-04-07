#! python3
# rename_files.py - Renames the downloaded spreadsheets from Odoo to 'raw_data', 
# 'breached' and 'last24' based on the time they were created (downloaded)

import os                   
from pathlib import Path    

# Get the current working directory. Save it to a variable.
cwd = Path.cwd()

# Use the glob method from the cwd object to get all the files named helpdesk.
# Put them on a list.
files = list(cwd.glob('helpdesk*.xlsx'))

# Sort the previous list: as the sorting rule, use the getctime method
# to order the list by creation time of the file
files.sort(key=os.path.getctime)

# Put the new file names on a list, in order.
filenames = [
        'raw_data.xlsx',
        'last24.xlsx',
        'breached.xlsx']

# Loop through our helpdesk files list
for i in range(0, len(files)):
    # If the file has been already renamed, show a message
    if os.path.isfile(cwd / filenames[i]):
        print("The file already exists")
    else:
        # Use the rename method of the os library. For making the new name, 
        # just grab the current working directory and concatenate it to the current
        # value on the filenames list.
        os.rename(files[i], cwd / filenames[i])









