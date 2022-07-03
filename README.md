# Overview

If you just want to track your own exams, open the `examsLinks.xlsx` and use the markers to update 
things manually.

If you want to understand how to run the script continue reading.

# Setup

Run the `setup.sh` script or do a simpler (install package globally):
```bash
pip3 install openpyxl
```

# Generate new Excel table

1. Make sure all of the exams are located in the relative path `./prevExams/` and that they don't 
have descriptive name (the script will tell you if a file is found with bad name)
2. Update the `examLinksTemplate.xlsx` to contain all of the years that may appear in the 
`prevExams` folder.
3. Make sure to retain the current formal of the template (e.g. don't sort the template itself, the 
first 
row of each 4 will have the year, after each row with a semester name, there's an empty one for 
Moed B, etc...)
4. Run the python script

   ```bash
   python3 addHyperlinksToExcel.py
   ```

   
5. A file called `examsLinks.xlsx` was created (**or overwritten!!**)
