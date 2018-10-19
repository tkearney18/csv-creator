Medicare Benefits Sheet Converted to CSV for Feeds Consumption
v0.1 8/22/2018 Troy Kearney

The plan-csv-creator.py is a Python script to convert the SalesConnect benefits sheet into a csv to be used by Feeds to import the Medicare plans for the tenant.

Before using this script you must have the Excel Workbook with the plan benefits for a particular tenant. You will need to go through each sheet and ensure that a few items are removed/present:
	- The first column must contain the label for the data that is being represented. This should always be there as we need to know how the data is to be mapped. If there is no such column present but are confident that the mapping is the same as the template file then insert one blank column to the front of the sheet.
	- Clean up any erronous data *
		- IE. Notes that are added to a cell, data that is striked out, county listings and images at the bottom of the sheet, etc. If something is questionable, ask.
		* Leave the rows at the top of the sheet. This will be removed by the script.
	- Ensure that the name of the sheet is the corresponding plan type. Med Supp, Med Adv, Med Cost, Med Rx
	- For any plans that say ALL * in the Service Location row, place all of the counties/zipcodes provided in that cell.

Usage:
	- Place the script, MedicareFeedTemplate.xlsx, and the benefits sheet in the same directory
	- Execute script by running python plan-csv-creator.py
	- The following questions will be asked:
		- Enter file to be converted to plan csv:
			- You do not need to add the file extension to the end. The script will look for fileName + '.xlsx'
		- What plan type is to be converted? (Worksheet name)
			- A sheet name that exists in the Workbook
		- How many rows should be skipped?
			- Can be any number. These are the rows at the top that are not used for data (header image, product header, data element row)
		- How many columns should be skipped?
			- Should always have one to skip (label column) but this will be accounted for by the script. Some tenants may have an example plan column.
		- Do you want to delete the Excel file created (y/n)?
			- A workbook with the same data as the csv will exist if needed, else it can be deleted.
	- After the csv is completed validate that all of the data is present and in the correct location.
	- Remove the field_data_zipcode column if it is not needed.
		- Working on adding a question for this item.
	- Upload the csv using the Feeds importer
	

	
TODOS:
	- Update code to Python Coding Standards/clean up the messes.
	- Handle failures with grace.
	- Add fuction to move all counties/zipcodes from an excel file when ALL is used in the cell.
	- Create question for removing the zipcode column if it is not needed.
	- Add a function to remove all the columns with no header.
	- Add file to write out all available sheets.
