# Wayleave-Template-Automation
- Copying specific columns from all sheets in a workbook to a template and saving each sheet as a new .xlsx file in a specific folder 

## Preparing Blank Grantors Workbook for Automation

1. On exporting the missing grantors, make the following changes to column names
   - Landlord/Wayleave Contact: Full Name > GRANTOR_NAME
   - City: City Name > CITY  
2. Changes can be made to any column name you wish to copy onto the template. Make sure the column you wish to copy onto the template has the same name as in the template.
3. Next, find the missing grantor names and wayleave information as required
4. You can choose to create a Pivot Table for the cities before or after the above step
5. Rename the sheets with the City Name
6. Save File
   

## Wayleave Template Code 

1. On opening wayleave-template.xlsm, you might get a warning on enabling macros, Click 'Enable Macros'
2. Next, Go to the Developer Tab.
   - If you do not see the Developer tab on Excel. Go to File > More > Options > Customise Ribbon and Select the Developer check box.
3. On the Developer tab, Click Visual Basic. This will open the VB Studio
4. Here you will see the required code on file 'Module 1'
5. On the line commented as 'CHANGE 1', Insert the file path to your source file (blank grantors file)
   - To get your file path, head over to the folder where your file is saved and right click on file address then click on 'Copy address as text'
   - At the end of your path, add \filename.xlsx
   - Make sure this is enclosed in quotes
6. The line commented as 'CHANGE 2' directs the program to the folder you wish to save the file in.
   - Here you are only required to change the file path of the required folder
   - Copy file path as instructed above and add a \ at the end
   - If you wish to change the suffix added to these new files, edit " Blanks.xlsx" as required
7. Save the file and run it

## Importing new contacts on Salesforce
1. Add Data to contact-list.csv as required
2. Go to Setup > Search Data Import Wizard on the Quick Find bar
3. Launch Wizard, Select 'Add new records' and drag csv file to upload 
4. If any field is marked as Unmapped, Map it to its respective field
5. Click Next and run the tool.
   
