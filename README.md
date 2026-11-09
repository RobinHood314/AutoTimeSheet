# AutoTimeSheet

This program is intended to make filling out the timesheet a more user-friendly experience, as opposed to manually editing the Excel Spreadsheet

## List of Functions
- Verifies the information on the top 2 rows and allows the user to change it if it's inaccurate
- Automatic date updates (If the Date Column is filled out)
  -Considers leap years! 
- Asks the user for the starting date in case the date column is blank
- Considers the work categories based off the schedule provided in the top 2 rows
- Asks the user to insert hours worked based off the categories
- Prints out the breakdown of the timesheet at the end
- Added more error handling for inserting the date

## Functions to implement
- [x] Handling for the year inserted when the user is asked to insert the start date
- [x] The option to go back when inserting hours worked (To allow for typo correction)*
- [x] Allowing universal compatability between timesheets as there are different versions of the timesheet

      *Will allow the user to go back within the day
