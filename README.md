# AutoTimeSheet

This program is intended to make filling out the timesheet a more user-friendly experience, as opposed to manually editing the Excel Spreadsheet

## Instructions
1. Delete any version of this program that is 1.3.5 or older, they do NOT work as intended (I should REALLY get testers)
2. Fork this repository!
   `git clone https://github.com/RobinHood314/AutoTimeSheet`

3. Copy your timesheet breakdown in the same folder as this program and name it TimeSheet.xlsx
4. Have fun!

### Important Side Note...
Sooooo AutoTimesheet versions 1.3.5 and older are ALL broken, they will NOT work as intended.
Here's why:
   1. The program will pass the user given values as strings as opposed to floats, which is problematic when the values are
      mostly numbers, which prevents most functionality of the timesheet, like adding hours in the pay period
   2. There is a line in the code (data_only = True) which essentially purges the formulas in the timesheet, which prevents any future use of the Timesheet
   3. I added a new function which allows the addition of hours in the day and in the entire pay period since I had to remove the "data_only = True" line
That being said, please download this version as this one fixed all the bugs I explained above

## List of Functions
- Verifies the information on the top 2 rows and allows the user to change it if it's inaccurate
- Automatic date updates (If the Date Column is filled out)
  - Considers leap years! 
- Asks the user for the starting date in case the date column is blank
- Prints the date range for the timesheet
- Considers the work categories based off the schedule provided in the top 2 rows
- Asks the user to insert hours worked based off the categories
  - Allows the user to go back to allow for typo correction (within the day)
- Prints out the breakdown of the timesheet at the end (now it looks nicer)

## Functions to implement
- No new functions to implement! (yet)
