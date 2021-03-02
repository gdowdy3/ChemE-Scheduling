# ChemE-Scheduling
Code for scheduling visitor-professor meetings during the MIT ChemE Recruiting Weekends

## Prerequisites
1. `Python`
2. `pip` (can be installed along with python)
3. `ortools` (an optimization package that can be installed with `pip install ortools`)
4. `pandas` (can be installed with `pip install pandas`)


## How to use this code
1. Download the `GenerateSchedule.py` and `Input Data.xlsx` files into the same directory on your local machine.  In the instructions that follow, we will refer to this as the `Working Directory`.  It doesn't matter which directory you pick, so long as you know where it is.
2. Open `Input Data.xlsx` and go to the `Visitor Preferences` sheet.  Fill in the visitor information, following the format of the existing data and replacing the existing data.
3. Go to the `Professor Availability` sheet.  Indicate when each professor is available for meetings by placing a `1` in the appropriate column.
4. Save and close the workbook.
5. Open a terminal and navigate to your `Working Directory`.
6. Enter the following command: `python GenerateSchedule.py`.
7. The outputs will be created in your `Working Directory`.

## Questions
Create an "Issue" on this GitHub repository if you have any problems/questions.
