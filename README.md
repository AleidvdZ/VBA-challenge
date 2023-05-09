# VBA-challenge
Notes and resources for VBA project:

Did the coding in the alphabetical_testing file as recommended
  By just doing the math to make sure my numbers were coming out correctly I determined that I needed the following results for the first ticker
    The numbers I am looking for AAB are  -0.17 (yearly change), -0.73 (%change), and the total stock 455750984

Started with assignment 06_credit_charges_solution.xlsm for coding the loops with output to a separate table.

Got LastRow info from 07_census_data… solution file 

To figure out how to subtract to the beginning of the loop I used code from the following online source: How to find the first and last values in a conditional loop in VBA - Stack Overflow: https://stackoverflow.com/questions/59441711/how-to-find-the-first-and-last-values-in-a-conditional-loop-in-vba
  Specifically: assigning the annual open value at the end of the “If” statement and doing the math for yearly change before this
  Realized I had to put in an initial value reference for the annual open value for this to work.

I got a run time error 6 and set the total stock value to double instead of long to fix it. Thanks to this resource: run time error '6' Overflow in visual basic 6.0: https://stackoverflow.com/questions/20855032/run-time-error-6-overflow-in-visual-basic-6-0

For conditional formatting referenced the 03_grader.xlsm file and the color index file shared by TA: http://dmcritchie.mvps.org/excel/colors.htm 
  Red color index is 3
  Green color index is 4

More formatting
  Rounding: Round function (Visual Basic for Applications) | Microsoft Learn https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/round-     function
  Column width: census data solution

Last table - figured it out by myself!  Not sure if it is the most straightforward way to do it but it made sense to me ;)

Formatting % cells with “%” symbol
   https://excelchamps.com/vba/functions/formatpercent/

Result for single sheet in alphabetical testing Sheet A - yay!! - it worked (print screen)

Multiple sheets:
  Used the commands from the Stu_Census_Pt1 file have the commands through each of the worksheets.  It took a bit to figure out “ws” before the cell and range.value and it   makes complete sense so that as the code runs through it sheets it is placed correctly.
  https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit for a reminder on how to autofit the applicable columns.

Submission:
GitHub/GitLab Submission (15 points)

All three of the following are uploaded to GitHub/GitLab:
Screenshots of the results (5 points) - 5 screenshots included: first and last sheet of alphabetical_testing workbook and one screenshot for each year of multiple_year_stock_data workbook

Separate VBA script files (5 points) - I created one VBA script file and copied and pasted it from the alphabetical_testing workbook to multiple_year_stock_data workbook

README file (5 points)
