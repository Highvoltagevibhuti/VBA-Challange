# VBA-Challange
Extraction of ticker symbol
-find last raw of ticker column A
-Use for Loop to loop through all rows to find unique ticker and place it in column "I"
- move on to next i when ticker values are different in cells for first column

Extarction of Yearly change and % change
-Check all rows in first colum via for loop
-If next row is different then get the close price
-If previous row is different then get the open price
-yearly change = close price- open price
- % change = (yearly change/ open price)*100

Extraction of total stock volumn
- Sum of all rows from volumn colum for unique ticker

Extarction of greatest volumn along with greatest % increase and decrease
- Use worksheetfunction.max and worksheetfunction.min
- Use For loop on columns "J" and "K" to find associated ticker

Conditional formating
-Use if function to do conditional formating

-define ws to run it on each worksheet and use next ws to loop all tabs/worksheets of excel
