# Module-2-VBA-challenge

Contents of repository

1. Screenshoots of each of the 3 sheets in the workbook (Screenshot_Multiple_year_stock_data_2018, Screenshot_Multiple_year_stock_data_2019, Screenshot_Multiple_year_stock_data_2020)
2. A VBA file exported from excel (StockSummaryVBAFile).


Description of code

1. The first loop (i) cycles through the worksheets in the workbook
2. For each sheet there is a loop (j) which cycles through all of the rows of stocks on the sheet. It loops from the second row to the last row which is found by using the command Worksheets(i).Cells(ActiveWorkbook.Worksheets(i).Rows.Count, "A").End(xlUp).Row
3. At the start of the loop a variable (NewStockStartRow = 2) which allocates the start row for the first stock is set.
4. Whilst cycling through the rows, the code compares the present name of the stock with the stock name in the next row (j+1). If they are different, it flags that a new ticker has been found.
5. When a new ticker has been found it write the name of the previous ticker into the first summary table. It also copies the variable NewStockStartRow into OldStockStartRow and allocates the row of the new ticker into NewStockStartRow. This allows us to keep track of where the new ticker started.
6. It is now known where the old ticker started (OldStockStartRow) and where it ended (NewStockStartRow - 1)
7. Calculations are now done on the old ticker to find the yearly change, %yearly change, total stock volume (using the sum function), and the cells and column are formatted.
8. The loop continues to do this for all the stck tickers in the data.
9. Once the program has looped through all the rows and produced a summary table, it produces a second summary table to show the greatest increase, decrease and total volume of stock. It used the Max and Min functions. Once the value is found, the row index is found and used to assign the ticker name to the value.
10. The cells and columns in the sheet are formatted to be more readable.
11. The program then cycles to the next sheet and repeats the process.
