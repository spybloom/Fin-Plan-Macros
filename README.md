# Fin-Plan-Macros
These are VBA macros I developed for my job at Financial Planning. 

From TD Ameritrade's website, all of the client's accounts are exported to .csv files and AllTDSheets is run to process them, sorting the sheets and copying the values over to the client's portfolio worksheet. After this, Dist is run and calls Grid_Sort to sort the equity grid and create a line on the client's "third sheet" to show their overall performance over time.
AddStock is used to add a stock's name, ticker and grid location to a master list of stocks our clients own, so they aren't added in the "Morningstar" sub of AllTDSheets.
NewAccount creates a template for adding a new account to the client's portfolio page.
