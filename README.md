
# VBA on Ticker Tracker
To recreate my results, go click [This Excel File](Resources/Multiple_year_stock_data.xlsx) to download the stock data and then paste the text from my [VBA Script](VBAscript.bas) into a module within VBA. Finally click the developer tab and run the macros.

This VBA script loops through all the stocks for one year and output the following information before repeating this process on the next years (worksheets) data:

* The ticker symbol.
* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.
* The stocks "Greatest % increase", "Greatest % decrease" and "Greatest total volume".


### Tech Stack
* Miscrosoft Excel and VBA 
