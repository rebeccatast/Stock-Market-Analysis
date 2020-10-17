# VBA-StockMarket-Data
For this project I created a VBA script in Excel to analyze stock market data from 2014, 2015, and 2016. The script was created using  the Visual Basic editor in Microsoft 365 Excel for Mac, version 16.42. 

The script, titled "Stock Challenge Script" can be run in Excel using the developer tab. Once in the developer tab, open the Visual Basic editor. You can copy the script from the file and paste into the editing window, and then run the script using the run script (play button) in the toolbar. This script will run for all sheets in the workbook, so there is no need to do each sheet individually. 

The script will loop through all stocks for each year, for all sheets in the workbook, providing an output of the following information: 
   -  Ticker Symbols
   -  Overall price change between the opening and closing price of each stock for a given year. 
   -  Overall percentage change between the opening and closing price of each stock for a given year. 
   -  Total stock volume
    
The script also uses conditional formatting to distinguish between stocks that had an overall positive yearly price change (in green) and stocks that had an overall negative yearly price change (in red). 

Finally, the script provides the ticker symbol for the stock with greatest percent increase and the value associated with that percent, the greatest percent decrease and the value associated with that percent, and the stock with the greatest total volume and the value associated with that volume. 
