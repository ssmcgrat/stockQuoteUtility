# stockQuoteUtility

This is a utility application I wrote for my father, who was frustrated with how Google finance had changed its UI, no longer allowing him to easily export the current stock or mutual fund prices of his portfolio to Excel. As he would say, "It's totally lame."

This application takes as input stocks.xlsx, reading each stock symbol from the A column and outputting the current price of the stock in the B column.

# Requirements

Java 8 installed on your local machine.
Windows 7+

# Usage

For an end user, there are three files of interest

    1. GetStockQuotes.bat
    
    2. PatsStocks.jar
    
    3. stocks.xlsx
    
    
Simply download these three files to a common directory on your machine. Open stocks.xlsx, add your stock symbols of interest in the A column, beginning in cell A1. Do not leave any empty cells between stock symbols. Close stocks.xlsx. Double click GetStockQuotes.bat. This will open a windows command terminal, executing the runnable jar file. Once the program is complete, hit any key to close the window. Open stocks.xlsx and your stock prices will be listed.

# Notes

It is possible that the third part server we query to retrieve quotes has issues. In the console output, if your see any errors related to servers (i.e. 503) try running the program again, this is most likely a network issue.

Should something unexpected go wrong, stocks.xlsx could become corrupted. We recommend copy/pasting the results in stocks.xlsx into another excel file that you use to keep track of your portfolio information.
