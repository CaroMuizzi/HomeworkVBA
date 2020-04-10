The VBA of Wall Street

We will use VBA scripting to analyze real stock market data. The files that we will be using for are:

* [Test Data](Resources/alphabtical_testing.xlsx) - Used for developing the scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) -  Used for Run the scripts and to generate the final report.

Objective:

* Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.

* Will also need to display the ticker symbol to coincide with the total volume.

* Create a script that will loop through all the stocks and take the following info.

  * Yearly change from what the stock opened the year at to what the closing price was.

  * The percent change from the what it opened the year at to what it closed.

  * The total Volume of the stock

  * Ticker symbol

* Should also have conditional formatting that will highlight positive change in green and negative change in red.

* Be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

* Make the appropriate adjustments to your script that will allow it to run on every worksheet just by running it once.


Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing the code. This dataset is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.


