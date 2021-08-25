# VBA-challenge
The VBA of Wall Street. This script will automatically analyze real stock market data from a data set.

Data set **Multiple_year_stock_data**

First, it will get the ticker's name and initial open price. It will loop throughout the rows until the ticker's name changes.
If the ticker's name is different, the for loop will stop and calculate the Yearly change by substracting the Open price to the last Close price.
If Yearly Change is a positive number, it will fill the cell with green, but if the value is negative, the cell will be filled with red.

Then it will calculate the Percent Change by (Yearly Change / Open Price) * 100.
Lastly, it will calculate the total stock volume by adding all the volumes of the ticker.

### Bonus Part included:

* Get Greatest % Increase with ticket name and value
* Get Greatest % Decrease with ticket name and value
* Get Greatest Total Volume with ticket name and value
