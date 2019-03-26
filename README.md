# Automated_Returns_Reporting
I made this program to automate stock return reporting for a wealth management firm in NYC

This program can can get historical stock price data, dividend data, and earnings release dates from Yahoo Finance and Finviz (respectively). 
The instructions for how to use it are as follows:

Instructions



1. Run DS1.exe (For windows dist → DS1_Windows → DS1_Windows.exe ||  For Mac dist → DS1 → DS1.exe)
2. Choose or create your google sheet that you want it to write to
3. Share this google sheet with: yahoofinance@voya-181121.iam.gserviceaccount.com

4. Format your google sheet to retrieve the information you need:
    - Write "Tickers" as the column header above all your tickers, make sure this is the first column in your sheer
    - For the other columns, choose which data you want by writing the following functions:

          - 'Earnings Release Dates' ----> Will give you the next release date for each stock
          - 'Historical Price (YYYYMMDD)' ----> Will give you the Adj. Close of each stock on this day
          - 'Historical Price (yesterday)' ----> Will give you the Adj. Close of the previous trading day. (may not work for 8 would-be trading days that are holidays)
          - 'DivX (YYYYMMDD:YYYYMMDD)' ----> Gives you the sum of all the dividends between the two dates (First date exclusive, last date inclusive)

     - Note: These column headers have to match the text above exactly in order for them to  work. If you want to retrieve data for one column and not the others, simply change the header of the other columns.

5. Answer the prompt in the terminal window using the name of the google sheet you want to write to, make sure to use correct capitalization and no space at the beginning. Also make sure that it is shared with: yahoofinance@voya-181121.iam.gserviceaccount.com

6. Wait for the program to fill out all the data (works for multiple sheets), this might take a while      (3+ minutes)
Note: Sometimes the program will crash and give the error message ‘No such file in directory’, this is usually due to internet connectivity issues, just run the program again.

7. If you run into any problems please let me know and I will help you, learning how to set up and use the program might take a minute but will save you a lot of time in the long-run. My email is rpn241@nyu.edu

For reference, I have also attached the .py and the .json so you can see the source code if you would like. The file sizes for the .exe’s zip folders are large because they include all the modules that are imported to the original .py file.

Feel free to test out the program on the ‘Test File’ in the main ‘Program’ folder. Just run the program and type in ‘Test File’, the file is already shared with the program.




