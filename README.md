# GetRecentDataFromYahoo

Yet another Excel VBA tool to fetch financial data from Yahoo.


About the spreadsheet
-----------------------------------
As Yahoo discontinued their ichart download API since 2017, it becomes more difficult to fetch financial data from Yahoo Finance, especially historical data, which requires a valid cookie to access. This Excel VBA tool doe not deal with historical data, but it allows users to fetch multiple type of real-time information from Yahoo Finance website, using some of the Yahoo APIs that programmers found out on the Internet.

Instructions
----------------------------------
See the Excel spreadsheet for details. The instructions tell you how to integrate it into your workbook.
The main part of the code is put into the Excel worksheet object rather than into a standard module at  request. For convenience, I also exported the code to GetDataModule.bas. You do not need to import anything to the spreadsheet, the spreadsheet is already usable. 

Change history
---------------------------------

**Ver 2.00**					

1. Significantly improved the stability when fetching basic information, including prices, time and exchanges.					

2. Added Sector and Industry. Other asset profile information is provided in the work area.					

3. Use the open source VBA-JSON to parse response from Yahoo. With VBA-JSON it can parse complex JSON structures, 					
    including asset profile information from Yahoo website.					
    VBA-JSON is published under MIT License. 

**Ver 2.01**

1. Added 'On Error' statement    					

Links
---------------------------------
[VBA-JSON](https://github.com/VBA-tools/VBA-JSON)  (c) Tim Hall, JSON conversion and parsing for VBA, [MIT Licnese](./VBA-JSON-2.3.1/LICENSE)

[Stack Overflow answer about Yahoo API](https://stackoverflow.com/questions/44030983/yahoo-finance-url-not-working/47505102#47505102)

