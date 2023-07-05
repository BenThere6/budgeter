## AutoBudgeter

# Description
This program automatically scrapes bank transactions from emails, sorts the transactions, loads them into a budget spreadsheet, then sends text updates. 
I had built this program to work specifically for me and my device. It worked great for me, but it has not been updated for general use. I uploaded for future reference, and so others can see what I've completed in the past.

# Process
* Test internet connection
* Check if workbook is open
* Delete unrelated emails (with no bank transactions)
* If there are emails with transactions, save the html code to computer
* Parse html to get the transaction amounts and descriptions
* Check if there is more than one transaction per email, make sure all data has been retrieved
* Open excel, save all retrieved data to a tab in budget spreadsheet
* Sort data in spreadsheet based on predetermined categories
* If a transaction is not recognized, send text to user with transaction details. User will reply with description keywords to look for next time, and which category this transaction should go in. When the program receives the text, the keywords and category will be saved to a file on the computer so it will know what to do next time it sees this kind of transaction.
* Updated budget information (percentages of categories based on targets and actuals) is sent by text to the user
* Unecessary files are removed from computer
* If there are any issues throughout this process a text is sent to the user with releveant information

# Credit
All of this code was written or compiled by me. Multiple functions in this program were sampled from code I found online, most notably the functions dealing with email and scraping html. I did not make note of everywhere I took code from. 
