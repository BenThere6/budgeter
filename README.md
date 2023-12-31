[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

# Auto Budgeter 
  
## Description
  
Experience streamlined financial management with this innovative program designed to automate the extraction, categorization, and integration of bank transactions from emails. Seamlessly sorting these transactions and populating them into a budget spreadsheet, the program goes above and beyond by providing text updates for added convenience.<br><br>Initially customized to match my specific device and preferences, this program has not been subsequently adapted for universal use. It is made available as a deliberate reference and a testament to my prior accomplishments. Its capabilities provide a glimpse into the potential of automated financial processing.

## Table of Contents

* [Process](#process)<br>
* [Contact Information](#contact-information)<br>
* [Credits](#credits)<br>
* [License](#license)

## Process

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

## Contact Information

For any further inquiries, please feel free to reach out to me through the following channels:
* GitHub: [My GitHub Profile](https://www.github.com/BenThere6)
* Email: benjaminbirdsall@icloud.com

I am here to assist you with any questions or feedback you may have. Thank you for your interest!

## Credits

The entirety of this codebase has been authored or curated by me. Notably, several functions within this program have been adapted from online sources, particularly those related to email handling and HTML scraping. While comprehensive attribution might not have been recorded for all instances of borrowed code, the integration of external resources has contributed to the program's robust functionality.

## License 

[MIT License](https://opensource.org/licenses/MIT)

The MIT License is a permissive open-source license that allows others to use, modify, and distribute your code for both commercial and non-commercial purposes. It requires that the original license notice and copyright notice are included in any redistributions.
