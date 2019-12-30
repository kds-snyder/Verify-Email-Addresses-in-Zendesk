This program uses Zendesk User APIs to bulk verify user email addresses contained in an Excel file

The Microsoft Open XML SDK is used to read the Excel file, as explained in https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet. 
The DOM approach is used, which according to https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet#approaches-to-parsing-open-xml-files, 
could result in an Out of Memory exception if working with very large files. This program ran with no exceptions for a file containing 673 rows.

Program verifies each user email address as follows:
1. Get the user ID corresponding to the email address
2. Get the user identity corresponding to the user ID and email address
3. If the 'verified' property in the user identity is false, set it to true

To use the program, you need to have the email address of a Zendesk user who is allowed to modify user data, and an API token for Zendesk. 
To run the program, open it in Visual Studio and then enter CTRL + F5.