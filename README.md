This program uses the Zendesk Search and User APIs to bulk verify user email addresses contained in an Excel file. Basic authentication with an email address and API token is used as explained in https://developer.zendesk.com/rest_api/docs/support/introduction#api-token.

The Microsoft Open XML SDK is used to read the Excel file, as explained in https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet. The DOM approach is used, which according to https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet#approaches-to-parsing-open-xml-files, could result in an Out of Memory exception if working with very large files. This program ran with no exceptions for a file containing 673 rows.

The program verifies each user email address as follows:
1. Get the user ID corresponding to the email address
2. Get the user identity corresponding to the user ID and email address
3. If the 'verified' property in the user identity is false, set it to true

When you use the program, you will need to input the following:
1. Support API base (e.g. https://xyz.zendesk.com)
2. Email address of a Zendesk user who is an admin or agent
3. API token for Zendesk
4. Excel file name
5. Column in the Excel file containing the email addresses to verify (e.g. A)

To run the program, open it in Visual Studio and then enter CTRL + F5. The program will open a console window and start asking for the input specified above, and then process the file. For each email address processed, the program outputs a line to the console window containing the row number, email address, user ID, user identity ID, verified setting, and whether the email needed to be verified. When the program is finished, it outputs the following:

Total # email addresses verified

Total # email addresses already verified

Total # email addresses that couldn't be verified

Total # email addresses processed

Program duration time