# The-Dark-Night
This project utilizes the full potential of google's G SUITE APIs. APIs used in this project are - Docs, Drive, Sheets and Gmail. 
This is made to automate a process and reduce complexity in creating seperate documents(using mail merge), then extracting data of each document and mail it to seperate users.

- So we perform mail merge using Google Docs and sheets API. 
- All the source data is stored in a sheet in Google Sheets.
- A Doc template is created with variables( like- {{NAME}} ) inside the template to be replaced with actual data from the source.
- Data from the source is merged into the docs and a new document is generated for each row of data in source.
- Now we extract the content of this newly generated document from DOCS API.
- Put the content in a message body and generate a message to be sent via E-Mail using Gmail API.
- Now each document is sent to seperate email address(extracted from the source sheet). 
