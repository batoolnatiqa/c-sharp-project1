# About my first project on github
This C# Windows Forms application enables users to import data from Excel files into a SQL Server database. It features duplicate detection to prevent redundant entries and provides functionalities to log messages for each record.

Features
Excel Data Import: Allows users to select and load data from Excel files into a DataGridView for preview.​

Duplicate Detection: Before inserting data into the SQL Server database, the application checks for existing records based on the RegNo field and alerts the user if duplicates are found.​

Message Logging: Enables users to insert messages into a log table for each RegNo entry, facilitating communication tracking.​

Components
1. User Interface
DataGridView: Displays the content of the selected Excel file for user review before database insertion.​

Buttons:

Import: Opens a dialog to select an Excel file and loads its data into the DataGridView.​

Save to Database: Inserts the data from the DataGridView into the SQL Server database, performing duplicate checks.​

Insert Message: Logs a user-defined message for each RegNo in the database.​

2. Functionalities
Excel File Selection and Loading:

Utilizes an OpenFileDialog to allow users to select Excel files.​
ASPSnippets

Reads data from the chosen file and populates the DataGridView for preview.​

Duplicate Entry Handling:

Before inserting data, the application queries the database to check for existing RegNo entries.​

If duplicates are detected, the application prompts the user with a message and excludes those entries from insertion.​

Message Logging:

Iterates through all RegNo entries in the database.​

Inserts a user-defined message along with a timestamp into a log table for each RegNo.​

Technologies Used
C# Windows Forms: For building the graphical user interface.​

SQL Server: To store imported data and log messages.​

ExcelDataReader Library: For reading data from Excel files
