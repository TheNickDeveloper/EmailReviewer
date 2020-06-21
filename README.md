# EmailReviewer
 
An Outlook automation script which can help auto classify the email based on the defined types. Tech highlight: Outlook, Entity Framework, SQLite.

## Design
1. Outlook dll for controlling the behavior of outlook.
2. Sqlite Entity Framework code first for storing email information.

## Function

1. Auto classify the email and move to corresponding folder.
2. Can also classify the email based on the ticket number if there is no priority in email subject.
3. Store the email information in sqlite db further data analysis.
4. Can auto reply email with defined contents if priority reach the critial level. 

## How to Use

#### 1. Configuration setting

Configure source folder name and log output folder path in app.config file.

#### 2. Define the script execute plan

Set execute plan in Scheduler for email scripting.
