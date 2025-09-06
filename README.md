# Excel Reader
A simple console app that reads data from an Excel spreadsheet, saves it to a database, then displays it on the console. Developed with C# and SQLite

## Given Requirements
* When the application starts, it should delete the database if it exists, create a new one, create all tables, read from Excel, seed into the database.
* EPPlus package should be used.
* You shouldn't read into Json first.
* You can use SQLite or SQL Server (or MySQL if you're using a Mac).
* Once the database is populated, you'll fetch data from it and show it in the console.
* You don't need any user input
* You should print messages to the console letting the user know what the app is doing at that moment (i.e. reading from excel; creating tables, etc.)
* The application will be written for a known table, you don't need to make it dynamic.
* When submitting the project for review, you need to include an xls file that can be read by your application.

## Features 
* Users has the ability to input a file path to any Excel file they wish to transpose into a DB.
* Console interface letting the user know what is currently happening.
* Spectre.Console package for beautified console logs and displaying data in a table.

## Challenges
* Making the application able to dynamically read from any header no matter what the number and the contents of the headers are.

## Lessons Learned
* Working with EPPlus is very important for Office files I/O.
* Spectre Console is worth it for the visual improvement.

## Areas To Improve
* Using EPPlus
* Using EFCore with dynamic shadow properties

## Resources
* EPPlus videos on YT

