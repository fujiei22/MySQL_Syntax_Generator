# SqlSyntaxGenerator

The SqlSyntaxGenerator is a command-line tool that facilitates the generation of SQL statements based on data from Excel files. This tool allows users to convert Excel data into INSERT or UPDATE SQL statements, providing a seamless way to interact with relational databases.

## Usage
### 1. Install Dependencies:  
Ensure you have Node.js installed on your machine. Install the required Node.js packages using the following command:  
```
npm install
```

### 2. Run the Program:  
Execute the program using the following command:  
```
node db.js
```

### 3. Follow Prompts:  
 + Enter the Excel file name when prompted.  
 + Select the desired worksheet from the list.  
 + Provide the table name for SQL statement generation.  
 + Choose between INSERT or UPDATE SQL statements.  
 + Follow additional prompts based on your selection.

### 4. Output:  
The generated SQL statements will be saved to an output file located in the "output" folder. The file name includes a timestamp for uniqueness.

## Features  
 + Supports both INSERT and UPDATE SQL statement generation.
 + Handles special characters in Excel data, such as quotes and backticks, to prevent SQL errors.  
 + Allows users to choose the target worksheet from multi-sheet Excel files.

## Note  
 + Ensure the Excel file is in the "xlsx" directory.
 + Review and modify the generated SQL statements as needed for your specific database.
