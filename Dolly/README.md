The script analyzes the excel table, selecting from it the column names and their values.
Then replaces it in the template and stores the new instances in the specified folder
To create Word documents, you need to use the docx templates
If you need PDF files use the html template
The number of instances corresponds to the number of rows in the table.
To run the script from the command line, type
dolly.py [-h] [-t TYPE_FILE] data template output_folder