The script parses the excel table by selecting the column names and their meanings from it.
Then substitutes it in a template and stores new instances in the specified folder
The number of instances corresponds to the number of rows in the table
To run the script from the command line, type 
python dolly.py --data path/to/Data.xlsx --template path/to/Invoice.docx --output path/to/output/