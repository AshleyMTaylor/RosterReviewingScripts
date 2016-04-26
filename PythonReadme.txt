Install Python from here:

https://www.python.org/ftp/python/2.7.11/python-2.7.11.amd64.msi

Then you can open a command prompt and run:

python CloseReview.py -r roster.csv -c accounts.csv -o journalEntries.csv

You need to create the roster and account csv files with the following columns (no headers):

Roster:
employeeId | charge_type | company_code | department | location | project

Costing:
charge_code | amount | employeeId | company_code | department | location | project

The roster and costing csv files need to be in the same directory as the python script file.  The output file name can be whatever you want but no spaces.