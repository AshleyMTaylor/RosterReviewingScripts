#!/usr/bin/python

import csv, argparse, xlsxwriter, xlrd

__author__ = 'Ashley Taylor'
 
class EmployeeRosterEntry:
    def __init__(self, employeeId, charge_type, company_code, department, location, project ):
        self.employeeId = employeeId
        self.charge_type = charge_type
        self.company_code = company_code
        self.department = department
        self.location = location
        self.project = project

class AccountCharge:
    def __init__(self, charge_code, amount, rosterEntryInfo ):
        self.charge_code = charge_code
        self.amount = amount
        self.rosterEntryInfo = rosterEntryInfo
        
class JournalEntry:
    def __init__(self, myAccountCharge, myEmployeeRosterEntryList, reasons_flagged):
        self.sourceAccountCombo = [myAccountCharge.rosterEntryInfo.company_code, myAccountCharge.rosterEntryInfo.department, myAccountCharge.rosterEntryInfo.project, myAccountCharge.charge_code, myAccountCharge.rosterEntryInfo.location]
        self.amount = myAccountCharge.amount
        self.charge_type = myAccountCharge.rosterEntryInfo.charge_type
        self.employeeId = myAccountCharge.rosterEntryInfo.employeeId
        self.employeeRosterInfo = myEmployeeRosterEntryList
        self.reasons_flagged = reasons_flagged

def main():
    parser = argparse.ArgumentParser(description='This script generated required journal entries for labor reclass.')
    parser.add_argument('-r','--rosterInputCsvFile', help='Roster Input file name',required=True)
    parser.add_argument('-c','--costingInputCsvFile', help='Costing Input file name',required=True)
    parser.add_argument('-s','--siteProjectsInputCsvFile', help='Site Projects Input file name',required=True)
    parser.add_argument('-o','--reclassJournalEntriesOutput',help='Reclass Output file name', required=True)
    args = parser.parse_args()
           
    rosterEntryListCSVFilename = args.rosterInputCsvFile
    accountChargeListCSVFilename = args.costingInputCsvFile
    validSiteProjectsListFilename = args.siteProjectsInputCsvFile
    journalEntryListCSVFilename = args.reclassJournalEntriesOutput
    
    print "Review the close!"

# --------------------------------------------------------------------------------
# Generate the RosterEntryList from the Roster CSV File
# --------------------------------------------------------------------------------

# Create a new empty list
    RosterEntryList = []

# Open the CSV file
    with open(rosterEntryListCSVFilename, 'rb') as csvRosterFile:
        csvRosterFileReader = csv.reader(csvRosterFile, delimiter=',')
        for row in csvRosterFileReader:
            # Parse each line for the account charge data
            # employeeId | charge_type | company_code | department | location | project
            RosterEntryList.append(EmployeeRosterEntry(row[0], row[1], row[2], row[3], row[4], row[5] ) )
    
# Create a new employee_dictionary
    employee_dictionary = {}

# Now loop through and create the employee dictionary to compare the account charges to
    for nextEmployeeRosterEntry in RosterEntryList:

# Create the employee in the dictionary if not present
        if not employee_dictionary.has_key(nextEmployeeRosterEntry.employeeId):
            employee_dictionary[nextEmployeeRosterEntry.employeeId] = []

# Now add this roster entry info to the list of roster entries for this employee'
        employee_dictionary[nextEmployeeRosterEntry.employeeId].append(nextEmployeeRosterEntry)

# --------------------------------------------------------------------------------
# Generate the account charges list from the Account Charges CSV File
# --------------------------------------------------------------------------------

# Create a new empty list
    ValidAccountChargeList = []
    InvalidAccountChargeList = []
    
    # Open the CSV file containing the list of valid sites and project codes
    # site_code | project_code
    
    validSiteProjectsDictionary = {}
    with open(validSiteProjectsListFilename, 'rb') as validSiteProjectsDictionaryFile:
        validSiteProjectsDictionaryFileReader = csv.reader(validSiteProjectsDictionaryFile, delimiter=',')
       
        for row in validSiteProjectsDictionaryFileReader:
            if validSiteProjectsDictionary.has_key(row[0]):
                if row[1] not in validSiteProjectsDictionary[row[0]]:
                    validSiteProjectsDictionary[row[0]].add(row[1])
            else:
                validSiteProjectsDictionary[row[0]] = {row[1]}
    
    # Open the account charges list file
    # charge_code | amount | employeeId | company_code | department | location | project
    
    # we only ever care about account charge codes between 4200 and 5000
    # we also only ever care about GP departments, which are 401, 402, 403, 405, and 471
    
    with open(accountChargeListCSVFilename, 'rb') as csvAccountChargeFile:
        csvAccountChargeFileReader = csv.reader(csvAccountChargeFile, delimiter=',')
        entriesToConsider = 0
        validGpDepartments = {'401','402','403','405','471'}
        
        for row in csvAccountChargeFileReader:
            if 4200 <= int(row[0]) <= 5000:
                if validSiteProjectsDictionary.has_key(row[5]):
                    if row[4] in validGpDepartments:
                        entriesToConsider += 1

                        if row[6] in validSiteProjectsDictionary[row[5]]:
                            # Right now we don't really care about the valid charges, but we might later on if we want to check these vs the roster
                            ValidAccountChargeList.append(AccountCharge(row[0], row[1], EmployeeRosterEntry(row[2], "N/A", row[3], row[4], row[5], row[6]) ))

                        else:
                            InvalidAccountChargeList.append(AccountCharge(row[0], row[1], EmployeeRosterEntry(row[2], "N/A", row[3], row[4], row[5], row[6]) ))                    
        
        print "Entries to Consider: " , entriesToConsider
        
# --------------------------------------------------------------------------------
# Iterate through the account charges and create journal entries for discrepancies
# --------------------------------------------------------------------------------

    journalEntryList = []
    for nextAccountCharge in InvalidAccountChargeList:
        if employee_dictionary.has_key(nextAccountCharge.rosterEntryInfo.employeeId):
            matches = 0
            for nextRosterEntry in employee_dictionary[nextAccountCharge.rosterEntryInfo.employeeId]:
                if nextRosterEntry.employeeId == nextAccountCharge.rosterEntryInfo.employeeId:
                    if nextRosterEntry.company_code == nextAccountCharge.rosterEntryInfo.company_code:
                        if nextRosterEntry.department == nextAccountCharge.rosterEntryInfo.department:            
                            if nextRosterEntry.location == nextAccountCharge.rosterEntryInfo.location:
                                if nextRosterEntry.project == nextAccountCharge.rosterEntryInfo.project: 
                                    matches += 1
                
            if matches == 0:
                journalEntryList.append(JournalEntry(nextAccountCharge, employee_dictionary[nextAccountCharge.rosterEntryInfo.employeeId], "INVALID_ACCT_CHARGE_NO_MATCHING_ROSTER_ENTRY"))
            else:
                journalEntryList.append(JournalEntry(nextAccountCharge, employee_dictionary[nextAccountCharge.rosterEntryInfo.employeeId], "INVALID_ACCT_CHARGE_BUT_MATCHED_ROSTER"))
                
        else:
            
            journalEntryList.append(JournalEntry(nextAccountCharge,[],"INVALID_ACCOUNT_CHARGE_EMPLOYEE_NOT_IN_ROSTER"))
    
    # Don't care right now about valid account charge entries, so don't need to check the roster for these
    
    #for nextAccountCharge in ValidAccountChargeList:
        #if employee_dictionary.has_key(nextAccountCharge.rosterEntryInfo.employeeId):
            #matches = 0
            #for nextRosterEntry in employee_dictionary[nextAccountCharge.rosterEntryInfo.employeeId]:
                #if nextRosterEntry.employeeId == nextAccountCharge.rosterEntryInfo.employeeId:
                    #if nextRosterEntry.company_code == nextAccountCharge.rosterEntryInfo.company_code:
                        #if nextRosterEntry.department == nextAccountCharge.rosterEntryInfo.department:            
                            #if nextRosterEntry.location == nextAccountCharge.rosterEntryInfo.location:
                                #if nextRosterEntry.project == nextAccountCharge.rosterEntryInfo.project: 
                                    #matches += 1
        
            #if matches == 0:
                #journalEntryList.append(JournalEntry(nextAccountCharge, employee_dictionary[nextAccountCharge.rosterEntryInfo.employeeId], "VALID_ACCT_CHARGE_NO_MATCHING_ROSTER_ENTRY"))

            #else:
        
                #journalEntryList.append(JournalEntry(nextAccountCharge,[],"VALID_ACCOUNT_CHARGE_EMPLOYEE_NOT_IN_ROSTER"))        

# --------------------------------------------------------------------------------
# Output journal entries to CSV file
# --------------------------------------------------------------------------------

    with open(journalEntryListCSVFilename, 'wb') as cvsJournalEntryFile:
        workbook = xlsxwriter.Workbook("journalEntries.xlsx")    
        sheet1 = workbook.add_worksheet("Journal Entries")
        
        # Write Header Row
        fieldnames = ['Info Source','Company Code', 'Department','Project','Account','Location','Amount','EmployeeID','Charge Type','Reason for Change']
        journalWriter = csv.DictWriter(cvsJournalEntryFile, fieldnames=fieldnames)
        journalWriter.writeheader()
        column = 0
        row = 0
        for field in fieldnames:
            sheet1.write(row, column, field)
            column += 1
        row += 1
        
        journalEntriesProcessed = 0
        nextTenPercent = 10
        percentComplete = 0
        journalEntryListSize = len(journalEntryList)
        print "Exporting ", journalEntryListSize, " Entries"
        print "        0 %"
        # Write the journal entries and supporting information
        for nextJournalEntry in journalEntryList:
            # Journal Entry
            rowDictionary = {'Info Source':"Costing Report",'Company Code' : nextJournalEntry.sourceAccountCombo[0], 'Department' : nextJournalEntry.sourceAccountCombo[1],'Project' : nextJournalEntry.sourceAccountCombo[2],'Account' : nextJournalEntry.sourceAccountCombo[3],'Location' : nextJournalEntry.sourceAccountCombo[4],'Amount' : nextJournalEntry.amount,'EmployeeID' : nextJournalEntry.employeeId,'Charge Type' : nextJournalEntry.charge_type, 'Reason for Change' : nextJournalEntry.reasons_flagged}            
            journalWriter.writerow(rowDictionary)                 
            column = 0
            for columnEntry in rowDictionary:
                sheet1.write(row, column, rowDictionary[columnEntry])
                column += 1
            row += 1
            
            # Journal Entry Employee Supporting Information
            for nextEmployeeRosterEntry in nextJournalEntry.employeeRosterInfo:
                rowDictionary = {'Info Source':"Employee Roster",'Company Code' : nextEmployeeRosterEntry.company_code, 'Department' : nextEmployeeRosterEntry.department,'Project' : nextEmployeeRosterEntry.project,'Account' : nextJournalEntry.sourceAccountCombo[3],'Location' : nextEmployeeRosterEntry.location,'Amount' : "",'EmployeeID' : nextEmployeeRosterEntry.employeeId,'Charge Type' : nextEmployeeRosterEntry.charge_type, 'Reason for Change' : nextJournalEntry.reasons_flagged}
                journalWriter.writerow(rowDictionary)
                column = 0
                for columnEntry in rowDictionary:
                    sheet1.write(row, column, rowDictionary[columnEntry])
                    column += 1
                row += 1
                
            # Write a blank line    
            rowDictionary = {'Info Source':"",'Company Code' : "", 'Department' : "",'Project' : "",'Account' : "",'Location' : "",'Amount' : "",'EmployeeID' : "",'Charge Type' : "", 'Reason for Change' : ""}
            journalWriter.writerow(rowDictionary)
            column = 0
            for columnEntry in rowDictionary:
                sheet1.write(row, column, rowDictionary[columnEntry])
                column += 1
            row += 1
            
            journalEntriesProcessed += 1
            percentComplete = (100 * journalEntriesProcessed) / journalEntryListSize
            if percentComplete >= nextTenPercent:
                print "     ", percentComplete, "%"
                nextTenPercent += 10
        
        print "    100%"    
        print ""
        
        # We're done writing, so go ahead and close the Excel file    
        workbook.close()
        
    print "Required journal entries written to", journalEntryListCSVFilename
    

    
    return

if __name__ == '__main__':
    main()
    
    ## List sheet names, and pull a sheet by name
        ##
        #sheet_names = xl_workbook.sheet_names()
        #print('Sheet Names', sheet_names)
    
        #xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])
    
        ## Or grab the first sheet by index 
        ##  (sheets are zero-indexed)
        ##
        #xl_sheet = xl_workbook.sheet_by_index(0)
        #print ('Sheet name: %s' % xl_sheet.name)
    
        ## Pull the first row by index
        ##  (rows/columns are also zero-indexed)
        ##
        #row = xl_sheet.row(0)  # 1st row
    
        ## Print 1st row values and types
        ##
        #from xlrd.sheet import ctype_text   
    
        #print('(Column #) type:value')
        #for idx, cell_obj in enumerate(row):
        #cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
        #print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))
    
        ## Print all values, iterating through rows and columns
        ##
        #num_cols = xl_sheet.ncols   # Number of columns
        #for row_idx in range(0, xl_sheet.nrows):    # Iterate through rows
        #print ('-'*40)
        #print ('Row: %s' % row_idx)   # Print row number
        #for col_idx in range(0, num_cols):  # Iterate through columns
        #cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        #print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))