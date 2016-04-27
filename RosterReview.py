#!/usr/bin/python

import csv, argparse, xlsxwriter, xlrd

__author__ = 'Ashley Taylor'
 
class RosterErrorsEntry:
    def __init__(self, employeeId, company_code, department, location, project, problem ):
        self.employeeId = employeeId
        self.company_code = company_code
        self.department = department
        self.location = location
        self.project = project
        self.problem = problem

def main():
    parser = argparse.ArgumentParser(description='This script generated required journal entries for labor reclass.')
    parser.add_argument('-r','--rosterInputCsvFile', help='Roster Input file name (*.csv)',required=True)
    parser.add_argument('-s','--siteProjectsInputCsvFile', help='Site Projects Input file name (*.csv)',required=True)
    parser.add_argument('-b','--budgetDeptSummaryCsvFile', help='Budget Summary 200/300 Valid Departments file name (*.csv)',required=True)
    parser.add_argument('-o','--rosterErrorsOutput',help='Roster Errors Output file name (*.xlsx)', required=True)
    args = parser.parse_args()
           
    rosterEntryListCSVFilename = args.rosterInputCsvFile
    validSiteProjectsListFilename = args.siteProjectsInputCsvFile
    budgetDeptSummaryFilename = args.budgetDeptSummaryCsvFile
    rosterErrorsOutputXLSXFilename = args.rosterErrorsOutput
    
    validGpDepartments = {'401','402','403','405','471'}
    
    validClientServiceDepts = set()
    for x in range(200, 299):
        validClientServiceDepts.add(str(x))
        
    validMgmtSupportDepts = set()
    for x in range(300, 399):
        validMgmtSupportDepts.add(str(x))
        
    print "Review the roster!"

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

# --------------------------------------------------------------------------------
# Generate the list of valid client services and management support employee IDs
# from the 200/300 Dept Budget Summary CSV File
# --------------------------------------------------------------------------------

# Open the CSV file containing the list of valid employee IDs for 200/300 Depts
# employee ID | Department
    
    validMgmtSupportEmployeeIDs = set()
    validClientServicesEmployeeIDs = set()
    with open(budgetDeptSummaryFilename, 'rb') as budgetDeptSummaryFile:
        budgetDeptSummaryFileReader = csv.reader(budgetDeptSummaryFile, delimiter=',')
           
        for row in budgetDeptSummaryFileReader:
            if row[1] in validClientServiceDepts:
                if row[0] not in validClientServicesEmployeeIDs:
                    validClientServicesEmployeeIDs.add(row[0])
                    
            if row[1] in validMgmtSupportDepts:
                if row[0] not in validMgmtSupportEmployeeIDs:
                    validMgmtSupportEmployeeIDs.add(row[0])

# --------------------------------------------------------------------------------
# Generate the roster errors CSV file
# --------------------------------------------------------------------------------

# we only ever care about account charge codes between 4200 and 5000
# we also only ever care about GP departments, which are 401, 402, 403, 405, and 471

# Create a new empty list
    RosterErrorsEntryList = []

# Open the CSV file
    with open(rosterEntryListCSVFilename, 'rb') as csvRosterFile:
        csvRosterFileReader = csv.reader(csvRosterFile, delimiter=',')
        for row in csvRosterFileReader:
            # Parse each line
            # employeeId | charge_code | company_code | department | location | project
            if row[2] in validGpDepartments:
                if validSiteProjectsDictionary.has_key(row[3]): 
                    if row[4] not in validSiteProjectsDictionary[row[3]]:
                        # invalid project for location specified
                        RosterErrorsEntryList.append(RosterErrorsEntry(row[0], row[1], row[2], row[3], row[4],  "Invalid project/location combination" ) )
                else:
                    # invalid location code specified
                    RosterErrorsEntryList.append(RosterErrorsEntry(row[0], row[1], row[2], row[3], row[4],  "Invalid location code specified" ) )
            else:
                if row[2] in validClientServiceDepts:
                    if row[0] not in validClientServicesEmployeeIDs:
                        RosterErrorsEntryList.append(RosterErrorsEntry(row[0], row[1], row[2], row[3], row[4],  "Employee was not budgeted to client service department" ) )
                else:
                    if row[3] in validMgmtSupportDepts:
                        if row[0] not in validMgmtSupportEmployeeIDs:
                            RosterErrorsEntryList.append(RosterErrorsEntry(row[0], row[1], row[2], row[3], row[4], "Employee was not budgeted to management support department" ) )         
                    #else:
                        # This is not a 200 or 300 department so we don't care
                        
    print "Entries to Consider: " , len(RosterErrorsEntryList)
        
# --------------------------------------------------------------------------------
# Output journal entries to CSV file
# --------------------------------------------------------------------------------

    workbook = xlsxwriter.Workbook(rosterErrorsOutputXLSXFilename)    
    sheet1 = workbook.add_worksheet("Roster Entry Errors")
    
    # Write Header Row
    fieldnames = ['Employee ID','Company Code', 'Department','Location','Project','Problem']
    column = 0
    row = 0
    for field in fieldnames:
        sheet1.write(row, column, field)
        column += 1
    row += 1
    
    rosterEntriesProcessed = 0
    nextTenPercent = 10
    percentComplete = 0
    RosterErrorsEntryListSize = len(RosterErrorsEntryList)
    print "Exporting ", RosterErrorsEntryListSize, " Entries"
    print "        0 %"
    
    # Write the roster error entries and supporting information
    column = 0
    for RosterError in RosterErrorsEntryList:
        sheet1.write(row, column, RosterError.employeeId)
        column += 1

        sheet1.write(row, column, RosterError.company_code)
        column += 1  
        
        sheet1.write(row, column, RosterError.department)
        column += 1
        
        sheet1.write(row, column, RosterError.location)
        column += 1               

        sheet1.write(row, column, RosterError.project)
        column += 1
        
        sheet1.write(row, column, RosterError.problem)
        column = 0           
        
        row += 1
        
        rosterEntriesProcessed += 1
        percentComplete = (100 * rosterEntriesProcessed) / RosterErrorsEntryListSize
        if percentComplete >= nextTenPercent:
            print "     ", percentComplete, "%"
            nextTenPercent += 10
    
    print "    100%"    
    print ""
    
    # We're done writing, so go ahead and close the Excel file    
    workbook.close()
        
    print "Roster entry errors written to", rosterErrorsOutputXLSXFilename
      
    return

if __name__ == '__main__':
    main()