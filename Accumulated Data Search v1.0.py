# Written 2016-04-21 v1.0
# This program searches a directory and opens all excel files found in
# it, searching for work order number and or letter returning all rows with
# query found and saving it to an cvs file.
import xlrd
import os
import csv


# Look for query in each sheet and print rows containing query
def find_in_workbook_sheets(file_location, workorder, letter, new_file_name):

    # Open file and store in workbook variable
    workbook = xlrd.open_workbook(file_location)
   
        
    # Get number of sheets in workbook
    number_of_sheets = workbook.nsheets

    # Open each sheet, search for criteria entered by user and write matching rows to new file
    for index in range(number_of_sheets):
        sheet = workbook.sheet_by_index(index)
        for index in range(sheet.nrows):
            if workorder in sheet.row_values(index) and letter.upper() in sheet.row_values(index):
                with open(new_file_name,'a',newline='') as fp:
                    a =  csv.writer(fp,delimiter=',')
                    data = [sheet.row_values(index)]
                    a.writerows(data)


# Get the .xls and .xlsx files from the root and sub dirs
def get_file_contents(workorder, letter, new_file_name):
    file_location = r'S:\Production\Time Sheet Accumulated data'

    for roots, dirs, files in os.walk(file_location):
        for file in files:
            file_name = roots + '\\' + file
            print(file_name)
            if '.xlsx' in file_name:
                try:
                    find_in_workbook_sheets(file_name,workorder,letter,new_file_name)
                except:
                    continue
            elif '.xls' in file_name:
                try:
                    find_in_workbook_sheets(file_name,workorder,letter,new_file_name)
                except:
                    continue
            
# main
def main():

    # Get user query
    workorder = input('Enter first search criteria:  ')
    letter = input('Enter second search criteria or press enter to skip:  ')
    print('Looking.....')

    new_file_name = "Accumulated Data Search " + workorder + letter + ".csv"
    print(new_file_name)
    
    # Open new .xls file and insert headers
    try:
        with open(new_file_name,'w',newline='') as fp:
            a =  csv.writer(fp,delimiter=',')
            data = [['GROUP','EMPLOYEE NAME','DATE','TASK',
                     'W.O.#','LETTER','INSTALL HOURS','TO HOURS','UTO HOURS','HOURS WORKED','REPAIR HOURS',
                     'TTL HRS PER W.O.',' ','TTL HRS PER SUPERV','TTL HOURS PUNCHED',]]
            a.writerows(data)
    except:
        print('Please close the Accumulated Data Search Document and try again.')
        error = input('Press enter to exit')
        exit()



    # If the user enters a number parse the str to int
    if workorder.isdigit():
        workorder = int(workorder)
        get_file_contents(workorder,letter, new_file_name)
    else:
        get_file_contents(workorder,letter, new_file_name)
   
    print('Done...')
    input('Press enter to exit')
    
# Call main
main()
