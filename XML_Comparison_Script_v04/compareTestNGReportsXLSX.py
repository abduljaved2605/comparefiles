import pprint
from deepdiff import DeepDiff
import csv
import random
import logging
import xlsxwriter
import os
import xml.etree.ElementTree as ET     # Library Reference https://docs.python.org/3/library/xml.etree.elementtree.html

"""
File 1 is termed as testng-results_old.xml OR testNG_report_01*
File 2 is termed as testng-results_new.xml OR testNG_report_02*
"""

### Variables
# Used as Run ID
rand_value = random.randint(1000, 9999)
# File Paths
file_1_path = './data_set1/testng-results_old.xml'
file_2_path = './data_set2/testng-results_new.xml'
result_file_name = "results/report_" + str(rand_value) + ".xlsx"
log_file_name = "results/logs_" + str(rand_value) + ".txt"

# Counter to fetch the number of test cases that do not match in both files
global mismatch_counter
mismatch_counter = 0  

# List to store the data to be written to the results CSV file
csv_row_data_list = []   
# To store the TestNG statistics from all files
testNGStatsList = [] 

# CSV Fields
csv_header_fields = ['test_case_file1', 'test_case_file_2', 'status_file1', 'status_file2', 'description_file1', 'description_file2']

# Initializing the logger file
logging.basicConfig(filename=log_file_name, filemode='w', level=logging.DEBUG, format='%(name)s - %(levelname)s - %(message)s')


### Functions
"""
Function to fetch all test cases along with it's results in the given testng-results.xml file 
Argument: 
   file_path - The path to the testNG XML results
Returns:
    List of test cases present in the file 
    List of test cases along with its status and description as a Dictionary item
"""
def fetch_testcase_stats(file_path):
    logging.info('Statistics fetched from the file: %s', file_path)    
    # Load the XML file
    tree1 = ET.parse(file_path)
    root1 = tree1.getroot()

    for temp in root1.findall('.'):    
        ignored = temp.get('ignored')
        total = temp.get('total')
        passed = temp.get('passed')
        failed = temp.get('failed')
        skipped = temp.get('skipped')        
        logging.info('Passed: %s', passed)
        logging.info('Failed: %s', failed)
        logging.info('Skipped: %s', skipped)
        logging.info('Ignored: %s', ignored)
        logging.info('Total: %s', total)

    for temp in root1.findall('./suite'):    
        startedAt = temp.get('started-at')
        finishedAt = temp.get('finished-at') 
        durationMs = temp.get('duration-ms')
        logging.info('Started At: : %s', startedAt)
        logging.info('Finished At: %s', finishedAt)
        logging.info('Duration in ms: %s', durationMs)
        # Storing the stats as dictionary        
        dictValues = {"File Name": os.path.basename(file_path), "Passed": passed, "Failed": failed, "Skipped": skipped, 
                      "Ignored": ignored, "Total": total, "Started At": startedAt, 
                      "Finished At": finishedAt, "Duration in ms": durationMs}                
        testNGStatsList.append(dictValues)
    return  testNGStatsList


"""
Function to fetch all test cases along with it's results in the given testng-results.xml file 
Argument: 
   file_path - The path to the testNG XML results
Returns:
    List of test cases present in the file 
    List of test cases along with its status and description as a Dictionary item
"""
def fetch_testcase_with_results(file_path):
    logging.info("Processing " + file_path + " file...")
    testNGTestCasesList = []    # Initializing list to store the 'test cases' found in testng-results XML file
    testNGResultsList = []      # Initializing list to store the 'results' found in testng-results XML file
    # Load the XML file
    tree1 = ET.parse(file_path)
    root1 = tree1.getroot()

    for temp in root1.findall('./suite/test/class/test-method'):    
        name = temp.get('name')
        status = temp.get('status')
        description = temp.get('description')
        # Skipping through the unwanted test cases
        if name == 'TSQA_afterMethod' or name == 'TSQA_beforeMethod' or name == 'setTestSuite' or name == 'TSQA_setUp' or name == 'TSQA_tearDown':
            # Skipping suite and method setup and teardown test cases
            continue
        else:
            testNGTestCasesList.append(name)
            dictValues = {"status": status, "description": description}
            tempDic = {name: dictValues}
            testNGResultsList.append(tempDic)
    return testNGTestCasesList, testNGResultsList

"""
Function to fetch the results of test cases that are present in both files
and append the results to csv_row_data_list 
"""
def fetch_tests_present_in_both_file():
    # Fetching the results for test cases that are present in both File 1 and File 2
    logging.info('Fetching the results for test cases that are present in both File 1 and File 2')
    for test in unique_tests_in_both_files:
        test_file_1 = [item for item in testNG_report_01_results_list if item.get(test)]
        test_file_2 = [item for item in testNG_report_02_results_list if item.get(test)]
        if (test_file_1[0][test]['status'] == test_file_2[0][test]['status']) and (test not in missing_tests_in_file_1) and (test not in missing_tests_in_file_2):
            continue
        else:
            logging.debug('-----------------------------------')
            logging.debug(test)
            logging.debug('test_file_1 %s', test_file_1[0][test]['status'])
            logging.debug('test_file_2 %s', test_file_2[0][test]['status'])
            logging.debug('-----------------------------------')

            # Writing to CSV file with the following headers 
            # ['test_case_file1', 'test_case_file_2', 'status_file1', 'status_file2', 'description_file1', 'description_file2']
            row_data = dict(test_case_file1 = test, 
                            test_case_file_2 = test, 
                            status_file1=test_file_1[0][test]['status'], 
                            status_file2=test_file_2[0][test]['status'],
                            description_file1=test_file_1[0][test]['description'], 
                            description_file2=test_file_2[0][test]['description'])
            csv_row_data_list.append(row_data)
            global mismatch_counter
            mismatch_counter += 1


"""
Function to fetch the results of test cases that are that are not present in File 1
and append the results to csv_row_data_list 
"""
def fetch_tests_missing_in_file1():
    # Fetching the results for test cases that are not present in File 1
    logging.info('Fetching the results for test cases that are not present in File 1')
    for test in missing_tests_in_file_1:
        test_file_2 = [item for item in testNG_report_02_results_list if item.get(test)]
        logging.debug('-----------------------------------')
        logging.debug(test)
        logging.debug('test_file_1 - NULL')
        logging.debug('test_file_2 %s', test_file_2[0][test]['status'])
        logging.debug('-----------------------------------')
        row_data = dict(test_case_file1 = '', 
                        test_case_file_2 = test, 
                        status_file1= '', 
                        status_file2=test_file_2[0][test]['status'],
                        description_file1= '', 
                        description_file2=test_file_2[0][test]['description'])
        csv_row_data_list.append(row_data)
        global mismatch_counter
        mismatch_counter += 1

"""
Function to fetch the results of test cases that are that are not present in File 2
and append the results to csv_row_data_list 
"""
def fetch_tests_missing_in_file2():
    # Fetching the results for test cases that are not present in File 2
    logging.info('Fetching the results for test cases that are not present in File 2')
    for test in missing_tests_in_file_2:
        test_file_1 = [item for item in testNG_report_01_results_list if item.get(test)]
        logging.debug('-----------------------------------')
        logging.debug(test)
        logging.debug('test_file_1 %s', test_file_1[0][test]['status'])
        logging.debug('test_file_2 - NULL')
        logging.debug('-----------------------------------')
        row_data = dict(test_case_file1 = test, 
                        test_case_file_2 = '', 
                        status_file1=test_file_1[0][test]['status'], 
                        status_file2= '',
                        description_file1=test_file_1[0][test]['description'], 
                        description_file2= '')
        csv_row_data_list.append(row_data)
        global mismatch_counter
        mismatch_counter += 1

"""
Function to write the results stored in csv_row_data_list to the results.csv file
"""
def write_to_csv_file():
    logging.info('Number of test cases that do not match in both files: %s', mismatch_counter) 
    with open(result_file_name, 'w', newline='') as csvfile:
        # Creating a csv writer object 
        writer = csv.DictWriter(csvfile, fieldnames = csv_header_fields) 
            
        # Writing the fields defined in the variables sections
        writer.writeheader() 
        # Writing the list of dictionaries that are present stored in 'csv_row_data_list'
        writer.writerows(csv_row_data_list)

"""
Function to write the statistics overview to the results file
"""
def write_overview_to_results(workbook, sheet_name):
    # Sheet 1, Overview of the statistics
    worksheet = workbook.add_worksheet(sheet_name)
    bold = workbook.add_format({'bold': True}) 

    # Write the headers to the first row of the worksheet
    headers = list(testNGStatsList[0].keys())
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, bold)

    # Write the data to the remaining rows of the worksheet
    for row, row_data in enumerate(testNGStatsList, start=1):
        for col, cell_data in enumerate(row_data.values()):
            worksheet.write(row, col, cell_data)

"""
Function to write the Test cases found in the given input file to the results
"""
def write_file_results_to_results(workbook, sheet_name, results_list):
    worksheet = workbook.add_worksheet(sheet_name)
    bold = workbook.add_format({'bold': True})
    font_green = workbook.add_format({'font_color': 'green'})
    font_red = workbook.add_format({'font_color': 'red'})
    font_orange = workbook.add_format({'font_color': 'orange'})

    # Write headers
    worksheet.write(0, 0, 'Test Case ID', bold)
    worksheet.write(0, 1, 'Status', bold)
    worksheet.write(0, 2, 'Description', bold)

    # Write data
    for row, data in enumerate(results_list, start=1):
        # Get the test case ID
        test_case_id = list(data.keys())[0]
        
        # Get the test case data
        test_case_data = data[test_case_id]
        # Write the test case ID, status, and description to the worksheet
        if test_case_data['status'] == 'PASS':
            worksheet.write(row, 0, test_case_id, font_green)
            worksheet.write(row, 1, test_case_data['status'], font_green)
            worksheet.write(row, 2, test_case_data['description'])
        elif test_case_data['status'] == 'FAIL':
            worksheet.write(row, 0, test_case_id, font_red)
            worksheet.write(row, 1, test_case_data['status'], font_red)
            worksheet.write(row, 2, test_case_data['description'])
        else:
            worksheet.write(row, 0, test_case_id, font_orange)
            worksheet.write(row, 1, test_case_data['status'], font_orange)
            worksheet.write(row, 2, test_case_data['description'])

"""
Function to write the comparison output from the input files to the results
"""
def write_comparison_results_to_report(workbook, sheet_name, results_list):
    worksheet = workbook.add_worksheet(sheet_name)
    bold = workbook.add_format({'bold': True})    
    font_green = workbook.add_format({'font_color': 'green'})
    font_red = workbook.add_format({'font_color': 'red'})
    font_orange = workbook.add_format({'font_color': 'orange'})

    # Write headers
    headers = ['Test Case', 'Status (File1)', 'Status (File2)', 'Description (File1)', 'Description (File2)']
    
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, bold)    

    # Write data
    for row, data in enumerate(results_list, start=1):
        # Get the test case ID and data for both files
        test_case_id = data['test_case_file1']        
        status_file1 = data['status_file1']
        status_file2 = data['status_file2']
        desc_file1 = data['description_file1']
        desc_file2 = data['description_file2']
        # Write the data to the worksheet
        worksheet.write(row, 0, test_case_id)
        # Writing status_file1 with respective font colors
        if status_file1 == "PASS":
            worksheet.write(row, 1, status_file1, font_green)
        elif status_file1 == "FAIL":
            worksheet.write(row, 1, status_file1, font_red)
        else:
            worksheet.write(row, 1, status_file1, font_orange)
        # Writing status_file2 with respective font colors
        if status_file2 == "PASS":
            worksheet.write(row, 2, status_file2, font_green)
        elif status_file2 == "FAIL":
            worksheet.write(row, 2, status_file2, font_red)
        else:
            worksheet.write(row, 2, status_file2, font_orange)   
        
        worksheet.write(row, 3, desc_file1)
        worksheet.write(row, 4, desc_file2) 

def main():
    logging.info('RUN ID: %s', rand_value)
    logging.info('File 1 is termed as testng-results_old.xml OR testNG_report_01*')
    logging.info('File 2 is termed as testng-results_new.xml OR testNG_report_02')

    # Processing File 1 
    global testNG_report_01_testcases_list 
    testNG_report_01_testcases_list = []
    global testNG_report_01_results_list 
    testNG_report_01_results_list = []
    testNG_report_01_testcases_list, testNG_report_01_results_list =  fetch_testcase_with_results(file_1_path)
    fetch_testcase_stats(file_1_path)
    # Processing File 2 
    global testNG_report_02_testcases_list
    testNG_report_02_testcases_list = []
    global testNG_report_02_results_list
    testNG_report_02_results_list = []
    testNG_report_02_testcases_list, testNG_report_02_results_list =  fetch_testcase_with_results(file_2_path)
    fetch_testcase_stats(file_2_path)
    logging.debug("List of test cases in File 1...")
    logging.debug(testNG_report_01_testcases_list)
    logging.debug("List of test cases in File 2...")
    logging.debug(testNG_report_02_testcases_list)

    logging.info("Number of test cases found in File 1: %s", len(testNG_report_01_testcases_list))
    logging.info("Number of test cases found in File 2: %s", len(testNG_report_02_testcases_list))

    # Comparing if all test cases are present in both files
    logging.info('Comparing the test cases found from both files...')
    testNG_report_01_testcases_list.sort()
    testNG_report_02_testcases_list.sort()
    if testNG_report_01_testcases_list == testNG_report_02_testcases_list:
        logging.info('RESULT:' + 'Test cases from File 1 match with test cases found in File 2...')
    else:
        logging.info('RESULT:' + 'Test cases from File 1 DO NOT match with test cases found in File 2...')

    logging.debug("List of test cases with status in File 1... ")
    logging.debug(testNG_report_01_results_list)

    logging.debug("List of test cases with status in File 2...")
    logging.debug(testNG_report_02_results_list)

    # Comparing the results
    logging.info('Comparing the results from both files...')
    if testNG_report_01_results_list == testNG_report_02_results_list:
        logging.info('RESULT:' + 'Results from File 1 match with test cases found in File 2...')
    else:
        logging.info('RESULT:' + 'Results from File 1 DO NOT match with test cases found in File 2...')

    logging.info('Comparing the test cases and logging the results...')
    logging.info('Looping through all results from File 1 and File 2... ')
    logging.debug('---------------------- \n RESULT \n ----------------------')

    # # Results from DeepDiff function
    # result = DeepDiff(testNG_report_01_results_list, testNG_report_02_results_list)
    # pprint.pprint(result)

    ### This logic is to handle the test cases that are missing/additional in each file
    # Iterating over each test in testNG_report_01_testcases_list and checks if it is not present 
    # in testNG_report_02_testcases_list. 
    # We store the missing tests in a new list called missing_tests_in_file_2.
    global missing_tests_in_file_2
    missing_tests_in_file_2 = [item for item in testNG_report_01_testcases_list if item not in testNG_report_02_testcases_list]
    # Similar to above logic
    global missing_tests_in_file_1
    missing_tests_in_file_1 = [item for item in testNG_report_02_testcases_list if item not in testNG_report_01_testcases_list]

    # Test cases from both File 1 and File 2
    all_test_cases = list(set(testNG_report_01_testcases_list + testNG_report_02_testcases_list))

    # Unique Test cases present in both files
    global unique_tests_in_both_files
    unique_tests_in_both_files = list(set(all_test_cases) - set(missing_tests_in_file_1) - set(missing_tests_in_file_2))

    logging.info('List of all test cases that are found in File1 and File2: %s', all_test_cases)
    logging.info('List of missing test cases that are found in File2 and not found in File1: %s', missing_tests_in_file_1)
    logging.info('List of missing test cases that are found in File1 and not found in File2: %s', missing_tests_in_file_2)

    # Calling the functions to compare and fetch the results and write it to a CSV file
    fetch_tests_present_in_both_file()
    fetch_tests_missing_in_file1()
    fetch_tests_missing_in_file2()
    
    # Generating the output results file
    logging.info('Generating the output results file %s...', result_file_name)
    workbook = xlsxwriter.Workbook(result_file_name)
    write_overview_to_results(workbook, "Overview")
    write_file_results_to_results(workbook, "File 1 Results", testNG_report_01_results_list)
    write_file_results_to_results(workbook, "File 2 Results", testNG_report_02_results_list)
    write_comparison_results_to_report(workbook, "Comparison Results", csv_row_data_list)
    workbook.close()
    logging.info('Completed generating the output results file %s...', result_file_name)

if __name__ == "__main__":
    main()