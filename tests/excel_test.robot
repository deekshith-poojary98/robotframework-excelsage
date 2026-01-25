*** Settings ***
Library    ../ExcelSage/ExcelSage.py


*** Variables ***
${excel}    ./data/sample_original.xlsx
@{columns}    A    B

*** Test Cases ***
Test Excel Library
    [Documentation]    This is a sample test case for testing ExcelSage Lib
    Open Workbook    ${excel}
    ${sheets}    Get sheets
    Log    ${sheets}    WARN

    Get Row Count    sheet_name=${sheets}[1]    include_header=True    starting_cell=D6
    ${row_count_before}    Get Row Count    sheet_name=${sheets}[1]    include_header=True    starting_cell=D6
    Log    Row count before deletion: ${row_count_before}    WARN
    ${duplicates}    Find Duplicates  output_format=list    starting_cell=D6    sheet_name=${sheets}[1]    delete=True    output_filename=output.xlsx    overwrite_if_exists=True
    Log    Deleted ${duplicates} rows    WARN
    Close Workbook
    
    
    Open Workbook    output.xlsx
    ${row_count_after}    Get Row Count    sheet_name=${sheets}[1]    include_header=True    starting_cell=D6
    Log    Row count after deletion: ${row_count_after}    WARN
    Should Be True    ${row_count_before} > ${row_count_after}
