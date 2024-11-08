*** Settings ***
Library    ..//src//ExcelSage.py


*** Variables ***
${excel}    ..//data//sample_original.xlsx


*** Test Cases ***
Test Excel Library
    [Documentation]    This is a sample test case for testing ExcelSage Lib
    Open Workbook    ${excel}
    Log    Workbook opened successfully
    ${sheets}    Get sheets
    Log    ${sheets}    WARN

