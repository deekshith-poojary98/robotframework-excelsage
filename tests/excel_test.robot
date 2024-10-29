*** Settings ***
Library    C:\\vscode\\ExcelLibrary.tar\\ExcelLibrary\\src\\ExcelSage.py


*** Variables ***
${excel}    ${CURDIR}//data//sample.xlsx


*** Test Cases ***
Test Excel Library
    [Documentation]    This is Test case for testing ExcelSage Lib
    Open Workbook    ${excel}
    Log    Workbook opened successfully
    ${sheets}    Get sheets
    Log    ${sheets}    WARN
    # Delete Sheet     TestSheet
    # Delete Sheet     SheetToRename
    # Delete Sheet     test_sheet
    # Delete Sheet     RenamedSheet


