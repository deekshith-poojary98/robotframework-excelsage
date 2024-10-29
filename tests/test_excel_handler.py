from src.ExcelSage import *
from assertpy import assert_that
from openpyxl import Workbook
import openpyxl as xl
import pytest

exl = ExcelSage()

EXCEL_FILE_PATH = r".\data\sample.xlsx"
INVALID_EXCEL_FILE_PATH = r"..\data\sample1.xlsx"
NEW_EXCEL_FILE_PATH = r".\data\new_excel.xlsx"

@pytest.fixture
def setup_teardown():
    yield
    if exl.active_workbook:
        exl.close_workbook()

def test_open_workbook_success(setup_teardown):
    workbook = exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    assert_that(workbook).is_instance_of(Workbook)

def test_open_workbook_file_not_found(setup_teardown):
    try:
        exl.open_workbook(workbook_name=INVALID_EXCEL_FILE_PATH)
        assert False, "Expected ExcelNotFoundError but did not get one"
    except ExcelFileNotFoundError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_create_workbook_success(setup_teardown):
    sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
    workbook = exl.create_workbook(workbook_name=NEW_EXCEL_FILE_PATH, sheet_data=sheet_data, overwrite_if_exists=True)
    assert_that(workbook).is_instance_of(Workbook)

def test_create_workbook_file_already_exists(setup_teardown):
    try:
        exl.create_workbook(workbook_name=NEW_EXCEL_FILE_PATH)
        assert False, "Expected FileAlreadyExistsError but did not get one"
    except FileAlreadyExistsError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_create_workbook_type_error(setup_teardown):
    try:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], "John"]
        exl.create_workbook(workbook_name=NEW_EXCEL_FILE_PATH, sheet_data=sheet_data, overwrite_if_exists=True)
        assert False, "Expected TypeError but did not get one"
    except TypeError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_get_sheets_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    sheets = exl.get_sheets()
    assert_that(sheets).is_length(2).contains("Sheet1", "Sheet2")

def test_get_sheets_workbook_not_open(setup_teardown):
    try:
        exl.get_sheets()
        assert False, "Expected WorkbookNotOpenError but did not get one"
    except WorkbookNotOpenError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_add_sheet_success(setup_teardown):
    sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.add_sheet(sheet_name="Sheet3", sheet_data=sheet_data, sheet_pos=2)

    workbook = xl.load_workbook(filename=EXCEL_FILE_PATH)
    sheets = workbook.sheetnames
    assert_that(sheets).is_length(3).contains("Sheet1", "Sheet2", "Sheet3")
    sheet_to_delete = workbook['Sheet3']
    workbook.remove(sheet_to_delete)
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

def test_add_sheet_workbook_not_open(setup_teardown):
    try:
        exl.add_sheet(sheet_name="Sheet4")
        assert False, "Expected WorkbookNotOpenError but did not get one"
    except WorkbookNotOpenError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_add_sheet_sheet_exists(setup_teardown):
    try:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.add_sheet(sheet_name="Sheet1", sheet_data=sheet_data, sheet_pos=2)
        assert False, "Expected SheetAlreadyExistsError but did not get one"
    except SheetAlreadyExistsError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_add_sheet_sheet_position(setup_teardown):
    try:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.add_sheet(sheet_name="Sheet5", sheet_data=sheet_data, sheet_pos=5)
        assert False, "Expected InvalidSheetPositionError but did not get one"
    except InvalidSheetPositionError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_add_sheet_type_error(setup_teardown):
    try:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], "John"]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.add_sheet(sheet_name="Sheet6", sheet_data=sheet_data, sheet_pos=1)
        assert False, "Expected TypeError but did not get one"
    except TypeError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_delete_sheet_success(setup_teardown):
    workbook = xl.load_workbook(filename=EXCEL_FILE_PATH)
    workbook.create_sheet(title="Sheet_to_delete")
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    deleted_sheet = exl.delete_sheet(sheet_name="Sheet_to_delete")
    assert_that(deleted_sheet).is_equal_to("Sheet_to_delete")

    workbook = xl.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook.sheetnames
    assert_that(sheet).is_length(2).contains("Sheet1", "Sheet2")
    workbook.close()

def test_delete_sheet_workbook_not_open(setup_teardown):
    try:
        exl.delete_sheet(sheet_name="Sheet_to_delete")
        assert False, "Expected WorkbookNotOpenError but did not get one"
    except WorkbookNotOpenError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_delete_sheet_doesnt_exists(setup_teardown):
    try:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.delete_sheet(sheet_name="Sheet_to_delete")
        assert False, "Expected SheetDoesntExistsError but did not get one"
    except SheetDoesntExistsError:
        assert True
    except Exception as e:
        assert False, f"Unexpected error occurred: {e}"

def test_fetch_sheet_data_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    sheet_data = exl.fetch_sheet_data(sheet_name="Sheet2",output_format="list", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, list)).is_true()

    sheet_data = exl.fetch_sheet_data(sheet_name="Sheet2",output_format="dict", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, list) and all(isinstance(item, dict) for item in sheet_data)).is_true()

    sheet_data = exl.fetch_sheet_data(sheet_name="Sheet2",output_format="dataframe", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, DataFrame)).is_true()
