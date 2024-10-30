from src.ExcelSage import *
from assertpy import assert_that
from openpyxl import Workbook
import openpyxl as xl
import pytest

exl = ExcelSage()

EXCEL_FILE_PATH = r".\data\sample.xlsx"
INVALID_EXCEL_FILE_PATH = r"..\data\sample1.xlsx"
NEW_EXCEL_FILE_PATH = r".\data\new_excel.xlsx"
INVALID_SHEET_NAME = "invalid_sheet"
INVALID_CELL_ADDRESS = "AAAA1"


@pytest.fixture
def setup_teardown():
    yield
    if exl.active_workbook:
        exl.close_workbook()


def test_open_workbook_success(setup_teardown):
    workbook = exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    assert_that(workbook).is_instance_of(Workbook)


def test_open_workbook_file_not_found(setup_teardown):
    with pytest.raises(ExcelFileNotFoundError) as exc_info:
        exl.open_workbook(workbook_name=INVALID_EXCEL_FILE_PATH)

    assert_that(str(exc_info.value)).is_equal_to(f"Excel file '{INVALID_EXCEL_FILE_PATH}' not found. Please give the valid file path.")


def test_create_workbook_success(setup_teardown):
    sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
    workbook = exl.create_workbook(workbook_name=NEW_EXCEL_FILE_PATH, sheet_data=sheet_data, overwrite_if_exists=True)
    assert_that(workbook).is_instance_of(Workbook)


def test_create_workbook_file_already_exists(setup_teardown):
    with pytest.raises(FileAlreadyExistsError) as exc_info:
        exl.create_workbook(workbook_name=NEW_EXCEL_FILE_PATH)
        assert False, "Expected FileAlreadyExistsError but did not get one"

    assert_that(str(exc_info.value)).is_equal_to(f"Unable to create workbook. The file '{NEW_EXCEL_FILE_PATH}' already exists. Set 'overwrite_if_exists=True' to overwrite the existing file.")


def test_create_workbook_type_error(setup_teardown):
    with pytest.raises(TypeError) as exc_info:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], "John"]
        exl.create_workbook(workbook_name=NEW_EXCEL_FILE_PATH, sheet_data=sheet_data, overwrite_if_exists=True)

    assert_that(str(exc_info.value)).is_equal_to("Invalid row at index 3 of type 'str'. Each row in 'sheet_data' must be a list.")


def test_get_sheets_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    sheets = exl.get_sheets()
    assert_that(sheets).is_length(2).contains("Sheet1", "Sheet2")


def test_get_sheets_workbook_not_open(setup_teardown):
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.get_sheets()

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


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
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.add_sheet(sheet_name="Sheet4")

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


def test_add_sheet_sheet_exists(setup_teardown):
    with pytest.raises(SheetAlreadyExistsError) as exc_info:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.add_sheet(sheet_name="Sheet1", sheet_data=sheet_data, sheet_pos=2)

    assert_that(str(exc_info.value)).is_equal_to("Sheet 'Sheet1' already exists.")


def test_add_sheet_sheet_position(setup_teardown):
    with pytest.raises(InvalidSheetPositionError) as exc_info:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.add_sheet(sheet_name="Sheet5", sheet_data=sheet_data, sheet_pos=5)

    assert_that(str(exc_info.value)).is_equal_to("Invalid sheet position: 5. Maximum allowed is 2.")


def test_add_sheet_type_error(setup_teardown):
    with pytest.raises(TypeError) as exc_info:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], "John"]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.add_sheet(sheet_name="Sheet6", sheet_data=sheet_data, sheet_pos=1)

    assert_that(str(exc_info.value)).is_equal_to("Invalid row at index 3 of type 'str'. Each row in 'sheet_data' must be a list.")


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


def test_delete_sheet_doesnt_exists(setup_teardown):
    with pytest.raises(SheetDoesntExistsError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.delete_sheet(sheet_name=INVALID_SHEET_NAME)

    assert_that(str(exc_info.value)).is_equal_to(f"Sheet '{INVALID_SHEET_NAME}' doesn't exists.")


def test_fetch_sheet_data_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    sheet_data = exl.fetch_sheet_data(sheet_name="Sheet2",output_format="list", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, list)).is_true()

    sheet_data = exl.fetch_sheet_data(sheet_name="Sheet2",output_format="dict", starting_cell="A1", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, list) and all(isinstance(item, dict) for item in sheet_data)).is_true()

    sheet_data = exl.fetch_sheet_data(sheet_name="Sheet2",output_format="dataframe", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, DataFrame)).is_true()


def test_fetch_sheet_data_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.fetch_sheet_data(sheet_name="Sheet2",output_format="list", starting_cell=INVALID_CELL_ADDRESS, ignore_empty_columns=True, ignore_empty_rows=True)

    assert_that(str(exc_info.value)).is_equal_to(f"Cell '{INVALID_CELL_ADDRESS}' doesn't exists.")


def test_fetch_sheet_data_invalid_output_format(setup_teardown):
    with pytest.raises(ValueError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.fetch_sheet_data(sheet_name="Sheet2",output_format="invalid", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)

    assert_that(str(exc_info.value)).is_equal_to("Invalid output format. Use 'list', 'dict', or 'dataframe'.")


def test_rename_sheet_success(setup_teardown):
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    workbook.create_sheet(title="Sheet_to_rename")
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.rename_sheet(old_name="Sheet_to_rename", new_name="Sheet_renamed")

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheets = workbook.sheetnames
    assert_that(sheets).contains("Sheet_renamed")
    sheet_to_delete = workbook["Sheet_renamed"]
    workbook.remove(sheet_to_delete)
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_rename_sheet_workbook_not_open(setup_teardown):
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.rename_sheet(old_name="Sheet1", new_name="Sheet2")

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


def test_rename_sheet_sheet_exists(setup_teardown):
    with pytest.raises(SheetAlreadyExistsError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.rename_sheet(old_name="Sheet1", new_name="Sheet2")

    assert_that(str(exc_info.value)).is_equal_to("Sheet 'Sheet2' already exists.")


def test_rename_sheet_sheet_doesnt_exists(setup_teardown):
    with pytest.raises(SheetDoesntExistsError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.rename_sheet(old_name=INVALID_SHEET_NAME, new_name="Sheet2")

    assert_that(str(exc_info.value)).is_equal_to(f"Sheet '{INVALID_SHEET_NAME}' doesn't exists.")


def test_get_cell_value_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    cell_value = exl.get_cell_value(sheet_name="Sheet1",  cell_name="A1")
    assert_that(cell_value).is_equal_to("First Name")


def test_get_cell_value_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_cell_value(sheet_name="Sheet1",  cell_name=INVALID_CELL_ADDRESS)

    assert_that(str(exc_info.value)).is_equal_to(f"Cell '{INVALID_CELL_ADDRESS}' doesn't exists.")


def test_close_workbook_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.close_workbook()

    assert_that(exl.active_workbook).is_none()
    assert_that(exl.active_workbook_name).is_none()
    assert_that(exl.active_sheet).is_none()


def test_close_workbook_workbook_not_open(setup_teardown):
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.close_workbook()

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


def test_save_workbook_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.add_sheet(sheet_name="New_sheet")
    exl.save_workbook()

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook.sheetnames
    assert_that(sheet).is_length(3).contains("New_sheet")
    sheet_to_delete = workbook["New_sheet"]
    workbook.remove(sheet_to_delete)
    workbook.save(EXCEL_FILE_PATH)

def test_save_workbook_workbook_not_open(setup_teardown):
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.save_workbook()

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


def test_set_active_sheet_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    active_sheet = exl.set_active_sheet(sheet_name="Sheet2")

    assert_that(active_sheet).is_equal_to("Sheet2")
    assert_that(str(exl.active_sheet)).contains("Sheet2")


def test_set_active_sheet_workbook_not_open(setup_teardown):
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.set_active_sheet(sheet_name="Sheet2")

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


def test_set_active_sheet_sheet_doesnt_exists(setup_teardown):
    with pytest.raises(SheetDoesntExistsError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.set_active_sheet(sheet_name=INVALID_SHEET_NAME)

    assert_that(str(exc_info.value)).is_equal_to(f"Sheet '{INVALID_SHEET_NAME}' doesn't exists.")


def test_write_to_cell_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.write_to_cell(cell_name="AAA1", cell_value="New_value", sheet_name="Sheet1")

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook["Sheet1"]
    cell_value = sheet["AAA1"].value
    assert_that(cell_value).is_equal_to("New_value")
    sheet["AAA1"] = None
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_write_to_cell_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.write_to_cell(cell_name=INVALID_CELL_ADDRESS, cell_value="New_value", sheet_name="Sheet1")

    assert_that(str(exc_info.value)).is_equal_to(f"Cell '{INVALID_CELL_ADDRESS}' doesn't exists.")


def test_get_column_count_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    column_count = exl.get_column_count(sheet_name="Sheet2", starting_cell="D6", ignore_empty_columns=True)
    assert_that(column_count).is_equal_to(7)


def test_get_column_count_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_column_count(sheet_name="Sheet2", starting_cell=INVALID_CELL_ADDRESS)

    assert_that(str(exc_info.value)).is_equal_to(f"Cell '{INVALID_CELL_ADDRESS}' doesn't exists.")


def test_get_row_count_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    row_count = exl.get_row_count(sheet_name="Sheet2", include_header=True, starting_cell="D6", ignore_empty_rows=True)
    assert_that(row_count).is_equal_to(52)


def test_get_row_count_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_row_count(sheet_name="Sheet2", starting_cell=INVALID_CELL_ADDRESS)

    assert_that(str(exc_info.value)).is_equal_to(f"Cell '{INVALID_CELL_ADDRESS}' doesn't exists.")


def test_append_row_success(setup_teardown):
    data = ['Marisa', 'Pia', 'Female', 33, 'France', '21/05/2015', None, 1946]
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.append_row(sheet_name="Sheet1", row_data=data)

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook['Sheet1']

    for row in sheet.iter_rows():
        row_values = [cell.value for cell in row]
        if row_values == data:
            row_to_delete = row[0].row
            sheet.delete_rows(row_to_delete)
            break
    else:
        assert False, "Row not appended"

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_insert_row_success(setup_teardown):
    data = ['Mark', 'Pia', 'Male', 53, 'France', '11/09/2014', None, 1946]
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.insert_row(sheet_name="Sheet1", row_data=data, row_index=10)

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook['Sheet1']

    for row in sheet.iter_rows():
        row_values = [cell.value for cell in row]
        if row_values == data:
            row_to_delete = row[0].row
            assert_that(row_to_delete).is_equal_to(10)
            sheet.delete_rows(row_to_delete)
            break
    else:
        assert False, "Row not inserted at index 10"

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_insert_row_invalid_row_index(setup_teardown):
    with pytest.raises(InvalidRowIndexError) as exc_info:
        data = ['Mark', 'Pia', 'Male', 53, 'France', '11/09/2014', None, 1946]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.insert_row(row_data=data, row_index=1012323486523)

    assert_that(str(exc_info.value)).is_equal_to(f"Row index 1012323486523 is invalid or out of bounds. The valid range is 1 to 1048576.")


def test_delete_row(setup_teardown):
    data = ['Dee', 'Pia', 'Male', 53, 'France', '11/09/2014', None, 1946]
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook['Sheet1']
    sheet.insert_rows(2)
    for col_index, value in enumerate(data, start=1):
        sheet.cell(row=2, column=col_index, value=value)

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.delete_row(row_index=2)

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook['Sheet1']

    for row in sheet.iter_rows():
        row_values = [cell.value for cell in row]
        if row[0].row > 2:
            break
        if row_values == data:
            assert False, "Row not deleted at index 2"

    workbook.close()
