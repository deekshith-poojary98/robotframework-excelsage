from src.ExcelSage import *
from assertpy import assert_that
from openpyxl import Workbook
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.protection import SheetProtection
import openpyxl as xl
import pytest
import shutil
from pandas import DataFrame

exl = ExcelSage()

EXCEL_FILE_PATH = r".\data\sample.xlsx"
INVALID_EXCEL_FILE_PATH = r"..\data\sample1.xlsx"
NEW_EXCEL_FILE_PATH = r".\data\new_excel.xlsx"
INVALID_SHEET_NAME = "invalid_sheet"
INVALID_CELL_ADDRESS = "AAAA1"
INVALID_ROW_INDEX = 1012323486523
INVALID_COLUMN_INDEX = 163841


@pytest.fixture
def setup_teardown(scope='function', autouse=False):
    yield
    if exl.active_workbook:
        exl.close_workbook()


def copy_test_excel_file():
    source_file = r".\data\sample_original.xlsx"
    destination_file = r".\data\sample.xlsx"
    shutil.copy(source_file, destination_file)


def delete_the_test_excel_file():
    os.remove(EXCEL_FILE_PATH)


@pytest.fixture(scope="session", autouse=True)
def setup():
    copy_test_excel_file()
    yield
    delete_the_test_excel_file()


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
    assert_that(sheets).is_length(3).contains("Sheet1", "Offset_table", "Invalid_header")


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
    assert_that(sheets).is_length(4).contains("Sheet1", "Offset_table", "Sheet3", "Invalid_header")
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


def test_add_sheet_sheet__invalid_position(setup_teardown):
    with pytest.raises(InvalidSheetPositionError) as exc_info:
        sheet_data = [["Name", "Age"], ["Dee", 26], ["Mark", 56], ["John", 30]]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.add_sheet(sheet_name="Sheet5", sheet_data=sheet_data, sheet_pos=5)

    assert_that(str(exc_info.value)).is_equal_to("Invalid sheet position: 5. Maximum allowed is 3.")


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
    assert_that(sheet).is_length(3).contains("Sheet1", "Offset_table", "Invalid_header")
    workbook.close()


def test_delete_sheet_doesnt_exists(setup_teardown):
    with pytest.raises(SheetDoesntExistsError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.delete_sheet(sheet_name=INVALID_SHEET_NAME)

    assert_that(str(exc_info.value)).is_equal_to(f"Sheet '{INVALID_SHEET_NAME}' doesn't exists.")


def test_fetch_sheet_data_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    sheet_data = exl.fetch_sheet_data(sheet_name="Offset_table",output_format="list", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, list)).is_true()

    sheet_data = exl.fetch_sheet_data(sheet_name="Offset_table",output_format="dict", starting_cell="A1", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, list) and all(isinstance(item, dict) for item in sheet_data)).is_true()

    sheet_data = exl.fetch_sheet_data(sheet_name="Offset_table",output_format="dataframe", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)
    assert_that(isinstance(sheet_data, DataFrame)).is_true()


def test_fetch_sheet_data_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.fetch_sheet_data(sheet_name="Offset_table",output_format="list", starting_cell=INVALID_CELL_ADDRESS, ignore_empty_columns=True, ignore_empty_rows=True)

    assert_that(str(exc_info.value)).is_equal_to(f"Cell '{INVALID_CELL_ADDRESS}' doesn't exists.")


def test_fetch_sheet_data_invalid_output_format(setup_teardown):
    with pytest.raises(ValueError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.fetch_sheet_data(sheet_name="Offset_table",output_format="invalid", starting_cell="D6", ignore_empty_columns=True, ignore_empty_rows=True)

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
        exl.rename_sheet(old_name="Sheet1", new_name="Offset_table")

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


def test_rename_sheet_sheet_exists(setup_teardown):
    with pytest.raises(SheetAlreadyExistsError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.rename_sheet(old_name="Sheet1", new_name="Offset_table")

    assert_that(str(exc_info.value)).is_equal_to("Sheet 'Offset_table' already exists.")


def test_rename_sheet_sheet_doesnt_exists(setup_teardown):
    with pytest.raises(SheetDoesntExistsError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.rename_sheet(old_name=INVALID_SHEET_NAME, new_name="Offset_table")

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
    assert_that(sheet).is_length(4).contains("New_sheet")
    sheet_to_delete = workbook["New_sheet"]
    workbook.remove(sheet_to_delete)
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

def test_save_workbook_workbook_not_open(setup_teardown):
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.save_workbook()

    assert_that(str(exc_info.value)).is_equal_to("Workbook isn't open. Please open the workbook first.")


def test_set_active_sheet_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    active_sheet = exl.set_active_sheet(sheet_name="Offset_table")

    assert_that(active_sheet).is_equal_to("Offset_table")
    assert_that(str(exl.active_sheet)).contains("Offset_table")


def test_set_active_sheet_workbook_not_open(setup_teardown):
    with pytest.raises(WorkbookNotOpenError) as exc_info:
        exl.set_active_sheet(sheet_name="Offset_table")

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
    column_count = exl.get_column_count(sheet_name="Offset_table", starting_cell="D6", ignore_empty_columns=True)
    assert_that(column_count).is_equal_to(7)


def test_get_column_count_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_column_count(sheet_name="Offset_table", starting_cell=INVALID_CELL_ADDRESS)

    assert_that(str(exc_info.value)).is_equal_to(f"Cell '{INVALID_CELL_ADDRESS}' doesn't exists.")


def test_get_row_count_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    row_count = exl.get_row_count(sheet_name="Offset_table", include_header=True, starting_cell="D6", ignore_empty_rows=True)
    assert_that(row_count).is_equal_to(52)


def test_get_row_count_invalid_cell_address(setup_teardown):
    with pytest.raises(InvalidCellAddressError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_row_count(sheet_name="Offset_table", starting_cell=INVALID_CELL_ADDRESS)

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
        exl.insert_row(row_data=data, row_index=INVALID_ROW_INDEX)

    assert_that(str(exc_info.value)).is_equal_to(f"Row index {INVALID_ROW_INDEX} is invalid or out of bounds. The valid range is 1 to 1048576.")


def test_delete_row_success(setup_teardown):
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


def test_delete_row_invalid_row_index(setup_teardown):
    with pytest.raises(InvalidRowIndexError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.delete_row(row_index=INVALID_ROW_INDEX)

    assert_that(str(exc_info.value)).is_equal_to(f"Row index {INVALID_ROW_INDEX} is invalid or out of bounds. The valid range is 1 to 1048576.")


def test_append_column_success(setup_teardown):
    data = ["New Column", "data1", "data2", "data3", "data4", "data5"]
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.append_column(sheet_name="Sheet1", col_data=data)

    expected_values = ["data1", "data2", "data3", "data4", "data5"]
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook['Sheet1']
    last_column = sheet.max_column
    header = sheet.cell(row=1, column=last_column).value
    assert_that(header).is_equal_to("New Column")
    new_column_values = [sheet.cell(row=row, column=last_column).value for row in range(2, 7)]
    assert_that(new_column_values).is_equal_to(expected_values)

    sheet.delete_cols(last_column)
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_insert_column_success(setup_teardown):
    data = ["New Column", "data1", "data2", "data3", "data4", "data5"]
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.insert_column(sheet_name="Sheet1", col_data=data, col_index=1)

    expected_values = ["data1", "data2", "data3", "data4", "data5"]
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook['Sheet1']
    header = sheet.cell(row=1, column=1).value

    assert_that(header).is_equal_to("New Column")
    new_column_values = [sheet.cell(row=row, column=1).value for row in range(2, 7)]
    assert_that(new_column_values).is_equal_to(expected_values)

    sheet.delete_cols(1)
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_insert_column_invalid_column_index(setup_teardown):
    with pytest.raises(InvalidColumnIndexError) as exc_info:
        data = ["New Column", "data1", "data2", "data3", "data4", "data5"]
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.insert_column(sheet_name="Sheet1", col_data=data, col_index=INVALID_COLUMN_INDEX)

    assert_that(str(exc_info.value)).is_equal_to(f"Column index {INVALID_COLUMN_INDEX} is invalid or out of bounds. The valid range is 1 to 16384.")


def test_delete_column_success(setup_teardown):
    data = ["New Column", "data1", "data2", "data3", "data4", "data5"]
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook["Sheet1"]
    sheet.insert_cols(1)

    for row_index, value in enumerate(data, start=1):
        sheet.cell(row=row_index, column=1, value=value)

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.delete_column(sheet_name="Sheet1", col_index=1)

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook['Sheet1']
    header = sheet.cell(row=1, column=1).value
    workbook.close()

    assert_that(header).is_not_equal_to("New Column")


def test_delete_column_invalid_column_index(setup_teardown):
    with pytest.raises(InvalidColumnIndexError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.delete_column(sheet_name="Sheet1", col_index=INVALID_COLUMN_INDEX)

    assert_that(str(exc_info.value)).is_equal_to(f"Column index {INVALID_COLUMN_INDEX} is invalid or out of bounds. The valid range is 1 to 16384.")


@pytest.mark.parametrize("column_name, expected_length, expected_values, output_format", [
    ("A", 52, ["Tommie", "Nereida", "Stasia"], "list"),
    (["A", "B"], 52, [["Tommie", "Nereida", "Stasia"], ["Mccrystal", "Partain", "Hanner"]], "list"),
    ("First Name", 52, ["Tommie", "Nereida", "Stasia"], "list"),
    (["First Name", "Last Name"], 52, [["Tommie", "Nereida", "Stasia"], ["Mccrystal", "Partain", "Hanner"]], "list"),

    ("A", 52, [["Tommie", "Nereida", "Stasia"]], "dict"),
    (["A", "B"], 52, [["Tommie", "Nereida", "Stasia"], ["Mccrystal", "Partain", "Hanner"]], "dict"),
    ("First Name", 52, [["Tommie", "Nereida", "Stasia"]], "dict"),
    (["First Name", "Last Name"], 52, [["Tommie", "Nereida", "Stasia"], ["Mccrystal", "Partain", "Hanner"]], "dict"),

    ("A", 52, ["Tommie", "Nereida", "Stasia"], "DataFrame"),
    (["A", "B"], 52, [["Tommie", "Nereida", "Stasia"], ["Mccrystal", "Partain", "Hanner"]], "DataFrame"),
    ("First Name", 52, ["Tommie", "Nereida", "Stasia"], "DataFrame"),
    (["First Name", "Last Name"], 52, [["Tommie", "Nereida", "Stasia"], ["Mccrystal", "Partain", "Hanner"]], "DataFrame")
])
def test_get_column_values_success(setup_teardown, column_name, expected_length, expected_values, output_format):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    column_values = exl.get_column_values(column_names_or_letters=column_name, sheet_name="Offset_table", starting_cell="D6", output_format=output_format)

    if isinstance(column_name, str) or output_format in ["dict", "DataFrame"]:
        assert_that(isinstance(column_values, eval(output_format))).is_true()

        if output_format == "list":
            assert_that(column_values).is_length(expected_length).contains(*expected_values)
        elif output_format == "dict":
            keys = list(column_values.keys())
            for i, key in enumerate(keys):
                values = column_values.get(key)
                assert_that(values).is_length(expected_length).contains(*expected_values[i])
        else:
            row_count, column_count = column_values.shape
            assert_that(row_count).is_equal_to(expected_length)
            column_headers = column_values.columns.tolist()

            if column_count == 1:
                assert_that(column_values[column_headers[0]].tolist()).contains(*expected_values)
            else:
                assert_that(column_values[column_headers[0]].tolist()).contains(*expected_values[0])
                assert_that(column_values[column_headers[1]].tolist()).contains(*expected_values[1])
    else:
        assert_that(isinstance(column_values, eval(output_format))).is_true()
        assert_that(column_values).is_length(2)
        for sublist in column_values:
            assert_that(isinstance(sublist, list)).is_true()
            assert_that(sublist).is_length(expected_length)


def test_get_column_values_invalid_output_format(setup_teardown):
    with pytest.raises(ValueError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_column_values(column_names_or_letters="A", sheet_name="Offset_table", starting_cell="D6", output_format="invalid")

    assert_that(str(exc_info.value)).is_equal_to("Invalid output format. Use 'list', 'dict', or 'dataframe'.")


def test_get_column_values_invalid_header(setup_teardown):
    with pytest.raises(ValueError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_column_values(column_names_or_letters="G", sheet_name="Invalid_header", starting_cell="A1", output_format="list")

    assert_that(str(exc_info.value)).is_equal_to("Column letter 'G' does not have a valid string header: '1234' found.")


@pytest.mark.parametrize("invalid_columns",["AAAA", "Invalid Column Header"])
def test_get_column_values_invalid_column(setup_teardown, invalid_columns):
    with pytest.raises(ValueError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_column_values(column_names_or_letters=invalid_columns, sheet_name="Sheet1", starting_cell="A1", output_format="list")

    assert_that(str(exc_info.value)).is_equal_to(f"Invalid column name or letter: '{invalid_columns}'")


def test_get_column_values_column_out_of_bound(setup_teardown):
    with pytest.raises(ValueError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_column_values(column_names_or_letters="Z", sheet_name="Sheet1", starting_cell="A55", output_format="list")

    assert_that(str(exc_info.value)).is_equal_to(f"Column letter 'Z' is out of bounds for the provided sheet.")


@pytest.mark.parametrize("row_index, expected_values, expected_length, output_format", [
    (2,["Lester", "Prothro", "Male", 20, "France", "15/10/2017", None, 6574], 8, "list"),
    ([2, 3],[["Lester", "Prothro", "Male", 20, "France", "15/10/2017", None, 6574], ["Francesca", "Beaudreau", "Female", 21, "France", "15/10/2017", 5412]], 8, "list"),
    (2,[["Lester", "Prothro", "Male", 20, "France", "15/10/2017", None, 6574]], 8, "dict"),
    ([2, 3],[["Lester", "Prothro", "Male", 20, "France", "15/10/2017", None, 6574], ["Francesca", "Beaudreau", "Female", 21, "France", "15/10/2017", 5412]], 8, "dict")
 ])
def test_get_row_values_success(setup_teardown, row_index, expected_length, expected_values, output_format):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    row_value = exl.get_row_values(sheet_name="Sheet1", row_indices=row_index, output_format=output_format)
    assert_that(isinstance(row_value, eval(output_format))).is_true()

    if isinstance(row_index, int) or output_format == "dict":
        if output_format == "list":
            assert_that(row_value).is_length(expected_length).contains(*expected_values)
        else:
            keys = list(row_value.keys())
            for i, key in enumerate(keys):
                values = row_value.get(key)
                assert_that(values).is_length(expected_length).contains(*expected_values[i])
    else:
        assert_that(isinstance(row_value, eval(output_format))).is_true()
        assert_that(row_value).is_length(2)
        for sublist in row_value:
            assert_that(isinstance(sublist, list)).is_true()
            assert_that(sublist).is_length(expected_length)


def test_get_row_values_invalid_output_format(setup_teardown):
    with pytest.raises(ValueError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_row_values(sheet_name="Sheet1", row_indices=2, output_format="invalid")

    assert_that(str(exc_info.value)).is_equal_to("Invalid output format. Use 'list' or 'dict'.")


def test_get_row_values_invalid_row_index(setup_teardown):
    with pytest.raises(InvalidRowIndexError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.get_row_values(sheet_name="Sheet1", row_indices=INVALID_ROW_INDEX, output_format="list")

    assert_that(str(exc_info.value)).is_equal_to(f"Row index {INVALID_ROW_INDEX} is invalid or out of bounds. The valid range is 1 to 1048576.")


def test_protect_sheet_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.protect_sheet(sheet_name="Invalid_header", password="password")

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook["Invalid_header"]
    assert_that(sheet.protection.sheet).is_true()
    sheet.protection.set_password("password")
    sheet.protection.sheet = False
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_protect_sheet_sheet_already_protected(setup_teardown):
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook["Invalid_header"]
    sheet.protection.set_password("password")
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    with pytest.raises(SheetAlreadyProtectedError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.protect_sheet(sheet_name="Invalid_header", password="password")

    assert_that(str(exc_info.value)).is_equal_to("The sheet 'Invalid_header' is already protected and cannot be protected be again.")


def test_unprotect_sheet_success(setup_teardown):
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook["Invalid_header"]
    sheet.protection.set_password("password")
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.unprotect_sheet(sheet_name="Invalid_header", password="password")

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook["Invalid_header"]
    assert_that(sheet.protection.sheet).is_false()
    workbook.close()


def test_unprotect_sheet_sheet_not_protected(setup_teardown):
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    sheet = workbook["Invalid_header"]
    sheet.protection.set_password("password")
    sheet.protection.sheet = False
    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    with pytest.raises(SheetNotProtectedError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.unprotect_sheet(sheet_name="Invalid_header", password="password")

    assert_that(str(exc_info.value)).is_equal_to("The sheet 'Invalid_header' is not currently protected and cannot be unprotected.")


def test_protect_workbook_success(setup_teardown):
    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.protect_workbook(password="password", protect_sheets=True)

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    assert_that(workbook.security.lockStructure).is_true()
    workbook.security.lockStructure = False

    for sheet in workbook.worksheets:
        if sheet.protection.sheet:
            sheet.protection.sheet = False

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()


def test_protect_workbook_workbook_already_protected(setup_teardown):
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    protection = WorkbookProtection()
    protection.workbookPassword = "password"
    protection.lockStructure = True
    protection.lockWindows = True

    workbook.security = protection

    for sheet in workbook.worksheets:
        sheet.protection = SheetProtection(password="password")
        sheet.protection.sheet = True
        sheet.protection.formatCells = False
        sheet.protection.formatColumns = False
        sheet.protection.formatRows = False
        sheet.protection.insertColumns = False
        sheet.protection.insertRows = False
        sheet.protection.deleteColumns = False
        sheet.protection.deleteRows = False
        sheet.protection.sort = False
        sheet.protection.autoFilter = False
        sheet.protection.objects = False
        sheet.protection.scenarios = False

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    with pytest.raises(WorkbookAlreadyProtectedError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.protect_workbook(protect_sheets=True, password="password")

    assert_that(str(exc_info.value)).is_equal_to("The workbook is already protected and cannot be protected be again.")


def test_unprotect_workbook_success(setup_teardown):
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    protection = WorkbookProtection()
    protection.workbookPassword = "password"
    protection.lockStructure = True
    protection.lockWindows = True

    workbook.security = protection

    for sheet in workbook.worksheets:
        sheet.protection = SheetProtection(password="password")
        sheet.protection.sheet = True
        sheet.protection.formatCells = False
        sheet.protection.formatColumns = False
        sheet.protection.formatRows = False
        sheet.protection.insertColumns = False
        sheet.protection.insertRows = False
        sheet.protection.deleteColumns = False
        sheet.protection.deleteRows = False
        sheet.protection.sort = False
        sheet.protection.autoFilter = False
        sheet.protection.objects = False
        sheet.protection.scenarios = False

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
    exl.unprotect_workbook(unprotect_sheets=True)

    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    assert_that(workbook.security and workbook.security.lockStructure).is_false()
    workbook.close()


def test_unprotect_workbook_workbook_not_protected(setup_teardown):
    workbook = excel.load_workbook(filename=EXCEL_FILE_PATH)
    workbook.security.lockStructure = False

    for sheet in workbook.worksheets:
        if sheet.protection.sheet:
            sheet.protection.sheet = False

    workbook.save(EXCEL_FILE_PATH)
    workbook.close()

    with pytest.raises(WorkbookNotProtectedError) as exc_info:
        exl.open_workbook(workbook_name=EXCEL_FILE_PATH)
        exl.unprotect_workbook(unprotect_sheets=True)

    assert_that(str(exc_info.value)).is_equal_to("The workbook is not currently protected and cannot be unprotected.")
