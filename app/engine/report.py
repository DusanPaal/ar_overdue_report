# pylint: disable = E0110, E1101

"""The module generates Excel overdue reports form the evaluated data."""

from datetime import date
import pandas as pd
from pandas import DataFrame, Series, ExcelWriter
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.workbook import Workbook

FilePath = str

def _get_col_width(vals: Series, col_name: str, add_width: int = 0) -> int:
	"""Returns an iteger representing the width of a column calculated
	as the maximum number of characters contained in column name and
	column values plus additional points provided with the 'add_width'
	argument (default 0 points).
	"""

	offset = 1 # additional offset factor in points

	if col_name.isnumeric():
		return 14 + add_width

	if col_name == "Agreement":
		return 11 + add_width

	if col_name in ("Valid_From", "Valid_To"):
		return 11 + add_width

	if col_name == "Payments":
		return 12 + add_width

	data_vals = vals.astype("string").dropna().str.len()
	data_vals = list(data_vals)
	data_vals.append(len(str(col_name)))
	width = max(data_vals) + offset + add_width

	return width

def _col_to_rng(
		data: DataFrame, first_col: str, last_col: str = None,
		row: int = -1, last_row: int = -1) -> str:
	"""
	Converts data position in a DataFrame object into excel range notation (e.g. 'A1:D1', 'B2:G2').
	If 'last_col' is None, then only single-column range will be generated (e.g. 'A:A', 'B1:B1').
	If 'row' is '-1', then the generated range will span all the column(s) rows (e.g. 'A:A', 'E:E').
	If 'last_row' is provided, then the generated range will include all data records up to the last
	row (including).

	Params:
	-------
	data: Data for which colum names should be converted to a range.
	first_col: Name of the first column.
	last_col: Name of the last column.
	row: Index of the row for which the range will be generated.
	last_row: Index of the last data row which location will be considered in the resulting range.

	Returns:
	---------
	Excel data range notation.
	"""

	if isinstance(first_col, str):
		first_col_idx = data.columns.get_loc(first_col)
	elif isinstance(first_col, int):
		first_col_idx = first_col
	else:
		assert False, "Argument 'first_col' has invalid type!"

	first_col_idx += 1
	prim_lett_idx = first_col_idx // 26
	sec_lett_idx = first_col_idx % 26

	lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
	lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
	lett = "".join([lett_a, lett_b])

	if last_col is None:
		last_lett = lett
	else:

		if isinstance(last_col, str):
			last_col_idx = data.columns.get_loc(last_col)
		elif isinstance(last_col, int):
			last_col_idx = last_col
		else:
			assert False, "Argument 'last_col' has invalid type!"

		last_col_idx += 1
		prim_lett_idx = last_col_idx // 26
		sec_lett_idx = last_col_idx % 26

		lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
		lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
		last_lett = "".join([lett_a, lett_b])

	if row == -1:
		rng = ":".join([lett, last_lett])
	elif first_col == last_col and row != -1 and last_row == -1:
		rng = f"{lett}{row}"
	elif first_col == last_col and row != -1 and last_row != -1:
		rng = ":".join([f"{lett}{row}", f"{lett}{last_row}"])
	elif first_col != last_col and row != -1 and last_row == -1:
		rng = ":".join([f"{lett}{row}", f"{last_lett}{row}"])
	elif first_col != last_col and row != -1 and last_row != -1:
		rng = ":".join([f"{lett}{row}", f"{last_lett}{last_row}"])
	else:
		assert False, "Undefined argument combination!"

	return rng

def _convert_dates(vals: Series) -> Series:
	"""Converts string dates to datetime.date objects."""

	converted = vals.apply(
		lambda x: (x - date(1899, 12, 30)).days if not pd.isna(x) else x
	)

	return converted

def _generate_formats(report) -> dict:

	formats = {}

	formats["money"] = report.add_format({"num_format": "#,##0.00", "align": "center"})
	formats["category"] = report.add_format({"num_format": "000", "align": "center"})
	formats["date"] = report.add_format({"num_format": "dd.mm.yyyy", "align": "center"})
	formats["general"] = report.add_format({"align": "center"})
	formats["header"] = report.add_format({
		"align": "center",
		"bg_color": "#F06B00",
		"font_color": "white",
		"bold": True
	})

	return formats

def _format_header(data: DataFrame, sheet: Worksheet, cell_format: Format) -> None:
	"""Applies a specific visual format to the header row of a worksheet."""

	header_index = 1

	excel_sheet_range = _col_to_rng(
		data, data.columns[0],
		data.columns[-1],
		row = header_index)

	params = {"type": "no_errors", "format": cell_format}
	sheet.conditional_format(excel_sheet_range, params)

def _write(writer: ExcelWriter, data: DataFrame, sht_name: str) -> Workbook:
	"""Writes data to a named workbook sheet."""

	data.columns = data.columns.str.replace("_", " ", regex = False)
	data.to_excel(writer, index = False, sheet_name = sht_name)

	# replace spaces in column names back with underscores
	# for a better field manupulation further in the code
	data.columns = data.columns.str.replace(" ", "_", regex = False)
	workbook = writer.book

	return workbook

def _apply_column_formats(data, sheet: Worksheet, formats: dict, date_fields: list) -> None:
	"""Applies specific visual formats to the columns of a worksheet."""

	for col_name in data.columns:

		col_width = _get_col_width(data[col_name], col_name)
		col_rng = _col_to_rng(data, col_name)

		if col_name == "DC_Amount":
			fmt = formats["money"]
		elif col_name == "Category":
			fmt = formats["category"]
		elif col_name in date_fields:
			fmt = formats["date"]
		else:
			fmt = formats["general"]

		# apply new column params
		sheet.set_column(col_rng, col_width, fmt)

def create_report_obi_de(
		file: FilePath, evaluated: DataFrame,
		fields: list, sht_names: dict
	) -> None:
	"""Creates Excel overdue report for
	OBI Germany from the evaluated data.

	Parameters:
	-----------
	file:
		The path to the .xlsx report file to be created.

	evaluated:
		The evaluated data that will be written to the report sheet.

	fields:
		A list of fields that defines columns to include 
		in the overdue report, specifying their order.

	sht_names:
		A dictionary mapping data types to their
		respective sheet names in the Excel file.
	"""

	data = evaluated.copy()
	data = data.reindex(columns = fields) # reorder fields
	date_fields = ("Document_Date", "Due_Date", "Created_On")

	# convert datetime format to excel native serial date format
	# prior to printing to file in order for the date vals
	# appearing in a correct format on the report sheet
	for field in date_fields:
		data[field] = _convert_dates(data[field])

	# print all and cleared items to separate sheets of a workbook
	with pd.ExcelWriter(file, engine = "xlsxwriter") as wrtr:
		sht_name = sht_names["data_sheet_name"]
		report = _write(wrtr, data, sht_name)
		data_sht = wrtr.sheets[sht_name]
		formats = _generate_formats(report)
		_apply_column_formats(data, data_sht, formats, date_fields)
		_format_header(data, data_sht, formats["header"])

def create_report_austria(
		file: FilePath, evaluated: DataFrame,
		fields: list, sht_names: dict
	) -> None:
	"""Creates Excel overdue report
	for Austria from the evaluated data.

	Parameters:
	-----------
	file:
		The path to the .xlsx report file to be created.

	evaluated:
		The evaluated data that will be written to the report sheet.

	fields:
		A list of fields that defines columns to include 
		in the overdue report, specifying their order.

	sht_names:
		A dictionary mapping data types to their
		respective sheet names in the Excel file.
	"""

	data = evaluated.copy()

	# reorder fields
	data = data.reindex(fields, axis = 1)
	sales_data = data[[
		"Head_Office",
		"Customer_Name",
		"Salesperson",
		"Channel"
	]].copy()

	data.rename({
		"Head_Office": "Debitor",
		"Customer_Name": "Debitor_Name",
		"Country": "Land"
	}, axis = 1, inplace = True)

	sales_data = sales_data.assign(Aenderung = pd.NA, Bemerkung = pd.NA)

	sales_data.rename({
		"Head_Office": "Debitor",
		"Customer_Name": "Debitor_Text",
		"Salesperson": "Responsible",
		"Channel": "Kanal",
		"Aenderung": "Änderung"
	}, axis = 1, inplace = True)

	sales_data_fields = (
		"Debitor", "Debitor_Text",
		"Responsible", "Änderung",
		"Kanal", "Bemerkung"
	)

	sales_data = sales_data.reindex(sales_data_fields, axis = 1)

	# convert datetime format to excel native serial date format
	# prior to printing to file in order for the date vals
	# appearing in a correct format on the report sheet
	date_fields = ("Document_Date", "Due_Date", "Clearing_Date", "Created_On")

	for field in date_fields:
		data[field] = _convert_dates(data[field])

	# print all and cleared items to separate sheets of a workbook
	with ExcelWriter(file, engine = "xlsxwriter") as wrtr:

		_write(wrtr, data, sht_names["data_sheet_name"])
		_write(wrtr, sales_data, sht_names["sales_sheet_name"])

		report = wrtr.book
		data_sht = wrtr.sheets[sht_names["data_sheet_name"]]
		sales_sht = wrtr.sheets[sht_names["sales_sheet_name"]]

		formats = _generate_formats(report)
		_apply_column_formats(data, data_sht, formats, date_fields)
		_format_header(data, data_sht, formats["header"])

		_apply_column_formats(sales_data, sales_sht, formats, date_fields)
		_format_header(sales_data, sales_sht, formats["header"])
