"""
The module contains procedures for performing
operations on accounting data such as parsing,
cleaning, conversion, evaluation and calculations.
"""

import re
from io import StringIO
import pandas as pd
from pandas import DataFrame, Series
import numpy as np

FilePath = str

PENALTY_GEENRAL = 10
PENALTY_DELIVRY_QUOTE = 11
PENALTY_DELIVRY_DELAY = 12

_penalty_category_codes = [
	PENALTY_GEENRAL,
	PENALTY_DELIVRY_QUOTE,
	PENALTY_DELIVRY_DELAY
]

def _extract_case_ids(vals: Series, rx_patt: str) -> Series:
	"""Parses SAP-exported string number to a standard format
	string number.
	"""

	matches = vals.str.findall(rx_patt, re.I)
	matches.mask(matches.str.len() == 0, pd.NA, inplace = True)
	matches.mask(matches.notna(), matches.str[0], inplace = True)
	matches.mask(matches.notna(), matches.str[1], inplace = True)
	matches = pd.to_numeric(matches).astype("UInt64")

	return matches

def _parse_numbers(vals: Series, dtype: str) -> Series:
	"""Converts string amounts in the SAP format to floating point literals."""

	repl = vals.str.replace(".", "", regex = False).str.replace(",", ".", regex = False)
	repl = repl.mask(repl.str.endswith("-"), "-" + repl.str.rstrip("-"))
	parsed = pd.to_numeric(repl).astype(dtype)

	return parsed

def _strip_line_pipes(text: str) -> None:
	"""Strips pipes form every line of a text"""

	replaced = re.sub(r"^\|", "", text, flags = re.M)
	replaced = re.sub(r"\|$", "", replaced, flags = re.M)
	replaced = re.sub(r"\"", "", replaced, flags = re.M)

	return replaced

def convert_fbl5n_data(data: str, case_id_rx: str) -> DataFrame:
	"""Converts plain FBL5N text data into a panel dataset.

	Parameters:
	-----------
	data:
		The plain text data exported form the FBL5N transaction.

	case_id_rx:
		Regex pattern for matching and extraction of country-specific
		case ID numbering in the strings of the data 'Text' field.

	Returns:
	--------
	The converted data (column name and data type):
		- "Head_Office": `UInt64`
		- "Branch: `UInt64`
		- "Currency": `string`
		- "Document_Number": `UInt64`
		- "Document_Type": `string`
		- "Document_Date": `object`
		- "Due_Date": `object`
		- "Arrears": `Int32`
		- "Clearing_Document": `UInt64`
		- "DC_Amount": `float64`
		- "Account_Assignment": `string`
		- "Tax": `string`
		- "Text": `string`
		- "Clearing_Date": `datetime64[ns]`
		- "Case_ID": `UInt64`
	"""

	assert data != "", "Data is an unexpected empty string!"
	matches = re.findall(r"^\|\s*\d+.*\|$", data, re.M)
	raw_txt = "\n".join(matches)
	stripped = _strip_line_pipes(raw_txt)

	parsed = pd.read_csv(
		StringIO(stripped),
		dtype = "string",
		sep = "|",
		names = [
			"Head_Office",
			"Branch",
			"Currency",
			"Document_Number",
			"Document_Type",
			"Document_Date",
			"Due_Date",
			"Arrears",
			"Clearing_Document",
			"DC_Amount",
			"Account_Assignment",
			"Tax",
			"Text",
			"Clearing_Date"
	])

	# clean parsed data
	cleaned = parsed.apply(lambda x: x.str.strip())
	cleaned["Tax"].replace("**", "", inplace = True)

	# convert fields to approprite data types
	cleaned["DC_Amount"] = _parse_numbers(cleaned["DC_Amount"], dtype = "float64")
	cleaned["Head_Office"] = pd.to_numeric(cleaned["Head_Office"]).astype("UInt64")
	cleaned["Arrears"] = _parse_numbers(cleaned["Arrears"], dtype = "Int32")
	cleaned["Document_Number"] = cleaned["Document_Number"].astype("UInt64")
	cleaned["Branch"] = pd.to_numeric(cleaned["Branch"]).astype("UInt64")
	cleaned["Clearing_Document"] = pd.to_numeric(cleaned["Clearing_Document"]).astype("UInt64")
	cleaned["Document_Date"] = pd.to_datetime(cleaned["Document_Date"], dayfirst = True).dt.date
	cleaned["Due_Date"] = pd.to_datetime(cleaned["Due_Date"], dayfirst = True).dt.date
	cleaned["Clearing_Date"] = pd.to_datetime(cleaned["Clearing_Date"], dayfirst = True).dt.date

	# add new fields to data frame
	enriched = cleaned.assign(Case_ID = pd.NA)

	# case ID(s) extracted from item text
	rx_patt = fr"(\A|[^a-zA-Z])D[P]?\s*[-_/]?\s*({case_id_rx})"
	enriched["Case_ID"] = _extract_case_ids(enriched["Text"], rx_patt)

	return enriched

def convert_dms_data(data: str) -> DataFrame:
	"""Converts plain DMS text data to a panel dataset.

	Parameters:
	-----------
	data:
		The plain text data exported form the UDM_DISPUTE transaction.

	Returns:
	--------
	The converted data (column name and data type):
		- "Case_ID": `UInt64`
		- "Debitor": `UInt64`
		- "Created_On": `object`
		- "Processor": `string`
		- "Status_Sales": `string`
		- "Status_AC": `string`
		- "Notification": `UInt64`
		- "Category_Description": `string`
		- "Category": `UInt8`
		- "Root_Cause": `category`
		- "Autoclaims_Note": `string`
		- "Fax_Number": `string`
		- "Status": `UInt8`
		- "DMS_Assignment": `object`
	"""

	assert data != "", "Data is an unexpected empty string!"
	matches = re.findall(r"^\|.*?\|.*?\|\d+.*$", data, re.M)
	raw_txt = "\n".join(matches)
	stripped = _strip_line_pipes(raw_txt)

	parsed = pd.read_csv(
		StringIO(stripped),
		sep = "|",
		dtype = "string",
		names = [
			"Case_ID",
			"Debitor",
			"Created_On",
			"Processor",
			"Status_Sales",
			"Status_AC",
			"Notification",
			"Category_Description",
			"Category",
			"Root_Cause",
			"Autoclaims_Note",
			"Fax_Number",
			"Status",
			"DMS_Assignment",
		]
	)

	# trim string data
	cleaned = parsed.apply(lambda x: x.str.strip())

	cleaned["DMS_Assignment"] = cleaned["DMS_Assignment"].mask(
		cleaned["DMS_Assignment"] == "", pd.NA
	)

	# convert fields to appropriate data types
	cleaned["Debitor"] = pd.to_numeric(cleaned["Debitor"]).astype("UInt64")
	cleaned["Notification"] = cleaned["Notification"].astype("UInt64")
	cleaned["Case_ID"] = cleaned["Case_ID"].astype("UInt64")
	cleaned["Created_On"] =  pd.to_datetime(cleaned["Created_On"], dayfirst = True).dt.date
	cleaned["Root_Cause"] = cleaned["Root_Cause"].astype("category")
	cleaned["Category"] = pd.to_numeric(cleaned["Category"]).astype("UInt8")
	cleaned["Status"] = cleaned["Status"].astype("UInt8")
	cleaned["DMS_Assignment"] = cleaned["DMS_Assignment"].astype("object")

	return cleaned

def evaluate_obi_de(
		fbl5n_data: DataFrame, dms_data: DataFrame,
		queries: dict, acc_data_paths: str
	) -> DataFrame:
	"""Evaluates overdue parameters for OBI Germany
	from the exported FBL5N and DMS data.

	Parameters:
	-----------
	fbl5n_data:
		The converted FBL5N data.

	dms_data:
		The converted DMS data.

	queries:
		Pandas data queries and their respective descriptions.

		The queries are executed to the data, and, if any
		records are found, the respective query description is
		written to the "Note" column of the queried data.

	acc_data_paths:
		Paths to Excel (xlsx) files that
		contain an additional info for each
		disputed case.

		The information is then joined to the
		case ID numbers in the FBL5N data.
		
	Returns:
	--------
	The evaluation result.
	"""

	merged = pd.merge(fbl5n_data, dms_data, how = "left", on = "Case_ID")
	sorted_data = merged.sort_values(["Case_ID"], ascending = False)
	acc_datasets = []

	for data_path in acc_data_paths:
		acc_datasets.append(pd.read_excel(data_path))

	# leave only overdue items
	data = sorted_data.drop(sorted_data[(sorted_data["Arrears"] < 0)].index)

	# prep data for analysis
	data = data.assign(Note = pd.NA)

	data["Overdue_Days"] = pd.cut(
		x = data["Arrears"], right = True,
		labels = ["〈0 - 30〉", "(30 - 60〉", "(60 - 90〉", "(90 - 120〉", "(120 - ∞)"],
		bins = [-np.inf, 30, 60, 90, 120, np.inf]
	).astype("category")

	penalty_cases = data["Category"].isin(_penalty_category_codes)
	data.loc[penalty_cases.index, "Category_Description"] = "Penalties"

	customer_documents = data["Document_Type"].isin(["DZ", "DA"])
	data.loc[customer_documents.index, "Document_Type"] = "DZ/DA"

	data["Lower_Status_Sales"] = data["Status_Sales"].str.lower()
	data["Lower_Text"] = data["Text"].str.lower()

	# generate notes using query strings stated in processing rules
	for q_string in queries:
		idx = data.query(q_string).index
		data.loc[idx, "Note"] = queries[q_string]

	# prep field for merging by replacing NA with a dummy int
	data["Case_ID"].fillna(0, inplace = True)

	for acc_data in acc_datasets:

		if "Case_ID" in acc_data.columns:
			merging_key = "Case_ID"
		elif "Document_Number" in acc_data.columns:
			merging_key = "Document_Number"
		else:
			assert False, "Could not detect data merging key!"

		data = pd.merge(
			data, acc_data, how = "left",
			left_on = merging_key, right_on = merging_key
		)

		# copy the new description to the "Notes" field
		data["Note"].mask(data["Description"].notna(), data["Description"], inplace = True)

		# delete the "Description" column
		data.drop(["Description"], axis = 1, inplace = True)

	# restore field post-merging by replacing the dummy int with NA
	data["Case_ID"].mask(data["Case_ID"] == 0, pd.NA, inplace = True)

	# drop helper fields
	data.drop(["Lower_Status_Sales", "Lower_Text"], axis = 1, inplace = True)

	return data

def evaluate_austria(
		fbl5n_data: DataFrame, dms_data: DataFrame,
		customer_data: FilePath, queries: dict,
	) -> DataFrame:
	"""Evaluates overdue parameters for Austria
	from the exported FBL5N and DMS data.

	Parameters:
	-----------
	fbl5n_data:
		The converted FBL5N data.

	dms_data:
		The converted DMS data.

	customer_data:
		Path to the customers.xlsx file that
		contains an additional info for each
		customer account, such as the responsible 
		Sales Person, Channel, customer name and
		country.

		The information is then joined to the
		customer accounts in the FBL5N data.

	queries:
		Pandas data queries and their respective descriptions.

		The queries are executed to the data, and, if any
		records are found, the respective query description is
		written to the "Note" column of the queried data.

	Returns:
	--------
	The evaluation result.
	"""

	merged = pd.merge(fbl5n_data, dms_data, how = "left", on = "Case_ID")
	sorted_data = merged.sort_values(["Case_ID"], ascending = False)
	cust_data  = pd.read_excel(customer_data)

	data = pd.merge(
		sorted_data, cust_data, how = "left",
		left_on = "Head_Office", right_on = "Account"
	)

	# prep data for analysis
	data = data.assign(
		Note = pd.NA,
		Lower_Status_Sales = data["Status_Sales"].str.lower(),
		Lower_Text = data["Text"].str.lower(),
		Lower_Customer_Name = data["Customer_Name"].str.lower()
	)

	penalty_cases = data["Category"].isin(_penalty_category_codes)
	data.loc[penalty_cases.index, "Category_Description"] = "Penalties"

	data["Overdue_Days"] = pd.cut(
		x = data["Arrears"], right = True,
		bins = [-np.inf, -1, 0, 30, 60, 90, 120, np.inf],
		labels = [
			"nicht fällig","fällig heute",
			"1 - 30", "31 - 60", "61 - 90",
			"91 - 120", "> 120"
	]).astype("category")

	customer_documents = data["Document_Type"].isin(["DZ", "DA"])
	data.loc[customer_documents.index, "Document_Type"] = "DZ/DA"

	# generate notes using query strings stated in processing rules
	for q_string in queries:
		data.loc[data.query(q_string).index, "Note"] = queries[q_string]

	data = data.assign(Debitor_Combined = pd.NA)
	data.drop(["Lower_Status_Sales", "Lower_Text", "Arrears"], axis = 1, inplace = True)

	return data
