# pylint: disable = C0103, W0603, W0703, W1203

"""
Description:
------------
The module autoates operations in the SAP transaction UDM_DISPUTE:

	- search_disputes():
		For searching of dipute cases by case ID.

	- export_disputes_data():
		For exporting the table of disputed cases
		found by the search_disputes() procedure.

How to use:
-----------
The module requires the pyperclip package to be installed
in the python environmnt. The package is used for copying 
the contents of a collection to the clipboard.

The UDM_DISPUTE must be started by calling the `start()` procedure.

Attempt to use an exclusive procedure when UDM_DISPUTE has not been 
started results in the `UninitializedModuleError` exception.

After using the module, the transaction should be closed
and the resources released by calling the `close()` procedure.

Version history:
----------------
1.0.20210526 - initial version
1.0.20210908 - removed 'srch_mask' parameter from 'close()' procedure and
			   any related logic
			 - added assertions as input check to public procedures
1.0.20220427 - fixed bug in 'modify_case_parameters()' when edit mode was
			   icorrectly identified as active following an error
			   during processing of a previous case.
1.0.20220504 - removed unused virtual key mapping from _vkeys{}
1.0.20231213 - the search_dispute() and search_disputes() procedures merged  
			   into a single search_disputes() procedure. The search_dispute()
			   procedure has been removed.
			   The export_disputes_data() procedure now returns the exported
			   dispute data as a plain text, while the export file is removed
			   when the export finishes. 

"""

import logging
import os
from os.path import exists, isfile, split, splitext
from typing import Union
from pyperclip import copy as copy_to_clipboard
from win32com.client import CDispatch

FilePath = str

_sess = None
_main_wnd = None
_stat_bar = None

log = logging.getLogger("master")


# keyboard to SAP virtual keys mapping
_virtual_keys = {
	"Enter":    0,
	"F3":       3,
	"F8":       8,
	"CtrlS":    11,
	"F12":      12,
	"ShiftF4":  16,
	"ShiftF12": 24
}


class DataExportError(Exception):
	"""Raised when data export fails."""

class FolderNotFoundError(Exception):
	"""Raised when a folder is reruired but doesn't exist."""

class LayoutNotFoundError(Exception):
	"""Raised when the used layout is not
	found in the list of available layouts.
	"""

class CasesNotFoundError(Exception):
	"""Raised when the number of records found
	is lower than the number of	cases searched.
	"""


def _press_key(name: str) -> None:
	"""Simulates pressing a keyboard button."""
	_main_wnd.SendVKey(_virtual_keys[name])

def _is_popup_dialog() -> bool:
	"""Checks if the active window is a popup dialog window."""
	return _sess.ActiveWindow.type == "GuiModalWindow"

def _close_popup_dialog(confirm: bool, max_attempts: int = 3) -> None:
	"""Confirms or declines a pop-up dialog."""

	nth = 0

	dialog_titles = (
		"Information",
		"Status check error",
		"Document lines: Display messages"
	)

	while _sess.ActiveWindow.text in (dialog_titles) and nth < max_attempts:
		if confirm:
			_press_key("Enter")
		else:
			_press_key("F12")

		nth += 1

	if _sess.ActiveWindow.type == "GuiModalWindow":
		dialog_title = _sess.ActiveWindow.text
		raise RuntimeError(f"Could not close the dialog window: '{dialog_title}'!")

	btn_caption = "Yes" if confirm else "No"

	for child in _sess.ActiveWindow.children:
		for grandchild in child.children:
			if grandchild.Type != "GuiButton":
				continue
			if btn_caption != grandchild.text.strip():
				continue
			grandchild.Press()
			return

def _get_grid_view() -> CDispatch:
	"""Returns a GuiGridView object representing
	the DMS window containing search results.
	"""

	splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
	grid_view = splitter_shell.FindAllByName("shell", "GuiGridView")(6)

	return grid_view

def _execute_query() -> int:
	"""Simulates pressing the 'Search' button
	located on the DMS main search mask.
	Returned is the number of cases found.
	"""

	splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
	qry_toolbar = splitter_shell.FindAllByName("shell", "GuiToolbarControl")(5)
	qry_toolbar.PressButton("DO_QUERY")

	num = _stat_bar.Text.split(" ")[0]
	num = num.strip().replace(".", "")

	return int(num)

def _find_and_click_node(tree: CDispatch, node: CDispatch, node_id: str) -> bool:
	"""Traverses the left-sided DMS menu tree to find the item with the given node ID.
	Once the item is found, the procedure simulates clicking on that item to open
	the corresponding subwindow.
	"""

	# find and double click the target root node
	if tree.IsFolder(node):
		tree.CollapseNode(node)
		tree.ExpandNode(node)

	# double clisk the target node
	if node.strip() == node_id:
		tree.DoubleClickNode(node)
		return True

	subnodes = tree.GetsubnodesCol(node)

	if subnodes is None:
		return False

	iter_subnodes = iter(subnodes)

	if _find_and_click_node(tree, next(iter_subnodes), node_id):
		return True

	try:
		next_node = next(iter_subnodes)
	except StopIteration:
		return False
	else:
		return _find_and_click_node(tree, next_node, node_id)

def _get_search_mask() -> CDispatch:
	"""Returns the GuiGridView object representing
	the DMS case search window.
	"""

	# find the target node by traversing the search tree
	tree_id = "shellcont/shell/shellcont[0]/shell/shellcont[1]/shell/shellcont[1]/shell"
	tree = _main_wnd.findById(tree_id)
	nodes = tree.GetNodesCol()
	iter_nodes = iter(nodes)
	clicked = _find_and_click_node(tree, next(iter_nodes), node_id = "4")

	assert clicked, "Target node not found!"

	# get reference to the search mask object found
	splitter_shell = _main_wnd.FindByName("shell", "GuiSplitterShell")
	srch_mask = splitter_shell.FindAllByName("shell", "GuiGridView")(4)

	return srch_mask

def _apply_layout(grid_view: CDispatch, name: str) -> None:
	"""Searches a layout by name in the DMS layouts list. If the layout is
	found in the list of available layouts, this gets selected.
	"""

	# Open Change Layout Dialog
	grid_view.PressToolbarContextButton("&MB_VARIANT")
	grid_view.SelectContextMenuItem("&LOAD")
	apo_grid = _sess.findById("wnd[1]").findAllByName("shell", "GuiShell")(0)

	for row_idx in range(0, apo_grid.RowCount):
		if apo_grid.GetCellValue(row_idx, "VARIANT") == name:
			apo_grid.setCurrentCell(str(row_idx), "TEXT")
			apo_grid.clickCurrentCell()
			return

	raise LayoutNotFoundError(f"Layout not found: {name}")

def _select_data_format(grid_view: CDispatch, idx: int) -> None:
	"""Selects the 'Unconverted' file format
	from file export format option window.
	"""

	grid_view.PressToolbarContextButton("&MB_EXPORT")
	grid_view.SelectContextMenuItem("&PC")
	option_wnd = _sess.FindById("wnd[1]")
	option_wnd.FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(idx).Select()

def _set_case_id(srch_mask: CDispatch, val: Union[str, int]) -> None:
	"""Enters case ID value into the corresponding field located on the search mask."""

	case_id_digits_count = 7

	if str(val).isnumeric() and len(str(val)) != case_id_digits_count:
		raise ValueError(f"Invalid case ID: {val}!")

	srch_mask.ModifyCell(0, "VALUE1", str(val))

def _set_hits_limit(srch_mask: CDispatch, n: int) -> None:
	"""Enters an iteger that restricts the he number of found records."""

	max_disputes = 5000

	if n > max_disputes:
		raise ValueError(f"Argument 'cases' cannot contain more than {max_disputes} cases!")

	if n == 0:
		raise ValueError("Argument 'cases' contains no case ID!")

	srch_mask.ModifyCell(23, "VALUE1", max_disputes)

def _copy_to_searchbox(srch_mask, cases: list) -> None:
	"""Copies case ID numbers intothe search listbox."""

	invalid_cases = []

	for case in cases:
		if not str(case).isnumeric():
			invalid_cases.append(case)

	if len(invalid_cases) != 0:
		vals = ';'.join(invalid_cases)
		raise ValueError(f"Argument 'cases' contains invalid value: {vals}")

	srch_mask.PressButton(0, "SEL_ICON1")
	cases = map(str, cases)
	_press_key("ShiftF4")       			# clear any previous values
	copy_to_clipboard("\r\n".join(cases))   # copy accounts to clipboard
	_press_key("ShiftF12")      			# confirm selection
	copy_to_clipboard("")                   # clear the clipboard content
	_press_key("F8")          				# confirm

def _export_to_file(grid_view: CDispatch, file: FilePath) -> None:
	"""Exports loaded accounting data to a text file."""

	folder_path, file_name = split(file)

	if not file_name.endswith(".txt"):
		ext = splitext(file)[1]
		_press_key("F3")
		raise ValueError(f"Invalid file type: '{ext}!")

	if not exists(folder_path):
		_press_key("F12")
		raise FolderNotFoundError(
			"The export folder not found at the "
			f"path specified: '{folder_path}'!")

	# select 'Unconverted' data format
	# and confirm the selection
	_select_data_format(grid_view, 0)
	_press_key("Enter")

	# enter data export file name, folder path and encoding
	# then click 'Replace' an existing file button
	if not folder_path.endswith("\\"):
		folder_path = "".join([folder_path, "\\"])

	# enter data export file name and folder path
	encoding_utf8 = "4120"

	_sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = folder_path
	_sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = file_name
	_sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = encoding_utf8

	# replace an exiting file
	_press_key("CtrlS")

	# double check if data export succeeded
	if not isfile(file):
		raise DataExportError(f"Failed to export data to file: {file}")

def _read_exported_data(file_path: str) -> str:
	"""Reads exported FBL5N data from the text file."""

	with open(file_path, encoding = "utf-8") as stream:
		text = stream.read()

	return text

def start(sess: CDispatch) -> None:
	"""Starts the UDM_DISPUTE transaction.

	If the FBL5N has already been started,
	then the transaction will be restarted.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object (wrapped in
		the win32:CDispatch class)that represents
		an active user SAP GUI _sess.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if sess is None:
		raise UnboundLocalError("Argument 'sess' is unbound!")

	# close the transaction
	# if it is already running
	close()

	_sess = sess
	_main_wnd = _sess.findById("wnd[0]")
	_stat_bar = _main_wnd.findById("sbar")

	_sess.StartTransaction("UDM_DISPUTE")

	return _get_search_mask()

def close() -> None:
	"""Closes a running UDM_DISPUTE transaction.

	Attempt to close the transaction that has not been
	started by the `start()` procedure is ignored.
	"""

	global _sess
	global _main_wnd
	global _stat_bar

	if _sess is None:
		return

	_sess.EndTransaction()

	if _is_popup_dialog():
		_close_popup_dialog(confirm = True)

	_sess = None
	_main_wnd = None
	_stat_bar = None

def search_dispute(search_mask: CDispatch, case: int) -> Union[CDispatch, None]:
	"""Searches a disputed case in DMS based on database ID.

	Parameters:
	-----------
	search_mask:
		An SAP `GuiGridView` object that
		represents the DMS search window.

	case:
		ID number of the searched case.

	Returns:
	--------
	If a record is found, then a `GuiGridView` object
	that contains the search results is returned.

	If no record is found, then `None` is returned.
	"""

	_set_case_id(search_mask, case)
	n_found = _execute_query()

	if n_found == 0:
		return None

	return _get_grid_view()

def search_disputes(search_mask: CDispatch, case: Union[int, str, list]) -> CDispatch:
	"""Searches a disputed case in DMS based on database ID.

	If the number of records found is lower than the number of
	cases searched, then `CasesNotFoundError` exception is raised.

	Parameters:
	-----------
	search_mask:
		An SAP `GuiGridView` object that
		represents the DMS search window.

	case:
		An ID number of the searched case, stored as
		an `int` or `str`, or a list of these numbers.

	Returns:
	--------
	A `GuiGridView` object that represents the
	DMS table with the search results.
	"""

	if isinstance(case, (int, str)):
		_set_case_id(search_mask, case)
		n_found = _execute_query()

		if n_found == 0:
			return None

		return _get_grid_view()

	n_total = len(case)

	# hit limit should be equal to the num of cases
	_set_hits_limit(search_mask, n_total)

	# search cases
	_copy_to_searchbox(search_mask, case)
	n_found = _execute_query()

	if n_found > 0:
		search_results = _get_grid_view()
	else:
		search_results = None

	if 0 < n_found < n_total:
		raise CasesNotFoundError(
			f"Incorrect disputes detected: {n_total - n_found}. "
			"There might be a typo in the case ID provided in item 'Text' value(s).")

	return search_results

def export_disputes_data(search_result: CDispatch, file: FilePath, layout: str) -> str:
	"""Exports a case search result.

	Parameters:
	-----------
	search_result:
		A `GuiGridView` object that contains case search result.

	file:
		Path to a temporary .txt file to which the data will be exported.

		The file is removed when the data reading is complete.

		If the file path points to an invalid folder, \n
		then a `FolderNotFoundError` exception is raised.

		A `DataExportError` exception is raised
		if the attempt to export accounting data fails.

	layout:
		The name of the layout that defines the column format of the exported data.

		By default, no specific layout name is used,
		and a the predefined FLB3N layout is used.

		A `LayoutNotFoundError` exception is raised when the used
		layout is not found in the list of the available layouts.

	Returns:
	--------
	The exported dispute data as plain text.
	"""

	_apply_layout(search_result, layout)
	_export_to_file(search_result, file)

	data = _read_exported_data(file)

	try:
		os.remove(file)
	except (PermissionError, FileNotFoundError) as exc:
		log.error(exc)

	return data
