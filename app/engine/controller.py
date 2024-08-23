
"""
Description:
------------
The controller.py represents the middle layer of the application design
and mediates communication between the top layer (app.py) and the highly
specialized modules situated on the bottom layer of the design (mails.py,
report.py, processor.py, sap.py, dms.py fbl5n.py).

Version history:
----------------
1.0.20210720 - Initial version.
"""

import logging
import os
import re
from datetime import date
from datetime import datetime as dt
from datetime import timedelta
from glob import glob
from logging import Logger, config
from os.path import basename, isfile, join
from typing import Union
import yaml
from pandas import DataFrame, Series
from win32com.client import CDispatch

from . import dms, fbl5n, mails, processor, report, sap

log = logging.getLogger("master")


# ====================================
# initialization of the logging system
# ====================================

def _compile_log_path(log_dir: str) -> str:
	"""Compiles the path to the log file
	by generating a log file name and then
	concatenating it to the specified log
	directory path."""

	date_tag = dt.now().strftime("%Y-%m-%d")
	nth = 0

	while True:
		nth += 1
		nth_file = str(nth).zfill(3)
		log_name = f"{date_tag}_{nth_file}.log"
		log_path = join(log_dir, log_name)

		if not isfile(log_path):
			break

	return log_path

def _read_log_config(cfg_path: str) -> dict:
	"""Reads logging configuration parameters from a yaml file."""

	# Load the logging configuration from an external file
	# and configure the logging using the loaded parameters.

	if not isfile(cfg_path):
		raise FileNotFoundError(f"The logging configuration file not found: '{cfg_path}'")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	return yaml.safe_load(content)

def _update_log_filehandler(log_path: str, logger: Logger) -> None:
	"""Changes the log path of a logger file handler."""

	prev_file_handler = logger.handlers.pop(1)
	new_file_handler = logging.FileHandler(log_path)
	new_file_handler.setFormatter(prev_file_handler.formatter)
	logger.addHandler(new_file_handler)

def _print_log_header(logger: Logger, header: list, terminate: str = "\n") -> None:
	"""Prints header to a log file."""

	for nth, line in enumerate(header, start = 1):
		if nth == len(header):
			line = f"{line}{terminate}"
		logger.info(line)

def _remove_old_logs(logger: Logger, log_dir: str, n_days: int) -> None:
	"""Removes old logs older than the specified number of days."""

	old_logs = glob(join(log_dir, "*.log"))
	n_days = max(1, n_days)
	curr_date = dt.now().date()

	for log_file in old_logs:
		log_name = basename(log_file)
		date_token = log_name.split("_")[0]
		log_date = dt.strptime(date_token, "%Y-%m-%d").date()
		thresh_date = curr_date - timedelta(days = n_days)

		if log_date < thresh_date:
			try:
				logger.info(f"Removing obsolete log file: '{log_file}' ...")
				os.remove(log_file)
			except PermissionError as exc:
				logger.error(str(exc))

def configure_logger(log_dir: str, cfg_path: str, *header: str) -> None:
	"""Configures application logging system.

	Parameters:
	-----------
	log_dir:
		Path to the directory to store the log file.

	cfg_path:
		Path to a yaml/yml file that contains
		application configuration parameters.

	header:
		A sequence of lines to print into the log header.
	"""

	log_path = _compile_log_path(log_dir)
	log_cfg = _read_log_config(cfg_path)
	config.dictConfig(log_cfg)
	logger = logging.getLogger("master")
	_update_log_filehandler(log_path, logger)
	if header is not None:
		_print_log_header(logger, list(header))
	_remove_old_logs(logger, log_dir, log_cfg.get("retain_logs_days", 1))


# ====================================
# 		application configuration
# ====================================

def load_app_config(cfg_path: str) -> dict:
	"""Reads application configuration
	parameters from a file.

	Parameters:
	-----------
	cfg_path:
		Path to a yaml/yml file that contains
		application configuration parameters.

	Returns:
	--------
	Application configuration parameters.
	"""

	log.info("Loading application configuration ...")

	if not cfg_path.endswith((".yaml", ".yml")):
		raise ValueError("The configuration file not a YAML/YML type!")

	with open(cfg_path, encoding = "utf-8") as stream:
		content = stream.read()

	cfg = yaml.safe_load(content)
	log.info("Configuration loaded.")

	return cfg

def load_processing_rules(file_path: str) -> dict:
	"""Loads customer-specific parameters for data evaluation.

	Parameters:
	-----------
	file_path:
		Path to the file containing the processing rules.

	Returns:
	--------
	Data evaluation parameters.
	"""

	log.info("Loading data evaluation rules ...")
	with open(file_path, encoding = "utf-8") as stream:
		content = stream.read()

	rules = yaml.safe_load(content)
	log.info("Rules loaded.")

	return rules

# ====================================
# 		Management of SAP connection
# ====================================

def connect_to_sap(system: str) -> CDispatch:
	"""Creates connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	system:
		The SAP system to use for connecting to the scripting engine.

	Returns:
	--------
	An SAP `GuiSession` object that represents active user session.
	"""

	log.info("Connecting to SAP ...")
	sess = sap.connect(system)
	log.info("Connection created.")

	return sess

def disconnect_from_sap(sess: CDispatch) -> None:
	"""Closes connection to the SAP GUI scripting engine.

	Parameters:
	-----------
	sess:
		An SAP `GuiSession` object (wrapped in the `win32:CDispatch` class)
		that represents an active user SAP GUI session.
	"""

	log.info("Disconnecting from SAP ...")
	sap.disconnect(sess)
	log.info("Connection to SAP closed.")


# ====================================
# 		Fetching of user input
# ====================================

def fetch_user_input(msg_cfg: dict, email_id: str) -> tuple:
	"""Fetches the processing parameters and data provided by the user.

	Parameters:
	-----------
	msg_cfg:
		Application 'messages' configuration parameters.

	email_id:
		The string ID of the message.

	Returns:
	--------
	Names of the processing parameters and their values:
	- "error_message": `str` Default: ""
		Message that details any errors in user email.
	- "message_type": `str` Default: "I"
		Type of the message:
			- "I": The message is an information.
			- "W": The message is a warning.
			- "E": The message is an error.
	- "email": `str` Default: ""
		Email address of the sender.
	- "entity": `str` Default: ""
		Entity for which the overdue report is
		created (a worklist or a company code).
	- "overdue_day": `date` Default: None
		Date for which the open line items are exported from FBL5N.
	"""

	log.info("Retrieving user message ...")

	params = {
		"error_message": "",
		"message_type": "I",
		"email": "",
		"entity": "",
		"overdue_day": None,
	}

	acc = mails.get_account(
		msg_cfg["requests"]["mailbox"],
		msg_cfg["requests"]["account"],
		msg_cfg["requests"]["server"]
	)

	messages = mails.get_messages(acc, email_id)

	if len(messages) == 0:
		raise RuntimeError(
			f"Could not find message with the specified ID: '{email_id}'")

	msg = messages[0]
	log.info("User message retrieved.")

	log.info("Extracting relevant contents from the user message ...")
	params.update({"email": msg.sender.email_address})
	params_pattern = r"(?P<entity>\w+)/(?P<date>\d{2}.\d{2}.\d{4})"
	match = re.search(params_pattern, msg.text_body, re.M)

	if match is None:
		log.error("Could not recognize user parameter(s) in messsage text.")
		params.update({"email": "The parameters you have entered are invalid!"})
		return params

	entity = match.group("entity").upper()
	exp_date = match.group("date")
	exp_date = dt.strptime(exp_date, "%d.%m.%Y").date()

	params.update({"entity": entity, "overdue_day": exp_date})

	return params

# ====================================
# 			Data export
# ====================================

def fetch_fbl5n_data(
		temp_dir: str, entity: str,
		overdue_day: str, rules: dict,
		data_cfg: dict, session: CDispatch
	) -> DataFrame:
	"""Fetches open line items from the FBL5N transaction.

	First, the data is exported form the FBL5N as raw text.
	Then, the FBL5N text data is parsed into a DataFrame object.

	Parameters:
	-----------
	temp_dir:
		Path to the directory where temporary files are stored.

	entity:
		Entity for which the open line items are exported.

		Can be either a company code or a worklist.

	overdue_day:
		Day for which the open line items are exported.

	rules:
		Entity-specific data processing rules.

	data_cfg:
		Application 'data' configuration parameters.

	session:
		An SAP `GuiSession` object.

	Returns:
	--------
	The FBL5N accounting data.
	"""

	exp_path = join(temp_dir, "fbl5n_export.txt")

	log.info("Starting FBL5N ...")
	fbl5n.start(session)
	log.info("FBL5N started.")

	try:
		log.info("Exporting FBL5N data ...")
		if rules["type"] == "country":
			data = fbl5n.export_line_items(
				exp_path, rules["company_code"],
				from_day = overdue_day,
				layout = data_cfg["fbl5n_layout"])
		elif rules["type"] == "worklist":
			data = fbl5n.export_line_items(
				exp_path, rules["company_code"],
				entity, from_day = overdue_day,
				layout = data_cfg["fbl5n_layout"])
		log.info("FBL5N data successfully exported.")
	except Exception as exc:
		raise RuntimeError(str(exc)) from exc
	finally:
		log.info("Closing FBL5N ...")
		fbl5n.close()
		log.info("FBL5N closed.")

	converted = processor.convert_fbl5n_data(data, rules["case_id_rx"])
	assert not converted.empty, "Data conversion failed!"

	return converted

def fetch_dms_data(
		temp_dir: str, data_cfg: dict,
		cases: Series, session: CDispatch
	) -> DataFrame:
	"""Fetches data for disputed cases
	from the UDM_DISPUTE transaction.

	Parameters:
	-----------
	temp_dir:
		Path to the directory where
		temporary files are stored.

	data_cfg:
		Application 'data' configuration parameters.

	cases:
		Case ID numbers extracted from
		the "Text" field of the FBL5N data.

	session:
		An SAP `GuiSession` object.

	Returns:
	--------
	The DMS cases data.
	"""

	exp_path = join(temp_dir, "dms_export.txt")
	cases = cases.dropna().unique()

	log.info("Starting UDM_DISPUTE ...")
	search_mask = dms.start(session)
	log.info("UDM_DISPUTE started.")

	try:
		log.info(f"Searching {len(cases)} cases ...")
		search_result = dms.search_disputes(search_mask, list(cases))
		log.info("Exportig disputes data from DMS ...")
		data = dms.export_disputes_data(search_result, exp_path, data_cfg["dms_layout"])
		log.info("DMS data successfully exported.")
	except Exception as exc:
		raise RuntimeError(str(exc)) from exc
	finally:
		log.info("Closing UDM_DISPUTE ...")
		dms.close()
		log.info("UDM_DISPUTE closed.")

	converted = processor.convert_dms_data(data)
	assert not converted.empty, "Data conversion failed!"

	return converted


# ====================================
# 			Data evaluaton
# ====================================

def evaluate_data(
		fbl5n_data: DataFrame, dms_data: DataFrame,
		data_dir: str, entity: str, rules: dict,
	) -> DataFrame:
	"""Evaluates the FBL5N and DMS data.

	Parameters:
	-----------
	fbl5n_data:
		The FBL5N accounting data.

	dms_data:
		The DMS dispute data.

	data_dir:
		Path to the application directory where accounting
		data used for the evaluation is stored.

	entity:
		Entity for which the open line items are exported.

	rules:
		Entity-specific data processing rules.

	Returns:
	--------
	The result of the data evaluation.
	"""

	entity_subdir = "_".join([entity, rules["company_code"]])

	if entity == "OBI" and rules["company_code"] == "1001":

		info_data_dir = join(data_dir, entity_subdir)
		info_data_paths = glob(join(info_data_dir, "*.xlsx"))

		evaluated = processor.evaluate_obi_de(
			fbl5n_data, dms_data, rules["queries"], info_data_paths)

	elif entity == "AUSTRIA":

		customer_data_path = join(
			data_dir, entity_subdir,"customers.xlsx")

		evaluated = processor.evaluate_austria(
			fbl5n_data, dms_data,
			customer_data_path,
			rules["queries"])

	else:
		raise NotImplementedError(
			"No data evaluation proedure has been "
			f"implemented for entity '{entity}'!")

	return evaluated


# ====================================
# 			Reporting
# ====================================

def create_report(
		data: DataFrame,
		report_cfg: dict,
		entity: str,
		rules: dict,
		overdue_day: date,
		temp_dir: str
	) -> str:
	"""Creates user report from the processing result.

	Parameters:
	-----------

	data:
		The evaluation result from which report will be generated.

	data_cfg:
		Application 'data' configuration parameters.

	entity:
		Entity for which the open line items are exported.

	rules:
		Entity-specific data processing rules.

	overdue_day:
		Day for which the open line items are exported.

	data_dir:
		The application directory where additional
		accounting info files are stored, which are
		necessary for creating pivot tables in the
		user report

	temp_dir:
		Path to the directory where temporary files are stored.

	Returns:
	--------
	Path to the report file.
	"""

	log.info("Creating user report ...")

	company_code = rules["company_code"]
	report_date = overdue_day.strftime("%d%b%Y")
	report_name = report_cfg["report_name"]

	report_name = report_name.replace("$entity$", entity)
	report_name = report_name.replace("$company_code$", company_code)
	report_name = report_name.replace("$date$", report_date)

	report_path = join(temp_dir, report_name)

	report_fields = rules["report_fields"]
	sheet_names = rules["report_sheets"]

	if entity == "AUSTRIA":
		report.create_report_austria(
			report_path, data,
			report_fields, sheet_names)
	elif entity == "OBI" and company_code == "1001":
		report.create_report_obi_de(
			report_path, data,
			report_fields, sheet_names)
	else:
		raise NotImplementedError(
			"No report creation procedure has been "
			f"implemented for '{entity} {company_code}'!")

	return report_path

def send_notification(
		msg_cfg: dict,
		user_mail: str,
		template_dir: str,
		attachment: Union[dict, str] = None,
		error_msg: str = ""
	) -> None:
	"""Sends a notification with processing result to the user.

	Parameters:
	-----------
	msg_cfg:
		Application 'messages' configuration parameters.

	user_mail:
		Email address of the user who requested processing.

	template_dir:
		Path to the application directory
		that contains notification templates.

	attachment:
		Attachment name and data or a file path.

	error_msg:
		Error message that will be included in the user notification.
		By default, no erro message is included.
	"""

	log.info("Sending notification to user ...")

	notif_cfg = msg_cfg["notifications"]

	if not notif_cfg["send"]:
		log.warning(
			"Sending of notifications to users "
			"is disabled in 'app_config.yaml'.")
		return

	if error_msg != "":
		templ_name = "template_error.html"
	else:
		templ_name = "template_completed.html"

	templ_path = join(template_dir, templ_name)

	with open(templ_path, encoding = "utf-8") as stream:
		html_body = stream.read()

	if error_msg != "":
		html_body = html_body.replace("$error_msg$", error_msg)

	if attachment is None:
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body
		)
	elif isinstance(attachment, dict):
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body,
			{attachment["name"]: attachment["content"]}
		)
	elif isinstance(attachment, str):
		msg = mails.create_smtp_message(
			notif_cfg["sender"], user_mail,
			notif_cfg["subject"], html_body,
			attachment
		)
	else:
		raise ValueError(f"Unsupported data type: '{type(attachment)}'!")

	try:
		mails.send_smtp_message(msg, notif_cfg["host"], notif_cfg["port"])
	except Exception as exc:
		log.error(exc)
		return

	log.info("Notification sent.")


# ====================================
# 			Data cleanup
# ====================================

def delete_temp_files(temp_dir: str) -> None:
	"""Removes all temporary files.

	Parameters:
	-----------
	temp_dir:
		Path to the directory where temporary files are stored.
	"""

	file_paths = glob(join(temp_dir, "*.*"))

	if len(file_paths) == 0:
		log.warning("No temporary files to remove detected.")
		return

	log.info("Removing temporary files ...")

	for file_path in file_paths:
		try:
			os.remove(file_path)
		except Exception as exc:
			log.exception(exc)

	log.info("Files successfully removed.")
