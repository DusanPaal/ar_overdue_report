# pylint: disable = C0103, W1203

"""
The "AR Overdue Report" application automates
generating of excel reports that summarize
overdue items located on customer accounts
for sake of monthly review by AR department.
"""

import argparse
import logging
import sys
import datetime as dt
from os.path import join
from engine import controller

log = logging.getLogger("master")

def main(**args) -> int:
	"""Program entry point.

	Controls the overall execution of
	the core component of the application.

	Parameters (args):
	------------------
	- "email_id":
		The string ID of the user message
		that triggers the application. 
		                
	Returns:
	--------
	Program completion state:
		- 0: Program successfully completes.
		- 1: Program fails during the initialization phase.
		- 2: Program fails during the user input fetch phase.
		- 3: Program fails during the processing phase.
		- 4: Program fails during the reporting phase.
	"""

	app_dir = sys.path[0]
	log_dir = join(app_dir, "logs")
	temp_dir = join(app_dir, "temp")
	data_dir = join(app_dir, "data")
	template_dir = join(app_dir, "notification")
	rules_path = join(app_dir, "rules.yaml")
	app_cfg_path = join(app_dir, "app_config.yaml")
	log_cfg_path = join(app_dir, "log_config.yaml")
	curr_date = dt.datetime.now().strftime("%d-%b-%Y")

	try:
		controller.configure_logger(
			log_dir, log_cfg_path,
			"Application name: AR Overdue Report",
			"Application version: 1.0.20220721",
			f"Log date: {curr_date}")
	except Exception as exc:
		print(exc)
		print(
			"CRITICAL: Unhandled exception while "
			"trying to configuring the logging system!")
		return 1

	try:
		log.info("=== Initialization START ===")
		cfg = controller.load_app_config(app_cfg_path)
		ruleset = controller.load_processing_rules(rules_path)
		session = controller.connect_to_sap(cfg["sap"]["system"])
		log.info("=== Initialization END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("Unhandled exception while initializing the application!")
		return 1

	try:
		log.info("=== Fetching user input START ===")
		user_input = controller.fetch_user_input(
			cfg["messages"], args["email_id"])
		log.info("=== Fetching user input END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("=== Fetching user input FAILURE ===\n")
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(session)
		log.info("=== Cleanup END ===\n")
		return 2

	if user_input["message_type"] == "E":
		log.error(user_input["error_message"])
		controller.send_notification(
			cfg["messages"], user_input["email"], template_dir,
			error_msg = user_input["error_message"])
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(session)
		log.info("=== Cleanup END ===\n")
		return 2
	  
	if not user_input["entity"] in ruleset:
		log.error(user_input["error_message"])
		controller.send_notification(
			cfg["messages"], user_input["email"], template_dir,
			error_msg = "The entity you have entered is not valid!")
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(session)
		log.info("=== Cleanup END ===\n")
		return 2

	log.info("=== Processing START ===")
	rules = ruleset[user_input["entity"]]

	try:

		fbl5n_data = controller.fetch_fbl5n_data(
			temp_dir, user_input["entity"],
			user_input["overdue_day"], rules,
			cfg["data"], session)

		dms_data = controller.fetch_dms_data(
			temp_dir, cfg["data"],
			fbl5n_data["Case_ID"], session)

		evaluated = controller.evaluate_data(
			fbl5n_data, dms_data, data_dir,
			user_input["entity"], rules)

	except Exception as exc:
		log.exception(exc)
		log.info("=== Processing FAILURE ===\n")
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(session)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")
		return 3

	log.info("=== Processing END ===\n")

	try:
		log.info("=== Reporting START ===")

		report_path = controller.create_report(
			evaluated, cfg["report"],user_input["entity"],
			rules, user_input["overdue_day"], temp_dir)

		controller.send_notification(
			cfg["messages"], user_input["email"],
			template_dir, attachment = report_path)

		log.info("=== Reporting END ===\n")
	except Exception as exc:
		log.exception(exc)
		log.critical("=== Reporting FAILURE ===\n")
		return 4
	finally:
		log.info("=== Cleanup START ===")
		controller.disconnect_from_sap(session)
		controller.delete_temp_files(temp_dir)
		log.info("=== Cleanup END ===\n")

	return 0


if __name__ == "__main__":
	parser = argparse.ArgumentParser()
	parser.add_argument("-e", "--email_id", required = False, help = "Sender email id.")
	arguments = vars(parser.parse_args())
	exit_code = main(email_id = arguments["email_id"])
	log.info(f"=== System shutdown with return code: {exit_code} ===")
	logging.shutdown()
	sys.exit(exit_code)
