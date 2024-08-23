# pylint: disable = E0611, W0603

"""
Description:
------------
The module provides an interface for managing
connection to the SAP GUI Scripting Engine.

Version history:
----------------
1.0.20231122: Initial version.
"""

from os.path import isfile
from subprocess import Popen, TimeoutExpired

import win32com.client
from win32com.client import CDispatch
from win32ui import FindWindow
from win32ui import error as WinError

FilePath = str

_systems = {
	"P25": "OG ERP: P25 Productive SSO",
	"Q25": "OG ERP: Q25 Quality Assurance SSO"
}

class SapConnectionError(Exception):
	"""Raised when attempt to connect to the scripting engine fails."""

def connect(system: str, exe: FilePath="") -> CDispatch:
	"""Connects to the SAP GUI Scripting Engine.

	A `SapConnectionError` exception is raised when
	logging into the SAP GUI Scripting Engine fails.

	Parameters:
	-----------
	system:
		SAP system with the GUI Scripting Engine to connect:
		- "P25": Productive SSO.
		- "Q25": Quality Assurance SSO.

	exe:
		Path to the local SAP GUI executable.

		If the path is not specified, then the default local SAP
		installation directory is searched for the executable file.
		If the executable is not found, then a `FileNotFoundError`
		exception is raised.

	Returns:
	--------
	An SAP `GuiSession` context object representing
	an active session for running transactions.
	"""

	# SAP should always be installed to the same directory for all users
	# If a specific path is used, then the default path is overridden.

	default_exe = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
	exe_path = default_exe if exe == "" else exe

	if not isfile(exe_path):
		raise FileNotFoundError(
			"SAP GUI executable not found at "
			f"the specified path: {exe_path}!")

	if system.upper() not in _systems:
		raise ValueError(f"Unrecognized SAP system to connect: {system}!")

	try:
		FindWindow("", "SAP Logon 750")
	except WinError:
		try:
			proc = Popen(exe_path)
			proc.communicate(timeout = 8)
		except TimeoutExpired:
			pass # does not impact getting a SapGui reference in next step
		except Exception as exc:
			raise SapConnectionError("Communication with the process failed!") from exc

	try:
		sap_gui_auto = win32com.client.GetObject("SAPGUI")
	except Exception as exc:
		raise SapConnectionError("Could not get the 'SAPGUI' object.") from exc

	engine = sap_gui_auto.GetScriptingEngine

	if engine.Connections.Count == 0:
		sys_name = _systems[system.upper()]
		engine.OpenConnection(sys_name, Sync = True)

	conn = engine.Connections(0)

	return conn.Sessions(0)

def disconnect(sess: CDispatch) -> None:
	"""Disconnects from the SAP GUI Scripting Engine.

	Parameteres:
	------------
	sess:
		An SAP `GuiSession` object.
	"""

	if sess is None:
		raise UnboundLocalError("Uninitialized session used!")

	conn = sess.Parent
	conn.CloseSession(sess.ID)
	conn.CloseConnection()
