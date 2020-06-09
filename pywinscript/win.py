"""
pywinscript.win is a high-level OS wrapper for Windows desktop that starts 
programs and creates directories without raising unnecessary exceptions. The 
intent of the function calls are considered when determining the legitimacy of 
an exception.

"""
import subprocess
from os.path import join as osjoin
from os import startfile, mkdir


def start(program, path=None):
	"""Start a program if not already running.

	Parameters
	----------
	program : str
		Name ending with '.exe'.
	path : str or None
		Absolute path to the EXE, optional.

	Raises
	------
	WindowsError
                The system cannot find the file specified.

	"""
	if not is_running(program):
		if path is not None:
			program = osjoin(path, program)
		startfile(program)


def is_running(program):
	"""Returns True if a program is running.

	Parameters
	----------
	program : str
		Executable program ending with '.exe'.

	"""
	if program in subprocess.check_output('tasklist', shell=True):
		return True


def create_folder(path):
	"""Create a new folder if it does not already exist.

	Parameters
	----------
	path : str
		Absolute path to destination folder.

	Returns
	-------
	path : str
		Absolute path to destination folder.

	Raises
	------
	WindowsError
		If there is a missing link in `path`.
	
	"""
	try:
		mkdir(path)
	except WindowsError as error:
		if error.winerror == 183:
            # Folder already exists
			return path
		else:
			raise error
	else:
		return path