

import subprocess
from os import startfile


def start(program):
	"""Start a given program if not already running.

	Parameters
	----------
	program : str
		Executable program ending with '.exe'

	"""
	if not is_running(program):
		startfile(program)


def is_running(program):
	"""
	Parameters
	----------
	program : str
		Executable program ending with '.exe'

	Returns
	-------
	True
		`program` is running
	
	"""
	if program in subprocess.check_output('tasklist', shell=True):
		return True


