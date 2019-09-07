#!/usr/bin/env python
# -*- coding: utf-8 -*-

import subprocess
from os import startfile


def start(program):
	"""Start an application if not already running.

	Parameters
	----------
	program : str
		Executable program ending with '.exe'

	"""
	if not is_running(program):
		startfile(program)


def is_running(program):
	"""Verify that an application is running.

	Parameters
	----------
	program : str
		Executable program ending with '.exe'

	Returns
	-------
	True
		If `program` is running.
	
	"""
	if program in subprocess.check_output('tasklist', shell=True):
		return True


