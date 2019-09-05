#!/usr/bin/env python
# -*- coding: utf-8 -*-


import subprocess
from os import startfile

__author__ = 'Brandon McCleary'
__version__ = '1.0.1'
__maintainer__ = 'Brandon McCleary'
__email__ = 'brandon.shane.mccleary@gmail.com'


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


