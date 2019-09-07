#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import win32api
import win32con
import pythoncom
import comtypes.client
import win32com.client as win32
from pywintypes import com_error
from win import start
from constants import POLYWORKS, INSPECTOR


class Polyworks(object):
	"""
	Provides an interface to InnovMetric PolyWorks commands.

	"""
	# TODO
	# Consider making ROOT a list of paths. Since IT might install things 
	# differently from person to person, this might assist the program in 
	# finding a working path in all cases.
	ROOT = 'C:\\Program Files\\InnovMetric\\Polyworks MS 2018\\bin'
	POLYWORKS_PATH = os.path.join(ROOT, POLYWORKS)
	INSPECTOR_PATH = os.path.join(ROOT, INSPECTOR)
	MODULES = {
		'polyworks' : POLYWORKS_PATH,
		'inspector' : INSPECTOR_PATH
		}

	def __init__(self):
		pass

	def start_exe(self, module):
		"""
		Parameters
		----------
		module : {'polyworks', 'inspector'}
		
		"""
		start(self.MODULES[module])

	def connect_to_inspector(self):
		"""Connect to an instance of PolyWorks Inspector.

		This connection allows the passing of commands from a python program to
		Inspector in the form of PolyWorks' own macro scripting language.

		If inspector is already open this method will grab the current project, 
		otherwise a new project is created. PolyWorks will handle missing 
		dongle errors internally so no errors will be raised for that issue.

		Returns
		-------
		inspector : project.CommandCenterCreate
			Connection to the current project.

		Raises
		------
		OSError
			If 'Error loading type library/DLL' (bad path).

		"""
		key = win32api.RegOpenKeyEx(
			win32con.HKEY_CLASSES_ROOT, 
			'InnovMetric.PolyWorks.IMInspect', 
			0, 
			win32con.KEY_READ
		)
		module = comtypes.client.GetModule(self.INSPECTOR_PATH)
		init = comtypes.client.CreateObject(
			win32api.RegQueryValue(
				key, 
				'CLSID'
			)
		)
		project = init.QueryInterface(module.IIMInspect).ProjectGetCurrent()
		return project.CommandCenterCreate()




