"""
pywinscript.polyworks provides an interface to Innovmetric's PolyWorks 
Inspector.

"""
import os
import comtypes.client
from win import start
from constants import POLYWORKS, POLYWORKS_PATH, INSPECTOR


class Polyworks(object):
	"""
	An interface to native PolyWorks commands.

	Attributes
	----------
	ROOT : str
		Absolute path to PolyWorks application files.
	CLSID : str
		COM object identifier.
	PW_PATH : str
		Absolute path to polyworks.exe.
	INSPECTOR_PATH : str
		Absolute path to iminspect.exe.
	MODULES : dict
	inspector : IIMInspect.CommandCenterCreate
		A connection to an active project.

	"""

	CLSID = '{23E630FD-A4D5-47AD-A3B2-7D2779FFF888}'
	PW_PATH = os.path.join(POLYWORKS_PATH, POLYWORKS)
	INSPECTOR_PATH = os.path.join(POLYWORKS_PATH, INSPECTOR)
	MODULES = {
		'polyworks' : PW_PATH,
		'inspector' : INSPECTOR_PATH
		}

	def __init__(self):
		pass

	def start_exe(self, module):
		"""Launch EXE if not already running.

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
		otherwise a new project is created. PolyWorks will handle missing dongle 
		errors internally so no errors will be raised for that issue.

		Returns
		-------
		inspector : IIMInspect.CommandCenterCreate
			A connection to current project.

		Raises
		------
		OSError
			Error loading type library/DLL (bad path)
		WindowsError
			Server execution failed

		"""
		module = comtypes.client.GetModule([self.CLSID, 1, 0])
		init = comtypes.client.CreateObject('InnovMetric.PolyWorks.IMInspect')
		project = init.QueryInterface(module.IIMInspect).ProjectGetCurrent()
		self.inspector = project.CommandCenterCreate()
		return self.inspector


if __name__ == '__main__':
	polyworks = Polyworks()
	polyworks.connect_to_inspector()
	# polyworks.start_exe('inspector')
