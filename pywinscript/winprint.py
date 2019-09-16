

import win32api
import win32print



class Printer(object):
	"""Represents a physical printer.

	Parameters
	----------
	document : str
		Absolute path to printable document.
	papersize : {'A', 'B', 'D'}, optional
	copies : int, optional

	Attributes
	----------
	document
	printer
	papersize
	copies

	"""

	_PROPERTIES = {
		'A' : {
			'constant' : 1, 
			'inches' : '8.5 X 11', 
			'printer': r'\\sn0085piprn001\PRSUSHN022'
		},
		'B' : {
			'constant' : 17, 
			'inches' : '11 X 17', 
			'printer': r'\\sn0085piprn001\PRSUSHN022'
		},
		'D' : {
			'constant' : 25, 
			'inches' : '24 X 36', 
			'printer': r'\\sn0085piprn001\PRSUSHN023'
		},
	}

	# The amount of printer settings that are returned when calling
	# win32print.GetPrinter().
	_LEVEL = 2

	def __init__(self, document, papersize='A', copies=1):
		self._document = document
		self._papersize = papersize
		self._copies = copies

	@property
	def document(self):
		"""str: The absolute path to a printable document."""
		return self._document
	
	@document.setter
	def document(self, filepath):
		self._document = filepath

	@property
	def printer(self):
		"""str: The active printer."""
		return self._PROPERTIES[self._papersize]['printer']

	@property
	def papersize(self):
		"""str: The papersize to be printed."""
		return self._PROPERTIES[self._papersize]['constant']

	@papersize.setter
	def papersize(self, size):
		self._papersize = size

	@property
	def copies(self):
		"""int: The number of copies to be printed."""
		return self._copies
	
	@copies.setter
	def copies(self, copies):
		self._copies = copies

	def start(self):
		"""Print a document on the active printer."""
		# Set printer
		win32print.SetDefaultPrinter(self.printer)
		handle = win32print.OpenPrinter(self.printer, None)

		# Get existing printer settings
		settings = win32print.GetPrinter(handle, self._LEVEL)
		settings['pDevMode'].PaperSize = self.papersize
		settings['pDevMode'].Copies = self.copies

		# Update printer settings
		# Exceptions are raised, but changes are applied.
		try:
			win32print.SetPrinter(handle, self._LEVEL, settings, 0)
		except:
			pass

		win32api.ShellExecute(0, 'print', self.document, None, '.', 0)
		win32print.ClosePrinter(handle)


if __name__ == '__main__':
	printer = Printer('C:\\Users\\mcclbra\\Desktop\\1.docx')
	printer.start()

