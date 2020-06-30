"""
pywinscript.autocad uses the pyautocad API to satisfy specific automation 
requirements.

"""
from comtypes import COMError
from constants import AUTOCAD, AUTOCAD_PATH
from win import start, is_running
from pyautocad import Autocad, ACAD, APoint


class AutoCAD(Autocad):
	"""
	Provides automation functions for Autodesk AutoCAD.

	Attributes
	----------
	layouts : list
	layout_names : list

	"""
	def __init__(self):
		super(AutoCAD, self).__init__()
		start(AUTOCAD, AUTOCAD_PATH)

	@property
	def layouts(self):
		"""list: The layout objects from the active document."""
		return [layout for layout in self.iter_layouts(doc=self.doc)]

	@property
	def layout_names(self):
		"""list: The layout names from the active document."""
		return [layout.Name for layout in self.layouts]	

	def set_layout(self, name):
		"""Activate a document layout tab.

		Parameters
		----------
		name : str

		Returns
		-------
		True
			If the layout was set.

		Raises
		------
		AttributeError
			If the layout could not be found.
		COMError
			If the Document object could not be found.

		"""
		try:
			for layout in self.iter_layouts(doc=self.doc):
				if layout.Name == name:
					self.doc.ActiveLayout = layout
					return True
			raise AttributeError('Layout "%s" could not be found.' % name)
		except COMError as error:
			raise error

	def regen(self):
		"""Update the active document."""
		self.doc.Regen(ACAD.acAllViewports)


class CADTable(object):
	"""
	Provides a native AutoCAD table for the active layout.

	Parameters
	----------
	cad : AutoCAD
		An instance of AutoCAD with a valid active document.
	x : int
		Table X position.
	y : int
		Table Y position.
	row_count : int
	col_count : int
	row_height : int
	table_width : int

	Attributes
	----------
	cad : AutoCAD
	table : AutoCAD Table object

	Raises
	------
	CADLayerError
		If the current layer is locked.

	"""
	def __init__(self, cad, x, y, row_count, col_count, row_height, table_width):
		# Cell walls
		self.masks = [
			ACAD.acBottomMask, 
			ACAD.acLeftMask, 
			ACAD.acRightMask, 
			ACAD.acTopMask
		]
		# Lineweights
		self.lw_thin = ACAD.acLnWt013
		self.lw_thick = ACAD.acLnWt040

		# Create table
		try:
			self.table = cad.doc.paperspace.AddTable(
				APoint(x, y),
				row_count,
				col_count,
				row_height,
				table_width/col_count
			)
		except COMError:
			raise CADLayerError()

	def set_title_row(self, title, text_height, row_height):
		"""
		Parameters
		----------
		title : str
		text_height : int
		row_height : int

		"""
		self.table.SetText(0, 0, title)
		self.table.SetCellTextHeight(0, 0, text_height)
		self.table.SetRowHeight(0, row_height)

	def add_contrast(self):
		"""Thicken exterior walls and thin interior walls of table."""
		for row in range(self.table.rows):
			for col in range(self.table.columns):
				for mask in self.masks:
					self.table.SetCellGridVisibility(row, col, mask, True)
					self.table.SetCellGridVisibility(
						row, col, mask, self.lw_thin
					)
				if col == 0:
					self.table.SetCellGridLineWeight(
						row, col, self.masks[1], self.lw_thick
					)
				if col == (self.table.columns - 1):
					self.table.SetCellGridLineWeight(
						row, col, self.masks[2], self.lw_thick)
				if row == 0:
					self.table.SetCellGridLineWeight(
						row, col, self.masks[3], self.lw_thick
					)
				if row == (self.table.rows - 1):
					self.table.SetCellGridLineWeight(
						row, col, self.masks[0], self.lw_thick
					)


class CADLayerError(Exception):
	"""Raised during an attempt to place an object on a locked layout."""
	def __init__(self):
		self.message = 'The layer you are trying to access is locked.'

class CADDocError(Exception):
	def __init__(self):
		self.message = 'The active document could not be found.'

class CADOpenError(Exception):
	def __init__(self):
		self.message = 'This operation is unavailable. '
		self.message += 'Make sure AutoCAD is available and try again.'
