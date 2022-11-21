''' a layer on top of xlsxwriter to support panda style excel formula creation '''
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

ALL = None
const = {}
dom   = {}
views = []

def writexls(filename):
	''' write all views and constants into an xlsx file '''
	with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
		for view in views:
			view.value.to_excel(writer, sheet_name=view.sheet
				, startrow=view.anchor[0], startcol=view.anchor[1])
			for iconst in const:
				const[iconst].define(writer.book)

class Constant():
	''' xlsx named constants '''
	irow = 0
	constsheet = None
	constsheetname = 'contants'

	def __init__(self, name, value):
		self.name   = name
		self.value  = value
		self.row    = Constant.irow
		Constant.irow += 1
		const[self.name] = self

	def define(self, workbook):
		if workbook.get_worksheet_by_name(Constant.constsheetname) is None:
			Constant.constsheet = workbook.add_worksheet(Constant.constsheetname)
		nameloc=xl_rowcol_to_cell(self.row, 0)
		valueloc=xl_rowcol_to_cell(self.row, 1, row_abs=True, col_abs=True)
		workbook.define_name(self.name, f'={Constant.constsheetname}!{valueloc}')
		Constant.constsheet.write(nameloc, self.name)
		Constant.constsheet.write(valueloc, self.value)


class Formula():
	''' xlsx formula '''

	def __init__(self, datadef):
		self.rows = datadef["rows"]
		self.cols = datadef["cols"]
		self.vals = datadef["vals"]
		self.index = pd.MultiIndex.from_product([datadef[i] for i in self.rows], names=self.rows)
		self.columns = pd.MultiIndex.from_product([datadef[i] for i in self.cols], names=self.cols)
		self.values = [[self.vals(*list(irow + icol)) for icol in self.columns] for irow in self.index]

	def data(self):
		return pd.DataFrame( self.values, index=self.index, columns=self.columns )


class View():
	''' xlsx view '''

	def __init__(self, viewdef):
		self.sheet       = viewdef["sheet"]
		self.anchor      = viewdef["anchor"]
		self.value       = viewdef["value"]
		if isinstance(self.value, Formula):
			self.value = self.value.data()
		if (isinstance(self.value.index, pd.core.indexes.base.Index)
			and not isinstance(self.value.index, pd.core.indexes.multi.MultiIndex)):
			self.value.index = pd.MultiIndex.from_product([list(self.value.index),], names=[self.value.index.name,] )
		if (isinstance(self.value.columns, pd.core.indexes.base.Index)
			and not isinstance(self.value.columns, pd.core.indexes.multi.MultiIndex)):
			self.value.columns = pd.MultiIndex.from_product([list(self.value.columns),], names=[self.value.columns.name,] )
		if (not isinstance(self.value.index, pd.core.indexes.multi.MultiIndex)
			or not isinstance(self.value.columns, pd.core.indexes.multi.MultiIndex)):
			raise TypeError("xlsxview dataframe indices and columns must be indices or multiindices,"
				+" example: index=pandas.MultiIndex.from_product([[1,2,3]], names=['myidx'])")
		self.indexdim    = self.value.index.levshape
		self.indexdimlen = len(self.value.index.levshape)
		self.columndim   = self.value.columns.levshape
		self.columndimlen= len(self.value.columns.levshape)
		for i, val in enumerate(list(self.value.index.names)):
			if val in dom:
				if dom[val] != list(self.value.index.get_level_values(i).unique()):
					raise ImportError(f"Conflicting domains. Name={val} has already ben defined in a different way.")
			else:
				dom[val] = list(self.value.index.get_level_values(i).unique())
				print(f"Imported dom[{val}]={dom[val]}")
		for i, val in enumerate(list(self.value.columns.names)):
			if val in dom:
				if dom[val] != list(self.value.columns.get_level_values(i).unique()):
					raise ImportError(f"Conflicting domains. Name={val} has already ben defined in a different way.")
			else:
				dom[val] = list(self.value.columns.get_level_values(i).unique())
				print(f"Imported dom[{val}]={dom[val]}")
		views.append(self)

	# example[  0  ][  1  ][  2  ][  3  ][  4  ][  5  ]
	# [  0  ]               col11         col12
	# [  1  ]               col21  col22  col21  col22
	# [  2  ] row1   row2  ____________________________
	# [  3  ]   A      0      a      b      c      d
	# [  4  ]          1      e      f      g      h
	def ref(self, row, col, sheetref=False, debug=False):
		extrarowforindexheader=1
		if not isinstance(row,tuple):
			row = (row,)
		if not isinstance(col,tuple):
			col = (col,)
		row = list(row)
		col = list(col)
		retval={}
		for mode in ['min', 'max']:
			minmaxrow = tuple([ self.value.index.get_level_values(i).unique()[0 if mode=='min' else -1]
				if val is None else val for i, val in enumerate(row) ])
			minmaxcol = tuple([ self.value.columns.get_level_values(i).unique()[0 if mode=='min' else -1]
				if val is None else val for i, val in enumerate(col) ])
			ridx = self.value.index.get_indexer([minmaxrow])[0]
			cidx = self.value.columns.get_indexer([minmaxcol])[0]
			retval[mode] = xl_rowcol_to_cell(
				self.anchor[0] + self.columndimlen + ridx + extrarowforindexheader
				, self.anchor[1] + self.indexdimlen + cidx )
			print(f"REF{mode} = " + str(retval[mode])
			 + " ROW row="+str(row)+" anchor="+str(self.anchor[0])+" offset=" +str(self.columndimlen)+" "+str(ridx)+" +1"
			 +" COL col="+str(col)+" anchor="+str(self.anchor[1])+" offset=" +str(self.indexdimlen)+" "+str(cidx) ) if debug else None
		retvalue = ("'"+self.sheet+"'!" if sheetref else "") + (retval['min'] if retval['min']==retval['max'] else retval['min']+':'+retval['max'])
		print(retvalue) if debug else None
		return retvalue

