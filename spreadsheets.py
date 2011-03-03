"""Spreadsheets.py Written by Nikola Dipanov (nikola.djipanov@gmail.com), December, 2010
This module is a python script representing a solution to the task described below. 
The orginal problem is found on the following url: http://uva.onlinejudge.org/index.php?option=com_onlinejudge&Itemid=8&category=3&page=show_problem&problem=132

Task:
n 1979, Dan Bricklin and Bob Frankston wrote VisiCalc, the first spreadsheet application. 
It became a huge success and, at that time, was the killer application for the Apple II computers. Today, spreadsheets are found on most desktop computers.
The idea behind spreadsheets is very simple, though powerful. A spreadsheet consists of a table where each cell contains either a number or a formula. 
A formula can compute an expression that depends on the values of other cells. Text and graphics can be added for presentation purposes.

You are to write a very simple spreadsheet application. Your program should accept several spreadsheets. 
Each cell of the spreadsheet contains either a numeric value (integers only) or a formula, which only support sums. 
After having computed the values of all formulas, your program should output the resulting spreadsheet where all formulas have been replaced by their value. 

Input:

The first line of the input file contains the number of spreadsheets to follow. A spreadsheet starts with a line consisting of two integer numbers, 
separated by a space, giving the number of columns and rows. The following lines of the spreadsheet each contain a row. 
A row consists of the cells of that row, separated by a single space.
A cell consists either of a numeric integer value or of a formula. A formula starts with an equal sign (=). 
After that, one or more cell names follow, separated by plus signs (+). The value of such a formula is the sum of all values found in the referenced cells. 
These cells may again contain a formula. There are no spaces within a formula.

The name of a cell consists of one to three letters for the column followed by a number between 1 and 999 (including) for the row. 
The letters for the column form the following series: A, B, C, ..., Z, AA, AB, AC, ..., AZ, BA, ..., BZ, CA, ..., ZZ, AAA, AAB, ..., AAZ, ABA, ..., ABZ, ACA, ..., ZZZ. 
These letters correspond to the number from 1 to 18278. The top left cell has the name A1

Output:

The output of your program should have the same format as the input, except that the number of spreadsheets and the number of columns and rows are not repeated. 
Furthermore, all formulas should be replaced by their value. 
"""

import types, re, itertools, csv, sys, string

def columns_naive():
	"""Generate column labels - naive generator approach"""
	alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
	for l in alphabet:
		yield l
	for l in alphabet:
		for l1 in alphabet:
			yield l + l1
	for l in alphabet:
		for l1 in alphabet:
			for l2 in alphabet:
				yield l + l1 + l2

def columns_batteries():
	"""Generate column labels using itertools batteries"""
	alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
	rl = list(alphabet)
	rl.extend(["".join(l) for l in itertools.product(alphabet, repeat=2)])
	rl.extend(["".join(l) for l in itertools.product(alphabet, repeat=3)])
	return rl

columns = columns_batteries

#Dicitonaries mapping column references to their indexes for quick lookup (to get indices in lists we'll have to subtract one)
column_map = dict(zip(columns(), range(len(columns())+1)[1:]))
column_reverse_map = dict(map(lambda t: (t[1], t[0]), column_map.items()))
	
class Cell(object):
	"""Class representing a cell in the spreadsheet - verifies that it has the propper formula passed, 
	and also saves all referenced cells to be used in the SpreadSheet class to determine circular references 
	and to calculate the values"""
	formula_re = re.compile('^(\d+)|=([A-Z]{1,3}\d{1,3})((?:[+][A-Z]{1,3}\d{1,3})*)$')
	def __init__(self, formula):
		mo = self.formula_re.match(formula)
		if mo:
			self.formula = formula
			if not formula.startswith("="):
				self.references = []
				self.value = int(formula)
			else:
				self.references = [mo.group(2)] 
				#Because python re engine returns only the last of multiple matched groups - split the 3rd group
				if mo.group(3):
					self.references.extend([s for s in mo.group(3).split("+") if s])
				self.value = None
		else:
			raise ValueError, "Inconsistent formula %s" %formula
	
	def __repr__(self):
		return str(self.value or self.formula)
			
class SpreadSheet(object):
	"""Class representing a spreadsheet with a list of cell rows - cell matrix"""
	ref_re = re.compile('([A-Z]{1,3})(\d{1,3})')
	
	def _get_cell(self, ref):
		"""Get the cell object based on the textual reference"""
		mo = self.ref_re.match(ref)
		if mo:
			col, row = mo.groups()
			try:
				return self.matrix[int(row)-1][column_map[col]-1]
			except IndexError:
				raise ValueError, "%s is not a valid refernece for this spreadsheet" %ref
		else:
			raise ValueError, "%s is not a valid refernece" %ref
	
	def _has_cir_ref(self):
		"""Method that determines if the given spreadsheet has circular references"""
		def is_reffed(cell_ref, base_cell_ref):
			"""Inner function to recursively check if there are circular 
			reference to a given cell. Does depth-first search"""
			if base_cell_ref in self._get_cell(cell_ref).references:
				return True
			else:
				for c in self._get_cell(cell_ref).references:
					if is_reffed(c, cell_ref):
						return True
				return False
		#The _has_cir_ref method body strarts here
		for i in xrange(len(self.matrix)):
			for j in xrange(len(self.matrix[i])):
				ref = column_reverse_map[j+1] + str(i+1)
				if is_reffed(ref, ref):
					return True
		return False
		
	def __init__(self, matrix = None):
		self.matrix = []
		if matrix:
			for row in matrix:
				self.matrix.append(map(lambda v: Cell(v), row))
			if self._has_cir_ref():
				raise ValueError, "Matrix contains circular references!"
	
	def compute(self):
		"""Function that computes all the missing values from the spreadsheet"""
		def compute_cell(cell_ref):
			"""Inner function to recursively compute values of all linked cells and the cell itself"""
			cell = self._get_cell(cell_ref)
			if cell.value:
				return cell.value
			else:
				cell.value = 0
				for c in cell.references:
					cell.value += compute_cell(c)
				return cell.value
		#The compute method body strarts here
		for i in xrange(len(self.matrix)):
			for j in xrange(len(self.matrix[i])):
				ref = column_reverse_map[j+1] + str(i+1)
				compute_cell(ref)
		
	def __repr__(self):
		rs = ""
		for row in self.matrix:
			rs += " ".join(map(str, row)) + "\n"
		return rs

def parse_input_file(f):
	"""Function that parses the input file and returns the list of SpreadSheet objects
	In case there is an error while parsing the file ValueError is raised outlining the 
	line where there was an inconsistency.
	"""
	ss_list = []
	int_re = re.compile("\d+")
	reader = csv.reader(f, delimiter=' ')
	try:
		first_row = reader.next()
		if len(first_row) == 1 and int_re.match(first_row[0]):
			ss_no = int(first_row[0])
			#Loop for spreadsheets
			for ss in xrange(ss_no):
				matrix = []
				#Get the line that contains rows and colums
				line = reader.next()
				col_no, row_no = line
				if len(line) == 2 and int_re.match(col_no) and int_re.match(row_no):
					for row in xrange(int(row_no)):
						line = reader.next()
						if len(line) == int(col_no):
							matrix.append(line)
						else:
							raise ValueError, "Error in the input file line %d" %reader.line_num
					#Now create the SpreadSheet object - this can raise ValueError (either bad formula or circular ref)
					#Our plan is to catch it and inform the caller of the function with another ValueError
					try:
						ss_list.append(SpreadSheet(matrix))
					except ValueError as ve:
						raise ValueError, "Error in the input file somewhere before the line %d:\n%s" %(reader.line_num, str(ve))
				else:
					raise ValueError, "Error in the input file line %d" %reader.line_num
			return ss_list
		else:
			raise ValueError, "Error in the input file line %d" %reader.line_num
	except StopIteration:
		raise ValueError, "Error in the input file - Input file incomplete"
	
if __name__ == '__main__':
	if len(sys.argv) == 2:
		if sys.argv[1] == '?':
			print "Usage: spreadsheets.py <filename>.\n\n"
			print "The first line of the input file contains the number of spreadsheets to follow. A spreadsheet starts with a line consisting of two integer numbers, separated by a space, giving the number of columns and rows. The following lines of the spreadsheet each contain a row. A row consists of the cells of that row, separated by a single space.\n"
			print "A cell consists either of a numeric integer value or of a formula. A formula starts with an equal sign (=). After that, one or more cell names follow, separated by plus signs (+). The value of such a formula is the sum of all values found in the referenced cells. These cells may again contain a formula. There are no spaces within a formula.\n" 
		else:
			try:
				f = open(sys.argv[1])
				spreadsheets = parse_input_file(f)
				for ss in spreadsheets:
					ss.compute()
					print ss
			except IOError:
				print "Error oppening file %s" %sys.argv[1]
				exit()
			except ValueError as ve:
				print ve
				exit()
			f.close()
	else:
		print "Wrong number of program arguments. Call spreadsheets.py with ? argument to see full help on how to use this script."
		exit()