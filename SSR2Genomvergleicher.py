#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri May  8 17:34:46 2026

@author: greengooscenario

Format of output data (Genomvergleicher2 standard):
	- The first row is the column names, all following rows are data
	- Anything from column 8 ('H') on is genotype data,
	  anything before is metadata

Note: openpyxl (a backend library used by pandas.read_excel) version 3.1.2
is buggy and might cause the cryptic error message
"Value must be either numerical or a string containing a wildcard".
Avoid this version in your python setup.
"""

# libraries:
import sys
import argparse
from pathlib import Path
import pandas as pd

# configuration:

# Verbosity level for dubugging
# (preliminary setting; CL arguments "-d" and "-v" can increase the level)
verbosity = 1

# Program version string
VERSION = "0.1"


def log(msg, header=""):
	""" Prints a message under a header when global variable verbosity > 0.
	Args:
		msg (any type): A message string or just any other datatype. Will be printed.
		header (str, optional): A title that will be printed above the message. Defaults to "".
	Returns:
		None.
	"""
	if verbosity > 0:
		print(">>>" + header + ":")
		print(msg)


def debug(msg, header=""):
	""" Prints a message under a header when global variable verbosity > 1.
	Args:
		msg (any type): A message string or just any other datatype. Will be printed.
		header (str, optional): A title that will be printed above the message. Defaults to "".
	Returns:
		None.
	"""
	if verbosity > 1:
		print(">>>" + header + ":")
		print(msg)


def parseParams():
	""" Handles the argparse library and evaluates the command line (CL).
	Returns:
		params (dict): parameters parsed from CL.
	"""
	parser = argparse.ArgumentParser(
	  description="""Transforms a spreadsheet file or csv table
	  with SSR data into a csv table fit to be loaded into Genomvergleicher2.
	  Optionally attach to a existant table, omitting duplicates.

	  Table columns can be given by name, letter or number (starting at 1).
	  Where adequate, information not given explicitly
	  will be guessed or omitted.
	  """,
	  add_help=True)

	parser.add_argument("-V", "--version",
	  action='version',
	  version=f"SSR2Genomvergleicher version {VERSION} -- try '--help' for further instructions")
	parser.add_argument("-v", "--verbose",
	  action='store_const',
	  const=1,
	  default=0,
	  help="Be communicative about what is being done.\n")
	parser.add_argument("-d", "--debug",
	  action='store_const',
	  const=2,
	  default=0,
	  help="Be very communicative about what is being done.\n")
	parser.add_argument("infile",
	  default="",
	  help="SSR fingerprint table in csv or arbitrary spreadsheet format.\n")
	parser.add_argument("-s", "--sheet",
	  default=1,
	  help="(Default =%(default)s) For a spreadsheet input file with more than one sheet, specify here which one to work on. '1'=First sheet and so on.\n")
	parser.add_argument("-c", "--startcol",
	  default=None,
	  help="(Default =%(default)s) Column in which the genetic data start.\n")
	parser.add_argument("-r", "--startrow",
	  default=None,
	  help="Row in which the genetic data start, counting from 1.\n")

	parser.add_argument("-n", "--name",
	  default="",
	  help="Which column to use as the 'Name' (cultivar name) column.\n")
	parser.add_argument("-g", "--genetic_group",
	  default="",
	  help="Which column to use as the 'Gene Group' (molecular group) column.\n")
	parser.add_argument("-f", "--reference",
	  default="",
	  help="Which column to use as the 'Reference' column (with information weather this is a reference genotype).\n")
	parser.add_argument("-t", "--trueness",
	  default="",
	  help="Which column to use as the 'Trueness-to-Type' column.\n")
	parser.add_argument("-m", "--munq",
	  default="",
	  help="Which column to use as the 'MUNQ' (unique genotype code) column.\n")
	parser.add_argument("-p", "--ploidy",
	  default="",
	  help="Which column to use as the 'Ploidy' (diploid, triploid, or tetraploid) column.\n")
	parser.add_argument("-i", "--includes_mutations",
	  default="",
	  help="Which column to use as the 'Includes mutations' (that can't be differentiated using SSR markers) column.\n")

	parser.add_argument("-a", "--attachto",
	  default=" ", # space char - important marker!
	  help="Attach results to this preexistant table formatted for Genomvergleicher.\n")
	parser.add_argument("-w", "--overwrite",
	  action="store_const",
	  const="w",
	  default="x",
	  help="""Specify this to overwrite the output file if it already exists.
	  Else, the program will terminate in that case.\n""")
	parser.add_argument("-o", "--outfile",
	  default=" ", # again the space char
	  help="Filename for output.\n")

	outputDict = vars(parser.parse_args())
	##outputDict["help"] = parser.format_help()
	return outputDict


def guessOrGet(parameter, inputCols):
	""" If column mapping was not completely defined on CL:
		Try to guess by name which column of input file to use as which output column.
		Where our guess is ambiguos, we ask the user.
	Args:
		parameter (str): The parameter / target column to guess.
		inputCols (list): Columns from the input file to choose from.
	Returns:
		str: A column name from the input file.
	"""
	if len(inputCols) == 0:
		print("Error: 'guessOrGet' called without any inputCols to choose from!")
	menu = []
	for colName in inputCols:
		if colName.upper().find(parameter.upper()) > -1: ##ToDo: Use a fuzzy find instead?
			# parameter is among the input columns
			menu.append(colName)

	if len(menu) == 0:
		# parameter not found among the input cols
		# => let user pick from all metadata columns
		print("\nUnable to guess " + parameter.upper() + " column. Please pick manually:")
		return usersPick(inputCols)
	if len(menu) == 1:
		# only one match found
		return menu[0]
	elif len(menu) > 1:
		# more than one candidate among metadata cols - let user pick one
		print("More than one candidate column for " + parameter.upper() + " found. Please pick one:")
		return usersPick(menu)


def usersPick(l=None, required=False, optSpacer="\n"):
	""" Have the user interactively pick one item from a list.
	Args:
		l (list-like, or None):
			enumerate()-able list of items to offer to the user,
			or None to not display the options
			(typically when this has been handled elsewhere)
		required (boolean):
			If True, the user must choose an option,
			or quit the program.
			If False, the user is presented with the option
			to reject to choose, without quitting.
	Returns:
		None if required == False and the user entered 'x'
		An item from parameter l as picked by the user if l was passed
		The string entered by the user if no list l was passed
	"""
	prompt = "\n"
	specialOpts = ["q"]
	thirdOpt = ""
	if required is False:
		thirdOpt = "letter 'x' to leave blank, "
		specialOpts = ["q", "x"]
	if l is not None:
		for index, item in enumerate(l):
			# we want to show a 1-based list, so we add 1 to the indices:
			prompt += (f"{index + 1}) {item}" + optSpacer)
	a = ""
	while not (a.isdecimal() or (a in specialOpts)):
		# we'll keep asking until the user makes sense!
		print(prompt)
		a = input("\nPlease enter number, " + thirdOpt + "or 'q' to quit: ")
	if a == "q":
		print("Good choice! Feel free to come back when you have made up your mind!")
		sys.exit(0)
	elif a == "x":
		return None
	if l is None:
		return a
	else:
		return l[int(a) - 1] # we return the value, not the index


def askForStartRow():
	""" Make sure we have a data start row that defines the header / column labels row
	Returns:
		int: The first row of data, 1-based.
	"""
	# load files, preliminary
	if params["infile"].lower().endswith(".csv"):
		prelimDf = pd.read_csv(params["infile"],
		  sep=";",
		  header=None,
		  engine="pyarrow",
		  escapechar="\\", # we need to escape the backslash
		  skip_blank_lines=True,
		  skipinitialspace=True)
	else:
		# not a .csv, so probly some sort of officeware spreadsheet...
		# pandas.read_excel can digest them all.
		prelimDf = pd.read_excel(
		  params["infile"],
		  header=None,
		  sheet_name=int(params["sheet"]) - 1
		  )
	# purpose-made df to reshape the input,
	# for the user to pick start row in a 1-based manner:
	(ourRow, ourCol) = prelimDf.shape
	zeroDf = pd.DataFrame("", index=range(1), columns=range(ourCol))
	prelimDf = pd.concat([zeroDf, prelimDf], axis="index", ignore_index=True)

	print("""
  The row in which the genetic data start has not been defined -
  please pick the correct row number manually.
  (Hint: We are looking for the first row that is NOT column labels.
  If you don't know what to do, '2' would be a good guess.)
  """)
	print(prelimDf.iloc[1:6])
	return int(usersPick(required=True))


def positiveOrNone(n, printIfNone="", printIfNotNone=""):
	""" Returns None if a value is not >= 0
	Args:
		n (int or float): Input
		printIfNone (str): Message to print if n is Negative
		printIfNotNone (str): Message to print if n is 0 or Positive
	Returns:
		n if it is 0 or Positive,
		None if n is Negative
	"""
	if n >= 0:
		if printIfNotNone != "":
			print(printIfNotNone)
		return n
	else:
		if printIfNone != "":
			print(printIfNone)
		return None


def maskSpecialChars(inputData):
	""" Replace certain characters with other, in a hard-coded manner.
		Will mostly be called from Lamda expressions.
	Args:
		inputData (any, as string): Typically, a string-like to have certain chars replaced.
	Returns:
		str: A string in which certain characters have been substituted with placeholders.
	"""
	outData = str(inputData).replace(";",",")
	return str(outData).replace("|","~")


def askForFilenames():
	""" This makes sure the program has a file to write its output to;
		either the filename passed as CL option "-o / --outfile"
		or the file passed with "-a / --attach_to".
		If the files are not not given or not in .csv format,
		we attempt to derive a name from the infile name.
	Returns:
		None. Mofifies the global 'params' dictionary.
	"""
	if params["attachto"] == " ":
		# no file to attach the data to was given
		params["attachto"] = ""
		if len(sys.argv) < 3:
			# we've been called with only the infile as argument --
			# perhaps we are in drag+drop-mode?
			# just in case, we'll ask the user for a premade file
			# to attach data to
			attfilename = "/"
			while attfilename == "/": ##ToDo: Use a decent file chooser!
				print("\nWould you like to attach your genetic data to a preexistant table?")
				attfilename = input("Please enter filename, or press ENTER to not attach anywhere.")
				if attfilename != "" and (not Path(attfilename).is_file()):
					attfilename = "/"
					print("File not found. Please try again.")
			params["attachto"] = attfilename
			log(params["attachto"], "Attach to file")
	else: # a file to attach data to was named on CL
		if not params["attachto"].lower().endswith(".csv"):
			print("File to attach data to must be in .csv format. Please reconsider. Bye!")
			sys.exit(0)

	if params["outfile"] == " ": # no outfile given
		params["outfile"] = ""
		if params["attachto"] == "":
			# no file to attach data to either
			if params["infile"].lower().endswith(".csv"):
				# must ask
				outfilename = input("Please enter output filename, or 'q' to quit': ")
				if outfilename == "q":
					print("Good idea to take some time to think! Bye!")
					sys.exit(0)
				if not outfilename.lower().endswith(".csv"):
					outfilename += ".csv"
				params["outfile"] = outfilename
				print(f"Will write output to {params['outfile']}")
			else:
				params["outfile"] = Path(params["infile"]).stem + ".csv"
				print(f"No output filename given. Assuming {params['outfile']}")

	if params["overwrite"] == "x":
		if Path(params["outfile"]).is_file(): #file exists
			# overwriting has not been explicitely specified
			answer = input(f" File {params['outfile']} already exists. Choose 'y' to overwrite or 'q' to quit.")
			if str(answer).lower() == "y":
				params["overwrite"] = "w"
			else:
				print("Bye, see you later!")
				sys.exit(0)
	log(params["overwrite"], "outfile overwrite mode")


def convertToColLabels(paramDict, dframe): #currently unused?
	""" Make sure the column specifications are labels, not numbers.
	Args:
		paramDict (dict): A dict that holds the column specifications, among other things.
		dframe (DataFrame): The dataframe that the column specs refer to.
	Returns:
		paramDict (dict): The dict with the column specs now as labels.
	"""
	for key, value in paramDict.items():
		log(value, key)
		if key in ["name",
		  "genetic_group",
		  "reference",
		  "trueness",
		  "munq",
		  "ploidy",
		  "includes_mutations",
		  "startcol"]: # we only need to sanitize the column specs
			if str(value).isdecimal():
				log(value , str(key) + " column is given as a number")
				log(dframe.columns[int(value)-1], "Corresponds to name")
				paramDict[key] = dframe.columns[int(value) - 1]
				continue
			if len(str(value)) == 1 and str(value).isalpha():
				debug(key, "Column is given as letter")
				i = "abcdefghijklmnopqrstuvwxyz".index(value.lower())
				paramDict[key] = dframe.columns[i]
	return paramDict


def convertTo0basedColIndex(paramDict):
	""" Converts column specification given as label, number or letter
		to 0-based column index number.
	Args:
		paramDict (dict): The name-value dict that stores the arguments for this program
	Returns:
		paramDict (dict): The name-value dict with all values that denote columns
		 converted to 0-based indices.
	TODO: Stick an additional column to the input df,
	and turn all undefined cols into references to this empty col
	"""
	for key, value in paramDict.items():
		log(value, key)
		if key in ["name",
		  "genetic_group",
		  "reference",
		  "trueness",
		  "munq",
		  "ploidy",
		  "includes_mutations",
		  "startcol"]: # we only need to sanitize params that refer to column
			if value is None:
				paramDict[key] = -1 #we have to supply an empty last column in the input!
				continue
			elif str(value).isdecimal():
				log(value , str(key) + " column is given as a number")
				paramDict[key] = int(value) - 1
				continue
			elif (len(str(value)) == 1) and str(value).isalpha():
				debug(key, "Column is given as letter")
				paramDict[key] = int("abcdefghijklmnopqrstuvwxyz".index(value.lower()))
				continue
			elif len(str(value)) > 1:
				debug(key, "Column is given by label")
				try:
					paramDict[key] = int(list(df1.columns).index(value))
				except:
					print(f"Failed to find the column {value} you gave as parameter {key}.")
					print("Perhaps you misspelled it?")
					print("Remember: Column labels are Case Sensitive!")
					sys.exit(1)
			else:
				log(str(key) + ": " + str(value), "Can't make shaft nor bolt from this")
	return paramDict


def addColumn(dfA, dfB, colNum, newName):
	""" Appends a column to a dataFrame. A wrapper for DataFrame.insert().
	Args:
		dfA (DataFrame): DataFrame to append to.
		dfB (DataFrame): Source of column to append.
		colNum (int): Column to append, referenced by index number.
		newName (string): Name of the newly attached column.
	Returns:
		dfA (DataFrame): DataFrame with a newly attached column.
	"""
	dfA.insert(loc=len(dfA.columns),
	  column=newName,
	  value=dfB.iloc[:, colNum],
	  allow_duplicates=False)
	return dfA



# main program
###############

# check and parse the command line
log(len(sys.argv), "Length of CL:")
params = parseParams()
verbosity += int(params["verbose"]) + int(params["debug"])
log(f"Parameters: {params}")

if not Path(params["infile"]).is_file():
	print(f"Error: Input file {params['infile']} not found. Perhaps you mistyped?")
	sys.exit(1)

# first make sure we have a startrow
if params["startrow"] is None:
	params["startrow"] = askForStartRow()

# now load input file:
if params["infile"].lower().endswith(".csv"):
	df1 = pd.read_csv(params["infile"],
	  sep=";",
	  header=positiveOrNone(int(params["startrow"]) - 2,
	   printIfNone="Warning: Your choice of start row results in a table without column labels!"),
	  engine="pyarrow",
	  escapechar="\\",
	  skip_blank_lines=True,
	  skipinitialspace=True)
else:
	# not a .csv, so probly some sort of officeware spreadsheet...
	# pandas.read_excel can digest them all.
	#
	# first we get the column names...
	inputColumns = pd.read_excel(
	  params["infile"],
	  header=int(params["startrow"]) - 2,
	  sheet_name=int(params["sheet"]) - 1,
	  nrows=0
	  ).columns
	# ...then we read the whole table, converting all ";" to "," and "|" to "~"
	# (see function "maskSpecialChars")
	df1 = pd.read_excel(
	  params["infile"],
	  header=int(params["startrow"]) - 2,
	  sheet_name=int(params["sheet"]) - 1,
	  converters={col: maskSpecialChars for col in inputColumns}
	  )

log(df1.head(),"The input:")

# make sure we have a start column
if params["startcol"] is None:
	print("""
  The column in which the genetic data start has not been defined -
  please pick the correct column number manually.
  (Hint: We are looking for the first column that is NOT metadata.)
  """)
	params["startcol"] = usersPick(df1.columns, required=True, optSpacer="  |  ")

# make sure we have understood what the startcol arg wants to tell us:
if str(params["startcol"]).isdecimal(): # start column given as column number
	params["startcol"] = int(params["startcol"]) # just to make sure
	metaCols = df1.iloc[ : 1, : params["startcol"]-1].columns
elif len(str(params["startcol"])) == 1: # start column given as column letter
	# remember: For easier interaction with the user and conceptual clarity,
	# we are using 1-based indexing almost until the end
	startColNum = "abcdefghijklmnopqrstuvwxyz".index(str(params["startcol"]).lower()) + 1
	metaCols = df1.iloc[ : 1, : startColNum - 1].columns
else: # start column probably given by label
	try:
		startColNum = list(df1.columns).index(params["startcol"]) + 1
		metaCols = df1.iloc[ : 1, : startColNum - 1].columns
	except:
		metaCols = df1.columns
		log("proceeding anyway...", "Could not identify start column")

# now we file through the parameters:

##helpMessageDisplayed = False
for s in params.keys():
	debug(params[s], s)
	if params[s] == "" or params[s] == None:
		"""
		if helpMessageDisplayed == False:
			# we haven't yet printed the help message, so we do now:
			##print(params["help"])
			##log(s, "This help message was brought to you on account of")
			helpMessageDisplayed = True # we won't show this again
		"""
		params[s] = guessOrGet(s, metaCols) # get parameter as column label
if "help" in params.keys():
	del params["help"]

# verify file names for output
askForFilenames()
log(params, "Final config:")

# sanitize column names, letters or 1-based numbers to 0-based column indices
# ('None'-values will be turned into -1)
params = convertTo0basedColIndex(params)
log(params, "Sanitized:")

# stick an empty last column to the input frame
# so that 'None'-column indications give an empty column in the output
(nRow, nCol) = df1.shape
dfEmptyCol = pd.DataFrame("", index=range(nRow), columns=range(1))
log(dfEmptyCol, "Empty DF")
df1.insert(loc=len(df1.columns),
  column="AnEmptyColumn",
  value=dfEmptyCol.iloc[:,0],
  allow_duplicates=True)

# let's begin to construct the output table!
dfOut = pd.DataFrame()
dfOut = addColumn(dfOut, df1, params["name"], "Name")
dfOut = addColumn(dfOut, df1, params["genetic_group"], "Gene Group")
dfOut = addColumn(dfOut, df1, params["reference"], "Reference")
dfOut = addColumn(dfOut, df1, params["trueness"], "Trueness to Type")
dfOut = addColumn(dfOut, df1, params["munq"], "MUNQ")
dfOut = addColumn(dfOut, df1, params["ploidy"], "Ploidy")
dfOut = addColumn(dfOut, df1, params["includes_mutations"], "Includes Mutations")

#get rid of the empty last column
del df1["AnEmptyColumn"]

# add the actual genetic data
dfOut = dfOut.join(df1.iloc[ : , (params["startcol"]) : ],
  how="left",
  rsuffix="XXX"
  )

log(dfOut, "Output")
log(dfOut.columns, "Columns")

# write output to file
dfOut.to_csv(
  path_or_buf=params["outfile"],
  ##mode="w" if params["overwrite"] == "w" else None,
  mode=params["overwrite"],
  sep=";",
  index=False,
#  decimal=",",
  escapechar="\\"
  )

print("Done.")
log("Wrote output to file " + params["outfile"])


