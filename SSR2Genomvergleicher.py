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
VERSION = "0.2"


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
	  description="""  Transforms a spreadsheet file or csv table
  with SSR data into a csv table fit to be loaded into Genomvergleicher2.
  Optionally attach to a existant table.

  Where options need table columns as arguments,
  these can be given by name, letter or number (starting at 1).
  Specify '0' to leave a column empty.
  Where adequate, information not given explicitly will be guessed or omitted.
  """,
	  add_help=True)
	##ToDo: Try to add second help option '-?'
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
	parser.add_argument("INFILE",
	  default="",
	  help="SSR fingerprint table in csv or arbitrary spreadsheet format.\n")
	parser.add_argument("-s", "--sheet",
	  default=1,
	  help="(Default =%(default)s) For a spreadsheet input file with more than one sheet, specify the default sheet that holds the genetic data.\n'1'=First sheet and so on.\n")
	parser.add_argument("-c", "--startcol",
	  default=None,
	  help="(Default =%(default)s) Column in which the genetic data start.\n")
	parser.add_argument("-r", "--startrow",
	  default=None,
	  help="Row in which the genetic data start, counting from 1.\n")

	parser.add_argument("-n", "--name",
	  default="",
	  help="Which column to use as the 'Name' (cultivar name) column.\n")
	parser.add_argument("-e", "--prefix",
	  default=" ", # the space char makes sure the user is not prompted about this any further; it is ignored or gets automatically dropped next time the data are read from file
	  help="Add this prEfix text to the texts of the 'Name' column.\n")
	parser.add_argument("-u", "--suffix",
	  default=" ", # the space char makes sure the user is not prompted about this any further; it is ignored or gets automatically dropped next time the data are read from file
	  help="Add this sUffix text behind the texts of the 'Name' column.\n")
	parser.add_argument("-g", "--genetic_group",
	  default="",
	  help="Which column to use as the 'Gene Group' (molecular group) column.\n")
	parser.add_argument("-f", "--reference",
	  default="",
	  help="Which column to use as the 'ReFerence' column (with information weather this is a reference genotype).\n")
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

	parser.add_argument("-q", "--no_deduplication",
	  action='store_const',
	  const="y",
	  default="n",
	  help="Don't remove duplicate fingerprints.\n")

	parser.add_argument("-a", "--attachto",
	  default="",
	  help="Attach results to this preexistant table formatted for Genomvergleicher.\n")

	parser.add_argument("-w", "--overwrite",
	  action="store_const",
	  const="w",
	  default="x",
	  help="""Specify this to overwrite the output file if it already exists.
	  Else, the program will terminate in that case.\n""")
	parser.add_argument("-o", "--outfile",
	  default="", ## again the space char
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
	debug(parameter, "guessing Parameter")
	debug(inputCols, "from columns")
	if len(inputCols) == 0:
		print("Warning: 'guessOrGet' called without any inputCols to choose from!")
	menu = []
	for colName in inputCols:
		debug(colName, "querying input column")
		if str(colName).upper().find(parameter.upper()) > -1: ##ToDo: Use a fuzzy find instead?
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
	if params["INFILE"].lower().endswith(".csv"):
		prelimDf = pd.read_csv(params["INFILE"],
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
		  params["INFILE"],
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
	#return str(outData).replace("|","~")
	return outData


def askForFilenames():
	""" This makes sure the program has a file to write its output to;
		either the filename passed as CL option "-o / --outfile"
		or the file passed with "-a / --attach_to".
		If the files are not not given or not in .csv format,
		we attempt to derive a name from the infile name.
	Returns:
		None. Mofifies the global 'params' dictionary.
	"""
	if len(params["attachto"]) > 0: # a file to attach data to was named on CL
		if not params["attachto"].lower().endswith(".csv"):
			print("File to attach data to must be in .csv format, with a .csv ending. Please reconsider. Bye!")
			sys.exit(1)
		if not Path(params["attachto"]).is_file():
			print("File to attach data to not found. Perhaps you misspelled? I quit. Bye!")
			sys.exit(1)
	else: # no file to attach the data to was given
		##params["attachto"] = ""
		if len(sys.argv) < 3:
			# we've been called with only the infile as argument --
			# perhaps we are in drag+drop-mode?
			# just in case, we'll ask the user for a premade file
			# to attach data to
			attfilename = "/"
			while attfilename == "/": ##ToDo: Use a decent file chooser!
				print("""
  Would you like to attach your genetic data to a preexistant table?
  (Hint: It might be more convenient to name a file to attach data to
  on the command line with the '-a filename.csv' switch.)
				  """)
				attfilename = input("Please enter filename, 'q' to quit, or 'x' or just ENTER to not attach to another file.")
				if attfilename == "q":
					print("See you later!")
					sys.exit(0)
				if attfilename == "" or attfilename == "x":
					attfilename = ""
					break
				if not Path(attfilename).is_file():
					attfilename = "/"
					print("File not found. Please try again.")
			params["attachto"] = attfilename
			log(params["attachto"], "Attach to file")

	# check for outfile
	if params["outfile"] == "": # no outfile given
		##params["outfile"] = ""
		if params["attachto"] == "":
			# no file to attach data to either
			if params["INFILE"].lower().endswith(".csv"):
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
				params["outfile"] = Path(params["INFILE"]).stem + ".csv"
				print(f"No output filename given. Assuming {params['outfile']}")

	if params["overwrite"] == "x":
		if Path(params["outfile"]).is_file(): #file exists
			# overwriting has not been explicitely specified
			answer = input(f"\n  File {params['outfile']} already exists. Choose 'y' to overwrite or 'q' to quit. ")
			if str(answer).lower() == "y":
				params["overwrite"] = "w"
			else:
				print("Bye, see you later!")
				sys.exit(0)
	log(params["overwrite"], "outfile overwrite mode")

	if params["attachto"] == "":
		del params["attachto"]


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
			if (value is None) or (str(value) == ""):
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


def notEmpty(x):
	""" Checks whether arg is an empty string or NaN / None / NA.
	Args:
		x (str): String to check.
	Returns:
		True if x is not "",
		False if x is "".
	"""
	if str(x) == "" or x is None or x.isna:
		return False
	else:
		return True

def dedup(df, scol, protocolFile): ##ToDo: Remove
	"""	Finds partial duplicates in a DataFrame, removes them
		and writes a protocol to file.
	Args:
		df (DataFrame): Input data.
		scol (int): Column from which on the relevant data start.
		protocolFile (str): Filename to write the deduplication protocol to.
	Returns:
		DataFrame: The deduplicated data.
	"""
	maskDup = df.duplicated(
	  subset=df.columns[scol :],
	  keep="first")

	del df.loc[maskDup, : ]
	pass

def removeDuplicates(df, scol=7, protFileStem=""):
	""" Remove rows with duplicate genetic fingerprints, ignoring metadata.
		The first occurrence of every fingerprint is kept.
		Every further occurrence is removed.
	Args:
		df (DataFrame): Table to be deduplicated.
		scol (int): Column where the genetic data start, 0-based.
	Returns:
		cleaned_df (DataFrame): DESCRIPTION.
	"""
	# separate metadata and genetic profile columns:
	metadata_cols = list(df.columns[:scol])
	profile_cols = list(df.columns[scol:])

	debug(metadata_cols, "Metadata Columns for deduplication")
	debug(profile_cols, "Profile Columns for deduplication")

	# we make a working copy, so the original df is left alone
	working_df = df.copy().reset_index(drop=True)

	# store the current row numbers in a column
	# => allows us to identify the original rows later:
	working_df["row_id"] = working_df.index

	# identify duplicate fingerprints -
	# This produces a series of Booleans reflecting the index of the df, where:
	# False = first occurrence (original)
	# True  = duplicate
	duplicate_mask = working_df.duplicated(
	  subset=profile_cols,
	  keep="first")

	# for every profile group, determine the row number of first occurrence.
	# => transform("first") returns one value for every row, namely
	# the row_id of the original profile.
	original_row = (working_df
	  .groupby(profile_cols, sort=False)["row_id"]
	  .transform("first"))

	# Build a report DataFrame --
	# metadata of duplicate rows:
	duplicate_metadata = (
	  working_df.loc[duplicate_mask, metadata_cols]
	  .reset_index(drop=True))

	# metadata of the corresponding originals:
	original_metadata = (
	  working_df.loc[original_row[duplicate_mask], metadata_cols]
	  .reset_index(drop=True))

	# differentiate the 'original' columns
	original_metadata.columns = [
	  f"{col}_original" for col in metadata_cols
	  ]

	# create a 'separator' column
	separator = pd.DataFrame({
	  "<Relation>": [" <is a genetic duplicate of> "] * duplicate_mask.sum()
	  })

	# assemble the report
	report_df = pd.concat(
	  [
	   duplicate_metadata,
	   separator,
	   original_metadata
	  ],
	  axis=1)

	print("Removed duplicate fingerprints:")
	print(report_df)
	if protFileStem != "":
		print("Writing deduplication report to " + protFileStem + ".deduplication-report.csv")
		report_df.to_csv(
		  path_or_buf=protFileStem + ".deduplication-report.csv",
		  mode="w",
		  sep="	", # Tab character
		  index=False,
		#  decimal=",",
		  escapechar="\\"
		  )

	# Finally: create output df with duplicate rows removed
	cleaned_df = (
	  working_df.loc[~duplicate_mask]
	  .drop(columns="row_id")
	  .reset_index(drop=True)
	)
	return cleaned_df


# main program
###############

# check and parse the command line
log(len(sys.argv), "Length of CL:")
params = parseParams()
verbosity += int(params["verbose"]) + int(params["debug"])
log(f"Parameters: {params}")

if not Path(params["INFILE"]).is_file():
	print(f"Error: Input file {params['infile']} not found. Perhaps you mistyped?")
	sys.exit(1)

# make sure we know in which row the actual genetic data start
if params["startrow"] is None:
	params["startrow"] = askForStartRow()
if not str(params["startrow"]).isdecimal(): # start row number was given explicitely, but in bogus format
	print("""
  Error: The row where the genetic data start (CL argument '-r' / '--startrow')
  must be supplied as a row number, counting from 1.""")
	sys.exit(1)

"""
#Handle multiple header rows as multiindex...
if int(params["startrow"]) == 2:
	# The data start in the second row, so the column labels ('header') are in first row, or in pythonic: The 0st row
	debug("startrow is 2")
	columnhead = 0
elif int(params["startrow"]) > 2:
	# There probably are several rows with header info; lets hope it makes sense to engulf them all
	columnhead = list(range(int(params["startrow"]) - 1))
else:
	columnhead = None
log(columnhead, "Rows to use as column headers")
"""

# now load input file:
if params["INFILE"].lower().endswith(".csv"):
	dfPreview = pd.read_csv(params["INFILE"],
	  sep=";",
	  header=None,
	  engine="pyarrow",
	  escapechar="\\",
	  skip_blank_lines=True,
	  keep_default_na=False,
	  skipinitialspace=True,
	  on_bad_lines="warn")
else:
	# not a .csv, so probly some sort of officeware spreadsheet...
	# pandas.read_excel can digest them all.

	# first we get the column names...
	if str(params["sheet"]).isdecimal():
		params["sheet"] = int(params["sheet"]) - 1
	inputColumns = pd.read_excel(
	  params["INFILE"],
	  header=None,
	  sheet_name=params["sheet"],
	  nrows=0,
	  keep_default_na=False
	  ).columns
	# ...then we read the whole table, converting all ";" to "," and "|" to "~"
	# (see function "maskSpecialChars")
	dfPreview = pd.read_excel(
	  params["INFILE"],
	  header=None,
	  sheet_name=params["sheet"],
	  keep_default_na=False,
	  converters={col: maskSpecialChars for col in inputColumns}
	  )

# make sure we get all available information for the column labels in cases with more than 1 header rows
if str(params["startrow"]).isdecimal() and int(params["startrow"]) > 2:
	debug(dfPreview.iloc[int(params["startrow"]) - 2 : int(params["startrow"]) - 1], "Header row")
	for col, content in dfPreview.iloc[int(params["startrow"]) - 2].items():
		# walk through the rows above the genetic data, look for column label candidates
		if content == "": #content.isna().all():
			log("###################################No column label!")
			debug(str(content) +" @Type: "+ str(type(content)) +" @Size: "+ str(len(content)), "Column " + str(col))
			dfPreview.iloc[int(params["startrow"]) - 2, col] = "" # Make sure this is an empty string, not float or whatever
			for i in range(int(params["startrow"]) - 3, -1, -1):
				debug(str(i) +" "+ str(col), "Position")
				debug(str(dfPreview.iloc[int(i), int(col)]), "Row " + str(i))
				debug(str(dfPreview.iloc[int(params["startrow"]) - 2, col]) + " #Type: " + str(type(dfPreview.iloc[int(params["startrow"]) - 2, col])), "Amend")
				debug("", "amending...")
				dfPreview.iloc[int(params["startrow"]) - 2, col] = str(dfPreview.iloc[int(params["startrow"]) - 2, col]) + str(dfPreview.iloc[int(i), int(col)])

df1 = dfPreview.rename(columns=dfPreview.iloc[int(params["startrow"]) - 2]).drop(dfPreview.index[ : int(params["startrow"]) - 1]).reset_index(drop=True)

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
		print("Could not identify start column.")
		if str(params["startcol"]).find(":") > -1:
			print("Perhaps you tried to specify the start column in SHEET:COL format?")
			print("Please use the '--sheet' argument to specify a table sheet other than the first.")
		print("...trying to proceed anyway.")

# now we file through the parameters:

# verify file names for output
askForFilenames()

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
	##if str(params[s]) == "!":
	##	params[s] = ""
# in case the help message was saved in params...
if "help" in params.keys():
	del params["help"]

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
dfProcessed = pd.DataFrame()
dfProcessed = addColumn(dfProcessed, df1, params["name"], "Name")
dfProcessed = addColumn(dfProcessed, df1, params["genetic_group"], "Gene Group")
dfProcessed = addColumn(dfProcessed, df1, params["reference"], "Reference")
dfProcessed = addColumn(dfProcessed, df1, params["trueness"], "Trueness to Type")
dfProcessed = addColumn(dfProcessed, df1, params["munq"], "MUNQ")
dfProcessed = addColumn(dfProcessed, df1, params["ploidy"], "Ploidy")
dfProcessed = addColumn(dfProcessed, df1, params["includes_mutations"], "Includes Mutations")

#get rid of the empty last column
del df1["AnEmptyColumn"]

# add the actual genetic data
dfProcessed = dfProcessed.join(
  df1.iloc[ : , (params["startcol"]) : ].astype(int, errors="ignore"),
  how="left"
  )

debug(dfProcessed.dtypes, "Types before NaN-removal")

# turn empty string values in the genetic data into integer 0
sub = dfProcessed.iloc[0 : , 7 : ]
dfProcessed.iloc[0 : , 7 : ] = sub.mask(sub.eq("") | sub.isna(), int(0))

debug(dfProcessed.dtypes, "Types after NaN-removal")

# turn empty string values in 'Name' column to '—'
sub = dfProcessed.Name
dfProcessed.Name = sub.mask(sub.eq("") | sub.isna(), "—")

# add prefix and suffix to 'Name' column
if str(params["prefix"]) != " " or str(params["suffix"]) != " ":
	log(params["prefix"], "Prefix")
	log(params["suffix"], "Suffix")
	dfProcessed["Name"] = str(params["prefix"]) + dfProcessed["Name"] + str(params["suffix"])
	##dfProcessed.assign(Name = lambda thisDF: (str(params["prefix"]) + thisDF.Name + str(params["suffix"])))
					##= str(params["prefix"]) + str(dfProcessed.Name) + str(params["suffix"]))

# trim leading and trailing spaces
#dfProcessed.Name = dfProcessed.Name.replace(r"^ +| +$", r"", regex=True)

log(dfProcessed, "Output")
log(dfProcessed.columns, "Output Columns")

# write output to file
if not "attachto" in params.keys():
	# we write output to its own file
	dfProcessed.to_csv(
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
else:
	# we attach output to another file, which must be in Genomvergleicher2 csv format
	log(params["attachto"], "Loading master file to attach to...")
	dfAttach = pd.read_csv(
	  params["attachto"],
	  sep=";",
	  header=0,
	  #engine="pyarrow",
	  escapechar="\\",
	  keep_default_na=False,
	  skip_blank_lines=True,
	  skipinitialspace=True)

	log(str(dfAttach.columns[0:7]), "Checking column labels of target file...")
	if list(dfAttach.columns[0:7]) != list(dfProcessed.columns[0:7]):
		print("""
  The file you want to attach your data to doesn't seem to be in Genomvergleicher2 format.
  Maybe you want to run it through this program first?
  Bye!""")
		sys.exit(1)

	debug(dfAttach.Hi02c07_6.dtypes,"dfAttach Type Hi02c07_6:")

	dfCombined = pd.concat([dfAttach, dfProcessed],
	  axis="index",
	  #ignore_index=True, # resulting df gets a new, clean row numbering
	  join="outer"
	  )

	# turn empty string values and NaNs in the combined genetic data into integer 0
	sub = dfCombined.iloc[0 : , 7 : ]
	dfCombined.iloc[0 : , 7 : ] = sub.mask(sub.eq("") | sub.isna(), int(0))
	dfCombined.iloc[0 : , 7 : ] = dfCombined.iloc[0 : , 7 : ].astype(int)

	debug(dfCombined.Hi02c07_6.dtypes,"dfCombined Type Hi02c07_6:")

	# rearrange the genotype columns alphabetically
	newColOrder = list(dfCombined.columns[0 : 7]) + sorted(list(dfCombined.columns[7 : ]))
	dfCombined = dfCombined.reindex(columns=newColOrder)

	debug(dfCombined.Hi02c07_6.dtypes,"dfCombined Type Hi02c07_6:")

	if "no_deduplication" in params.keys() and params["no_deduplication"] == "n":
		dfCombined = removeDuplicates(dfCombined,
		  protFileStem=Path(params["attachto"]).stem + "--" + Path(params["INFILE"]).stem)

	# write result to file
	dfCombined.to_csv(
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

