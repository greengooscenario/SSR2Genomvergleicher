#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri May  8 17:34:46 2026

@author: greengooscenario

Format of output data (Genomvergleicher2 standard):
	- The first row is the column names, all following rows are data
	- Anything from column 8 ('H') on is genotype data,
	  anything before is metadata

Note: openpyxl (a backend library used by pandas.read_excel) version 3.1.2 is buggy
and might cause the cryptic error message "Value must be either numerical or a string containing a wildcard".
Avoid this version in your python setup.
"""

# libraries:
import sys
import pandas as pd
from pathlib import Path
##from io import StringIO
import argparse

# configuration:

# Verbosity level for dubugging
# (preliminary; variable leter is set according to CL arguments)
verbosity = 2

# Program version string
version = "0.1"


def log(msg, header=""):
	if verbosity > 0:
		print(">>>" + header + ":")
		print(msg)


def debug(msg, header=""):
	if verbosity > 1:
		print(">>>" + header + ":")
		print(msg)


def parseParams():
	""" Handles the argparse library and evaluates the command line (CL).
	Returns:
		params (dict): parameters parsed from CL.
	"""
	parser = argparse.ArgumentParser(
	  description="""Transforms a spreadsheet file or csv table with SSR data into a csv table fit to be loaded into Genomvergleicher2.
	  Optionally attach to a existant table, omitting duplicates.
	  \n Table columns can be given by name, letter or number (starting at 1).
	  Where adequate, information not given explicitly will be guessed or omitted.""",
	  add_help=True)

	parser.add_argument("-V", "--version", action='version', version=f"SSR2Genomvergleicher version {version} -- try '--help' for further instructions")
	parser.add_argument("-v", "--verbose",
	  action='store_const',
	  const=1,
	  default=0,
	  help="Be communicative about what is being done.")
	parser.add_argument("-d", "--debug",
	  action='store_const',
	  const=2,
	  default=0,
	  help="Output a lot of debug information about what is being done.")
	parser.add_argument("infile",
	  default="",
	  help="SSR fingerprint table in csv or arbitrary spreadsheet format.")
	parser.add_argument("-s", "--sheet",
	  default=1,
	  help="(Default =%(default)s) For a spreadsheet input file with more than one sheet, specify here which one to work on. '1'=First sheet and so on.")
	parser.add_argument("-c", "--startcol",
	  default=8,
	  help="(Default =%(default)s) Column in which the genetic data start. ")
	parser.add_argument("-r", "--startrow",
	  default=2,
	  help="(Default =%(default)s) Row in which the genetic data start, beginning at 1.")

	parser.add_argument("-n", "--name",
	  default="",
	  help="Which column to use as the 'Name' (cultivar name) column.   ")
	parser.add_argument("-g", "--genetic_group",
	  default="",
	  help="Which column to use as the 'Gene Group' (molecular group) column. ")
	parser.add_argument("-f", "--reference",
	  default="",
	  help="Which column to use as the 'Reference' column (with information weather this is a reference genotype). ")
	parser.add_argument("-t", "--trueness",
	  default="",
	  help="Which column to use as the 'Trueness-to-Type' column. When no column is given, we will assume '4' = not tested.")
	parser.add_argument("-m", "--munq",
	  default="",
	  help="Which column to use as the 'MUNQ' (unique genotype code) column. ")
	parser.add_argument("-p", "--ploidy",
	  default="",
	  help="Which column to use as the 'Ploidy' (diploid, triploid, or tetraploid) column. ")
	parser.add_argument("-i", "--includes_mutations",
	  default="",
	  help="Which column to use as the 'Includes mutations' (that can't be differentiated using SSR markers) column. ")

	parser.add_argument("-a", "--attachto",
	  default=" ", # space char - important marker!
	  help="Attach subject table to this existant table formatted for Genomvergleicher.")
	parser.add_argument("-w", "--overwrite",
	  action="store_const",
	  const="w",
	  default="x",
	  help="""Specify this to overwrite the output file if it already exists.
	  Else, the program will terminate in that case.""")
	parser.add_argument("-o", "--outfile",
	  default=" ", # again the space char
	  help="Filename for output.")

	params = vars(parser.parse_args())
	params["help"] = parser.format_help()
	return params


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
	out = []
	for colName in inputCols:
		if colName.upper().find(parameter.upper()) > -1:
			out.append(colName)
	log(len(out),"Number of candidate columns")
	if len(out) == 0:
		# prompt user
		print("Unable to guess " + parameter.upper() + " column. Please pick manually:")
		out = [usersPick(inputCols)]
	if len(out) > 1:
		# more than one candidate - let user pick one
		print("More than one candidate column for " + parameter.upper() + " found. Please pick one:")
		out = [usersPick(out)]
	return out[0]


def usersPick(l):
	""" Have the user interactively pick one item from a list.
	Args:
		l (list-like): enumerate()-able list of items to pick from.
	Returns:
		item from parameter l, or None: What the user picked.
	"""
	prompt = "\n"
	for index, item in enumerate(l):
		# we want to show a 1-based list, so we add 1 to the indices:
		prompt += f"{index + 1}) {item}\n"
	a = "§" # we'll keep asking until the user makes sense!
	while not (a.isdecimal() or a == "x" or a == "q"):
		print(prompt)
		a = input("\nPlease enter number, letter 'x' to omit column, or 'q' to quit: ")
	if a == "q":
		print("Bye! Better luck next time!")
		sys.exit(0)
	elif a == "x":
		return None
	return l[int(a) - 1]


def maskSpecialChars(input):
	""" Replace certain characters with other, in a hard-coded manner.
		Will mostly be called from Lamda expressions.
	Args:
		input (any, as string): Typically, a string-like to have certain chars replaced.
	Returns:
		str: A string in which certain characters have been substituted with placeholders.
	"""
	s = str(input).replace(";",",")
	return str(s).replace("|","~")


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
		if len(sys.argv) < 2:
			# we've been called with only the infile as argument --
			# perhaps we are in drag+drop-mode?
			# just in case, we'll ask the user for a premade file
			# to attach data to
			attfilename = "/"
			while attfilename == "/":
				print("""
				  Would you like to attach your genetic data to a preexistant table?
				  Please enter filename, or press ENTER to not attach anywhere.
				  """)
				attfilename = input()
				if not os.path.isfile(attfilename):
					attfilename = "/"
					print("File not found. Please try again.")
			params["attachto"] = attfilename
	else:
		if not params["attachto"].lower().endswith(".csv"):
			print("File to attach data to must be in .csv format. Please reconsider. Bye!")
			sys.exit(0)

	if params["outfile"] == " ": # no outfile given
		params["outfile"] == ""
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


def convertToColLabels(paramDict, dframe): #currently unused?
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


def convertToColNums(paramDict, dframe):
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
				paramDict[key] = int(value) - 1
				continue
			elif len(str(value)) == 1 and str(value).isalpha():
				debug(key, "Column is given as letter")
				paramDict[key] = int("abcdefghijklmnopqrstuvwxyz".index(value.lower()))
				continue
			elif len(str(value)) > 1:
				debug(key, "Column is given by label")
				try:
					paramDict[key] = int(list(df1.columns).index(value))
				except:
					print(f"Can't find the column {value} you gave as parameter {key}.")
					print("Perhaps you misspelled it?")
					print("Remember: Column labels are Case Sensitive!")
					sys.exit(1)
			else:
				log("Can't make shaft nor bolt from this.", "")
	return paramDict


# main program
###############

# check and parse the command line
params = parseParams()
verbosity = int(params["verbose"]) + int(params["debug"])
log(f"Parameters: {params}")

# load files
log(f"File 1: {params['infile']}")
if params["infile"].lower().endswith(".csv"):
	df1 = pd.read_csv(params["infile"],
	  sep=";",
	  header=params["startrow"] - 2,
	  engine="pyarrow",
	  escapechar="Ŧ",
	  skipinitialspace=True)
else:
	# not a .csv, so probly some sort of officeware spreadsheet...
	# pandas.read_excel can digest them all
	# first we get the column names...
	inputColumns = pd.read_excel(
	  params["infile"],
	  header=(params["startrow"] - 2),
	  sheet_name=params["sheet"] - 1,
	  nrows=0,
	  verbose=True
	  ).columns
	# ...then we read the whole table, converting all ";" to "," and "|" to "~"
	# (see function "maskSpecialChars")
	df1 = pd.read_excel(
	  params["infile"],
	  header=(params["startrow"] - 2),
	  sheet_name=params["sheet"] - 1,
	  converters={col: maskSpecialChars for col in inputColumns},
	  verbose=True)

# guess or ask for the parameters we didn't get from CL
helpMessageDisplayed = False
# first make sure we have understood what the startcol arg wants to tell us:
if str(params["startcol"]).isdecimal(): # start column given as column number
	params["startcol"] = int(params["startcol"]) # just to make sure
	metaCols = df1.iloc[ : 1, : params["startcol"]].columns
elif len(str(params["startcol"])) == 1: # start column given as column letter
	startColNum = "abcdefghijklmnopqrstuvwxyz".index(str(param["startcol"]).lower()) + 1
	metaCols = df1.iloc[ : 1, : startColNum].columns
else: # start column probably given by label
	try:
		startColNum = list(df1.columns).index(params["startcol"])
		metaCols = df1.iloc[ : 1, : startColNum].columns
	except:
		metaCols = df1.columns
		log("proceeding anyway...", "Could not identify start column")
# now we file through the parameters:
for s in params.keys():
	log(params[s], s)
	if params[s] == "" or params[s] == None:
		if helpMessageDisplayed == False:
			# a parameter was not given on CL
			# and we haven't yet printed the help message, so we do now:
			print(params["help"])
			helpMessageDisplayed = True # we won't show this again
		params[s] = guessOrGet(s, metaCols)
		log(params[s], "guess or get user input")
del params["help"]

askForFilenames()
log(params, "Final config:")

debug(df1.columns, "Columns of Infile")
debug(type(df1.columns), "Type of Column list")
debug(df1.columns.array, "Columns of Infile by Index no.")
debug(df1.columns.values, "Columns of Infile by Value")

# sanitize column names or letters to column numbers
params = convertToColNums(params, df1)
debug(params, "Sanitized:")

# let's begin to construct the output table!
dfOut = df1.iloc[ : , params["name"] : params["name"]]
dfOut = dfOut.join(df1.iloc[ : , params["genetic_group"]], how="left", rsuffix="XXX")
dfOut = dfOut.join(df1.iloc[ : , params["reference"]], how="left", rsuffix="XXX")
dfOut = dfOut.join(df1.iloc[ : , params["trueness"]], how="left", rsuffix="XXX")
dfOut = dfOut.join(df1.iloc[ : , params["munq"]], how="left", rsuffix="XXX")
dfOut = dfOut.join(df1.iloc[ : , params["ploidy"]], how="left", rsuffix="XXX")
dfOut = dfOut.join(df1.iloc[ : , params["includes_mutations"]], how="left", rsuffix="XXX")

# let's rename the columns to Genomvergleicher2 standard
dfOut = dfOut.rename(columns={
  params["name"]: "Name",
  params["genetic_group"]: "Gene Group",
  params["reference"]: "Reference",
  params["trueness"]: "Trueness to Type",
  params["munq"]: "MUNQ",
  params["ploidy"]: "Ploidy",
  params["includes_mutations"]: "Includes Mutations"
  })

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
  mode=params["overwrite"],
  sep=";",
  index=False,
#  decimal=",",
  escapechar="Ŧ"
  )

print("Done.")
log(params["outfile"], "Wrote output to file")







