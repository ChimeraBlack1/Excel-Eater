import openpyxl as xl

def Home():
  """
                  Home
  -------------------------------------
  1) Add a Master sheet
        - what is the name of the book
        - what is the name of the sheet
        - column(s) to match against child (index column)
  2) Add Child Sheet
        - what is the name of the book
        - what is the name of the sheet
        - column(s) to match against master (index column)
        - column(s) (values) to copy/paste into master
  3) Consume Child Sheet
        - overwrite master entries?
        - save updated master sheet
  4) Print Master values
  5) Print Child values
  6) Print values in Child not in Master
  7) Print values in Master not in Child
  8) Print values consumed by Master (and from which child)

  *** TYPE EXIT TO STOP THE PROGRAM *** 

  """

  print("                Home                ")
  print("-------------------------------------")
  print("1) Add a Master sheet")
  print("2) Add a Child sheet")
  print("3) Consume Child Sheet")
  print("4) Print Master values")
  print("5) Print Child values")
  print("-------------------------------------")
  print("*** TYPE EXIT TO STOP THE PROGRAM ***")

def FindLastRow(sheet, row=0, col=0):
  """
  Find the number of populated rows in a given column of an excel worksheet
  """  
  content = sheet.cell(row, col).value
  rowCount = 0

  while content != "":
    try:
      content = sheet.cell_value(row + rowCount, col)
    except:
      break
    rowCount = rowCount + 1
    
  return rowCount


def GetValues(sheet, whichCol, index, start, end):
  """
  Return a nested dictionary of values from the target sheet
  create object of objects that contain the value to update as the key, and the values and xl cell as the details
  ie: {123456: {xlRow:144, xlCol:5,color: "blue", type:"sale", angle:90}, 890128:{xlRow:145, xlCol:5, color: "red", type:"return", angle:180}}
  """
  # print("this is whichCol: " + str(whichCol))
  destDict = {}

  for i in range(start, end):
    indexVal = sheet.cell(i,index).value
    valueDict = {}
    for j in whichCol:
      valueDict[whichCol[j]] = sheet.cell(i,j).value
    
    destDict[indexVal] = valueDict
  return destDict


def ConsumeChild():
  pass

def GetInput(message, err):
  """
  Get Input in a while loop
  """  
  stop = False
  while stop == False:
    print(message)
    try:
      userInput = input("> ")
      stop = True
    except:
      print(err)

  return userInput
  

def GetBook(prompt):
  """
  Return WorkBook
  """
  destination = False
  while destination == False:
    master = input(prompt)
    try:
      Book = xl.load_workbook(filename=master)
      destination = True
    except:
      print("Sorry, I couldn't find that filename in the current directory.  Did you forget the extension? xlsx, xlsm, etc...)> ")
  return Book


def GetSheet(book, prompt):
  """
  Return target sheet
  """
  worksheetID = False
  while worksheetID == False:
    worksheet = input(prompt)
    try:
      sheet = book[worksheet]
      worksheetID = True
    except: 
      print("Sorry, that sheet does not exist. Try again... ")
  return sheet


def ValidateMaxRow(sheet):
  """
  Did the user enter a valid ENDING cell?
  """
  valid = False
  while valid == False:
    maxRow_ = GetInput("Please enter the row you want to read the excel sheet to ", "Sorry, that is not a valid row. Please try again... ")
    if maxRow_ == "":
      print("Searching entire document...")
      maxRow_ = sheet.max_row +1
      valid = True
    try:
      maxRow_ = int(maxRow_)
      valid = True
    except:
      print("you must enter a valid number")
      continue
    if maxRow_ < 1:
      print("Searching entire document...")
      maxRow_ = sheet.max_row +1
      valid = True
  return maxRow_



def ValidateRowStart():
  """
  Did the user enter a valid STARTING cell?
  """
  valid = False
  while valid == False:
    rowStart_ = GetInput("Which row would you like to start on?", "Sorry that's not a valid row.  Please try again ")
    if rowStart_ == "":
      print("Starting at row 1...")
      rowStart_ = 1
      valid = True
    try:
      rowStart_ = int(rowStart_)
      valid = True
    except:
      print("you must enter a valid number")
      continue
    if rowStart_ <= 0:
      print("Starting at row 1...")
      rowStart_ = 0
      valid = True
  return rowStart_


def ValidateCol():
  """
  Did the user enter a VALID INDEX COLUMN?
  """
  valid = False
  while valid == False:
    colStart_ = GetInput("Which column are we going to use for the index?", "Sorry that's not a valid row.  Please try again ")
    if colStart_ == "":
      print("Using Column 1...")
      colStart_ = 1
      valid = True
    try:
      colStart_ = int(colStart_)
      valid = True
    except:
      print("you must enter a valid number")
      continue
    if colStart_ < 1:
      print("You did not specify a valid number. The program will guess at the which row is the START")
      colStart_ = 1
      valid = True
  return colStart_


def GetColumns(colCount):
  """
  Return an object containing the names and locations of the columns to retrieve data from
  """
  goodData = False
  whichCol = {}
  while goodData == False:
    for i in range(0, colCount):
      colName = GetInput("Name column " +str(i+1), "Invalid column name")
      colNum = GetInput("Enter Column number for value " +str(i+1), "Invalid column number")
      try:
        colNum = int(colNum)
        goodData = True
      except:
        print("Invalid column number...")
      whichCol[colNum] = colName
    
  return whichCol

def UploadSuccessful(sheet):
  if sheet == "master".lower():
    print("***************************************")
    print("*Master document uploaded successfully*")
    print("***************************************")
  if sheet == "child".lower():
    print("**************************************")
    print("*Child document uploaded successfully*")
    print("**************************************")

def PrintValues():
  """
  Print the values in a more readable way
  """
  pass
  

#master globals
masterBook = {}
masterSheet = {}
masterValues = {}
masterCol = {}
masterRowStart_ = 1
maxRow_ = 1

#child globals
childBook = {}
childSheet = {}
childDict = {}
childCol = {}
childRowStart_ = 1
childMaxRow_ = 1

#program loop
run = True
while run:
  Home()
  selection = str(GetInput("Please make a selection", "I'm sorry, that's not a valid selection.  Please try again"))
  print("you entered: " + selection)
  if selection.lower() == "exit":
    exit()
  if selection == "1":
    #get master sheet
    masterBook = GetBook("What is the name of the MASTER book? ")
    masterSheet = GetSheet(masterBook, "What is the name of the target sheet in the MASTER book? ")
    maxRow_ = ValidateMaxRow(masterSheet)
    masterRowStart_ = ValidateRowStart()
    masterColIndex = ValidateCol()
    colCount = GetInput("how many columns are we taking values from?","Sorry that's not a valid number.  Try again.")
    # colCount = int(colCount)
    try:
      colCount = int(colCount)
    except:
      print("please enter a valid number")
      continue

    whichCol = GetColumns(colCount)

    masterValues = GetValues(masterSheet, whichCol, masterColIndex, masterRowStart_, maxRow_)
    UploadSuccessful(sheet="master")
  elif selection == "2":
    #get child sheet
    childBook_ = GetBook("What is the name of the CHILD book? ")
    childSheet_ = GetSheet(childBook_, "What is the name of the target sheet in the CHILD book? ")
    maxRow_ = ValidateMaxRow(childSheet_)
    childRowStart_ = ValidateRowStart()
    childCol_ = ValidateCol()

    childValues = GetValues(childSheet_, childCol_, childRowStart_, maxRow_)
    UploadSuccessful(sheet="child")
  elif selection == "3":
    #consume child sheet
      # column(s) to match against master (index column)
      # column(s) (values) to copy/paste into master
    pass
  elif selection == "4":
    # print master values
    try:
      print(masterValues)
    except:
      print("no Master sheet has been uploaded yet")
  elif selection == "5":
    # print child values
    try:
      print(childValues)
    except:
      print("no Child sheet has been uploaded yet...")
      #select child
    pass

# masterValues = GetValues(masterSheet, )

# 1) - Get the last row in the worksheet
# 2) - Get the last row in a given column 
# 3) - Provide the last row manually

#MASTER
#what's the file type? xls, xlsx, xlsm?
#how many columns are we comparing to the child documents?
  #what letters are they?
#which row to start in on the master
  #go right to end?
  #specify end?

#CHILD
#how many col are we searching through in the child
  #what letters are they
#which col do the target values live in, in the child?

#CONSUME
# how many values be combined in the child?
  #if > 0, which letter ranges should be combined?
# which columns are we putting the values into in the master?

#OPTIONS
#ignore overwrite values in master from child?


# def update_xlsx(wb, wb2, ws, ws2):
#   #get the last excel entry in first sheet
#   destEnd = ws.max_row +1
  
#   #get the last exccel entry in the second sheet
#   srcEnd = ws2.max_row +1

#   srcDict = {}

#   # create object of objects that contain the value to update as the key, and the values and xl cell as the details
#   # ie: {123456: {xlRow:144, xlCol:5,color: "blue", type:"sale", angle:90}, 890128:{xlRow:145, xlCol:5, color: "red", type:"return", angle:180}}

#   for i in range(1, destEnd):
#     val = ws.cell(i,1).value
#     destDict[i] = val

#   for i in range(1, srcEnd):
#     val = ws2.cell(i,1).value
#     if val not in destDict:
#       ws.cell(i,1).value = val
  
#   wb.save(filename=dest)
#   print("saved updated workbook")

# update_xlsx(wb, wb2, ws, ws2)