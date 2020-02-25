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


def GetValues(sheet, col, row, end):
  """
  Return a nested dictionary of values from the target sheet
  """
  # create object of objects that contain the value to update as the key, and the values and xl cell as the details
  # ie: {123456: {xlRow:144, xlCol:5,color: "blue", type:"sale", angle:90}, 890128:{xlRow:145, xlCol:5, color: "red", type:"return", angle:180}}
  destDict = {}

  for i in range(row, end):
    val = sheet.cell(i,col).value
    destDict[i] = val
  
  return destDict


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

#program loop
run = True
while run:
  Home()
  selection = GetInput("Please make a selection", "I'm sorry, that's not a valid selection.  Please try again")
  if selection == "exit":
    exit()
  if selection == 1:
    #get master sheet
    masterBook = GetBook("What is the name of the MASTER book? ")
    masterSheet = GetSheet(masterBook, "What is the name of the target sheet in the MASTER book? ")
    maxRow_ = GetInput("Please enter the row you want to read the excel sheet to ", "Sorry, that is not a valid row. Please try again... ")
    masterRowStart_ = GetInput("Which row would you like to start on?", "Sorry that's not a valid row.  Please try again ")
  elif selection == 2:
    #get child sheet
    childBook = GetBook("What is the name of the CHILD book? ")
    childSheet = GetSheet(childBook, "What is the name of the target sheet in the CHILD book? ")
  elif selection == 3:
    #consume child sheet
      # column(s) to match against master (index column)
      # column(s) (values) to copy/paste into master
      pass
  elif selection == 4:
    # print master values
     pass
  elif selection == 5:
    # print child values
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


def update_xlsx(wb, wb2, ws, ws2):
  #get the last excel entry in first sheet
  destEnd = ws.max_row +1
  
  #get the last exccel entry in the second sheet
  srcEnd = ws2.max_row +1

  srcDict = {}

  # create object of objects that contain the value to update as the key, and the values and xl cell as the details
  # ie: {123456: {xlRow:144, xlCol:5,color: "blue", type:"sale", angle:90}, 890128:{xlRow:145, xlCol:5, color: "red", type:"return", angle:180}}

  for i in range(1, destEnd):
    val = ws.cell(i,1).value
    destDict[i] = val

  for i in range(1, srcEnd):
    val = ws2.cell(i,1).value
    if val not in destDict:
      ws.cell(i,1).value = val
  
  wb.save(filename=dest)
  print("saved updated workbook")

# update_xlsx(wb, wb2, ws, ws2)