import openpyxl as xl

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


def GetInput(message, err):
  """
  Get Input in a while loop
  """  
  stop = False
  while stop == False:
    try:
      userInput = input(message)
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


masterBook = GetBook("What is the name of the MASTER book? ")
masterSheet = GetSheet(masterBook, "What is the name of the target sheet in the MASTER book? ")
childBook = GetBook("What is the name of the CHILD book? ")
childSheet = GetSheet(childBook, "What is the name of the target sheet in the CHILD book? ")

maxRow_ = GetInput("Please enter the row you want to read the excel sheet to ", "Sorry, that is not a valid row. Please try again... ")



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


def GetValues(book, start=1, end=maxRow_):
  """
  Return a nested dictionary of values from the target sheet
  """


def update_xlsx(wb, wb2, ws, ws2):
  #get the last excel entry in first sheet
  destEnd = ws.max_row +1

  #get the last exccel entry in the second sheet
  srcEnd = ws2.max_row +1

  destDict = {}
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

update_xlsx(wb, wb2, ws, ws2)