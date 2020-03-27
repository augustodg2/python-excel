import openpyxl
import glob

def search_for(key, dataframe):
  for row in dataframe.rows:
    for cell in row:
      if cell.value == key:
        return cell.column_letter + str(cell.row)
  return False


files = glob.glob(r"B:\Desktop\*\*.xlsx")

for file in files:
  workbook = openpyxl.load_workbook(file)
  dataframe = workbook.active

  email = search_for('value', dataframe)
  print email

  if email:
    dataframe[email].value = 'new_value'
    print dataframe[email].value

    workbook.save(file)
    print 'Saved as ' + file
