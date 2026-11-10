import xlrd
from pathlib import Path
import dateutil
import datetime
import sys
from openpyxl import Workbook, load_workbook

def process_xlsx(file, lud):
  try:
    wb = load_workbook(file, read_only=True)
  except:
    print("file open exception: ", file.name)
    return False
  for sheet in wb:
    try:
      d = dateutil.parser.parse(sheet.title)
      if(d > lud):
        print("Needs update: ", file.name, sheet.title, d)
    except dateutil.parser.ParserError:
      print("Error parsing date", file.name, sheet.title)
  return True

def process_xls(file, lud):
  try:
    wb = xlrd.open_workbook(file)
  except:
    print("file open exception: ", file.name)
    return False
  for sheet in wb.sheet_names():
    try:
      d = dateutil.parser.parse(sheet)
      if(d > lud):
        print("Needs update: ", file.name, sheet, d)
    except dateutil.parser.ParserError:
      print("Error parsing date", file.name, sheet)
  return True

def main():
  argv = sys.argv
  if len(argv) < 3:
    print("script accepts 2 mandatory arguments: wbs-job directory_to_traverse last_update_date_file")
    exit(1)
  p = Path(argv[1])
  last_updated = Path(argv[2])
  if not p.exists() or not p.is_dir() or not last_updated.exists() or not last_updated.is_file:
    print("first argument must be an existing directory")
    print("second argument must be a text file")
    exit(1)

  lus = last_updated.read_text()
  lud = dateutil.parser.parse(lus)

  files = list(p.glob('*.xls*'))
  for file in files:
    if file.name.endswith("xls"):
      if not process_xls(file, lud):
        continue
    else:
      if not process_xlsx(file, lud):
        continue
if __name__ == '__main__':
  main()