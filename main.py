from uwstyle.dialogs import dialog, select

app = None

def main():
  global app
  app = Excel()
  wb = app.chooseworkbook()
  if wb is None:
    exit
  ws = app.chooseworksheet(wb)
  if ws is None:
    exit
  print(ws.name)

import win32com.client

class Excel:
  """
  Require win32com module(pip install pywin32)
  """
  def __init__(self, active=True):
    """
    Get Microsoft Excel Object

    Parameters
    ----
    active: bool
      True to get Excel that is already booting.
      False if necessary to start a new Excel.

    Exceptions
    ----
    com_error
      Occurs when Excel is not running when active is true.
    """
    if active:
      self.excel = win32com.client.GetActiveObject("Excel.Application")
    else:
      self.excel = win32com.client.Dispatch("Excel.Application")

  @property
  def workbooks(self):
    """
    Get a list of open workbooks.

    Returns
    ----
    list: list(Workbook)
      Workbook object list.
    """
    result = []
    for i in range(1, self.excel.workbooks.count + 1):
      result.append(self.excel.workbooks(i))
    return result

  def getsheets(self, workbook):
    """
    Get a list of all the sheet objects in the workbook.

    Parameters
    ----
    workbook: Workbook
      Target workbook.

    Returns
    ----
    list: list(Sheet)
      Sheet list.
    """
    result = []
    for i in range(1, workbook.sheets.count + 1):
      result.append(workbook.sheets(i))
    return result

  ### Utility Methods.

  def chooseworkbook(self, message="Choose Workbook"):
    """
    Select a workbook.
    When two or more workbooks are open, a dialog is displayed to prompt the user for selection.
    If there is only one workbook open, it will return the file name without displaying anything.

    Parameters
    ----
    message: str
      The message that is displayed in the dialog.

    Returns
    ----
    workbook: Workbook|None
      The Workbook object. None if the dialog has been cancelled.
    """
    wb = self.workbooks
    files = [f.name for f in wb]
    if len(files) == 1:
      return wb[0]
    else:
      r = select(message, files)
      if r is None:
        return None
      else:
        return self.excel.workbooks(r[1])

  def chooseworksheet(self, workbook, message="Choose Worksheet"):
    """
    Select a worksheet.
    When two or more worksheets are open, a dialog is displayed to prompt the user for selection.
    If there is only one worksheet open, it returns the name of the worksheet without displaying anything.

    Parameters
    ----
    workbook: Workbook
      Workbook object.
    message: str
      The message that is displayed in the dialog.

    Returns
    ----
    worksheet: Sheet|None
      the worksheet object. None if the dialog has been cancelled.
    """
    ws = self.getsheets(workbook)
    sheets = [s.name for s in ws]
    if len(sheets) == 1:
      return ws[0]
    else:
      r = select(message, sheets)
      if r is None:
        return None
      else:
        return workbook.sheets(r[1])

if __name__ == "__main__":
  main()