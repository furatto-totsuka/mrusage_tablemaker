from uwstyle.dialogs import dialog, select
from uwstyle.excel import Excel

def main():
  excel = Excel()
  wb = excel.chooseworkbook()
  if wb is None:
    exit
  ws = wb.chooseworksheet()
  if ws is None:
    exit
  print(ws.name)


if __name__ == "__main__":
  main()