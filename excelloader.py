from datetime import datetime, timedelta
from datetimerange import DateTimeRange
import re

def load(excel):
  def dividetime(baseday, timestr):
    ts = list(map(
      lambda p:
        list(map(lambda n: int(n), p.split(":"))),
      timestr.split("～")
    ))
    return DateTimeRange(
      baseday + timedelta(hours=ts[0][0], minutes=ts[0][1]),
      baseday + timedelta(hours=ts[1][0], minutes=ts[1][1]))

  wb = excel.chooseworkbook()
  if wb is None:
    return None
  ws = wb.chooseworksheet()
  if ws is None:
    return None
  table = ws.cells(1, 1).currentRegion
  datas = []
  first = True
  print("> データ読み込み中")
  holiday = re.compile('定\s*休\s*日')
  for r in table.rows:
    if first:
      first = False
      continue
    if r.cells(4).value is None or holiday.match(r.cells(4).value):
      continue
    # Collect Items
    data = {
      "week": r.cells(3).value,
      "name": r.cells(5).value,
      "category": "" if r.cells(7).value is None else r.cells(7).value,
      "note": r.cells(8).value,
    }
    # Special Items
    d = r.cells(1).value #pywin32type/datetime
    data["day"] = datetime(d.year, d.month, d.day)
    data["time"] = dividetime(data["day"], r.cells(4).value)
    data["resv"] = dividetime(data["day"], r.cells(6).value) if r.cells(6).value is not None else data["time"]
    print("{0:%m/%d} {1}".format(data["day"], data["name"]))
    datas.append(data)

  return datas

if __name__ == "__main__":
  from uwstyle.excel import Excel
  excel = Excel()
  data = load(excel)
