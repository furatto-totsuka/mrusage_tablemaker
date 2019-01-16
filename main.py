from datetime import datetime, timedelta
from uwstyle.dialogs import dialog, select
from uwstyle.excel import Excel


def load():
  def dividetime(baseday, timestr):
    ts = list(map(
      lambda p:
        list(map(lambda n: int(n), p.split(":"))),
      timestr.split("～")
    ))
    return {
      "start": baseday + timedelta(hours=ts[0][0], minutes=ts[0][1]),
      "end": baseday + timedelta(hours=ts[1][0], minutes=ts[1][1]),
    }
  excel = Excel()
  wb = excel.chooseworkbook()
  if wb is None:
    return None
  ws = wb.chooseworksheet()
  if ws is None:
    return None
  table = ws.cells(1, 1).currentRegion
  datas = []
  first = True
  for r in table.rows:
    if first:
      first = False
      continue
    if r.cells(3).mergecells or r.cells(5).value == "定休日":
      continue
    # Collect Items
    data = {
      "week": r.cells(2).value,
      "name": r.cells(4).value,
      "category": r.cells(6).value,
      "note": r.cells(7).value,
    }
    # Special Items
    d = r.cells(1).value #pywin32type/datetime
    data["day"] = datetime(d.year, d.month, d.day)
    data["time"] = dividetime(data["day"], r.cells(3).value)
    data["resv"] = dividetime(data["day"], r.cells(5).value) if r.cells(5).value is not None else data["time"]
    datas.append(data)

  return datas

if __name__ == "__main__":
  data = load()
  if data is None:
    exit