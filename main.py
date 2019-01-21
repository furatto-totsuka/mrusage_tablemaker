from datetime import datetime, timedelta
from uwstyle.dialogs import dialog, select
from uwstyle.excel import Excel


def load(excel):
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
    print("{0:%m/%d} {1}".format(data["day"], data["name"]))
    datas.append(data)

  return datas

def create_sheet(excel, data):
  d = data[0]["day"]
  print("> データ出力中")
  start = datetime(d.year, d.month, 1)
  ws = excel.add().sheets[0]
  ws.name = "{0}月会議室利用状況".format(start.minute)
  # Header
  ws.cells(1, 1).value = "日"
  d = t = start
  t += timedelta(hours=10)
  for i in range(14):
    ws.cells(1, 3 + i).value = "{0:%H:%M}".format(t)
    t += timedelta(minutes=30)

  ws.cells(1, 18).value = "備考"


if __name__ == "__main__":
  excel = Excel()
  data = load(excel)
  if data is None:
    exit
  create_sheet(excel, data)