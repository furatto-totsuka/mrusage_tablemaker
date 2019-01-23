from datetime import datetime, timedelta
from datetimerange import DateTimeRange
from uwstyle.dialogs import dialog, select
from uwstyle.excel import Excel

WD_CLOSED = 3
xlThemeColorAccent2 = 6
xlCenter = -4108
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
  for r in table.rows:
    if first:
      first = False
      continue
    if r.cells(3).mergecells or r.cells(5).value == "定休日":
      continue
    # Collect Items
    data = {
      "week": r.cells(3).value,
      "name": r.cells(5).value,
      "category": r.cells(7).value,
      "note": r.cells(8).value,
    }
    # Special Items
    d = r.cells(1).value #pywin32type/datetime
    data["day"] = datetime(d.year, d.month, d.day)
    data["time"] = dividetime(data["day"], r.cells(4).value)
    data["resv"] = dividetime(data["day"], r.cells(6).value) if r.cells(6).value is not None else data["time"]
    print("{0:%m/%d} {1}".format(data["day"], data["name"]))
    if data["note"] is None or not "ホール" in data["note"]:
    datas.append(data)

  return datas

def create_sheet(excel, data):
  timelist = list(map(lambda i: i["resv"], data))
  d = data[0]["day"]
  print("> データ出力中")
  start = datetime(d.year, d.month, 1)
  ws = excel.add_workbook().sheets[0]
  ws.name = "{0}月会議室利用状況".format(start.month)
  # Header
  ws.cells(1, 1).value = "日"
  d = t = start
  t += timedelta(hours=10)
  if "header":
    r = 2
    while t.hour < 17:
      ws.cells(1, r).value = "{0:%H:%M}".format(t)
      ws.cells(1, r).horizontalAlignment = xlCenter
      r += 1
      t += timedelta(minutes=30)
    ws.cells(1, r).value = "備考"

  if "rows":
    c = 2
    while d.month == start.month:
      print("{0:%m/%d}".format(d))
      ws.cells(c, 1).value = "{0:%Y/%m/%d}".format(d)
      ws.cells(c, 1).numberFormatLocal = "m/d(aaa);@"
      if d.weekday() == WD_CLOSED:
        ws.cells(c, 16).value = "休日"
        ws.cells(c, 16).horizontalAlignment = xlCenter
        ws.raw.Range(ws.cells(c, 2), ws.cells(c, 16)) \
          .interior.themeColor = xlThemeColorAccent2
        ws.raw.Range(ws.cells(c, 2), ws.cells(c, 16)) \
          .interior.tintAndShade = 0.599993896298105
      else:
        t = d
        t += timedelta(hours=10)
        r = 2
        while t.hour < 17:
          ws.cells(c, r).value = "●" \
            if any(t in tl for tl in timelist) else "○"
          ws.cells(c, r).horizontalAlignment = xlCenter
          r += 1
          t += timedelta(minutes=30)

      c += 1
      d += timedelta(days=1)

    ws.cells(c + 1, 2).value = "凡例"
    ws.cells(c + 2, 2).value = "●＝予約あり"
    ws.cells(c + 3, 2).value = "○＝予約なし"

  lo = ws.raw.listObjects.add(1, ws.cells(1, 1).currentRegion, None, 1)
  lo.showAutoFilterDropDown = False
  print("> Finished")



if __name__ == "__main__":
  excel = Excel()
  data = load(excel)
  if data is None:
    exit
  excel.excel.screenUpdating = False
  try:
    create_sheet(excel, data)
  finally:
    excel.excel.screenUpdating = True