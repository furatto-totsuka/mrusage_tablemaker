import csv
from uwstyle.excel import Excel
import excelloader

def writecsv(data):
  print("> Create CSV")
  header = [
    "Start Date",
    "Start Time",
    "End Date",
    "End Time",
    "Subject",
    "Location",
    "Description"
  ]
  items = [header]
  for d in data:
    item = []
    item.append(d["day"].strftime("%Y/%m/%d"))
    item.append(d["time"].start_datetime.strftime("%H:%M"))
    item.append(d["day"].strftime("%Y/%m/%d"))
    item.append(d["time"].end_datetime.strftime("%H:%M"))
    item.append(d["name"].encode('cp932', "ignore").decode('cp932'))
    item.append("ふらっとステーション・とつか")
    item.append(d["category"].encode('cp932', "ignore").decode('cp932'))
    items.append(item)
  with open("import.csv", "w") as f:
    writer = csv.writer(f, lineterminator="\n")
    for item in items:
      writer.writerow(item)
  print("> Finished")

if __name__ == "__main__":
  excel = Excel()
  data = excelloader.load(excel)
  if data is None:
    exit
  writecsv(data)
