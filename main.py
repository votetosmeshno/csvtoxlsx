import os, csv, datetime
from openpyxl import Workbook
import time


def parseToObj(id, filepath):
    csv_reader = openCSV(filepath)
    csv_headings = next(csv_reader)
    s = csv_headings[0]

    data = s.split(" ")
    result = {}
    result["id"] = id
    result["symbol"] = data[1]
    try:
        data[5]
        result["action"] = data[2] + " " + data[3]
        result["tradeSize"] = data[4]
        result["price"] = data[5].replace(".", ",")
    except IndexError:
        result["action"] = data[2]
        result["tradeSize"] = data[3]
        result["price"] = data[4].replace(".", ",")

    # creation time
    stat = os.stat(filepath)
    creationTime = 1
    try:
        creationTime = stat.st_birthtime
    except AttributeError:
        creationTime = stat.st_mtime
    year, month, day, hour, minute, second = time.localtime(creationTime)[:-3]

    result["creation_date"] = "%02d/%02d/%d " % (day, month, year)
    result["creation_time"] = "%02d:%02d:%02d" % (hour, minute, second)

    return result


def openCSV(filepath):
    f = open(filepath, newline='')
    csv_reader = csv.reader(f)
    return csv_reader


def main():
    infolder = './in/'
    outfolder = './out/'

    infiles = os.listdir(infolder)
    # array of maps
    toWrite = []
    i = 1
    for filename in infiles:  # по всем файлам
        print("\n", filename)
        if filename == '.DS_Store':
            continue
        filepath = infolder + filename

        # appending created dict to the list we write to excel
        obj = parseToObj(i, filepath)
        if not (obj["action"].lower() == "sell" or obj["action"].lower() == "buy"):
            continue

        toWrite.append(obj)
        i += 1

    print("Got " + str(len(toWrite)) + " records")

    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "ID"
    sheet["B1"] = "Trade Date"
    sheet["C1"] = "Trade Time"
    sheet["D1"] = "Symbol"
    sheet["E1"] = "Action"
    sheet["F1"] = "Trade Size"
    sheet["G1"] = "Price"

    for obj in toWrite:
        newId = str(obj["id"] + 1)
        sheet["A" + newId] = obj["id"]
        sheet["B" + newId] = obj["creation_date"]
        sheet["C" + newId] = obj["creation_time"]
        sheet["D" + newId] = obj["symbol"]
        sheet["E" + newId] = obj["action"]
        sheet["F" + newId] = obj["tradeSize"]
        sheet["G" + newId] = obj["price"]

    outputName = outfolder + "result_out_{0}.xlsx".format(datetime.datetime.now().strftime("%m-%d-%Y,%H:%M:%S"))
    workbook.save(filename=outputName)


main()