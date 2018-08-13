import xlrd
import xlwt
import time
import datetime
import uszipcode
import urllib.request

book = xlrd.open_workbook(r"C:\Users\Connor\Documents\TestAPIRequestBook.xlsx")
sheetRead = book.sheet_by_index(0)

databook = xlrd.open_workbook(r"C:\Users\Connor\Documents\TestAPIRawData.xls")
dataRead = databook.sheet_by_index(0)

workbook = xlwt.Workbook()
sheetWrite = workbook.add_sheet('test')

workbookData = xlwt.Workbook()
dataWrite = workbookData.add_sheet('test')

print("Starting " + str(sheetRead.nrows-1) + " rows calculations")

#Duplicate read workbook
for i in range(sheetRead.nrows):
    for p in range(sheetRead.ncols):
        if sheetRead.cell_value(i,p) != "":
            sheetWrite.write(i, p, sheetRead.cell_value(i,p))

for i in range(dataRead.nrows):
    for p in range(dataRead.ncols):
        if dataRead.cell_value(i,p) != "":
            dataWrite.write(i, p, dataRead.cell_value(i,p))

UnixTimeData = []
for i in range(dataRead.nrows):
    UnixTimeData.append(dataRead.cell_value(i,0))

print(UnixTimeData)

search = uszipcode.ZipcodeSearchEngine()

WrittenIterations = 0

for i in range(sheetRead.nrows-1):
    print("StartingRow" + str(i+1))
    Longitude = search.by_zipcode(str(int(sheetRead.cell_value(i + 1, 0)))).Longitude
    Lattertude = search.by_zipcode(str(int(sheetRead.cell_value(i + 1, 0)))).Latitude
    sheetWrite.write(i + 1, 4, Longitude)
    sheetWrite.write(i + 1, 5, Lattertude)

    Day = str(int(sheetRead.cell_value(i+1,1)))
    Month = str(int(sheetRead.cell_value(i + 1, 2)))
    Year = str(int(sheetRead.cell_value(i + 1, 3)))

    UnixTime = time.mktime(datetime.datetime.strptime(Day+"/"+Month+"/"+Year, "%d/%m/%Y").timetuple())

    sheetWrite.write(i + 1, 6, UnixTime)

    if(UnixTime in UnixTimeData):
        print("Data Found On Computer")
        contents = dataRead.cell_value(UnixTimeData.index(UnixTime),1)

    else:
        contents = urllib.request.urlopen("https://api.darksky.net/forecast/8066119e9963349cca1c07b7b9740e45/" + str(Lattertude) + "," + str(Longitude) + "," + str(int(UnixTime))).read()
        print("Data Found Online")
        contents = contents.decode("utf-8")

        dataWrite.write(dataRead.nrows + WrittenIterations, 0, UnixTime)
        dataWrite.write(dataRead.nrows + WrittenIterations, 1, contents)

        WrittenIterations = WrittenIterations + 1

        workbookData.save(r"C:\Users\Connor\Documents\TestAPIRawData.xls")

    contentssplit = contents.split(",")

    Titles = contentssplit
    # Filter data
    for Char in '1234567890- [{}":.]':
        Titles = [s.replace(Char, '') for s in Titles]
    Titles = [s.replace('summary', '') for s in Titles]

    RemoveChar = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQ[RSTUVWXYZ['{]}:/"

    for Char in RemoveChar:
        contentssplit = [s.replace(Char, '') for s in contentssplit]

    contentssplit = [s.replace('"', '') for s in contentssplit]

    sheetWrite.write(i + 1, 7, Titles[4])   # Summary

    try:
        sheetWrite.write(i + 1, 8, contentssplit[Titles.index("precipIntensityMax")])   # PrecipIntensityHigh
    except ValueError:
        sheetWrite.write(i + 1, 8, "")  # PrecipIntensityHigh

    try:
        sheetWrite.write(i + 1, 9, contentssplit[Titles.index("precipProbability")])   # PrecipProbabilityHigh
    except ValueError:
        sheetWrite.write(i + 1, 9, "")  # PrecipProbabilityHigh

    try:
        sheetWrite.write(i + 1, 10, contentssplit[Titles.index("temperatureHigh")])   # TempratureHigh
    except ValueError:
        sheetWrite.write(i + 1, 10, "")  # TempratureHigh

    try:
        sheetWrite.write(i + 1, 11, contentssplit[Titles.index("temperatureLow")])  # TempratureLow
    except ValueError:
        sheetWrite.write(i + 1, 11, "")  # TempratureLow

    try:
        sheetWrite.write(i + 1, 12, contentssplit[Titles.index("apparentTemperatureHigh")])  # ApparentTempratureHigh
    except ValueError:
        sheetWrite.write(i + 1, 12, "")  # ApparentTempratureHigh

    try:
        sheetWrite.write(i + 1, 13, contentssplit[Titles.index("apparentTemperatureLow")])  # ApparentTempratureLow
    except ValueError:
        sheetWrite.write(i + 1, 13, "")  # ApparentTempratureLow

    try:
        sheetWrite.write(i + 1, 14, contentssplit[Titles.index("dewPoint")])  # DewPoint
    except ValueError:
        sheetWrite.write(i + 1, 14, contentssplit[Titles.index("dewPoint")])  # DewPoint

    try:
        sheetWrite.write(i + 1, 15, contentssplit[Titles.index("humidity")])  # Humidity
    except ValueError:
        sheetWrite.write(i + 1, 15, contentssplit[Titles.index("humidity")])  # Humidity

    try:
        sheetWrite.write(i + 1, 16, contentssplit[Titles.index("pressure")])  # Pressure

        indices = [pp for pp, x in enumerate(Titles) if x == "pressure"]
        lowest = float(contentssplit[indices[0]])
        highest = float(contentssplit[indices[0]])
        for i2 in indices:
            if (highest < float(contentssplit[i2])):
                highest = float(contentssplit[i2])
            if (lowest > float(contentssplit[i2])):
                lowest = float(contentssplit[i2])
        Change = highest - lowest
        sheetWrite.write(i + 1, 17, Change)  # PressureChange
    except ValueError:
        sheetWrite.write(i + 1, 16, "")  # Pressure
        sheetWrite.write(i + 1, 17, "")  # PressureChange

    try:
        sheetWrite.write(i + 1, 18, contentssplit[Titles.index("windSpeed")])  # WindSpeed
    except ValueError:
        sheetWrite.write(i + 1, 18, "")  # WindSpeed

    try:
        sheetWrite.write(i + 1, 19, contentssplit[Titles.index("windGust")])  # WindGust
    except ValueError:
        sheetWrite.write(i + 1, 19, "")  # WindGust

    try:
        sheetWrite.write(i + 1, 20, contentssplit[Titles.index("windBearing")])  # WindBearing
    except ValueError:
        sheetWrite.write(i + 1, 20, "")  # WindBearing

    try:
        sheetWrite.write(i + 1, 21, contentssplit[Titles.index("cloudCover")])  # CloudCover
    except ValueError:
        sheetWrite.write(i + 1, 21, "")  # CloudCover

    try: #Some past things don't have uv or ozone numbers

        indices = [pp for pp, x in enumerate(Titles) if x == "uvIndex"]
        if (len(indices) > 0):
            UVHighest = float(contentssplit[indices[0]])
            for i2 in indices:
                if (UVHighest < float(contentssplit[i2])):
                    UVHighest = float(contentssplit[i2])
            sheetWrite.write(i + 1, 22, UVHighest)  # uvIndex
        else:
            sheetWrite.write(i + 1, 22, "")

    except ValueError:
        sheetWrite.write(i + 1, 22, "")

    try:
        sheetWrite.write(i + 1, 23, contentssplit[Titles.index("visibility")])  # Visibility
    except ValueError:
        sheetWrite.write(i + 1, 23, "")  # Visibility

    try:
        sheetWrite.write(i + 1, 24, contentssplit[Titles.index("ozone")])  # Ozone
    except ValueError:
        sheetWrite.write(i + 1, 24, "")

    try:
        sheetWrite.write(i + 1, 25, contentssplit[Titles.index("moonPhase")])  # moonPhase
    except ValueError:
        sheetWrite.write(i + 1, 25, "")  # moonPhase

    print("Saving")
    workbook.save(r"C:\Users\Connor\Documents\TextBookWrite.xls")
    print("Saved")

print("Saving Final")
workbook.save(r"C:\Users\Connor\Documents\TextBookWrite.xls")
print("Saved Final")
print("Complete")