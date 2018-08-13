import xlrd
import xlwt
import time
import datetime
import uszipcode
import urllib.request

# Open base workbook
book = xlrd.open_workbook(r"C:\Users\Connor\Documents\TestAPIRequestBook.xlsx")
sheetRead = book.sheet_by_index(0)

# Open offline data workbook
databook = xlrd.open_workbook(r"C:\Users\Connor\Documents\TestAPIRawData.xls")
dataRead = databook.sheet_by_index(0)

# Create output workbook
workbook = xlwt.Workbook()
sheetWrite = workbook.add_sheet('test')

# Create output offline data workbook
workbookData = xlwt.Workbook()
dataWrite = workbookData.add_sheet('test')

print("Starting " + str(sheetRead.nrows-1) + " rows calculations")

# Duplicate read workbook
for i in range(sheetRead.nrows):
    for p in range(sheetRead.ncols):
        # If cell is empty don't copy any data since you can only write to each cell once
        if sheetRead.cell_value(i, p) != "":
            sheetWrite.write(i, p, sheetRead.cell_value(i, p))

# Duplicate offline data workbook
for i in range(dataRead.nrows):
    for p in range(dataRead.ncols):
        # If cell is empty don't copy any data since you can only write to each cell once
        if dataRead.cell_value(i, p) != "":
            dataWrite.write(i, p, dataRead.cell_value(i, p))

# Make an array of all the unix times stored in offline data workbook
UnixTimeData = []
for i in range(dataRead.nrows):
    UnixTimeData.append(dataRead.cell_value(i, 0))

print(UnixTimeData)

# Define what search is
search = uszipcode.ZipcodeSearchEngine()

# Count number written to offline data workbook, used later to determine what line to write to in offline data workbook
WrittenIterations = 0

# Iterate through each row in base workbook
for i in range(sheetRead.nrows-1):
    print("StartingRow" + str(i+1))
    # Determine the longitude and lattitude and write them to the output workbook
    Longitude = search.by_zipcode(str(int(sheetRead.cell_value(i + 1, 0)))).Longitude
    Latitude = search.by_zipcode(str(int(sheetRead.cell_value(i + 1, 0)))).Latitude
    sheetWrite.write(i + 1, 4, Longitude)
    sheetWrite.write(i + 1, 5, Latitude)

    # Get Day Month and Year from the workbook
    Day = str(int(sheetRead.cell_value(i+1, 1)))
    Month = str(int(sheetRead.cell_value(i + 1, 2)))
    Year = str(int(sheetRead.cell_value(i + 1, 3)))

    # Use time package to calculate Unix time
    UnixTime = time.mktime(datetime.datetime.strptime(Day+"/"+Month+"/"+Year, "%d/%m/%Y").timetuple())

    # Write the unix time to the output workbook
    sheetWrite.write(i + 1, 6, UnixTime)

    # If the unix time is in the offline data workbook use the data from the offline data workbook
    if UnixTime in UnixTimeData:
        print("Data Found On Computer")
        contents = dataRead.cell_value(UnixTimeData.index(UnixTime), 1)

    # If the unix time isn't in the offline data workbook request it from the website
    else:
        contents = urllib.request.urlopen("https://api.darksky.net/forecast/8066119e9963349cca1c07b7b9740e45/" +
                                          str(Latitude) + "," + str(Longitude) + "," + str(int(UnixTime))).read()
        print("Data Found Online")
        contents = contents.decode("utf-8")

        # Write the new data that was found to the offline data workbook
        dataWrite.write(dataRead.nrows + WrittenIterations, 0, UnixTime)
        dataWrite.write(dataRead.nrows + WrittenIterations, 1, contents)

        WrittenIterations = WrittenIterations + 1

        workbookData.save(r"C:\Users\Connor\Documents\TestAPIRawData.xls")

    # Filter data into into 2 arrays Titles containing names, and contentSplit containing numbers
    contentSplit = contents.split(",")
    Titles = contentSplit

    # Remove all numbers from Titles, also remove summary before summary title
    for Char in '1234567890- [{}":.]':
        Titles = [s.replace(Char, '') for s in Titles]
    Titles = [s.replace('summary', '') for s in Titles]

    # Remove all characters from content split
    RemoveChar = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQ[RSTUVWXYZ['{]}:/"
    for Char in RemoveChar:
        contentSplit = [s.replace(Char, '') for s in contentSplit]

    # remove " characters from content split
    contentSplit = [s.replace('"', '') for s in contentSplit]

    # Write summary title to output sheet
    sheetWrite.write(i + 1, 7, Titles[4])

    # Try:
    #     Write to output sheet the first number that has the title of "Something"
    # If there is a error
    #     Write nothing

    # PrecipitationIntensityHigh
    try:
        sheetWrite.write(i + 1, 8, contentSplit[Titles.index("precipIntensityMax")])
    except ValueError:
        sheetWrite.write(i + 1, 8, "")

    # PrecipitationProbabilityHigh
    try:
        sheetWrite.write(i + 1, 9, contentSplit[Titles.index("precipProbability")])
    except ValueError:
        sheetWrite.write(i + 1, 9, "")

    # TempratureHigh
    try:
        sheetWrite.write(i + 1, 10, contentSplit[Titles.index("temperatureHigh")])
    except ValueError:
        sheetWrite.write(i + 1, 10, "")

    # TempratureLow
    try:
        sheetWrite.write(i + 1, 11, contentSplit[Titles.index("temperatureLow")])
    except ValueError:
        sheetWrite.write(i + 1, 11, "")

    # ApparentTempratureHigh
    try:
        sheetWrite.write(i + 1, 12, contentSplit[Titles.index("apparentTemperatureHigh")])
    except ValueError:
        sheetWrite.write(i + 1, 12, "")

    # ApparentRempratureLow
    try:
        sheetWrite.write(i + 1, 13, contentSplit[Titles.index("apparentTemperatureLow")])
    except ValueError:
        sheetWrite.write(i + 1, 13, "")

    # DewPoint
    try:
        sheetWrite.write(i + 1, 14, contentSplit[Titles.index("dewPoint")])
    except ValueError:
        sheetWrite.write(i + 1, 14, "")

    # Humidity
    try:
        sheetWrite.write(i + 1, 15, contentSplit[Titles.index("humidity")])
    except ValueError:
        sheetWrite.write(i + 1, 15, "")

    # Pressure
    try:
        sheetWrite.write(i + 1, 16, contentSplit[Titles.index("pressure")])

        # Find the indices of all things titled pressure
        indices = [pp for pp, x in enumerate(Titles) if x == "pressure"]

        # Initialize the lowest and highest to the first number in the list
        lowest = float(contentSplit[indices[0]])
        highest = float(contentSplit[indices[0]])

        # Iterate through the indices and
        # replace highest if i2 is higher than highest and lower if i2 is lower than lowest
        for i2 in indices:
            if highest < float(contentSplit[i2]):
                highest = float(contentSplit[i2])
            if lowest > float(contentSplit[i2]):
                lowest = float(contentSplit[i2])

        # Calculate the change and write it to output sheet
        Change = highest - lowest
        sheetWrite.write(i + 1, 17, Change)
    except ValueError:
        sheetWrite.write(i + 1, 16, "")
        sheetWrite.write(i + 1, 17, "")

    # WindSpeed
    try:
        sheetWrite.write(i + 1, 18, contentSplit[Titles.index("windSpeed")])
    except ValueError:
        sheetWrite.write(i + 1, 18, "")

    # WindGust
    try:
        sheetWrite.write(i + 1, 19, contentSplit[Titles.index("windGust")])
    except ValueError:
        sheetWrite.write(i + 1, 19, "")

    # WindBearing
    try:
        sheetWrite.write(i + 1, 20, contentSplit[Titles.index("windBearing")])
    except ValueError:
        sheetWrite.write(i + 1, 20, "")

    # CloudCover
    try:
        sheetWrite.write(i + 1, 21, contentSplit[Titles.index("cloudCover")])
    except ValueError:
        sheetWrite.write(i + 1, 21, "")

    # UVIndex
    try:  # Find the highest uv index

        # Find indicies of things titled uvIndex
        indices = [pp for pp, x in enumerate(Titles) if x == "uvIndex"]

        # If something was found titled uvindex
        if len(indices) > 0:

            # Initialize highest value to first value in indices
            UVHighest = float(contentSplit[indices[0]])

            # Iterate through indices, replace highest if i2 is higher than highest
            for i2 in indices:
                if UVHighest < float(contentSplit[i2]):
                    UVHighest = float(contentSplit[i2])

            # Write to output workbook
            sheetWrite.write(i + 1, 22, UVHighest)
        else:
            sheetWrite.write(i + 1, 22, "")

    except ValueError:
        sheetWrite.write(i + 1, 22, "")

    # Visibility
    try:
        sheetWrite.write(i + 1, 23, contentSplit[Titles.index("visibility")])
    except ValueError:
        sheetWrite.write(i + 1, 23, "")

    # Ozone
    try:
        sheetWrite.write(i + 1, 24, contentSplit[Titles.index("ozone")])
    except ValueError:
        sheetWrite.write(i + 1, 24, "")

    # Moon Phase
    try:
        sheetWrite.write(i + 1, 25, contentSplit[Titles.index("moonPhase")])
    except ValueError:
        sheetWrite.write(i + 1, 25, "")

    # Save output workbook after each line
    print("Saving")
    workbook.save(r"C:\Users\Connor\Documents\TextBookWrite.xls")
    print("Saved")

# Save at the end
print("Saving Final")
workbook.save(r"C:\Users\Connor\Documents\TextBookWrite.xls")
print("Saved Final")
print("Complete")
