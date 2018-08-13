import urllib.request
import xlrd
import xlwt

#Read file
book = xlrd.open_workbook(r"C:\Users\Connor\Documents\TextBook.xlsx")
sheet = book.sheet_by_index(0)

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('test')

#find weather for day
contents = urllib.request.urlopen("https://api.darksky.net/forecast/8066119e9963349cca1c07b7b9740e45/40.3083,-105.0811,1532368895").read()
contentssplit = contents.decode("utf-8").split(",")

Titles = contentssplit
#Filter data
for Char in '1234567890- [{}":.]':
    Titles = [s.replace(Char, '') for s in Titles]

RemoveChar = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQ[RSTUVWXYZ['{]}:/"

for Char in RemoveChar:
    contentssplit = [s.replace(Char, '') for s in contentssplit]

contentssplit = [s.replace('"', '') for s in contentssplit]

#Write to file

print(len(Titles))
print(len(contentssplit))

for i in range(len(Titles)):
    if(Titles[i] == "datatime" or Titles[i] == "time" or Titles[i] == "dailydatatime" or Titles[i] == "currentlytime"):
        print("")
    if(i < 255):
        sheet.write(0,i+1,Titles[i])
        sheet.write(1,i+1,contentssplit[i])
    print(str(i) + " " + Titles[i] + " " + contentssplit[i])

indices = [i for i, x in enumerate(Titles) if x == "pressure"]
lowest = float(contentssplit[indices[0]])
highest = float(contentssplit[indices[0]])
for i in indices:
    if (highest < float(contentssplit[i])):
        highest = float(contentssplit[i])
    if (lowest > float(contentssplit[i])):
        lowest = float(contentssplit[i])
    print(highest)
    print(lowest)

Change = highest - lowest
print(Change)
