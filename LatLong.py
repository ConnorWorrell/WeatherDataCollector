import uszipcode

search = uszipcode.ZipcodeSearchEngine()
zipcode = search.by_zipcode("80513")

print(zipcode.Latitude)
print(zipcode.Longitude)