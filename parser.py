import unicodecsv as csv
import requests
import xlrd

COUNTIES = {
    "AND": "Androscoggin",
    "ARO": "Aroostook",
    "CUM": "Cumberland",
    "FRA": "Franklin",
    "HAN": "Hancock",
    "KEN": "Kennebec",
    "KNO": "Knox",
    "LIN": "Lincoln",
    "OXF": "Oxford",
    "PEN": "Penobscot",
    "PIS": "Piscataquis",
    "SAG": "Sagadahoc",
    "SOM": "Somerset",
    "WAL": "Waldo",
    "WAS": "Washington",
    "YOR": "York"
}

results = []

office = 'U.S. House'
for party in ['REP', 'DEM']:
    url = "http://www.maine.gov/sos/cec/elec/results/2016/Rep%20to%20Congress%20" + party + ".xlsx"
    r = requests.get(url)
    workbook = xlrd.open_workbook(file_contents=r.content)
    for ws in workbook.sheets():
        candidates = ws.row_values(0)[2:]
        for row in range(2, ws.nrows):
            if ws.row_values(row)[0] == '':
                continue
            county = COUNTIES[ws.row_values(row)[0]]
            district = ws.name
            for candidate in candidates:
                results.append([county, ws.row_values(row)[1].strip(), office, district, party, candidate, ws.row_values(row)[candidates.index(candidate)+2]])

office = 'State Senate'
for party in ['REP', 'DEM']:
    print party
    url = "http://www.maine.gov/sos/cec/elec/results/2016/State%20Senate%20" + party + ".xlsx"
    r = requests.get(url)
    workbook = xlrd.open_workbook(file_contents=r.content)
    for ws in workbook.sheets():
        for row in range(0, ws.nrows):
            if ws.row_values(row)[0] == '' and ws.row_values(row)[4]:
                continue
            if 'Total' in ws.row_values(row)[1] or 'TOTAL' in ws.row_values(row)[1]:
                candidates = []
                continue
            if ws.row_values(row)[1] == '' or ws.row_values(row)[1].strip() == 'Town':
                continue
            if ws.row_values(row)[1].strip() == 'CTY':
                candidates = ws.row_values(row)[3:]
            elif ws.row_values(row)[1] == '':
                county = None
            else:
                county = COUNTIES[ws.row_values(row)[1]]
            district = ws.row_values(row)[0]
            for candidate in candidates:
                if ws.row_values(row)[candidates.index(candidate)+3] == '':
                    continue
                if ws.row_values(row)[candidates.index(candidate)+3] == candidate:
                    continue
                try:
                    results.append([county, ws.row_values(row)[2].strip(), office, district, party, candidate, ws.row_values(row)[candidates.index(candidate)+3]])
                except:
                    print ws.row_values(row)
                    print candidates
                    raise


with open('2016/20160614__me__primary__town.csv','wb') as csvfile:
        csvwriter = csv.writer(csvfile, encoding='utf-8')
        csvwriter.writerow(['county','town', 'office', 'district', 'party', 'candidate', 'votes'])
        csvwriter.writerows(results)
