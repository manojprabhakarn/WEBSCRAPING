from bs4 import BeautifulSoup
import requests
import openpyxl


# select the active worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = 'Companies'

# add some data to the worksheet
worksheet['A1'] = 'CIN'
worksheet['B1'] = 'Company Name'
worksheet['C1'] = 'City'
worksheet['D1'] = 'Status'

# send GET request to first page
url = 'https://www.zaubacorp.com/company-list/nic-D/city-BANGALORE/company-type-FTC-company.html'
response = requests.get(url)

# create BeautifulSoup object
soup = BeautifulSoup(response.text, 'html.parser')
# print(soup)

# initialize list to store table data
data = []
i=1
# loop over pages
try: 
    while True:
        # extract table data from current page
        i=i+1
        table = soup.find('table')
        # print(table)
        rows = table.find_all('tr')
        for row in rows:
            cells = row.find_all('td')
            data.append([cell.text for cell in cells])

        # check if there is a next page
        next_page_link = soup.find('a', {'rel': 'nofollow'})
        if next_page_link:
            # send GET request to next page
            url = 'https://www.zaubacorp.com/company-list/nic-D/city-BANGALORE/company-type-FTC/p-'+str(i)+'-company.html'
            print(i)
            response = requests.get(url)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            if (len(response.content))>80000:
                continue
            else:
                print(".....page exists......")
                break

        else:
            print("...finished....")
            break
except KeyboardInterrupt:
    print("....interrupted ......... sorry")

# print the extracted data
for row in data:
    worksheet.append(row)

workbook.save("companies.xlsx")