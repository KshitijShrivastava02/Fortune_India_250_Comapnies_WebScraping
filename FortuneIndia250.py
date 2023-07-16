from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title="Top 250 companies of India"
print(excel.sheetnames)
sheet.append(['serial No.' , 'Company name' , 'Industry type' , 'Total Income'])


try:
    source =requests.get('https://www.fortuneindia.com/fortune-500/company-listing/?year=2022&page=1&query=&per_page=500')
    source.raise_for_status()

    soup = BeautifulSoup(source.text,'html.parser')
    
    company_name=soup.find('tbody').find_all('tr')
    print(len(company_name))
    
    for company in company_name:

        serialNum=company.find('td' ,class_="f-500-row-td serial-num").text
        Company_name=company.find('td' , class_="company-name-container f-500-row-td").a.get_text(strip=True).split(' ')[0]
        Industry=company.find('td' , class_="f-500-row-td industry").text
        TotalIncome=company.find('td' , class_="f-500-row-td align-right").text
        #Netprofit=company.find('td' , class_="f-500-row-td").text


        print(serialNum, Company_name, Industry, TotalIncome)
        
        sheet.append([serialNum , Company_name , Industry , TotalIncome])

except Exception as e:
    print(e)

excel.save('Top 250 companies of India.xlsx')
