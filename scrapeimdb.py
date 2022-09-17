import requests,openpyxl
from bs4 import BeautifulSoup

excel=openpyxl.Workbook()
print(excel.sheetnames)

#activate the sheet as active sheet
sheet=excel.active

#give a title to excel sheet
sheet.title="IMDB Top Ranked Movies"
print(excel.sheetnames)

#create the column heading
sheet.append(['Movie rank','Movie Name','Movie Released Year','Movie Rating '])

try:
    source=requests.get("https://www.imdb.com/chart/top/")
    source.raise_for_status()

    soup=BeautifulSoup(source.text,'html.parser')
    
    movies= soup.find('tbody',class_='lister-list').find_all('tr')
    for mov in movies:
        name=mov.find('td',class_='titleColumn').a.text
        rank=mov.find('td',class_='titleColumn').get_text(strip='True').split('.')[0]
        year=mov.find('td',class_='titleColumn').span.text.strip('()')
        rating=mov.find('td',class_="ratingColumn imdbRating").strong.text  
            
        print(rank, name, year, rating)
        #add the column values
        sheet.append([rank, name, year, rating])


except Exception as err:
    print(err)

#save the excel file
excel.save("C:/Users/Lenovo/youtubescraper/IMDB Movie Rating.xlsx")
    