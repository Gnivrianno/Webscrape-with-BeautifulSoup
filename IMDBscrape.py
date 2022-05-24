from csv import excel
from unicodedata import ucd_3_2_0
from bs4 import BeautifulSoup

import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating', 'Total User Ratings'])




try:
    source = requests.get('https://www.imdb.com/chart/top/?ref_=nv_mv_250')
    source.raise_for_status()

    soup = BeautifulSoup(source.text,'html.parser')
    

    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    for movie in movies:

        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        rating = movie.find('td', class_="ratingColumn").strong.text
        userrating = movie.find('td', class_="ratingColumn").strong['title']
        userrating = userrating[13:]
        userrating = userrating.replace(" ","")
        userrating = userrating.replace("userratings","")
        userrating = userrating.replace(",","")
        userrating = int(userrating)
       
        print(userrating)

        sheet.append([rank, name, year, rating, userrating])








except Exception as e:
    print(e)


excel.save('IMDB Movie Rating.xlsx')
