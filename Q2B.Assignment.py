from bs4 import BeautifulSoup
import requests , openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])


try:
    source = requests.get('https://www.imdb.com/chart/boxoffice/')
    source.raise_for_status()#throws an error of Client Error

    soup = BeautifulSoup(source.text,'html.parser')
    # print(soup)
    
    movies = soup.find('tbody').find_all('tr') #class_ always be written in this form
    # print(len(movies))
    # movies = int(input("Enter the n value"))
    for movie in movies:
        
        name = movie.find('td', class_="titleColumn").a.text

        weekend = movie.find('td', class_="ratingColumn").get_text(strip=True)#.split('.')[0]

        gross = movie.find('td', class_="weeksColumn").get_text(strip=True)#.strip('()')

        # crew = movie.find([a.attrs.get('title') for a in movie.select('td.titleColumn a')])
        crew = movie.find('td',class_="titleColumn").a.attrs.get('title')

        
        print((name, weekend, gross, crew))
        sheet.append([name, weekend, gross, crew])
            
        # break
        # print(year)
        # print(rating)
        


except Exception as e:
    print(e)
excel.save('IMDB Movie.xlsx')