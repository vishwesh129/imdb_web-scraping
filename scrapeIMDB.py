from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
# print(excel.sheetnames)

sheet = excel.active
sheet.title = 'Top Rated Movie List'
# print(excel.sheetnames)
sheet.append(['RANK' , 'MOVIE NAME' , 'YEAR OF RELEASE' , ' IMDB RATING'])


url = "https://www.imdb.com/chart/top"
try:
    pageSource = requests.get(url , headers={'user-Agent' : 'Mozilla/5.0'})
    pageSource.raise_for_status()

    soup = BeautifulSoup(pageSource.text, 'html.parser')

    movies = soup.find('ul' , class_ = 'ipc-metadata-list ipc-metadata-list--dividers-between sc-3a353071-0 wTPeg compact-list-view ipc-metadata-list--base').find_all('li')

    for movie in movies:

        movie_name = movie.find('h3' , class_ = "ipc-title__text").text.split('. ')[1]

        rank = movie.find('h3' , class_ = "ipc-title__text").text.split('.')[0]

        year = movie.find('span' , class_='sc-14dd939d-6 kHVqMR cli-title-metadata-item').text

        rating = movie.find('span' , class_='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').text

        print(rank, movie_name , year, rating)
        sheet.append([rank, movie_name , year, rating])

except Exception as e:
    print(e)

excel.save("IMDB_MOVIES_LIST.xlsx")

