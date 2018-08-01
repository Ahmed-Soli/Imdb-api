
import xlsxwriter # used to export data to excel
import  requests # used to open links and get their source code
from bs4 import  BeautifulSoup # used to parse html 
import os # used to get the directory 
import json # used to handle json response 


# Defining a function that take the imdb id 
def Exportexcel(id):

    print('Connecting ....') # printing on the screen the word connecting ... 

    r = requests.get("http://www.omdbapi.com/?apikey=48ebd72b&i=" + id) # ssending get request to the api 

    print('Reading content ...')  # printing on the screen the word connecting ... 

    soup = BeautifulSoup(r.content, 'html.parser')  # reading the request response 

    json_data = json.loads(str(soup))  # reading the json response 

    file_path = os.getcwd()  +'/sheet.csv'  # get current path and add to it the file sheet.csv

    print('Creating sheet ...') # printing on the screen the word Creating ... 

   
    workbook =  xlsxwriter.Workbook(file_path)  # Create a workbook and add a worksheet.

    worksheet = workbook.add_worksheet('Sheet 1')  # add a worksheet.
    
    bold = workbook.add_format({'bold': 1}) # Add a bold format to use to highlight cells.
    
    row = 1 # the cell under headers cell.
    col = 0 # first column.

    Title = json_data['Title'] # geting the title value out of the json response

    worksheet.write('A1', 'Title', bold)  # creating a head with the value of title
    worksheet.write_string( row , col , Title )  # adding the value of title to the second cell after the header of title
    print('Title added ...') # printing on the screen the word connecting ... 

    ReleaseYear = json_data['Year'] 

    worksheet.write('B1', 'ReleaseYear', bold) 
    worksheet.write_string( row , col + 1 , ReleaseYear ) 
    print('ReleaseYear added ...') 

    ReleaseDate = json_data['Released']

    worksheet.write('C1', 'ReleaseDate', bold)
    worksheet.write_string( row , col + 2 , ReleaseDate )
    print('ReleaseDate added ...') 

    Runtime = json_data['Runtime']

    worksheet.write('D1', 'Runtime', bold)
    worksheet.write_string( row , col +3, Runtime )
    print('Runtime added ...') 


    Genre = json_data['Genre']

    worksheet.write('E1', 'Genre', bold)
    worksheet.write_string( row , col + 4 , Genre )
    print('Genre added ...') 

    Director = json_data['Director']

    worksheet.write('F1', 'Director', bold)
    worksheet.write_string( row , col + 5 , Director )
    print('Director added ...') 

    Writer = json_data['Writer']

    worksheet.write('G1', 'Writer', bold)
    worksheet.write_string( row , col + 6 , Writer )
    print('Writer added ...') 
 
    Actors = json_data['Actors']

    worksheet.write('H1', 'Actors', bold)
    worksheet.write_string( row , col + 7 , Actors )
    print('Actors added ...') 

    Plot = json_data['Plot']

    worksheet.write('I1', 'Plot', bold)
    worksheet.write_string( row , col + 8 , Plot )
    print('Plot added ...') 

    Language = json_data['Language']

    worksheet.write('J1', 'Language', bold)
    worksheet.write_string( row , col + 9 , Language )
    print('Language added ...') 

    Country = json_data['Country']

    worksheet.write('K1', 'Country', bold)
    worksheet.write_string( row , col + 10 , Country )
    print('Country added ...') 

    Awards = json_data['Awards']

    worksheet.write('L1', 'Awards', bold)
    worksheet.write_string( row , col + 11 , Awards )
    print('Awards added ...') 

    Poster = json_data['Poster']

    worksheet.write('M1', 'Poster', bold)
    worksheet.write_string( row , col + 12, Poster )
    print('Poster added ...') 

    Ratings = json_data['imdbRating']

    worksheet.write('N1', 'Ratings', bold)
    worksheet.write_string( row , col + 13 , Ratings )
    print('Ratings added ...') 

    imdbID = json_data['imdbID']

    worksheet.write('O1', 'imdbURL', bold)
    worksheet.write_string( row , col + 14 , 'https://www.imdb.com/title/' + url_id )
    print('imdbURL added ...') 

    Type = json_data['Type']

    worksheet.write('P1', 'Type', bold)
    worksheet.write_string( row , col + 15 , Type )
    print('Type added ...') 

    if 'BoxOffice' in json_data:
        BoxOffice = json_data['BoxOffice']

        worksheet.write('Q1', 'BoxOffice', bold)
        worksheet.write_string( row , col + 16 , BoxOffice )
        print('BoxOffice added ...') 

        Production = json_data['Production']

        worksheet.write('R1', 'Production', bold)
        worksheet.write_string( row , col + 17 , Production )
        print('Production added ...') 

        Website = json_data['Website']

        worksheet.write('S1', 'Website', bold)
        worksheet.write_string( row , col + 18 , Website )
        print('Website added ...') 

    else:
        totalSeasons = json_data['totalSeasons']

        worksheet.write('Q1', 'totalSeasons', bold)
        worksheet.write_string( row , col + 16 , totalSeasons )
        print('totalSeasons added ...') 
    



    x = 0
    while x < 20 :
        worksheet.set_column(1,x, 15) # set the width of the columns in the excel sheet to 15
        x = x + 1

    workbook.close() # close the excel sheet 
    print('sheet.csv Exported Successfully')


url = input("Enter the url : ") # url in the format of 'https://www.imdb.com/title/tt6452574'

if '?' in url and 'title' in url:
    new_url = url.split('title/')
    url_id = new_url[1].split('/')
    if 'tt' in url_id[0]:
        url_id = url_id[0]
    else:
        url_id = url_id[1]    
    Exportexcel(url_id)
elif '?' not in url and 'title' in url:
    url_id = url.split('title/')
    url_id = url_id[1]
    Exportexcel(url_id)
else:
    print('sorry invalid link')

