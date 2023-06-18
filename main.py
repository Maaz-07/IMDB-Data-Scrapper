import requests
from bs4 import BeautifulSoup
import xlwt

# Function to scrape the movie data
def scrape_movies():
    url = "https://www.imdb.com/chart/top"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    movies = []
    
    # Find the movie data within the table
    table = soup.find("table", {"class": "chart full-width"})
    rows = table.find_all("tr")
    
    # Iterate through each row and extract the relevant information
    for row in rows[1:]:  # Skip the header row
        cells = row.find_all("td")
        name = cells[1].find("a").text
        genre = cells[1].find("span", {"class": "secondaryInfo"}).text.strip("() ")
        year = cells[1].find("span", {"class": "secondaryInfo"}).next_sibling.strip("() ")
        rating = cells[2].find("strong").text
        movies.append((name, genre, year, rating))
    
    return movies

# Function to save movie data to an Excel file
def save_to_excel(movies):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Top Movies")
    
    # Write column headers
    headers = ["Name", "Genre", "Year", "Rating"]
    for col, header in enumerate(headers):
        sheet.write(0, col, header)
    
    # Write movie data
    for row, movie in enumerate(movies, start=1):
        for col, value in enumerate(movie):
            sheet.write(row, col, value)
    
    # Save the Excel file
    workbook.save("top_movies.xls")

# Scrape the movies and save the data to Excel
movies_data = scrape_movies()
save_to_excel(movies_data)

print("Movie data has been successfully scraped and saved to 'top_movies.xls' file.")
