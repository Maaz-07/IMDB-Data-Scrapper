import requests
from bs4 import BeautifulSoup
import xlwt

# Function to scrape the movie names
def scrape_movie_names():
    url = "https://www.imdb.com/chart/top"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    movie_names = []
    
    # Find the movie names within the table
    table = soup.find("table", {"class": "chart full-width"})
    rows = table.find_all("tr")
    
    # Iterate through each row and extract the movie name
    for row in rows[1:]:  # Skip the header row
        name = row.find("td", {"class": "titleColumn"}).find("a").text
        movie_names.append(name)
    
    return movie_names

# Function to save movie names to an Excel file
def save_to_excel(movie_names):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Top Movie Names")
    
    # Write column header
    sheet.write(0, 0, "Movie Name")
    
    # Write movie names
    for row, name in enumerate(movie_names, start=1):
        sheet.write(row, 0, name)
    
    # Save the Excel file
    workbook.save("top_movie_names.xls")

# Scrape the movie names and save them to Excel
movie_names_data = scrape_movie_names()
save_to_excel(movie_names_data)

print("Movie names have been successfully scraped and saved to 'top_movie_names.xls' file.")
