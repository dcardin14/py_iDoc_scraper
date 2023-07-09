import requests
from bs4 import BeautifulSoup

# Set the URL of the website
url = "https://www.idocmarket.com/RIOCO1/Document/Search"

# Set your credentials
username = "dardin14@gmail.com"
password = "sumtoor99A+"

# Create a session object to persist cookies across requests
session = requests.Session()

# Send a GET request to the website to initiate a session
response = session.get(url)

# Create a BeautifulSoup object to parse the HTML content
soup = BeautifulSoup(response.content, "html.parser")

# Find the form input fields for username and password
username_field = soup.find("input", attrs={"name": "Username"})
password_field = soup.find("input", attrs={"name": "Password"})

# Update the values of the username and password fields
username_field["value"] = username
password_field["value"] = password

# Submit the login form
response = session.post(url, data=soup.form.attrs)

# Retrieve the search results page
response = session.get(url)

# Create a new BeautifulSoup object to parse the search results page
soup = BeautifulSoup(response.content, "html.parser")

# Find and extract the desired information from the parsed content
search_results = soup.find_all("div", class_="resultTitle")
urls = [result.a["href"] for result in search_results]

# Process the extracted data as per your needs
for url in urls:
    print(url)

