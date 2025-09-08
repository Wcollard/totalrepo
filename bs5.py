from bs4 import BeautifulSoup
import requests

# Example URL of a Google Patent page
url = "https://patents.google.com/patent/US10325221B2/en"
response = requests.get(url)

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(response.content, 'html.parser')

# Find the <a> tag containing the text "Download PDF"
pdf_link_tag = soup.find('a', string='Download PDF')

# Extract the href attribute if the tag is found
pdf_link = pdf_link_tag['href'] if pdf_link_tag else "PDF link not found"

print(f"PDF link: {pdf_link}")