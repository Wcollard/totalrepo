import requests
from bs4 import BeautifulSoup

def get_patent_details(google_url):
    try:
        response = requests.get(google_url, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 1. Get Abstract
            abstract = "Abstract not found"
            meta = soup.find('meta', attrs={'name': 'DC.description'})
            if meta and meta.get('content'):
                abstract = meta['content'].strip()
            else:
                abstract_div = soup.find(lambda tag: tag.name in ["div", "section"] and "abstract" in (tag.get('class', []) + [tag.get('id', '')]))
                if abstract_div:
                    abstract = abstract_div.get_text().strip()
            
            # 2. Get Title
            title = "Title not found"
            title_tag = soup.find('meta', attrs={'name': 'DC.title'})
            if title_tag:
                title = title_tag.get('content', '').strip()

            # 3. Get Inventor
            inventor = "Inventor not found"
            inventor_tag = soup.find('meta', attrs={'name': 'DC.contributor'})
            if inventor_tag:
                inventor = inventor_tag.get('content', '').strip()

            # 4. Get Assignee
            assignee = "Assignee not found"
            assignee_tag = soup.find('meta', attrs={'name': 'DC.assignee'})
            if assignee_tag:
                assignee = assignee_tag.get('content', '').strip()


            # 5. Get Publication Date
            publication_date = "Publication date not found"
            pub_date_tag = soup.find('meta', attrs={'name': 'DC.date'})
            if pub_date_tag:
                publication_date = pub_date_tag.get('content', '').strip()

            return {
                "abstract": abstract,
                "title": title,
                "inventor": inventor,
                "assignee": assignee,
                "publication_date": publication_date
            }
        return "Failed to retrieve patent details"
    except Exception as e:
        return f"Error: {e}"

# Example usage
url = "https://patents.google.com/patent/US20040143644A1/en"
details = get_patent_details(url)
print(details)