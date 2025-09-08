from serpapi import GoogleSearch

patent= input ("what is the patent number?")

params = {
  "engine": "google_patents_details",
  "patent_id": f"patent/{patent}/en",
  "api_key": "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
}

search = GoogleSearch(params)
results = search.get_dict()

print (results)