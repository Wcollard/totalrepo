from serpapi import GoogleSearch

params = {
  "engine": "google_patents",
  "q": "(US6111111)",
  "api_key": "7bf2aaaeab13938ea4fc3920bbde495841f0877f96803a1dc060447b0091867d"
}

search = GoogleSearch(params)
results = search.get_dict()
organic_results = results["organic_results"]
print (results)