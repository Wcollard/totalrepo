
#Altcode5.py
 # Create workbook and worksheet
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Add formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'border': 1
    })

    abstract_format = workbook.add_format({
        'text_wrap': True,
        'valign': 'top',
        'border': 1
    })


    link_format = workbook.add_format({
        'underline': True,
        'color': 'blue',
        'border': 1
    })

    # Set column widths
    column_widths = [20, 20, 25, 15, 15, 15, 15, 70, 50]
    for col, width in enumerate(column_widths):
        worksheet.set_column(col, col, width)

    # Write headers
    headers = ["Ref No.", "Inventor", "Title", "Publication Date", "Google Link", "Espacenet Link", "USPTO Link", "ABSTRACT", "NOTES"]
    worksheet.write_row(0, 0, headers, header_format)

    # Write data
    row = 1
    for number in patent_numbers:
        cleaned_number = number.strip()
        if not cleaned_number:
            continue

        google_url = f"https://patents.google.com/patent/{cleaned_number}/en"
        espacenet_url = f"https://worldwide.espacenet.com/patent/search?q={cleaned_number}"
        uspto_number = cleaned_number.replace("US", "")
        uspto_url = f"https://ppubs.uspto.gov/pubwebapp/external.html?q={uspto_number}.pn."

        title, inventor, publication_date, abstract = get_patent_details(google_url)

        # Write data with appropriate formats
        worksheet.write(row, 0, cleaned_number)
        worksheet.write(row, 1, inventor)
        worksheet.write(row, 2, title, abstract_format)
        worksheet.write(row, 3, publication_date)
        worksheet.write_url(row, 4, google_url, link_format, string=cleaned_number)
        worksheet.write_url(row, 5, espacenet_url, link_format, string=cleaned_number)
        worksheet.write_url(row, 6, uspto_url, link_format, string=cleaned_number)
        worksheet.write(row, 7, abstract, abstract_format)

        # Set row height based on abstract content
        text_lines = len(abstract) // 70 + abstract.count('\n') + 1  # Rough estimate
        row_height = min(text_lines * 15, 409)  # 409 is Excel's
        worksheet.set_row(row, row_height)

        row += 1

#Serpapi10.py
def write_to_excel(data, filepath):
    df = pd.DataFrame(data)
    print(df.head())  # Debugging line to check DataFrame contents
    
    with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Patents')
        workbook  = writer.book
        worksheet = writer.sheets['Patents']
        
        # Set column widths
        worksheet.set_column('A:A', 20)  # Patent No
        worksheet.set_column('B:B', 30)  # Title
        worksheet.set_column('C:C', 20)  # PDF link
        worksheet.set_column('D:D', 20)  # Inventors
        worksheet.set_column('E:E', 30)  # Assignees
        worksheet.set_column('F:F', 20)  # Publication Date
        worksheet.set_column('G:G', 70)  # Abstract
        worksheet.set_column('H:H', 20)  # Description Link
        worksheet.set_column('I:I', 20)  # Claims
        worksheet.set_column('J:J', 20)  # External Links

        # Format the abstract column to auto-fit row height
        for idx, row in enumerate(data, start=1):
            abstract = row["abstract"]
            if abstract:
                worksheet.write_string(idx, 6, abstract)
                worksheet.set_row(idx, None, None, {'text_wrap': True})