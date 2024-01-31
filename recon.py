import requests
import PyPDF2
import json
import io
import re
import openpyxl

# This function will return a list of all the id's for the pdfs that we want to search through
def accessID(startYear, endYear, chamberArray):
    idList = []
    for year in range(startYear, endYear + 1):
        response = requests.get("https://api.parliament.nsw.gov.au/api/hansard/search/year/" + str(year))
        if response.status_code == 200:
            response = json.loads(response.text)

        # for each item in the response, get its events, and in the events array, loop through and obtain the id for each event
        for dateItem in response:
            for eventItems in dateItem['Events']:
                chamberArray.append(eventItems['Chamber'])
                idList.append(eventItems['PdfDocId'])
        pass
    return idList

# This function will return the pdf content for a given id
def getPDF(id):
    print("Accessing id", id, "...")
    response = requests.get("https://api.parliament.nsw.gov.au/api/hansard/search/daily/pdf/" + id)
    if response.status_code == 200:
        print("Success!!")
        return response.content
    else:
        print("failure")

# This function will search through the pdf content for a given search term
def search_pdf(pdf_content, search_term):
    pdf_file = io.BytesIO(pdf_content)
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    
    total_occurrences = 0

    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()
        
        # Remove line breaks and other whitespace characters
        cleaned_text = re.sub(r'\s+', ' ', text)

        occurrences_on_page = cleaned_text.lower().count(search_term.lower())
        total_occurrences += occurrences_on_page
        
    print(f"Total occurrences of '{search_term}': {total_occurrences}")
    return total_occurrences

# This function will create an excel file with the given headings
def createExcelFile(filePath, headings):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for col_num, heading in enumerate(headings, 1):
        sheet.cell(row=1, column=col_num, value=heading)
    workbook.save(filePath)

# This function will add data to an existing excel file
def add_data_to_excel(file_path, data):
    # Open the existing Excel workbook
    workbook = openpyxl.load_workbook(file_path)

    # Select the default sheet (index 0)
    sheet = workbook.active

    # Find the next empty row
    next_row = sheet.max_row + 1

    # Add data to the sheet
    for col_num, value in enumerate(data, 1):
        sheet.cell(row=next_row, column=col_num, value=value)

    # Save the changes to the workbook
    workbook.save(file_path)

# ******************* This is the main function *******************************
# TODO: allow user input for the years and search term. Create default values for these

#Could have an array of chamerID
chambers = []
pdf_ID = accessID(2022, 2023, chambers)
headings = [
    "PDF ID", 
    "Chamber", 
    "Reconciliation Australia” mentioned?", 
    "Reconciliation NSW” mentioned?", 
    "Reconciliation” mentioned?",
    "# of times “Reconciliation Australia” mentioned?",
    "# of times “Reconciliation NSW” mentioned?",
    "# of times “Reconciliation” mentioned?"
]
file_path = 'output.xlsx'
createExcelFile(file_path, headings)

for PDFid in pdf_ID:
    response = getPDF(PDFid)
    dataToAdd = []

    dataToAdd.append(PDFid)

    # need to change this because not all of them are Upper House
    dataToAdd.append(chambers[pdf_ID.index(PDFid)])


    #  do a loop for the headings
    #TODO: create a system that allows for customised search terms
    searchResults = []
    searchResults.append(search_pdf(response, "Reconciliation Australia"))
    searchResults.append(search_pdf(response, "Reconciliation NSW"))
    searchResults.append(search_pdf(response, "Reconciliation"))

    for results in searchResults:
        if results > 0:
            dataToAdd.append("Yes")
        else:
            dataToAdd.append("No")

    dataToAdd.extend(searchResults)
    add_data_to_excel(file_path, dataToAdd)
    print("\n\n")