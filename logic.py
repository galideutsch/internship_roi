from logging import raiseExceptions
import routes 
import requests

def getFolderId(ENDPOINT, drive_id, parent_folder_id, folder_name, headers):
    folder_contents = routes.getFilesById(ENDPOINT, drive_id, parent_folder_id, headers)
    for item in folder_contents:
        if item['name'] == folder_name:
            return item['id']

def getFiles(ENDPOINT, drive_id, folder_id, headers):
    folder_contents = routes.getFilesById(ENDPOINT, drive_id, folder_id, headers)
    return folder_contents

def getAllWorkbooks(ENDPOINT, drive_id, folder, headers):
    workbooks = []
    for file in folder:
        if file['file']['mimeType'] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            workbook = getWorkbook(ENDPOINT, drive_id, file['id'], headers)
            workbooks.append(workbook)
    return workbooks

def getWorkbook(ENDPOINT, drive_id, file, headers, options=None):
    workbook_content = routes.getExcelWorkbook(ENDPOINT, drive_id, file['id'], headers)

    # get worksheets and their contents
    sheets = getWorksheets(ENDPOINT, drive_id, file['id'], workbook_content, headers, options)
    return sheets

def getWorksheets(ENDPOINT, drive_id, file_id, workbook, headers, options=None):
    worksheets = []
    for sheet in workbook:   
        try:
            sheet_data = routes.getAllWorksheetCells(ENDPOINT, drive_id, file_id, sheet['id'], headers)
            
            ws = sheet_data
            worksheets.append(ws)
        except requests.exceptions.HTTPError as err:
            print(f'Error logic::getWorksheets: {err}')
    return worksheets