import requests
import urllib

# ********************************** General Access to Drive, Folders and Files **********************************

#TODO: explanations of routes
def getUser(ENDPOINT: str, headers: dict):
    res = requests.get(f'{ENDPOINT}/me', headers=headers)
    res.raise_for_status()
    print(res.json())

def getSiteId(ENDPOINT, SHAREPOINT_HOSTNAME, SITE_NAME, headers):
    res = requests.get(f'{ENDPOINT}/sites/{SHAREPOINT_HOSTNAME}:/sites/{SITE_NAME}', headers=headers)
    site_info = res.json()
    return site_info['id']

def getDriveId(ENDPOINT, site_id, headers):
    res = requests.get(f'{ENDPOINT}/sites/{site_id}/drive', headers=headers)
    res.raise_for_status()
    drive_info = res.json()
    return drive_info['id']  

def getRootFolderId(ENDPOINT, folder_name, drive_id, headers):
    folder_url = urllib.parse.quote(folder_name)
    res = requests.get(f'{ENDPOINT}/drives/{drive_id}/root:/{folder_url}', headers=headers)
    res.raise_for_status()
    folder_info = res.json()
    return folder_info['id']

def getRootFolderInfo(ENDPOINT, folder_name, drive_id, headers):
    folder_url = urllib.parse.quote(folder_name)
    res = requests.get(f'{ENDPOINT}/drives/{drive_id}/root:/{folder_url}', headers=headers)
    res.raise_for_status()
    return res.json()

def getFilesById(ENDPOINT, drive_id, folder_id, headers):
    url = f'{ENDPOINT}/drives/{drive_id}/items/{folder_id}/children'
    res = requests.get(f'{ENDPOINT}/drives/{drive_id}/items/{folder_id}/children', headers=headers)
    res.raise_for_status()
    return res.json()['value']

# *************** Workbooks ***************

def getExcelWorkbook(ENDPOINT, drive_id, file_id, headers):
    res = requests.get(f'{ENDPOINT}/drives/{drive_id}/items/{file_id}/workbook/worksheets', headers=headers)
    res.raise_for_status()
    return res.json()['value']

# get all cells in a worksheet
def getAllWorksheetCells(ENDPOINT, drive_id, file_id, sheet_id, headers):
    res = requests.get(f'{ENDPOINT}/drives/{drive_id}/items/{file_id}/workbook/worksheets/{sheet_id}/usedRange(valuesOnly=true)?$select=values', headers=headers)
    res.raise_for_status()
    return res.json()['values']

# get a specific range in a worksheet
def getWorksheetRange(ENDPOINT, drive_id, file_id, sheet_id, address, headers):
    res = requests.get(f"{ENDPOINT}/drives/{drive_id}/items/{file_id}/workbook/worksheets/{sheet_id}/range(address='{address}')", headers=headers)
    res.raise_for_status()
    return res.json()['values']

    