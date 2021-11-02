import os
import atexit
import msal
import routes
import logic
from config import TENANT_ID, CLIENT_ID, SHAREPOINT_HOSTNAME, SITE_NAME, ROOT_FOLDER_NAME

AUTHORITY = 'https://login.microsoftonline.com/' + TENANT_ID
ENDPOINT = 'https://graph.microsoft.com/v1.0'

SCOPES = [
    'Files.ReadWrite.All',
    'Sites.ReadWrite.All',
    'User.Read',
    'User.ReadBasic.All'
    ]

def main():
    cache = msal.SerializableTokenCache()

    if os.path.exists('token_cache.bin'):
        cache.deserialize(open('token_cache.bin', 'r').read())

    atexit.register(lambda: open('token_cache.bin', 'w').write(cache.serialize()) if cache.has_state_changed else None)

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    accounts = app.get_accounts()
    res = None
    if len(accounts) > 0:
        res = app.acquire_token_silent(SCOPES, account=accounts[0])

    if res is None:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if 'user_code' not in flow:
            raise Exception('Failed to create device flow')

        print(flow['message'])

        res = app.acquire_token_by_device_flow(flow)

    if 'access_token' in res:
        headers={'Authorization': 'Bearer ' + res['access_token']}

        # hard-coded ranges incase of "ResponsePayloadSizeLimitExceeded error"
        worksheet_columns = [{"GSS Combined RF.xlsx": ['A', 'J']}]
        worksheet_rows = [{"GSS Combined RF.xlsx": [[1, 10000], [10001, 20000],[20001, 30099]]}]

        # get user info
        routes.getUser(ENDPOINT, headers)

        # get the site id
        SITE_ID = routes.getSiteId(ENDPOINT, SHAREPOINT_HOSTNAME, SITE_NAME, headers)

        # get the drive id
        DRIVE_ID = routes.getDriveId(ENDPOINT, SITE_ID, headers)
        # get the root folder id
        ROOT_FOLDER_ID = routes.getRootFolderId(ENDPOINT, ROOT_FOLDER_NAME, DRIVE_ID, headers)
        
        # get folder id
        folder_name = 'Career Center'
        folder_id = logic.getFolderId(ENDPOINT, DRIVE_ID, ROOT_FOLDER_ID, folder_name, headers)
        
        # get subfolder contents
        subfolder_name = 'Gali - Internship ROI'
        subfolder_id = logic.getFolderId(ENDPOINT, DRIVE_ID, folder_id, subfolder_name, headers)
        subfolder_contents = logic.getFiles(ENDPOINT, DRIVE_ID, subfolder_id, headers)

    else:
        raise Exception('no access token in res')


if __name__ == "__main__":
    main()