import pandas as pd
from io import StringIO

import re

import pydrive2
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive, GoogleDriveFile

from googleapiclient.discovery import build
from tempfile import NamedTemporaryFile


def authenticate_google(conf_path: str="./conf"):
    GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = f"{conf_path}/client_secret.json"
    
    gauth = GoogleAuth()
    gauth.LoadCredentialsFile(f"{conf_path}/credentials.json")
    
    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()
    
    gauth.SaveCredentialsFile(f"{conf_path}/credentials.json")

    return gauth

def download_responses(gauth: pydrive2.auth, folder_id: str) -> dict:
    drive = GoogleDrive(gauth)
    
    def list_files_recursive(folder_id: str):
        files = drive.ListFile({
            'q': f"'{folder_id}' in parents and trashed=false"
        }).GetList()
        
        result = {}
        for file in files:
            if file['mimeType'] == 'application/vnd.google-apps.folder':
                result[file['title']] = list_files_recursive(file['id'])
            else:
                result[file['title']] = file
        
        return result
    
    return list_files_recursive(folder_id)

def clean_filename(filename: str) -> str:
    # Extract number at start
    match = re.search(r'^\d+', filename)
    if not match:
        return ''  # Mark for removal
    
    number = int(match.group())
    
    # Remove extension
    name = filename.rsplit('.', 1)[0]
    # Remove number prefix (with any separators including underscores)
    name = re.sub(r'^\d+\s*[-_\s]*', '', name)
    # Remove "O'Clock" prefix if present
    name = re.sub(r"^\d+\s+O'Clock\s+", '', name, flags=re.IGNORECASE)
    # Replace underscores with spaces BEFORE other processing
    name = name.replace('_', ' ')
    # Remove instrument suffixes
    name = re.sub(r'\s+(alto\d*|flute|drums|bass|soprano|edit|orig|concert|NEW).*$', '', name, flags=re.IGNORECASE)
    # Remove anything in parentheses or after dashes
    name = re.sub(r'\s*[\(\-].*$', '', name)
    # Remove apostrophes, commas, periods, and numbers
    name = re.sub(r"[''\u2018\u2019,.\d]", '', name)
    # Add spaces before capital letters (for SnakeCase/camelCase)
    name = re.sub(r'([a-z])([A-Z])', r'\1 \2', name)
    # Collapse multiple spaces into one
    name = re.sub(r'\s+', ' ', name)
    # Remove trailing whitespace and dashes
    name = name.strip().rstrip('-').strip()
    # Lowercase everything
    name = name.lower()
    
    # Mark for removal if empty
    if not name:
        return ''
    
    return f"{number} {name}"
def create_file_matrix(folder_dict: dict) -> pd.DataFrame:
    # Clean all filenames first
    cleaned_dict = {}
    for folder, files in folder_dict.items():
        if isinstance(files, GoogleDriveFile):
            continue
        elif isinstance(files, dict):
            cleaned_dict[folder] = {clean_filename(f): v for f, v in files.items()}
    
    all_filenames = set()
    for files in cleaned_dict.values():
        all_filenames.update(files.keys())
    
    # Remove empty strings
    all_filenames.discard('')
    
    matrix = pd.DataFrame(index=sorted(all_filenames), 
                         columns=sorted(cleaned_dict.keys()))
    
    for folder, files in cleaned_dict.items():
        matrix[folder] = matrix.index.isin(files.keys())
    
    matrix = matrix.replace({True: "yes", False: "no"})
    
    
    # Add number column
    matrix.insert(0, 'Number', matrix.index.str.split(' ').str[0].astype(int))
    
    return matrix

def filter_matrix(df) -> pd.DataFrame:
    return (
        df
        .drop(columns=['VJE MISCELLANEOUS PARTS (VOCAL, VIBRAPHONE, ETC.)'])
    )

def upload_to_google_sheets(matrix: pd.DataFrame, gauth: pydrive2.auth, folder_id: str, filename: str, file_id: str = None):
    
    drive = GoogleDrive(gauth)
    
    # Write to temporary Excel file
    with NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        matrix.to_excel(tmp.name, index=True)
        tmp_path = tmp.name
    
    # Create or update Google Sheet
    if file_id:
        file = drive.CreateFile({'id': file_id})
    else:
        file = drive.CreateFile({
            'title': filename,
            'parents': [{'id': folder_id}]
        })
    
    file.SetContentFile(tmp_path)
    file.Upload({'convert': True})
    
    return file

def upload_with_formatting(matrix: pd.DataFrame, gauth: pydrive2.auth, folder_id: str, filename: str):
    # First upload with PyDrive2
    file = upload_to_google_sheets(matrix, gauth, folder_id, filename)
    
    # Then add formatting with Sheets API
    service = build('sheets', 'v4', credentials=gauth.credentials)
    
    # Get the actual sheet ID
    spreadsheet = service.spreadsheets().get(spreadsheetId=file['id']).execute()
    sheet_id = spreadsheet['sheets'][0]['properties']['sheetId']
    
    requests = [
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{'sheetId': sheet_id, 'startRowIndex': 1, 'startColumnIndex': 2}],
                    'booleanRule': {
                        'condition': {'type': 'TEXT_CONTAINS', 'values': [{'userEnteredValue': 'yes'}]},
                        'format': {'backgroundColor': {'red': 0.7, 'green': 1, 'blue': 0.7}}
                    }
                }
            }
        },
        {
            'autoResizeDimensions': {
                'dimensions': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': len(matrix.columns) + 1
                }
            }
        }
    ]
    
    service.spreadsheets().batchUpdate(
        spreadsheetId=file['id'],
        body={'requests': requests}
    ).execute()


if __name__ == "__main__":

    gauth = authenticate_google()

    FOLDER_ID = "1EVQ39t7olGOLrzAIPnLQNvgXzpcUwYDx"
    file_list = download_responses(gauth, folder_id=FOLDER_ID)

    matrix = create_file_matrix(file_list).pipe(filter_matrix)
    upload_with_formatting(matrix, gauth, FOLDER_ID, "file_availability_matrix")



