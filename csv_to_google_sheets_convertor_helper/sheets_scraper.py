import csv
import sys
import gspread
from gspread import Client
from gspread.exceptions import APIError
from oauth2client.service_account import ServiceAccountCredentials


ACCESS_SCOPE: [str] = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file',
                       'https://www.googleapis.com/auth/drive']  # Telling the bot what to access.

JSON_KEY_FILE_NAME: str = input('\nPlease type the name of your google cloud credentials json file (with file '
                                'extension) & Press Enter to continue: ')

AUTHENTICATION: ServiceAccountCredentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE_NAME,
                                                                                             ACCESS_SCOPE)
CLIENT_ID: Client = gspread.authorize(AUTHENTICATION)  # Authorising our credentials.

ACCESS_EMAIL: str = input('Type your gmail address here to access the file: ')  # Giving the user's email, access to
# the new spreadsheet.


def reading_from_a_csv(file_name: str):

    with open(file_name, 'r', encoding='utf-8') as my_file:
        csv_reader: csv.DictReader = csv.DictReader(my_file)

        columns: [str] = csv_reader.fieldnames

        rows: [list] = []

        [rows.append(list(row.values())) for row in csv_reader]  # Get the values and add them to a list.

        return columns, rows


def writing_to_google_sheets(file_name: str):

    try:
        print('\nReading our spreadsheet file...')
        file_contents = reading_from_a_csv(file_name)
        file_columns = file_contents[0]
        file_rows = file_contents[1]

        print('\nCreating our Google-Sheets Workbook...')
        viewership_comments_workbook = CLIENT_ID.create(file_name)
        viewership_comments_workbook.share(ACCESS_EMAIL, perm_type='user', role='writer')

        print('\nWriting our Column Headers...')
        viewership_comments_workbook_sheet1 = CLIENT_ID.open(file_name).sheet1
        viewership_comments_workbook_sheet1.insert_row(file_columns)

        try:
            print('\nWriting our data to our new Google sheet...')
            [viewership_comments_workbook_sheet1.insert_row(file_row, idx+2) for idx, file_row in enumerate(file_rows)]
            print('\nWriting Completed Successfully!!')
        except APIError as api_err:
            print(api_err)

    except FileNotFoundError and UnicodeDecodeError as uncderr:
        if FileNotFoundError:
            print('\n\t*** File is not readable. Please make sure you provide a correct and valid spreadsheet file name'
                  'including the file extension ***')
        elif UnicodeDecodeError:
            print(str(uncderr))
        input('\nPress Enter to Exit & Try Again.')
        sys.exit(0)
