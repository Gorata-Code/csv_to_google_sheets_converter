import sys
from ssl import SSLEOFError, SSLError
from urllib3.exceptions import MaxRetryError
from csv_to_google_sheets_convertor_helper.sheets_scraper import writing_to_google_sheets


def script_summary() -> None:
    print('''
               ***----------------------------------------------------------------------------------------***
         \t***------------------------ DUMELANG means GREETINGS! ~ G-CODE -----------------------***
                     \t***------------------------------------------------------------------------***\n
              
        \t"CSV-TO-GOOGLE~SHEETS-CONVERTER" Version 1.0.0\n
        
        This bot will help you convert a csv file (excel / spreadsheet type of file)
        to a Google Sheet file hosted in your Google Cloud Account. All you need to do
        is provide a filename (for the file you want to read / convert), your gmail address
        and your Google Service Account Credentials (just provide the json file name) and
        then you are all set. 
        
        Cheers!!
    ''')


def csv_gsheets_bot(file_name: str) -> None:
    try:
        writing_to_google_sheets(file_name)

    except FileNotFoundError and SSLEOFError and MaxRetryError and SSLError as conn_err:
        if FileNotFoundError:
            print(
                '\n\t*** This file is not accessible. Please make sure you provide a valid spreadsheet file name & '
                'file extension ***')
        elif SSLEOFError or SSLError or MaxRetryError or 'Max retries exceeded with url' in str(conn_err):
            print('\nThere is a problem with your internet connection. Please make sure you are connected to the '
                  'internet before you try again.')

    input('\nPress Enter to Exit.')
    sys.exit(0)


def main() -> None:
    script_summary()
    file_name: str = input('\nPlease type file name (with extension) and Press Enter: ')

    if len(file_name.strip()) >= 5:
        csv_gsheets_bot(file_name)

    elif len(file_name.strip()) < 5:
        print('\nPlease provide a valid file name.')
        input('\nPress Enter to Exit: ')
        sys.exit(1)


if __name__ == '__main__':
    main()
