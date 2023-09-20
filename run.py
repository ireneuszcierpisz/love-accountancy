import gspread
from google.oauth2.service_account import Credentials
from pprint import pprint

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)

GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('love-accountancy')

tb = SHEET.worksheet("Trial Balance").get_all_values()


def get_gl_codes():
    """
    Get GL Codes or Titles as an input from the user.
    Run a while loop to collect a valid string of data from the user
    via the terminal. The data must be a string of elements separated
    by commas. The loop will repeatedly request data, until it is valid.
    """
    statements = ['SOPL', 'SOFP']
    data = {}
    i = 0
    while True:

        print(f"Please enter first and last GL code for {statements[i]}.")
        print("It must be exactly same strings of signs as the GL Code in TB.")
        print('Separate two numbers by comma, without spaces.')

        user_data = input(f"Input two GL Codes for {statements[i]} here:\n")
        user_data = user_data.split(',')

        if validate_data(user_data):
            print(f'\nGot unique GL Codes for {statements[i]} : {user_data}\n')
            # print("If it is wrong, don't worry just run program again.")
            data[statements[i]] = user_data
            if i == 1:
                break
            else:
                i += 1

    # print('Going to proceed with:')
    # for k,v in data.items():
    #     print(f'{k}: {v}')

    return data


def validate_data(values):
    """
    Inside the try, converts all string values into integers.
    Raises ValueError if strings cannot be converted into int,
    or if there aren't exactly 2 values.
    """
    try:
        for value in values:
            n = 0
            for row in tb:
                n += row.count(value)
            # print(f"! Found {n} GL Code '{value}' !")
            if n != 1:
                raise ValueError(f'"{value}" is in {n} celles, check GL Codes')

        if len(values) != 2:
            raise ValueError(f'Two values required, provided {len(values)}')
    except ValueError as e:
        print(f'Invalid data: {e}, please try again.\n')
        return False

    return True


def check_balance(data, type_of_fs):
    """
    Checking if balance in given column of data is equel zero
    """
    financial_periods = ['current period', 'previous period']
    col = 2  # financial results for current period
    for i in range(len(financial_periods)):
        print(f'Checking balance on {type_of_fs} data for {financial_periods[i]}...\n')
        column = [row[col].replace('(', '-').replace(')', '') for row in data]
    # print('Got column to calculate balance: ', column)
        float_column = [float(num.replace(',', '')) for num in column]
    # print('float column: ', float_column)
        balance = sum(float_column)
        print(f'Balance for {financial_periods[i]} on {type_of_fs} data is: {balance:,.2f}\n')
        col += 2


def get_data_for_fs(type_of_fs):
    # gets trial balance as a list of lists of strings
    # gets data entered by user as a dictionary

    print(f'Extracting data for {type_of_fs}...')
    user_data = get_gl_codes()
    first_row, last_row = 0, 0
    for row in tb:
        if row[0] == user_data[type_of_fs][0]:
            first_row = tb.index(row) + 1
            # print('first row num: ', first_row)
        if row[0] == user_data[type_of_fs][1]:
            last_row = tb.index(row) + 1
            # print('last row num: ', last_row)
    table_for_fs = tb[first_row - 1:last_row]
    # return needed part of TB worksheet as a list of lists
    return table_for_fs, type_of_fs


def main():
    sofp_in_tb, type_of_fs = get_data_for_fs('SOFP')
    check_balance(sofp_in_tb, type_of_fs)
    sopl_data, fs_type = get_data_for_fs('SOPL')
    check_balance(sopl_data, fs_type)



print('Welcome! This tool is only for accountants :)')
print('You can use it for preparing FS from your Trial Balance.\n')

# sofp = SHEET.worksheet("SOFP").get_all_values()
# # sofp_data = sofp.get_all_values()
# pprint(sofp)
# print(sofp[2][4], sofp[4][0])
# cell_list = SHEET.worksheet('Trial Balance').findall("Other Liabilities")
# print(cell_list)

main()
