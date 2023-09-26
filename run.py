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

# fetches values from all of the cells of the sheet to reduce API calls:
tb = SHEET.worksheet("Trial Balance").get_all_values()
sopl_worksheet = SHEET.worksheet("SOPL")
sofp_worksheet = SHEET.worksheet("SOFP")

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

        user_data = input("Type GL Codes separated by a comma(e.g.: 400-e,3531) here:\n")
        user_data = user_data.split(',')

        if validate_data(user_data):
            print(f'Got unique GL Codes for {statements[i]} : {user_data}\n')
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
    financial_periods = ['Current Period', 'Previous Period']
    col = 2  # financial results for current period
    # dictionary to collect all data needed to create financial raports as SOFP and SOPL
    collection = {type_of_fs: {}}
    for i in range(len(financial_periods)):
        print(f'{financial_periods[i]}')
        column = [row[col].replace('(', '-').replace(')', '') for row in data]
        float_column = [float(num.replace(',', '')) for num in column]
        balance = sum(float_column)
        print(f'    {type_of_fs} Balance is:  {balance:,.2f}')
        balance = sum([round(n) for n in float_column])
        print(f'    {type_of_fs} Balance on Rounded Numbers is: {balance:,.2f}\n')

        # collecting data for financial statement worksheet
        collection[type_of_fs][financial_periods[i]] = {}
        for title in set([row[6] for row in data]):
            collection[type_of_fs][financial_periods[i]][title] = []
            for row in data:
                if title == row[6]:
                    num = float(row[col].replace('(', '-').replace(')', '').replace(',', ''))
                    collection[type_of_fs][financial_periods[i]][title].append(num)
        for k, v in collection[type_of_fs][financial_periods[i]].items():
            collection[type_of_fs][financial_periods[i]][k] = sum(v)
        # print(collection)
        col += 2

    return collection


def get_data_for_fs(type_of_fs, user_data):
    # gets trial balance as a list of lists of strings
    # gets data entered by user as a dictionary

    print(f'\nExtracting data from TB for {type_of_fs}...')
    first_row, last_row = 0, 0
    for row in tb:
        if row[0] == user_data[type_of_fs][0]:
            first_row = tb.index(row) + 1
        if row[0] == user_data[type_of_fs][1]:
            last_row = tb.index(row) + 1
    table_for_fs = tb[first_row - 1:last_row]
    # retuart of TB worksheet as a list of lists
    # return type of financial statement
    return type_of_fs, table_for_fs


def make_raport(raport_name, data_collection, g_worksheet):
    print(f'Generating raport {raport_name}...')
    data_list = g_worksheet.get_all_values()
    col = 4
    for p in ['Current Period', 'Previous Period']:
        # print(f'  {p}:')
        for k, v in data_collection[raport_name][p].items():
            row_num = 1
            found = False
            for row in data_list:
                if row[0] == k:
                    try:
                        if found:
                            raise ValueError(f'"{k}" repeated in {raport_name}')
                    except ValueError as e:
                        print(f'\n   Check your FS! {e}, row {row_num}.\n')
                        input('Press Enter to continue.')
                    g_worksheet.update_cell(row_num, col, v)
                    # print(f'{k} = {v} , {raport_name} row:{row_num} col:{col}')
                    found = True
                row_num += 1
            try:
                if not found:
                    raise ValueError(f'"{k}" NOT found in {raport_name}!')
            except ValueError as e:
                print(f'\nCheck FS! {e}, value {v:,.2f} not assigned!\n')
                input('Press Enter to continue.')
        col += 2


def handle_data(g_worksheet):
    # get financial statement as a list
    print('Computing totals in FS...')
    fs = g_worksheet.get_all_values()
    for col in [3, 5]:
        for i in range(len(fs)):
            total = 0
            if fs[i][0][:5] == "Total":
                j = 1
                while True:
                    if fs[i - j][col] == '':
                        break
                    total += float(fs[i - j][col].replace(',', '.'))
                    j += 1
                g_worksheet.update_cell(i + 1, col + 1, total)


def main():
    user_data = get_gl_codes()

    sofp, fp_table = get_data_for_fs('SOFP', user_data)
    sofp_collection = check_balance(fp_table, sofp)

    sopl, pl_table = get_data_for_fs('SOPL', user_data)
    sopl_collection = check_balance(pl_table, sopl)

    make_raport('SOFP', sofp_collection, sofp_worksheet)
    handle_data(sofp_worksheet)
    print('Your SOFP raport is ready. Check your Google Spreadsheet, please.')

    make_raport('SOPL', sopl_collection, sopl_worksheet)


print('Welcome! This tool is  useful only for accountants :)')
print('You can use it for preparing FS from your Trial Balance.\n')


main()
