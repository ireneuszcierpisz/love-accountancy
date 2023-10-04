import gspread
from google.oauth2.service_account import Credentials
import time
import sys


SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)

GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('love-accountancy')

# fetches values from all of the cells of the Trial Balance (TB) spreadsheet
# to reduce API calls:
tb = SHEET.worksheet("Trial Balance").get_all_values()
# Statement of Profit and Loss (SOPL) and
# Statement of Financial Position (SOFP)
# worksheets as the global variables
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

        print(f"Enter the first and last GL code \
set in TB for {statements[i]}!")
        print("In this case: 40001,86000 (for SOPL) or 1,35001 (for SOFP)")

        user_data = input("Type two GL Codes separated by a comma here:\n")
        user_data = user_data.split(',')

        if validate_data(user_data):
            print(f'Got unique GL Codes for {statements[i]} : {user_data}\n')
            # print("If it is wrong, don't worry just run program again.")
            data[statements[i]] = user_data
            if i == 1:
                break
            else:
                i += 1

    return data


def validate_data(values):
    """
    Inside the try, converts all string values into integers.
    Raises ValueError if strings cannot be found as the GL code,
    or there are more than one position with the same GL code
    or if there aren't exactly 2 values provided by the user.
    """
    try:
        # check if the GL code is unique
        for value in values:
            n = 0
            for row in tb:
                n += row.count(value)
            if n != 1:
                raise ValueError(f'"{value}" is in {n} cells, check GL Codes')
        # check if there are two values provided
        if len(values) != 2:
            raise ValueError(f'Two values required, provided {len(values)}')
    except ValueError as e:
        print(f'\n   ! Invalid data: {e}, please try again.\n')
        return False

    return True


def check_balance(data, type_of_fs):
    """
    Checking the sum of a given column of numbers in the Trial Balance
    worksheet as related to the relevant financial statements
    and collecting data for the financial statement.
    """
    print('Checking the balance...')
    financial_periods = ['Current Period', 'Previous Period']
    col = 2  # financial results column in TB for the current period
    # dictionary to collect all data from TB needed to create
    # financial raports SOFP and SOPL:
    collection = {type_of_fs: {}}
    for i in range(len(financial_periods)):
        print(f'  {financial_periods[i]}')
        column = [row[col].replace('(', '-').replace(')', '') for row in data]
        float_column = [float(num.replace(',', '')) for num in column]
        balance = sum(float_column)
        print(f'    {type_of_fs} Balance is:  {balance:,.2f}')
        # collecting data for financial statement worksheet
        collection[type_of_fs][financial_periods[i]] = {}
        for title in set([row[6] for row in data]):
            collection[type_of_fs][financial_periods[i]][title] = []
            for row in data:
                if title == row[6]:
                    num = float(row[col].replace(
                        '(', '-').replace(')', '').replace(',', ''))
                    collection[type_of_fs][financial_periods[i]][title].append(
                        num)
        for k, v in collection[type_of_fs][financial_periods[i]].items():
            collection[type_of_fs][financial_periods[i]][k] = sum(v)
        col += 2
    print('-----------------\n')
    return collection


def get_data_for_fs(type_of_fs, user_data):
    """
    Uses trial balance as a list of lists,
    gets as a dictionary data entered by the user
    and return part of tb as a data for one of financial statements
    """

    print(f'Extracting data from TB for {type_of_fs}...')
    first_row, last_row = 0, 0
    for row in tb:
        if row[0] == user_data[type_of_fs][0]:
            first_row = tb.index(row) + 1
        if row[0] == user_data[type_of_fs][1]:
            last_row = tb.index(row) + 1
    try:
        if first_row > last_row:
            raise ValueError(
                'The first GL code in the entered pair \
must be higher in the TB table than the second!')
    except ValueError as e:
        print(f'Error: {e}.')
        print('Invalid GL code pair! Start over.')
        sys.exit(1)
    table_for_fs = tb[first_row - 1:last_row]
    # return part of the TB worksheet as a list of lists
    # returns the name of the financial statement type
    return type_of_fs, table_for_fs


def make_raport(raport_name, data_collection, g_worksheet):
    """
    Prepares financial statement (fs) of a given type.
    Retrieves the data set for the fs and
    updates the worksheet with rounded values and
    changes the sign of the numbers where required.
    Throws an error if any item in fs is not unique
    or there is no item that exists in TB.
    """
    print(f'Generating report {raport_name}...')
    data_list = g_worksheet.get_all_values()
    # gets range of worksheet rows containing assets
    first, last = 0, 0
    for i in range(len(data_list)):
        if data_list[i][0] == 'ASSETS':
            first = i
        if data_list[i][0] == 'EQUITY':
            last = i
    col = 4
    for p in ['Current Period', 'Previous Period']:
        for k, v in data_collection[raport_name][p].items():
            row_num = 1
            found = False
            for row in data_list:
                if row[0] == k:
                    try:
                        if found:
                            raise ValueError(
                                f'"{k}" repeated in {raport_name}')
                    except ValueError as e:
                        print(f'\n   Check your FS! {e}, row {row_num}.\n')
                        input('Press Enter to continue or Run Program again.')
                    if row_num < last and row_num > first:
                        g_worksheet.update_cell(row_num, col, round(v))
                    else:
                        g_worksheet.update_cell(row_num, col, round(-1 * v))
                    found = True
                row_num += 1
            try:
                if not found:
                    raise ValueError(f'"{k}" NOT found in {raport_name}!')
            except ValueError as e:
                print(f'\nCheck FS! {e}, value {v:,.2f} not assigned!\n')
                input('Fix {raport_name} and start over \
                    or press Enter to continue.')
        col += 2


def handle_data(g_worksheet):
    """
    Adds a style to the sheet header.
    Calculates totals in columns for given groups of data,
    calculates the required sum of totals in the statement and
    updates the appropriate spreadsheet cells.
    """
    # adds style to the worksheet heading
    g_worksheet.format("A1:F2", {
        "backgroundColor": {
            "red": 0.5,
            "green": 1.0,
            "blue": 0.5
        },
        "horizontalAlignment": "Left",
        "textFormat": {
            "foregroundColor": {
                "red": 0.50,
                "green": 0.50,
                "blue": 0.50
            },
            "fontSize": 11,
        }
    })
    print('   Computing totals...')
    # gets financial statement as a list
    fs = g_worksheet.get_all_values()
    for col in [3, 5]:
        t_o_t = 0   # total of totals variable
        for i in range(len(fs)):
            time.sleep(1)
            total = 0
            if fs[i][0][:5] == "Total":
                j = 1
                # the flag indicates whether the cell needs to be updated
                # with the t_o_t value
                flag = False
                while True:
                    if fs[i - j][col] == '':
                        if j == 1:
                            flag = True
                        break
                    total += int(fs[i - j][col])
                    j += 1
                t_o_t += total
                if flag:
                    g_worksheet.update_cell(i + 1, col + 1, t_o_t)
                    t_o_t = 0
                else:
                    g_worksheet.update_cell(i + 1, col + 1, total)

    if g_worksheet == sopl_worksheet:
        compute_loss(fs, g_worksheet)


def compute_loss(fs, g_worksheet):
    """
    Computes Loss Before Tax and Loss for the Financial Period
    for fs: Statement of Profit and Loss.
    Updating the appropriate cells in the SOPL report with loss values.
    """
    # Retrieves the appropriate data from fs and performs the calculations
    for i in range(len(fs)):
        if "Operating loss" in fs[i][0]:
            ol_current = int(fs[i][3])
            ol_previous = int(fs[i][5])
            continue
        if "Total finance costs" in fs[i][0]:
            tfc_current = int(fs[i][3])
            tfc_previous = int(fs[i][5])
            # Loss before tax, current period:
            lbtcp = ol_current + tfc_current
            g_worksheet.update_cell(i + 3, 4, lbtcp)
            # Loss before tax, previous period:
            lbtpp = ol_previous + tfc_previous
            g_worksheet.update_cell(i + 3, 6, lbtpp)
            continue
        if "Tax for the financial period" in fs[i][0]:
            tftfp_current = int(fs[i][3])
            tftfp_previous = int(fs[i][5])
            # Loss for the financial period, current period:
            lftfpc = lbtcp + tftfp_current
            g_worksheet.update_cell(i + 3, 4, lftfpc)
            # Loss for the financial period, previous period:
            lftfpp = lbtpp + tftfp_previous
            g_worksheet.update_cell(i + 3, 6, lftfpp)


def main():
    user_data = get_gl_codes()

    sofp, fp_table = get_data_for_fs('SOFP', user_data)
    sofp_collection = check_balance(fp_table, sofp)

    sopl, pl_table = get_data_for_fs('SOPL', user_data)
    sopl_collection = check_balance(pl_table, sopl)

    make_raport('SOFP', sofp_collection, sofp_worksheet)
    handle_data(sofp_worksheet)
    print('\nSOFP report is ready in google spreadsheet.\n')

    make_raport('SOPL', sopl_collection, sopl_worksheet)
    handle_data(sopl_worksheet)
    print('\nSOPL report is ready in google spreadsheet.\n')


print('\n    Welcome! Please follow the instructions below.')
print('You can use that tool for preparing Financial \
Statements from the Trial Balance.\n')


main()
