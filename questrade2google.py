import sys
import csv
import xlrd
from pandas.io.data import DataReader
from datetime import datetime


def price_on_day(symbol, date):
    data = DataReader(symbol + '.to',  "yahoo", date, date)
    price = data['Close'][date]
    return price


def write_csv(transactions, filename):
    with open(filename, 'w') as csvfile:
        fieldnames = ['Symbol', 'Type', 'Date', 'Shares', 'Price', 'Commission']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for transaction in transactions:
            writer.writerow(transaction)


def parse_columns(sheet):
    cols = {'Type': ''}
    for x in range(sheet.ncols):
        text = sheet.cell(0, x).value
        if text == 'TransactionDate':
            cols['Date'] = x
        elif text == 'Symbol':
            cols['Symbol'] = x
        elif text == 'Quantity':
            cols['Shares'] = x
        elif text == 'Price':
            cols['Price'] = x
        elif text == 'Commission':
            cols['Commission'] = x
        elif text == 'AccountNumber':
            cols['X-Account'] = x
        elif text == 'ActivityType':
            cols['X-ActivityType'] = x
        elif text == 'Action':
            cols['X-Action'] = x
        elif text == 'CurrencyDisplay':
            cols['X-Currency'] = x
    return cols


def process_xlsx(filename, account):
    sheet = xlrd.open_workbook(filename).sheet_by_index(0)

    cols = parse_columns(sheet)

    transactions = []
    for row in range(1, sheet.nrows):
        transaction = parse_row(sheet, row, cols, account)
        if transaction is not None:
            transactions.append(transaction)

    return transactions


def parse_row(sheet, row, cols, account):
    transaction = {}

    # only process CAD transactions
    xcurrency = sheet.cell(row, cols['X-Currency']).value
    if (xcurrency != 'CAD'):
        return None
    # only process for the account specified
    xaccount = sheet.cell(row, cols['X-Account']).value
    if (xaccount != account):
        return None
    olddate = sheet.cell(row, cols['Date']).value
    day = int(olddate[0:2])
    month = int(olddate[3:5])
    year = int(olddate[6:10])
    date = datetime(year, month, day)
    xaction = sheet.cell(row, cols['X-Action']).value
    xactivity = sheet.cell(row, cols['X-ActivityType']).value

    transaction['Date'] = str(year) + '-' + str(month) + '-' + str(day)

    for col in cols.keys():
        if col[0:2] == 'X-':
            # extra column
            continue
        elif col == 'Symbol':
            if not sheet.cell(row, cols[col]).value:
                # transaction not involving shares
                return None
            else:
                symbol = sheet.cell(row, cols[col]).value
                # chop off '.TO'
                if symbol[-3:] == '.TO':
                    symbol = symbol[0:-3]

                if xactivity == 'Withdrawals':
                    transaction['Price'] = price_on_day(symbol, date)
                    transaction['Type'] = 'Sell'
                elif xactivity == 'Deposits':
                    transaction['Price'] = price_on_day(symbol, date)
                    transaction['Type'] = 'Buy'
                transaction[col] = symbol
        elif col == 'Commission':
            # use positive commissions
            commission = sheet.cell(row, cols[col]).value
            commission = abs(commission)
            transaction[col] = commission
        elif col == 'Type':
            if xaction == 'Buy':
                transaction[col] = 'Buy'
            elif xaction == 'Sell':
                transaction[col] = 'Sell'
            elif xactivity == 'Dividends':
                return None
            else:
                continue
        elif col == 'Date':
            continue
        elif col == 'Price':
            if xactivity == 'Withdrawals' or xactivity == 'Deposits':
                continue
            else:
                transaction[col] = sheet.cell(row, cols[col]).value
        elif col == 'Shares':
            # use positive shares only (ie for Sells)
            shares = sheet.cell(row, cols[col]).value
            shares = abs(shares)
            transaction[col] = shares
        else:
            transaction[col] = sheet.cell(row, cols[col]).value

    return transaction


def main(argv):
    infile = '2014.xlsx'
    outfile = 'output.csv'
    account = '########'
    if len(argv) == 3:
        infile = argv[1]
        outfile = argv[2]
        account = argv[3]
    else:
        sys.stdout.write('Usage: quest2goog.py input.xlsx output.csv accountNumber\n')
        sys.exit(0)

    transactions = process_xlsx(infile, account)
    write_csv(transactions, outfile)
    sys.stdout.write('Done!')
    #os.system('type output.csv')

if __name__ == "__main__":
    main(sys.argv)
