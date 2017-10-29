


from twstock import Stock
import twstock
import time
import datetime
import os
import re
import xlsxwriter
import sys

# type,code,name,ISIN,start,market,group,CFI
# 股票,1101,台泥,TW0001101004,1962/02/09,上市,水泥工業,ESVUFR

class Global:
    wait_time = 6


def handle_data(stock, y, m, sheet, row):

    time.sleep(Global.wait_time) # small delay for website block
    Global.wait_time -= 1
    if Global.wait_time < 2:
        Global.wait_time = 6

    date_list = stock.fetch(y, m)
    # if not stock.low:
    #     raise ValueError

    sheet.write_row('A' + str(row),
                    [str(y) + " %02d"%m,
                     min(stock.low),
                     max(stock.high),
                     date_list[0].open,
                     date_list[-1].close,
                     sum(stock.capacity),
                     round(date_list[-1].close - date_list[0].open, 2),
                     sum(stock.turnover)])


history_head = ['日期', '最低', '最高', '開盤', '收盤', '成交量', '價差', '金額']

if __name__ == '__main__':
    #
    # time.sleep(1)
    # stock = twstock.realtime.get('2412')

    for number in range(1000, 2000, 1):
        if str(number) in twstock.codes:

            time.sleep(2) # small delay for website block

            stock_info = twstock.codes[str(number)]
            stock_name = stock_info[1] + re.sub('\*', '', stock_info[2]) # strip some special character
            if '-DR' in stock_name:
                continue

            path = 'StockList//' + stock_info[6]
            if not os.path.isdir(path):
                os.mkdir(path)

            path += '//' + stock_name
            if not os.path.isdir(path):
                os.mkdir(path)

            path += '//' + stock_name + '_History.xlsx'
            if not os.path.isfile(path):
                # if not exist, get all data

                file = xlsxwriter.Workbook(path)
                sheet = file.add_worksheet()
                sheet.write_row('A1', history_head)

                start_date = stock_info[4].split('/')
                start_date = datetime.date(int(start_date[0]), int(start_date[1]), int(start_date[2]))

                current_year = datetime.date(2017, 1, 1)
                result = current_year - start_date
                year = datetime.date.today().year
                month = datetime.date.today().month

                stock = Stock(str(number))


                print(stock_name, start_date)
                if result.days > 16:
                    # only handle recently 15 years
                    for i in range(15*12):
                        # handle data
                        # year = 2017
                        # month = 4
                        print('    %d %d parsing...' % (year, month))
                        handle_data(stock, year, month, sheet, i+2)
                        # sys.exit()
                        month -= 1
                        if month == 0:
                            year -= 1
                            month = 12

                else:
                    i = 0
                    while True:
                        # handle data
                        print('    %d %d parsing...' % (year, month))
                        handle_data(stock, year, month, sheet, i+2)

                        if start_date.year == year and start_date.month == month:
                            print(start_date.year, start_date.month, ' end')
                            break
                        i += 1
                        month -= 1
                        if month == 0:
                            year -= 1
                            month = 12
                            print('    %d parsing...' % year)

                file.close()
            else:
                # if exist, just scan latest data.
                pass

            #
            # print(int(result.days/365))

            # stock = Stock('2421')
            # price = stock.fetch(2017, 10)
            # print(stock)

            # stock = Stock('2421')
            # price = stock.fetch(2017, 9)
            # print(stock.high)
            # break
            # print(twstock.codes[str(i)][4])

            # print(stock.price)


    # print(stock.price)
    # print(twstock.codes['2412'])
