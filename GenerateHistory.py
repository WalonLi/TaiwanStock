


from twstock import Stock
import twstock
import time
import datetime
import os
import re
#import xlsxwriter
import openpyxl
import sys

# type,code,name,ISIN,start,market,group,CFI
# 股票,1101,台泥,TW0001101004,1962/02/09,上市,水泥工業,ESVUFR
#
# DATATUPLE = namedtuple('Data', ['date', 'capacity', 'turnover', 'open',
#                                 'high', 'low', 'close', 'change', 'transaction'])
class Global:
    wait_time = 6
    skip = True # some stock information can't get.

def handle_data(stock, y, m, raw_path, sheet, row):
    Global.skip = True

    for loop in range(3):
        try:
            time.sleep(Global.wait_time) # small delay for website block
            Global.wait_time -= 1
            if Global.wait_time < 2:
                Global.wait_time = 6

            date_list = stock.fetch(y, m)
            # if not stock.low:
            #     raise ValueError

            sheet.append([str(y) + " %02d"%m,
                         min(stock.low),
                         max(stock.high),
                         date_list[0].open,
                         date_list[-1].close,
                         sum(stock.capacity),
                         round(date_list[-1].close - date_list[0].open, 2),
                         sum(stock.turnover)])
            # flush raw data
            file = open(raw_path, 'w')
            for i in range(len(date_list)):
                file.write(str(date_list[i])+'\n')
            file.close()
            Global.skip = False
            break
        except ValueError as e:
            print('unknown value error.', e)


history_head = ['日期', '最低', '最高', '開盤', '收盤', '成交量', '價差', '金額']

def get_history():
    # for number in range(3031, 4000, 1):
    for number in range(3031, 4000, 1):
        
        if str(number) in twstock.codes:

            # time.sleep(2) # small delay for website block

            stock_info = twstock.codes[str(number)]
            stock_name = stock_info[1] + re.sub('\*', '', stock_info[2]) # strip some special character
            if '-DR' in stock_name:
                continue

            path = 'StockList\\' + stock_info[6]
            if not os.path.isdir(path):
                os.mkdir(path)

            path += '\\' + stock_name
            if not os.path.isdir(path):
                os.mkdir(path)

            excel_path = path + '\\' + stock_name + '_History.xlsx'
            if not os.path.isfile(excel_path):
                # if not exist, get all data
                file = openpyxl.Workbook()
                sheet = file.active
                sheet.title = 'Sheet1'
                sheet.append(history_head)
                
                start_date = stock_info[4].split('/')
                start_date = datetime.date(int(start_date[0]), int(start_date[1]), int(start_date[2]))

                target_date = datetime.date(2002, 11, 1)
                if start_date > target_date:
                    target_date = start_date

                year = datetime.date.today().year
                month = datetime.date.today().month

                stock = Stock(str(number))
                print(stock_name, start_date)
                
                
                i = 0
                while True:
                    # handle data
                    print('    %d %d parsing...' % (year, month))
                    handle_data(stock, year, month, path+'\\raw%d%02d'%(year,month), sheet, i+2)
                    if Global.skip:
                        error_file = open(path+'\\raw%d%02d.fail'%(year,month), 'w')
                        error_file.close()
                        break

                    if target_date.year == year and target_date.month == month:
                        print(target_date.year, target_date.month, ' end')
                        break
                    i += 1
                    month -= 1
                    if month == 0:
                        year -= 1
                        month = 12
                file.save(excel_path)
                file.close()
            else:
                # if exist, just scan latest data.
                pass
            
            
def fix_history():
    for root, dirs, files in os.walk(os.getcwd() + '\\StockList'):
        for f in files:
            if f[-5:] == '.fail':
                file = openpyxl.load_workbook(root + '\\' + files[0])
                sheet = file.get_sheet_by_name('Sheet1')
                
                print(root, files[0], f)
                number = int(files[0][:4])
                fail_year = int(f[3:7])
                fail_month = int(f[7:9])
                print('Stock:%d Y:%d M:%02d' % (number, fail_year, fail_month))
                start_date = twstock.codes[str(number)][4].split('/')
                start_date = datetime.date(int(start_date[0]), int(start_date[1]), int(start_date[2]))

                target_date = datetime.date(2002, 11, 1)
                if start_date > target_date:
                    target_date = start_date

                year = fail_year
                month = fail_month

                stock = Stock(str(number))
                
                i = 0
                while True:
                    # handle data
                    print('    %d %d parsing...' % (year, month))
                    handle_data(stock, year, month, root+'\\raw%d%02d'%(year,month), sheet, i+2)
                    if Global.skip:
                        error_file = open(root +'\\raw%d%02d.fail'%(year,month), 'w')
                        error_file.close()
                        break
                    else:
                        if os.path.isfile(root + '\\' + f):
                            os.remove(root + '\\' + f)

                    if target_date.year == year and target_date.month == month:
                        print(target_date.year, target_date.month, ' end')
                        break
                    i += 1
                    month -= 1
                    if month == 0:
                        year -= 1
                        month = 12
                file.save(root + '\\' + files[0])
                file.close()
    
    
if __name__ == '__main__':
    #
    # time.sleep(1)
    # stock = twstock.realtime.get('2412')
    print('Choose functions.')
    print('11:Generate History.')
    print('22:Fix unresolved History.')
    option = input()
    # option = '22'
    if option == '11':
        get_history()
    elif option == '22':
        fix_history()

