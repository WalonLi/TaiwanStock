
pip install twstock
pip install xlutils
pip install openpyxl
copy Patch\stock.py C:\Python36-32\Lib\site-packages\twstock\stock.py
copy Patch\worksheet.py C:\Python36-32\Lib\site-packages\openpyxl\worksheet\worksheet.py
copy Patch\tpex_equities.csv C:\Python36-32\Lib\site-packages\twstock\codes\tpex_equities.csv
copy Patch\twse_equities.csv C:\Python36-32\Lib\site-packages\twstock\codes\twse_equities.csv
