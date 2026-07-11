установка:
1.

'''
git clone https://github.com/Bogdan-Stolb/Ozon-parser.git

'''
3.

'''
pip install openpyxl undetected-chromedriver selenium
'''

требования: 
1. python 3.8+
2. Chrome 149

как проверить версию chrome:
'''
# Windows
reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version

# Linux
google-chrome --version

# macOS
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version
'''

если версия отличает, то найдите строку: self.driver = uc.Chrome(options=options, version_main=149)  # укажи свою версию в version_main

запуск:
'''
python script.py
'''

далее введите параметры: 
1. запрос
2. мин.цена
3. макс.цена
4. лимит

