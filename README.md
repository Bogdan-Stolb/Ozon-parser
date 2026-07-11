```markdown
# Ozon Parser

Парсер товаров Ozon с фильтрацией по цене и экспортом в Excel.

## Требования

- Python 3.8+
- Google Chrome (версия 149)

## Установка

1. Клонируйте репозиторий:
```bash
git clone https://github.com/Bogdan-Stolb/Ozon-parser.git
cd Ozon-parser
```

2. Установите зависимости:
```bash
pip install openpyxl undetected-chromedriver selenium
```

## Проверка версии Chrome

**Windows:**
```cmd
reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version
```

**Linux:**
```bash
google-chrome --version
```

**macOS:**
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version
```

Если версия отличается от 149, откройте файл скрипта и найдите строку:
```python
self.driver = uc.Chrome(options=options, version_main=149)
```
Замените `149` на вашу версию Chrome.

## Запуск

```bash
python ozon_parser.py
```

Введите параметры:
1. Запрос
2. Мин. цена
3. Макс. цена
4. Лимит (или Enter для сбора всех товаров)

## Результат

Excel-файл `ozon_[запрос]_[мин]-[макс].xlsx` с колонками:
- Название
- Цена в выдаче
- Цена на странице
- URL
- Характеристики

## Обновление

```bash
git pull origin main
```
```
