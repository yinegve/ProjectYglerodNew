import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import string
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import re
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


def extract_digits(s): # Функция для извлечения цифр из строки
    digits = re.findall(r'\d+\.\d+|\d+', str(s))
    return digits

def divide_and_multiply(nums): # Функциия вычисления объема
    if len(nums) == 3:
        nums = [round(float(num) / 1000, 2) for num in nums]
        result = 1
        for num in nums:
            result *= num
        return round(result, 3)
    else:
        return None

dfyg = pd.read_excel("THE_THEATRE.xlsx")
yg = pd.DataFrame(dfyg)

yg = yg.iloc[7:, 1:]
yg.columns = ['Description', 'Units', 'Quantity', 'Unit cost', 'Total cost', 'Unit price', 'Total price', 'Remark for customer']

colums = ['Unit cost', 'Total cost', 'Unit price', 'Total price', 'Remark for customer']
yg.drop(colums, axis=1, inplace=True)
yg.columns = ['Description', 'Units', 'Quantity']
yg = yg[yg['Units'] != 'job']
yg = yg[yg['Units'] != 'Job']
yg = yg[yg['Units'].notna()]

yg['Quantity'] = pd.to_numeric(yg['Quantity'], errors='coerce')

yg.to_excel("THE_THEATRE_new.xlsx", index = False)

# Загрузка данных
df = pd.read_excel('THE_THEATRE_new.xlsx')
units = df['Units']
quantity = df['Quantity']

# Выбор колонки
column_data = df['Description'].astype(str)

# Токенизация
nltk.download('punkt')
tokenized_data = column_data.apply(nltk.word_tokenize)

# Приведение к нижнему регистру
lowercased_data = tokenized_data.apply(lambda tokens: [word.lower() for word in tokens])

    # Удаление знаков препинания
punctuation_table = str.maketrans('', '', string.punctuation)
without_punctuation = lowercased_data.apply(lambda tokens: [word.translate(punctuation_table) for word in tokens])

    # Удаление стоп-слов
nltk.download('stopwords')
english_stopwords = stopwords.words('english')

def remove_stopwords(text_tokens):
    return [word for word in text_tokens if not word in english_stopwords and word]

filtered_data = without_punctuation.apply(remove_stopwords)

    # Лемматизация
nltk.download('wordnet')
nltk.download('omw-1.4')
lemmatizer = WordNetLemmatizer()

def lemmatize_words(text_tokens):
    return ' '.join([lemmatizer.lemmatize(word) for word in text_tokens])

lemmatized_data = filtered_data.apply(lemmatize_words)

    # Создание нового DataFrame для сохранения результатов
new_df = pd.DataFrame({
    'Description': lemmatized_data,
    'Units': units,
    'Quantity': quantity
})

new_df['Size, mm'] = new_df['Description'].apply(extract_digits)

new_df['Volume, m^3'] = new_df['Size, mm'].apply(divide_and_multiply)


for index, row in new_df.iterrows():
     # Проверим, пусто ли значение в столбце "Quantity"
    if pd.isnull(row['Quantity']):
            # Если значение пустое, заменим его на 1
        new_df.at[index, 'Quantity'] = 1

for index, row in new_df.iterrows():
    if "steel" in row['Description'] and not pd.isnull(row['Volume, m^3']):
        new_df.at[index, 'Mass, kg'] = row['Volume, m^3'] * 7000
for index, row in new_df.iterrows():
    if "porcelain" in row['Description'] and not pd.isnull(row['Volume, m^3']):
        new_df.at[index, 'Mass, kg'] = row['Volume, m^3'] * 2800
for index, row in new_df.iterrows():
    if "gypsum" in row['Description'] and not pd.isnull(row['Volume, m^3']):
        new_df.at[index, 'Mass, kg'] = row['Volume, m^3'] * 800
for index, row in new_df.iterrows():
    if "plywood" in row['Description'] and not pd.isnull(row['Volume, m^3']):
        new_df.at[index, 'Mass, kg'] = row['Volume, m^3'] * 700
for index, row in new_df.iterrows():
    if "acrylic" in row['Description'] and not pd.isnull(row['Volume, m^3']):
        new_df.at[index, 'Mass, kg'] = row['Volume, m^3'] * 1200
for index, row in new_df.iterrows():
    if "steel" in row['Description']:
        new_df.at[index, 'Сarbon, kg'] = row['Mass, kg'] * 1.5
for index, row in new_df.iterrows():
    if "porcelain" in row['Description']:
        new_df.at[index, 'Сarbon, kg'] = row['Mass, kg'] * 0.8
for index, row in new_df.iterrows():
    if "gypsum" in row['Description']:
        new_df.at[index, 'Сarbon, kg'] = row['Mass, kg'] * 5.4
for index, row in new_df.iterrows():
    if "plywood" in row['Description']:
        new_df.at[index, 'Сarbon, kg'] = row['Mass, kg'] * 0.9
for index, row in new_df.iterrows():
    if "acrylic" in row['Description']:
        new_df.at[index, 'Сarbon, kg'] = row['Mass, kg'] * 0.9


new_df['Carbon footprint, T'] = (new_df['Сarbon, kg'] * new_df['Quantity'])/1000

total_sales = round(new_df['Carbon footprint, T'].sum(), 0)



new_df.insert(8, 'Result', '')
if total_sales <=200:
    new_df.at[0, 'Result']=('The carbon footprint is equal to ' + str(total_sales) + '.'+ ' The carbon footprint in this estimate is low, and the estimate can be considered environmentally efficient.')
elif 200 < total_sales <= 400:
    new_df.at[0, 'Result'] = ('The carbon footprint is equal to ' + str(total_sales) + '.' + ' In this case, the average carbon footprint is observed, some building materials can be changed.')
else:
    new_df.at[0, 'Result'] = ('The carbon footprint is equal to ' + str(total_sales) + '.' + ' In this estimate, the carbon footprint is high, it is recommended to choose another estimate for construction.')

    # Сохранение результатов в новый Excel-файл
new_df.to_excel('lemmatized_data.xlsx', index=False)

    # Вывод сообщения об успешном сохранении
print("Результаты были успешно сохранены в файл 'lemmatized_data.xlsx'.")
print(round(total_sales, 0))
    #rint('Углеродный след составляет ' + str(total_sales) + '.')
    #if total_sales > 300:
       # print('Углеродный след данной сметы строительногопроекта достаночно высок. Рекомендуем подобрать материалы с меньшим выбросом углеродного следа.')