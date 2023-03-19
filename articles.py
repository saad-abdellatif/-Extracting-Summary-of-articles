# pip install spacy pandas selenium
# python -m spacy download en_core_web_md
import pandas as pd
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
import csv
import openpyxl
import sqlite3
from selenium.webdriver import ActionChains
import spacy
from spacy.lang.en.stop_words import STOP_WORDS
import os

# Define function to remove duplicates and irrelevant data
def clean_data(df):
    df.drop_duplicates(inplace=True, keep='first')
    df.reset_index(inplace=True, drop=True)
    df = df[df.Website != "Website"]
    df = df[df["ArticleTitle"] != "No Data"]
    df.dropna(inplace=True)
    df.dropna(subset=['ArticleTitle'], inplace=True)
    return df

# To remove duplicates and irrelevant data using spaCy
def extract_unique_entities(text):
    doc = nlp(text)
    entities = set()
    for ent in doc.ents:
        if ent.label_ in ['ORG', 'PRODUCT', 'PERSON', 'NORP', 'FAC', 'GPE', 'LOC']:
            entities.add(ent.text)
    return frozenset(entities)

# Load data
df = pd.read_csv('content.csv')

# Clean data
df = clean_data(df)

# Available trained pipelines for English
nlp = spacy.load("en_core_web_md")

df['UniqueEntities_content'] = df['ArticleContent'].apply(extract_unique_entities)
df['UniqueEntities_title'] = df['ArticleTitle'].apply(extract_unique_entities)

df.drop_duplicates(subset='UniqueEntities_content', keep='first', inplace=True)
df.drop_duplicates(subset='UniqueEntities_title', keep='first', inplace=True)

# Drop the 'UniqueEntities' column if no longer needed
df.drop(columns=['UniqueEntities_content','UniqueEntities_title'], inplace=True)

# Exported the cleaned data to a CSV file
df.to_csv('Data_Cleaned.csv', index= False)

#reading the file
with open('Data_Cleaned.csv', 'r') as csvfile:
    reader = csv.DictReader(csvfile)
    # Load the CSV data into a list
    data = list(reader)

# Connect to the database and create the articles table if it doesn't exist
conn = sqlite3.connect('articles.db')
c = conn.cursor()
c.execute('''
          CREATE TABLE IF NOT EXISTS articles (
              id INTEGER PRIMARY KEY,
              Website TEXT,
              ArticleURL TEXT,
              ArticleTitle TEXT,
              ArticleDate TEXT,
              ArticleImage_url TEXT,
              ArticleContent TEXT,
              Summary TEXT
          )
          ''')
conn.commit()

# Summarizing the data using QuillBot
base_url = "https://quillbot.com/summarize"
driver = webdriver.Chrome()
driver.maximize_window()
driver.implicitly_wait(5)
driver.get(base_url)
sleep(5)

try:
    driver.find_element(By.XPATH,'//*[text()="Accept All"]').click()
    sleep(0.5)
except: pass

'''
try:
    fromElement =  driver.find_element(By.XPATH,'//*[@id="root-client"]/div[2]/div[3]/section[1]/div/div/div/div[1]/div/div/div/div[1]/div[2]/div[3]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/span/span[4]')
    toElement = driver.find_element(By.XPATH,'//*[@id="root-client"]/div[2]/div[3]/section[1]/div/div/div/div[1]/div/div/div/div[1]/div[2]/div[3]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/span/span[5]')
    #toElement = driver.find_element(By.XPATH,'//*[@id="root-client"]/div[2]/div[3]/section[1]/div/div/div/div[1]/div/div/div/div[1]/div[2]/div[3]/div/div[1]/div/div[2]/div/div[2]/div/div[2]/span/span[3]')
    actions = ActionChains(driver)
    actions.drag_and_drop(fromElement, toElement).perform() 
    sleep(1)
except:
    print("Please, dragt it to the short length Summary Manually")
    sleep(30)
'''

# Insert the articles data into the database
with open('Data_Cleaned.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        Website = row['Website']
        ArticleURL = row['ArticleURL']
        ArticleTitle = row['ArticleTitle']
        ArticleDate = row['ArticleDate']
        ArticleImage_url = row['ArticleImage_url']
        ArticleContent = row['ArticleContent']

        # Checking if the record has a summary or not
        c.execute('SELECT id, ArticleContent, Summary FROM articles WHERE ArticleURL=?', (ArticleURL,))
        existing_data = c.fetchone()
        if existing_data:
            existing_id = existing_data[0]
            article_content = existing_data[1]
            summary = existing_data[2]
            
            if summary and summary.strip() != '':
                print(f"{existing_id} - Article with URL '{ArticleURL}' has summary. Skipping it.\n")
                continue
        
        
        # Get the text and summarize it
        input_box = driver.find_element(By.ID,'inputBoxSummarizer')
        input_box.clear()
        input_box.send_keys(ArticleContent)
        sleep(1)
        submit_button = driver.find_element(By.XPATH,'//button[@type="button"]/div[text()="Summarize"]')
        driver.execute_script("arguments[0].click();", submit_button)
        sleep(5)
        summary_box = driver.find_element(By.ID,'outputBoxSummarizer')
        summary = summary_box.text
        
        # Update the database with the summary text
        if existing_data:
            # Update the existing record with new data
            c.execute('UPDATE articles SET Website=?, ArticleTitle=?, ArticleDate=?, ArticleImage_url=?, ArticleContent=?, Summary=? WHERE id=?',
                    (Website, ArticleTitle, ArticleDate, ArticleImage_url, ArticleContent, summary, existing_data[0]))
        else:
            # Insert a new record
            c.execute('INSERT INTO articles (Website, ArticleURL, ArticleTitle, ArticleDate, ArticleImage_url, ArticleContent, Summary) VALUES (?, ?, ?, ?, ?, ?, ?)',
                    (Website, ArticleURL, ArticleTitle, ArticleDate, ArticleImage_url, ArticleContent, summary))

        conn.commit()
        

# Close the database connection
conn.close()

# Export the data to an Excel file
conn = sqlite3.connect('articles.db')
c = conn.cursor()
c.execute('SELECT Website, ArticleURL, ArticleTitle, ArticleDate, ArticleImage_url, ArticleContent, Summary FROM articles')
data = c.fetchall()
conn.close()

wb = openpyxl.Workbook()
ws = wb.active
header = ['Website', 'ArticleURL', 'ArticleTitle', 'ArticleDate', 'ArticleImage_url', 'ArticleContent', 'Summary']
ws.append(header)

for row in data:
    ws.append(row)
wb.save('Articles.xlsx')

# Close the browser
driver.quit()

# Delete the cleaned file and the database file
os.remove("Data_Cleaned.csv")
os.remove("articles.db")

