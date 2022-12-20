#!/usr/bin/env python
# coding: utf-8

# In[1]:
!pip install beatifulsoup4

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import pandas
import os
import re
import xlwt
from pandas import Series, ExcelWriter
import glob
import random
from datetime import datetime
from lxml import etree
import streamlit as st


# TITLE FOR APP

st.title('Google PAA & Related Search Finder')
st.subheader('Find the most frequent PAAs and Related Searches for a given set of keywords.')


# SET USER AGENT LIST
user_agent_list = ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
                   'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
                   'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1.1 Safari/605.1.15']


# Full list of just the User Agent to rotate into loop
#user_agent_list = [
#'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1.1 Safari/605.1.15',
#'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:77.0) Gecko/20100101 Firefox/77.0',
#'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
#'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:77.0) Gecko/20100101 Firefox/77.0',
#'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
#'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36']


# Wait for FILE UPLOAD

st.write('Upload a an excel or csv file with two columns (Keyword and Search Volume) containing your list of keywords and their search volume')

file = st.file_uploader("Select your file.")

if file is not None:
    st.write("Your File is: ", uploaded_file.name)
    df=pd.read_excel(file)
    df['Google_URL'] = 'https://www.google.com/search?q=' + df['Keyword'].str.replace(' ', '%20')
else:
    st.warning('Please upload a file to get started.')
    st.stop()
    

st.button("Continue")

# function will pick a random user-agent
# inserts it into the headers variable to use in the loop
# new user_agent for each keyword
def randomize_headers():
    for i in range(1,5):
        #Pick a random user agent
        user_agent = random.choice(user_agent_list)
        headers = {'user-agent': user_agent,         
               'referrer': 'http://localhost:8888/',
               'sec-ch-ua': 'Not?A_Brand;v=8, Chromium;v=108, Google Chrome;v=108',
               'sec-ch-ua-full-version-list': '"Not?A_Brand";v="8.0.0.0", "Chromium";v="108.0.5359.100", "Google Chrome";v="108.0.5359.100"',
               'sec-ch-ua-mobile': '?0',
               'sec-ch-ua-model': ''}
    print("User agent is: " + str(user_agent))
    print(headers)
    return headers, user_agent


# In[10]:


# text_input = input('Cut and paste keywords separated by a comma or semi-colon. Or, \033[1m Press Enter \033[0mto upload a file from your computer.')


# In[11]:


# "C:\Users\Spencer Baselice\Documents\paa-test.csv"
print('Upload a file with two columns: \033[1mKeyword and Search Volume.\033[0m')
file = input("Copy and paste the full path and filename:\n")
print(file)

file = file.replace('"','').replace('\\' , '/')
print(file)


# In[12]:


header = input('Does your excel file have a header?\nType Yes or No:\n ')


# In[13]:


# if user types in No, n, no or N, or even Nada header is set to 0
# if they type in Yes, y, Y or even Yay! It does! then the header is set to 1 default
if 'N' in header:
    header=None
elif 'n' in header:
    header=None
else:
    header=0


# In[14]:


print('Your file is: '+ str(file))


# In[15]:


error = 'There has been an error loading your file. Make sure your file path does not contain\n extra spaces, periods or commas and is a valid csv or excel format.'


if '.csv' in file:
    try:
        df = pd.read_csv(file, header=header)
    except FileNotFoundError:
        print("File not found. Please enter the full path of the file.")
    except pd.errors.EmptyDataError:
        print("No data")
    except pd.errors.ParserError:
        print("Parse error")
    except Exception:
        print(error)
elif '.xlsx' in file:
    try:
        df = pd.read_excel(file, header=header, engine=None)
    except FileNotFoundError:
        print("File not found. Please enter the full path of the file.")
    except pd.errors.EmptyDataError:
        print("No data")
    except pd.errors.ParserError:
        print("Parse error")
    except Exception:
        print(error)
        
elif '.xls' in file:
    try:
        df = pd.read_excel(file, header=header, engine=None)
    except FileNotFoundError:
        print("File not found. Please enter the full path of the file.")
    except pd.errors.EmptyDataError:
        print("No data")
    except pd.errors.ParserError:
        print("Parse error")
    except Exception:
        print("Some other exception")        
elif '' in file:
    print(error)


# In[16]:


df.rename(columns={df.columns[0]: 'Keyword', df.columns[1]: 'Volume'},inplace=True)


# In[17]:


print('Here\'s the first 5 rows of your file: ') 
cols = df.head(5)
cols


# In[18]:


colist = list(cols)
for x in colist:
    print(x,cols[x][0:4])


# In[19]:


input("Verify this is the correct list. Type Yes and press ENTER to continue.")


# In[20]:


df['Google_URL'] = 'https://www.google.com/search?q=' + df['Keyword'].str.replace(' ', '%20')
df


# In[23]:


# loop creates two dicts, oe for PAAs and one for Related Searches
# Each request for the serp yeilds both PAAs and RS for that keyword
# Radnomized user agents are used to avoid bot detection
# DFs are transposed
count = 0
paa_scrape_dict = {'Keyword':[],
                   'Google_URL':[],
                   'Response_Code':[],
                   'People_Also_Ask':[]
                   }
related_scrape_dict = {'Keyword': [],
                       'Google_URL': [],
                       'Response_Code':[],
                       'Related_Search':[]
                      }
paa_div_capture = {'Keyword':[],
                   'Google_URL':[],
                   'Response_Code':[],
                   'People_Also_Ask':[],
                   'PAA_Div_Tag':[],
                   'Soup_Obj':[],
                   'User_Agent':[]}

#for (colname,colval) in df.iteritems():
# for i in colval.values: # gives you each value in the list
index=0
for column in df[['Google_URL']]:
    col_values = df['Google_URL']    
    for u in col_values:
        #execute function to pick a random user agent
        headers, user_agent = randomize_headers()
        time.sleep(2)
        #print(user_agent)
        url = u
        print(url)
        word = url.replace('https://www.google.com/search?q=', '').replace('%20', ' ')
        print(word)
        url_no = count=+1
        res = requests.get(url,headers=headers)
        response_code = res.status_code
        soup = BeautifulSoup(res.text, 'html.parser', from_encoding='iso-8859-1')
        #paa_extract = soup.find_all('div', class_='iDjcJe IX9Lgd wwB5gf')
        paa_div = soup.find_all('div', class_='wQiwMc related-question-pair')
        print(paa_div)
        paa_div_capture['People_Also_Ask'].append(paa_div)
        paa_div_capture['Keyword'].append(word)
        paa_div_capture['Google_URL'].append(url)
        paa_div_capture['Response_Code'].append(response_code)
        paa_div_capture['PAA_Div_Tag'].append(paa_div)
        paa_div_capture['Soup_Obj'].append(soup)
        paa_div_capture['User_Agent'].append(user_agent)
        if not paa_div:
                q = 'No PAAs Found'
                print(q)
                paa_scrape_dict['People_Also_Ask'].append(q)
                paa_scrape_dict['Keyword'].append(word)
                paa_scrape_dict['Google_URL'].append(url)
                paa_scrape_dict['Response_Code'].append(response_code)   
        else:
            for question in paa_div:
                print(question)
                q = question.get_text()
                print('Variable: ' + str(q))
                            
                if "Search for:" in q:
                    q = q.split('Search for:', 1)[0]
                    print('SearchFor found. Split q is ' + str(q))
                else:
                    continue
                paa_scrape_dict['People_Also_Ask'].append(q)
                paa_scrape_dict['Keyword'].append(word)
                paa_scrape_dict['Google_URL'].append(url)
                paa_scrape_dict['Response_Code'].append(response_code)
        dom = etree.HTML(str(soup))
        related_searches = dom.xpath('//span[contains(text(),"Related")]/ancestor::div[@data-hveid[string-length()>0]][position() = 1]//div[text()[string-length()>0]]')
        related_search_text_list = [BeautifulSoup(etree.tostring(s), "html.parser").get_text() for s in related_searches]
        for search in related_search_text_list:
            if '...' in search:
                print('found ... in ' + str(search) + ' continuing on.')
                continue
            else:
                related_scrape_dict['Related_Search'].append(search)
                print(search)
                related_scrape_dict['Keyword'].append(word)
                related_scrape_dict['Google_URL'].append(url)
                related_scrape_dict['Response_Code'].append(response_code)
        count = count+1

df_paa_scrape = pd.DataFrame.from_dict(paa_scrape_dict, orient='index')
df_paa = df_paa_scrape.transpose()
df_related_scrape = pd.DataFrame.from_dict(related_scrape_dict, orient='index')
df_related = df_related_scrape.transpose()
df_div = pd.DataFrame.from_dict(paa_div_capture, orient='index')
df_div.transpose()


# In[24]:


# Merge the scrape data with the input data to get the volume
paa_merge = df_paa.merge(df, how='left', on='Keyword')


# In[25]:


# clean it up, we don't need the google urls
paa_final = paa_merge.drop(['Google_URL_y', 'Google_URL_x'], axis=1).rename(columns={'Volume_y':'Volume'})


# In[26]:


print('Building top PAA with total search volume...')


# In[27]:


paa_final


# In[28]:


#Pivot PAAs to find the top amongst the set
paa_pivot = paa_final.groupby('People_Also_Ask').agg({'Volume': ['sum','count']}).sort_values(by=[('Volume', 'sum')], ascending=False)


# In[35]:


paa_pivot


# In[29]:


# Merge the Related Searches with input to get volume
related_merge = df_related.merge(df, how='left', on='Keyword')
related_final = related_merge.drop(['Google_URL_y', 'Google_URL_x'], axis=1).rename(columns={'Volume_y':'Volume'})


# In[36]:


related_final


# In[30]:


print('Building top Related Searches with total search volume...')


# In[31]:


# pivot the related searches to get the top amonst the set
related_pivot = related_final.groupby('Related_Search').agg({'Volume': ['sum','count']}).sort_values(by=[('Volume', 'sum')], ascending=False)


# In[37]:


related_pivot


# In[33]:


# Save the data to separate sheets in the same excel file
current_path = os.getcwd()
current_time = time.strftime("%m%d%y_%H%M%S")
path = str(current_path) + '\serp_scraper_results_' + str(current_time) + '.xlsx'
writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
paa_pivot.to_excel(writer, sheet_name = 'top_paas')
related_pivot.to_excel(writer, sheet_name = 'top_related_searches')
paa_final.to_excel(writer, sheet_name = 'paas_all')
related_final.to_excel(writer, sheet_name = 'related_searches_all')
df_div.to_excel(writer, sheet_name = 'scrape_data')
writer.close()

file_saved = glob.glob(path)


# In[34]:


# tell user where the file has been saved
print('Your file has been saved at: ' + str(path))


# In[ ]:



