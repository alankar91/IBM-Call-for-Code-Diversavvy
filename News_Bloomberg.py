#!/usr/bin/env python
# coding: utf-8

# In[1]:


import http.client

conn = http.client.HTTPSConnection("bb-finance.p.rapidapi.com")

headers = {
    'X-RapidAPI-Key': "201a4897bamsh4dc90f4f1dd54c7p113f64jsnff37c6c6f4fb",
    'X-RapidAPI-Host': "bb-finance.p.rapidapi.com"
    }

conn.request("GET", "/market/auto-complete?query=oracle", headers=headers)

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))


# In[2]:


#import libraries used below
import http.client
import json
from datetime import datetime
from pathlib import Path


# In[3]:


# res = response.getresponse()
# data = response
# Load API response into a Python dictionary object, encoded as utf-8 string
json_dictionary = json.loads(data.decode("utf-8"))


# In[4]:


list_title=[]
list_link=[]
list_published=[]
list_date=[]
for i in json_dictionary['news']:
    list_title.append(i['title'])
    list_link.append(i['longURL'])
    list_published.append(i['card'])
    list_date.append(i['date'])
#     print(i['title'])


# In[5]:


import pandas as pd
# Calling DataFrame after zipping both lists, with columns specified 
df_l_2 = pd.DataFrame(list(zip(list_title,list_link,list_published,list_date)), columns =['title', 'website','published','date']) 


# In[6]:


get_ipython().system('pip install vaderSentiment')
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer


# In[7]:


analyzer = SentimentIntensityAnalyzer()


# In[8]:


single=df_l_2['title']
train_idx = [i for i in range(len(single.index))]
# Convert to numpy
x_train = single.values[train_idx]


# In[9]:


sentences=x_train
analyzer = SentimentIntensityAnalyzer()
for sentence in sentences:
    vs = analyzer.polarity_scores(sentence)


# In[10]:


predicted_class_id=[]
for sentence in sentences:
    x = analyzer.polarity_scores(sentence)
    predicted_class_id.append(x)


# In[11]:


df_pred=pd.DataFrame(predicted_class_id)
df_pred.head()


# In[12]:


f_df=pd.merge(df_l_2, df_pred, left_index=True, right_index=True)


# In[13]:


f_df.head()


# In[15]:


f_df.to_excel("News_Data.xlsx", index = False)


# In[ ]:




