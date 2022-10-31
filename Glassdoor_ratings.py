#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.request import Request, urlopen


# In[2]:


start_page = 1
end_page = 5
company_list = []

for page in range(start_page, end_page+1):
    search_url = f"https://www.glassdoor.com/Explore/browse-companies.htm?overall_rating_low=3&page={page}&sector=10013&filterType=RATING_DIVERSITY_AND_INCLUSION"
    hdr = {'User-Agent': 'Mozilla/5.0'}
    search_req = Request(search_url,headers=hdr)
    search_res = urlopen(search_req)
    search_soup = BeautifulSoup(search_res, "html.parser")
    companies = search_soup.select('.row.flex-wrap')

    for com in companies:
        rating_dict = {}
        rating_dict['Company'] = com.h2.text
        rating_dict['Company Review'] = com.select_one('.employerInfo__EmployerInfoStyles__ratingSeparator b').text
        rating_dict['D&I Score'] = com.select_one('.pl-md-xsm b').text

        company_list.append(rating_dict)


# In[3]:


rating_df = pd.DataFrame(company_list)


# In[4]:


rating_df.to_excel('Glassdoor_rating.xlsx', index= False)


# In[5]:


rating_df


# In[ ]:




