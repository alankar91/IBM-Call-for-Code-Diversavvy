#!/usr/bin/env python
# coding: utf-8

# In[31]:


# ! pip install PyMuPDF==1.18.9
# !pip install nltk

# !pip install PyPDF2


# In[32]:


import fitz
import os
import re
from io import BytesIO
import pandas as pd
import nltk
from nltk import tokenize
from IPython.display import display, HTML


# In[36]:


def highlight_pdf(file_name, search_terms):
    pdf_doc = fitz.open(f'Input_Files/{file_name}')
    output_buffer = BytesIO()

    for pg_num in range(pdf_doc.pageCount):
        page = pdf_doc[pg_num]
        highlight = None

        for term in search_terms:
            highlight_area = page.searchFor(term)
            highlight = page.addHighlightAnnot(highlight_area)
        
        highlight.update()
    
    pdf_doc.save(output_buffer)
    pdf_doc.close()

    with open(f'{file_name}', mode='wb') as f:
        f.write(output_buffer.getbuffer())
    
    return


def extract_text(file_name):
    pdf_doc = fitz.open(f'Input_Files/{file_name}')
    text = ''

    for pg_num in range(pdf_doc.pageCount):
        page = pdf_doc[pg_num]
        page_text = page.get_text()
        page_text = page_text.replace('\n', ' ')
        text += page_text
    
    return text

def highlight_sentences(sentences, key_terms):
    for sent in sentences:
        sent_lower = sent.lower()
        for term in key_terms:
            if term in sent_lower:
                start = sent_lower.index(term)
                end = start + len(term)
                html = '<p>' + sent[:start] + '<mark>' + sent[start:end] + '</mark>' + sent[end:] + '</p>'
                display(HTML(html))
    return 


def highlight_sentences_by_goal(df):
    found_goals_df = df[df['No_Of_Occurence'] > 0]
    goals = found_goals_df['Goals'].unique().tolist()

    for goal in goals:
        print(goal.center(120))
        print()
        goal_df = found_goals_df[found_goals_df['Goals'] == goal]

        for ind, row in goal_df.iterrows():
            sentences = row['Sentence'].splitlines()

            for sent in sentences:
                start = sent.lower().index(row['Search_Key'].lower())
                end = start + len(row['Search_Key'])
                html = '<p>' + sent[:start] + '<mark>' + sent[start:end] + '</mark>' + sent[end:] + '<b>' + row['File_Name'] + '</b>' + '</p>'
                display(HTML(html))
        print('\n\n')

    return


# In[37]:


def find_key_sentences(sentences, key_term):
    terget_sentences = []
    count = 0

    for sent in sentences:
        sent_lower = sent.lower()
        if key_term in sent_lower:
            terget_sentences.append(sent)
            count += 1
    
    return count, terget_sentences


def count_no_of_occurences(file_name, sentences, search_key_df):
    list_of_data = []

    for s_key, goal in zip(search_key_df['Key Terms'], search_key_df['Goals']):
        count, key_sentences = find_key_sentences(sentences, s_key.lower())

        result_dict = {}
        result_dict['File_Name'] = file_name
        result_dict['Search_Key'] = s_key
        result_dict['No_Of_Occurence'] = count
        result_dict['Goals'] = goal
        result_dict['Sentence'] = ' \n'.join(key_sentences)

        list_of_data.append(result_dict)
    
    return list_of_data



# In[46]:


# @hidden_cell
# The following code contains the credentials for a file in your IBM Cloud Object Storage.
# You might want to remove those credentials before you share your notebook.
credentials_3 = {
    'IAM_SERVICE_ID': 'iam-ServiceId-4e110c47-4f05-44c1-a9f7-7393b1d9d4a1',
    'IBM_API_KEY_ID': '5guqPumJkfgaqbO3tmASQRBfnvxwoUPM24XYyis0qE4p',
    'ENDPOINT': 'https://s3.private.eu.cloud-object-storage.appdomain.cloud',
    'IBM_AUTH_ENDPOINT': 'https://iam.cloud.ibm.com/oidc/token',
    'BUCKET': 'annualreportsmining-donotdelete-pr-wvxdzzga1kwkci',
    'FILE': 'META.pdf'
}
# @hidden_cell
# The following code contains the credentials for a file in your IBM Cloud Object Storage.
# You might want to remove those credentials before you share your notebook.
credentials_2 = {
    'IAM_SERVICE_ID': 'iam-ServiceId-4e110c47-4f05-44c1-a9f7-7393b1d9d4a1',
    'IBM_API_KEY_ID': '5guqPumJkfgaqbO3tmASQRBfnvxwoUPM24XYyis0qE4p',
    'ENDPOINT': 'https://s3.private.eu.cloud-object-storage.appdomain.cloud',
    'IBM_AUTH_ENDPOINT': 'https://iam.cloud.ibm.com/oidc/token',
    'BUCKET': 'annualreportsmining-donotdelete-pr-wvxdzzga1kwkci',
    'FILE': 'APPLE.pdf'
}
# @hidden_cell
# The following code contains the credentials for a file in your IBM Cloud Object Storage.
# You might want to remove those credentials before you share your notebook.
credentials_1 = {
    'IAM_SERVICE_ID': 'iam-ServiceId-4e110c47-4f05-44c1-a9f7-7393b1d9d4a1',
    'IBM_API_KEY_ID': '5guqPumJkfgaqbO3tmASQRBfnvxwoUPM24XYyis0qE4p',
    'ENDPOINT': 'https://s3.private.eu.cloud-object-storage.appdomain.cloud',
    'IBM_AUTH_ENDPOINT': 'https://iam.cloud.ibm.com/oidc/token',
    'BUCKET': 'annualreportsmining-donotdelete-pr-wvxdzzga1kwkci',
    'FILE': 'ALPHABET.pdf'
}
# @hidden_cell
# The following code contains the credentials for a file in your IBM Cloud Object Storage.
# You might want to remove those credentials before you share your notebook.
credentials_4 = {
    'IAM_SERVICE_ID': 'iam-ServiceId-4e110c47-4f05-44c1-a9f7-7393b1d9d4a1',
    'IBM_API_KEY_ID': '5guqPumJkfgaqbO3tmASQRBfnvxwoUPM24XYyis0qE4p',
    'ENDPOINT': 'https://s3.private.eu.cloud-object-storage.appdomain.cloud',
    'IBM_AUTH_ENDPOINT': 'https://iam.cloud.ibm.com/oidc/token',
    'BUCKET': 'annualreportsmining-donotdelete-pr-wvxdzzga1kwkci',
    'FILE': 'MICROSOFT.pdf'
}
# @hidden_cell
# The following code contains the credentials for a file in your IBM Cloud Object Storage.
# You might want to remove those credentials before you share your notebook.
credentials_5 = {
    'IAM_SERVICE_ID': 'iam-ServiceId-4e110c47-4f05-44c1-a9f7-7393b1d9d4a1',
    'IBM_API_KEY_ID': '5guqPumJkfgaqbO3tmASQRBfnvxwoUPM24XYyis0qE4p',
    'ENDPOINT': 'https://s3.private.eu.cloud-object-storage.appdomain.cloud',
    'IBM_AUTH_ENDPOINT': 'https://iam.cloud.ibm.com/oidc/token',
    'BUCKET': 'annualreportsmining-donotdelete-pr-wvxdzzga1kwkci',
    'FILE': 'ORACLE.pdf'
}


# In[40]:


import os, types
import pandas as pd
from botocore.client import Config
import ibm_boto3

def __iter__(self): return 0

# @hidden_cell
# The following code accesses a file in your IBM Cloud Object Storage. It includes your credentials.
# You might want to remove those credentials before you share the notebook.
client_0be1e42846844fd4b0d5f43192b1bb02 = ibm_boto3.client(service_name='s3',
    ibm_api_key_id='5guqPumJkfgaqbO3tmASQRBfnvxwoUPM24XYyis0qE4p',
    ibm_auth_endpoint="https://iam.cloud.ibm.com/oidc/token",
    config=Config(signature_version='oauth'),
    endpoint_url='https://s3.private.eu.cloud-object-storage.appdomain.cloud')

body = client_0be1e42846844fd4b0d5f43192b1bb02.get_object(Bucket='annualreportsmining-donotdelete-pr-wvxdzzga1kwkci',Key='SearchTerms.xlsx')['Body']

df_data_0 = pd.read_excel(body.read())
df_data_0.head()

df=df_data_0


# In[47]:


from ibm_botocore.client import Config
import ibm_boto3

cos = ibm_boto3.client(service_name='s3',
    ibm_api_key_id=credentials_1['IBM_API_KEY_ID'],
    ibm_service_instance_id=credentials_1['IAM_SERVICE_ID'],
    ibm_auth_endpoint=credentials_1['IBM_AUTH_ENDPOINT'],
    config=Config(signature_version='oauth'),
    endpoint_url=credentials_1['ENDPOINT'])


# In[ ]:


for bucket in cos.list_buckets()['Buckets']:
    print(bucket['Name'])


# In[ ]:


files = os.listdir('Input_Files')
pdf_files = sorted([f for f in files if f.endswith('.pdf')])

search_df = pd.read_excel('SearchTerms.xlsx')
search_terms = df['Key Terms']

for file_name in pdf_files:
    highlight_pdf(file_name, search_terms)

result_list = []

for file_name in pdf_files:
    text = extract_text(file_name)
    sentences = tokenize.sent_tokenize(text)

    print(file_name.center(120))
    print()
    highlight_sentences(sentences, search_terms.str.lower())
    print('\n\n\n')


    file_result = count_no_of_occurences(file_name, sentences, search_df)
    result_list.extend(file_result)


df = pd.DataFrame(result_list)


# In[ ]:


pivot_table = df.groupby(['File_Name', 'Goals']).sum(['No of Occurance']).unstack()
pivot_table = pivot_table.droplevel(0, axis=1).reset_index()
df.to_excel('D&I-Results.xlsx', index=False)
pivot_table.to_excel('D&I-Pivot Table.xlsx', index=False)

