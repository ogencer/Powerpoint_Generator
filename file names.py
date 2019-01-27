
# coding: utf-8

# In[1]:


import os
import xlwt
import pandas as pd
import numpy as np


# In[2]:


#Change these accordingly.
excel_address = 'C:/Users/ozgur/OneDrive/Masaüstü/c.xlsx'
photo_address = 'C:/Users/ozgur/OneDrive/Masaüstü/PHOTO'


# In[3]:


def surnames(photo):
    file_list = [f for f in os.listdir(photo) if os.path.isfile(os.path.join(photo, f))]
    #print (file_list)
    surnames = []
    for a in file_list:
        split = a.split()
        result = split[-1].split('.')
        surnames.append([result[0], a]) 
        #print([result[0], a])
    return surnames
surnames(photo_address)


# In[11]:


df = pd.read_excel(excel_address)
df.describe()
df['Last Name:'] = df['Last Name:'].apply(lambda x: x.lower())
df['Last Name:'] = df['Last Name:'].apply(lambda x: x.capitalize())
df.columns = df.columns.str.replace('Last Name:','last')


# In[12]:


df['filenames'] = 'default value'
for x in surnames(photo_address):
    y = (x[0].lower()).capitalize()
    address = photo_address + x[1]
    address = address.replace("/", "\\")
    df.loc[df['last'] == y, 'filenames'] = address
    #df[df.last == y, 'file names'] = 'C:/Users/ozgur/OneDrive/Masaüstü/PHOTO/'+x[1]
    #df['Last Name:' == (x[0].lower()).capitalize(), 'file names'] = 'C:/Users/ozgur/OneDrive/Masaüstü/PHOTO/'+x[1]
    #df['file names'] = np.where(df['Last Name:' == x[0].lower].str.lower()==(x[0].lower()), 'C:/Users/ozgur/OneDrive/Masaüstü/PHOTO/'+x[1], ' ')
    #df.drop(columns=['(False, file names)'])
    #df.drop(df.columns[[6]], axis=1, inplace=True)
df


# In[13]:


df.to_excel(excel_address, engine='xlsxwriter')

