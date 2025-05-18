#!/usr/bin/env python
# coding: utf-8

# In[2]:


get_ipython().system('pip install pandas openpyxl')


# In[3]:


import os
import pandas as pd


# In[41]:


data_file = r"C:\Users\USER\Desktop\Cleaning of Data & Merging into single excel\Payout Summary & Order Level Sales"


# In[42]:


invoice_files = [f for f in os.listdir(data_file) if f.endswith('.xlsx')]
print(f"Found {len(invoice_files)} invoice files.")


# In[43]:


sample_file = os.path.join(data_file, invoice_files[0])
excel = pd.ExcelFile(sample_file)
print("Sheet names:", excel.sheet_names)


# In[50]:


summary_data = []

for file in invoice_files:
    file_path = os.path.join(folder_path, file)
    print(f"Reading file: {file}")  # Debug line

    try:
        xl = pd.ExcelFile(file_path, engine='openpyxl')  # 👈 Specify engine
        df = xl.parse(xl.sheet_names[0])
        df['Source File'] = file
        summary_data.append(df)
    except Exception as e:
        print(f"❌ Error reading {file}: {e}")



# In[45]:


summary_data = []

for file in invoice_files:
    file_path = os.path.join(folder_path, file)
    print(f"Reading file: {file}")  # Debug line
    
    try:
        xl = pd.ExcelFile(file_path)
        df = xl.parse(xl.sheet_names[0])
        df['Source File'] = file
        summary_data.append(df)
    except Exception as e:
        print(f"❌ Error reading {file}: {e}")


# In[49]:


xl = pd.ExcelFile("C:/Users/USER/Desktop/Cleaning of Data & Merging into single excel/Payout Summary & Order Level Sales", engine='openpyxl')


# In[ ]:




