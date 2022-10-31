#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd  
import xlwings as xw





initial_version ='./TDGData/SecondTable.csv'
print("SecondTable loaded")
updated_version ='./TDGData/FirstTable.csv'
print("FirstTable loaded")

# In[3]:


df_initial = pd.read_csv(initial_version)
df_initial.head(10)


# In[4]:


df_update = pd.read_csv(updated_version)
df_update.head(10)


# In[5]:


df_initial.shape

# In[6]:


df_update.shape


# In[7]:


df_initial.shape == df_update.shape


# In[8]:


df_update = df_update.reset_index()
df_update.head(3)


# In[9]:


df_diff = pd.merge(df_initial, df_update, how="outer", indicator="Exist")
d={"left_only":"Only present in new file", "right_only":"Only present in old file","both":"Present in Both"}
df_diff['Exist']= df_diff['Exist'].map(d)
df_diff


# In[10]:


df_diff = df_diff.query("Exist != 'Present in Both'")
df_diff


# In[11]:


df_highlightNew = df_diff.query("Exist == 'Only present in new file'")
df_highlightNew


df_highlightNew.to_csv("./TDGData/Diff_High_New.csv")
print("Differences in New file created")


# In[12]:


df_highlightOld = df_diff.query("Exist == 'Only present in old file'")
df_highlightOld

df_highlightOld.to_csv("./TDGData/Diff_High_Old.csv")
print("Differences in Old file created")



# In[13]:


highlight_rows = df_highlightOld['index'].tolist()
highlight_rows = [int(row) for row in highlight_rows]
highlight_rows


# In[14]:


first_row_in_excel = 2

highlight_rows = [x + first_row_in_excel for x in highlight_rows]
highlight_rows

print("Row number = ")
print(highlight_rows)


# In[15]:


# with xw.App(visible=False) as app:
#     updated_wb = app.books.open(updated_version)
#     updated_ws = updated_wb.sheets(1)
#     rng = updated_ws.used_range

#     print(f"Used Range: {rng.address}")

#     # Hightlight the rows in Excel
#     for row in rng.rows:
#         if row.row in highlight_rows:
#             row.color = (255, 71, 76)  # light red

#     updated_wb.save("./TDGData/Difference_Highlighted.xlsx")
#     print("Completed")


# In[ ]:




