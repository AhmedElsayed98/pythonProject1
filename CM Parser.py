#!/usr/bin/env python
# coding: utf-8

# In[72]:


import pandas as pd
import xml.etree.ElementTree as et
import numpy as np


# In[73]:


xtree = et.parse("idLeModeLBXML.xml")
xroot = xtree.getroot()


# In[74]:


xroot[0][1].attrib


# In[75]:


classes


# In[76]:


Classes = np.unique(classes)


# In[77]:


Classes


# In[78]:


total = []
for class_ in Classes:
    cl = xroot.findall(".//*[@class='{0}']".format(class_))
    class_list = []
    for elem in cl:
        for i in range(len(elem)):
            elem.attrib.update({elem[i].attrib['name'] : elem[i].text})
        class_list.append(elem.attrib)
    total.append(class_list)


# In[99]:


total[0][0]['distName'].split("/")[1].split("-")[1]


# In[35]:


for i in range(len(total)):
    for j in range(len(total[i])):
        total[i][j]['class'] = total[i][j]['class'].split(":")[1]


# In[44]:


pd.DataFrame(total[3])


# In[50]:


class1 = []
for c in cl:
    class1.append(c.attrib)


# In[51]:


class1DF = pd.DataFrame(class1)


# In[52]:


class1DF


# In[49]:


class1DF[['Class']] = class1DF['class'].str.split(":", expand=True).str[1]


# In[19]:


class1DF.drop(['class'], axis=1, inplace=True)


# In[20]:


class1DF['MRBTS'] = class1DF['distName'].str.split("-").str[2]


# In[21]:


class1DF


# In[22]:


total = []
for class_ in len(Classes):
    cl = xroot.findall(".//*[@class='{0}']".format(class_))
    class_list = []
    for c in cl:
        class_list.append(c.attrib)
    total.append(class_list)


# In[23]:


pd.DataFrame(total[31])


# In[ ]:




