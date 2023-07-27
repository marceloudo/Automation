#!/usr/bin/env python
# coding: utf-8

# # <font color=green>SAP FS10N - VARIABLE EXPENSES</font>
# ***
# DATA EXTRACTION
# ***

# In[236]:


# Importing the Libraries
import win32com.client
import pandas as pd
from datetime import datetime
import subprocess


# ## <font color=green>1 - START</font>
# ***
# Always we have those variables except st_string

# In[237]:


SapGuiAuto = win32com.client.GetObject('SAPGUI')
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)
now = datetime.now()
dt_string = now.strftime("%Y%m%d %H%M")
dt_string


# ## <font color=green>2 - Folder and Filename</font>
# ***
# Take a look at folderdir that has those "\\" marks different from we see by using OS.

# In[238]:


filename = "FS10N-" + dt_string + ".XLSX"
folderdir = "C:\\Users\\20056306\\HASH\\"
folderdir + filename


# ## <font color=green>3 - CALL FS10N</font>
# ***
# Calling SAP Transaction

# In[239]:


session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nFS10N"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/tbar[1]/btn[17]").press()


# ## <font color=green>4 - VARIANT BOX</font>
# ***
# Dealing with the variant box by letting it empty.

# In[240]:


session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press()


# ## <font color=green>5 - VARIANT REGARDING VARIABLE EXPENSES</font>
# ***
# In the list of variants, we gotta choose the VARIABLE EXPENSES variant.

# In[241]:


session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 15
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").firstVisibleRow = 9
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "15"
session.findById("wnd[1]/tbar[0]/btn[2]").press()


# ## <font color=green>6- SET UP TRANSACTION</font>
# ***
# Setting up transaction by defining Year and Month.

# In[242]:


session.findById("wnd[0]/usr/txtGP_GJAHR").Text = 2023
session.findById("wnd[0]/usr/txtGP_GJAHR").SetFocus()
session.findById("wnd[0]/usr/txtGP_GJAHR").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press()


# ## <font color=green>7 - MONTHS AND VARIABLE EXPENSES FIGURES</font>
# ***
# Choose one of them to have the details.

# In[243]:


session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").setCurrentCell("1", "BALANCE")
session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").doubleClickCurrentCell()
session.findById("wnd[0]/tbar[1]/btn[33]").press()


# ## <font color=green>8 - LAYOUT BOX</font>
# ***
# Choosing one of the layouts.

# In[244]:


session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(151, "TEXT")
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 145
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "151"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()


# ## <font color=green>9 - DATA EXPORT</font>
# ***
# Generating excel file.

# In[245]:


session.findById("wnd[0]").maximize()
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderdir
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()
session.findById("wnd[0]/tbar[0]/btn[3]").press()


# ## <font color=green>10 - END OF SESSION VARIABLES</font>
# ***
# 

# In[246]:


# Run SAP Scriptsession = None
connection = None
application = None
SapGuiAuto = None


# ## <font color=green>11 - DATA TABLE</font>
# ***
# Bringing data to the notebook.

# In[247]:


df = pd.read_excel(folderdir + filename, sheet_name = 'Sheet1')
df


# ## <font color=green>12 - END OF EXCEL SESSION</font>
# ***

# In[248]:


subprocess.call(["taskkill","/f","/im","EXCEL.EXE"])

