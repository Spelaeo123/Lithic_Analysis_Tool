#!/usr/bin/env python
# coding: utf-8

# # Jacobi Lithic Analysis Data Entry Tool

# ## Import libraries and modules

# In[1]:


import os
import sys
import PySimpleGUI as sg
import pandas as pd
import openpyxl
from openpyxl import load_workbook


# ## Import input/output file

# In[ ]:





# ## Columns in excel file

# In[2]:


cols = ['SF Number', 'Site', 'Category', 'Post-deposition damage', 'Cortication type', 'Butt type', 'Cortex type', 'Cortex thickness', 'Other details']


# ## Attribute Lists

# In[3]:


lst = ('Site 1','Site 2','Site 3','Site 4','Site 5','Site 6','Site 7','Other')
CatLst = ('1. Flake', '2. Blade', '3. Bladelet','4. Blake-like flake (DO NOT USE)','5. Irregular Waste','6. Chip','7. Microburin','8. Burin spall','9. Rejuvenation flake','10. Core tablet','11. Rejuvenation flake other','12. Levallois flake','13. Janus flake (= thinning)','14. Axe working flake','15. Flake from ground implement','16. Single platform blade core','17. Opposed platform blade core','18. Other Blade core','19. Tested nodule/ bashed lump','20. Single platform flake core','21. Multiplatform flake core','22. Keeled non-discoidal flake core','23. Levallois/ other discoidal flake core','24. Unclassifiable/ fragmentary core','25. Microlith','26. Petit tranchet arrowhead','27. Leaf arrowhead','28. Chisel arrowhead','29. Oblique arrowhead','30. Barbed and tanged arrowhead','31. Triangular arrowhead','32. Hollow-based arrowhead','33. Laurel leaf','34. Unfinished arrowhead/blank','35. Fragmentary/unclass/other arrowhead','36. End scraper','37. Side scraper','38. End and side scraper','39. Disc scraper','40. Thumbnail scraper','41. Scraper on a non-flake blank','42. Other scraper','43. Awl','44. Piercer','45. Spurred piece','46. Other borer','47. Serrated flake','48. Saw','49. Denticulate','50. Notch','51. Backed knife','52. Scale flaked knife','53. Plano-convex knife','54. Other knife','55. Flake retouched','56. Single-piece sickle','57. Fabricator','58. Axe','59. Other heavy implement','60. Misc retouch','61. Other retouch','62. Burnt unworked','63. Hammerstone','64. Natural','65. Core on a flake','66. Gun flint','67. Axe sharpening flake','68. Sieved chips 10-4mm','69. Sieved chips 4-2mm','70. Sieved chips','71. Bruised blade/flake','72. Burin','73. Blade crested','74. Blade retouched','75. End truncation straight','76. End truncation oblique','77. Backed bladelet','78. Bipolar core')
lst2 = ('1. Fresh','2. Slight post depositional damage', '3. Moderate post depositional damage','4. Heavy post depositional damage','5. Plough damaged','6. Rolled','7. Glossed','8. Modern damage')
lst3 = ('1. No cortication','2. Light cortication','3. Moderate cortication','4. Heavy cortication','5. Very heavy cortication','6. Iron stained')
lst4 = ('1. Cortical (cortex)','2. Plain (unaltered inner flint)','3. Dihedral (2 scars)','4. Faceted (more than 2 scars)','5. Linear (very narrow and usually quite long)','6. Shattered','7. Thermal (old recorticated or patinated surface)','8. Punctiform','9. Modified (retouched or otherwise altered)','10. Indeterminate (if you can’t decide or it doesn’t fit any of the above)')
lst5 = ('1. Chalk – thick, creamy white, chalky', '2. Weathered Chalk – like above but clearly worn and probably thin or moderate in thickness','3. Weathered – Similar to WC above, will be rough to the touch but clearly ground down.','4. Rolled – typical beach pebble/ river gravel flint, smooth','5. Thermal – re-corticated old surface, often will be in','6. Indeterminate','7. Banded -eg. Bullhead flint')
lst6 = ('1. Thin – up to 2mm thick (judge don’t measure)', '2. Moderate 2-4mm thickness', '3. Thick 4-10mm', '4. Very thick, greater than 10mm')


# ## Column sizes

# In[4]:


colsize = [20,60]


# ## Font sizes

# In[5]:


hfont = ("Calibri", 12)
font = ("Calibri", 10)


# ## Columns layout

# In[6]:


#Select Post-deposition damage
col1 = [[sg.Text('Post-deposition damage:', font=hfont)],
        [sg.Radio(text, "Radio2", enable_events=True, key=f"Radio2 {i}", size=15, font=font, pad=0) 
         for i, text in enumerate(lst2)]
       ]

#Select Cortication
col2 = [[sg.Text('Cortication type:', font=hfont)],
        [sg.Radio(text, "Radio3", enable_events=True, key=f"Radio3 {i}", size=15, font=font, pad=0) 
         for i, text in enumerate(lst3)]
       ]

#Select Butt type
col3 = [[sg.Text('Butt type:', font=hfont)],
        [sg.Radio(text, "Radio4", enable_events=True, key=f"Radio4 {i}", size=15, font=font, pad=0) 
         for i, text in enumerate(lst4)]
       ]

#Select Cortex type 
col4 = [[sg.Text('Cortex type:', font=hfont)], 
        [sg.Radio(text, "Radio5", enable_events=True, key=f"Radio5 {i}", size=15, font=font, pad=0) 
         for i, text in enumerate(lst5)]
       ]

#Select Cortex thickness
col5 = [[sg.Text('Cortex thickness:', font=hfont)], 
        [sg.Radio(text, "Radio6", enable_events=True, key=f"Radio6 {i}", size=15, font=font, pad=0) 
         for i, text in enumerate(lst6)]
       ]


# ## Layout and window

# In[7]:


layout = [[
    #Browse to file  
    [sg.Text('Select output file:', font=hfont), sg.Input(), sg.FilesBrowse(key='-IN-'), sg.B('OK')],
    #Title
    [sg.Text('Please fill the following fields:', font=hfont)],
    #Type SF Number
    [sg.Text('SF Number', size=(15,1), font=hfont), sg.InputText(key='SF_Number')],
    #Select Site
    [sg.Text('Site:', font=hfont)],
    [sg.Radio(text, "Radio", enable_events=True, key=f"Radio1 {i}", font=font) 
        for i, text in enumerate(lst)],
    #Select Category
    [sg.Text('Category', size=(15,1), font=hfont), sg.Combo(['1. Flake', '2. Blade', '3. Bladelet','4. Blake-like flake (DO NOT USE)','5. Irregular Waste','6. Chip','7. Microburin','8. Burin spall','9. Rejuvenation flake','10. Core tablet','11. Rejuvenation flake other','12. Levallois flake','13. Janus flake (= thinning)','14. Axe working flake','15. Flake from ground implement','16. Single platform blade core','17. Opposed platform blade core','18. Other Blade core','19. Tested nodule/ bashed lump','20. Single platform flake core','21. Multiplatform flake core','22. Keeled non-discoidal flake core','23. Levallois/ other discoidal flake core','24. Unclassifiable/ fragmentary core','25. Microlith','26. Petit tranchet arrowhead','27. Leaf arrowhead','28. Chisel arrowhead','29. Oblique arrowhead','30. Barbed and tanged arrowhead','31. Triangular arrowhead','32. Hollow-based arrowhead','33. Laurel leaf','34. Unfinished arrowhead/blank','35. Fragmentary/unclass/other arrowhead','36. End scraper','37. Side scraper','38. End and side scraper','39. Disc scraper','40. Thumbnail scraper','41. Scraper on a non-flake blank','42. Other scraper','43. Awl','44. Piercer','45. Spurred piece','46. Other borer','47. Serrated flake','48. Saw','49. Denticulate','50. Notch','51. Backed knife','52. Scale flaked knife','53. Plano-convex knife','54. Other knife','55. Flake retouched','56. Single-piece sickle','57. Fabricator','58. Axe','59. Other heavy implement','60. Misc retouch','61. Other retouch','62. Burnt unworked','63. Hammerstone','64. Natural','65. Core on a flake','66. Gun flint','67. Axe sharpening flake','68. Sieved chips 10-4mm','69. Sieved chips 4-2mm','70. Sieved chips','71. Bruised blade/flake','72. Burin','73. Blade crested','74. Blade retouched','75. End truncation straight','76. End truncation oblique','77. Backed bladelet','78. Bipolar core'], key='Category')],   
    [sg.Column(col1)],
    [sg.Column(col2)],
    [sg.Column(col3)],
    [sg.Column(col4)],
    [sg.Column(col5)], 
    [sg.Text('Other comments:', font=hfont),sg.Input('', key='Other Details', do_not_clear=False)],
    [sg.Push()], 
    [sg.Button("Submit"), sg.Button("Clear"), sg.Button('Exit')],
]]

window = sg.Window("Jacobi: Lithic Analyis Data Entry Tool v0.5", layout, resizable=True, element_justification='l', finalize=True)

print(layout)


# ## clear input function

# In[8]:


def clear_input():
    keys_to_clear = ["SF_Number", "Category", "Other Details"]
    for key in keys_to_clear:
        window[key]('')
    window["Radio1 0"].reset_group()
    window["Radio2 0"].reset_group()
    window["Radio3 0"].reset_group()
    window["Radio4 0"].reset_group()
    window["Radio5 0"].reset_group()
    window["Radio6 0"].reset_group()
    return None


# ## While statement

# In[9]:


while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, 'Exit'):
        break
    if event == 'Clear':
        clear_input()
    if event == "OK":
        filename = values['-IN-']
    elif event == 'Submit':
       
        radio_value = window['Radio1 0'].TKIntVar.get()
        text = lst[radio_value % 1000] if radio_value else None
        
        radio_value2 = window['Radio2 0'].TKIntVar.get()
        text2 = lst2[radio_value2 % 1000] if radio_value2 else None
        
        radio_value3 = window['Radio3 0'].TKIntVar.get()
        text3 = lst3[radio_value3 % 1000] if radio_value3 else None       
        
        radio_value4 = window['Radio4 0'].TKIntVar.get()
        text4 = lst4[radio_value4 % 1000] if radio_value4 else None
        
        radio_value5 = window['Radio5 0'].TKIntVar.get()
        text5 = lst5[radio_value5 % 1000] if radio_value5 else None
        
        radio_value6 = window['Radio6 0'].TKIntVar.get()
        text6 = lst6[radio_value6 % 1000] if radio_value6 else None
        
        record = [values['SF_Number'], text, values['Category'], text2, text3, text4, text5, text6, values['Other Details']]
        print(record)
        df_record = pd.DataFrame([record], columns=cols)
        print(df_record)
        df = pd.read_excel(filename, engine="openpyxl")
        df = df.append(df_record, ignore_index=True, sort=False)
        print(df)
        with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, index=False)
            writer.save()
        #writer.close()
        sg.popup('Data saved!')
        clear_input()

window.close()
sys.exit(1)


# In[ ]:





# In[ ]:




