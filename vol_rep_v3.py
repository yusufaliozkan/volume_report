import pandas as pd
import re
import numpy as np
from operator import index
import io
import streamlit as st
from io import BytesIO
from datetime import date
from io import StringIO

# Setting the app page layout
st.set_page_config(layout = "centered", page_title='Volume Report processing tool', page_icon='https://res.cloudinary.com/crunchbase-production/image/upload/c_lpad,h_256,w_256,f_auto,q_auto:eco,dpr_1/ol2lcon3zrdrcu6v1jo7')
path='https://upload.wikimedia.org/wikipedia/en/thumb/3/32/Logo_for_Imperial_College_London.svg/2560px-Logo_for_Imperial_College_London.svg.png'
st.image(path, width=300)

st.title('Spiral Symplectic Report processing tool') 

# st.sidebar.markdown("# Spiral Symplectic Report processor")

with st.expander('Instructions'):
    st.markdown('* Copy the text of volume report') 
    st.markdown('* Paste into the box below')
    st.markdown("* Press '**Ctrl and Enter**' to apply")
    st.markdown("* Click '**Download report**'")
    st.markdown('* Save the file to [SharePoint](https://imperiallondon.sharepoint.com/:f:/r/sites/GreenTeam2/Shared%20Documents/General/Spiral%20Symplectic%20(Volume)%20report?csf=1&web=1&e=okEv4g)')

txt = st.text_area('Paste the volume report here:', '''''')
df = pd.DataFrame(StringIO(txt))

if len(txt)>0:
    st.success('Thank you for inserting the text! You can now download the report.')
    if df is not None:
        df1 = df.drop([0])    
        df1[1] = df1[0].str.extract('Spiral:(.*)')
        pattern = r'(https?:\/\/(?:www\.)?[-a-zA-Z0-9@:%._+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}[-a-zA-Z0-9()@:%_+.~#?&/=]*)' 
        df1[2] = df1[1].str.extract(pattern, expand=True)
        df_url = df1.drop([0,1], axis=1)
        df_url =  df_url.dropna()
        df_show = df_url.copy()
        def make_hyperlink(value):
            url = "{}"
            return '=HYPERLINK("%s", "%s")' % (url.format(value), value)
        df_url[2] = df_url[2].apply(lambda x: make_hyperlink(x))

        number = st.number_input("How many sheets you'd like to create?", max_value=4, min_value=1, value=3)
        
        df_split = np.array_split(df_url, number)
        if df_split is not None:
            buffer = io.BytesIO()
            today = date.today().isoformat()
            a = 'Weekly Spiral Symplectic Report - '+today
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Write each dataframe to a different worksheet.
                for i in range(number):
                    if i==0:
                        sheet_name='K'
                    if i==1:
                        sheet_name='H'
                    if i==2:
                        sheet_name='C'
                    if i==3:
                        sheet_name='P4'
                    df_split[i].to_excel(writer, sheet_name=sheet_name, header=False, index=False)

                # df_split[0].to_excel(writer, sheet_name='K', header=False, index=False)
                # df_split[1].to_excel(writer, sheet_name='Y', header=False, index=False)

                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.close()

            st.download_button(
                label="Download report",
                data=buffer,
                file_name= a+".xlsx",
                mime="application/vnd.ms-excel"
            )
        st.dataframe(df_show, 400)
        
else:
    st.error('Paste the volume report and press "**Ctrl + Enter**" to download the spreadsheet!', icon="🚨")
    st.stop()
