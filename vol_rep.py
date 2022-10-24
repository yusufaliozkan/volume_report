# streamlit run "D:\OneDrive - Imperial College London\07.Projects\13.Data\vol_report\vol_rep.py"

import pandas as pd
import openpyxl
# from urlextract import URLExtract
import re
import numpy as np
from operator import index
import io
import streamlit as st
from io import BytesIO
from datetime import date


# Setting the app page layout
st.set_page_config(layout = "wide", page_title='Copyright statements dashboard', page_icon="https://upload.wikimedia.org/wikipedia/commons/thumb/b/b0/Copyright.svg/220px-Copyright.svg.png")
st.markdown("# Copyright statements dashboard")

st.sidebar.markdown("# Copyright statements dashboard")


# text = st.text_area('Paste the volume report here: ', ' ')
df=pd.read_clipboard(st.text_area('Paste the volume report here: ', ' '), header=None)

with st.expander('Do not check', expanded=False):
    df1 = df.drop([0])
    df1[1] = df1[0].str.extract('Spiral:(.*)')
    pattern = r'(https?:\/\/(?:www\.)?[-a-zA-Z0-9@:%._+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}[-a-zA-Z0-9()@:%_+.~#?&/=]*)' 
    df1[2] = df1[1].str.extract(pattern, expand=True)
    df_url = df1.drop([0,1], axis=1)

    def make_hyperlink(value):
        url = "{}"
        return '=HYPERLINK("%s", "%s")' % (url.format(value), value)
    df_url[2] = df_url[2].apply(lambda x: make_hyperlink(x))
    df_split = np.array_split(df_url, 2)

    # df['url'] = df['Journal articles with Symplectic Volume details'].str.extract('Spiral:(.*)')
    # pattern = r'(https?:\/\/(?:www\.)?[-a-zA-Z0-9@:%._+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}[-a-zA-Z0-9()@:%_+.~#?&/=]*)' 
    # df['Spiral url'] = df['url'].str.extract(pattern, expand=True)
    # df_url = df.drop(['Journal articles with Symplectic Volume details','url'], axis=1).copy()

    # def make_hyperlink(value):
    #     url = "{}"
    #     return '=HYPERLINK("%s", "%s")' % (url.format(value), value)
    # df_url['Spiral url'] = df_url['Spiral url'].apply(lambda x: make_hyperlink(x))
    # df_split = np.array_split(df_url, 2)

buffer = io.BytesIO()

today = date.today().isoformat()
a = 'Weekly Spiral Symplectic Report - '+today

with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    # Write each dataframe to a different worksheet.
    df_split[0].to_excel(writer, sheet_name='Kim', header=False, index=False)
    df_split[1].to_excel(writer, sheet_name='Yusuf', header=False, index=False)

    # Close the Pandas Excel writer and output the Excel file to the buffer
    writer.save()

    st.download_button(
        label="Download Excel worksheets",
        data=buffer,
        file_name= a+".xlsx",
        mime="application/vnd.ms-excel"
    )
    
