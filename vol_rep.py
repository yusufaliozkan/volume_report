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
from xlsxwriter import Workbook


from bokeh.models.widgets import Button
from bokeh.models import CustomJS
from streamlit_bokeh_events import streamlit_bokeh_events
import pandas as pd
from io import StringIO

# Setting the app page layout
st.set_page_config(layout = "centered", page_title='Spiral Symplectic Report processing tool', page_icon='https://pbs.twimg.com/profile_images/1509826209563263008/cNh9JRjd_400x400.jpg')
path='https://upload.wikimedia.org/wikipedia/en/thumb/3/32/Logo_for_Imperial_College_London.svg/2560px-Logo_for_Imperial_College_London.svg.png'
st.image(path, width=300)

st.title('Spiral Symplectic Report processing tool') 

# st.sidebar.markdown("# Spiral Symplectic Report processor")
st.markdown('* Copy the text of volume report') 
st.markdown('* Paste into the box below')
st.markdown("* Click 'Press here to start'")
st.markdown("* Click 'Process report' (if a notification comes up, click 'Allow'.)")
st.markdown("* Click 'Download report'")

txt=st.text_area('Paste the volume report text here: ', ' ', placeholder='Enter') #, header=None
if st.button('Press here to start'):    
    copy_button = Button(label="Process report")
    copy_button.js_on_event("button_click", CustomJS(code="""
        navigator.clipboard.readText().then(text => document.dispatchEvent(new CustomEvent("GET_TEXT", {detail: text})))
        """))
    result = streamlit_bokeh_events(
        copy_button,
        events="GET_TEXT",
        key="get_text",
        refresh_on_update=False,
        override_height=75,
        debounce_time=0)

    if result:
        if "GET_TEXT" in result:        
            df = pd.DataFrame(StringIO(result.get("GET_TEXT")))
            # st.table(df)
            if df is not None:
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
                if df_split is not None:
                    buffer = io.BytesIO()
                    today = date.today().isoformat()
                    a = 'Weekly Spiral Symplectic Report - '+today
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        # Write each dataframe to a different worksheet.
                        df_split[0].to_excel(writer, sheet_name='K', header=False, index=False)
                        df_split[1].to_excel(writer, sheet_name='Y', header=False, index=False)

                        # Close the Pandas Excel writer and output the Excel file to the buffer
                        writer.save()

                    st.download_button(
                        label="Download report",
                        data=buffer,
                        file_name= a+".xlsx",
                        mime="application/vnd.ms-excel"
                    )
# else:
#         st.write("**Please paste the volume report!**")                    
